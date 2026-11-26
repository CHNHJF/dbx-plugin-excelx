mod convert;

use std::path::{Path, PathBuf};

use convert::analyze::{parse_any, read_source_file, sheet_to_json};
use convert::build::{build_plans, list_tables, looks_like_sqlite, write_database};
use convert::dbx_integrate as dbx;
use convert::export::{csv_with_bom, export_csv, export_xlsx_with_notes, XLSX_MAX_ROWS};
use dbx_plugin_sdk::{PluginEmitter, PluginError, PluginHandler, PluginMetadata, PluginServer, RequestContext};
use serde_json::{json, Value};

const PLUGIN_ID: &str = "io.github.chnhjf.excelx";

fn main() -> std::io::Result<()> {
    let metadata = PluginMetadata::new(PLUGIN_ID, env!("CARGO_PKG_VERSION"));
    PluginServer::new(metadata, Plugin).serve()
}

struct Plugin;

impl PluginHandler for Plugin {
    fn handle(
        &self,
        _context: RequestContext,
        method: &str,
        params: Value,
        emitter: &PluginEmitter,
    ) -> Result<Value, PluginError> {
        match method {
            // --- import pipeline ---
            "excelx/pickFiles" => pick_files(emitter),
            "excelx/analyze" => analyze(&params),
            "excelx/ingestFile" => ingest_file(&params),
            "excelx/convert" => convert(&params, emitter),

            // --- dbx integration ---
            "excelx/registerConnection" => register_connection(&params),
            "excelx/notifyReload" => dbx::notify_reload().map(|_| json!({ "success": true })).map_err(rpc_err),
            "excelx/openInDbx" => open_in_dbx(&params),
            "excelx/listDbxConnections" => Ok(json!({ "connections": dbx::list_all_sqlite_connections() })),

            // --- export pipeline ---
            "excelx/pickDbFile" => pick_db_file(emitter),
            "excelx/dbInfo" => db_info(&params),
            "excelx/export" => export(&params),
            "excelx/openFolder" => open_folder(&params),
            _ => Err(PluginError::method_not_found(method)),
        }
    }
}

fn rpc_err(e: String) -> PluginError {
    PluginError::new(-32000, e)
}



// ---------------------------------------------------------------- import

/// Native file picker: the sandboxed iframe only sees File objects without
/// real paths, so the sidecar opens the OS dialog and returns plain paths.
/// File picking is fire-and-forget: the OS dialog blocks for as long as the
/// user browses, which outlives the host bridge's hard 120s invoke cap. So
/// this RPC only launches the dialog and returns immediately; the result (or
/// cancellation) arrives as an `excelx/pickFilesDone` event.
fn pick_files(emitter: &PluginEmitter) -> Result<Value, PluginError> {
    spawn_dialog(
        emitter.clone(),
        "excelx/pickFilesDone",
        "选择要转换的表格",
        "Spreadsheets|*.xlsx;*.xlsm;*.xlsb;*.xls;*.csv;*.tsv;*.txt|All files|*.*",
        true,
    );
    Ok(json!({ "started": true }))
}

/// Export-side picker restricted to SQLite files (any SQLite, not just ours).
fn pick_db_file(emitter: &PluginEmitter) -> Result<Value, PluginError> {
    spawn_dialog(
        emitter.clone(),
        "excelx/pickDbDone",
        "选择 SQLite 数据库文件",
        "SQLite database|*.db;*.sqlite;*.sqlite3;*.db3",
        false,
    );
    Ok(json!({ "started": true }))
}

/// Runs the PowerShell dialog on a background thread and pushes the outcome
/// as a plugin event; see pick_files for why the RPC cannot wait.
fn spawn_dialog(emitter: PluginEmitter, event: &'static str, title: &'static str, filter: &'static str, multiselect: bool) {
    std::thread::spawn(move || {
        let result = native_dialog_pick(title, filter, multiselect);
        let payload = match result {
            Ok(paths) if !paths.is_empty() => json!({ "paths": paths, "cancelled": false }),
            Ok(_) => json!({ "paths": [], "cancelled": true }),
            Err(e) => json!({ "paths": [], "error": e }),
        };
        let _ = emitter.event(event, payload);
    });
}

fn native_dialog_pick(title: &str, filter: &str, multiselect: bool) -> Result<Vec<String>, String> {
    #[cfg(target_os = "macos")]
    {
        return mac_dialog_pick(title, filter, multiselect);
    }
    #[cfg(target_os = "windows")]
    {
        return win_dialog_pick(title, filter, multiselect);
    }
    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    {
        let _ = (title, filter, multiselect);
        Err("此平台暂不支持文件选择对话框".to_string())
    }
}

/// Extensions listed in a Windows-style filter ("Name|*.a;*.b|Name2|*.*").
/// The catch-all `*` is dropped: it has no meaningful per-platform mapping.
#[cfg_attr(not(target_os = "macos"), allow(dead_code))]
fn filter_extensions(filter: &str) -> Vec<String> {
    filter
        .split('|')
        .skip(1)
        .step_by(2)
        .flat_map(|group| group.split(';'))
        .filter_map(|pat| pat.trim().strip_prefix("*."))
        .filter(|e| !e.is_empty() && *e != "*")
        .map(|e| e.to_ascii_lowercase())
        .collect()
}

/// macOS: system file picker via osascript. `choose file of type` takes real
/// Uniform Type Identifiers; made-up ones (`public.xlsx`) match nothing and
/// grey every file out (issue #1), so extensions are mapped to the UTIs that
/// macOS/Office actually declare. The result is converted to POSIX paths
/// inside AppleScript — the raw `alias "HD:Users:..."` text is an HFS path
/// and would never open. Output is plain UTF-8, one path per line.
#[cfg(target_os = "macos")]
fn mac_dialog_pick(title: &str, filter: &str, multiselect: bool) -> Result<Vec<String>, String> {
    let script = mac_pick_script(title, filter, multiselect);
    let out = std::process::Command::new("osascript")
        .args(["-e", &script])
        .output()
        .map_err(|e| format!("启动文件选择对话框失败: {e}"))?;
    if !out.status.success() {
        let detail = String::from_utf8_lossy(&out.stderr).trim().to_string();
        return Err(if detail.is_empty() {
            "文件选择对话框无法运行".to_string()
        } else {
            format!("文件选择对话框无法运行：{}", truncate_chars(&detail, 200))
        });
    }
    let text = String::from_utf8_lossy(&out.stdout).into_owned();
    let paths: Vec<String> = text
        .lines()
        .map(|l| l.trim().trim_end_matches('\u{feff}').to_string())
        .filter(|l| !l.is_empty())
        .collect();
    if paths.is_empty() && !text.trim().is_empty() {
        // Defensive: a non-empty reply that is not a path list means the
        // script output format changed; surface it instead of a bogus path.
        return Err(format!("文件选择结果无法识别：{}", truncate_chars(text.trim(), 200)));
    }
    Ok(paths)
}

/// The AppleScript source for the picker. Pure so the macOS CI job can
/// compile-check it with `osascript` without ever showing a dialog, and so
/// the UTI mapping is unit-testable on every platform.
#[cfg_attr(not(target_os = "macos"), allow(dead_code))]
fn mac_pick_script(title: &str, filter: &str, multiselect: bool) -> String {
    let mut utis: Vec<&'static str> = Vec::new();
    for ext in filter_extensions(filter) {
        for uti in mac_utis_for_ext(&ext) {
            if !utis.contains(&uti) {
                utis.push(uti);
            }
        }
    }
    if utis.is_empty() {
        utis.push("public.data");
    }
    let type_list = utis.iter().map(|u| format!("\"{u}\"")).collect::<Vec<_>>().join(", ");
    let choose = if multiselect { "choose file with multiple selections allowed" } else { "choose file" };
    // Double quotes inside the prompt would terminate the AppleScript string.
    let prompt = title.replace('"', "’");
    format!(
        "try\n\
         \tset pickedFiles to {choose} with prompt \"{prompt}\" of type {{{type_list}}}\n\
         \tif class of pickedFiles is not list then set pickedFiles to {{pickedFiles}}\n\
         \tset outText to \"\"\n\
         \trepeat with f in pickedFiles\n\
         \t\tset outText to outText & (POSIX path of (contents of f)) & linefeed\n\
         \tend repeat\n\
         \treturn outText\n\
         \ton error number -128\n\
         \treturn \"\"\n\
         \tend try"
    )
}

/// Extension → real macOS UTI. Only real, system-or-Office-declared UTIs:
/// `public.xlsx` / `public.csv` do not exist and would grey out every file
/// (the original bug in issue #1).
#[cfg_attr(not(target_os = "macos"), allow(dead_code))]
fn mac_utis_for_ext(ext: &str) -> &'static [&'static str] {
    match ext {
        "xlsx" => &["org.openxmlformats.spreadsheetml.sheet"],
        "xlsm" => &["org.openxmlformats.spreadsheetml.sheet.macroEnabled"],
        "xlsb" => &["com.microsoft.excel.xlsb"],
        "xls" => &["com.microsoft.excel.xls"],
        "csv" => &["public.comma-separated-values-text", "public.plain-text"],
        "tsv" => &["public.tab-separated-values-text", "public.plain-text"],
        "txt" => &["public.plain-text"],
        // SQLite files have no system-declared UTI; accept any file.
        _ => &["public.data"],
    }
}

fn truncate_chars(s: &str, n: usize) -> String {
    s.chars().take(n).collect()
}

#[cfg(target_os = "windows")]
fn win_dialog_pick(title: &str, filter: &str, multiselect: bool) -> Result<Vec<String>, String> {
    // PowerShell + Windows Forms file dialog; STA is required for dialogs.
    // Title text is ASCII-safe by construction (callers pass literals).
    //
    // Selected paths travel via a UTF-8 temp file, NOT stdout: a piped
    // PowerShell 5.1 console writes GBK, so Chinese filenames arrive as
    // U+FFFD garbage and every later file read fails with os error 2.
    let ms = if multiselect { "$true" } else { "$false" };
    let script = format!(
        r#"
Add-Type -AssemblyName System.Windows.Forms
$top = New-Object System.Windows.Forms.Form
$top.TopMost = $true
$top.MinimizeBox = $false
$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Filter = '{filter}'
$dlg.Multiselect = {ms}
$dlg.Title = '{title}'
if ($dlg.ShowDialog($top) -eq [System.Windows.Forms.DialogResult]::OK) {{
  $dlg.FileNames | Out-File -FilePath $env:EXCELX_PICK_OUT -Encoding utf8
  $top.Dispose()
}} else {{
  '' | Out-File -FilePath $env:EXCELX_PICK_OUT -Encoding utf8
  $top.Dispose()
}}
"#,
    );
    let encoded = base64_of(&script);
    let out_file = std::env::temp_dir().join(format!("excelx-pick-{}.txt", std::process::id()));
    let out_path = out_file.to_string_lossy().to_string();
    // Windows PowerShell 5.1 first; fall back to PowerShell 7 (`pwsh`) when
    // the legacy executable is unavailable or blocked by policy. Both ship
    // System.Windows.Forms on Windows and accept -STA / -EncodedCommand.
    let mut last_spawn_error = String::new();
    let mut out = None;
    for exe in ["powershell", "pwsh"] {
        match std::process::Command::new(exe)
            .args(["-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", &encoded])
            .env("EXCELX_PICK_OUT", &out_path)
            .output()
        {
            Ok(result) => {
                // A successful launch that did not write the result file means
                // the dialog itself could not run (e.g. constrained language
                // mode); try the next interpreter before giving up.
                if result.status.success() || out_file.is_file() {
                    out = Some(result);
                    break;
                }
                last_spawn_error = String::from_utf8_lossy(&result.stderr).trim().to_string();
            }
            Err(e) => last_spawn_error = e.to_string(),
        }
    }
    let Some(out) = out else {
        return Err(format!(
            "启动文件选择对话框失败：{}。可以直接把文件的完整路径粘贴到输入框里。",
            truncate_chars(&last_spawn_error, 160)
        ));
    };
    let picked = std::fs::read(&out_file);
    let _ = std::fs::remove_file(&out_file);
    if !out.status.success() {
        return Err("文件选择对话框无法运行。可以直接把文件的完整路径粘贴到输入框里。".to_string());
    }
    let text = match picked {
        Ok(bytes) => {
            // Out-File utf8 emits a BOM; strip it, then decode as UTF-8
            // strictly so any encoding surprise surfaces here, not later.
            let body = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(&bytes);
            String::from_utf8(body.to_vec()).map_err(|_| "文件选择结果编码异常".to_string())?
        }
        Err(_) => return Err("文件选择结果读取失败".to_string()),
    };
    Ok(text
        .lines()
        .map(|l| l.trim().trim_end_matches('\u{feff}').to_string())
        .filter(|l| !l.is_empty())
        .collect())
}

/// UTF-16LE base64: -EncodedCommand expects it and this sidesteps console
/// codepage mangling of the script text.
fn base64_of(s: &str) -> String {
    let utf16: Vec<u8> = s.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
    base64_encode(&utf16)
}

fn base64_encode(data: &[u8]) -> String {
    const TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity((data.len() + 2) / 3 * 4);
    for chunk in data.chunks(3) {
        let b = [chunk[0], *chunk.get(1).unwrap_or(&0), *chunk.get(2).unwrap_or(&0)];
        let n = ((b[0] as u32) << 16) | ((b[1] as u32) << 8) | b[2] as u32;
        out.push(TABLE[(n >> 18) as usize & 63] as char);
        out.push(TABLE[(n >> 12) as usize & 63] as char);
        out.push(if chunk.len() > 1 { TABLE[(n >> 6) as usize & 63] as char } else { '=' });
        out.push(if chunk.len() > 2 { TABLE[n as usize & 63] as char } else { '=' });
    }
    out
}

fn analyze(params: &Value) -> Result<Value, PluginError> {
    let path = str_param(params, "path")?;
    let path = normalize_pasted_path(path);
    let source = Path::new(path);
    if source.is_dir() {
        return Err(rpc_err("这是一个文件夹，请选择具体的表格文件（.xlsx / .xls / .csv）".to_string()));
    }
    if !source.exists() {
        return Err(rpc_err(format!("文件不存在：{path}")));
    }
    let bytes = read_source_file(path).map_err(rpc_err)?;
    analyze_bytes(path, &bytes, params).map(|v| {
        let mut v = v;
        if let Some(obj) = v.as_object_mut() {
            obj.insert("path".into(), Value::String(path.to_string()));
        }
        v
    })
}

/// Shared analyze core over already-loaded bytes, so both the path-based and
/// the streamed (host fileTransfer fallback) entry points behave identically.
fn analyze_bytes(path: &str, bytes: &[u8], params: &Value) -> Result<Value, PluginError> {
    // Optional per-request header row (1-based; 0 = no header). The preview UI
    // re-analyzes with this when the user moves the header-row selector, so
    // the preview matches what convert will build.
    let header_row = params.get("headerRow").and_then(Value::as_u64).map(|v| v as usize);
    let sheets = match header_row {
        Some(h) if h != 1 => convert::analyze::reparse_with_headers(path, bytes, &[h]).map_err(rpc_err)?,
        _ => parse_any(path, bytes).map_err(rpc_err)?,
    };
    let total_rows: usize = sheets.iter().map(|s| s.rows.len()).sum();
    Ok(json!({
        "sheets": sheets.iter().map(sheet_to_json).collect::<Vec<_>>(),
        "totalRows": total_rows,
    }))
}

/// Fallback import path for hosts whose native dialog cannot run (or where
/// the UI used the host's own picker): the UI streams the picked file's
/// bytes (base64, chunks up to the host's 8 MiB limit) and the sidecar
/// materialises it under the plugin's staging directory. Returns the staging
/// path plus the analyze payload so the UI proceeds exactly like the native
/// path. `name` is sanitised and never used for directory traversal.
///
/// Protocol: `offset` is the RAW BYTE offset of this chunk within the file
/// (not the base64 string offset), `totalSize` the raw byte size. Chunks may
/// be dispatched concurrently by the SDK worker pool, so a chunk with
/// `offset > 0` is never trusted to arrive in order: each chunk is parked
/// under `<key>.<offset>.part` and folded into the `.partial` file in
/// strict offset order. Only the fold that reaches `totalSize` analyses.
fn ingest_file(params: &Value) -> Result<Value, PluginError> {
    let name = sanitized_ingest_name(
        params.get("name").and_then(Value::as_str).unwrap_or("dropped"),
    );
    let data_b64 = str_param(params, "dataBase64")?;
    let offset = params.get("offset").and_then(Value::as_u64).unwrap_or(0);
    let total = params.get("totalSize").and_then(Value::as_u64);
    let bytes = base64_decode(data_b64).map_err(|e| rpc_err(format!("文件数据解码失败: {e}")))?;
    let dir = ingest_dir();
    std::fs::create_dir_all(&dir).map_err(|e| rpc_err(format!("创建暂存目录失败: {e}")))?;
    // One staging slot per name: a re-pick of the same file reuses it so
    // stale partials from an aborted upload never leak forever.
    let stem = Path::new(&name).file_stem().and_then(|s| s.to_str()).unwrap_or("dropped").to_string();
    let ext = Path::new(&name).extension().and_then(|s| s.to_str()).unwrap_or("bin").to_ascii_lowercase();
    let key = format!("{stem}-{}", total.unwrap_or(0));
    let partial = dir.join(format!("{key}.partial"));
    let target = dir.join(format!("{key}.{ext}"));

    // Single-shot: the whole file in one request (common for small files).
    if offset == 0 && total.is_some_and(|t| t as usize == bytes.len()) {
        std::fs::write(&partial, &bytes).map_err(|e| rpc_err(format!("写入暂存文件失败: {e}")))?;
    } else {
        // Park this chunk at its offset, then fold contiguous chunks from
        // the current file length forward. The fold is guarded by a lock file
        // so two concurrent folds cannot interleave appends.
        let chunk_path = dir.join(format!("{key}.{offset}.part"));
        std::fs::write(&chunk_path, &bytes).map_err(|e| rpc_err(format!("写入暂存分片失败: {e}")))?;
        // Lock covers the fold so the partial length is always observed after
        // any earlier fold completed.
        let _guard = ingest_lock(&dir, &key);
        fold_chunks(&dir, &key, &partial).map_err(rpc_err)?;
        drop(_guard);
    }
    let written = std::fs::metadata(&partial).map(|m| m.len()).unwrap_or(0);
    let done = total.map(|t| written >= t).unwrap_or(true);
    if !done {
        return Ok(json!({ "path": partial.to_string_lossy(), "received": written, "complete": false }));
    }
    let bytes = std::fs::read(&partial).map_err(|e| rpc_err(format!("读取暂存文件失败: {e}")))?;
    let _ = std::fs::rename(&partial, &target);
    ingest_dir_gc(&dir, &key);
    let path = target.to_string_lossy().to_string();
    let mut payload = analyze_bytes(&path, &bytes, &params)?;
    if let Some(obj) = payload.as_object_mut() {
        obj.insert("path".into(), Value::String(path.clone()));
        obj.insert("staged".into(), Value::Bool(true));
    }
    Ok(payload)
}

/// Appends every parked chunk (`<key>.<offset>.part`) that is contiguous
/// with the current end of `partial`, in order, removing each as it is
/// consumed. Returns the resulting length of `partial`.
fn fold_chunks(dir: &Path, key: &str, partial: &Path) -> Result<u64, String> {
    use std::io::Write;
    let mut written = std::fs::metadata(partial).map(|m| m.len()).unwrap_or(0);
    loop {
        let next = dir.join(format!("{key}.{written}.part"));
        if !next.is_file() {
            break;
        }
        let chunk = std::fs::read(&next).map_err(|e| format!("读取暂存分片失败: {e}"))?;
        let mut file = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(partial)
            .map_err(|e| format!("写入暂存文件失败: {e}"))?;
        file.write_all(&chunk).map_err(|e| format!("写入暂存文件失败: {e}"))?;
        drop(file);
        written += chunk.len() as u64;
        let _ = std::fs::remove_file(&next);
    }
    Ok(written)
}

/// File-system mutex for the chunk fold: `create_new` fails when another
/// fold holds it. Held only for microseconds; a stale lock (crash mid-fold)
/// is broken after a short wait.
fn ingest_lock(dir: &Path, key: &str) -> IngestLock {
    let lock = dir.join(format!("{key}.lock"));
    for _ in 0..50 {
        match std::fs::OpenOptions::new().create_new(true).write(true).open(&lock) {
            Ok(_) => return IngestLock(lock),
            Err(_) => std::thread::sleep(std::time::Duration::from_millis(20)),
        }
    }
    // Stale lock (holder crashed mid-fold): break it and proceed.
    let _ = std::fs::remove_file(&lock);
    let _ = std::fs::OpenOptions::new().create_new(true).write(true).open(&lock);
    IngestLock(lock)
}

impl Drop for IngestLock {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}

struct IngestLock(PathBuf);

/// Bounds the staging dir: removes leftovers (partials, parts, locks) older
/// than a day so aborted uploads never accumulate. Never touches the file
/// currently being assembled.
fn ingest_dir_gc(dir: &Path, keep_key: &str) {
    let cutoff = std::time::SystemTime::now() - std::time::Duration::from_secs(24 * 60 * 60);
    let Ok(entries) = std::fs::read_dir(dir) else { return };
    for entry in entries.flatten() {
        let name = entry.file_name().to_string_lossy().to_string();
        let Ok(meta) = entry.metadata() else { continue };
        if meta.is_dir() || name.starts_with(keep_key) {
            continue;
        }
        if let Ok(modified) = meta.modified() {
            if modified < cutoff {
                let _ = std::fs::remove_file(entry.path());
            }
        }
    }
}

/// Sanitises a streamed file's name for the staging dir: collapses `..`,
/// strips path separators and filesystem-hostile characters, and keeps the
/// extension so `parse_any` can pick the right parser.
fn sanitized_ingest_name(raw: &str) -> String {
    let base = raw.rsplit(['/', '\\']).next().unwrap_or(raw).trim();
    // Collapse any run of dots to a single dot first: safe_name would turn
    // ".." into its "export" placeholder, and a leading dot makes a hidden file.
    let mut collapsed = String::with_capacity(base.len());
    let mut prev_dot = false;
    for c in base.chars() {
        if c == '.' {
            if !prev_dot && !collapsed.is_empty() {
                collapsed.push('.');
            }
            prev_dot = true;
        } else {
            collapsed.push(c);
            prev_dot = false;
        }
    }
    let cleaned = safe_name(&collapsed).replace("..", ".");
    if cleaned.is_empty() || cleaned == "." || cleaned == "export" && base.trim_start_matches('.').is_empty() {
        "dropped.bin".to_string()
    } else {
        cleaned
    }
}

/// Directory where streamed-in files are materialised (host fileTransfer
/// fallback only). Temp dir keeps it out of the user's documents folder and
/// the OS cleans it eventually; stale files are bounded by `ingest_dir_gc`.
fn ingest_dir() -> PathBuf {
    std::env::temp_dir().join("excelx-ingest")
}

/// Display stem for a staged file: strips the `-<size>` suffix the staging
/// name carries, so `report-12.xlsx` becomes table/connection name `report`.
pub fn staged_display_stem(path: &str) -> String {
    let raw = Path::new(path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("excelx");
    match raw.rsplit_once('-') {
        Some((head, digits)) if !digits.is_empty() && digits.chars().all(|c| c.is_ascii_digit()) => head.to_string(),
        _ => raw.to_string(),
    }
}

fn base64_decode(s: &str) -> Result<Vec<u8>, String> {
    const fn val(c: u8) -> Option<u8> {
        match c {
            b'A'..=b'Z' => Some(c - b'A'),
            b'a'..=b'z' => Some(c - b'a' + 26),
            b'0'..=b'9' => Some(c - b'0' + 52),
            b'+' => Some(62),
            b'/' => Some(63),
            _ => None,
        }
    }
    let mut out = Vec::with_capacity(s.len() / 4 * 3);
    let mut acc: u32 = 0;
    let mut bits = 0u32;
    for &c in s.as_bytes() {
        if c == b'=' || c == b'\n' || c == b'\r' || c == b' ' {
            continue;
        }
        let v = val(c).ok_or("invalid base64")?;
        acc = (acc << 6) | v as u32;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((acc >> bits) as u8);
        }
    }
    Ok(out)
}

/// Pasted paths: tolerate the quotes Explorer/"Copy as path" adds and the
/// whitespace around them, so a manual paste works first time.
fn normalize_pasted_path(path: &str) -> &str {
    path.trim().trim_matches('"').trim()
}

fn convert(params: &Value, emitter: &PluginEmitter) -> Result<Value, PluginError> {
    let path = normalize_pasted_path(str_param(params, "path")?);
    let plans = params.get("plans").and_then(Value::as_array).cloned().unwrap_or_default();
    // overwrite: replace an existing .db at the default path (re-convert flow)
    // instead of minting _2/_3 copies. Defaults to false for safety.
    let overwrite = params.get("overwrite").and_then(Value::as_bool).unwrap_or(false);
    let out_path = if overwrite {
        output_path_overwrite(path)
    } else {
        output_path_for(path)
    };
    // Self-overwrite guard: converting a .db-suffixed text file would resolve
    // the output to the input itself and destroy the source on rename.
    let same_target = std::fs::canonicalize(&out_path).ok()
        .zip(std::fs::canonicalize(path).ok())
        .map(|(a, b)| a == b)
        .unwrap_or_else(|| out_path == Path::new(path));
    if same_target {
        return Err(rpc_err("输出路径与源文件相同，已取消转换以免覆盖源文件。".to_string()));
    }

    let bytes = read_source_file(path).map_err(rpc_err)?;
    let header_rows: Vec<usize> = plans
        .iter()
        .map(|p| p.get("headerRow").and_then(Value::as_u64).map(|v| v as usize).unwrap_or(1))
        .collect();
    let sheets = if header_rows.iter().any(|h| *h != 1) {
        convert::analyze::reparse_with_headers(path, &bytes, &header_rows).map_err(rpc_err)?
    } else {
        parse_any(path, &bytes).map_err(rpc_err)?
    };
    let planned = build_plans(&sheets, &plans);

    // Build beside the target first; rename in so a failure never leaves a
    // half-written database at the final path.
    if let Some(parent) = out_path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let staging = out_path.with_extension("db.partial");
    let _ = std::fs::remove_file(&staging);
    {
        let conn = rusqlite::Connection::open(&staging)
            .map_err(|e| rpc_err(format!("创建数据库失败: {e}")))?;
        write_database(&conn, &planned, path, |p| {
            let _ = emitter.event(
                "excelx/progress",
                json!({
                    "stage": "convert",
                    "table": planned.get(p.table_index).map(|(_, x)| x.table_name.clone()).unwrap_or_default(),
                    "tableIndex": p.table_index + 1,
                    "tablesTotal": p.tables_total,
                    "rowsDone": p.rows_done,
                    "rowsTotal": p.rows_total,
                }),
            );
        })
        .map_err(rpc_err)?;
    }
    std::fs::rename(&staging, &out_path).map_err(|e| rpc_err(format!("落盘数据库失败: {e}")))?;

    let tables: Vec<String> = planned.iter().map(|(_, p)| p.table_name.clone()).collect();
    let total_rows: usize = planned.iter().map(|(s, _)| s.rows.len()).sum();
    let first_table = tables.first().cloned();

    let (conn_id, conn_name) = match dbx::register_connection(
        &out_path.to_string_lossy(),
        &out_path.file_stem().and_then(|s| s.to_str()).unwrap_or("excelx"),
    ) {
        Ok(v) => v,
        Err(e) => {
            // The database itself is fine; registration is best-effort. Report
            // it as a warning field instead of failing the whole convert.
            return Ok(json!({
                "dbPath": out_path.to_string_lossy(),
                "tables": tables,
                "totalRows": total_rows,
                "registered": false,
                "registerError": e,
            }));
        }
    };
    let reload_ok = dbx::notify_reload().is_ok();
    let _ = dbx::open_in_dbx(&conn_id, first_table.as_deref());

    Ok(json!({
        "dbPath": out_path.to_string_lossy(),
        "tables": tables,
        "totalRows": total_rows,
        "registered": true,
        "reloadOk": reload_ok,
        "connectionId": conn_id,
        "connectionName": conn_name,
    }))
}

/// <source>.xlsx -> <source>.db next to the file; falls back to a numbered
/// name if the target exists (never silently overwrites a previous export).
/// Files that were streamed into the temp staging dir (`excelx-ingest`)
/// get their database in the user's Documents folder instead, because a .db
/// buried in %TEMP% is effectively lost for the user.
fn output_path_for(source: &str) -> PathBuf {
    let first = default_output_candidate(source);
    if !first.exists() {
        return first;
    }
    let dir = first.parent().unwrap_or(Path::new(".")).to_path_buf();
    let stem = first.file_stem().and_then(|s| s.to_str()).unwrap_or("excelx").to_string();
    let mut candidate = first;
    let mut n = 2;
    while candidate.exists() {
        candidate = dir.join(format!("{stem}_{n}.db"));
        n += 1;
    }
    candidate
}

/// Default `<stem>.db` location for a source file: beside the original,
/// except for staged (streamed) files which resolve to Documents/ExcelX.
fn default_output_candidate(source: &str) -> PathBuf {
    if is_staged_source(source) {
        return documents_dir()
            .unwrap_or_else(std::env::temp_dir)
            .join("ExcelX")
            .join(format!("{}.db", staged_display_stem(source)));
    }
    let stem = Path::new(source)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("excelx");
    let dir = Path::new(source).parent().unwrap_or(Path::new("."));
    dir.join(format!("{stem}.db"))
}

fn is_staged_source(source: &str) -> bool {
    let staged_root = ingest_dir();
    Path::new(source).starts_with(&staged_root)
}

fn documents_dir() -> Option<PathBuf> {
    #[cfg(target_os = "windows")]
    {
        // Avoid SHGetKnownFolderPath (no extra crates): %USERPROFILE%\Documents
        // is correct for default profiles and harmless when redirected.
        std::env::var("USERPROFILE").ok().map(|h| PathBuf::from(h).join("Documents"))
    }
    #[cfg(target_os = "macos")]
    {
        std::env::var("HOME").ok().map(|h| PathBuf::from(h).join("Documents"))
    }
    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    {
        std::env::var("HOME").ok().map(|h| PathBuf::from(h).join("Documents"))
    }
}

/// Overwrite mode: the default .db path even when it already exists. Used by
/// the re-convert flow after the user confirms replacement in the UI.
fn output_path_overwrite(source: &str) -> PathBuf {
    default_output_candidate(source)
}

fn file_stem(path: &str) -> String {
    Path::new(path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("excelx")
        .to_string()
}

// ---------------------------------------------------------------- dbx bridge

fn register_connection(params: &Value) -> Result<Value, PluginError> {
    let db_path = str_param(params, "dbPath")?;
    let base_name = params
        .get("name")
        .and_then(Value::as_str)
        .map(String::from)
        .unwrap_or_else(|| file_stem(db_path));
    let (id, name) = dbx::register_connection(db_path, &base_name).map_err(rpc_err)?;
    dbx::notify_reload().map_err(rpc_err)?;
    Ok(json!({ "connectionId": id, "connectionName": name }))
}

fn open_in_dbx(params: &Value) -> Result<Value, PluginError> {
    let connection_id = str_param(params, "connectionId")?;
    let table = params.get("table").and_then(Value::as_str);
    dbx::open_in_dbx(connection_id, table).map_err(rpc_err)?;
    Ok(json!({ "success": true }))
}

// ---------------------------------------------------------------- export

fn db_info(params: &Value) -> Result<Value, PluginError> {
    let db_path = str_param(params, "path")?;
    check_db_not_locked(db_path)?;
    let p = Path::new(db_path);
    if !p.is_file() {
        return Err(rpc_err("文件不存在".to_string()));
    }
    if !looks_like_sqlite(p) {
        return Err(rpc_err("这不是 SQLite 数据库文件".to_string()));
    }
    let conn = open_sqlite_readonly(db_path)?;
    let tables = list_tables(&conn).map_err(rpc_err)?;
    let table_infos: Vec<Value> = tables
        .iter()
        .map(|(name, count)| {
            json!({
                "name": name,
                "rowCount": count,
                "overXlsxLimit": *count as usize > XLSX_MAX_ROWS,
            })
        })
        .collect();
    Ok(json!({
        "path": db_path,
        "fileSize": p.metadata().map(|m| m.len()).unwrap_or(0),
        "tables": table_infos,
    }))
}

/// Detects a live writer lock: if another process (e.g. DBX) holds the
/// database open for writing, an immutable read may serve a stale snapshot —
/// surface a clear message instead.
fn check_db_not_locked(db_path: &str) -> Result<(), PluginError> {
    let p = Path::new(db_path);
    if !p.is_file() {
        return Ok(()); // missing file handled by the caller
    }
    // A write-mode SQLite open + BEGIN IMMEDIATE acquires the same lock a
    // live writer (e.g. DBX) holds; busy here means "in use elsewhere".
    let conn = match rusqlite::Connection::open(p) {
        Ok(c) => c,
        Err(e) if e.to_string().contains("unable to open") => {
            return Err(rpc_err("该数据库正在被其他程序使用（可能是 DBX 已打开该连接）。请先关闭对应连接再导出。".to_string()))
        }
        Err(_) => return Ok(()),
    };
    let _ = conn.busy_timeout(std::time::Duration::from_millis(200));
    match conn.execute_batch("BEGIN IMMEDIATE; ROLLBACK;") {
        Ok(()) => Ok(()),
        Err(e) => {
            let msg = e.to_string();
            if msg.contains("locked") || msg.contains("busy") {
                Err(rpc_err("该数据库正在被其他程序使用（可能是 DBX 已打开该连接）。请先关闭对应连接再导出。".to_string()))
            } else {
                Ok(())
            }
        }
    }
}

/// Read-only + immutable URI: avoids creating/touching WAL sidecar files of
/// databases that may be open in DBX right now.
fn open_sqlite_readonly(db_path: &str) -> Result<rusqlite::Connection, PluginError> {
    let uri = format!(
        "file:{}?immutable=1",
        db_path.replace('\\', "/").replace('?', "%3F").replace('#', "%23")
    );
    rusqlite::Connection::open_with_flags(uri, rusqlite::OpenFlags::SQLITE_OPEN_READ_ONLY)
        .map_err(|e| rpc_err(format!("打开数据库失败: {e}")))
}

fn export(params: &Value) -> Result<Value, PluginError> {
    let db_path = str_param(params, "dbPath")?;
    let format = params.get("format").and_then(Value::as_str).unwrap_or("xlsx");
    if format != "xlsx" && format != "csv" {
        return Err(rpc_err("仅支持导出 XLSX 或 CSV".to_string()));
    }
    let tables: Vec<String> = params
        .get("tables")
        .and_then(Value::as_array)
        .map(|a| a.iter().filter_map(Value::as_str).map(String::from).collect())
        .unwrap_or_default();
    check_db_not_locked(db_path)?;
    let conn = open_sqlite_readonly(db_path)?;
    let listed = list_tables(&conn).map_err(rpc_err)?;
    let chosen: Vec<String> = if tables.is_empty() {
        listed.iter().map(|(n, _)| n.clone()).collect()
    } else {
        tables
    };
    // Over-limit tables are dropped from xlsx with a warning rather than
    // silently truncating; CSV has no row limit so nothing is dropped there.
    let mut warnings: Vec<String> = Vec::new();
    let chosen: Vec<String> = chosen
        .into_iter()
        .filter(|t| {
            let count = listed.iter().find(|(n, _)| n == t).map(|(_, c)| *c as usize).unwrap_or(0);
            if format != "csv" && count > XLSX_MAX_ROWS {
                warnings.push(format!("表 {t} 有 {count} 行，超过 xlsx 单表上限 {XLSX_MAX_ROWS} 行，已跳过（可改用 CSV 导出）"));
                false
            } else {
                true
            }
        })
        .collect();

    if chosen.is_empty() {
        return Err(rpc_err("没有可导出的表".to_string()));
    }
    let out_dir = Path::new(&db_path).parent().unwrap_or(Path::new(".")).to_path_buf();
    let base = safe_name(&file_stem(db_path));

    // CSV: one file per table, UTF-8 with BOM so Excel opens Chinese
    // correctly on double-click, no row limit. Single-table exports are
    // named `<db>.csv`; multi-table exports `<db>_<table>.csv` so they never
    // collide. `unique_path` still steps around an existing file (e.g. the
    // source spreadsheet that produced this database).
    if format == "csv" {
        let mut paths: Vec<String> = Vec::new();
        for table in &chosen {
            let csv_text = export_csv(&conn, table, usize::MAX).map_err(rpc_err)?;
            let file_name = if chosen.len() == 1 {
                format!("{base}.csv")
            } else {
                format!("{base}_{}.csv", safe_name(table))
            };
            let out_path = unique_path(out_dir.join(&file_name));
            std::fs::write(&out_path, csv_with_bom(&csv_text)).map_err(|e| rpc_err(format!("写出文件失败: {e}")))?;
            paths.push(out_path.to_string_lossy().to_string());
        }
        let total_bytes: u64 = paths
            .iter()
            .filter_map(|p| std::fs::metadata(p).ok())
            .map(|m| m.len())
            .sum();
        if paths.len() > 1 {
            warnings.push(format!("已导出 {} 个 CSV 文件（每张表一个），与数据库放在同一文件夹", paths.len()));
        }
        return Ok(json!({
            "path": paths[0],
            "paths": paths,
            "bytes": total_bytes,
            "format": "csv",
            "warnings": warnings,
        }));
    }

    let (out, truncated) = export_xlsx_with_notes(&conn, &chosen, XLSX_MAX_ROWS).map_err(rpc_err)?;
    let (bytes, file_name) = (out, format!("{base}.xlsx"));
    if truncated > 0 {
        warnings.push(format!("有 {truncated} 个单元格超过 Excel 单元格 32,767 字符上限，已截断保留前 32,000 字符"));
    }
    let out_path = unique_path(out_dir.join(&file_name));
    std::fs::write(&out_path, &bytes).map_err(|e| rpc_err(format!("写出文件失败: {e}")))?;
    Ok(json!({
        "path": out_path.to_string_lossy(),
        "paths": [out_path.to_string_lossy()],
        "bytes": bytes.len(),
        "format": "xlsx",
        "warnings": warnings,
    }))
}

/// Strips filesystem-hostile characters from a table name used as file name.
fn safe_name(name: &str) -> String {
    let cleaned: String = name
        .chars()
        .filter(|c| !matches!(c, '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|') && !c.is_control())
        .collect();
    let trimmed = cleaned.trim().trim_start_matches('.').to_string();
    if trimmed.is_empty() { "export".to_string() } else { trimmed }
}

fn unique_path(candidate: PathBuf) -> PathBuf {
    if !candidate.exists() {
        return candidate;
    }
    let stem = candidate.file_stem().and_then(|s| s.to_str()).unwrap_or("export").to_string();
    let ext = candidate.extension().and_then(|s| s.to_str()).unwrap_or("").to_string();
    let dir = candidate.parent().unwrap_or(Path::new(".")).to_path_buf();
    let mut n = 2;
    loop {
        let name = if ext.is_empty() { format!("{stem}_{n}") } else { format!("{stem}_{n}.{ext}") };
        let p = dir.join(name);
        if !p.exists() {
            return p;
        }
        n += 1;
    }
}

fn open_folder(params: &Value) -> Result<Value, PluginError> {
    let path = str_param(params, "path")?;
    let target = Path::new(path);
    let dir = if target.is_dir() { target } else { target.parent().unwrap_or(Path::new(".")) };
    #[cfg(target_os = "windows")]
    let opener = ("explorer", None::<&str>);
    #[cfg(target_os = "macos")]
    let opener = ("open", None::<&str>);
    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    let opener = ("xdg-open", None::<&str>);
    std::process::Command::new(opener.0)
        .arg(dir)
        .spawn()
        .map_err(|e| rpc_err(format!("打开文件夹失败: {e}")))?;
    Ok(json!({ "success": true, "folder": dir.to_string_lossy() }))
}

fn str_param<'a>(params: &'a Value, key: &str) -> Result<&'a str, PluginError> {
    params
        .get(key)
        .and_then(Value::as_str)
        .ok_or_else(|| PluginError::new(-32602, format!("Missing required parameter: {key}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ingest_chunks_arriving_out_of_order_are_reassembled() {
        let dir = std::env::temp_dir().join("excelx_test_ingest");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        let data = b"hello, world\nsecond line\nthird\n".to_vec();
        let (a, b) = data.split_at(10);
        // Chunk 2 (offset 10) lands first, exactly as a concurrent dispatch
        // could deliver it; then chunk 1 folds both in order.
        let part2 = dir.join("k.10.part");
        std::fs::write(&part2, b).unwrap();
        let partial = dir.join("k.partial");
        let _ = fold_chunks(&dir, "k", &partial);
        assert!(!partial.exists(), "nothing can fold before offset 0 arrives");
        let part1 = dir.join("k.0.part");
        std::fs::write(&part1, a).unwrap();
        let written = fold_chunks(&dir, "k", &partial).unwrap();
        assert_eq!(written as usize, data.len());
        assert_eq!(std::fs::read(&partial).unwrap(), data);
        assert!(!part1.exists() && !part2.exists());
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn ingest_names_are_sanitised() {
        assert_eq!(sanitized_ingest_name("../../evil/../sales q1.csv"), "sales q1.csv");
        assert_eq!(sanitized_ingest_name("/etc/passwd"), "passwd");
        assert_eq!(sanitized_ingest_name(".."), "dropped.bin");
        assert_eq!(sanitized_ingest_name("C:\\Users\\me\\报表.XLSX"), "报表.XLSX");
    }

    #[test]
    fn base64_utf16le_roundtrip_marker() {
        assert_eq!(base64_of("A"), "QQA="); // verified against Python base64
        assert_eq!(base64_encode(b"abc"), "YWJj");
        assert_eq!(base64_encode(b"ab"), "YWI=");
        assert_eq!(base64_encode(b"a"), "YQ==");
    }

    #[test]
    fn base64_decode_roundtrip_and_padding() {
        for sample in [b"".as_slice(), b"a", b"ab", b"abc", b"abcd", "中文路径".as_bytes()] {
            let encoded = base64_encode(sample);
            assert_eq!(base64_decode(&encoded).unwrap(), sample);
        }
        assert_eq!(base64_decode("YWJj\n").unwrap(), b"abc");
        assert!(base64_decode("YWJ*").is_err());
    }

    #[test]
    fn filter_exts_drop_catch_all() {
        assert_eq!(
            filter_extensions("Spreadsheets|*.xlsx;*.csv;*.txt|All files|*.*"),
            vec!["xlsx".to_string(), "csv".to_string(), "txt".to_string()]
        );
        assert!(filter_extensions("All files|*.*").is_empty());
    }

    #[test]
    fn pasted_paths_are_normalised() {
        assert_eq!(normalize_pasted_path("  \"C:\\dir\\a.csv\"  "), "C:\\dir\\a.csv");
        assert_eq!(normalize_pasted_path("/tmp/x.xlsx"), "/tmp/x.xlsx");
    }

    #[test]
    fn mac_picker_uses_only_real_utis() {
        // Every UTI the picker asks for must be a real declared type; the
        // made-up "public.xlsx" style was the root cause of issue #1.
        let filter = "Spreadsheets|*.xlsx;*.xlsm;*.xlsb;*.xls;*.csv;*.tsv;*.txt|All files|*.*";
        let script = mac_pick_script("选择要转换的表格", filter, true);
        for uti in [
            "org.openxmlformats.spreadsheetml.sheet",
            "com.microsoft.excel.xls",
            "public.comma-separated-values-text",
            "public.plain-text",
        ] {
            assert!(script.contains(&format!("\"{uti}\"")), "missing UTI {uti} in:\n{script}");
        }
        assert!(!script.contains("public.xlsx"), "made-up UTI leaked into script:\n{script}");
        assert!(!script.contains("public.csv"), "made-up UTI leaked into script:\n{script}");
        // The catch-all `*.*` contributes nothing (no `public.*`).
        assert!(!script.contains("\"public.\""), "catch-all leaked: {script}");
        // Multi-select and POSIX conversion are present.
        assert!(script.contains("choose file with multiple selections allowed"));
        assert!(script.contains("POSIX path of (contents of f)"));
        // Cancel (-128) returns "" instead of an error.
        assert!(script.contains("on error number -128"));
        // Braces are balanced so the generated AppleScript parses.
        let open = script.matches('{').count();
        let close = script.matches('}').count();
        assert_eq!(open, close, "unbalanced braces:\n{script}");
        // A double quote in the title is neutralised, not injected.
        let quoted = mac_pick_script("a \"quoted\" title", filter, false);
        assert!(!quoted.contains("\\\"quoted\\\""), "quote injection:\n{quoted}");
        assert!(quoted.contains("with prompt \"a ’quoted’ title\""));
    }

    /// macOS CI only: prove the generated AppleScript actually compiles and
    /// runs headlessly on a real system. The interactive `choose file …`
    /// statement is replaced by an empty list, so no dialog blocks the
    /// runner; everything else (try/on error, list coercion, POSIX path
    /// loop) executes for real. Also proves the POSIX-path statement returns
    /// what the picker relies on.
    #[cfg(target_os = "macos")]
    #[test]
    fn mac_picker_script_compiles_and_runs_headless() {
        let script = mac_pick_script("选择要转换的表格", "Spreadsheets|*.xlsx;*.csv|All files|*.*", true);
        // Replace the WHOLE `set pickedFiles to choose file … of type {…}`
        // statement — leaving `… of type {…}` after `{}` is a syntax error.
        let headless = script
            .lines()
            .map(|line| {
                if line.contains("set pickedFiles to choose file") {
                    "\tset pickedFiles to {}".to_string()
                } else {
                    line.to_string()
                }
            })
            .collect::<Vec<_>>()
            .join("\n");
        assert!(!headless.contains("choose file"), "stub failed:\n{headless}");
        let path = std::env::temp_dir().join("excelx-picker-headless.applescript");
        std::fs::write(&path, &headless).unwrap();
        let out = std::process::Command::new("osascript").arg(&path).output().expect("osascript present");
        assert!(
            out.status.success(),
            "generated AppleScript failed to compile/run:\n{}",
            String::from_utf8_lossy(&out.stderr)
        );
        assert_eq!(String::from_utf8_lossy(&out.stdout).trim(), "");

        // Also compile the REAL script (dialog included) without running it:
        // `osasscript` compiles a script file before executing, and a syntax
        // error is reported even when we never reach the dialog. We can't run
        // it (it blocks), so compile via `run script`-free path: use
        // `osascript -e 'compile'`-equivalent by checking the parse error
        // of the full text with a non-executing `STORE` — simplest reliable
        // form is `osascript -e "compile" ` which is not available, so we
        // rely on the stub above for semantics and on brace/quote balance
        // (checked cross-platform in mac_picker_uses_only_real_utis) for the
        // dialog line itself.

        // The POSIX conversion statement on a real file with a non-ASCII name.
        let f = std::env::temp_dir().join("excelx posix 测试.csv");
        std::fs::write(&f, b"x").unwrap();
        let conv = format!(
            "set f to POSIX file \"{}\"\nreturn POSIX path of (contents of {{f}})",
            f.to_string_lossy()
        );
        let out = std::process::Command::new("osascript").args(["-e", &conv]).output().unwrap();
        assert!(out.status.success(), "{}", String::from_utf8_lossy(&out.stderr));
        assert_eq!(String::from_utf8_lossy(&out.stdout).trim(), f.to_string_lossy());
        let _ = std::fs::remove_file(&f);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn staged_sources_resolve_to_documents_dir() {
        let staged = ingest_dir().join("report-12.xlsx");
        assert!(is_staged_source(staged.to_str().unwrap()));
        assert_eq!(staged_display_stem(staged.to_str().unwrap()), "report");
        let out = default_output_candidate(staged.to_str().unwrap());
        assert_eq!(out, documents_dir().unwrap().join("ExcelX").join("report.db"));
        // regular files stay beside the source (build the expected path with
        // PathBuf so the separator is platform-correct).
        let regular = std::env::temp_dir().join("data").join("a.xlsx");
        assert!(!is_staged_source(regular.to_str().unwrap()));
        assert_eq!(default_output_candidate(regular.to_str().unwrap()), std::env::temp_dir().join("data").join("a.db"));
        // a stem that merely contains a dash is not a staged size suffix
        let dashed = std::env::temp_dir().join("sales-2024-q1.xlsx");
        assert_eq!(staged_display_stem(dashed.to_str().unwrap()), "sales-2024-q1");
    }

    #[test]
    fn output_paths_never_overwrite() {
        let dir = std::env::temp_dir().join("excelx_test_paths");
        let _ = std::fs::create_dir_all(&dir);
        let first = dir.join("data.db");
        std::fs::write(&first, b"x").unwrap();
        let next = output_path_for(dir.join("data.xlsx").to_str().unwrap());
        assert_eq!(next, dir.join("data_2.db"));
        let _ = std::fs::remove_file(&first);
        let _ = std::fs::remove_file(&next);
    }

    #[test]
    fn safe_names() {
        assert_eq!(safe_name("订单/明细:*?"), "订单明细");
        assert_eq!(safe_name("..."), "export");
        assert_eq!(safe_name("表1"), "表1");
    }

    #[test]
    fn unique_path_appends_counter() {
        let dir = std::env::temp_dir().join("excelx_test_unique");
        let _ = std::fs::create_dir_all(&dir);
        let first = dir.join("out.xlsx");
        std::fs::write(&first, b"x").unwrap();
        assert_eq!(unique_path(dir.join("out.xlsx")), dir.join("out_2.xlsx"));
        std::fs::write(dir.join("out_2.xlsx"), b"x").unwrap();
        assert_eq!(unique_path(dir.join("out.xlsx")), dir.join("out_3.xlsx"));
        let _ = std::fs::remove_file(&first);
        let _ = std::fs::remove_file(dir.join("out_2.xlsx"));
    }
}
