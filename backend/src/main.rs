mod convert;

use std::path::{Path, PathBuf};

use convert::analyze::{parse_any, read_source_file, sheet_to_json};
use convert::build::{build_plans, list_tables, looks_like_sqlite, write_database};
use convert::dbx_integrate as dbx;
use convert::export::{export_xlsx_with_notes, XLSX_MAX_ROWS};
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
        "Spreadsheets|*.xlsx;*.xlsm;*.xlsb;*.xls;*.csv;*.tsv|All files|*.*",
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
    let out = std::process::Command::new("powershell")
        .args(["-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", &encoded])
        .env("EXCELX_PICK_OUT", &out_path)
        .output()
        .map_err(|e| format!("启动文件选择对话框失败: {e}"))?;
    let picked = std::fs::read(&out_file);
    let _ = std::fs::remove_file(&out_file);
    if !out.status.success() {
        return Err("文件选择对话框无法运行".to_string());
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
    let bytes = read_source_file(path).map_err(rpc_err)?;
    // Optional per-request header row (1-based; 0 = no header). The preview UI
    // re-analyzes with this when the user moves the header-row selector, so
    // the preview matches what convert will build.
    let header_row = params.get("headerRow").and_then(Value::as_u64).map(|v| v as usize);
    let sheets = match header_row {
        Some(h) if h != 1 => convert::analyze::reparse_with_headers(path, &bytes, &[h]).map_err(rpc_err)?,
        _ => parse_any(path, &bytes).map_err(rpc_err)?,
    };
    let total_rows: usize = sheets.iter().map(|s| s.rows.len()).sum();
    Ok(json!({
        "path": path,
        "sheets": sheets.iter().map(sheet_to_json).collect::<Vec<_>>(),
        "totalRows": total_rows,
    }))
}

fn convert(params: &Value, emitter: &PluginEmitter) -> Result<Value, PluginError> {
    let path = str_param(params, "path")?;
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

    let (conn_id, conn_name) = match dbx::register_connection(&out_path.to_string_lossy(), &file_stem(path)) {
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

/// Overwrite mode: the default .db path even when it already exists. Used by
/// the re-convert flow after the user confirms replacement in the UI.
fn output_path_overwrite(source: &str) -> PathBuf {
    let stem = Path::new(source)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("excelx");
    let dir = Path::new(source).parent().unwrap_or(Path::new("."));
    dir.join(format!("{stem}.db"))
}

/// <source>.xlsx -> <source>.db next to the file; falls back to a numbered
/// name if the target exists (never silently overwrites a previous export).
fn output_path_for(source: &str) -> PathBuf {
    let stem = Path::new(source)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("excelx");
    let dir = Path::new(source).parent().unwrap_or(Path::new("."));
    let mut candidate = dir.join(format!("{stem}.db"));
    let mut n = 2;
    while candidate.exists() {
        candidate = dir.join(format!("{stem}_{n}.db"));
        n += 1;
    }
    candidate
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
    if format != "xlsx" {
        return Err(rpc_err("仅支持导出 XLSX".to_string()));
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
    // silently truncating; CSV caps at the same limit per file.
    let mut warnings: Vec<String> = Vec::new();
    let chosen: Vec<String> = chosen
        .into_iter()
        .filter(|t| {
            let count = listed.iter().find(|(n, _)| n == t).map(|(_, c)| *c as usize).unwrap_or(0);
            if format != "csv" && count > XLSX_MAX_ROWS {
                warnings.push(format!("表 {t} 有 {count} 行，超过 xlsx 单表上限 {XLSX_MAX_ROWS} 行，已跳过（可单独导出 CSV）"));
                false
            } else {
                true
            }
        })
        .collect();

    if chosen.is_empty() {
        return Err(rpc_err("没有可导出的表".to_string()));
    }
    let (out, truncated) = export_xlsx_with_notes(&conn, &chosen, XLSX_MAX_ROWS).map_err(rpc_err)?;
    let (bytes, file_name) = (out, format!("{}.xlsx", safe_name(&file_stem(db_path))));
    let _ = format;
    if truncated > 0 {
        warnings.push(format!("有 {truncated} 个单元格超过 Excel 单元格 32,767 字符上限，已截断保留前 32,000 字符"));
    }
    let out_dir = Path::new(&db_path).parent().unwrap_or(Path::new(".")).to_path_buf();
    let out_path = unique_path(out_dir.join(&file_name));
    std::fs::write(&out_path, &bytes).map_err(|e| rpc_err(format!("写出文件失败: {e}")))?;
    Ok(json!({ "path": out_path.to_string_lossy(), "bytes": bytes.len(), "warnings": warnings }))
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
    std::process::Command::new("explorer")
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
    fn base64_utf16le_roundtrip_marker() {
        assert_eq!(base64_of("A"), "QQA="); // verified against Python base64
        assert_eq!(base64_encode(b"abc"), "YWJj");
        assert_eq!(base64_encode(b"ab"), "YWI=");
        assert_eq!(base64_encode(b"a"), "YQ==");
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
