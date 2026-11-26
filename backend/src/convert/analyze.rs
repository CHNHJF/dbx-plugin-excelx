//! Spreadsheet parsing + type inference. Pure logic (no I/O policy) except for
//! file reading, so it can be unit-tested against in-memory fixtures.

use serde_json::{json, Value};

use super::names::{sanitize_column_name, sanitize_sheet_label};

pub const SAMPLE_ROWS: usize = 50;
/// Rows scanned for type inference. Sheets are typically homogeneous well
/// before this; a full scan would put 1M-row sheets out of reach.
const INFER_ROWS: usize = 5000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColType {
    Text,
    Integer,
    Real,
    Date,
}

impl ColType {
    pub fn as_str(self) -> &'static str {
        match self {
            ColType::Text => "TEXT",
            ColType::Integer => "INTEGER",
            ColType::Real => "REAL",
            ColType::Date => "DATE",
        }
    }
    pub fn from_str(s: &str) -> Option<Self> {
        match s {
            "TEXT" => Some(ColType::Text),
            "INTEGER" => Some(ColType::Integer),
            "REAL" => Some(ColType::Real),
            "DATE" => Some(ColType::Date),
            _ => None,
        }
    }
}

/// A parsed sheet: raw string cells (row-major) plus inference results.
#[derive(Debug, Clone)]
pub struct SheetData {
    pub source_name: String,
    pub label: String,
    pub headers: Vec<String>,
    pub rows: Vec<Vec<String>>,
    pub inferred_types: Vec<ColType>,
    pub skipped_header_rows: usize,
    /// Hidden sheets (Excel hidden / very-hidden) convert only when the user
    /// explicitly opts in; the preview marks them.
    pub hidden: bool,
}

pub fn col_type_from_cells(cells: impl Iterator<Item = String>) -> ColType {
    let mut saw_int = false;
    let mut saw_real = false;
    let mut saw_date = false;
    let mut saw_any = false;
    for cell in cells {
        let t = cell.trim();
        if t.is_empty() {
            continue;
        }
        saw_any = true;
        if parse_date_loose(t).is_some() {
            saw_date = true;
            if saw_int || saw_real {
                return ColType::Text;
            }
            continue;
        }
        if saw_date {
            return ColType::Text;
        }
        // Leading-zero digits are identifiers (001, 007): TEXT, never numbers.
        if t.len() > 1 && t.starts_with('0') && t.chars().all(|c| c.is_ascii_digit()) {
            return ColType::Text;
        }
        // Long all-digit runs (≥ 7 digits, no separators) are phone numbers,
        // ID cards, tracking numbers... never quantities. Keep them TEXT so a
        // 11-digit mobile number stays an exact string instead of a REAL.
        if t.len() >= 7 && t.chars().all(|c| c.is_ascii_digit()) {
            return ColType::Text;
        }
        match parse_number_loose(t) {
            Some(v) if v.fract() == 0.0 && !t.contains(['.', 'e', 'E']) => saw_int = true,
            Some(_) => saw_real = true,
            None => return ColType::Text,
        }
        if saw_int && saw_real {
            return ColType::Real;
        }
    }
    if !saw_any {
        return ColType::Text;
    }
    if saw_date {
        return ColType::Date;
    }
    if saw_int {
        ColType::Integer
    } else {
        ColType::Real
    }
}

/// Parses numbers the way humans type them in spreadsheets: allows thousands
/// separators and surrounding spaces; keeps sign and decimals. `001` still
/// parses, which is why ID-like columns must be pinned to TEXT in the preview
/// (documented behavior, not silent magic).
///
/// Comma handling is syntax-driven so both locales work without configuration:
/// - `1,234` / `1,234,567.50` → thousands separators (groups of exactly 3)
/// - `12,50` / `0,5` → decimal comma (European / Chinese-locale Excel exports
///   with `;` field separators). Previously this became 1250 — a 100× error.
fn parse_number_loose(t: &str) -> Option<f64> {
    let cleaned = normalise_number_text(t)?;
    if cleaned.is_empty() {
        return None;
    }
    cleaned.parse::<f64>().ok().filter(|v| v.is_finite())
}

/// Rewrites a numeric cell into ASCII `.`-decimal form with no separators.
/// Returns None for text that contains no digit.
fn normalise_number_text(t: &str) -> Option<String> {
    let t = t.trim().replace('\u{a0}', " ");
    let t = t.trim();
    if t.is_empty() || !t.chars().any(|c| c.is_ascii_digit()) {
        return None;
    }
    if !t.contains(',') {
        return Some(t.replace(' ', ""));
    }
    // Split off any exponent so the mantissa can be inspected alone.
    let (mantissa, exponent) = match t.find(['e', 'E']) {
        Some(i) if i > 0 && t[..i].chars().all(|c| c.is_ascii_digit() || matches!(c, '.' | ',' | '-' | '+')) => (&t[..i], t[i..].to_string()),
        _ => (t, String::new()),
    };
    let sign_len = usize::from(mantissa.starts_with(['-', '+']));
    let body = &mantissa[sign_len..];
    // Inspect comma groups on the integer part only, so `1,234,567.50`'s
    // last group is `567` (not `567.50`).
    let int_part = match body.rfind('.') {
        Some(d) if !body[..d].contains(',') || body[..d].rfind(',') < Some(d) => &body[..d],
        _ => body,
    };
    let comma_groups: Vec<&str> = int_part.split(',').collect();
    // `1,234` / `1,234,567` / `1,234,567.50`: every group after the first is
    // exactly three digits → thousands separators. Any other comma shape is a
    // decimal comma (`12,50`, `0,5`, `1,5e3`) or European dot-thousands
    // (`1.234,56`), where the dot comes BEFORE the comma.
    let thousands_shaped = comma_groups.len() > 1
        && !comma_groups[0].is_empty()
        && comma_groups[1..].iter().all(|g| g.len() == 3 && g.chars().all(|c| c.is_ascii_digit()))
        && body.rfind('.').map_or(true, |d| body.rfind(',').map_or(true, |c| c < d));
    let cleaned: String = if thousands_shaped {
        // strip commas, keep the dot
        format!("{}{}", &mantissa[..sign_len], body.replace(',', ""))
    } else if body.contains('.') {
        // European: dot = thousands, comma = decimal
        let (head, tail) = body.rsplit_once(',').unwrap_or((body, ""));
        format!("{}{}.{}", &mantissa[..sign_len], head.replace('.', "").replace(',', ""), tail.replace('.', "").replace(',', ""))
    } else {
        // lone decimal comma(s): `12,50` → `12.50`, `1,5` → `1.5`
        let (head, tail) = body.split_once(',').unwrap_or((body, ""));
        format!("{}{}.{}", &mantissa[..sign_len], head, tail.replace(',', ""))
    };
    let cleaned = cleaned.chars().filter(|c| *c != ' ').collect::<String>();
    Some(format!("{cleaned}{exponent}"))
}

/// Exact integer parse: parses digit strings directly as i64 without going
/// through f64, so values like 9007199254740993 keep full precision.
/// Returns None when the value is not an integer or overflows i64.
fn parse_integer_exact(t: &str) -> Option<i64> {
    let cleaned = normalise_number_text(t)?;
    // Reject decimal/exponent forms — those are not integers even if f64-round.
    if cleaned.contains(['.', 'e', 'E']) {
        return None;
    }
    cleaned.parse::<i64>().ok()
}

/// True when the integer text is within i64 range but cannot be represented
/// exactly (used to decide TEXT demotion instead of silent corruption).
fn integer_overflows_i64(t: &str) -> bool {
    let Some(cleaned) = normalise_number_text(t) else { return false };
    if cleaned.is_empty() || cleaned.contains(['.', 'e', 'E']) {
        return false;
    }
    cleaned.parse::<i128>().map(|v| v > i64::MAX as i128 || v < i64::MIN as i128).unwrap_or(true)
}

/// Recognizes common spreadsheet date/datetime shapes and returns them in
/// SQLite-preferred `YYYY-MM-DD[ HH:MM:SS]` form. DATE columns store as TEXT
/// in ISO form — sorts correctly, stays human-readable, round-trips.
pub fn parse_date_loose(t: &str) -> Option<String> {
    parse_date_loose_impl(t)
}

fn parse_date_loose_impl(t: &str) -> Option<String> {
    let t = t.trim();
    let len = t.len();
    if !(8..=32).contains(&len) {
        return None;
    }
    // Chinese office format: 2024年1月2日 or 2024年1月2日 14:30
    if let Some(day_pos) = t.find('日') {
        let (date_cn, tail) = t.split_at(day_pos + "日".len());
        let rest = date_cn.strip_suffix("日").unwrap_or(date_cn);
        if let Some(normalized) = normalize_cn_date_with_time(rest, tail) {
            return Some(normalized);
        }
        return None;
    }
    let all_ok = t.chars().all(|c| c.is_ascii_digit() || matches!(c, '-' | '/' | ':' | ' ' | '.' | 'T'));
    if !all_ok {
        return None;
    }
    let (date_part, time_part) = split_date_time(t)?;
    let (y, mo, d) = split_ymd(date_part)?;
    if !(1990..=2100).contains(&y) {
        return None;
    }
    if !(1..=12).contains(&mo) || !(1..=31).contains(&d) {
        return None;
    }
    let time = match time_part {
        Some(tp) => {
            let parts: Vec<u32> = tp.split(':').filter_map(|p| p.parse().ok()).collect();
            if parts.len() < 2 || parts.iter().any(|p| *p > 59) {
                return None;
            }
            let secs = parts.get(2).copied().unwrap_or(0);
            format!(" {:02}:{:02}:{:02}", parts[0], parts[1], secs)
        }
        None => String::new(),
    };
    Some(format!("{y:04}-{mo:02}-{d:02}{time}"))
}

/// "2024年1月2" (日 stripped by caller) -> normalized "2024-01-02" with an
/// optional trailing time "14:30" preserved as " 14:30:00".
fn normalize_cn_date_with_time(rest: &str, tail: &str) -> Option<String> {
    let time_part = tail.trim();
    let time_part = if time_part.is_empty() { None } else { Some(time_part) };
    let date_part = rest;
    let mut parts = date_part.split('年');
    let y: u32 = parts.next()?.parse().ok()?;
    let rest = parts.next()?;
    let mut sub = rest.split('月');
    let mo: u32 = sub.next()?.parse().ok()?;
    let d: u32 = sub.next()?.trim().parse().ok()?;
    if !(1990..=2100).contains(&y) || !(1..=12).contains(&mo) || !(1..=31).contains(&d) {
        return None;
    }
    let time = match time_part {
        Some(tp) => {
            let tp = tp.trim_start_matches([' ', ',']);
            let segs: Vec<u32> = tp.split(':').filter_map(|x| x.parse().ok()).collect();
            if segs.len() < 2 || segs.iter().any(|v| *v > 59) {
                return None;
            }
            format!(" {:02}:{:02}:{:02}", segs[0], segs[1], segs.get(2).copied().unwrap_or(0))
        }
        None => String::new(),
    };
    Some(format!("{y:04}-{mo:02}-{d:02}{time}"))
}

fn split_date_time(t: &str) -> Option<(&str, Option<&str>)> {
    if let Some(pos) = t.find(|c| c == ' ' || c == 'T') {
        // Reject "1 2 3"-style noise: only one separator allowed.
        if t[pos + 1..].contains(|c| c == ' ' || c == 'T') {
            return None;
        }
        Some((&t[..pos], Some(&t[pos + 1..])))
    } else {
        Some((t, None))
    }
}

fn split_ymd(date_part: &str) -> Option<(u32, u32, u32)> {
    let seps = date_part.contains('/') as u8 + date_part.contains('-') as u8 + date_part.contains('.') as u8;
    if seps > 1 {
        return None;
    }
    let parts: Vec<&str> = date_part.split(['-', '/', '.']).collect();
    if parts.len() != 3 {
        return None;
    }
    if parts.iter().any(|p| p.is_empty() || p.len() > 4) {
        return None;
    }
    let y: u32 = parts[0].parse().ok()?;
    let mo: u32 = parts[1].parse().ok()?;
    let d: u32 = parts[2].parse().ok()?;
    Some((y, mo, d))
}

/// Detects UTF-8 vs UTF-8-BOM vs UTF-16 (LE/BE with BOM) vs GBK fallback.
/// The BOM and UTF-16 cases are unambiguous; GBK is the pragmatic guess for
/// Chinese exports that are not UTF-8 — anything that still fails GBK decodes
/// with replacement so import never hard-fails on encoding.
pub fn sniff_and_decode(bytes: &[u8]) -> (String, &'static str) {
    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        return (String::from_utf8_lossy(&bytes[3..]).into_owned(), "utf-8-bom");
    }
    if bytes.starts_with(&[0xFF, 0xFE]) {
        let (cow, _, _) = encoding_rs::UTF_16LE.decode(&bytes[2..]);
        return (cow.into_owned(), "utf-16le");
    }
    if bytes.starts_with(&[0xFE, 0xFF]) {
        let (cow, _, _) = encoding_rs::UTF_16BE.decode(&bytes[2..]);
        return (cow.into_owned(), "utf-16be");
    }
    match std::str::from_utf8(bytes) {
        Ok(s) => (s.to_string(), "utf-8"),
        Err(_) => {
            let (cow, _, had_errors) = encoding_rs::GBK.decode(bytes);
            if had_errors {
                (String::from_utf8_lossy(bytes).into_owned(), "utf-8-lossy")
            } else {
                (cow.into_owned(), "gbk")
            }
        }
    }
}

/// Splits a decoded delimited text file into rows with quote handling.
///
/// The delimiter is chosen per file: `.tsv` forces a tab; otherwise the
/// first non-empty line decides between `,` `\t` `;` by which yields the most
/// fields (ties go to the comma). This is what makes European / Chinese-
/// locale Excel exports (`分隔符 = ;`) and "tab saved as .csv" files work
/// without the user having to rename or re-save anything.
pub fn split_csv_rows(text: &str) -> Vec<Vec<String>> {
    let delimiter = sniff_delimiter(text);
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .trim(csv::Trim::All)
        .delimiter(delimiter)
        .from_reader(text.as_bytes());
    rdr.records()
        .filter_map(|r| r.ok())
        .map(|r| r.iter().map(|c| c.to_string()).collect())
        .collect()
}

/// Picks `,` `\t` or `;` from the first few non-empty lines by field count.
pub fn sniff_delimiter(text: &str) -> u8 {
    let mut best = (b',', 1usize);
    for line in text.lines().filter(|l| !l.trim().is_empty()).take(8) {
        for cand in [b',', b'\t', b';'] {
            let n = line.matches(cand as char).count() + 1;
            if n > best.1 {
                best = (cand, n);
            }
        }
        // A clear winner on the first line is enough.
        if best.1 > 1 {
            break;
        }
    }
    best.0
}

/// `header_row_1based` is 1-based (1 = first row is the header). Rows above it
/// are title/banner rows and get skipped. 0 means "no header" (col_1..col_n).
pub fn parse_sheet_with_header(raw_rows: Vec<Vec<String>>, source_name: &str, header_row_1based: usize) -> SheetData {
    let skipped = header_row_1based.saturating_sub(1);
    // Keep the header row as the first element of `body`: skip only the
    // banner rows above it. header_row_1based == 0 consumes nothing.
    let body: Vec<Vec<String>> = raw_rows.into_iter().skip(skipped).collect();
    let header_mode = if header_row_1based == 0 { 0 } else { 1 };
    parse_sheet_body(body, header_mode, skipped, source_name)
}

/// Consumes the first body row as the header when `has_header_row` is true.
fn parse_sheet_body(body: Vec<Vec<String>>, header_mode: usize, skipped: usize, source_name: &str) -> SheetData {
    let mut rows_iter = body.into_iter().peekable();
    let headers_raw: Vec<String> = if header_mode == 0 {
        Vec::new()
    } else {
        rows_iter.next().unwrap_or_default()
    };
    let mut width = if header_mode == 0 {
        rows_iter.peek().map_or(headers_raw.len(), |r| r.len())
    } else {
        headers_raw.len().max(rows_iter.peek().map_or(0, |r| r.len()))
    };
    // Banner rows like "标题,,,," pad the header with empty cells; trailing
    // columns that are empty in the header AND every data row are noise.
    {
        let mut body_tail = |i: usize| rows_iter.peek().map_or(true, |r| r.get(i).map_or(true, |c| c.trim().is_empty()));
        while width > 0
            && headers_raw.get(width - 1).map_or(true, |h| h.trim().is_empty())
            && body_tail(width - 1)
        {
            width -= 1;
        }
    }
    if width == 0 {
        return SheetData {
            source_name: source_name.to_string(),
            label: sanitize_sheet_label(source_name),
            headers: vec![],
            rows: vec![],
            inferred_types: vec![],
            skipped_header_rows: skipped,
            hidden: false,
        };
    }
    let headers: Vec<String> = if headers_raw.is_empty() {
        (0..width).map(|i| format!("col_{}", i + 1)).collect()
    } else {
        (0..width)
            .map(|i| sanitize_column_name(headers_raw.get(i).map(String::as_str).unwrap_or(""), i))
            .collect()
    };
    let headers = super::names::dedupe_names(headers, "col");
    let data_rows: Vec<Vec<String>> = rows_iter
        .map(|mut r| {
            r.resize(width, String::new());
            r
        })
        .filter(|r| r.iter().any(|c| !c.trim().is_empty()))
        .collect();
    let scan = data_rows.iter().take(INFER_ROWS);
    let inferred_types: Vec<ColType> = (0..width)
        .map(|i| col_type_from_cells(scan.clone().map(|r| r[i].clone())))
        .collect();
    SheetData {
        source_name: source_name.to_string(),
        label: sanitize_sheet_label(source_name),
        headers,
        rows: data_rows,
        inferred_types,
        skipped_header_rows: skipped,
        hidden: false,
    }
}

/// Auto-detect entry: skip empty leading rows and title/banner rows (a row
/// with exactly one non-empty cell while a later row has several — "2024年
/// 销售汇总", "制表人：…", blank line, then the real header). The first row
/// with ≥ 2 non-empty cells is the header. A single-column file still works:
/// if NO row ever has ≥ 2 cells, the first non-empty row is the header.
pub fn parse_sheet(mut raw_rows: Vec<Vec<String>>, source_name: &str) -> SheetData {
    let non_empty = |r: &[String]| r.iter().filter(|c| !c.trim().is_empty()).count();
    let any_wide = raw_rows.iter().any(|r| non_empty(r) >= 2);
    let mut skipped = 0;
    while let Some(first) = raw_rows.first() {
        let n = non_empty(first);
        let is_banner = any_wide && n == 1 && raw_rows[1..].iter().any(|r| non_empty(r) >= 2);
        if n == 0 || is_banner {
            raw_rows.remove(0);
            skipped += 1;
        } else {
            break;
        }
    }
    parse_sheet_body(raw_rows, 1, skipped, source_name)
}

pub fn parse_csv_file(bytes: &[u8], source_name: &str) -> Result<SheetData, String> {
    let (text, _) = sniff_and_decode(bytes);
    let sheet = parse_sheet(split_csv_rows(&text), source_name);
    if sheet.headers.is_empty() {
        return Err("CSV 文件为空或没有有效内容".to_string());
    }
    Ok(sheet)
}


/// Sheet visibility straight from `xl/workbook.xml` (calamine 0.30 keeps
/// `Metadata.sheets` private). Returns names with state hidden/veryHidden.
fn hidden_sheet_names(bytes: &[u8]) -> std::collections::HashSet<String> {
    use std::collections::HashSet;
    use std::io::Read;
    let mut out = HashSet::new();
    let mut zip = match zip::ZipArchive::new(std::io::Cursor::new(bytes)) {
        Ok(z) => z,
        Err(_) => return out,
    };
    let mut xml = match zip.by_name("xl/workbook.xml") {
        Ok(f) => f,
        Err(_) => return out,
    };
    let mut buf = String::new();
    if xml.read_to_string(&mut buf).is_err() {
        return out;
    }
    // <sheet name="..." sheetId="1" state="hidden" r:id="rId1"/>
    for cap in regex_lite_regex_find(&buf) {
        out.insert(cap);
    }
    out
}

/// Minimal scan for hidden sheet declarations without a full XML parser:
/// finds `name="X"` on `<sheet ...>` tags that also carry a hidden state.
fn regex_lite_regex_find(xml: &str) -> Vec<String> {
    let mut names = Vec::new();
    let mut rest = xml;
    while let Some(pos) = rest.find("<sheet ") {
        let tag_end = match rest[pos..].find('>') {
            Some(e) => pos + e,
            None => break,
        };
        let tag = &rest[pos..tag_end];
        if tag.contains("state=\"hidden\"") || tag.contains("state=\"veryHidden\"") {
            if let Some(npos) = tag.find("name=\"") {
                let start = npos + 6;
                if let Some(epos) = tag[start..].find('"') {
                    names.push(tag[start..start + epos].to_string());
                }
            }
        }
        rest = &rest[tag_end..];
    }
    names
}

pub fn parse_workbook(bytes: &[u8], _source_name: &str) -> Result<Vec<SheetData>, String> {
    use calamine::Reader;
    let cursor = std::io::Cursor::new(bytes);
    let mut workbook = calamine::open_workbook_auto_from_rs(cursor)
        .map_err(|e| format!("无法读取表格文件: {e}"))?;
    let hidden_names = hidden_sheet_names(bytes);
    let is_1904 = workbook_is_1904(bytes);
    let mut sheets = Vec::new();
    for (name, range) in workbook.worksheets() {
        if range.height() == 0 && range.width() == 0 {
            continue;
        }
        let raw: Vec<Vec<String>> = range
            .rows()
            .map(|row| row.iter().map(|c| cell_to_text(c, is_1904)).collect())
            .collect();
        let mut sheet = parse_sheet(raw, &name);
        sheet.hidden = hidden_names.contains(&name);
        if !sheet.headers.is_empty() {
            sheets.push(sheet);
        }
    }
    if sheets.is_empty() {
        return Err("工作簿中没有包含数据的工作表".to_string());
    }
    Ok(sheets)
}

/// Renders one spreadsheet cell as the text the rest of the pipeline works
/// with. Dates are the subtle part: calamine's `ExcelDateTime` Display (with
/// the `dates` feature off) prints the raw serial number — 2024-03-01 would
/// become "45352" and be inferred as a number. Convert the serial to ISO
/// ourselves instead: pure date → `YYYY-MM-DD`, date+time →
/// `YYYY-MM-DD HH:MM:SS`, which the DATE inference and SQLite ordering both
/// expect.
fn cell_to_text(cell: &calamine::Data, is_1904: bool) -> String {
    match cell {
        calamine::Data::Int(i) => i.to_string(),
        calamine::Data::Float(f) => {
            // Trim float noise: 3.0 stays "3", 3.14 keeps precision.
            if f.fract() == 0.0 && f.abs() < 1e15 {
                format!("{}", *f as i64)
            } else {
                format!("{f}")
            }
        }
        calamine::Data::DateTime(dt) => excel_serial_to_iso(dt.as_f64(), is_1904),
        calamine::Data::DateTimeIso(x) => x.clone(),
        calamine::Data::DurationIso(x) => x.clone(),
        calamine::Data::String(x) => x.trim().to_string(),
        calamine::Data::Bool(b) => b.to_string(),
        other => other.to_string().trim().to_string(),
    }
}

/// Excel serial date (days since 1899-12-30 in the 1900 system, or since
/// 1904-01-01 in the 1904 system) → ISO text. The 1899-12-30 epoch already
/// accounts for Excel's phantom 1900-02-29 for every date after 1900-03-01;
/// dates in Jan–Feb 1900 are off by one, which no spreadsheet in practical
/// use has. 1904-workbook serials are shifted by the 1462 days between the
/// two epochs.
pub fn excel_serial_to_iso(serial: f64, is_1904: bool) -> String {
    if !serial.is_finite() || serial < 0.0 || serial > 3_000_000.0 {
        // Not a plausible date serial (corrupt cell or absurd value): keep the
        // number rather than inventing a year-8000 date.
        return format!("{serial}");
    }
    let days = serial.floor() as i64 + if is_1904 { 1462 } else { 0 };
    let secs = (((serial - serial.floor()) * 86_400.0).round() as i64).clamp(0, 86_399);
    let (y, m, d) = civil_from_days(days);
    if secs == 0 {
        format!("{y:04}-{m:02}-{d:02}")
    } else {
        format!("{y:04}-{m:02}-{d:02} {:02}:{:02}:{:02}", secs / 3600, (secs % 3600) / 60, secs % 60)
    }
}

/// Days since the Excel 1899-12-30 epoch → (year, month, day).
/// Howard Hinnant's civil-from-days algorithm; 1899-12-30 is day -25569
/// relative to the Unix epoch.
fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z - 25569; // Excel 1899-12-30 epoch → Unix 1970-01-01
    let z = z + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = (z - era * 146_097) as u64;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
}

/// True when the workbook uses the 1904 date system
/// (`<workbookPr date1904="1"/>` in `xl/workbook.xml`); rare, but Mac Excel
/// default for years and silently wrong by 4 years if ignored.
fn workbook_is_1904(bytes: &[u8]) -> bool {
    use std::io::Read;
    let mut zip = match zip::ZipArchive::new(std::io::Cursor::new(bytes)) {
        Ok(z) => z,
        Err(_) => return false,
    };
    let mut xml = match zip.by_name("xl/workbook.xml") {
        Ok(f) => f,
        Err(_) => return false,
    };
    let mut buf = String::new();
    if xml.read_to_string(&mut buf).is_err() {
        return false;
    }
    // Only the <workbookPr …/> element matters; date1904="1" / "true".
    buf.split('<').find(|s| s.starts_with("workbookPr")).is_some_and(|tag| {
        tag.split_once('>').map(|(t, _)| t).unwrap_or(tag).contains("date1904=\"1\"")
            || tag.contains("date1904=\"true\"")
    })
}


/// Per-column profile for the preview: null ratio, distinct count and up to
/// three sample values, so office users can see how the data was understood.
pub fn column_profiles(sheet: &SheetData) -> Vec<Value> {
    let n = sheet.rows.len().max(1);
    (0..sheet.headers.len())
        .map(|ci| {
            let mut non_empty = 0usize;
            let mut distinct = std::collections::HashSet::new();
            let mut samples: Vec<String> = Vec::new();
            for row in sheet.rows.iter().take(5000) {
                let v = row.get(ci).map(String::as_str).unwrap_or("").trim();
                if v.is_empty() {
                    continue;
                }
                non_empty += 1;
                distinct.insert(v.to_lowercase());
                if samples.len() < 3 && !samples.iter().any(|s| s == v) {
                    samples.push(v.chars().take(12).collect());
                }
            }
            json!({
                "name": sheet.headers[ci],
                "nullRatio": 1.0 - (non_empty as f64 / n as f64),
                "distinct": distinct.len(),
                "samples": samples,
            })
        })
        .collect()
}

/// The `analyze` RPC payload: per-sheet sample + inferred types.
pub fn sheet_to_json(sheet: &SheetData) -> Value {
    let sample: Vec<Vec<String>> = sheet.rows.iter().take(SAMPLE_ROWS).cloned().collect();
    json!({
        "sourceName": sheet.source_name,
        "label": sheet.label,
        "headers": sheet.headers,
        "rowCount": sheet.rows.len(),
        "inferredTypes": sheet.inferred_types.iter().map(|t| t.as_str()).collect::<Vec<_>>(),
        "sampleRows": sample,
        "skippedHeaderRows": sheet.skipped_header_rows,
        "hidden": sheet.hidden,
        "columnProfiles": column_profiles(sheet),
    })
}

/// Coerces one string cell into the user-confirmed column type for insertion.
pub fn coerce_cell(cell: &str, ty: ColType) -> rusqlite::types::Value {
    let t = cell.trim();
    if t.is_empty() {
        return rusqlite::types::Value::Null;
    }
    match ty {
        ColType::Text => rusqlite::types::Value::Text(t.to_string()),
        ColType::Integer => match parse_integer_exact(t) {
            Some(v) => rusqlite::types::Value::Integer(v),
            // Integer that overflows i64: keep the exact text rather than
            // corrupting it through f64 (9007199254740993 must survive).
            None if integer_overflows_i64(t) => rusqlite::types::Value::Text(t.to_string()),
            // Not an integer literal at all (e.g. "2.5"): f64 path decides.
            None => match parse_number_loose(t) {
                Some(v) if v.fract() == 0.0 && v.abs() < 9.3e18 => rusqlite::types::Value::Integer(v as i64),
                _ => rusqlite::types::Value::Text(t.to_string()),
            },
        },
        ColType::Real => match parse_number_loose(t) {
            Some(v) => rusqlite::types::Value::Real(v),
            None => rusqlite::types::Value::Text(t.to_string()),
        },
        ColType::Date => match parse_date_loose(t) {
            Some(iso) => rusqlite::types::Value::Text(iso),
            None => rusqlite::types::Value::Text(t.to_string()),
        },
    }
}


/// Re-parses already-analyzed sheets with explicit per-sheet header rows
/// (1-based; 0 = no header). Returns a fresh SheetData set; sheets whose
/// header row lands past the data just produce synthetic columns.
pub fn reparse_with_headers(
    path: &str,
    bytes: &[u8],
    header_rows: &[usize],
) -> Result<Vec<SheetData>, String> {
    let ext = std::path::Path::new(path)
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();
    if ext == "csv" || ext == "tsv" {
        let (text, _) = sniff_and_decode(bytes);
        let raw = split_csv_rows(&text);
        let stem = std::path::Path::new(path)
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("Sheet1")
            .to_string();
        let hr = header_rows.first().copied().unwrap_or(1);
        let sheet = parse_sheet_with_header(raw, &stem, hr);
        return Ok(vec![sheet]);
    }
    use calamine::Reader;
    let cursor = std::io::Cursor::new(bytes);
    let mut workbook = calamine::open_workbook_auto_from_rs(cursor)
        .map_err(|e| format!("无法读取表格文件: {e}"))?;
    let hidden_names = hidden_sheet_names(bytes);
    let is_1904 = workbook_is_1904(bytes);
    let mut sheets = Vec::new();
    for (idx, (name, range)) in workbook.worksheets().into_iter().enumerate() {
        if range.height() == 0 && range.width() == 0 {
            continue;
        }
        let raw: Vec<Vec<String>> = range
            .rows()
            .map(|row| row.iter().map(|c| cell_to_text(c, is_1904)).collect())
            .collect();
        let hr = header_rows.get(idx).copied().unwrap_or(1);
        let mut sheet = parse_sheet_with_header(raw, &name, hr);
        sheet.hidden = hidden_names.contains(&name);
        if !sheet.headers.is_empty() {
            sheets.push(sheet);
        }
    }
    if sheets.is_empty() {
        return Err("工作簿中没有包含数据的工作表".to_string());
    }
    Ok(sheets)
}

pub fn read_source_file(path: &str) -> Result<Vec<u8>, String> {
    let bytes = std::fs::read(path).map_err(|e| format!("无法读取文件 {path}: {e}"))?;
    if bytes.len() > 512 * 1024 * 1024 {
        return Err("文件过大（超过 512 MB）".to_string());
    }
    Ok(bytes)
}

pub fn parse_any(path: &str, bytes: &[u8]) -> Result<Vec<SheetData>, String> {
    let ext = std::path::Path::new(path)
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();
    let stem = std::path::Path::new(path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Sheet1")
        .to_string();
    match ext.as_str() {
        "csv" | "tsv" => parse_csv_file(bytes, &stem).map(|s| vec![s]),
        "xlsx" | "xlsm" | "xlsb" | "xls" => parse_workbook(bytes, path),
        _ => parse_workbook(bytes, path).or_else(|_| parse_csv_file(bytes, &stem).map(|s| vec![s])),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn infers_types() {
        assert_eq!(col_type_from_cells(["1", "2", ""].iter().map(|s| s.to_string())), ColType::Integer);
        assert_eq!(col_type_from_cells(["1", "2.5"].iter().map(|s| s.to_string())), ColType::Real);
        assert_eq!(
            col_type_from_cells(["1,234", "5,678"].iter().map(|s| s.to_string())),
            ColType::Integer
        );
        assert_eq!(col_type_from_cells(["a", "1"].iter().map(|s| s.to_string())), ColType::Text);
        assert_eq!(col_type_from_cells(["001", "002"].iter().map(|s| s.to_string())), ColType::Text);
        assert_eq!(col_type_from_cells(std::iter::empty::<String>()), ColType::Text);
    }

    #[test]
    fn decimal_comma_vs_thousands_comma() {
        // thousands style: groups of exactly three
        assert_eq!(parse_number_loose("1,234"), Some(1234.0));
        assert_eq!(parse_number_loose("1,234,567.50"), Some(1_234_567.50));
        assert_eq!(parse_number_loose("-1,234"), Some(-1234.0));
        // decimal comma (not 3-digit groups) — the 100× bug
        assert_eq!(parse_number_loose("12,50"), Some(12.5));
        assert_eq!(parse_number_loose("0,5"), Some(0.5));
        assert_eq!(parse_number_loose("7,99"), Some(7.99));
        // European "1.234,56" (dot thousands, comma decimal)
        assert_eq!(parse_number_loose("1.234,56"), Some(1234.56));
        // plain and exponent forms unchanged
        assert_eq!(parse_number_loose("3.14"), Some(3.14));
        assert_eq!(parse_number_loose("1e3"), Some(1000.0));
        assert_eq!(parse_number_loose("1,5e3"), Some(1500.0));
        assert_eq!(parse_number_loose("abc"), None);
        assert_eq!(parse_number_loose(""), None);
        // integers with thousands still parse exactly
        assert_eq!(parse_integer_exact("1,234,567"), Some(1_234_567));
        assert_eq!(parse_integer_exact("12,50"), None);
    }

    #[test]
    fn infers_dates() {
        assert_eq!(col_type_from_cells(["2024-01-02", "2024-03-04"].iter().map(|s| s.to_string())), ColType::Date);
        assert_eq!(col_type_from_cells(["2024/1/2", "2024/3/4"].iter().map(|s| s.to_string())), ColType::Date);
        assert_eq!(
            col_type_from_cells(["2024-01-02 08:30:00"].iter().map(|s| s.to_string())),
            ColType::Date
        );
        // mixed date + number -> text
        assert_eq!(col_type_from_cells(["2024-01-02", "42"].iter().map(|s| s.to_string())), ColType::Text);
    }

    #[test]
    fn parses_dates_to_iso() {
        assert_eq!(parse_date_loose("2024年1月2日").as_deref(), Some("2024-01-02"));
        assert_eq!(parse_date_loose("2024年12月31日").as_deref(), Some("2024-12-31"));
        assert_eq!(parse_date_loose("2024年1月2日 14:30").as_deref(), Some("2024-01-02 14:30:00"));
        assert_eq!(parse_date_loose("2024.1.2").as_deref(), Some("2024-01-02"));
        assert_eq!(parse_date_loose("2024.12.31").as_deref(), Some("2024-12-31"));
        assert_eq!(parse_date_loose("2024年13月1日"), None);
        assert_eq!(parse_date_loose("这不是日期日"), None);
        assert_eq!(parse_date_loose("2024-1-2").as_deref(), Some("2024-01-02"));
        assert_eq!(parse_date_loose("2024/01/02").as_deref(), Some("2024-01-02"));
        assert_eq!(parse_date_loose("2024-01-02 8:30").as_deref(), Some("2024-01-02 08:30:00"));
        assert_eq!(parse_date_loose("2024-13-01"), None);
        assert_eq!(parse_date_loose("123"), None);
        assert_eq!(parse_date_loose("hello"), None);
        assert_eq!(parse_date_loose("1800-01-01"), None);
    }

    #[test]
    fn decodes_encodings() {
        let (s, enc) = sniff_and_decode("a,b\n1,2".as_bytes());
        assert_eq!((s.as_str(), enc), ("a,b\n1,2", "utf-8"));
        let gbk = b"\xd6\xd0\xce\xc4"; // "中文" in GBK
        let (s, _) = sniff_and_decode(gbk);
        assert_eq!(s, "中文");
        let (s, _) = sniff_and_decode(&[0xFF, 0xFE, b'a', 0, b'b', 0]);
        assert_eq!(s, "ab");
    }

    #[test]
    fn parses_csv_rows_with_quotes() {
        let rows = split_csv_rows("a,b\n\"x,1\",2\n");
        assert_eq!(rows, vec![vec!["a", "b"], vec!["x,1", "2"]]);
    }

    #[test]
    fn sniffs_delimiter() {
        assert_eq!(sniff_delimiter("a,b,c\n1,2,3\n"), b',');
        assert_eq!(sniff_delimiter("name;price;qty\nWidget;12,50;3\n"), b';');
        assert_eq!(sniff_delimiter("a\tb\tc\n1\t2\t3\n"), b'\t');
        // a line with no delimiters at all defaults to comma
        assert_eq!(sniff_delimiter("just one cell\n"), b',');
        // comma wins ties (it is checked first and only strictly-greater wins)
        assert_eq!(sniff_delimiter("a;b,c\n"), b',');
    }

    #[test]
    fn splits_tsv_and_semicolon_csv() {
        let tsv = split_csv_rows("a\tb\n1\t2\n");
        assert_eq!(tsv, vec![vec!["a", "b"], vec!["1", "2"]]);
        let semi = split_csv_rows("name;price\nWidget;12,50\n");
        assert_eq!(semi, vec![vec!["name", "price"], vec!["Widget", "12,50"]]);
    }

    #[test]
    fn phone_numbers_infer_text() {
        assert_eq!(col_type_from_cells(["13800000001", "13800000002"].iter().map(|s| s.to_string())), ColType::Text);
        // ordinary quantities (≤ 6 digits) still become INTEGER
        assert_eq!(col_type_from_cells(["123456", "654321"].iter().map(|s| s.to_string())), ColType::Integer);
        // and so do 7+ digit numbers that carry thousands separators
        assert_eq!(col_type_from_cells(["12,345,678", "1,234,567"].iter().map(|s| s.to_string())), ColType::Integer);
    }

    #[test]
    fn excel_serials_render_as_iso_dates() {
        // Verified against Excel: 45352 = 2024-03-01, 45352.395833 ≈ 09:30:00
        assert_eq!(excel_serial_to_iso(45352.0, false), "2024-03-01");
        assert_eq!(excel_serial_to_iso(45352.3958333333, false), "2024-03-01 09:30:00");
        assert_eq!(excel_serial_to_iso(0.0, false), "1899-12-30"); // the epoch itself
        assert_eq!(excel_serial_to_iso(45292.0, false), "2024-01-01");
        assert_eq!(excel_serial_to_iso(45657.0, false), "2024-12-31");
        // Serials in Jan/Feb 1900 land one day late because Excel invents a
        // 1900-02-29; serial 61 (Excel's "1900-03-01") is where the two
        // calendars re-align. Nobody imports 1900 data, so this is accepted.
        assert_eq!(excel_serial_to_iso(61.0, false), "1900-03-01");
        // 1904 system: 0 = 1904-01-01
        assert_eq!(excel_serial_to_iso(0.0, true), "1904-01-01");
        // implausible serials are kept as numbers
        assert_eq!(excel_serial_to_iso(-5.0, false), "-5");
        assert_eq!(excel_serial_to_iso(f64::NAN, false), "NaN");
    }

    #[test]
    fn rendered_dates_infer_as_date_column() {
        // The regression: a date cell used to arrive as "45352" → REAL.
        let cells = [excel_serial_to_iso(45352.0, false), excel_serial_to_iso(45353.5, false)];
        assert_eq!(col_type_from_cells(cells.iter().map(|s| s.to_string())), ColType::Date);
    }

    #[test]
    fn long_integers_keep_exact_precision() {
        use rusqlite::types::Value as V;
        assert_eq!(coerce_cell("9007199254740993", ColType::Integer), V::Integer(9007199254740993));
        assert_eq!(coerce_cell("9223372036854775806", ColType::Integer), V::Integer(9223372036854775806));
        // thousands separators
        assert_eq!(coerce_cell("1,234,567,890", ColType::Integer), V::Integer(1234567890));
    }

    #[test]
    fn overflowing_integers_stay_text() {
        use rusqlite::types::Value as V;
        assert_eq!(coerce_cell("9223372036854775808", ColType::Integer), V::Text("9223372036854775808".into()));
        assert_eq!(coerce_cell("-9223372036854775809", ColType::Integer), V::Text("-9223372036854775809".into()));
    }

    #[test]
    fn leading_zeros_infer_text_at_inference_time() {
        // The reviewer's case: UI default plans echo inferredTypes verbatim,
        // so inference itself must classify 001-style columns as TEXT.
        assert_eq!(col_type_from_cells(["001", "002"].iter().map(|s| s.to_string())), ColType::Text);
        assert_eq!(col_type_from_cells(["0", "42"].iter().map(|s| s.to_string())), ColType::Integer);
    }

    #[test]
    fn coerce_respects_pinned_type() {
        use rusqlite::types::Value as V;
        assert_eq!(coerce_cell("001", ColType::Text), V::Text("001".into()));
        assert_eq!(coerce_cell("001", ColType::Integer), V::Integer(1));
        assert_eq!(coerce_cell("1,234", ColType::Integer), V::Integer(1234));
        assert_eq!(coerce_cell("", ColType::Integer), V::Null);
        assert_eq!(coerce_cell("abc", ColType::Integer), V::Text("abc".into()));
        assert_eq!(coerce_cell("2024-1-2", ColType::Date), V::Text("2024-01-02".into()));
        assert_eq!(coerce_cell("not a date", ColType::Date), V::Text("not a date".into()));
    }

    #[test]
    fn skips_empty_leading_rows() {
        let sheet = parse_sheet(
            vec![
                vec![String::new(), String::new()],
                vec!["h1".into(), "h2".into()],
                vec!["1".into(), String::new()],
                vec![String::new(), String::new()],
            ],
            "S",
        );
        assert_eq!(sheet.skipped_header_rows, 1);
        assert_eq!(sheet.headers, vec!["h1", "h2"]);
        assert_eq!(sheet.rows.len(), 1);
    }

    #[test]
    fn skips_title_banner_rows_before_header() {
        // "title" / "subtitle" / blank / header / data — header on row 4.
        let sheet = parse_sheet(
            vec![
                vec!["2024年销售汇总".into(), String::new(), String::new()],
                vec!["制表人：采购部".into(), String::new(), String::new()],
                vec![String::new(), String::new(), String::new()],
                vec!["单据编号".into(), "供应商".into(), "金额".into()],
                vec!["0001".into(), "华东物资".into(), "15200.75".into()],
            ],
            "S",
        );
        assert_eq!(sheet.skipped_header_rows, 3);
        assert_eq!(sheet.headers, vec!["单据编号", "供应商", "金额"]);
        assert_eq!(sheet.rows.len(), 1);
    }

    #[test]
    fn single_column_file_is_not_treated_as_banner() {
        // Every row has exactly one cell: the first one IS the header.
        let sheet = parse_sheet(
            vec![vec!["名称".into()], vec!["甲".into()], vec!["乙".into()]],
            "S",
        );
        assert_eq!(sheet.skipped_header_rows, 0);
        assert_eq!(sheet.headers, vec!["名称"]);
        assert_eq!(sheet.rows.len(), 2);
    }

    #[test]
    fn trailing_padded_columns_dropped() {
        // "标题,,,," pads to 5 cols; data only has 4 -> no col_5
        let sheet = parse_sheet_with_header(
            vec![
                vec!["2024年销售大表".into(), String::new(), String::new(), String::new(), String::new()],
                vec!["订单号".into(), "日期".into(), "金额".into(), "门店".into(), String::new()],
                vec!["A-1".into(), "2024-1-2".into(), "1500".into(), "北京".into(), String::new()],
            ],
            "S",
            2,
        );
        assert_eq!(sheet.headers, vec!["订单号", "日期", "金额", "门店"]);
    }

    #[test]
    fn honors_header_row_selection() {
        let raw = vec![
            vec!["2024销售大表".into(), String::new()],
            vec!["标题行".into(), String::new()],
            vec!["日期".into(), "金额".into()],
            vec!["2024-01-01".into(), "100".into()],
        ];
        // header on row 3 (1-based): rows 1-2 skipped as banner
        let sheet = parse_sheet_with_header(raw, "S", 3);
        assert_eq!(sheet.skipped_header_rows, 2);
        assert_eq!(sheet.headers, vec!["日期", "金额"]);
        assert_eq!(sheet.rows.len(), 1);
        assert_eq!(sheet.rows[0][0], "2024-01-01");
    }

    #[test]
    fn header_row_zero_generates_columns() {
        let raw = vec![vec!["a".into(), "1".into()], vec!["b".into(), "2".into()]];
        let sheet = parse_sheet_with_header(raw, "S", 0);
        assert_eq!(sheet.headers, vec!["col_1", "col_2"]);
        assert_eq!(sheet.rows.len(), 2);
    }
}
