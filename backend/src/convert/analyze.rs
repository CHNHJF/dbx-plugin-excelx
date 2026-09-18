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
fn parse_number_loose(t: &str) -> Option<f64> {
    let cleaned: String = t.chars().filter(|c| *c != ',' && *c != '\u{a0}').collect();
    if cleaned.is_empty() {
        return None;
    }
    cleaned.parse::<f64>().ok().filter(|v| v.is_finite())
}

/// Recognizes common spreadsheet date/datetime shapes and returns them in
/// SQLite-preferred `YYYY-MM-DD[ HH:MM:SS]` form. DATE columns store as TEXT
/// in ISO form — sorts correctly, stays human-readable, round-trips.
pub fn parse_date_loose(t: &str) -> Option<String> {
    let t = t.trim();
    let len = t.len();
    if !(8..=26).contains(&len) {
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
    if date_part.contains('/') && date_part.contains('-') {
        return None;
    }
    let parts: Vec<&str> = date_part.split(['-', '/']).collect();
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

/// Splits a decoded CSV into rows with quote handling.
pub fn split_csv_rows(text: &str) -> Vec<Vec<String>> {
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .trim(csv::Trim::All)
        .from_reader(text.as_bytes());
    rdr.records()
        .filter_map(|r| r.ok())
        .map(|r| r.iter().map(|c| c.to_string()).collect())
        .collect()
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
    }
}

/// Legacy auto-detect entry: skip empty leading rows, first non-empty row is
/// the header.
pub fn parse_sheet(mut raw_rows: Vec<Vec<String>>, source_name: &str) -> SheetData {
    let mut skipped = 0;
    while let Some(first) = raw_rows.first() {
        if first.iter().all(|c| c.trim().is_empty()) {
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

pub fn parse_workbook(bytes: &[u8], _source_name: &str) -> Result<Vec<SheetData>, String> {
    use calamine::Reader;
    let cursor = std::io::Cursor::new(bytes);
    let mut workbook = calamine::open_workbook_auto_from_rs(cursor)
        .map_err(|e| format!("无法读取表格文件: {e}"))?;
    let mut sheets = Vec::new();
    for (name, range) in workbook.worksheets() {
        if range.height() == 0 && range.width() == 0 {
            continue;
        }
        let raw: Vec<Vec<String>> = range
            .rows()
            .map(|row| {
                row.iter()
                    .map(|cell| match cell {
                        calamine::Data::Int(i) => i.to_string(),
                        calamine::Data::Float(f) => {
                            // Trim float noise: 3.0 stays "3", 3.14 keeps precision.
                            if f.fract() == 0.0 && f.abs() < 1e15 {
                                format!("{}", *f as i64)
                            } else {
                                format!("{f}")
                            }
                        }
                        calamine::Data::DateTime(dt) => dt.to_string(),
                        calamine::Data::DateTimeIso(s) => s.clone(),
                        calamine::Data::DurationIso(s) => s.clone(),
                        calamine::Data::String(s) => s.trim().to_string(),
                        calamine::Data::Bool(b) => b.to_string(),
                        other => other.to_string().trim().to_string(),
                    })
                    .collect()
            })
            .collect();
        let sheet = parse_sheet(raw, &name);
        if !sheet.headers.is_empty() {
            sheets.push(sheet);
        }
    }
    if sheets.is_empty() {
        return Err("工作簿中没有包含数据的工作表".to_string());
    }
    Ok(sheets)
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
        ColType::Integer => match parse_number_loose(t) {
            Some(v) if v.fract() == 0.0 && v.abs() < 9.3e18 => rusqlite::types::Value::Integer(v as i64),
            _ => rusqlite::types::Value::Text(t.to_string()),
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
    let mut sheets = Vec::new();
    for (idx, (name, range)) in workbook.worksheets().into_iter().enumerate() {
        if range.height() == 0 && range.width() == 0 {
            continue;
        }
        let raw: Vec<Vec<String>> = range
            .rows()
            .map(|row| {
                row.iter()
                    .map(|cell| match cell {
                        calamine::Data::Int(i) => i.to_string(),
                        calamine::Data::Float(f) => {
                            if f.fract() == 0.0 && f.abs() < 1e15 {
                                format!("{}", *f as i64)
                            } else {
                                format!("{f}")
                            }
                        }
                        calamine::Data::DateTime(dt) => dt.to_string(),
                        calamine::Data::DateTimeIso(x) => x.clone(),
                        calamine::Data::DurationIso(x) => x.clone(),
                        calamine::Data::String(x) => x.trim().to_string(),
                        calamine::Data::Bool(b) => b.to_string(),
                        other => other.to_string().trim().to_string(),
                    })
                    .collect()
            })
            .collect();
        let hr = header_rows.get(idx).copied().unwrap_or(1);
        let sheet = parse_sheet_with_header(raw, &name, hr);
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
        assert_eq!(col_type_from_cells(["001", "002"].iter().map(|s| s.to_string())), ColType::Integer);
        assert_eq!(col_type_from_cells(std::iter::empty::<String>()), ColType::Text);
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
