//! Export any SQLite database back to csv / xlsx.

use rusqlite::Connection;

use super::names::{quote_ident, sanitize_sheet_label};

pub const XLSX_MAX_ROWS: usize = 1_048_576;

fn read_table(conn: &Connection, table: &str, max_rows: usize) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    let mut cols_stmt = conn
        .prepare(&format!("SELECT name FROM pragma_table_info({})", quote_ident(table)))
        .map_err(|e| e.to_string())?;
    let headers: Vec<String> = cols_stmt
        .query_map([], |r| r.get(0))
        .map_err(|e| e.to_string())?
        .filter_map(|r| r.ok())
        .collect();
    let select = format!(
        "SELECT {} FROM {}",
        headers.iter().map(|h| quote_ident(h)).collect::<Vec<_>>().join(", "),
        quote_ident(table)
    );
    let mut stmt = conn.prepare(&select).map_err(|e| e.to_string())?;
    let n_cols = headers.len();
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut rs = stmt.query([]).map_err(|e| format!("读取表 {table} 失败: {e}"))?;
    while let Some(row) = rs.next().map_err(|e| e.to_string())? {
        if rows.len() >= max_rows {
            break;
        }
        let mut out = Vec::with_capacity(n_cols);
        for i in 0..n_cols {
            let v = row.get_ref(i).map_err(|e| e.to_string())?;
            out.push(match v {
                rusqlite::types::ValueRef::Null => String::new(),
                rusqlite::types::ValueRef::Integer(i) => i.to_string(),
                rusqlite::types::ValueRef::Real(f) => {
                    if f.fract() == 0.0 && f.abs() < 1e15 {
                        format!("{}", f as i64)
                    } else {
                        format!("{f}")
                    }
                }
                rusqlite::types::ValueRef::Text(t) => String::from_utf8_lossy(t).into_owned(),
                rusqlite::types::ValueRef::Blob(b) => format!("<{} bytes>", b.len()),
            });
        }
        rows.push(out);
    }
    Ok((headers, rows))
}

pub fn export_csv(conn: &Connection, table: &str, max_rows: usize) -> Result<String, String> {
    let (headers, rows) = read_table(conn, table, max_rows)?;
    let mut wtr = csv::Writer::from_writer(vec![]);
    wtr.write_record(&headers).map_err(|_| "写入表头失败".to_string())?;
    for r in rows {
        wtr.write_record(&r).map_err(|_| "写入行失败".to_string())?;
    }
    let bytes = wtr.into_inner().map_err(|_| "完成 CSV 失败".to_string())?;
    Ok(String::from_utf8_lossy(&bytes).into_owned())
}

/// BOM so Excel opens UTF-8 CSV with Chinese text correctly on double-click.
pub fn csv_with_bom(csv: &str) -> Vec<u8> {
    let mut v = Vec::with_capacity(3 + csv.len());
    v.extend_from_slice(&[0xEF, 0xBB, 0xBF]);
    v.extend_from_slice(csv.as_bytes());
    v
}

pub fn export_xlsx(conn: &Connection, tables: &[String], max_rows: usize) -> Result<Vec<u8>, String> {
    let mut wb = rust_xlsxwriter::Workbook::new();
    let mut used_labels: std::collections::HashSet<String> = std::collections::HashSet::new();
    for table in tables {
        let (headers, rows) = read_table(conn, table, max_rows)?;
        let mut label = sanitize_sheet_label(table);
        let mut n = 2;
        while used_labels.contains(&label) {
            label = sanitize_sheet_label(&format!("{}_{n}", table));
            n += 1;
        }
        used_labels.insert(label.clone());
        let ws = wb.add_worksheet().set_name(&label).map_err(|e| format!("工作表命名失败: {e}"))?;
        for (ci, h) in headers.iter().enumerate() {
            ws.write_string(0, ci as u16, h)
                .map_err(|e| format!("写入表头失败: {e}"))?;
        }
        for (ri, row) in rows.iter().enumerate() {
            for (ci, cell) in row.iter().enumerate() {
                ws.write_string((ri + 1) as u32, ci as u16, cell)
                    .map_err(|e| format!("写入单元格失败: {e}"))?;
            }
        }
    }
    wb.save_to_buffer().map_err(|e| format!("生成 xlsx 失败: {e}"))
}

/// CSV-specific problem scan for the pre-export check: actionable findings
/// for content that cannot round-trip cleanly through CSV.
pub fn csv_export_warnings(conn: &Connection, table: &str) -> Vec<String> {
    let mut warnings = Vec::new();
    let (headers, _rows) = match read_table(conn, table, 1) {
        Ok(v) => v,
        Err(e) => {
            warnings.push(e);
            return warnings;
        }
    };
    let dup_headers = duplicate_elements(&headers);
    if !dup_headers.is_empty() {
        warnings.push(format!(
            "存在重复列名（{}），CSV 第一行表头会无法区分；建议改导出 xlsx，或先在表中重命名列",
            dup_headers.join("、")
        ));
    }
    warnings
}

fn duplicate_elements(items: &[String]) -> Vec<String> {
    // PRAGMA table_info renames duplicate result columns to "col", "col:1",
    // so compare after stripping a trailing ":<digits>" suffix as well.
    let base = |s: &str| {
        let lower = s.to_lowercase();
        match lower.rsplit_once(':') {
            Some((head, digits)) if !digits.is_empty() && digits.chars().all(|c| c.is_ascii_digit()) => head.to_string(),
            _ => lower,
        }
    };
    let mut seen = std::collections::HashSet::new();
    let mut dups = std::collections::HashSet::new();
    for i in items {
        if !seen.insert(base(i)) {
            dups.insert(i.clone());
        }
    }
    dups.into_iter().collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::convert::analyze::{col_type_from_cells, ColType, SheetData};
    use crate::convert::build::{build_plans, write_database};

    fn db_with_order_table() -> Connection {
        let sheet = SheetData {
            source_name: "订单".into(),
            label: "订单".into(),
            headers: vec!["编号".into(), "金额".into()],
            rows: vec![
                vec!["001".into(), "1234.5".into()],
                vec!["002".into(), "88".into()],
            ],
            inferred_types: vec![ColType::Text, ColType::Real],
            skipped_header_rows: 0,
        };
        let planned = build_plans(std::slice::from_ref(&sheet), &[]);
        let conn = Connection::open_in_memory().unwrap();
        write_database(&conn, &planned, "t.xlsx", |_, _| {}).unwrap();
        conn
    }

    #[test]
    fn roundtrips_csv() {
        let conn = db_with_order_table();
        let csv = export_csv(&conn, "订单", 1000).unwrap();
        assert!(csv.contains("编号,金额"));
        assert!(csv.contains("001,1234.5"));
        assert!(csv_with_bom(&csv).starts_with(&[0xEF, 0xBB, 0xBF]));
    }

    #[test]
    fn roundtrips_xlsx_bytes() {
        let conn = db_with_order_table();
        let bytes = export_xlsx(&conn, &["订单".to_string()], 1000).unwrap();
        assert_eq!(&bytes[..2], b"PK"); // xlsx is a zip
        assert!(bytes.len() > 500);
    }

    #[test]
    fn respects_max_rows() {
        let conn = db_with_order_table();
        let csv = export_csv(&conn, "订单", 1).unwrap();
        assert_eq!(csv.lines().count(), 2); // header + 1 row
    }

    #[test]
    fn duplicate_headers_flagged_for_csv() {
        // Physical tables cannot have duplicate column names (SQLite rejects
        // the DDL), but views/selects can produce them — guard that case.
        let conn = Connection::open_in_memory().unwrap();
        conn.execute_batch("CREATE TABLE t (a TEXT, b TEXT); INSERT INTO t VALUES ('x','y'); CREATE VIEW v AS SELECT a AS col, b AS col FROM t;").unwrap();
        assert!(csv_export_warnings(&conn, "t").is_empty());
        let w = csv_export_warnings(&conn, "v");
        assert_eq!(w.len(), 1);
        assert!(w[0].contains("重复列名"), "{w:?}");
    }

    #[test]
    fn infer_matches_export_formatting() {
        assert_eq!(col_type_from_cells(["1", "2"].iter().map(|s| s.to_string())), ColType::Integer);
    }
}
