//! SQLite database construction from parsed sheets.

use rusqlite::Connection;
use serde_json::Value;

use super::analyze::{coerce_cell, ColType, SheetData};
use super::names::{quote_ident, unique_table_names};

/// User adjustments from the preview step: table name overrides, pinned
/// column types, skipped columns, and the primary key column (None = no PK).
#[derive(Debug, Clone)]
pub struct TablePlan {
    pub source_sheet: String,
    pub table_name: String,
    pub column_types: Vec<ColType>,
    pub skip_columns: Vec<usize>,
    pub primary_key: Option<usize>,
}

pub fn build_plans(sheets: &[SheetData], plans: &[Value]) -> Vec<(SheetData, TablePlan)> {
    // Sheet selection: the preview UI lets users uncheck sheets; only the
    // enabled subset becomes tables.
    let enabled: Vec<usize> = (0..sheets.len())
        .filter(|i| {
            let opt = plans.get(*i).and_then(|p| p.get("enabled")).and_then(Value::as_bool);
            // absent plan entry = default enabled; explicit false = skipped
            opt.unwrap_or(true)
        })
        .collect();
    let sheets: Vec<SheetData> = enabled.iter().map(|i| sheets[*i].clone()).collect();
    let plans: Vec<Value> = enabled.iter().map(|i| plans.get(*i).cloned().unwrap_or(Value::Null)).collect();
    let table_names = unique_table_names(&sheets.iter().map(|s| s.source_name.clone()).collect::<Vec<_>>());
    sheets
        .iter()
        .enumerate()
        .map(|(i, sheet)| {
            let plan = plans.get(i).cloned().unwrap_or(Value::Null);
            let width = sheet.headers.len();
            let mut column_types = sheet.inferred_types.clone();
            // Leading-zero guard (dialog quick path has no type preview):
            // if an inferred-INTEGER column's sampled values carry leading
            // zeros, pin TEXT — storing 001 as 1 silently loses data.
            for (j, ty) in column_types.iter_mut().enumerate() {
                if *ty == ColType::Integer {
                    let has_leading_zero = sheet
                        .rows
                        .iter()
                        .take(5000)
                        .any(|r| {
                            let v = r.get(j).map(String::as_str).unwrap_or("").trim();
                            v.len() > 1 && v.starts_with('0') && v.chars().all(|c| c.is_ascii_digit())
                        });
                    if has_leading_zero {
                        *ty = ColType::Text;
                    }
                }
            }
            if let Some(types) = plan.get("columnTypes").and_then(Value::as_array) {
                for (j, ty) in types.iter().enumerate().take(width) {
                    if let Some(t) = ty.as_str().and_then(ColType::from_str) {
                        column_types[j] = t;
                    }
                }
            }
            let skip_columns: Vec<usize> = plan
                .get("skipColumns")
                .and_then(Value::as_array)
                .map(|a| a.iter().filter_map(Value::as_u64).map(|v| v as usize).collect())
                .unwrap_or_default();
            // Default: NO primary key (safer for office data with messy
            // columns). "primaryKey": <index> opts in; the convert step then
            // pre-validates uniqueness and fails with the duplicate values.
            let primary_key = match plan.get("primaryKey") {
                Some(Value::Null) => None,
                Some(v) => match v.as_u64().map(|u| u as usize) {
                    Some(idx) if idx < width && !skip_columns.contains(&idx) => Some(idx),
                    _ => None,
                },
                None => None,
            };
            let table_name = plan
                .get("tableName")
                .and_then(Value::as_str)
                .map(|s| sanitize_table_name(s, &table_names[i]))
                .unwrap_or_else(|| table_names[i].clone());
            (
                sheet.clone(),
                TablePlan {
                    source_sheet: sheet.source_name.clone(),
                    table_name,
                    column_types,
                    skip_columns,
                    primary_key,
                },
            )
        })
        .collect()
}

/// Table-name override from the user: sanitized, falling back to the
/// auto-generated name when empty.
fn sanitize_table_name(raw: &str, fallback: &str) -> String {
    let cleaned = super::names::sanitize_column_name(raw, 0);
    if cleaned.is_empty() {
        fallback.to_string()
    } else {
        cleaned
    }
}

/// Writes the sheet into the database. Progress callback receives
/// (table_index, rows_done).
pub fn write_database(
    conn: &Connection,
    planned: &[(SheetData, TablePlan)],
    source_path: &str,
    mut progress: impl FnMut(usize, usize),
) -> Result<(), String> {
    conn.execute_batch("PRAGMA journal_mode = OFF; PRAGMA synchronous = OFF;")
        .map_err(|e| e.to_string())?;
    // PK uniqueness pre-check: fail before writing anything so the user gets
    // the offending values instead of a mid-import SQLite error.
    for (sheet, plan) in planned {
        if let Some(pk) = plan.primary_key {
            let mut seen = std::collections::HashMap::new();
            let mut dup: Vec<String> = Vec::new();
            for row in &sheet.rows {
                let v = row.get(pk).map(String::as_str).unwrap_or("").trim().to_string();
                if v.is_empty() {
                    continue;
                }
                *seen.entry(v.clone()).or_insert(0usize) += 1;
                if seen[&v] == 2 && dup.len() < 5 {
                    dup.push(v);
                }
            }
            if !dup.is_empty() {
                return Err(format!(
                    "列「{}」被设为主键但存在重复值（如：{}）。请先修正重复值，或把主键改为其他列/无主键。",
                    sheet.headers.get(pk).map(String::as_str).unwrap_or("?"),
                    dup.join("、")
                ));
            }
        }
    }
    for (idx, (sheet, plan)) in planned.iter().enumerate() {
        write_table(conn, sheet, plan)?;
        progress(idx + 1, sheet.rows.len());
    }
    let _ = source_path;
    Ok(())
}

fn write_table(conn: &Connection, sheet: &SheetData, plan: &TablePlan) -> Result<(), String> {
    let skip: std::collections::HashSet<usize> = plan.skip_columns.iter().copied().collect();
    let kept: Vec<usize> = (0..sheet.headers.len()).filter(|i| !skip.contains(i)).collect();
    if kept.is_empty() {
        return Err(format!("表 {} 的所有列都被跳过了", plan.table_name));
    }
    let cols: Vec<String> = kept.iter().map(|&i| quote_ident(&sheet.headers[i])).collect();
    let sql_types: Vec<&str> = kept.iter().map(|&i| plan.column_types[i].as_str()).collect();
    let ddl_cols: Vec<String> = cols
        .iter()
        .zip(sql_types.iter())
        .enumerate()
        .map(|(pos, (c, t))| {
            if plan.primary_key == kept.get(pos).copied() {
                format!("{c} {t} PRIMARY KEY")
            } else {
                format!("{c} {t}")
            }
        })
        .collect();
    conn.execute(
        &format!("CREATE TABLE {} ({});", quote_ident(&plan.table_name), ddl_cols.join(", ")),
        [],
    )
    .map_err(|e| format!("建表 {} 失败: {e}", plan.table_name))?;

    let placeholders = vec!["?"; kept.len()].join(",");
    let insert = format!(
        "INSERT INTO {} ({}) VALUES ({})",
        quote_ident(&plan.table_name),
        cols.join(","),
        placeholders
    );
    let mut stmt = conn
        .prepare(&insert)
        .map_err(|e| format!("准备写入 {} 失败: {e}", plan.table_name))?;
    for row in &sheet.rows {
        let values: Vec<rusqlite::types::Value> = kept
            .iter()
            .map(|&i| coerce_cell(&row[i], plan.column_types[i]))
            .collect();
        let params: Vec<&dyn rusqlite::ToSql> = values.iter().map(|v| v as &dyn rusqlite::ToSql).collect();
        if let Err(e) = stmt.execute(params.as_slice()) {
            return Err(format_insert_error(e, &plan.table_name, row));
        }
    }
    Ok(())
}

/// Turns raw SQLite constraint errors into actionable guidance: PK duplicates
/// name an offending value and point at the fix, so office users can correct
/// the source file instead of reading SQLite error codes.
fn format_insert_error(e: rusqlite::Error, table: &str, row: &[String]) -> String {
    let msg = e.to_string();
    if msg.contains("UNIQUE constraint failed") {
        let sample = row.first().map(String::as_str).unwrap_or("");
        format!(
            "表 {table} 的主键列存在重复值（例如「{sample}」）。请修正源文件中的重复值，或将主键改为其他列/设为无主键后重试。"
        )
    } else {
        format!("写入 {table} 失败: {e}")
    }
}

/// Lists user tables of any SQLite database (no marker requirement anymore:
/// the export side accepts every SQLite file the user picks).
pub fn list_tables(conn: &Connection) -> Result<Vec<(String, u64)>, String> {
    let mut stmt = conn
        .prepare("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
        .map_err(|e| e.to_string())?;
    let names: Vec<String> = stmt
        .query_map([], |r| r.get::<_, String>(0))
        .map_err(|e| e.to_string())?
        .filter_map(|r| r.ok())
        .collect::<Vec<_>>();
    let mut result = Vec::new();
    for name in names {
        let count: i64 = conn
            .query_row(&format!("SELECT COUNT(*) FROM {}", quote_ident(&name)), [], |r| r.get(0))
            .map_err(|e| e.to_string())?;
        result.push((name, count as u64));
    }
    Ok(result)
}

/// True when the file looks like a SQLite database (header magic).
pub fn looks_like_sqlite(path: &std::path::Path) -> bool {
    let mut magic = [0u8; 16];
    match std::fs::File::open(path).and_then(|mut f| std::io::Read::read_exact(&mut f, &mut magic)) {
        Ok(()) => &magic == b"SQLite format 3\0",
        Err(_) => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn sheet() -> SheetData {
        SheetData {
            source_name: "订单".into(),
            label: "订单".into(),
            headers: vec!["id".into(), "金额".into(), "日期".into(), "备注".into()],
            rows: vec![
                vec!["001".into(), "1,234.5".into(), "2024-1-2".into(), "a".into()],
                vec!["002".into(), "88".into(), "2024/1/3".into(), String::new()],
            ],
            inferred_types: vec![ColType::Text, ColType::Real, ColType::Date, ColType::Text],
            skipped_header_rows: 0,
        }
    }

    #[test]
    fn builds_database_with_pk_and_date() {
        let s = sheet();
        let plans = vec![json!({
            "tableName": "订单表",
            "columnTypes": ["TEXT", "REAL", "DATE", "TEXT"],
            "primaryKey": 0
        })];
        let planned = build_plans(std::slice::from_ref(&s), &plans);
        assert_eq!(planned[0].1.primary_key, Some(0));
        let conn = Connection::open_in_memory().unwrap();
        write_database(&conn, &planned, "C:/t/订单.xlsx", |_, _| {}).unwrap();

        let value: (String, f64, String) = conn
            .query_row(
                "SELECT id, 金额, 日期 FROM \"订单表\" WHERE id = '001'",
                [],
                |r| Ok((r.get(0)?, r.get(1)?, r.get(2)?)),
            )
            .unwrap();
        assert_eq!(value, ("001".to_string(), 1234.5, "2024-01-02".to_string()));

        let pk: String = conn
            .query_row("SELECT sql FROM sqlite_master WHERE name='订单表'", [], |r| r.get(0))
            .unwrap();
        assert!(pk.replace('"', "").contains("id TEXT PRIMARY KEY"), "{pk}");

        let tables = list_tables(&conn).unwrap();
        assert_eq!(tables, vec![("订单表".to_string(), 2)]);
    }

    #[test]
    fn leading_zero_integers_pinned_to_text() {
        let mut s = sheet();
        s.rows[0][0] = "001".into();
        s.rows[1][0] = "002".into();
        s.inferred_types[0] = ColType::Integer;
        let planned = build_plans(std::slice::from_ref(&s), &[]);
        assert_eq!(planned[0].1.column_types[0], ColType::Text);
        // plain integers stay INTEGER
        let s2 = sheet();
        let planned2 = build_plans(std::slice::from_ref(&s2), &[]);
        assert_eq!(planned2[0].1.column_types[1], ColType::Real);
    }

    #[test]
    fn pk_defaults_to_none_and_can_opt_in() {
        let s = sheet();
        // default: no PK
        assert_eq!(build_plans(std::slice::from_ref(&s), &[])[0].1.primary_key, None);
        assert_eq!(build_plans(std::slice::from_ref(&s), &[json!({})])[0].1.primary_key, None);
        let moved = build_plans(std::slice::from_ref(&s), &[json!({ "primaryKey": 2 })]);
        assert_eq!(moved[0].1.primary_key, Some(2));
        let off = build_plans(std::slice::from_ref(&s), &[json!({ "primaryKey": null })]);
        assert_eq!(off[0].1.primary_key, None);
        // skipped PK column -> no PK rather than a broken DDL
        let skipped = build_plans(std::slice::from_ref(&s), &[json!({ "skipColumns": [0], "primaryKey": 0 })]);
        assert_eq!(skipped[0].1.primary_key, None);
    }

    #[test]
    fn pk_uniqueness_prechecked_with_examples() {
        let mut s = sheet();
        s.rows[1][0] = "001".into();
        let planned = build_plans(std::slice::from_ref(&s), &[json!({ "primaryKey": 0 })]);
        let conn = Connection::open_in_memory().unwrap();
        let err = write_database(&conn, &planned, "x", |_, _| {}).unwrap_err();
        assert!(err.contains("重复值") && err.contains("001"), "{err}");
        // unique PK passes
        let ok = sheet();
        let planned2 = build_plans(std::slice::from_ref(&ok), &[json!({ "primaryKey": 0 })]);
        let conn2 = Connection::open_in_memory().unwrap();
        assert!(write_database(&conn2, &planned2, "x", |_, _| {}).is_ok());
    }

    #[test]
    fn disabled_sheets_are_skipped() {
        let a = sheet();
        let mut b = sheet();
        b.source_name = "第二张".into();
        b.label = "第二张".into();
        let planned = build_plans(&[a, b], &[json!({}), json!({ "enabled": false })]);
        assert_eq!(planned.len(), 1);
        assert_eq!(planned[0].1.table_name, "订单");
    }

    #[test]
    fn pk_duplicates_rejected_with_actionable_message() {
        let mut s = sheet();
        s.rows[1][0] = "001".into(); // duplicate PK value
        let planned = build_plans(std::slice::from_ref(&s), &[json!({ "primaryKey": 0 })]);
        let conn = Connection::open_in_memory().unwrap();
        let err = write_database(&conn, &planned, "x", |_, _| {}).unwrap_err();
        assert!(err.contains("被设为主键但存在重复值"), "{err}");
        assert!(err.contains("001"), "{err}");
    }

    #[test]
    fn skips_marked_columns() {
        let s = sheet();
        let plans = vec![json!({ "skipColumns": [3], "primaryKey": 0 })];
        let planned = build_plans(std::slice::from_ref(&s), &plans);
        let conn = Connection::open_in_memory().unwrap();
        write_database(&conn, &planned, "x", |_, _| {}).unwrap();
        let cols: i64 = conn
            .query_row("SELECT COUNT(*) FROM pragma_table_info('订单')", [], |r| r.get(0))
            .unwrap();
        assert_eq!(cols, 3);
    }

    #[test]
    fn detects_sqlite_magic() {
        let dir = std::env::temp_dir().join("excelx_magic");
        let _ = std::fs::create_dir_all(&dir);
        let f = dir.join("a.db");
        std::fs::write(&f, b"SQLite format 3\0rest").unwrap();
        assert!(looks_like_sqlite(&f));
        std::fs::write(&f, b"not a db").unwrap();
        assert!(!looks_like_sqlite(&f));
        let _ = std::fs::remove_file(&f);
    }
}
