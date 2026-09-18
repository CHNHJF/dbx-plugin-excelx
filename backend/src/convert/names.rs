use std::collections::HashSet;

/// Escapes an identifier for SQLite DDL. Unusual-but-legal characters are kept
/// (double quotes doubled), so Chinese sheet headers stay readable.
pub fn quote_ident(name: &str) -> String {
    format!("\"{}\"", name.replace('"', "\"\""))
}

fn is_forbidden_sheet_char(c: char) -> bool {
    matches!(c, '\\' | '/' | '?' | '*' | '[' | ']') || c.is_control()
}

/// Truncates to Excel's 31-char sheet-name limit and strips characters Excel
/// forbids, so a table exported back to xlsx can reuse its original name.
pub fn sanitize_sheet_label(name: &str) -> String {
    let cleaned: String = name.chars().filter(|c| !is_forbidden_sheet_char(*c)).collect();
    let trimmed = cleaned.trim().trim_start_matches('\'');
    let mut label: String = trimmed.chars().take(31).collect();
    if label.is_empty() {
        label = "Sheet".to_string();
    }
    label
}

/// Normalizes a header cell into a table column name: non-empty, single-line.
/// Keeps Unicode letters/digits (Chinese headers stay usable in SQL), replaces
/// everything else with `_`.
pub fn sanitize_column_name(raw: &str, fallback_index: usize) -> String {
    let mut name = String::new();
    for c in raw.chars() {
        if c.is_alphanumeric() || c == '_' {
            name.push(c);
        } else if !name.ends_with('_') {
            name.push('_');
        }
    }
    let name = name.trim_matches('_').to_string();
    if name.is_empty() {
        return format!("col_{}", fallback_index + 1);
    }
    name
}

/// Makes a set of names unique (case-insensitively, matching SQLite's default
/// NOCASE-ish comparison for ASCII) by appending `_2`, `_3`, ...
pub fn dedupe_names(names: Vec<String>, fallback_prefix: &str) -> Vec<String> {
    let mut seen: HashSet<String> = HashSet::new();
    let mut result = Vec::with_capacity(names.len());
    for (i, name) in names.into_iter().enumerate() {
        let base = sanitize_column_name(&name, i).to_lowercase();
        let mut candidate = name;
        let mut counter = 2;
        while seen.contains(&candidate.to_lowercase()) {
            candidate = format!("{}_{}", base, counter);
            counter += 1;
        }
        if candidate.is_empty() {
            candidate = format!("{}_{}", fallback_prefix, i + 1);
        }
        seen.insert(candidate.to_lowercase());
        result.push(candidate);
    }
    result
}

/// Unique table names from sheet names, de-duplicated case-insensitively.
pub fn unique_table_names(sheet_labels: &[String]) -> Vec<String> {
    dedupe_names(sheet_labels.iter().map(|l| sanitize_column_name(l, 0)).collect(), "table")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn quotes_and_sanitizes() {
        assert_eq!(quote_ident("order\"s"), "\"order\"\"s\"");
        assert_eq!(sanitize_column_name("订单 明细", 0), "订单_明细");
        assert_eq!(sanitize_column_name("  ", 2), "col_3");
        assert_eq!(sanitize_column_name("2024销售额", 0), "2024销售额");
    }

    #[test]
    fn dedupes_case_insensitive() {
        assert_eq!(
            dedupe_names(vec!["a".into(), "A".into(), "a".into()], "t"),
            vec!["a", "a_2", "a_3"]
        );
    }

    #[test]
    fn sheet_label_limits() {
        let long = "x".repeat(40);
        assert_eq!(sanitize_sheet_label(&long).chars().count(), 31);
        assert_eq!(sanitize_sheet_label("a/b?c*d"), "abcd");
        assert_eq!(sanitize_sheet_label(""), "Sheet");
    }
}
