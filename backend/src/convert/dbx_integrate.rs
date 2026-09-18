//! DBX host integration: register the generated SQLite file as a real DBX
//! connection (db_type=sqlite, full built-in SQLite workspace), then tell the
//! running app to reload and open it.
//!
//! Path of least surprise, verified against dbx source:
//! - connections live in `%APPDATA%/com.dbx.app/dbx.db`, table `connections`
//!   (`id TEXT PRIMARY KEY, config_json TEXT`) — same store the MCP
//!   `dbx_add_connection` tool writes through storage.add_connection_for_mcp.
//! - the app keeps dbx.db in WAL mode; short-lived writes are safe, busy
//!   timeout rides out host checkpoints.
//! - new connections do NOT need sidebar_layout edits: reconcileLayout()
//!   appends unknown ids (sidebarLayout.ts).
//! - POST to the loopback MCP bridge (port in `mcp-bridge-port`) refreshes the
//!   UI (`/reload-connections`) and opens a query tab (`/open-table`).

use std::path::{Path, PathBuf};

use serde_json::{json, Value};

/// Prefix so ExcelX connections are recognizable in the sidebar at a glance
/// (db_type stays "sqlite" to keep the full built-in workspace).
/// Theme color: Excel brand green (#217346) darkened a step to a tea green —
/// deliberately NOT DBX's palette green (#22c55e).
pub const CONN_PREFIX: &str = "ExcelX ";
pub const CONN_COLOR: &str = "#1d6a40";

pub fn dbx_data_dir() -> Option<PathBuf> {
    let appdata = std::env::var("APPDATA").ok()?;
    let dir = Path::new(&appdata).join("com.dbx.app");
    dir.join("dbx.db").is_file().then_some(dir)
}

pub fn bridge_port(data_dir: &Path) -> Option<u16> {
    std::fs::read_to_string(data_dir.join("mcp-bridge-port"))
        .ok()?
        .trim()
        .parse()
        .ok()
}

/// Minimal POST with a tiny hand-rolled HTTP/1.1 write+read. The bridge never
/// keeps connections alive, so one shot per request is its expected usage.
fn bridge_post(port: u16, path: &str, body: &str) -> Result<String, String> {
    use std::io::{Read, Write};
    let addr = format!("127.0.0.1:{port}");
    let mut stream = std::net::TcpStream::connect(&addr)
        .map_err(|_| "无法连接 DBX（DBX 可能未运行）".to_string())?;
    let req = format!(
        "POST {path} HTTP/1.1\r\nHost: {addr}\r\nContent-Type: application/json\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{body}",
        body.len()
    );
    stream
        .write_all(req.as_bytes())
        .map_err(|e| format!("发送请求失败: {e}"))?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).map_err(|e| format!("读取响应失败: {e}"))?;
    let text = String::from_utf8_lossy(&buf).into_owned();
    if !text.starts_with("HTTP/1.1 200") {
        let body_start = text.split_once("\r\n\r\n").map(|(_, b)| b.to_string()).unwrap_or_default();
        return Err(format!("DBX 返回错误: {}", truncate(&body_start, 300)));
    }
    Ok(text.split_once("\r\n\r\n").map(|(_, b)| b.to_string()).unwrap_or_default())
}

fn truncate(s: &str, n: usize) -> String {
    s.chars().take(n).collect()
}

/// Collects config_json strings without borrowing the statement (statement
/// lifetimes cannot escape this function).
fn read_all_connection_configs(conn: &rusqlite::Connection) -> Result<Vec<String>, rusqlite::Error> {
    let mut stmt = conn.prepare("SELECT config_json FROM connections")?;
    let rows = stmt.query_map([], |r| r.get::<_, String>(0))?;
    Ok(rows.filter_map(|r| r.ok()).collect())
}

/// Registers `<name>` pointing at `db_path`. If a connection with the same
/// name exists, appends (2), (3)... — never overwrites, mirroring file output.
/// Returns (connection_id, final_name).
pub fn register_connection(db_path: &str, base_name: &str) -> Result<(String, String), String> {
    let data_dir = dbx_data_dir().ok_or_else(|| "未找到 DBX 数据目录（dbx.db）".to_string())?;
    let dbx_db = data_dir.join("dbx.db");
    let conn = rusqlite::Connection::open_with_flags(
        &dbx_db,
        rusqlite::OpenFlags::SQLITE_OPEN_READ_WRITE | rusqlite::OpenFlags::SQLITE_OPEN_NO_MUTEX,
    )
    .map_err(|e| format!("打开 DBX 配置库失败: {e}"))?;
    conn.busy_timeout(std::time::Duration::from_millis(3000))
        .map_err(|e| e.to_string())?;

    let existing = read_all_connection_configs(&conn).map_err(|e| format!("读取现有连接失败: {e}"))?;
    let existing_names: Vec<String> = existing
        .iter()
        .filter_map(|j| serde_json::from_str::<Value>(j).ok())
        .filter_map(|c| c.get("name").and_then(Value::as_str).map(String::from))
        .collect();
    let final_name = unique_connection_name(&format!("{CONN_PREFIX}{base_name}"), &existing_names);

    let id = new_uuid_v4();
    let config = json!({
        "id": id,
        "name": final_name,
        "db_type": "sqlite",
        "driver_profile": "sqlite",
        "driver_label": "SQLite",
        "url_params": "",
        "host": db_path,
        "port": 0,
        "username": "",
        "password": "",
        "database": null,
        "color": CONN_COLOR,
        "connect_timeout_secs": 5,
        "query_timeout_secs": 30,
        "idle_timeout_secs": 60,
        "keepalive_interval_secs": 60,
        "ssl": false,
        "sysdba": false,
        "connection_string": null,
        "external_config": null,
        "jdbc_driver_class": null,
        "jdbc_driver_paths": [],
        "save_password": true,
        "database_info": { "productName": "SQLite" }
    });
    let config_json = serde_json::to_string(&config).map_err(|e| e.to_string())?;
    conn.execute(
        "INSERT INTO connections (id, config_json) VALUES (?1, ?2)",
        rusqlite::params![id, config_json],
    )
    .map_err(|e| format!("写入连接失败: {e}"))?;
    Ok((id, final_name))
}

fn unique_connection_name(base: &str, existing: &[String]) -> String {
    let taken = |n: &str| existing.iter().any(|e| e.eq_ignore_ascii_case(n));
    if !taken(base) {
        return base.to_string();
    }
    for n in 2.. {
        let candidate = format!("{base} ({n})");
        if !taken(&candidate) {
            return candidate;
        }
    }
    unreachable!()
}

/// Poor-man's UUID v4 — random enough for a client-generated connection id,
/// same shape DBX itself uses (Uuid::new_v4).
fn new_uuid_v4() -> String {
    let mut bytes = [0u8; 16];
    if let Ok(mut f) = std::fs::File::open("\\\\.\\CNG\\") {
        // CryptGenRandom-backed RNG device; fall through to time-based on error.
        use std::io::Read;
        let _ = f.read_exact(&mut bytes);
    }
    if bytes == [0u8; 16] {
        let t = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        for (i, b) in bytes.iter_mut().enumerate() {
            *b = ((t >> (i * 8)) & 0xff) as u8 ^ (i as u8).wrapping_mul(31);
        }
    }
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    format!(
        "{:02x}{:02x}{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
        bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
        bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15]
    )
}

/// Tells the running DBX to re-read connections from disk (sidebar updates).
pub fn notify_reload() -> Result<(), String> {
    let dir = dbx_data_dir().ok_or_else(|| "未找到 DBX 数据目录".to_string())?;
    let port = bridge_port(&dir).ok_or_else(|| "DBX 未运行或桥不可用".to_string())?;
    bridge_post(port, "/reload-connections", "")?;
    Ok(())
}

/// Opens the connection in DBX: the host emits mcp-open-table which connects
/// and opens a query tab. Table is optional (None = just open the connection).
pub fn open_in_dbx(connection_id: &str, table: Option<&str>) -> Result<(), String> {
    let dir = dbx_data_dir().ok_or_else(|| "未找到 DBX 数据目录".to_string())?;
    let port = bridge_port(&dir).ok_or_else(|| "DBX 未运行或桥不可用".to_string())?;
    // database "main" targets SQLite's single catalog node so the sidebar
    // focuses on the database level (not just the connection root).
    let body = json!({
        "connection_id": connection_id,
        "connection_name": "",
        "database": "main",
        "table": table.unwrap_or(""),
    })
    .to_string();
    bridge_post(port, "/open-table", &body)?;
    Ok(())
}

/// All SQLite connections from the real DBX connection list (no filter):
/// the export tab shows exactly what the sidebar shows. Each entry carries
/// whether the file still exists.
pub fn list_all_sqlite_connections() -> Vec<Value> {
    let Some(dir) = dbx_data_dir() else { return vec![] };
    let conn = match rusqlite::Connection::open_with_flags(
        dir.join("dbx.db"),
        rusqlite::OpenFlags::SQLITE_OPEN_READ_ONLY,
    ) {
        Ok(c) => c,
        Err(_) => return vec![],
    };
    let rows = read_all_connection_configs(&conn).unwrap_or_default();
    rows.iter()
        .filter_map(|j| serde_json::from_str::<Value>(j).ok())
        .filter(|c| {
            let is_ours = |c: &Value| {
                let name = c.get("name").and_then(Value::as_str).unwrap_or("");
                let color = c.get("color").and_then(Value::as_str).unwrap_or("");
                name.starts_with(CONN_PREFIX) || color == CONN_COLOR
            };
            c.get("db_type").and_then(Value::as_str) == Some("sqlite") && is_ours(c)
        })
        .filter_map(|c| {
            let host = c.get("host").and_then(Value::as_str)?;
            let path = Path::new(host);
            Some(json!({
                "id": c.get("id").cloned().unwrap_or(Value::Null),
                "name": c.get("name").cloned().unwrap_or(Value::Null),
                "dbPath": host,
                "exists": path.is_file(),
            }))
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn uniquifies_connection_names() {
        assert_eq!(unique_connection_name("ExcelX a", &[]), "ExcelX a");
        assert_eq!(unique_connection_name("ExcelX a", &["ExcelX a".into()]), "ExcelX a (2)");
        assert_eq!(
            unique_connection_name("ExcelX a", &["ExcelX a".into(), "ExcelX a (2)".into()]),
            "ExcelX a (3)"
        );
        // case-insensitive collision
        assert_eq!(unique_connection_name("ExcelX A", &["ExcelX a".into()]), "ExcelX A (2)");
    }

    #[test]
    fn uuid_shape() {
        let id = new_uuid_v4();
        assert_eq!(id.len(), 36);
        assert_eq!(id.as_bytes()[14], b'4');
        assert!(matches!(id.as_bytes()[19], b'8'..=b'b' | b'9' | b'a' | b'b'));
    }
}
