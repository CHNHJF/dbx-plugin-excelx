# ExcelX

**ExcelX** is a DBX plugin that converts spreadsheets (`.xlsx` / `.xlsm` / `.xlsb` / `.xls` / `.csv` / `.tsv`) into SQLite databases — and back.

**一个表格文件 = 一个数据库，一个 sheet = 一张表。**

## Why

Many office users don't know SQL, but DBX ships a built-in AI assistant. ExcelX closes the loop:

```
Spreadsheet ──(ExcelX)──► SQLite .db ──► DBX connection ──► ask the AI in plain language
     ◄──(ExcelX)──  export back to csv / xlsx  ◄──
```

- **Import** — pick a file, review inferred column types (text / integer / real / date), header row and primary key; the generated `.db` lands next to the source file and is registered into DBX's connection list automatically.
- **Export** — list ExcelX-generated SQLite connections (or any local `.db`) and export tables back to CSV (BOM, single table) or XLSX (multi-sheet).

## Highlights

- Encoding auto-detection: UTF-8 / UTF-8 BOM / UTF-16 LE+BE / GBK
- Type inference with **leading-zero protection** (`001` stays TEXT)
- Column profiles: fill ratio, distinct count, samples — inline under each header
- Empty-value (>30%) warnings before conversion
- Primary key defaults to the first column; duplicates fail with actionable messages
- Connections are real `db_type=sqlite` entries (full built-in workspace + AI), branded with an `ExcelX ` prefix and a tea-green `#1d6a40` color
- Rust sidecar: `calamine` + `csv` + `rusqlite(bundled)` + `rust_xlsxwriter`

## Build

```bash
npx @dbx-app/plugin-cli create excelx --template rust   # scaffold reference
# or build this repo directly:
cd backend && cargo build --release
# package (requires DBX_PLUGIN_SDK_ROOT pointing to the CLI's bundled SDK):
export DBX_PLUGIN_SDK_ROOT=<plugin-cli>/sdk-root
dbx-plugin package .
```

Releases are cut by GitHub Actions on published releases (see `.github/workflows/plugin-release.yml`).

## Install

DBX → Plugin Center → enable unsigned dev packages → install the `.dbxp` from [Releases](../../releases). Reopen the workbench tab after upgrading.

## License

Apache-2.0
