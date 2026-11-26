# Changelog

## v0.9.9

### 修复

- **macOS 文件选择框里 CSV / XLSX 全部灰掉、无法选中**（#1）
  根因：`choose file of type` 用的是不存在的 UTI（`public.xlsx`、`public.csv`……），macOS 匹配不到任何文件；
  同时把 `alias "HD:Users:…"` 的 HFS 路径当 POSIX 路径用，即使选中了也打不开。
  现在映射到系统/Office 真实声明的 UTI（`public.comma-separated-values-text`、`org.openxmlformats.spreadsheetml.sheet`、`com.microsoft.excel.xls` …），
  并在 AppleScript 内部转成 POSIX path 返回。
- **Windows 文件选择框**（#1 跟进反馈「Excel 也不行」）：
  PowerShell 5.1 启动失败 / 被策略拦截时自动改用 PowerShell 7（`pwsh`）重试；
  彻底失败时不再卡死，错误信息直接指引改用备用选择器或粘贴路径。
- 表头行下拉只到「第 5 行」：现在可选 0–20 行，且第一行看起来是标题横幅时主动提示改表头。
- **Excel 里的日期单元格变成了 45352 这样的数字**（xlsx 用户必踩）：
  calamine 在未开启 `dates` feature 时把 Excel 序列号原样打印，`2024-03-01` 入库就成了 `45352`。
  现在自行把序列号转成 ISO（`YYYY-MM-DD` / `YYYY-MM-DD HH:MM:SS`），支持 1904 日期系统。
- **分号分隔的 CSV 里 `12,50` 被当成 1250**（100 倍错误）：
  欧式 / 中文 Excel「分隔符 = ;」导出的文件用逗号做小数点，旧逻辑一律当千分位剥掉。
  现在按语法判断：`1,234,567` 是千分位，`12,50` / `1.234,56` 是小数逗号。
- **TSV 没有按 Tab 拆列**：分隔符现在自动识别 `,` `\t` `;`，`.tsv` 强制 Tab。
- **11 位手机号被推断成整数**：≥ 7 位的纯数字串（手机 / 身份证 / 快递单号）现在保持文本。
- **第一行是标题横幅的表格被识别成 1 列**：自动跳过只有 1 个非空单元格、且后面有宽行的前导行作为标题 / 副标题。

### 新增

- **文件选择四条路，总有一条能用**：
  1. 系统对话框（原有）
  2. 「备用选择器」——走 DBX 宿主自带的 `fileTransfer` 文件选择器（新 RPC `excelx/ingestFile`，分片传到后端，乱序到达也能正确拼装）
  3. 把文件拖进工作台（宿主 `fileTransfer.onDrop`）
  4. 粘贴文件完整路径（Explorer「复制文件地址」带引号、带空格都行；路径是文件夹 / 文件不存在给出明确提示）
- **导出 CSV**：每张表一个 `.csv`（UTF-8 带 BOM，Excel 双击直接打开），没有 xlsx 的 1,048,576 行上限；导出页新增按钮。
- 拖入 / 备用选择器导入的文件，生成的 `.db` 放到「文档\ExcelX」而不是临时目录。

### 工程

- CI：新增 Windows 测试任务（之前只有 macOS），文件改名 `backend-tests.yml`。
- 发布流程 plugin-cli 固定到 0.1.10；README 构建说明去掉已不需要的 `DBX_PLUGIN_SDK_ROOT`。
- 后端测试 37 → 51 个；新增 10 个真实场景测试样例（编码 / 分隔符 / 标题横幅 / 多 sheet / 3000 行）做端到端回归。

## v0.9.8

- macOS support (Apple Silicon + Intel) joins windows-x64.
- macOS: system file picker (osascript), DBX data dir per ~/Library/Application Support, 'open' for folders
- Backend tests and handshake verified on macOS CI (arm64 & x64)
- Windows behavior unchanged from v0.9.7

## v0.9.7

- Real conversion progress (batched row events, host.events declared)
- Replace-rebuild: overwrite / new-copy / cancel prompt; self-overwrite guard
- Chinese date formats (2024年1月2日, 2024.3.4) parsed and inferred
- Exported numeric cells are real numbers
- Async TopMost file dialogs (no more 120s timeout, no sinking behind windows)
- Exact i64 integer parsing; leading-zero guard at inference
- >32,767-char cells truncated with warning; export lock detection with clear message
- PK duplicate errors carry row numbers; hidden sheets labeled and unchecked; wide-sheet preview cap

## v0.9.3 – v0.9.0

- Chinese store listing
- Exact integer parsing (no f64 round-trip); i64 overflow falls back to exact TEXT
- AI question suggestions exclude technical columns; multi-table suggestions
- Primary key defaults to none; per-sheet selection; export simplified to xlsx
