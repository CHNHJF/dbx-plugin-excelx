# ExcelX

**ExcelX** 是一个 DBX 插件：把表格文件（`.xlsx` / `.xlsm` / `.xlsb` / `.xls` / `.csv` / `.tsv`）转换成 SQLite 数据库——再转换回来。

**一个表格文件 = 一个数据库，一个 sheet = 一张表。**

## 它解决什么问题

会用 Excel 的人多，会写 SQL 的人少。DBX 内置了 AI 助手，但 AI 需要一个数据库才能回答问题。ExcelX 补上这一环：

```
Excel / CSV ──(ExcelX)──► SQLite .db ──► 出现在 DBX 连接列表 ──► 用大白话向 AI 提问
                ◄──(ExcelX)──  随时导出回表格文件  ◄──
```

转换完成后，你可以直接问 DBX 的 AI 助手：

- 「上个月销售额最高的门店是哪家？」
- 「这张表有重复记录吗？」
- 「按部门汇总金额，列出前 5 名」

AI 会自动写 SQL、执行并给出答案——不需要懂任何数据库知识。

## 使用方法

### 导入：表格 → 数据库

1. 打开 DBX → 插件中心 → 已安装 → ExcelX → 打开工作台
2. 选择一个表格文件（多 sheet 的 Excel，每个 sheet 成为一张表；可以勾选只转换其中几个）
3. 检查预览：每列的类型（文本 / 整数 / 小数 / 日期）、表头所在行、是否需要主键（默认无主键；设置主键时会先校验值是否唯一）
4. 点击「转换为数据库」

生成的 `.db` 文件与原文件在同一文件夹，并自动加入 DBX 连接列表（带 `ExcelX ` 前缀和茶绿色标识），自动打开。之后就像使用任何数据库一样使用它：写 SQL、编辑数据、或者直接问 AI 助手。

### 导出：数据库 → 表格

在工作台切到「导出表格」页，选择 ExcelX 转换过的数据库（或本地任意 `.db` 文件），勾选要导出的表，点击导出——多张表会成为一个多 sheet 的 Excel 文件。修改过数据后再导出，就完成了「表格 → 数据库 → 表格」的完整回路。

## 特性

- **编码自动识别**：UTF-8 / UTF-8 BOM / UTF-16 / GBK，中文文件名与中文内容无忧
- **智能类型推断**：编号、手机号等有前导零的列自动保持文本（`001` 不会变成 `1`）；日期统一为可排序的 ISO 格式
- **列画像**：每列的有值率、不同取值数、样本值直接显示在表头下方
- **主键可选**：默认无主键；设置时会先校验唯一性，重复值会明确列出
- **纯本地**：转换全部在本机完成，数据不出电脑
- 生成的连接是标准 SQLite 连接——DBX 的全部能力（SQL 编辑器、数据编辑、AI 助手）开箱即用

## 平台支持

当前仅支持 **Windows x64**：文件选择对话框与 DBX 数据目录发现使用 Windows 机制（PowerShell + `%APPDATA%`）。macOS / Linux 支持在计划中。

## 安装

- **插件中心**（上架后）：DBX → Plugin Center 搜索 ExcelX 安装
- **手动**：从 [Releases](../../releases) 下载对应平台的 `.dbxp`，DBX → 插件中心 → 开启「允许安装未签名开发包」→ 安装

> 升级后请关闭旧的 ExcelX 工作台标签页再重新打开。

## 构建

```bash
# 打包（需要 DBX_PLUGIN_SDK_ROOT 指向 plugin-cli 自带的 SDK）
export DBX_PLUGIN_SDK_ROOT=<path-to-plugin-cli>/sdk-root
npx @dbx-app/plugin-cli package .
```

发布由 GitHub Actions 在 Release 发布时自动构建五平台包（见 `.github/workflows/plugin-release.yml`）。

## 许可

Apache-2.0
