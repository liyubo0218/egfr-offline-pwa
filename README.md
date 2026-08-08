# 基因筛选离线 PWA

这是一个无需服务器处理 Excel 的纯前端 PWA。Excel 内容始终保留在设备本机；首次成功打开后，可离线使用。

## 使用

1. 将整个项目上传到任何静态网站服务器，并通过 HTTPS 打开（Safari 的 Service Worker 与“添加到主屏幕”需要 HTTPS；`localhost` 也可用于开发）。
2. 在 iPhone Safari 中打开页面，点分享按钮，选择“添加到主屏幕”。
3. 选择一个 `.xlsx` 文件，勾选一个或多个基因，开始筛选。
4. 分别下载每个基因的结果 Excel，以及 `筛选总统计.xlsx`。

每个结果工作簿会保留找到的表头，写入符合规则的数据，启用自动筛选、首行冻结（支持该 Excel 视图属性的客户端会生效）并按中英文内容估算列宽。

## 基因配置

筛选逻辑不写在 JavaScript 中。每个基因拥有一个独立的 JSON，例如 `configs/egfr.json`：

```json
{
  "id": "egfr",
  "name": "EGFR",
  "outputFile": "EGFR.xlsx",
  "includeKeywords": ["关键词"],
  "excludeKeywords": ["排除关键词"]
}
```

筛选实现保留原始 `build_egfr.mjs` 的 `normalizeText()`、`displayText()`、`columnLetter()`、`findHeaderRow()`、`keywordMatch()`、`rowMatches()` 与 `nonEmptyRow()` 语义。它严格在前 20 行用完全相等的“印象”定位表头；`includeKeywords` 任一命中且 `excludeKeywords` 均不命中时才导出。每个工作表同时记录原脚本的表头行、印象列、扫描数量、包含命中、排除数量与最终命中。

### 新增基因

新增一个 JSON 文件（如 `configs/alk.json`），然后在 `configs/index.json` 添加一项：

```json
{ "id": "alk", "file": "alk.json" }
```

不需要修改任何 JavaScript。浏览器出于安全限制不能自动列出静态目录，因此索引文件是让离线 PWA 发现新配置所需的唯一登记处；同时请把新 JSON 加入 `service-worker.js` 的 `ASSETS` 列表后重新部署，确保它也可离线缓存。

## 项目结构

```
index.html              单页界面
style.css               iPhone 风格样式
app.js                  ES2022 应用逻辑
manifest.json           PWA 清单
service-worker.js       离线缓存
configs/                一个基因一个 JSON
vendor/xlsx.full.min.js SheetJS 0.18.5
icons/                  应用图标
```

## 命令行版本（build_gene.mjs）

除了浏览器 PWA，项目还内置了一个可离线运行的 Node.js 命令行脚本 `build_gene.mjs`，
筛选逻辑与 `app.js` 完全一致（同一套 `normalizeText`/`findHeaderRow`/`rowMatches` 等函数）。

```bash
node build_gene.mjs <源Excel路径> <基因配置json路径> [输出路径]

# 示例：
node build_gene.mjs ./样本.xlsx ./configs/egfr.json ./EGFR.xlsx
# 不传输出路径时，默认使用该基因 config 里的 outputFile（如 EGFR.xlsx）
```

行为：
- 输出 Workbook 只包含一个名为“筛选结果”的工作表。
- 该工作表里**只写入命中的行**；未命中（含被排除关键词命中）的行不会出现在输出文件中，
  不是先写全部数据再筛选，而是从未被写入过。
- 一次只处理一个基因的 config，不会读取或合并其它基因的数据。

## 本地预览

不要双击 `index.html` 直接以 `file://` 打开；应在项目目录启动任意静态 HTTP 服务后访问。发布到 HTTPS 后，PWA 安装和离线功能即可在 Safari、Chrome、Edge 运行。
