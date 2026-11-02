# docx-model-v1

`office_ooxml_kit extract` 输出的块模型契约。字段名沿用旧 PowerShell 实现（原
`form_generate/tools/docx.ps1`，已归档到 `tools/reference/docx.ps1`）既有的
`docx-model-v1`，因此原有消费方（`cr-workflow.mjs`、`probe.mjs`、`outline.mjs`、历史
`runs/` 产物）不需要任何改动。

## 顶层

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `schemaVersion` | int | 恒为 `1` |
| `file` | string | 输入文件绝对路径 |
| `sha256` | string | 输入文件整包摘要（大写） |
| `metadata` | object | `{documentId, version, title}`，缺失为空串 |
| `blocks` | array | 块序列，见下 |
| `warnings` | array[string] | 需要人工确认的提示 |
| `ancillary` | array[object] | `{part, text}`，页眉页脚/脚注/批注的文字 |

## 块

所有块都有 `kind`、`text`、`locator`、`bodyIndex`、`section`、`sectionTitle`、
`category`。段落额外有 `number`、`style`；表格行额外有 `cells`、`cellImages`、
`spans`、`grid`、`layout`、`tableIndex`、`rowIndex`。

`kind` 取值：`paragraph`、`heading`、`tableRow`。

### section / sectionTitle

`section` 是**路径**，各级标题用 `' / '` 连接（例如 `5 详细设计 / 5.1 模块设计`）；
`sectionTitle` 是最内层标题文字。消费方按 `' / '` 拆分取祖先，分隔符是契约的一部分。

### category

| 值 | 含义 |
| --- | --- |
| `front` | 前置内容（封面、签署页） |
| `toc` | 目录项 |
| `revision` | 修订历史 |
| `body` | 正文 |

**注意一处刻意的不对称**：`section` 为「前置内容」的**段落**归类为 `body`，而同一章节下的
**表格行**归类为 `front`。`cr-workflow.mjs` 的封面处理依赖它。

### bodyIndex

`w:body` 下的**原始子元素序号**（含 `sectPr`），不是已提取块的计数。按块序号计数会让
表格行与段落的定位整体偏移。

### 单元格与图片

`cells[i]` 是该格可见文字：段落文字 + 图片 token，按出现顺序拼成。只有在既无文字
又无图片时才回退到自动编号。`cellImages[i]` 是同一格的图片数组。

图片身份是内容哈希 token `⟦图:ABCD1234⟧`（8 位十六进制），外链图片为 `⟦图:外链⟧`。
同一份字节在任何位置都得到同一个 token。token 会出现在 `cells[i]` 的文字里，因此
「换了一张图」能被纯文本比较发现。

图片对象的键名是契约的一部分：

| 键 | 说明 |
| --- | --- |
| `token` | `⟦图:XXXXXXXX⟧` 或 `⟦图:外链⟧` |
| `sha256` | 图片字节摘要（大写）；外链为空串 |
| `part` | 包内部件名；外链为 URL |
| `ext` / `bytes` | 扩展名 / 字节数 |
| `cx` / `cy` | EMU 宽高（`cx` 是宽，`cy` 是高） |
| `role` | `picture` 或 `olePreview` |
| `name` / `alt` | `docPr` 名称与替代文字 |
| `olePart` | OLE 对象部件名，非 OLE 为空串 |
| `external` | 仅外链图片出现，恒为 `true` |

`cx`/`cy` 用 EMU 而不是厘米，是因为消费方按 `cx/360000` 换算厘米；换算放在消费方，
组件只提供原始值。

`spans[i]` 是该格的 `gridSpan`（横向合并列数）；`rowIndex` 那行的 `vMerge` 通过
逐格 XML 判定，延续格不是独立可见单元格。

## 与 PowerShell 实现的差异

Python 额外输出 `style`（段落）、`vMerge`、`images`（段落图片）三个键，PowerShell 不输出。
消费方用 `?.` 与 `??` 取值，因此不受影响；新增键属于 v1 的**向后兼容扩展**。
