# 架构

## 一句话

Python + lxml 做所有 DOCX 读写的实际工作，暴露成 CLI；Node 项目用子进程调用它。
只有一个实现，两个项目共用。

## 为什么是 Python 而不是 Node

不是偏好问题，是**能力问题**：写回必须保住 38 个命名空间前缀，只有 lxml 能做到。

| 方案 | 读取 | 写回 | 结论 |
| --- | --- | --- | --- |
| Node + `docx4js` / 自写 DOM | 可行 | 需自己处理前缀 | 引入新依赖，且仍要自己踩坑 |
| `xml.etree.ElementTree` | 可行 | **丢 34 个前缀 → Word 报损坏** | 不可用 |
| Python + `python-docx` | 可行 | 受其 API 限制，写不进去的部分要绕过 | 只适合纯读写简单段落 |
| **Python + lxml** | 可行 | **保住全部 38 个前缀** | 采用 |

实测耗时（同一份 1.3 MB 规程）：

- Node 解析 JSON：0.029–0.089 s
- lxml 解析 `document.xml`：0.023 s
- `docx.ps1 extract` 一次：≈1.35 s（含 PowerShell 启动、.NET XML DOM、JSON 序列化）

也就是说，**Python 侧并不慢**，慢的是「为了一次解析而启动整个 PowerShell」。
改成 CLI 之后，Node 侧只付一次进程启动（约 0.1 s），比原来快一个数量级。

## 分层

```
office_ooxml_kit/
  ns.py         命名空间常量与 QName 工具；tostring 故意不清理命名空间
  package.py    zip 包读写，保序、保压缩方式、保时间戳；save 拒绝覆盖
  images.py     媒体清单、关系解析、图片嵌入（内容哈希即身份）
  numbering.py  自动编号重建（numId/ilvl 分别沿 basedOn 解析）
  docx_read.py  提取：块模型 + 警告 + 附属内容 + 元数据
  docx_write.py 填充：模板 + fill.json → 新文档
  verify.py     输出校验：单外层表、只改可改部分、页面几何、命名空间合法
  cli.py        extract / inspect / fill / verify
```

依赖方向是单向的：`docx_write` 依赖 `docx_read` 的模型，`verify` 只读包，不反向。

## 契约（docx-model-v1）

块字段冻结，因为两个项目的消费方已经在用：

```
kind text number style section sectionTitle category locator bodyIndex
tableIndex rowIndex cells cellImages spans grid layout images
```

图片字段同样冻结：`token sha256 part ext bytes cx cy role name alt olePart external`。

两条容易写错、且必须靠文档保住的语义：

1. **`section` 是路径**，用 `' / '` 连接 9 层标题；`sectionTitle` 是最内层标题。
2. **`category` 有意不对称**：`前置内容` 里的段落是 `body`，但 `前置内容` 下的表格行是 `front`。
   `cr-workflow.mjs` 的封面逻辑依赖这个不对称，改成"统一"会破坏封面识别。

## 三件必须共享的事

### 1. 命名空间：不清洗

```python
etree.tostring(node)   # 正确
etree.cleanup_namespaces(...)  # 禁止：会删掉 mc:Ignorable 按名字引用的前缀
```

新建节点时用显式 `nsmap`，避免 lxml 生成 `ns0`。

### 2. 取整：`round()` 而不是 `int()`

PowerShell 的 `[int]`/`[long]` 是**四舍六入五成双**，Python 的 `int()` 是**截断**。
表列宽、VML 尺寸换算（pt/in/cm → EMU）、图片绘制尺寸，只要有一处截断，
就会差 1 twip / 1 EMU。单看不明显，累积起来就是**排版差一页**（实测 19 页 vs 18 页）。

### 3. 单元格必须有段落结尾

Word 要求每个 `w:tc` 的**最后一个非 `tcPr` 子节点**是 `w:p`。
判据是"最后一个子节点是不是段落"，不是"里面有没有段落"——
单元格末尾是嵌套表格时，必须补一个空段落，否则文件损坏。

## 写回的完整性检查

`verify.check()` 输出四项，任何一项不过即视为失败：

| 检查 | 防的是什么 |
| --- | --- |
| `singleOuterTable` | 表头与内容被拆成两张表（原文档是一张） |
| `preserveOnlyPartChanges` | 改了不该改的部分（页眉页脚、样式、编号定义） |
| `sectionGeometryPreserved` | 页面尺寸/页边距被改动，导致重排 |
| `xmlNamespaceValid` | 命名空间前缀丢失或非法，Word 会判定损坏 |

`compare_parts` 只比较**两侧都存在**的部件：新增图片是"增加"，不是"改写"。
早期版本把 19 张新图片全判成"被改写"，是自检假阳性。

## 为什么不把 PowerShell 版删掉

`docx.ps1` 现在是回归基准。它的价值不是"还能用"，而是**它是另一个独立实现**：
两份实现逐块比对，不一致的地方必须逐条裁定。

本次比对 71 份文档、35,078 块，最终 9 处差异：

- 7 处是 PowerShell 漏了公式文字（Python 对，已用原始 XML 证明）；
- 2 处是 PowerShell 编号差一位（Python 对，已用 Word COM 的 `ListString` 证明）。

**"参考实现"不等于"正确实现"。** 保留它当对照物，比直接删掉更有用。
