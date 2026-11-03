# office_ooxml_kit

通用 Office 文档操作组件（仓库名 `office-ooxml-kit`）。把「填一份带格式的 Word 表单」这件事从一次性脚本变成可复用能力。

**支持 .docx（读写）与 .doc（只读，Word 97-2003 二进制）。** 名字里的 OOXML 指它工作的
格式层（Office Open XML / ECMA-376），Excel 与 PowerPoint 尚未实现——不要把它当成覆盖
三种格式的套件。`.doc` 直读走内置的 CFB 容器与 Word 二进制解析（`cfb.py`、`doc_read.py`），
不调用外部转换器：二进制格式的样式/修订/图片边界在输出 `warnings` 里逐项声明，块模型与
`.docx` 同为 `docx-model-v1`。

## 为什么有它

同一套动作在两个项目里各写了一遍：读 DOCX 结构（标题层级、表格行、图片）、
按 JSON 填充模板、生成后校验。每次都靠现场脚本，既无法复用也无法回归。
这个组件把共因抽出来，做成与具体业务无关的一层。

## 技术路线

Python 3 + lxml。不是偏好问题，是被 Word 逼出来的：

- 标准库 `xml.etree` 序列化时会丢掉约 34 个命名空间前缀，`mc:Ignorable="w14 w15 …"`
  引用的前缀一旦消失，Word 直接判为「文件已经损坏」。lxml 完整保留全部前缀。
- 反过来，`etree.cleanup_namespaces` 同样不能用：它会删掉只被 `mc:Ignorable` 按名字
  引用的前缀。本组件改用显式 `nsmap` 建节点，避免出现 `ns0` 前缀。

读写路径都依赖 lxml。读取理论上可用标准库解析，但那样会有两套解析代码，
两条路径的字段口径迟早分叉——这里宁可统一依赖，也不接受两份实现。

## 用法

```bash
python -m office_ooxml_kit extract --path 甲.docx --output 甲.json     # 块模型（docx-model-v1）
python -m office_ooxml_kit extract --path 乙.doc  --output 乙.json     # 同上，二进制直读
python -m office_ooxml_kit scan    --path 甲.docx                      # 只列需人工核对的内容
python -m office_ooxml_kit inspect --path 甲.docx --output 甲-包清单.json
python -m office_ooxml_kit fill    --path 模板.docx --output 结果.docx --data fill.json
python -m office_ooxml_kit verify  --path 结果.docx --baseline 甲-包清单.json
```

Node 项目通过子进程调用，和它们现在调 `docx.ps1` 的方式一致。

## 契约

块模型字段名沿用旧 PowerShell 实现（原 `form_generate/tools/docx.ps1`，已归档到
`tools/reference/docx.ps1`）既有的 `docx-model-v1`，既有消费方零改动即可切换。
详见 [docx-model-v1](docs/docx-model-v1.md)。

## 文档

- [共因分析](docs/common-cause.md)——两个项目为什么各写了一遍，以及各漏掉了什么
- [架构](docs/architecture.md)——为什么是 Python+lxml、分层、三件必须共享的事
- [契约 docx-model-v1](docs/docx-model-v1.md)——块与图片的字段语义
- [迁移指南](docs/migration.md)——两种接入方式、降级路径、检查清单

## 工具

```bash
# 把源码复制进下游项目（发布用），并写 VENDOR.json 指纹；--check 只校验
python tools/vendor.py <目标项目>/tools/office_ooxml_kit
python tools/vendor.py <目标项目>/tools/office_ooxml_kit --check

# 语料级比对：与 PowerShell 参考实现逐块逐字段对照
PYTHONPATH=. python tools/compare_against_powershell.py 甲.docx 乙.docx

# 差异归因：证明每处差异来自公式文字或编号，而不是「大概没问题」
PYTHONPATH=. python tools/attribute_differences.py 甲.docx

# 端到端重放：用既有运行目录的 fill.json 重生成，比对块模型与页面
PYTHONPATH=. python tools/replay_run.py runs/<运行目录>
```

`attribute_differences.py` 的判据是可证伪的：它分别计算「只取 `w:t`」与「按文档顺序取
`w:t`+`m:t`」两种文字，若两者恰好等于两侧结果，差异就被证明完全来自公式。

## 测试

```bash
PYTHONPATH=. python -m unittest discover -s tests
```

测试用真实文档，不打桩：断言的对象是「Word 能不能打开」「和参照实现是否逐块一致」
「页数是否相同」，而不是内部函数调用次数。

## 已验证结果

| 项目 | 结果 |
| --- | --- |
| office_ooxml_kit 自测 | 41 项通过 |
| `form_generate` | 48/48 通过；真实交付物重放 44 块 0 差异、18 页、15,407 字 |
| `DO-178C-review-workflow` | 145 项通过（跳过 1 项） |
| 语料比对（71 份文档、35,078 块） | 9 处差异，全部裁定为参照实现的问题：7 处漏公式文字、2 处编号差一位（Word 自身显示 `5.4.1`，与 Python 一致） |
