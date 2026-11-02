# 迁移指南

面向两个已经接好的项目，以及以后要接进来的项目。

## 安装

```bash
cd D:/project/office-ooxml-kit
python -m pip install -e .
```

装好后 `python -m office_ooxml_kit` 在任何目录都能用。要求 `lxml>=6`——读写都要。

发布给同事时不要让他装这个包，改用 `tools/vendor.py` 把源码复制进项目，见"发布形态"一节。

## 接入方式一：Node 调 CLI（`form_generate` 用的就是这个）

```js
import { execFileSync } from 'node:child_process';

const PYTHON = process.env.OFFICEKIT_PYTHON || 'python';

function office_ooxml_kit(action, { path: source, output, data } = {}) {
  const args = ['-m', 'office_ooxml_kit', action, '--path', path.resolve(source), '--output', path.resolve(output)];
  if (data) args.push('--data', path.resolve(data));
  // 控制台默认 GBK，不强制 UTF-8 时 Python 的报错会变成乱码。
  const env = { ...process.env, PYTHONIOENCODING: 'utf-8' };
  try {
    execFileSync(PYTHON, args, { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024, env });
  } catch (e) {
    const detail = String(e.stderr ?? e.message).trim().split('\n').filter(Boolean).at(-1) ?? '';
    throw new Error(`office_ooxml_kit ${action} 失败：${detail}`);
  }
}
```

两个容易忘的点：

1. **`PYTHONIOENCODING=utf-8` 必须设。** 不设的时候中文报错是乱码，
   面向用户的报错信息等于不可读。
2. **结果走文件，不走 stdout。** stdout 只放进度文字；结构化结果写到 `--output`。
   走管道会引入编码与缓冲问题，文件名里的中文最容易出事。

## 接入方式二：Python 直接 import（`DO-178C-review-workflow` 用的就是这个）

内嵌副本放在 `scripts/office_ooxml_kit/`，而 `review.py` 本来就把 `scripts/` 加进了 `sys.path`，
所以 `import office_ooxml_kit` 直接解析到内嵌副本，不需要任何额外配置。

```python
def _office_ooxml_kit():
    """延迟导入：只在解析 Office 文档时才需要，缺依赖不该让别的功能也无法运行。"""
    try:
        import office_ooxml_kit.docx_read as module
    except ImportError as exc:
        raise ImportError(f'缺少 office_ooxml_kit 内嵌副本：{exc}（应为 scripts/office_ooxml_kit/）') from exc
    return module

report = _office_ooxml_kit().scan(path)     # 只要风险清单
model = _office_ooxml_kit().extract(path)   # 要完整块模型
```

**`scan()` 与 `extract()` 的区别**是本次迁移的关键：

- `extract()` 返回 `docx-model-v1` 块模型（`kind/text/section/...`）；
- `scan()` 只返回 `{warnings, ancillary, parts}`，用来回答"哪些内容文本比较验证不了"。

`DO-178C` 的块形状是 `{id, text, location, style}`，配合 `body/tbl1/tr1/tc1/p1`
路径定位，被评审记录和 `sweep`/`chain` 消费，**不需要**改成 `docx-model-v1`。
它需要的只是那套判据——所以接 `scan()`，不接 `extract()`。

这是本次改造最重要的一条经验：**共用一个组件的正确粒度不一定是"共用数据结构"，
经常是"共用判定标准"。** 硬把两个块形状统一，会破坏两个项目的消费方，收益为零。

## 接入形态一：git 子模块（推荐）

下游项目把本仓库作为子模块引入，代码跟着项目一起走，但不产生第二份历史。
子模块指向的具体 commit 被主仓库记录，所以「用的是哪个版本」是可查证的，不是靠人记。

```powershell
# 在下游项目根目录
git -c protocol.file.allow=always submodule add https://github.com/<you>/office-ooxml-kit.git tools/office_ooxml_kit

# 别人拿到这个项目之后
git clone --recurse-submodules <项目地址>
# 已经 clone 过的补一步：git submodule update --init
```

**布局上有一层容易错的地方**：子模块根目录与包同名，真正的包在它内部。

```
tools/office_ooxml_kit/          ← 子模块根（仓库根）
├── office_ooxml_kit/            ← 真正的 Python 包
├── README.md
└── pyproject.toml
```

所以 `PYTHONPATH` 必须指向**子模块根**，指向 `tools/` 会让 Python 把仓库根当成命名空间包
（`office_ooxml_kit.__version__` 就不存在了）：

```js
const VENDOR = path.join(path.dirname(fileURLToPath(import.meta.url)), 'office_ooxml_kit');
const SEARCH_PATH = process.env.PYTHONPATH
  ? `${VENDOR}${path.delimiter}${process.env.PYTHONPATH}` : VENDOR;
// 子进程 env 里带 PYTHONPATH: SEARCH_PATH 与 PYTHONIOENCODING: 'utf-8'
```

**子模块的硬代价**：普通 `git clone`（不加 `--recurse-submodules`）之后，`tools/office_ooxml_kit`
是**空目录**，`import` 直接失败。这不是配置问题，是子模块的工作方式。两个应对：

- 下发给同事的形态用 `tools/export.py` 导出的普通目录，见下节；
- 代码里把缺失说清楚（`review.py` 的做法）：

```python
try:
    import office_ooxml_kit.docx_read as module
except ImportError as exc:
    raise ImportError('缺少 office_ooxml_kit 内嵌子模块（应为 scripts/office_ooxml_kit/，'
                      '若缺失请执行 git submodule update --init）') from exc
```

**验证跑的是子模块而不是全局安装**：卸载全局包后跑测试。

```powershell
python -m pip uninstall -y office_ooxml_kit
node --test tests/tools.test.mjs tests/workflow.test.mjs
```

这一步是必要的。只要全局装过同名包，`import` 就可能解析到那个包，测试照样全绿，
而子模块其实从未被执行——发布出去才发现缺文件。

## 发布给同事：导出成普通目录

`git archive` **不支持** `--recurse-submodules`（Git 2.56 仍未支持），所以导出分两步：
主仓库归档，再把子模块自己的归档填进对应路径。

```powershell
python tools/export.py <项目路径> <输出目录>
```

导出结果没有 `.git`，子模块内容已经变成普通文件，对方解压即可用。
安全性上有一条硬规则：**子模块当前 checkout 必须等于主仓库锁定的 commit**，不一致就拒绝导出，
不产生任何半成品目录。否则导出的会是一个仓库从未声明过的版本。

## 备选形态：复制源码副本（`vendor.py`）

不用 git 的场合（对方只要一个目录、或者项目本身不在版本控制里），可以把源码复制进去：

```powershell
python tools/vendor.py "D:/project/form_generate/tools/office_ooxml_kit"
python tools/vendor.py "D:/project/form_generate/tools/office_ooxml_kit" --check
```

它只复制运行所需文件（`*.py`、`schema/*.json`），并写出 `VENDOR.json` 记录版本与每个文件的
sha256。`--check` 能报出手改与「源已新增而副本未同步」。

这条路与子模块二选一，不要同时用：两种形态并存时，`import` 解析到哪一份会变得难以判断。

## 降级路径

判据不可用时必须显式说明，不能悄悄退回"看起来没问题"：

```python
try:
    warnings.extend(_office_ooxml_kit().scan(path)['warnings'])
except Exception as exc:
    warnings.append(f'版面内容核对清单未能生成（{exc}），图片、修订、内容控件等需人工确认。')
```

## 迁移检查清单

- [ ] `python -m office_ooxml_kit extract --path <真实文档> --output /tmp/x.json` 能跑通
- [ ] 输出块的 `section` 用 `' / '` 连接，`category` 保留前置内容的不对称
- [ ] 项目自己的测试全绿
- [ ] 用一份**真实交付物**回归：页数、字数、表格数、图片数
- [ ] 内嵌副本经 `vendor.py --check` 校验与源一致（用子模块时改为 `git submodule status`）
- [ ] **卸载全局 `office_ooxml_kit` 后测试仍全绿**——证明跑的是内嵌副本，不是全局安装
- [ ] 旧实现已移除（本项目删除了 `tools/docx.ps1`，参考副本归档在 `tools/reference/`）
- [ ] 依赖说明与实际一致（不再声称"只用标准库"，`lxml` 写进 requirements）
- [ ] 子模块布局下 `PYTHONPATH` 指向子模块根，`office_ooxml_kit.__version__` 取得出值
- [ ] 子模块为空时（模拟普通 clone）报错信息能指明 `git submodule update --init`

页数和字数要用 Word 的 `ComputeStatistics` 比，不要只比块模型——
块模型全对但排版差一页，说明取整或样式继承还有问题。

## 已知取舍

| 取舍 | 原因 |
| --- | --- |
| 保留 `tools/docx.ps1` 作为回归基准 | 它是独立实现，能发现 Python 的单侧错误。本次就靠它发现了公式与编号两类问题（虽然这次是它错） |
| Python 多输出 `style`/`vMerge`/`images` 字段 | 消费方用 `?.`/`??` 取值，多字段无害；删掉反而丢信息 |
| 媒体文件默认命名 `word/media/cr<N>.<ext>` | 既有流程的校验按这个名字认图片；自造前缀会让"图片是否真的进包"判断失效 |
| `verify` 只比较两侧都存在的部件 | 新增图片是"增加"而非"改写"，否则自检必然误报 |
