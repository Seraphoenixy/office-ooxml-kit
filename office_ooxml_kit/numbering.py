"""自动编号重建。

Word 里「序号」常是样式绑定的自动编号：单元格正文为空，显示 1、2、3…。只取 w:t
会得到空串——CR 里就成了 `--`。各项目此前处理不一致（form_generate 完整实现、
DO-178C 只写警告、word-review 依赖 python-docx 不做），这里收敛成一份。

计数规则来自真实文档的踩坑记录：
- 共用同一 abstractNum 的连续单元格继续累加（1 号之后的 26 号）；
- `startOverride` 只在该 numId **首次**出现时重置计数，之后沿用累加——
  SWTP-013 的三行共用 numId=54，Word 显示 1、2、3 而不是 1、1、1。
"""
from __future__ import annotations

from .ns import attr, q

__all__ = ['Numbering', 'format_value']

_ROMAN = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'), (100, 'C'),
          (90, 'XC'), (50, 'XL'), (40, 'XL'), (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
_ROMAN = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'), (100, 'C'),
          (90, 'XC'), (50, 'L'), (40, 'XL'), (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]


def _roman(value: int) -> str:
    if value <= 0:
        return str(value)
    out = []
    for unit, glyph in _ROMAN:
        while value >= unit:
            out.append(glyph)
            value -= unit
    return ''.join(out)


def format_value(value: int, fmt: str) -> str:
    """按 w:numFmt 渲染计数值；语料里出现的是 decimal/upperLetter/lowerLetter/lowerRoman/bullet。"""
    if fmt == 'upperLetter':
        return chr(ord('A') + (value - 1) % 26) if value > 0 else ''
    if fmt == 'lowerLetter':
        return chr(ord('a') + (value - 1) % 26) if value > 0 else ''
    if fmt == 'upperRoman':
        return _roman(value).upper()
    if fmt == 'lowerRoman':
        return _roman(value).lower()
    if fmt == 'decimalZero':
        return f'{value:02d}'
    if fmt in ('bullet', 'none'):
        return ''
    if fmt == 'chineseCounting' or fmt == 'chineseCountingThousand':
        return _chinese(value)
    return str(value)


def _chinese(value: int) -> str:
    digits = '零一二三四五六七八九'
    if value < 10:
        return digits[value]
    if value < 20:
        return '十' + (digits[value % 10] if value % 10 else '')
    return digits[value // 10] + '十' + (digits[value % 10] if value % 10 else '')


class Numbering:
    """styleId → 编号定义 → 显示文本。

    `paragraph_number` 用于正文段落（章节号），`cell_number` 用于表格单元格
    （序号），二者计数空间分离，与 Word 的实际行为一致。
    """

    def __init__(self, styles_xml=None, numbering_xml=None):
        self.styles = {}
        self.num_to_abstract = {}
        self.abstracts = {}
        self.overrides = {}
        self._counters = {}
        self._cell_counters = {}
        self._cell_started = set()
        if styles_xml is not None:
            self._load_styles(styles_xml)
        if numbering_xml is not None:
            self._load_numbering(numbering_xml)

    # ---- 解析定义 ----
    def _load_styles(self, root):
        for style in root.iter(q('w:style')):
            self.styles[attr(style, 'w:styleId')] = style

    def _load_numbering(self, root):
        for abstract in root.iter(q('w:abstractNum')):
            self.abstracts[attr(abstract, 'w:abstractNumId')] = abstract
        for num in root.iter(q('w:num')):
            num_id = attr(num, 'w:numId')
            self.num_to_abstract[num_id] = attr(num.find(q('w:abstractNumId')), 'w:val')
            self.overrides[num_id] = {
                attr(override, 'w:ilvl'): override for override in num.findall(q('w:lvlOverride'))
            }

    def _style_prop(self, style_id, tag):
        """沿 basedOn 向上查样式属性节点（outlineLvl 等）；带环路保护。"""
        seen, node = set(), style_id
        while node and node in self.styles and node not in seen:
            seen.add(node)
            style = self.styles[node]
            props = style.find(q('w:pPr'))
            found = props.find(q(tag)) if props is not None else None
            if found is not None:
                return found
            based = style.find(q('w:basedOn'))
            node = attr(based, 'w:val') if based is not None else ''
        return None

    def outline_level(self, paragraph, style):
        """大纲级别：段落自带 outlineLvl 优先，否则沿样式继承。

        这是标题的唯一可靠判据。按样式名匹配 Heading 会在中文模板上失效
        （样式名是「标题 1」），按是否带自动编号判定又会把编号列表当成标题。
        """
        direct = paragraph.find(f'{q("w:pPr")}/{q("w:outlineLvl")}')
        node = direct if direct is not None else self._style_prop(style, 'w:outlineLvl')
        value = attr(node, 'w:val')
        return value if value.isdigit() and 0 <= int(value) <= 8 else ''

    def _style_num_id(self, style_id):
        """沿 basedOn 向上找样式的 numId；带环路保护。"""
        node = style_id
        for _ in range(40):
            if not node or node not in self.styles:
                return ''
            style = self.styles[node]
            props = style.find(q('w:pPr'))
            num_pr = props.find(q('w:numPr')) if props is not None else None
            if num_pr is not None:
                value = attr(num_pr.find(q('w:numId')), 'w:val')
                if value:
                    return value
            based = style.find(q('w:basedOn'))
            node = attr(based, 'w:val') if based is not None else ''
        return ''

    def _resolve(self, num_id, level):
        """返回 (abstract 节点, lvl 节点, override 节点)。"""
        abstract_id = self.num_to_abstract.get(num_id)
        abstract = self.abstracts.get(abstract_id)
        if abstract is None:
            return None, None, None
        override = self.overrides.get(num_id, {}).get(level)
        node = override.find(q('w:lvl')) if override is not None else None
        if node is None:
            node = abstract.find(f'{q("w:lvl")}[@{q("w:ilvl")}="{level}"]')
        return abstract, node, override

    def _style_value(self, style_id, prop):
        """沿 basedOn 链查 `w:numPr/w:...` 的值。

        numId 与 ilvl 必须各自独立上溯：中文模板里样式 30 写了 numPr 但 numId 为空、
        只在 ilvl 上有值，numId 实际来自上层的样式 21。把整个 numPr 节点一起取下来
        就会拿到空 numId，编号整段丢失。
        """
        node = style_id
        for _ in range(40):
            if not node or node not in self.styles:
                return ''
            style = self.styles[node]
            props = style.find(q('w:pPr'))
            num_pr = props.find(q('w:numPr')) if props is not None else None
            if num_pr is not None:
                value = attr(num_pr.find(q(prop)), 'w:val')
                if value:
                    return value
            based = style.find(q('w:basedOn'))
            node = attr(based, 'w:val') if based is not None else ''
        return ''

    def _num_pr(self, node):
        """段落/单元格内首段的 numId、ilvl、样式。

        numId 与 ilvl 都可能在样式上（中文模板的章节编号几乎都是样式绑定），
        段落自带的 pPr/w:numPr 优先；丢掉样式上的 ilvl 会把二级标题算成一级，
        编号会变成 1、2、3 而不是 1、1.1、1.2。
        """
        paragraph = node if node.tag == q('w:p') else node.find(q('w:p'))
        if paragraph is None:
            return '', '', ''
        props = paragraph.find(q('w:pPr'))
        direct = props.find(q('w:numPr')) if props is not None else None
        style = attr(props.find(q('w:pStyle')), 'w:val') if props is not None else ''

        num_id = attr(direct.find(q('w:numId')), 'w:val') if direct is not None else ''
        level = attr(direct.find(q('w:ilvl')), 'w:val') if direct is not None else ''
        if not num_id:
            num_id = self._style_value(style, 'w:numId')
        if not level:
            level = self._style_value(style, 'w:ilvl')
        return num_id, level or '0', style

    def paragraph_number(self, paragraph):
        """章节的显示编号；没有自动编号时返回空串。"""
        num_id, level, _ = self._num_pr(paragraph)
        if not num_id or num_id not in self.num_to_abstract:
            return ''
        abstract, node, override = self._resolve(num_id, level)
        if node is None:
            return ''
        key = f'{num_id}/{level}'
        start_override = override.find(q('w:startOverride')) if override is not None else None
        if start_override is not None and key not in self._counters:
            value = int(attr(start_override, 'w:val') or '1')
        elif self._counters.get(key):
            value = self._counters[key] + 1
        else:
            initial = attr(node.find(q('w:start')), 'w:val')
            value = int(initial) if initial.lstrip('-').isdigit() else 1
        self._counters[key] = value
        # 上级编号前进时，其下各级编号必须重新开始。
        for other in list(self._counters):
            if other.startswith(f'{num_id}/') and other.split('/')[1].isdigit() \
                    and int(other.split('/')[1]) > int(level):
                self._counters.pop(other)
        return self._label(node, self.num_to_abstract[num_id], num_id, level, value, self._counters)

    def cell_number(self, cell):
        """表格单元格的序号；仅在单元格没有可见文字时使用。"""
        num_id, level, _ = self._num_pr(cell)
        if not num_id or num_id not in self.num_to_abstract:
            return ''
        abstract, node, override = self._resolve(num_id, level)
        if node is None:
            return ''
        abstract_id = self.num_to_abstract[num_id]
        start_override = override.find(q('w:startOverride')) if override is not None else None
        if start_override is not None and num_id not in self._cell_started:
            # 只在本次文档内该 numId 首次出现时重置，之后沿用累加。
            value = int(attr(start_override, 'w:val') or '1')
            self._cell_started.add(num_id)
        elif self._cell_counters.get(abstract_id):
            value = self._cell_counters[abstract_id] + 1
        else:
            initial = attr(node.find(q('w:start')), 'w:val')
            value = int(initial) if initial.lstrip('-').isdigit() else 1
        self._cell_counters[abstract_id] = value
        return self._label(node, abstract_id, num_id, level, value, self._cell_counters)

    def _label(self, node, abstract_id, num_id, level, value, counters):
        label = attr(node.find(q('w:lvlText')), 'w:val')
        if not label:
            return format_value(value, attr(node.find(q('w:numFmt')), 'w:val'))
        abstract = self.abstracts[abstract_id]
        for index in range(0, int(level) + 1):
            reference = abstract.find(f'{q("w:lvl")}[@{q("w:ilvl")}="{index}"]')
            current = counters.get(f'{num_id}/{index}')
            if current is None:
                current = counters.get(abstract_id, 1)
            fmt = attr(reference.find(q('w:numFmt')), 'w:val') if reference is not None else 'decimal'
            label = label.replace(f'%{index + 1}', format_value(int(current), fmt))
        return label
