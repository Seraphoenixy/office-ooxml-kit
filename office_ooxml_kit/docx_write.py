"""CR 模板填充：把差异左右两侧的原文与图片写进更改申请单。

这一层只做「按已知位置改写 XML」，不做任何判断——填哪个格、填什么内容由调用方决定。
三条硬约束来自既有工作流，改错任一都会让交付物失去意义：
- 拒绝覆盖已存在的输出文件（每次生成落在独立目录，旧结果保留）；
- 原文表格必须重建为真正的嵌套表格，不能把管道符文本塞进一个单元格；
- 图片按内容哈希去重后重新挂关系，漏了关系项或内容类型 Word 会报文件损坏。
"""
from __future__ import annotations

import copy
import hashlib
import json
import re

from lxml import etree

from . import images as images_mod
from .ns import A, PIC, R, W, WP, attr, build, el, local, q, tostring
from .package import Package, PackageError

__all__ = ['fill', 'FillError', 'MAX_IMAGE_HEIGHT_EMU', 'IMAGE_WIDTH_MARGIN']

# 图片等比缩放后再嵌入：宽度不超过单元格可用宽度，高度上限 9cm。只缩放，不裁切。
MAX_IMAGE_HEIGHT_EMU = 3240000
IMAGE_WIDTH_MARGIN = 180
EMU_PER_TWIP = 635
TABLE_WIDTH_MARGIN = 180
IMAGE_MARGIN_TWIPS = 240

TOKEN_RE = re.compile('⟦图:[0-9A-Fa-f]{8}⟧')
TOKEN_SPLIT = re.compile('(⟦图:[0-9A-Fa-f]{8}⟧)')
TOKEN_EXACT = re.compile('^⟦图:[0-9A-Fa-f]{8}⟧$')
UNCERTAIN = re.compile('待确认|待审核|待核实|待补充|需.*确认')

# 模板外层表结构：21 行且首格是单据标题。对不上就拒绝填充，避免按猜到的位置写错格。
EXPECTED_ROWS = 21
EXPECTED_TITLE = '更改申请单'

# 这些表头首列出现时，第一行是表头而非数据行。
HEADER_MARKERS = ('序号', '需求标识', '测试用例标识', '工作岗位', '数据项标识', '设备名称', '软件名称')


class FillError(RuntimeError):
    """填充失败：模板结构与预期不符、缺少图片字节来源、输出已存在。"""


def fill(template, output, data, *, overwrite=False):
    """按 `data` 填充 `template`，写出到 `output`。

    `data` 是 fill.json 的解析结果：
    {documentId, title, crId, sources: {old, new}, slots: [...], changes: [...]}
    """
    if isinstance(data, str):
        with open(data, encoding='utf-8') as handle:
            data = json.load(handle)
    _guard_output(output, overwrite)

    package = Package(template)
    document = etree.fromstring(package.read('word/document.xml'))
    table = document.find(f'{q("w:body")}/{q("w:tbl")}')
    if table is None:
        raise FillError(f'模板缺少外层表格：{template}')
    rows = table.findall(q('w:tr'))
    if len(rows) != EXPECTED_ROWS:
        raise FillError(f'模板结构不匹配（外层表 {len(rows)} 行，预期 {EXPECTED_ROWS} 行），拒绝按未知位置填充。')
    from .docx_read import text_of
    if text_of(rows[0]).strip() != EXPECTED_TITLE:
        raise FillError(f'模板首行不是「{EXPECTED_TITLE}」，拒绝填充。')

    writer = _Writer(package, data)
    for slot in data.get('slots') or []:
        index = int(slot['row'])
        if not 0 <= index < len(rows):
            raise FillError(f'槽位越界：row={slot["row"]}，外层表只有 {len(rows)} 行。')
        cells = rows[index].findall(q('w:tc'))
        if int(slot['cell']) >= len(cells):
            raise FillError(f'槽位越界：row={slot["row"]} cell={slot["cell"]}，该行只有 {len(cells)} 格。')
        writer.set_cell(cells[int(slot['cell'])], str(slot.get('text') or ''))

    # 第 21 行是变更行原型：克隆它来生成每一条差异行，再把原型本身移出表格。
    prototype = copy.deepcopy(rows[20])
    table.remove(rows[20])
    for change in data.get('changes') or []:
        row = copy.deepcopy(prototype)
        for tag in ('w:trHeight', 'w:cantSplit'):
            for node in row.findall(f'{q("w:trPr")}/{q(tag)}'):
                node.getparent().remove(node)
        cells = row.findall(q('w:tc'))
        for index, key in ((0, 'old'), (1, 'new'), (2, 'remark')):
            parts = change.get(f'{key}Parts')
            if parts:
                writer.set_rich_cell(cells[index], parts)
            else:
                writer.set_cell(cells[index], str(change.get(key) or ''))
        table.append(row)

    # 可见空格补「--」：空白状态列会让人以为漏填。
    # 纵向合并的延续格不是独立可见单元格；含图片的格文字为空是正常的，补占位符会把图片顶掉。
    for cell in table.iter(q('w:tc')):
        merge = cell.find(f'{q("w:tcPr")}/{q("w:vMerge")}')
        if merge is not None and attr(merge, 'w:val') != 'restart':
            continue
        if cell.find('.//' + q('w:drawing')) is not None:
            continue
        if cell.find(q('w:tbl')) is None and not text_of(cell).strip():
            writer.set_cell(cell, '--')

    _mark_uncertain(table, rows_before=EXPECTED_ROWS)

    package.set('word/document.xml', tostring(document, xml_declaration=True))
    _clear_metadata(package, data)
    _set_update_fields_false(package)
    package.save(output, overwrite=overwrite)
    return {
        'output': str(output),
        'rows': len(data.get('changes') or []),
        'slots': len(data.get('slots') or []),
        'images': writer.images_written,
    }


def _guard_output(output, overwrite):
    from pathlib import Path
    target = Path(output)
    if target.exists() and not overwrite:
        raise FillError(f'输出已存在，禁止覆盖：{output}')
    if target.parent and not target.parent.exists():
        raise FillError(f'输出目录不存在：{target.parent}')


class _Writer:
    """把文字、富单元格（文字 + 图片 + 嵌套表）写进指定单元格。"""

    def __init__(self, package, data):
        self.package = package
        self.data = data
        self.sources = {}
        self.images_written = 0
        self._by_sha = {}
        self._seq = 0
        for side in ('old', 'new'):
            path = (data.get('sources') or {}).get(side)
            if path:
                self.sources[side] = Package(path)
        self.media = images_mod.Media(package)

    # -- 单元格 ------------------------------------------------------
    def set_cell(self, cell, value):
        """写入纯文本：保留单元格段落格式与首个 run 的字体，按换行分段。"""
        paragraph = cell.find(q('w:p'))
        if paragraph is None:
            paragraph = el('w:p')
        template = copy.deepcopy(paragraph)
        for child in list(template):
            if local(child) != 'pPr':
                template.remove(child)
        run_props = None
        first_run = cell.find('.//' + q('w:r') + '/' + q('w:t'))
        for run in cell.iter(q('w:r')):
            if run.find(q('w:t')) is not None:
                run_props = run.find(q('w:rPr'))
                break
        for child in list(cell):
            if local(child) != 'tcPr':
                cell.remove(child)
        for line in str(value).split('\n'):
            dest = copy.deepcopy(template)
            run = el('w:r')
            if run_props is not None:
                run.append(copy.deepcopy(run_props))
            text = el('w:t')
            text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            text.text = line
            run.append(text)
            dest.append(run)
            for keep in ('w:keepNext', 'w:pageBreakBefore'):
                for node in dest.findall(f'{q("w:pPr")}/{q(keep)}'):
                    node.getparent().remove(node)
            cell.append(dest)

    def set_rich_cell(self, cell, parts, base=None):
        """富单元格：文字段、图片、嵌套表格按原顺序写入。

        `base` 是供段落与字体继承的样板单元格。原文表格里再嵌原文表格时，内层单元格
        是新建的、自身没有段落可继承，必须把外层样板一路传下来；否则内层段落丢掉
        pPr，行距与对齐和原文对不上，页数也会差一页。
        """
        if base is None:
            base = copy.deepcopy(cell)
        if base.find(q('w:p')) is None:
            base.append(el('w:p'))
        width_node = cell.find(f'{q("w:tcPr")}/{q("w:tcW")}')
        outer_width = int(attr(width_node, 'w:w') or 6000)
        cell_max_emu = (outer_width - IMAGE_MARGIN_TWIPS) * EMU_PER_TWIP
        for child in list(cell):
            if local(child) != 'tcPr':
                cell.remove(child)
        for part in parts:
            kind = part.get('type')
            if kind == 'paragraph':
                scratch = copy.deepcopy(base)
                self.set_cell(scratch, str(part.get('text') or ''))
                for paragraph in scratch.findall(q('w:p')):
                    cell.append(copy.deepcopy(paragraph))
            elif kind == 'image':
                cell.append(self._image_paragraph(part, cell_max_emu))
            elif kind == 'table':
                cell.append(self._nested_table(part, base, cell_max_emu))
        # Word 硬性要求：单元格最后一个子元素必须是段落。嵌套表后面缺段落，
        # 文档会直接报「文件已经损坏」——判断的是「末尾是不是段落」，不是「里面有段落」。
        children = [child for child in cell if local(child) != 'tcPr']
        if not children or local(children[-1]) != 'p':
            cell.append(el('w:p'))

    # -- 图片 --------------------------------------------------------
    def _image_paragraph(self, part, max_width_emu):
        """嵌入一张图片：按原尺寸等比缩放，宽度受格宽限制、高度不超过 9cm。"""
        data = self._image_bytes(part)
        digest = hashlib.sha256(data).hexdigest().upper()
        ext = (part.get('ext') or 'png').lower().lstrip('.')
        if digest in self._by_sha:
            rel_id, _name = self._by_sha[digest]
        else:
            self._seq += 1
            rel_id, name, _ = self.media.embed(self.package, data, ext, seq=self._seq)
            self._by_sha[digest] = (rel_id, name)
            self.images_written += 1
        cx = int(part.get('cx') or 0)
        cy = int(part.get('cy') or 0)
        if cx <= 0 or cy <= 0:
            cx, cy = 3600000, 2700000
        scale = min(1.0, max_width_emu / cx, MAX_IMAGE_HEIGHT_EMU / cy)
        if scale <= 0:
            scale = 1.0
        # PowerShell 的 [long] 转换是「四舍六入五成双」，等价于 Python 的 round()；
        # 用 int() 截断会每张图少 1 EMU，几十张图之间的行高差异会累积成版面偏移。
        draw_cx, draw_cy = round(cx * scale), round(cy * scale)

        # 新建的 drawing 元素所属命名空间在模板根上并未声明；在这里显式给出前缀，
        # lxml 才会沿用 a/pic/wp 而不是自造 ns0。
        paragraph = build('w:p')
        props = el('w:pPr')
        align = el('w:jc')
        align.set(q('w:val'), 'center')
        props.append(align)
        paragraph.append(props)
        run = el('w:r')
        drawing = el('w:drawing')
        inline = el('wp:inline')
        for key in ('distT', 'distB', 'distL', 'distR'):
            inline.set(key, '0')
        extent = el('wp:extent')
        extent.set('cx', str(draw_cx))
        extent.set('cy', str(draw_cy))
        inline.append(extent)
        effects = el('wp:effectExtent')
        for key in ('l', 't', 'r', 'b'):
            effects.set(key, '0')
        inline.append(effects)
        doc_pr = el('wp:docPr')
        doc_pr.set('id', str(self._seq))
        doc_pr.set('name', f'图片 {self._seq}')
        inline.append(doc_pr)
        frame = el('wp:cNvGraphicFramePr')
        locks = el('a:graphicFrameLocks')
        locks.set('noChangeAspect', '1')
        frame.append(locks)
        inline.append(frame)
        graphic = el('a:graphic')
        graphic_data = el('a:graphicData')
        graphic_data.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
        picture = el('pic:pic')
        nv = el('pic:nvPicPr')
        cnv = el('pic:cNvPr')
        cnv.set('id', str(self._seq))
        cnv.set('name', part.get('name') or f'image{self._seq}.{ext}')
        nv.append(cnv)
        nv.append(el('pic:cNvPicPr'))
        picture.append(nv)
        blip_fill = el('pic:blipFill')
        blip = el('a:blip')
        blip.set(q('r:embed'), rel_id)
        blip_fill.append(blip)
        stretch = el('a:stretch')
        stretch.append(el('a:fillRect'))
        blip_fill.append(stretch)
        picture.append(blip_fill)
        sp_pr = el('pic:spPr')
        xfrm = el('a:xfrm')
        off = el('a:off')
        off.set('x', '0')
        off.set('y', '0')
        xfrm.append(off)
        ext_node = el('a:ext')
        ext_node.set('cx', str(draw_cx))
        ext_node.set('cy', str(draw_cy))
        xfrm.append(ext_node)
        sp_pr.append(xfrm)
        geom = el('a:prstGeom')
        geom.set('prst', 'rect')
        geom.append(el('a:avLst'))
        sp_pr.append(geom)
        picture.append(sp_pr)
        graphic_data.append(picture)
        graphic.append(graphic_data)
        inline.append(graphic)
        drawing.append(inline)
        run.append(drawing)
        paragraph.append(run)
        return paragraph

    def _image_bytes(self, part):
        """按来源版本、部分名取媒体字节；不做转码，显示效果与内容哈希都与原文一致。"""
        side = part.get('source')
        source = self.sources.get(side)
        if source is None:
            raise FillError(f'缺少图片来源：{side}')
        name = part.get('part')
        if not source.has(name):
            raise FillError(f'来源 {side} 中找不到媒体部分：{name}')
        return source.read(name)

    # -- 嵌套表 ------------------------------------------------------
    def _nested_table(self, part, base_cell, cell_max_emu):
        """把原文表格重建成真正的嵌套表：保留列宽比例、表头、跨列与合并。"""
        from .docx_read import text_of as cell_text
        rows = []
        if part.get('header'):
            rows.append(part['header'])
        rows.extend(part.get('rows') or [])
        columns = int(part.get('cols') or 0)
        if columns <= 0:
            columns = max((len(row) for row in rows), default=1)
        grid = part.get('grid') or []
        available = max(1, (int(attr(base_cell.find(f'{q("w:tcPr")}/{q("w:tcW")}'), 'w:w') or 6000))
                        - TABLE_WIDTH_MARGIN)

        table = el('w:tbl')
        props = el('w:tblPr')
        width = el('w:tblW')
        width.set(q('w:w'), '5000')
        width.set(q('w:type'), 'pct')
        props.append(width)
        layout = el('w:tblLayout')
        layout.set(q('w:type'), 'fixed' if part.get('layout') == 'fixed' else 'autofit')
        props.append(layout)
        borders = el('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            border = el('w:' + edge)
            border.set(q('w:val'), 'single')
            border.set(q('w:sz'), '4')
            border.set(q('w:color'), '000000')
            borders.append(border)
        props.append(borders)
        table.append(props)

        if grid:
            total = sum(int(value or 0) for value in grid) or 1
            # 取整方式必须与参照实现一致：PowerShell 的 [int] 转换是四舍五入，
            # Python 的 int() 是截断，照抄会每列少 1 twip，累积后整行宽度和原表不符。
            positions = [round(available * int(value or 0) / total) for value in grid]
        else:
            weights = {2: (0.40, 0.60), 3: (0.12, 0.25, 0.63),
                       5: (0.08, 0.24, 0.15, 0.12, 0.41)}.get(columns)
            if weights is None:
                weights = tuple(1.0 / columns for _ in range(columns))
            positions = [round(available * weight) for weight in weights]
        grid_node = el('w:tblGrid')
        for position in positions:
            column = el('w:gridCol')
            column.set(q('w:w'), str(position))
            grid_node.append(column)
        table.append(grid_node)

        for index, values in enumerate(rows):
            row = el('w:tr')
            if part.get('header') and index == 0:
                row_props = el('w:trPr')
                row_props.append(el('w:tblHeader'))
                row.append(row_props)
            spans = self._row_spans(part, index)
            column_index = 0
            for cell_index, value in enumerate(values):
                span = spans[cell_index] if cell_index < len(spans) else 1
                span = max(1, int(span or 1))
                if column_index + span > columns:
                    span = max(1, columns - column_index)
                cell_width = sum(positions[column_index:column_index + span])
                cell = el('w:tc')
                cell_props = el('w:tcPr')
                cell_width_node = el('w:tcW')
                cell_width_node.set(q('w:type'), 'dxa')
                cell_width_node.set(q('w:w'), str(cell_width))
                cell_props.append(cell_width_node)
                if span > 1:
                    span_node = el('w:gridSpan')
                    span_node.set(q('w:val'), str(span))
                    cell_props.append(span_node)
                cell.append(cell_props)
                if isinstance(value, dict) and value.get('parts'):
                    self.set_rich_cell(cell, value['parts'], base_cell)
                else:
                    scratch = copy.deepcopy(base_cell)
                    self.set_cell(scratch, str(value))
                    for paragraph in scratch.findall(q('w:p')):
                        copied = copy.deepcopy(paragraph)
                        if part.get('header') and index == 0:
                            for run in copied.findall(q('w:r')):
                                run_props = run.find(q('w:rPr'))
                                if run_props is None:
                                    run_props = el('w:rPr')
                                    run.insert(0, run_props)
                                run_props.append(el('w:b'))
                        cell.append(copied)
                row.append(cell)
                column_index += span
            # 行短于网格时补一个跨列空格，否则 Word 会把缺列的行拉宽。
            if column_index < columns:
                filler = el('w:tc')
                filler_props = el('w:tcPr')
                filler_width = el('w:tcW')
                filler_width.set(q('w:type'), 'dxa')
                filler_width.set(q('w:w'), str(sum(positions[column_index:])))
                filler_props.append(filler_width)
                span_node = el('w:gridSpan')
                span_node.set(q('w:val'), str(columns - column_index))
                filler_props.append(span_node)
                filler.append(filler_props)
                filler.append(el('w:p'))
                row.append(filler)
            table.append(row)
            _ = cell_text
        return table

    def _row_spans(self, part, index):
        """该行的跨列信息：表头行与数据行分开存放，按 allRows 的序号取。"""
        header = bool(part.get('header'))
        if header and index == 0:
            return list(part.get('headerSpans') or [])
        data_index = index - 1 if header else index
        row_spans = part.get('rowSpans') or []
        if 0 <= data_index < len(row_spans):
            return list(row_spans[data_index] or [])
        return []


def _mark_uncertain(table, *, rows_before):
    """待确认类文字标红：表单值与编辑备注用红字，原文引用保留原格式。"""
    for index, row in enumerate(table.findall(q('w:tr'))):
        if index < rows_before:
            targets = [row]
        else:
            cells = row.findall(q('w:tc'))
            targets = cells[2:3]
        for target in targets:
            for paragraph in target.iter(q('w:p')):
                from .docx_read import text_of
                if not UNCERTAIN.search(text_of(paragraph)):
                    continue
                for run in paragraph.findall(q('w:r')):
                    if run.find(q('w:t')) is None:
                        continue
                    props = run.find(q('w:rPr'))
                    if props is None:
                        props = el('w:rPr')
                        run.insert(0, props)
                    for color in props.findall(q('w:color')):
                        props.remove(color)
                    color = el('w:color')
                    color.set(q('w:val'), 'FF0000')
                    props.append(color)


def _clear_metadata(package, data):
    """清空模板遗留的作者与标题信息，并按 `data` 里的 `headerReplacements` 替换页眉页脚旧单事实。

    具体替换什么由调用方提供，组件不内置任何业务模板事实：把旧单号、旧标题、旧申请人
    写死在通用组件里，等于把某一个业务模板的历史信息变成所有使用者的默认行为，
    而且这些值会随组件一起发布出去。
    """
    replacements = data.get('headerReplacements') or {}
    raw = package.optional('docProps/core.xml')
    if raw is not None:
        root = etree.fromstring(raw)
        for node in root:
            if local(node) in ('creator', 'lastModifiedBy', 'title', 'subject',
                               'description', 'keywords'):
                node.text = None
        package.set('docProps/core.xml', tostring(root, xml_declaration=True))
    for name in package.part_names():
        if not re.match(r'^word/(header|footer).*\.xml$', name):
            continue
        root = etree.fromstring(package.read(name))
        for text in root.iter(q('w:t')):
            value = text.text or ''
            for key, replacement in replacements.items():
                value = value.replace(key, replacement)
            text.text = value
        package.set(name, tostring(root, xml_declaration=True))


def _set_update_fields_false(package):
    """静态 CR 不应弹出域更新提示。"""
    raw = package.optional('word/settings.xml')
    if raw is None:
        return
    root = etree.fromstring(raw)
    node = root.find('.//' + q('w:updateFields'))
    if node is None:
        node = el('w:updateFields')
        root.append(node)
    node.set(q('w:val'), 'false')
    package.set('word/settings.xml', tostring(root, xml_declaration=True))
