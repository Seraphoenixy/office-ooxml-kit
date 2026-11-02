"""DOCX 结构提取：把一份文档变成可定位、可比较、可写回的块列表。

字段名沿用旧 PowerShell 实现的既有输出（kind/text/number/style/section/sectionTitle/
category/locator/bodyIndex/tableIndex/rowIndex/cells/cellImages/spans/grid/layout/images），
因此既有 CR 流程可以零改动接入。这份形状就是 docx-model-v1 契约本体。

判定口径与 PowerShell 版逐条对齐，实测输出可逐块比对：
- 标题由 outlineLvl（段落或样式继承）决定，目录样式（toc/目录）不算标题；
- section 是标题路径（用「 / 」连接），sectionTitle 是最内层标题；
- 单元格先取可见文字，无文字且无图片时才回落到自动编号；
- 图片在文字里以内容哈希 token 表示，保证换图一定被比较出来。
"""
from __future__ import annotations

import re

from lxml import etree

from . import images as images_mod
from .ns import W, attr, local, q
from .numbering import Numbering
from .package import Package, PackageError

__all__ = ['extract', 'scan', 'DocxError', 'normalize', 'text_of', 'rich_text', 'category_of']

FRONT = '前置内容'
REVISION = '修订历史'
HEADING_DEPTH = 9
TOC_STYLE = re.compile('toc|目录', re.IGNORECASE)


class DocxError(RuntimeError):
    """文档结构层面的错误：不是 Word 文档、缺 document.xml。"""


def normalize(text) -> str:
    """空白折叠：所有比较前先过这一道，避免空格差异被当成实质差异。"""
    return re.sub(r'\s+', ' ', str(text or '')).strip()


def text_of(node) -> str:
    """取节点的可见文字；非段落节点按子段落逐行拼接。

    w:t 取字、w:tab 补制表符、w:br/w:cr 补换行——与 PowerShell 版一致，
    两边的提取结果才能逐字比对。
    """
    if local(node) != 'p':
        return '\n'.join(filter(None, (text_of(p) for p in node.iter(q('w:p'))))).strip()
    parts = []
    for child in node.iter():
        name = local(child)
        if name == 't':
            parts.append(child.text or '')
        elif name == 'tab':
            parts.append('\t')
        elif name in ('br', 'cr'):
            parts.append('\n')
    return ''.join(parts).strip()


def rich_text(node, media):
    """段落文字 + 图片 token；图片按文档顺序追加在文字之后。"""
    text = text_of(node)
    pictures = media.paragraph_images(node) if local(node) == 'p' else media.cell_images(node)
    if pictures:
        text += ''.join(picture['token'] for picture in pictures)
    return text, pictures


def _row_cells(row):
    """一行里的可见单元格。

    内容控件（w:sdt）会包住整格——加载规范里的字段就是这样存的。不展开就会漏格，
    列数比 gridCol 少，后续按列对齐就会全错位。
    """
    cells = []
    for child in row:
        name = local(child)
        if name == 'tc':
            cells.append(child)
        elif name == 'sdt':
            content = child.find(q('w:sdtContent'))
            if content is not None:
                cells.extend(content.findall(q('w:tc')))
    return cells


def _cell_rich(cell, media, numbering):
    """单元格文字与图片。

    逐段落取「文字 + 图片 token」后按行拼接；只有既无文字也无图片时才用自动编号，
    否则带图无字的单元格会凭空多出序号。
    """
    lines, pictures = [], []
    for paragraph in cell.iter(q('w:p')):
        text, pics = rich_text(paragraph, media)
        if text:
            lines.append(text)
        pictures.extend(pics)
    rendered = '\n'.join(lines).strip()
    if not rendered and not pictures:
        rendered = numbering.cell_number(cell)
    return rendered, pictures


def category_of(section):
    """表格行的章节归属：修订历史归 revision，前置内容归 front。

    注意与段落的分类规则不同（段落前置区仍算 body）——这是既有实现的口径，
    cr-workflow 的封面临界逻辑依赖它，保持一致才不会改变 CR 差异的归属。
    """
    if section == REVISION:
        return 'revision'
    if section == FRONT:
        return 'front'
    return 'body'


def extract(path, *, with_text=True):
    """提取一份 DOCX 的块模型（docx-model-v1）。

    `with_text=False` 时跳过不透明块，只看结构化内容。
    """
    try:
        package = Package(path)
    except PackageError as error:
        # 调用方只需要认识 DocxError 一个异常类型；包层面的失败在这里归一。
        raise DocxError(str(error)) from error
    if not package.has('word/document.xml'):
        raise DocxError(f'不是 Word 文档（缺 word/document.xml）：{path}')

    document = etree.fromstring(package.read('word/document.xml'))
    styles_xml = package.optional('word/styles.xml')
    numbering_xml = package.optional('word/numbering.xml')
    numbering = Numbering(
        etree.fromstring(styles_xml) if styles_xml else None,
        etree.fromstring(numbering_xml) if numbering_xml else None,
    )
    media = images_mod.Media(package)

    body = document.find(q('w:body'))
    if body is None:
        raise DocxError(f'缺 w:body：{path}')

    blocks = []
    section, section_title = FRONT, FRONT
    heads = [''] * HEADING_DEPTH
    paragraph_index = table_index = body_index = 0

    for node in body:
        # bodyIndex 是 w:body 下的原始子元素序号（含 sectPr），既有的定位与插行
        # 逻辑按它回指原文，不能按「已提取的块」重新计数，否则位置会整体偏移。
        body_index += 1
        name = local(node)
        if name == 'p':
            paragraph_index += 1
            style = attr(node.find(f'{q("w:pPr")}/{q("w:pStyle")}'), 'w:val')
            level = numbering.outline_level(node, style)
            text, pictures = rich_text(node, media)
            category = 'toc' if TOC_STYLE.search(style or '') else 'body'
            heading = bool(level) and category != 'toc'
            if heading and text:
                depth = int(level)
                heads[depth] = text
                for below in range(depth + 1, HEADING_DEPTH):
                    heads[below] = ''
                section_title = text
                section = ' / '.join(value for value in heads if value)
            if text == REVISION:
                section, section_title = REVISION, text
            if section == REVISION:
                category = 'revision'
            if not text:
                continue
            blocks.append({
                'kind': 'heading' if heading else 'paragraph',
                'text': text, 'number': numbering.paragraph_number(node), 'style': style,
                'section': section, 'sectionTitle': section_title, 'category': category,
                'locator': f'word/document.xml:/w:document/w:body/w:p[{paragraph_index}]',
                'bodyIndex': body_index, 'images': pictures,
            })
        elif name == 'tbl':
            table_index += 1
            blocks.extend(_table_blocks(node, table_index, body_index, section, section_title,
                                        category_of(section), numbering, media))
        elif name == 'sectPr':
            continue
        elif with_text:
            text = text_of(node)
            if text:
                blocks.append({
                    'kind': 'opaque', 'text': text, 'number': '', 'style': '',
                    'section': section, 'sectionTitle': section_title, 'category': 'front',
                    'locator': f'word/document.xml:/w:document/w:body/*[{body_index}]',
                    'bodyIndex': body_index, 'images': [],
                })

    warnings = _warnings(body, media)
    ancillary = _ancillary(package)
    for item in ancillary:
        warnings.append(f'附属内容 {item["part"]} 已提取，需核对其变化。')

    return {
        'schemaVersion': 1,
        'file': str(package.path),
        'sha256': package.sha256(),
        'metadata': _metadata(blocks),
        'blocks': blocks,
        'warnings': warnings,
        'ancillary': ancillary,
        'numberingCaveat': '自动编号按 OOXML 样式与编号定义重建；复杂重启、章节关联编号需结合原文人工核对。',
    }


def _table_blocks(table, table_index, body_index, section, section_title, category,
                  numbering, media):
    props = table.find(q('w:tblPr'))
    layout = attr(props.find(q('w:tblLayout')), 'w:type') if props is not None else ''
    layout = 'fixed' if layout == 'fixed' else 'autofit'
    grid = [int(attr(col, 'w:w') or 0) for col in table.findall(f'{q("w:tblGrid")}/{q("w:gridCol")}')]
    out = []
    for row_index, row in enumerate(table.findall(q('w:tr')), start=1):
        texts, pictures, spans, merged = [], [], [], []
        for cell in _row_cells(row):
            text, cell_pictures = _cell_rich(cell, media, numbering)
            texts.append(text)
            pictures.append(cell_pictures)
            span_node = cell.find(f'{q("w:tcPr")}/{q("w:gridSpan")}')
            spans.append(int(attr(span_node, 'w:val') or 1))
            merged.append(_vmerge(cell))
        joined = ' | '.join(texts)
        if not joined.strip(' |'):
            continue
        out.append({
            'kind': 'tableRow', 'text': joined, 'cells': texts, 'cellImages': pictures,
            'spans': spans, 'vMerge': merged, 'grid': grid, 'layout': layout,
            'number': '', 'style': '',
            'section': section, 'sectionTitle': section_title, 'category': category,
            'tableIndex': table_index, 'rowIndex': row_index,
            'locator': f'word/document.xml:/w:document/w:body/w:tbl[{table_index}]/w:tr[{row_index}]',
            'bodyIndex': body_index, 'images': [],
        })
    return out


def _vmerge(cell):
    """纵向合并状态：restart 是可见首格，continue 是延续格（不是独立可见单元格）。"""
    node = cell.find(f'{q("w:tcPr")}/{q("w:vMerge")}')
    if node is None:
        return ''
    return 'restart' if attr(node, 'w:val') == 'restart' else 'continue'


def _warnings(body, media):
    """把文本比较无法验证的版面内容显式列出，不让工具假装看过。"""
    checks = [
        ('修订痕迹', './/*[local-name()="ins" or local-name()="del" or local-name()="moveFrom" or local-name()="moveTo"]'),
        ('文本框', './/*[local-name()="txbxContent"]'),
        ('公式', './/*[local-name()="oMath"]'),
        ('嵌套表格', './/w:tbl//w:tbl'),
        ('内容控件', './/w:sdt'),
    ]
    out = []
    for label, path in checks:
        count = len(body.xpath(path, namespaces={'w': W}))
        if count:
            out.append(f'{label}：{count} 处；需要人工核对，文字比较不能验证其完整语义或外观。')
    # 只统计最外层绘图节点：图表、形状、OLE 包装各算一处，不重复计嵌套节点。
    outer = body.xpath('.//*[local-name()="drawing" or local-name()="object"]'
                       '[not(ancestor::*[local-name()="drawing" or local-name()="object"])]')
    shapes = len(outer) - media.count
    if shapes > 0:
        out.append(f'图形（非图片）：{shapes} 处；形状与图表中的文字、数值无法由文本比较核实。')
    if media.count or media.files:
        unique = len(media.unique_media())
        out.append(f'图片：{media.count} 处（{unique} 个媒体文件，其中 OLE 对象预览 {media.ole} 处）；'
                   '按内容哈希比较，图中数值与界面文字须人工核对。')
    return out


def scan(path):
    """只报「文本比较验证不了什么」，不产出块模型。

    调用方自己的块形状（例如按 XML 路径定位的扁平段落列表）不需要 docx-model-v1；
    它们真正需要的是同一套判据：哪些版面内容必须人工核对。判据只该有一份实现。
    """
    package = Package(path)
    document = etree.fromstring(package.read('word/document.xml'))
    body = document.find(q('w:body'))
    media = images_mod.Media(package)
    if body is not None:
        # Media 的计数只在解析绘图节点时累加；先遍历一遍，图片统计才不是 0。
        for hit in list(media._hits(body)):
            media._picture(hit)
    ancillary = _ancillary(package)
    warnings = _warnings(body, media) if body is not None else []
    if ancillary:
        # 附属内容的文字能提出来，但它不在正文块序列里；调用方须自行决定是否提示。
        warnings.append('附属内容（%s）不在正文块序列内，相关文字需单独核对。'
                        % '、'.join(item['part'] for item in ancillary))
    return {
        'warnings': warnings,
        'ancillary': ancillary,
        'parts': package.part_names(),
    }


def _ancillary(package):
    """页眉页脚、脚注尾注、批注：提取文字供核对变化。"""
    out = []
    for name in package.part_names():
        if re.match(r'^word/(header|footer|footnotes|endnotes|comments).*\.xml$', name):
            root = etree.fromstring(package.read(name))
            out.append({'part': name, 'text': text_of(root)})
    return out


def _metadata(blocks):
    """文档标识与版本：从全文取第一个匹配，与既有实现一致。"""
    joined = '\n'.join(block['text'] for block in blocks)
    document_id = re.search(r'BDS[0-9A-Z]+[0-9]+', joined)
    version = re.search(r'A/[0-9]+', joined)
    return {
        'documentId': document_id.group(0) if document_id else '',
        'version': version.group(0) if version else '',
        'title': '',
    }
