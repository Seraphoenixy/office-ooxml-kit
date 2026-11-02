"""图片与媒体：按内容哈希识别，而不是按文件名。

Word 升版时会重排媒体编号（image7.emf → image9.emf）。按文件名比较会把整篇图片
都报成变化，按内容比较才看得出「这次真的换了一张图」。OLE 对象只取预览图。
"""
from __future__ import annotations

import hashlib
import re

from lxml import etree

from .ns import CT, R, local, q
from .package import MIME_BY_EXT, REL_IMAGE, normalize_part_target

__all__ = ['Media', 'TOKEN_RE', 'token_of', 'readable', 'split_tokens']

# 提取结果里用 ⟦图:哈希前8位⟧ 表示图片；由内容哈希生成，媒体重编号不影响比较。
TOKEN_RE = re.compile(r'⟦图:([0-9A-Fa-f]{8}|外链)⟧')
DRAWING_XPATH = './/*[local-name()="drawing" or local-name()="object"]'


def token_of(digest: str) -> str:
    return '⟦图:%s⟧' % digest[:8].upper()


def readable(text) -> str:
    """给人看的记录里不要把 token 原样贴出来。"""
    return TOKEN_RE.sub(lambda m: '〔图片 %s〕' % m.group(1), str(text or ''))


def split_tokens(text):
    """把含图片的文字切成 文本/图片 段；纯图片返回单段。"""
    body = str(text or '')
    parts, last = [], 0
    for match in TOKEN_RE.finditer(body):
        if match.start() > last:
            parts.append({'type': 'text', 'text': body[last:match.start()]})
        parts.append({'type': 'image', 'token': match.group(0)})
        last = match.end()
    if last < len(body):
        parts.append({'type': 'text', 'text': body[last:]})
    return parts


class Media:
    """一份文档的媒体清单与关系项解析。"""

    def __init__(self, package, *, part_prefix='cr', rel_prefix='rIdCr'):
        self.package = package
        # 嵌入图片的部件名前缀。默认 cr 与既有流程一致；换项目可改成自己的命名。
        self.part_prefix = part_prefix
        self.rel_prefix = rel_prefix
        self.files = {}   # part 名 → {part, sha256, ext, bytes, data}
        self.count = 0    # 引用处数
        self.ole = 0      # OLE 对象预览计数
        self._rels = {}
        self._load_rels()

    def _load_rels(self):
        raw = self.package.optional('word/_rels/document.xml.rels')
        if raw is None:
            return
        root = etree.fromstring(raw)
        for node in root.iter():
            if local(node) != 'Relationship':
                continue
            self._rels[node.get('Id')] = {
                'target': node.get('Target') or '',
                'external': (node.get('TargetMode') or '') == 'External',
                'type': node.get('Type') or '',
            }

    def rel(self, rel_id):
        return self._rels.get(rel_id)

    def file_of(self, rel_id):
        """关系项 → 媒体文件描述；外部链接与缺文件返回 None。"""
        rel = self._rels.get(rel_id or '')
        if not rel or rel['external']:
            return None
        part = normalize_part_target('word', rel['target'])
        if part in self.files:
            return self.files[part]
        data = self.package.optional(part)
        if data is None:
            return None
        ext = part.rsplit('.', 1)[-1].lower() if '.' in part else ''
        self.files[part] = {
            'part': part,
            'sha256': hashlib.sha256(data).hexdigest().upper(),
            'ext': ext,
            'bytes': len(data),
            'data': data,
        }
        return self.files[part]

    # ---- 引用 ----
    def _hits(self, node):
        """只取最外层绘图节点，避免嵌套结构重复计数。"""
        for hit in node.xpath(DRAWING_XPATH):
            if hit.xpath('ancestor::*[local-name()="drawing" or local-name()="object"]'):
                continue
            yield hit

    def _picture(self, node):
        """一个绘图节点对应的图片信息；识别不出来时返回 None。"""
        if local(node) == 'object':
            preview = node.find('.//{%s}imagedata' % 'urn:schemas-microsoft-com:vml')
            if preview is None:
                return None
            info = self.file_of(preview.get('{%s}id' % R))
            if info is None:
                return None
            self.count += 1
            self.ole += 1
            shape = node.find('.//{%s}shape' % 'urn:schemas-microsoft-com:vml')
            height, width = _shape_emu(shape.get('style') if shape is not None else '')
            ole = node.find('.//{%s}OLEObject' % 'urn:schemas-microsoft-com:office:office')
            ole_part = ''
            if ole is not None:
                info_ole = self.file_of(ole.get('{%s}id' % R))
                ole_part = info_ole['part'] if info_ole else ''
            # 键名沿用旧 PowerShell 实现的既有模型：消费方读的是 sha256/cx/cy，
            # 自造 hash/width/height 会让它们静默取到 undefined。
            return {'role': 'olePreview', 'token': token_of(info['sha256']),
                    'sha256': info['sha256'], 'part': info['part'],
                    'ext': info['ext'], 'bytes': info['bytes'],
                    'cx': width, 'cy': height, 'olePart': ole_part,
                    'name': '', 'alt': ''}

        blip = node.find('.//{%s}blip' % 'http://schemas.openxmlformats.org/drawingml/2006/main')
        if blip is None:
            return None
        rel_id = blip.get('{%s}embed' % R) or ''
        rel = self._rels.get(rel_id)
        if rel and rel['external']:
            return {'role': 'picture', 'token': '⟦图:外链⟧', 'sha256': '',
                    'part': rel['target'], 'ext': '', 'bytes': 0,
                    'cx': 0, 'cy': 0, 'name': '', 'alt': '',
                    'external': True}
        info = self.file_of(rel_id)
        if info is None:
            return None
        extent = node.find('.//{%s}extent' % 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
        height = int(extent.get('cy') or 0) if extent is not None else 0
        width = int(extent.get('cx') or 0) if extent is not None else 0
        doc_pr = node.find('.//{%s}docPr' % 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
        self.count += 1
        return {'role': 'picture', 'token': token_of(info['sha256']),
                'sha256': info['sha256'], 'part': info['part'],
                'ext': info['ext'], 'bytes': info['bytes'],
                'cx': width, 'cy': height, 'olePart': '',
                'name': (doc_pr.get('name') or '') if doc_pr is not None else '',
                'alt': (doc_pr.get('descr') or '') if doc_pr is not None else ''}

    def paragraph_images(self, paragraph):
        return [pic for pic in (self._picture(hit) for hit in self._hits(paragraph)) if pic]

    def cell_images(self, cell):
        return [pic for pic in (self._picture(hit) for hit in self._hits(cell)) if pic]

    def unique_media(self):
        return {info['part']: info for info in self.files.values()
                if info['part'].startswith('word/media/')}

    # ---- 写入 ----
    def embed(self, out_package, data, ext, *, seq, dry_run=False):
        """把图片字节加进输出包，返回可直接用于 a:blip/@r:embed 的关系 Id。

        同时补 `[Content_Types].xml` 的 Default 声明——漏了这一步 Word 会报损坏。
        """
        digest = hashlib.sha256(data).hexdigest().upper()
        # 部件名沿用调用方约定：既有流程（以及它的校验）按 word/media/cr<N>.<ext>
        # 认临时嵌入的图片，自造前缀会让「图片是否真的进包了」判断失效。
        part = self.part_name(seq, ext)
        if not dry_run and not out_package.has(part):
            out_package.add(part, data)
            _ensure_content_type(out_package, ext)
        rel_id = '%s%d' % (self.rel_prefix, seq)
        _ensure_relationship(out_package, rel_id, part)
        return rel_id, part, digest

    def part_name(self, seq, ext):
        return 'word/media/%s%d.%s' % (self.part_prefix, seq, ext)


def _ensure_content_type(package, ext):
    raw = package.optional('[Content_Types].xml')
    if raw is None:
        return
    root = etree.fromstring(raw)
    mime = MIME_BY_EXT.get(ext)
    if not mime:
        return
    existing = {node.get('Extension') for node in root.iter() if local(node) == 'Default'}
    if ext in existing:
        return
    node = etree.SubElement(root, '{%s}Default' % CT)
    node.set('Extension', ext)
    node.set('ContentType', mime)
    package.set('[Content_Types].xml',
                etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True))


def _ensure_relationship(package, rel_id, part):
    """向 document.xml.rels 追加图片关系；Id 冲突时保留原项。"""
    raw = package.optional('word/_rels/document.xml.rels')
    if raw is None:
        return
    root = etree.fromstring(raw)
    if any(node.get('Id') == rel_id for node in root.iter() if local(node) == 'Relationship'):
        return
    node = etree.SubElement(root, '{%s}Relationship' % root.nsmap.get(None, 'http://schemas.openxmlformats.org/package/2006/relationships'))
    node.set('Id', rel_id)
    node.set('Type', REL_IMAGE)
    node.set('Target', part.replace('word/', '', 1))
    package.set('word/_rels/document.xml.rels',
                etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True))


def _shape_emu(style):
    """VML style 里的 pt/in/cm → EMU。"""
    # 与参照实现一致：PowerShell 的 [long] 是四舍五入，int() 截断会少 1 EMU。
    to_emu = lambda value, unit: round(value * {'pt': 12700, 'in': 914400, 'cm': 360000}[unit])  # noqa: E731
    out = [0, 0]
    for axis, index in (('width', 1), ('height', 0)):
        match = re.search(r'%s:\s*([\d.]+)(pt|in|cm)' % axis, style or '')
        if match:
            out[index] = to_emu(float(match.group(1)), match.group(2))
    return out[0], out[1]
