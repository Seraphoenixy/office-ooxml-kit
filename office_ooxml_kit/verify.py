"""包与文档校验：给 CR 流程提供「这份文件里到底有什么」的客观证据。

校验只报告事实，不做判断：是不是通过由调用方的规则决定。这样同一份证据既能用于
生成前后比对，也能用于人工复核。
"""
from __future__ import annotations

import hashlib
import re
from pathlib import Path

from lxml import etree

from .ns import W, local, q, tostring
from .package import Package

__all__ = ['inspect', 'VerifyError', 'MUTABLE_PARTS', 'PRESERVED_PREFIX']


class VerifyError(RuntimeError):
    """文件不是可读的 OOXML 包。"""


MUTABLE_PARTS = frozenset({
    'word/document.xml',
    'word/settings.xml',
    'docProps/core.xml',
    '[Content_Types].xml',
    'word/_rels/document.xml.rels',
})

# 页眉页脚本来就允许被替换旧单信息，不该按「必须逐字节不变」对待。
PRESERVED_PREFIX = re.compile(r'^word/(header|footer).*\.xml$')


def inspect(path):
    """包清单 + 文档分区几何 + 外层表格 XML，供生成前后比对。

    这三个维度是 CR 校验实际依赖的：包清单能证明「除约定部分外没有别的部分被改」，
    sectPr 能证明页面设置没变，表格 XML 能证明结构合法（命名空间、表头未被拆散）。
    """
    path = Path(path)
    if not path.exists():
        raise VerifyError(f'文件不存在：{path}')
    package = Package(path)
    document = etree.fromstring(package.read('word/document.xml'))
    parts = [
        {'name': name, 'size': len(package.read(name)), 'sha256': package.hash(name)}
        for name in package.part_names()
    ]
    body = document.find(q('w:body'))
    tables = [tostring(table) for table in body.findall(q('w:tbl'))] if body is not None else []
    sections = [tostring(node) for node in document.iter(q('w:sectPr'))]
    return {
        'sha256': _file_hash(path),
        'parts': parts,
        'sections': sections,
        'tables': tables,
    }


def compare_parts(before, after):
    """除约定可改的部分外，原有部分必须逐字节不变。

    这是「只追加内容、不重排原有结构」的证明：媒体文件重排编号不算变化，
    但某个样式表被改写就是真变化。

    只比较**两边都存在**的部分：填充图片时会新增 word/media/crN.png 这类条目，
    新增文件不是「原有部分被改写」。把新增当成违规会让每次带图的生成都失败。
    """
    after_map = {part['name']: part['sha256'] for part in after.get('parts') or []}
    kept = [part for part in before.get('parts') or []
            if part['name'] not in MUTABLE_PARTS and not PRESERVED_PREFIX.match(part['name'])]
    return [part['name'] for part in kept
            if part['name'] in after_map and after_map[part['name']] != part['sha256']]


def check(path, *, baseline=None, fill=None):
    """对生成结果做结构校验，返回逐项结论。"""
    current = inspect(path)
    report = {
        'file': str(path),
        'sha256': current['sha256'],
        'singleOuterTable': len(current['tables']) == 1,
        'preserveOnlyPartChanges': compare_parts(baseline, current) if baseline else [],
        'sectionGeometryPreserved': (
            json_stable(baseline.get('sections')) == json_stable(current['sections'])
            if baseline else True),
        'xmlNamespaceValid': _namespace_valid(current['tables']),
        'literalImageTokensInWord': any('⟦图:' in table for table in current['tables']),
        'note': '结构校验不能代替逐页视觉检查。',
    }
    report['passed'] = all((
        report['singleOuterTable'],
        not report['preserveOnlyPartChanges'],
        report['sectionGeometryPreserved'],
        report['xmlNamespaceValid'],
        not report['literalImageTokensInWord'],
    ))
    return report


def _namespace_valid(tables):
    """重复声明内建 xml 命名空间会让 Word 拒绝打开。"""
    pattern = re.compile(r'xmlns:(?!xml=)[^=\s]+="http://www\.w3\.org/XML/1998/namespace"')
    return not any(pattern.search(table) for table in tables)


def _find_part(parts, name):
    for part in parts:
        if part['name'] == name:
            return part['sha256']
    return None


def _file_hash(path):
    digest = hashlib.sha256()
    with open(path, 'rb') as handle:
        for chunk in iter(lambda: handle.read(65536), b''):
            digest.update(chunk)
    return digest.hexdigest().upper()


def json_stable(value):
    """把结构压成可直接比较的字符串。"""
    import json
    return json.dumps(value, ensure_ascii=False, sort_keys=True)
