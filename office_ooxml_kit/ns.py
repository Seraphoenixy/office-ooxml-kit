"""OOXML 命名空间常量与元素工具。

写回路径必须用 lxml：标准库 xml.etree 在序列化时会丢弃 `mc:Ignorable` 引用的前缀
（w14/w15/w16…），生成的文件 Word 直接判定为损坏。lxml 原样保留前缀声明。
"""
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
V = 'urn:schemas-microsoft-com:vml'
O = 'urn:schemas-microsoft-com:office:office'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
XML_NS = 'http://www.w3.org/XML/1998/namespace'

# 读写 XML 时统一用这套前缀，避免 lxml 自造 ns0/ns1。
NSMAP = {'w': W, 'r': R, 'wp': WP, 'a': A, 'pic': PIC, 'v': V, 'o': O, 'mc': MC}


def q(tag):
    """`w:p` → Clark 记法，供 lxml 直接使用。"""
    prefix, _, local = tag.partition(':')
    return '{%s}%s' % (NSMAP[prefix], local)


def local(node):
    """元素的本地名，等价于 PowerShell 的 $node.LocalName。

    也接受属性名（Clark 记法字符串），便于统一处理元素与属性。
    """
    if isinstance(node, str):
        return node.rpartition('}')[2]
    tag = node.tag
    if not isinstance(tag, str):
        return ''  # 注释、处理指令
    return tag.rpartition('}')[2]


def el(tag, **attrs):
    """建一个元素；属性名同样用 `w:val` 形式。"""
    node = etree.Element(q(tag))
    for key, value in attrs.items():
        node.set(q(key.replace('__', ':')), str(value))
    return node


def attr(node, tag, default=''):
    """读属性；缺省返回空串，与 PowerShell 的 Attr 语义一致。"""
    if node is None:
        return default
    value = node.get(q(tag))
    return default if value is None else value


def xml_space(node, value='preserve'):
    node.set('{%s}space' % XML_NS, value)


def parse(data):
    """解析 XML 部分并保留原有前缀声明。"""
    return etree.fromstring(data)


def tostring(node, *, xml_declaration=False):
    """序列化。

    绝不能在这里调用 `etree.cleanup_namespaces`：它会删除「看起来没用到」的
    命名空间声明，而 `mc:Ignorable="w14 w15 w16se …"` 恰恰只按前缀名引用它们，
    声明一删 Word 立刻报「文件已经损坏」。要避免 ns0 前缀，应该在建立新节点时
    就传入 `nsmap`（见 `build()`），而不是事后清洗。
    """
    return etree.tostring(node, xml_declaration=xml_declaration,
                          encoding='UTF-8' if xml_declaration else 'unicode',
                          standalone=True if xml_declaration else None)


def build(tag, nsmap=None, **attrs):
    """建一个带命名空间前缀声明的元素，避免 lxml 自造 ns0。"""
    node = etree.Element(q(tag), nsmap=nsmap or NSMAP)
    for key, value in attrs.items():
        node.set(q(key.replace('__', ':')), str(value))
    return node


def standalone_declarations():
    """单元素没有祖先声明时，序列化仍能用标准前缀。"""
    return etree.tostring(etree.Element(q('w:p'), nsmap=NSMAP), encoding='unicode') and dict(NSMAP)
