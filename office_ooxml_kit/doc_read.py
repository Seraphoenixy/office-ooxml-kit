"""Word 97-2003 二进制文档（.doc）结构与文字提取。

与 docx_read 的关系：本模块产出同一份 docx-model-v1 块模型（schemaVersion=1），
所以评审流程里「按块引用原文」这件事对 .docx 与 .doc 是同一套定位口径。
差异必须在 warnings 里说清楚，不允许悄悄降级：

- 片段表（piece table）取正文文字，8 位片段按 CP1252 解码；中文正文在 Word 中
  以 UTF-16 存储，若 8 位片段出现非 ASCII 字节，说明该段可能用了本地代码页（如 GBK），
  这里显式告警而不是猜。
- 段落样式表（STSH/PAPX）未解析：标题层级、自动编号、章节路径未还原，
  section/sectionTitle 置「未还原」，标题与编号须人工对照原件。
- 修订标记、域代码、文本框、表格边框与合并单元格同样不在文字比较能力内，逐项告警。
- 图片以 0x01 占位，给出固定 token ⟦图:doc⟧（不计算图像哈希）；同名 token 表示
  「此处有图」，不代表两张图相同——这一点在告警里写明。

二进制格式的这些边界是刻意保留的：与其解析一半样式、给出看似精确的章节路径，
不如把「文本能证明什么」和「必须人工看什么」分清楚。
"""
from __future__ import annotations

import hashlib
import struct
from pathlib import Path

from .cfb import CfbError, CompoundFile

__all__ = ['DocError', 'extract', 'extract_text', 'lines', 'metadata', 'scan']

PIC = '⟦图:doc⟧'
OBJECT = '⟦对象⟧'
UNRESTORED = '未还原'
FIELD_BEGIN, FIELD_SEP, FIELD_END = '\x13', '\x14', '\x15'
PARA_MARK, CELL_MARK = '\x0d', '\x07'


class DocError(RuntimeError):
    """不是 Word 97-2003 文档，或关键结构缺失。"""


# ---- 二进制结构 ----

def _fib(word):
    """解析 FIB：返回 ccp 段长、表流名、Clx 位置与关键偏移。"""
    if len(word) < 0x44:
        raise DocError('WordDocument 流过短，不是 Word 文档')
    ident = struct.unpack_from('<H', word, 0)[0]
    if ident != 0xA5EC:
        raise DocError(f'FIB 标识不符（0x{ident:04X}），不是 Word 97+ 文档')
    flags = struct.unpack_from('<H', word, 0x0A)[0]
    csw = struct.unpack_from('<H', word, 0x20)[0]
    lw = 0x20 + 2 + csw * 2
    cslw = struct.unpack_from('<H', word, lw)[0]
    pairs = lw + 2 + cslw * 4
    cb_fclcb = struct.unpack_from('<H', word, pairs)[0]

    def pair(index):
        off = pairs + 2 + index * 8
        if off + 8 > len(word):
            raise DocError('FIB 结构表不完整（文件损坏或版本过旧）')
        return struct.unpack_from('<II', word, off)

    ccp = [struct.unpack_from('<i', word, lw + 2 + i * 4)[0] for i in range(min(cslw, 11))]
    ccp += [0] * (11 - len(ccp))
    return dict(
        nfib=struct.unpack_from('<H', word, 2)[0],
        table_stream='1Table' if flags & 0x0200 else '0Table',
        complex=bool(flags & 0x0004),
        ccp=ccp, ccp_text=max(0, ccp[3]),
        fc_clx=pair(33), fc_min=struct.unpack_from('<I', word, 0x18)[0],
        fc_mac=struct.unpack_from('<I', word, 0x1C)[0], cb_fclcb=cb_fclcb,
    )


def _pieces(fib, table):
    """按 Clx 还原「字符位置 → 文件偏移」的片段表。"""
    fc, lcb = fib['fc_clx']
    if not lcb or fc + lcb > len(table):
        raise DocError('缺少片段表（Clx）；文件可能不是 Word 97+ 或已损坏')
    clx = table[fc:fc + lcb]
    offset, pod = 0, None
    while offset < len(clx):
        kind = clx[offset]
        if kind == 0x01:  # Prc：属性修改，与文字位置无关
            size = struct.unpack_from('<H', clx, offset + 1)[0]
            offset += 3 + size
            continue
        if kind != 0x02:
            raise DocError(f'未知的片段表项 0x{kind:02X}')
        size = struct.unpack_from('<I', clx, offset + 1)[0]
        pod = clx[offset + 5:offset + 5 + size]
        break
    if pod is None:
        return []
    count = (len(pod) - 4) // 12
    cps = struct.unpack_from(f'<{count + 1}I', pod, 0)
    out = []
    for i in range(count):
        base = 4 * (count + 1) + i * 8
        raw_fc = struct.unpack_from('<I', pod, base + 2)[0]
        compressed = bool(raw_fc & 0x40000000)
        pos = raw_fc & 0x3FFFFFFF
        if compressed:
            pos //= 2
        out.append(dict(cp0=cps[i], cp1=cps[i + 1], fc=pos, compressed=compressed))
    return out


def _decode(fib, word, pieces):
    """把片段拼成完整字符流；返回 (字符流, 8位片段中非ASCII字节的首个字符位置)。

    fComplex=0 时 Word 本应给出连续正文，但实际文件（如 WPS 导出的 .doc）仍带
    完整 Clx；片段表存在时按片段走，只有缺 Clx 才退回 fcMin..fcMac 的连续段，
    且连续段一律按 UTF-16LE——Word 97+ 的正文是宽字符，按 8 位解会得到乱码。
    """
    if not pieces:
        raw = word[fib['fc_min']:fib['fc_mac']]
        return raw.decode('utf-16-le', 'replace'), []
    chunks, wide = [], []
    for piece in pieces:
        count = piece['cp1'] - piece['cp0']
        if count <= 0:
            continue
        if piece['compressed']:
            raw = word[piece['fc']:piece['fc'] + count]
            if len(raw) < count:
                raise DocError('文字片段越界（文件损坏）')
            if any(b > 0x7F for b in raw):
                wide.append(piece['cp0'])
            chunks.append(raw.decode('cp1252', 'replace'))
        else:
            raw = word[piece['fc']:piece['fc'] + count * 2]
            if len(raw) < count * 2:
                raise DocError('文字片段越界（文件损坏）')
            chunks.append(raw.decode('utf-16-le', 'replace'))
    return ''.join(chunks), wide


# ---- 字符流到块 ----

STAT_KEYS = ('pictures', 'objects', 'fields', 'linebreaks')


def _clean(text, stats):
    """把段落/单元格内部特殊字符转成可见文本，并累计需人工核对的对象数。"""
    out, index = [], 0
    while index < len(text):
        ch = text[index]
        if ch == FIELD_BEGIN:
            # 域代码（0x13 到 0x14）不是可见文字，只保留域结果（0x14 到 0x15）。
            stats['fields'] += 1
            index += 1
            while index < len(text) and text[index] not in (FIELD_SEP, FIELD_END):
                index += 1
            if index < len(text) and text[index] == FIELD_SEP:
                index += 1
            continue
        if ch in (FIELD_SEP, FIELD_END):
            index += 1
            continue
        if ch == '\x01':
            stats['pictures'] += 1
            out.append(PIC)
        elif ch == '\x08':
            stats['objects'] += 1
            out.append(OBJECT)
        elif ch in '\x0b\x0c\x0e':
            stats['linebreaks'] += 1
            out.append('\n')
        elif ch == '\x1e':
            out.append('-')
        elif ch in '\x1f\x02\x05':
            pass
        elif ch == '\t':
            out.append('\t')
        elif ch == '\n':
            out.append('\n')  # 单元格内多段换行（切分时已写入）
        elif ch in (PARA_MARK, CELL_MARK):
            out.append(' ')
        elif ord(ch) < 0x20:
            pass
        else:
            out.append(ch)
        index += 1
    return ''.join(out)


def _split(fib, text, stats):
    """按字符位置切分段落与表格行。

    二进制格式没有段落容器：段落以 0x0d 结束；单元格以 0x07 结束，行尾在最后一个
    单元格之后再补一个 0x07（相邻两个 0x07 即行结束）；表格结束由随后的段落标记隐含。
    表格结构按这条既定规则重建，规则不足时（例如孤立分隔符）不猜，留给告警与人工核对。
    """
    limit = min(len(text), fib['ccp_text'] or len(text))
    blocks, index = [], 0
    para_index, table_index, row_index = 0, 0, 0
    buf, cells, in_table = [], [], False

    def cell_text(raw):
        return _clean(raw, stats).strip()

    def emit_row():
        nonlocal row_index, cells
        if any(cells):
            blocks.append(dict(kind='tableRow', text=' | '.join(cells),
                               cells=list(cells), tableIndex=table_index,
                               rowIndex=row_index,
                               locator=f'body/t{table_index}/r{row_index}'))
        cells = []
        row_index += 1

    def emit_para():
        nonlocal para_index
        value = cell_text(''.join(buf))
        buf.clear()
        if not value:
            return
        para_index += 1
        blocks.append(dict(kind='paragraph', text=value, locator=f'body/p{para_index}',
                           style=None, paraIndex=para_index))

    while index < limit:
        ch = text[index]
        if ch == CELL_MARK:
            if not in_table:
                in_table, table_index, row_index = True, table_index + 1, 1
                cells = []
            cells.append(cell_text(''.join(buf)))
            buf.clear()
            if index + 1 < limit and text[index + 1] == CELL_MARK:
                # 相邻两个单元格结束符 = 行结束（行尾在最后一个单元格后再补一个 0x07）
                emit_row()
                index += 2
                if index < limit and text[index] == PARA_MARK:
                    in_table, index = False, index + 1  # 表格由随后的段落标记收尾
            else:
                index += 1
            continue
        if ch == PARA_MARK:
            if in_table:
                buf.append('\n')  # 单元格内多段：段内换行，不是行结束
                index += 1
                continue
            emit_para()
            index += 1
            continue
        buf.append(ch)
        index += 1
    if in_table:
        cells.append(cell_text(''.join(buf)))
        buf.clear()
        emit_row()
    emit_para()
    return blocks


def _ancillary(fib, text):
    """页眉页脚与脚注属于子文档：按 ccp 分界取文字，不并入正文块。"""
    out, cursor = [], fib['ccp_text']
    for label, slot in (('脚注', 4), ('页眉页脚', 5)):
        count = max(0, fib['ccp'][slot])
        chunk = text[cursor:cursor + count]
        cursor += count
        if chunk.strip():
            cleaned = _clean(chunk, dict.fromkeys(STAT_KEYS, 0)).strip()
            if cleaned:
                out.append(dict(part=label, text=cleaned))
    return out


def _metadata_from_blocks(blocks):
    """封面标识：编号/版本既可单独成段，也常放在封面表格的相邻单元格里。"""
    document_id, version = '', ''
    window = blocks[:60]
    for n, block in enumerate(window):
        if block['kind'] == 'tableRow':
            cells = [c.strip() for c in block.get('cells') or []]
            for i, cell in enumerate(cells):
                if not document_id and cell in ('编号', '文档编号') and i + 1 < len(cells):
                    document_id = cells[i + 1]
                if not version and cell in ('版本', '版本号') and i + 1 < len(cells):
                    version = cells[i + 1]
        else:
            text = block['text']
            following = window[n + 1]['text'] if n + 1 < len(window) else ''
            if not document_id and text in ('编号', '文档编号'):
                document_id = following
            if not version and text in ('版本', '版本号'):
                version = following
    return dict(documentId=document_id, version=version, title='')


def extract(path):
    """返回 docx-model-v1 兼容块模型；结构不符时抛 DocError。"""
    path = Path(path)
    data = path.read_bytes()
    try:
        cfb = CompoundFile(data)
        word = cfb.read('WordDocument')
    except CfbError as exc:
        raise DocError(str(exc)) from exc
    fib = _fib(word)
    try:
        table = cfb.read(fib['table_stream'])
    except CfbError as exc:
        raise DocError(f'缺少表流：{exc}') from exc
    text, wide = _decode(fib, word, _pieces(fib, table))
    stats = dict.fromkeys(STAT_KEYS, 0)
    blocks = _split(fib, text, stats)
    for block in blocks:
        block.setdefault('category', 'body')
        block.setdefault('section', UNRESTORED)
        block.setdefault('sectionTitle', '')
        block.setdefault('number', '')
        block.setdefault('images', [])
        block.setdefault('bodyIndex', block.get('paraIndex') or block.get('rowIndex') or 0)

    warnings = [
        '二进制 .doc 直接解析：未解析段落样式表，标题层级、自动编号与章节路径未还原，'
        'section 记为「未还原」，标题与编号须对照原件。',
        '修订标记、批注、文本框、表格合并与边框未还原；文字比较不能验证其外观与完整性。',
    ]
    if stats['pictures']:
        warnings.append(f'内嵌图片 {stats["pictures"]} 处：仅以 {PIC} 占位，未提取图像字节或哈希，'
                        '图中数值与界面文字须人工核对；相同占位符不代表图片相同。')
    if stats['objects']:
        warnings.append(f'OLE 或其他嵌入对象 {stats["objects"]} 处，内容未提取。')
    if stats['fields']:
        warnings.append(f'域代码 {stats["fields"]} 处：只保留域结果文字，目录/交叉引用是否为最新无法核实。')
    if stats['linebreaks']:
        warnings.append(f'手动换行/分页/分栏 {stats["linebreaks"]} 处，按换行并入段落文字。')
    if wide:
        warnings.append(f'8 位片段含非 ASCII 字节（首见于字符位置 {wide[0]}），可能按本地代码页'
                        '（如 GBK）存储；本解析按 CP1252 解码，该处文字须人工核对。')
    ancillary = _ancillary(fib, text)
    for item in ancillary:
        warnings.append(f'附属内容「{item["part"]}」已提取，不在正文块序列内，相关文字需单独核对。')

    return {
        'schemaVersion': 1,
        'file': str(path.resolve()),
        'sha256': hashlib.sha256(data).hexdigest().upper(),
        'metadata': _metadata_from_blocks(blocks),
        'blocks': blocks,
        'warnings': warnings,
        'ancillary': ancillary,
        'format': 'doc-97-2003',
    }


def extract_text(path):
    """只要文字：按块返回 (locator, text) 序列，便于检索与引用。"""
    return [(b['locator'], b['text']) for b in extract(path)['blocks']]


def lines(path, limit=60):
    """按块返回前置文字（封面标识识别用，不构造完整块模型）。"""
    return [b['text'] for b in extract(path)['blocks'][:limit]]


def metadata(path):
    return extract(path)['metadata']


def scan(path):
    """只报「文字比较验证不了什么」，与 docx_read.scan 同一套判据口径。"""
    payload = extract(path)
    return dict(warnings=payload['warnings'], category='doc')
