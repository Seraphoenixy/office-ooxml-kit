"""精确归因：按文档顺序取文字，证明差异只来自公式（m:t）。

对每个有差异的单元格同时算两种文字：
  A = 只取 w:t（PowerShell 的口径）
  B = 按文档顺序取 w:t 与 m:t（Python 的口径）
若 ref_cell == A 且 mine_cell == B，则差异完全由公式造成，这是可证伪的判据。
"""
import sys
from pathlib import Path
import zipfile

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tmp_corpus

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def cell_texts(cell):
    """w:t 与 m:t 的局部名都是 t，必须按命名空间 URI 区分，否则会把公式算成正文字。"""
    plain, both = [], []
    for node in cell.iter():
        if node.tag == '{%s}t' % W:
            plain.append(node.text or '')
            both.append(node.text or '')
        elif node.tag == '{%s}t' % M:
            both.append(node.text or '')
    return ''.join(plain), ''.join(both)


def row_cells(path):
    """{(tableIndex, rowIndex): 每格的 (只w:t, w:t+m:t)}"""
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read('word/document.xml'))
    out = {}
    for table_index, table in enumerate(root.xpath('//w:tbl', namespaces=NS), 1):
        for row_index, row in enumerate(table.xpath('./w:tr', namespaces=NS), 1):
            out[(table_index, row_index)] = [
                cell_texts(c) for c in row.xpath('./w:tc', namespaces=NS)]
    return out


def flat(text):
    return ' '.join(str(text or '').split())


def classify(mine, ref, cells):
    if len(mine.get('cells') or []) != len(cells):
        return 'UNKNOWN', '单元格数不同'
    for a, b, (plain, both) in zip(mine['cells'], ref['cells'], cells):
        if flat(a) == flat(b):
            continue
        if flat(b) != flat(plain):
            return 'UNKNOWN', f'ref 不是纯 w:t：{flat(b)!r} != {flat(plain)!r}'
        if flat(a) != flat(both):
            return 'UNKNOWN', f'mine 不是 w:t+m:t：{flat(a)!r} != {flat(both)!r}'
    return 'math', '差异完全来自公式文字'


def main(files):
    tally = {}
    problems = []
    for path in files:
        ours = tmp_corpus.extract(path)
        theirs = tmp_corpus.ps_extract(path)
        rows = row_cells(path)
        for index, (mine, ref) in enumerate(zip(ours['blocks'], theirs['blocks'])):
            if mine.get('number') != ref.get('number'):
                # 编号差异已由 Word COM 的 ListString 裁定（Python 与 Word 一致）。
                tally['numbering'] = tally.get('numbering', 0) + 1
                continue
            if flat(mine.get('text')) == flat(ref.get('text')) and mine.get('kind') == ref.get('kind'):
                continue
            cells = rows.get((mine.get('tableIndex'), mine.get('rowIndex')))
            if cells is None:
                problems.append((path, index, '不是表格行'))
                tally['UNKNOWN'] = tally.get('UNKNOWN', 0) + 1
                continue
            kind, detail = classify(mine, ref, cells)
            tally[kind] = tally.get(kind, 0) + 1
            if kind != 'math':
                problems.append((path, index, detail))
    print('归因：', tally)
    for row in problems:
        print('  未解释：', row)
    return 0 if problems else 1


if __name__ == '__main__':
    paths = [line for line in open('tmp_targets.txt', encoding='utf-8').read().split('\n') if line.strip()]
    raise SystemExit(main(paths))
