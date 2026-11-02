"""语料级比对：Python 提取结果 vs PowerShell docx.ps1 提取结果，逐块逐字段。

之前这一轮比对特意跳过了 images/cellImages，于是图片键名不一致没被发现。
现在图片也纳入比较，字段名与取值都要一致。
"""
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from office_ooxml_kit.docx_read import extract

SHELL = os.environ.get('OFFICE_OOXML_KIT_PWSH', r'C:\Program Files\PowerShell\7\pwsh.exe')
# 参考实现路径；换机器或换项目时用环境变量指定。
SCRIPT = os.environ.get('DOCX_PS1', r'D:\project\office-ooxml-kit\tools\reference\docx.ps1')


def ps_extract(path):
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / 'ps.json'
        subprocess.run([SHELL, '-NoProfile', '-File', SCRIPT, '-Action', 'extract',
                        '-Path', str(path), '-Output', str(out)],
                       check=True, capture_output=True, encoding='utf-8')
        return json.loads(out.read_text(encoding='utf-8'))


def canon_image(image):
    """只看契约字段：sha256/cx/cy/ext/part/token/role。"""
    return tuple((key, image.get(key)) for key in ('token', 'sha256', 'cx', 'cy', 'ext', 'role', 'part'))


def main(targets):
    total_blocks = 0
    problems = 0
    for path in targets:
        path = Path(path)
        ours = extract(path)
        theirs = ps_extract(path)
        mismatches = []
        for index, (mine, ref) in enumerate(zip(ours['blocks'], theirs['blocks'])):
            for key in ('kind', 'text', 'number', 'section', 'sectionTitle', 'category',
                        'locator', 'bodyIndex', 'tableIndex', 'rowIndex', 'cells', 'spans',
                        'grid', 'layout'):
                if key in ref or key in mine:
                    if mine.get(key) != ref.get(key):
                        mismatches.append((index, key, ref.get(key), mine.get(key)))
            mine_cells = [canon_image(i) for c in (mine.get('cellImages') or []) for i in c]
            ref_cells = [canon_image(i) for c in (ref.get('cellImages') or []) for i in c]
            if mine_cells != ref_cells:
                mismatches.append((index, 'cellImages', ref_cells[:2], mine_cells[:2]))
            mine_imgs = [canon_image(i) for i in (mine.get('images') or [])]
            ref_imgs = [canon_image(i) for i in (ref.get('images') or [])]
            if mine_imgs != ref_imgs:
                mismatches.append((index, 'images', ref_imgs[:2], mine_imgs[:2]))
        if len(ours['blocks']) != len(theirs['blocks']):
            mismatches.append(('count', 'blocks', len(theirs['blocks']), len(ours['blocks'])))
        total_blocks += len(ours['blocks'])
        status = 'OK' if not mismatches else f'{len(mismatches)} 差异'
        print(f'{status:>10}  {len(ours["blocks"]):>5} 块  {path.name}')
        for row in mismatches[:6]:
            print('    ', row[0], row[1], repr(row[2])[:80], '!=', repr(row[3])[:80])
        problems += len(mismatches)
    print(f'\n合计 {len(targets)} 份文档、{total_blocks} 块，差异 {problems} 处。')
    return 1 if problems else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
