"""端到端等价比对：用 office_ooxml_kit 重放既有运行目录的 fill.json，
把结果与当初由 docx.ps1 生成的交付物逐项比对。

比 sha256 没有意义（zip 内时间戳会变），所以比的是契约层面的东西：
块模型（含图片字段）、包清单里除可改部分外的内容、页面几何。
"""
import json
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from office_ooxml_kit.docx_read import extract
from office_ooxml_kit.docx_write import fill
from office_ooxml_kit.verify import check, inspect

# 运行目录由命令行指定，缺省取已验证过的那一份交付物。
RUN = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
    'D:/project/form_generate/runs/BDS2230DG007_FPGA-ZX2230总线数据传输电子硬件测试规程_A1-to-A2')


def canon_image(image):
    return tuple((key, image.get(key)) for key in ('token', 'sha256', 'cx', 'cy', 'ext', 'role', 'part'))


def canon_blocks(model):
    out = []
    for block in model['blocks']:
        rows = []
        for key in ('kind', 'text', 'number', 'section', 'sectionTitle', 'category',
                    'locator', 'bodyIndex', 'tableIndex', 'rowIndex', 'cells', 'spans', 'grid', 'layout'):
            if key in block:
                rows.append((key, block[key]))
        rows.append(('cellImages', [canon_image(i) for c in (block.get('cellImages') or []) for i in c]))
        rows.append(('images', [canon_image(i) for i in (block.get('images') or [])]))
        out.append(rows)
    return out


def main():
    manifest = json.loads((RUN / 'manifest.json').read_text(encoding='utf-8'))
    data = json.loads((RUN / 'fill.json').read_text(encoding='utf-8'))
    delivered = Path(manifest['output'])
    tmp = Path(tempfile.mkdtemp(prefix='e2e-'))
    regenerated = tmp / delivered.name
    print('重放生成…')
    result = fill(manifest['template'], regenerated, data)
    print(json.dumps(result, ensure_ascii=False))

    before, after = extract(delivered), extract(regenerated)
    a, b = canon_blocks(before), canon_blocks(after)
    diffs = []
    if len(a) != len(b):
        diffs.append(('block count', len(a), len(b)))
    for index, (x, y) in enumerate(zip(a, b)):
        if x != y:
            for (kx, vx), (ky, vy) in zip(x, y):
                if (kx, vx) != (ky, vy):
                    diffs.append((index, kx, str(vx)[:60], str(vy)[:60]))
    print(f'块数 {len(a)} vs {len(b)}，差异 {len(diffs)} 处')
    for row in diffs[:8]:
        print('   ', row)

    report = check(regenerated, baseline=inspect(manifest['template']))
    print('结构校验:', '通过' if report['passed'] else '未通过', report['preserveOnlyPartChanges'][:5])

    ps = subprocess.run(['pwsh', '-NoProfile', '-Command',
                         '$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0;'
                         'foreach($f in @("%s","%s")){try{$d=$w.Documents.Open($f,$false,$true);'
                         'Write-Output ("OK "+(Split-Path $f -Leaf)+" pages="+$d.ComputeStatistics(2)+" words="+$d.ComputeStatistics(0));'
                         '$d.Close($false)}catch{Write-Output ("FAILED "+$f+": "+$_.Exception.Message)}};$w.Quit()'
                         % (delivered, regenerated)],
                        capture_output=True, text=True, errors='replace')
    print((ps.stdout or '').strip())
    return 0 if not diffs and report['passed'] else 1


if __name__ == '__main__':
    sys.exit(main())
