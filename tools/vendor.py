"""把 office_ooxml_kit 源码内嵌到下游项目，并记录来源指纹。

为什么是「复制」而不是「安装」：
下游项目要单独发布给同事，对方拿到的是一个自包含目录，不能要求他先
`pip install -e D:/project/office-ooxml-kit`——那台机器上根本没有这个路径。
复制之后下游项目不依赖任何外部仓库。

代价是同一份代码存在多份副本，因此这里做两件事防止漂移：
1. 复制时把每个文件的 sha256 写进 VENDOR.json，事后能证明「这份副本来自哪个版本」；
2. --check 模式逐文件比对，副本被改动或源被改动都会报出来。

用法：
    python tools/vendor.py <目标目录>          # 复制并写指纹
    python tools/vendor.py <目标目录> --check  # 只校验，不一致则退出码 1
"""
from __future__ import annotations

import hashlib
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

SOURCE = Path(__file__).resolve().parent.parent
PACKAGE = SOURCE / 'office_ooxml_kit'
MARKER = 'VENDOR.json'
# 只带运行所需的文件；测试、缓存、打包元数据不进下游项目。
INCLUDE = ('*.py', 'schema/*.json')


def wanted_files():
    out = []
    for pattern in INCLUDE:
        for path in sorted(PACKAGE.glob(pattern)):
            if path.is_file() and '__pycache__' not in path.parts:
                out.append(path)
    return out


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def manifest():
    files = {path.relative_to(PACKAGE).as_posix(): digest(path) for path in wanted_files()}
    version = (PACKAGE / '__init__.py').read_text(encoding='utf-8')
    version_line = next(line for line in version.split('\n') if line.startswith('__version__'))
    return {
        'package': 'office_ooxml_kit',
        'version': version_line.split('=')[1].strip().strip("'\""),
        'source': str(SOURCE),
        'files': files,
    }


def check(target):
    marker = target / MARKER
    if not marker.exists():
        print(f'{target} 没有 {MARKER}，不是内嵌副本')
        return 1
    recorded = json.loads(marker.read_text(encoding='utf-8'))
    problems = []
    for name, expected in recorded['files'].items():
        path = target / name
        if not path.exists():
            problems.append(f'缺少 {name}')
        elif digest(path) != expected:
            problems.append(f'内容被改动 {name}')
    current = manifest()['files']
    for name in current:
        if name not in recorded['files']:
            problems.append(f'源已新增 {name}，副本未同步')
    if problems:
        print(f'{target} 与源不一致：')
        for item in problems:
            print(' -', item)
        return 1
    print(f'{target} 与源一致（{len(current)} 个文件，版本 {recorded["version"]}）')
    return 0


def copy(target):
    target.mkdir(parents=True, exist_ok=True)
    for path in wanted_files():
        destination = target / path.relative_to(PACKAGE)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, destination)
    data = manifest()
    data['syncedAt'] = datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')
    (target / MARKER).write_text(
        json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True) + '\n',
        encoding='utf-8', newline='\n')
    print(f'已复制 {len(data["files"])} 个文件到 {target}（版本 {data["version"]}）')
    return 0


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    target = Path(argv[0]).resolve()
    if '--check' in argv:
        return check(target)
    return copy(target)


if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
