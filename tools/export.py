"""把仓库导出成可发布的普通文件目录（内嵌子模块内容，不含 .git）。

**为什么需要导出而不是让同事 clone 带子模块的仓库：**
git submodule 的一个硬代价是：普通 `git clone`（不带 --recurse-submodules）之后，
子模块目录是空的，Python 直接 `ModuleNotFoundError`，工具跑不起来。
发布给同事的形态必须是「普通目录 + 完整代码」，而不是「仓库 + 一个需要额外命令才能填充的空洞」。

`git archive` 本身**不支持** --recurse-submodules（Git 2.56 仍不支持），
所以导出分两步：主仓库 archive 出来，再把每个子模块自己的 archive 填进对应路径。
子模块的归档来自该子模块当前 checkout 的 HEAD，而 HEAD 正是主仓库 `git ls-tree` 锁定的那个 commit，
两者一致时导出结果才等于"仓库所声明的版本"——不一致就报错，不导出。

用法：
    python tools/export.py <仓库路径> <输出目录> [--force]
    python tools/export.py <仓库路径> <输出目录> --check   # 只比对，不写
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tarfile
import tempfile
from pathlib import Path


class ExportError(Exception):
    pass


def git(*args, cwd):
    result = subprocess.run(['git', *args], cwd=str(cwd), capture_output=True, text=True,
                            encoding='utf-8', errors='replace')
    if result.returncode != 0:
        raise ExportError(f'git {" ".join(args)} 失败：{result.stderr.strip()}')
    return result.stdout


def submodules(repo):
    """读取 .gitmodules，返回 [(路径, url)]；没有子模块时返回空列表。"""
    path = repo / '.gitmodules'
    if not path.is_file():
        return []
    out = git('config', '--file', str(path), '--get-regexp', r'submodule\..*\.(path|url)', cwd=repo)
    entries, current = [], {}
    for line in out.splitlines():
        key, _, value = line.partition(' ')
        name, _, field = key.rpartition('.')
        if field == 'path':
            current = {'name': name.split('.', 1)[1], 'path': value}
        elif field == 'url':
            current['url'] = value
            if 'path' in current:
                entries.append((current['path'], current['url']))
            current = {}
    return entries


def pinned_commit(repo, path):
    """主仓库为子模块记录的 commit（即 gitlink 的 sha）。"""
    out = git('ls-tree', 'HEAD', '--', path, cwd=repo).strip()
    if not out:
        raise ExportError(f'HEAD 中没有子模块记录：{path}')
    meta, _, _ = out.partition('\t')
    mode, kind, sha = meta.split()
    if kind != 'commit':
        raise ExportError(f'{path} 不是子模块（kind={kind}）')
    return sha


def head_of(path):
    return git('rev-parse', 'HEAD', cwd=path).strip()


def export_submodule(workdir, sub_path, sub_repo_dir, target):
    """把子模块内容从它自己的仓库导出到 target，校验与主仓库锁定的一致。"""
    pinned = pinned_commit(workdir, sub_path)
    actual = head_of(sub_repo_dir)
    if pinned != actual:
        raise ExportError(f'子模块 {sub_path} 的 checkout（{actual[:8]}）与主仓库锁定的 '
                          f'{pinned[:8]} 不一致，拒绝导出——导出会得到一个仓库从未声明过的版本')
    tar = Path(tempfile.mkdtemp(prefix='export-')) / 'sub.tar'
    try:
        with tar.open('wb') as handle:
            subprocess.run(['git', 'archive', '--format=tar', f'--output={handle.name}', 'HEAD'],
                           cwd=str(sub_repo_dir), check=True, capture_output=True)
        with tarfile.open(tar) as archive:
            archive.extractall(target, filter='data')
    finally:
        shutil.rmtree(tar.parent, ignore_errors=True)
    return pinned


def run(repo, out, force=False):
    repo, out = Path(repo).resolve(), Path(out).resolve()
    if not (repo / '.git').exists():
        raise ExportError(f'不是 git 仓库：{repo}')
    # HEAD 必须存在：导出内容来自 HEAD，没有提交就没有可导出的东西。
    if subprocess.run(['git', 'rev-parse', '--verify', 'HEAD'], cwd=str(repo),
                      capture_output=True).returncode != 0:
        raise ExportError(f'{repo} 还没有任何提交，无法导出——先提交一次再发布')

    # 所有前置条件先校验完再动磁盘：半成品输出目录比没有输出更糟，
    # 发布时会有人把它当成成功产物。
    subs = submodules(repo)
    for sub_path, _url in subs:
        sub_repo_dir = repo / sub_path
        if not sub_repo_dir.exists():
            raise ExportError(f'子模块未检出：{sub_repo_dir}（先 git submodule update --init）')
        pinned = pinned_commit(repo, sub_path)
        actual = head_of(sub_repo_dir)
        if pinned != actual:
            raise ExportError(f'子模块 {sub_path} 的 checkout（{actual[:8]}）与主仓库锁定的 '
                              f'{pinned[:8]} 不一致，拒绝导出——导出会得到一个仓库从未声明过的版本')

    if out.exists() and any(out.iterdir()):
        if not force:
            raise ExportError(f'输出目录非空：{out}（用 --force 覆盖）')
        shutil.rmtree(out)
    out.mkdir(parents=True, exist_ok=True)

    tar = Path(tempfile.mkdtemp(prefix='export-')) / 'main.tar'
    try:
        with tar.open('wb') as handle:
            subprocess.run(['git', 'archive', '--format=tar', f'--output={handle.name}', 'HEAD'],
                           cwd=str(repo), check=True, capture_output=True)
        with tarfile.open(tar) as archive:
            archive.extractall(out, filter='data')
    finally:
        shutil.rmtree(tar.parent, ignore_errors=True)

    exported = []
    for sub_path, url in subs:
        sub_repo_dir = repo / sub_path
        target = out / sub_path
        if target.exists():
            shutil.rmtree(target)
        target.mkdir(parents=True)
        exported.append((sub_path, export_submodule(repo, sub_path, sub_repo_dir, target), url))
    return out, exported


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('repo')
    parser.add_argument('out')
    parser.add_argument('--force', action='store_true', help='输出目录非空时覆盖')
    args = parser.parse_args(argv)
    try:
        out, exported = run(args.repo, args.out, args.force)
    except ExportError as error:
        print(f'导出失败：{error}', file=sys.stderr)
        return 1
    print(f'已导出到 {out}')
    for sub_path, sha, url in exported:
        print(f'  子模块 {sub_path} ← {sha[:8]}  ({url})')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
