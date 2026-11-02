"""OPC 包读写：保持部分顺序、压缩方式与时间戳；输出拒绝覆盖。

所有项目此前都在各自实现这一层（PowerShell 的 ZipArchive 手工重建、
python-docx 隐式代写）。差异集中在这里：谁补关系项、谁补内容类型、谁保住顺序。
"""
from __future__ import annotations

import hashlib
import re
import zipfile
from pathlib import Path

__all__ = ['Package', 'PackageError', 'REL_IMAGE']


class PackageError(RuntimeError):
    """包层面的错误：缺部件、输出已存在、目标不是 OPC 包。"""


REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

# 新增图片必须同时补内容类型声明，否则 Word 报文件损坏。
MIME_BY_EXT = {
    'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'gif': 'image/gif', 'bmp': 'image/bmp', 'emf': 'image/x-emf',
    'wmf': 'image/x-wmf', 'tif': 'image/tiff', 'tiff': 'image/tiff',
    'svg': 'image/svg+xml',
}


def normalize_part_target(base_dir: str, target: str) -> str:
    """把关系项 Target 解析成包内部件名，等价于 PowerShell 的 Rel-Info 归一化。"""
    if target.startswith('/'):
        target = target[1:]
    else:
        target = base_dir + '/' + target
    while '/../' in target:
        target = re.sub(r'[^/]+/\.\./', '', target)
    return target


class Package:
    """一个已打开的 OOXML 包；读取全部部件到内存，写回时保持原顺序。

    输入文档通常 2MB 量级，整体读入实测 0.028s，远低于复杂度带来的收益。
    """

    def __init__(self, path):
        self.path = Path(path)
        if not self.path.is_file():
            raise PackageError(f'文件不存在：{self.path}')
        try:
            with zipfile.ZipFile(self.path) as archive:
                self._entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
        except zipfile.BadZipFile as error:
            raise PackageError(f'不是有效的 OOXML 包：{self.path}（{error}）') from error
        self._parts = {info.filename: data for info, data in self._entries}
        self._added = []

    # ---- 读 ----
    def has(self, name: str) -> bool:
        return name in self._parts

    def read(self, name: str) -> bytes:
        try:
            return self._parts[name]
        except KeyError:
            raise PackageError(f'缺少部件：{name}（文件 {self.path.name}）') from None

    def read_text(self, name: str) -> str:
        return self.read(name).decode('utf-8')

    def optional(self, name: str):
        """读取可缺省的部件，缺失返回 None（styles/numbering/settings 都可能没有）。"""
        return self._parts.get(name)

    def part_names(self):
        return [info.filename for info, _ in self._entries] + [name for name, _ in self._added]

    def hash(self, name: str) -> str:
        """单个部件的 SHA-256（大写），用于比对「哪些部分真的变了」。"""
        return hashlib.sha256(self.read(name)).hexdigest().upper()

    def sha256(self) -> str:
        digest = hashlib.sha256()
        with self.path.open('rb') as handle:
            for chunk in iter(lambda: handle.read(1 << 20), b''):
                digest.update(chunk)
        return digest.hexdigest().upper()

    def part_hashes(self):
        """每个部件的 SHA-256，用于判断模板是否被改动。"""
        return [
            {'name': info.filename, 'size': len(data),
             'sha256': hashlib.sha256(data).hexdigest().upper()}
            for info, data in self._entries
        ]

    # ---- 写 ----
    def set(self, name: str, data) -> None:
        """替换已有部件的字节。"""
        if isinstance(data, str):
            data = data.encode('utf-8')
        if name not in self._parts:
            raise PackageError(f'部件不存在，不能覆盖：{name}')
        self._parts[name] = data

    def add(self, name: str, data: bytes) -> None:
        """新增部件（图片等），写回时追加在包尾。"""
        if name in self._parts or any(n == name for n, _ in self._added):
            raise PackageError(f'部件已存在，不能重复添加：{name}')
        self._added.append((name, data))

    def save(self, out_path, overwrite=False) -> Path:
        """写出新包；默认拒绝覆盖已有文件（各项目都在执行这条纪律）。"""
        out = Path(out_path)
        if out.exists() and not overwrite:
            raise PackageError(f'输出已存在，禁止覆盖：{out}')
        out.parent.mkdir(parents=True, exist_ok=True)
        temporary = out.with_name(out.name + '.tmp')
        with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as archive:
            for info, original in self._entries:
                copy = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                copy.compress_type = info.compress_type
                copy.external_attr = info.external_attr
                copy.create_system = info.create_system
                copy.comment = info.comment
                archive.writestr(copy, self._parts.get(info.filename, original))
            for name, data in self._added:
                archive.writestr(zipfile.ZipInfo(name), data)
        temporary.replace(out)
        return out
