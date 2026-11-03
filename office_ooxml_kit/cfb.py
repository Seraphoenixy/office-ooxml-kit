"""MS-CFB（Compound File Binary）只读容器。

Word 97-2003 的 .doc 是复合文档：正文、表流、样式表都是容器里的流。
本模块只做「按名字取流」这一件事——不解释流的语义，那是 doc_read 的事。

为什么不用 olefile：本项目不引入第三方依赖（仅有 lxml），而 CFB 的读取面
很小（头、FAT、目录、迷你流），用标准库 struct 实现比新增依赖更可控；
并且解析失败必须报明确原因，而不是依赖库的通用异常。
"""
from __future__ import annotations

import struct

__all__ = ['CompoundFile', 'CfbError']

SIGNATURE = bytes.fromhex('d0cf11e0a1b11ae1')
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
MAXREGSECT = 0xFFFFFFFA


class CfbError(RuntimeError):
    """不是复合文档，或容器结构损坏。"""


class CompoundFile:
    """以 512 字节扇区打开一个复合文档；只提供读取。"""

    def __init__(self, data):
        if len(data) < 512 or data[:8] != SIGNATURE:
            raise CfbError('不是 OLE2 复合文档（签名不符）')
        self.data = data
        shift = struct.unpack_from('<H', data, 0x1E)[0]
        if shift not in (9, 12):
            raise CfbError(f'不支持的扇区大小：2^{shift}')
        self.sector_size = 1 << shift
        self.mini_size = 1 << struct.unpack_from('<H', data, 0x20)[0]
        fat_count = struct.unpack_from('<I', data, 0x2C)[0]
        self.dir_start = struct.unpack_from('<I', data, 0x30)[0]
        self.mini_cutoff = struct.unpack_from('<I', data, 0x38)[0]
        self.minifat_start = struct.unpack_from('<I', data, 0x3C)[0]
        self.minifat_count = struct.unpack_from('<I', data, 0x40)[0]
        difat_start = struct.unpack_from('<I', data, 0x44)[0]
        difat_count = struct.unpack_from('<I', data, 0x48)[0]

        fat_sectors = self._difat(fat_count, difat_start, difat_count)
        self.fat = self._read_fat(fat_sectors)
        self.entries = self._read_directory()
        root = self.entries[0]
        self.mini_stream = self._chain_bytes(root['start'], root['size'], self.fat,
                                            self.sector_size, self.data, 512)
        self.ministream = self.mini_stream
        self.minifat = self._read_minifat()

    # ---- 底层读取 ----

    def _sector(self, index):
        off = 512 + index * self.sector_size
        return self.data[off:off + self.sector_size]

    def _difat(self, fat_count, start, count):
        entries = list(struct.unpack_from('<109I', self.data, 0x4C))
        entries = [e for e in entries if e <= MAXREGSECT]
        per = self.sector_size // 4 - 1
        index = start
        seen = set()
        while index <= MAXREGSECT and index not in seen and len(entries) < fat_count:
            seen.add(index)
            block = self._sector(index)
            values = struct.unpack_from(f'<{self.sector_size // 4}I', block, 0)
            entries.extend(v for v in values[:per] if v <= MAXREGSECT)
            index = values[per]
        if len(entries) < fat_count:
            raise CfbError(f'FAT 扇区表不完整：应有 {fat_count} 项，实得 {len(entries)}')
        return entries[:fat_count]

    def _read_fat(self, sectors):
        entries = []
        for s in sectors:
            block = self._sector(s)
            entries.extend(struct.unpack_from(f'<{self.sector_size // 4}I', block, 0))
        return entries

    @staticmethod
    def _chain_bytes(start, size, fat, sector_size, data, base):
        """沿 FAT 链取流；size=None 表示读到链尾（目录用）。"""
        out = bytearray()
        index = start
        seen = set()
        while index <= MAXREGSECT:
            if index in seen:
                raise CfbError('扇区链成环（文件损坏）')
            seen.add(index)
            off = base + index * sector_size
            out.extend(data[off:off + sector_size])
            if index >= len(fat):
                raise CfbError('扇区链越界（文件损坏）')
            index = fat[index]
            if size is not None and len(out) >= size:
                break
        if size is not None and len(out) < size:
            raise CfbError(f'流数据不足：需要 {size} 字节，实得 {len(out)}')
        return bytes(out if size is None else out[:size])

    def _read_directory(self):
        raw = self._chain_bytes(self.dir_start, None, self.fat, self.sector_size,
                                self.data, 512)
        entries = []
        for off in range(0, len(raw), 128):
            chunk = raw[off:off + 128]
            if len(chunk) < 128:
                break
            name_len = struct.unpack_from('<H', chunk, 0x40)[0]
            kind = chunk[0x42]
            if kind == 0:
                continue
            name = chunk[:max(0, name_len - 2)].decode('utf-16-le', 'replace')
            start = struct.unpack_from('<I', chunk, 0x74)[0]
            size = struct.unpack_from('<Q', chunk, 0x78)[0]
            entries.append(dict(name=name, kind=kind, start=start, size=size))
        if not entries:
            raise CfbError('目录为空（文件损坏）')
        return entries

    def _read_minifat(self):
        if self.minifat_start > MAXREGSECT:
            return []
        raw = self._chain_bytes(self.minifat_start, self.minifat_count * self.sector_size,
                                self.fat, self.sector_size, self.data, 512)
        return list(struct.unpack_from(f'<{len(raw) // 4}I', raw, 0))

    # ---- 对外 ----

    def stream_names(self):
        return [e['name'] for e in self.entries if e['kind'] == 2]

    def read(self, name):
        for entry in self.entries:
            if entry['kind'] == 2 and entry['name'] == name:
                if entry['size'] < self.mini_cutoff:
                    return self._chain_bytes(entry['start'], entry['size'], self.minifat,
                                             self.mini_size, self.ministream, 0)
                return self._chain_bytes(entry['start'], entry['size'], self.fat,
                                         self.sector_size, self.data, 512)
        raise CfbError(f'缺少流：{name}（现存：{", ".join(self.stream_names())}）')
