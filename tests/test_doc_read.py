"""doc_read/cfb 的单元测试：合成容器与纯文本切分，不依赖真实项目文件。"""
from __future__ import annotations

import struct
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from office_ooxml_kit import doc_read  # noqa: E402
from office_ooxml_kit.cfb import CfbError, CompoundFile  # noqa: E402

FATSECT = 0xFFFFFFFD
ENDOFCHAIN = 0xFFFFFFFE


def build_cfb(streams):
    """构造一个只含常规流（≥4096 字节）的最小复合文档，用来验证读取器。"""
    sector = 512
    fat = [FATSECT, ENDOFCHAIN]  # 0:FAT 扇区, 1:目录扇区
    payload = bytearray()
    starts = {}
    for name, data in streams.items():
        assert len(data) >= 4096, '合成用例只覆盖常规流'
        # 流按扇区切分；跳过目录扇区（1）造成的小洞不影响读取。
        count = (len(data) + sector - 1) // sector
        start = len(fat)
        starts[name] = start
        for i in range(count):
            fat.append(ENDOFCHAIN - 1 if i == count - 1 else start + i + 1)
        payload += data.ljust(count * sector, b'\0')

    directory = bytearray()
    entries = [('Root Entry', 5, ENDOFCHAIN, 0)] + [
        (name, 2, starts[name], len(data)) for name, data in streams.items()]
    for name, kind, start, size in entries:
        raw = name.encode('utf-16-le') + b'\0\0'
        chunk = bytearray(128)
        chunk[:len(raw)] = raw
        struct.pack_into('<H', chunk, 0x40, len(raw))
        chunk[0x42] = kind
        struct.pack_into('<I', chunk, 0x44, ENDOFCHAIN)
        struct.pack_into('<I', chunk, 0x48, ENDOFCHAIN)
        struct.pack_into('<I', chunk, 0x4C, ENDOFCHAIN)
        struct.pack_into('<I', chunk, 0x74, start)
        struct.pack_into('<Q', chunk, 0x78, size)
        directory += chunk

    header = bytearray(512)
    header[:8] = bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<H', header, 0x1E, 9)
    struct.pack_into('<H', header, 0x20, 6)
    struct.pack_into('<I', header, 0x2C, 1)
    struct.pack_into('<I', header, 0x30, 1)
    struct.pack_into('<I', header, 0x38, 4096)
    struct.pack_into('<I', header, 0x3C, ENDOFCHAIN)
    struct.pack_into('<I', header, 0x40, 0)
    struct.pack_into('<I', header, 0x44, ENDOFCHAIN)
    struct.pack_into('<I', header, 0x48, 0)
    struct.pack_into('<I', header, 0x4C, 0)  # DIFAT[0] = FAT 扇区
    for i in range(1, 109):
        struct.pack_into('<I', header, 0x4C + i * 4, 0xFFFFFFFF)

    fat_bytes = bytearray(sector)
    for i in range(sector // 4):
        struct.pack_into('<I', fat_bytes, i * 4, fat[i] if i < len(fat) else 0xFFFFFFFF)
    directory_bytes = bytes(directory).ljust(sector, b'\0')

    return bytes(header) + bytes(fat_bytes) + directory_bytes + bytes(payload)


class CfbTests(unittest.TestCase):
    def test_reads_streams(self):
        data = build_cfb({'WordDocument': b'W' * 5000, '1Table': b'T' * 4096})
        cfb = CompoundFile(data)
        self.assertEqual(cfb.read('WordDocument'), b'W' * 5000)
        self.assertEqual(cfb.read('1Table'), b'T' * 4096)
        self.assertIn('WordDocument', cfb.stream_names())

    def test_rejects_non_compound(self):
        with self.assertRaises(CfbError):
            CompoundFile(b'%PDF-1.7 not ole')

    def test_missing_stream_lists_candidates(self):
        with self.assertRaises(CfbError) as ctx:
            CompoundFile(build_cfb({'WordDocument': b'W' * 4096})).read('0Table')
        self.assertIn('WordDocument', str(ctx.exception))


class SplitTests(unittest.TestCase):
    def test_paragraphs(self):
        stats = dict.fromkeys(doc_read.STAT_KEYS, 0)
        blocks = doc_read._split({'ccp_text': 0}, '甲\r乙\r', stats)
        self.assertEqual([b['text'] for b in blocks], ['甲', '乙'])
        self.assertTrue(all(b['kind'] == 'paragraph' for b in blocks))

    def test_table_row_and_following_paragraph(self):
        stats = dict.fromkeys(doc_read.STAT_KEYS, 0)
        text = '前置段落\r甲\x07乙\x07\x07\r后续段落\r'
        blocks = doc_read._split({'ccp_text': 0}, text, stats)
        self.assertEqual(blocks[0]['text'], '前置段落')
        self.assertEqual(blocks[1]['kind'], 'tableRow')
        self.assertEqual(blocks[1]['cells'], ['甲', '乙'])
        self.assertEqual(blocks[2]['text'], '后续段落')

    def test_paragraph_inside_cell_is_newline_not_row_break(self):
        stats = dict.fromkeys(doc_read.STAT_KEYS, 0)
        text = '甲\x07乙\r丙\x07\x07'
        blocks = doc_read._split({'ccp_text': 0}, text, stats)
        self.assertEqual(blocks[0]['cells'], ['甲', '乙\n丙'])

    def test_clean_drops_field_code_keeps_result(self):
        stats = dict.fromkeys(doc_read.STAT_KEYS, 0)
        value = doc_read._clean('\x13 PAGEREF x \x1412\x15', stats)
        self.assertEqual(value, '12')
        self.assertEqual(stats['fields'], 1)

    def test_clean_marks_pictures_and_objects(self):
        stats = dict.fromkeys(doc_read.STAT_KEYS, 0)
        value = doc_read._clean('a\x01b\x08c', stats)
        self.assertIn(doc_read.PIC, value)
        self.assertIn(doc_read.OBJECT, value)
        self.assertEqual((stats['pictures'], stats['objects']), (1, 1))


class ExtractTests(unittest.TestCase):
    def test_rejects_docx(self):
        from zipfile import ZipFile
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / '甲.docx'
            with ZipFile(path, 'w') as z:
                z.writestr('word/document.xml', '<w:document/>')
            with self.assertRaises(doc_read.DocError):
                doc_read.extract(path)

    def test_rejects_truncated_doc(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / '乙.doc'
            path.write_bytes(bytes.fromhex('d0cf11e0a1b11ae1') + b'\0' * 500)
            with self.assertRaises(doc_read.DocError):
                doc_read.extract(path)


if __name__ == '__main__':
    unittest.main()
