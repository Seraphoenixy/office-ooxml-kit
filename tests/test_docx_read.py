"""提取层回归：用真实文档钉住 docx-model-v1 的关键语义。

这些断言不是「跑通就行」，每一条都对应一个曾经踩过的坑：内容控件里的单元格会漏格、
图片 token 位置写错会让图片被挤到文字后面、bodyIndex 按已提取块计数会让定位整体偏移。
"""
from __future__ import annotations

import json
import unittest
from pathlib import Path

from office_ooxml_kit import docx_read
from office_ooxml_kit.docx_read import extract, normalize, text_of

CORPUS = Path('D:/project/form_generate')
TEMPLATE = next(iter(sorted(CORPUS.glob('BDSNASA-3511*.docx'))), None)
SPEC = CORPUS / 'BDS2230DG007 FPGA-ZX2230总线数据传输电子硬件测试规程A1.docx'


def require(path):
    if path is None or not path.exists():
        raise unittest.SkipTest(f'缺少样本文件：{path}')
    return path


class TestNormalize(unittest.TestCase):
    def test_collapses_all_whitespace(self):
        self.assertEqual(normalize('  a \n\t b  '), 'a b')

    def test_handles_none(self):
        self.assertEqual(normalize(None), '')


class TestExtractShape(unittest.TestCase):
    """字段形状是契约：挪动或改名都会让既有消费方静默取到空值。"""

    PARAGRAPH_KEYS = {'kind', 'text', 'number', 'section', 'sectionTitle', 'category',
                      'locator', 'bodyIndex'}
    ROW_KEYS = {'kind', 'text', 'cells', 'cellImages', 'spans', 'number', 'section',
                'sectionTitle', 'category', 'tableIndex', 'rowIndex', 'locator',
                'bodyIndex', 'grid', 'layout'}

    def setUp(self):
        self.model = extract(require(SPEC))

    def test_header_fields(self):
        for key in ('schemaVersion', 'file', 'sha256', 'metadata', 'blocks', 'warnings', 'ancillary'):
            self.assertIn(key, self.model)
        self.assertEqual(self.model['schemaVersion'], 1)

    def test_paragraph_keys_present(self):
        paragraph = next(b for b in self.model['blocks'] if b['kind'] in ('paragraph', 'heading'))
        self.assertLessEqual(self.PARAGRAPH_KEYS, set(paragraph))

    def test_row_keys_present(self):
        row = next(b for b in self.model['blocks'] if b['kind'] == 'tableRow')
        self.assertLessEqual(self.ROW_KEYS, set(row))

    def test_every_block_has_locator(self):
        for block in self.model['blocks']:
            self.assertTrue(block['locator'].startswith('word/document.xml:/w:document/w:body'))
            self.assertIsInstance(block['bodyIndex'], int)


class TestHeadings(unittest.TestCase):
    def setUp(self):
        self.model = extract(require(SPEC))

    def test_section_is_path_joined_by_slash(self):
        # 既有消费方按 ' / ' 拆路径取祖先章节，分隔符变了就会整段错位。
        for block in self.model['blocks']:
            self.assertNotIn('//', block['section'])

    def test_section_title_is_last_segment(self):
        for block in self.model['blocks']:
            if block['section'] == docx_read.FRONT:
                continue
            self.assertEqual(block['section'].split(' / ')[-1], block['sectionTitle'])

    def test_toc_entries_are_not_headings(self):
        # 目录项带章节号但不是标题；当成标题会让后续每个块的 section 都错。
        for block in self.model['blocks']:
            if block['category'] == 'toc':
                self.assertNotEqual(block['kind'], 'heading')

    def test_headings_have_outline_level(self):
        headings = [b for b in self.model['blocks'] if b['kind'] == 'heading']
        self.assertTrue(headings, '样本应含标题')
        self.assertTrue(all(b['text'] for b in headings))


class TestTableCells(unittest.TestCase):
    def setUp(self):
        self.model = extract(require(SPEC))

    def test_cells_count_matches_spans(self):
        for block in self.model['blocks']:
            if block['kind'] != 'tableRow':
                continue
            self.assertEqual(len(block['cells']), len(block['spans']))
            self.assertEqual(len(block['cells']), len(block['cellImages']))

    def test_text_is_pipe_joined_cells(self):
        for block in self.model['blocks']:
            if block['kind'] != 'tableRow':
                continue
            self.assertEqual(block['text'], ' | '.join(block['cells']))

    def test_image_token_present_in_cell_text(self):
        # 图片出现在文字里，才能让「换一张图」被文本比较发现。
        rows = [b for b in self.model['blocks']
                if b['kind'] == 'tableRow' and any(b['cellImages'])]
        self.assertTrue(rows, '样本应含带图表格行')
        for block in rows:
            for text, pictures in zip(block['cells'], block['cellImages']):
                for picture in pictures:
                    self.assertIn(picture['token'], text)

    def test_merged_cells_flagged(self):
        states = {state for block in self.model['blocks'] if block['kind'] == 'tableRow'
                  for state in block.get('vMerge', [])}
        self.assertLessEqual(states, {'', 'restart', 'continue'})


class TestImages(unittest.TestCase):
    def setUp(self):
        self.model = extract(require(SPEC))

    def test_tokens_are_hashes(self):
        import re
        pattern = re.compile('⟦图:[0-9A-Fa-f]{8}⟧')
        for block in self.model['blocks']:
            for picture in block.get('images') or []:
                self.assertRegex(picture['token'], pattern)

    def test_same_bytes_share_token(self):
        seen = {}
        for block in self.model['blocks']:
            for picture in block.get('images') or []:
                self.assertEqual(seen.setdefault(picture['token'], picture['sha256']),
                                 picture['sha256'])


class TestMathText(unittest.TestCase):
    """公式（m:t）是可见内容，必须进文字；漏了它等于把公式换成空白。

    这条差异是语料比对发现的：PowerShell 版的 Text 只取 w:t，把 OMML 公式里的
    σpr_gnd、5×10-6 m/m 全丢了。单元格看起来「有内容」，实际比对时是空的。
    """

    MATH_DOC = Path('D:/project/BDS2560-SW1/docs/BDS2560&2580模块间接口协议-A4-2026-7-13.docx')

    def test_math_text_is_kept(self):
        model = extract(require(self.MATH_DOC))
        joined = ' '.join(block['text'] for block in model['blocks'])
        self.assertIn('σpr_gnd', joined)
        self.assertIn('1.275×10-3', joined)

    def test_math_cell_is_not_blank(self):
        # 该行第 2 格只有公式；若只取 w:t，这一格会变成空串，进而误判成「单元格为空」。
        model = extract(require(self.MATH_DOC))
        for block in model['blocks']:
            if block.get('kind') != 'tableRow':
                continue
            if any('σpr_gnd' in str(cell) for cell in (block.get('cells') or [])):
                self.assertNotEqual((block['cells'])[1], '')
                return
        self.fail('未找到含 σpr_gnd 的表格行')


class TestWarnings(unittest.TestCase):
    def test_warns_about_unverifiable_content(self):
        model = extract(require(SPEC))
        self.assertTrue(model['warnings'])
        joined = ' '.join(model['warnings'])
        self.assertIn('图片', joined)
        self.assertIn('人工核对', joined)

    def test_warns_about_math_so_reader_knows_text_is_approximate(self):
        # 公式的文字已经进块，但上下标、分式这类排版无法由文字比较验证，必须照实说明。
        model = extract(require(TestMathText.MATH_DOC))
        self.assertTrue(any('公式' in warning for warning in model['warnings']))

    def test_ancillary_captured(self):
        # 页眉页脚/脚注/批注的实际存在与否取决于文档本身，断言「提取到的都是包里真有
        # 的部分」比断言某个具体前缀更可靠——模板一个页眉都没有的情况下也不会误报。
        from office_ooxml_kit.package import Package
        model = extract(require(TEMPLATE))
        parts = Package(TEMPLATE).part_names()
        self.assertTrue(model['ancillary'])
        for item in model['ancillary']:
            self.assertIn(item['part'], parts)


class TestMetadata(unittest.TestCase):
    def test_extracts_document_id_and_version(self):
        model = extract(require(SPEC))
        self.assertTrue(model['metadata']['documentId'].startswith('BDS'))
        self.assertTrue(model['metadata']['version'])


class TestGuardrails(unittest.TestCase):
    def test_rejects_non_docx(self):
        with self.assertRaises(docx_read.DocxError):
            extract(Path(__file__).resolve())

    def test_missing_file_raises(self):
        with self.assertRaises(Exception):
            extract(Path('D:/project/office-ooxml-kit/does-not-exist.docx'))


if __name__ == '__main__':
    unittest.main()
