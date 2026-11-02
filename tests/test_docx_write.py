"""写入层回归：填充必须产出 Word 能打开、且与参照实现一致的文件。

这里的断言对应三个曾让文件打不开的真实缺陷：
- 单元格末尾没有段落 → Word 报「文件已经损坏」；
- 嵌套表内层段落没有 pPr → 行距塌陷、页数差一页；
- 列宽取整用截断 → 每列少 1 twip，整表宽度对不上原文。
"""
from __future__ import annotations

import json
import shutil
import tempfile
import unittest
from pathlib import Path

from lxml import etree

from office_ooxml_kit.docx_write import FillError, fill
from office_ooxml_kit.ns import local, q, tostring
from office_ooxml_kit.package import Package
from office_ooxml_kit.verify import check, inspect

RUNS = Path('D:/project/form_generate/runs')
SAMPLE_RUN = RUNS / 'BDS2230DG007_FPGA-ZX2230总线数据传输电子硬件测试规程_A1-to-A2'


def require_run():
    if not SAMPLE_RUN.exists():
        raise unittest.SkipTest(f'缺少样本运行目录：{SAMPLE_RUN}')
    return SAMPLE_RUN


def require_sources(manifest):
    """样本目录里的 manifest 指向 work 目录之外的源文档，它们不属于本仓库。

    源文档可能被移走或清理；这时跳过相关用例而不是报错，否则环境缺少输入
    会表现为代码失败，掩盖真正的回归。
    """
    missing = [name for name in (manifest.get('old'), manifest.get('new'), manifest.get('template'))
               if name and not Path(name).is_file()]
    if missing:
        raise unittest.SkipTest('缺少样本源文档（不在本仓库内）：' + '、'.join(Path(x).name for x in missing))


class FillCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        run = require_run()
        manifest = json.loads((run / 'manifest.json').read_text(encoding='utf-8'))
        require_sources(manifest)
        cls.template = manifest['template']
        cls.data = json.loads((run / 'fill.json').read_text(encoding='utf-8'))
        cls.tmp = Path(tempfile.mkdtemp(prefix='office_ooxml_kit-'))
        cls.output = cls.tmp / 'filled.docx'
        cls.result = fill(cls.template, cls.output, cls.data)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def document(self):
        return etree.fromstring(Package(self.output).read('word/document.xml'))


class TestFillResult(FillCase):
    def test_reports_counts(self):
        self.assertEqual(self.result['rows'], len(self.data['changes']))
        self.assertEqual(self.result['slots'], len(self.data['slots']))
        self.assertGreater(self.result['images'], 0)

    def test_every_cell_ends_with_paragraph(self):
        # Word 硬性要求；缺了直接报文件损坏。
        for cell in self.document().iter(q('w:tc')):
            children = [child for child in cell if local(child) != 'tcPr']
            self.assertTrue(children, '空单元格')
            self.assertEqual(local(children[-1]), 'p')

    def test_contains_exactly_one_outer_table(self):
        body = self.document().find(q('w:body'))
        self.assertEqual(len(body.findall(q('w:tbl'))), 1)

    def test_nested_tables_are_real_tables(self):
        # 管道符文本塞进单元格会让原文表格不可读。
        root = self.document()
        nested = [table for table in root.iter(q('w:tbl')) if local(table.getparent()) == 'tc']
        self.assertGreater(len(nested), 0)
        for table in nested:
            for cell in table.iter(q('w:tc')):
                text = ''.join(node.text or '' for node in cell.iter(q('w:t')))
                self.assertNotIn(' | ', text)

    def test_no_image_tokens_in_word(self):
        # token 是提取层的记号，绝不作为文字留在 Word 里。
        raw = Package(self.output).read('word/document.xml').decode('utf-8')
        self.assertNotIn('⟦图:', raw)

    def test_namespaces_preserved(self):
        raw = Package(self.output).read('word/document.xml').decode('utf-8')
        self.assertIn('mc:Ignorable', raw)
        self.assertNotIn('xmlns:ns0', raw)

    def test_images_have_relationships(self):
        package = Package(self.output)
        rels = etree.fromstring(package.read('word/_rels/document.xml.rels'))
        ids = {node.get('Id') for node in rels.iter() if local(node) == 'Relationship'}
        raw = package.read('word/document.xml').decode('utf-8')
        import re
        used = set(re.findall(r'r:embed="([^"]+)"', raw))
        self.assertEqual(used, used & ids)

    def test_all_media_parts_declared(self):
        package = Package(self.output)
        content_types = etree.fromstring(package.read('[Content_Types].xml'))
        defaults = {node.get('Extension') for node in content_types.iter() if local(node) == 'Default'}
        for name in package.part_names():
            if name.startswith('word/media/'):
                self.assertIn(name.rsplit('.', 1)[-1].lower(), defaults)

    def test_update_fields_disabled(self):
        raw = Package(self.output).optional('word/settings.xml')
        self.assertIsNotNone(raw)
        root = etree.fromstring(raw)
        node = root.find('.//' + q('w:updateFields'))
        self.assertIsNotNone(node)
        self.assertEqual(node.get(q('w:val')), 'false')

    def test_structure_check_passes(self):
        report = check(self.output, baseline=inspect(self.template))
        self.assertTrue(report['passed'], report)


class TestRefusals(FillCase):
    def test_refuses_overwrite(self):
        with self.assertRaises(FillError):
            fill(self.template, self.output, self.data)

    def test_allows_overwrite_when_asked(self):
        target = self.tmp / 'again.docx'
        fill(self.template, target, self.data)
        fill(self.template, target, self.data, overwrite=True)
        self.assertTrue(target.exists())

    def test_rejects_wrong_template(self):
        wrong = Path(self.template)
        with self.assertRaises(FillError):
            fill(wrong, self.tmp / 'wrong.docx', {'slots': [{'row': 99, 'cell': 1, 'text': 'x'}],
                                                  'changes': []})


class TestSlots(FillCase):
    def test_slot_text_landed(self):
        root = self.document()
        table = root.find(f'{q("w:body")}/{q("w:tbl")}')
        rows = table.findall(q('w:tr'))
        for slot in self.data['slots'][:5]:
            if not slot.get('text'):
                continue
            cell = rows[int(slot['row'])].findall(q('w:tc'))[int(slot['cell'])]
            text = ''.join(node.text or '' for node in cell.iter(q('w:t')))
            self.assertEqual(text.strip(), str(slot['text']).split('\n')[0].strip())

    def test_empty_visible_cells_become_dashes(self):
        root = self.document()
        table = root.find(f'{q("w:body")}/{q("w:tbl")}')
        rows = table.findall(q('w:tr'))
        # 第 1 行的更改类别内圈格是空占位；空白会让人以为漏填。
        cell = rows[1].findall(q('w:tc'))[4]
        text = ''.join(node.text or '' for node in cell.iter(q('w:t'))).strip()
        self.assertEqual(text, '--')

    def test_uncertain_values_are_red(self):
        root = self.document()
        reds = 0
        for paragraph in root.iter(q('w:p')):
            text = ''.join(node.text or '' for node in paragraph.iter(q('w:t')))
            if not any(word in text for word in ('待确认', '待审核', '待核实', '待补充')):
                continue
            for run in paragraph.findall(q('w:r')):
                color = run.find(f'{q("w:rPr")}/{q("w:color")}')
                if color is not None and color.get(q('w:val')) == 'FF0000':
                    reds += 1
        self.assertGreater(reds, 0, '待确认内容应标红')


class TestRoundTrip(FillCase):
    def test_extract_after_fill_sees_all_rows(self):
        from office_ooxml_kit.docx_read import extract
        model = extract(self.output)
        rows = [b for b in model['blocks'] if b['kind'] == 'tableRow']
        self.assertGreater(len(rows), len(self.data['changes']))
        self.assertEqual(model['sha256'], Package(self.output).sha256())


if __name__ == '__main__':
    unittest.main()
