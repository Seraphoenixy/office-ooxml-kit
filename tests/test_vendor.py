"""vendor.py 的测试：内嵌副本是可发布形态，指纹校验不能失效。

这些测试针对发布机制本身，而不是 office_ooxml_kit 的读写能力：如果 --check 不能发现
副本被手改或落后于源，下游项目就会带着一份没人知道已经漂移的代码发给同事。
"""
import importlib.util
import json
import shutil
import tempfile
import unittest
from pathlib import Path

# tools/ 不是包（没有 __init__.py），按文件加载，和本仓库其他测试的做法一致。
ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location('vendor', ROOT / 'tools' / 'vendor.py')
vendor = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(vendor)


class VendorTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix='vendor-'))
        self.target = self.tmp / 'office_ooxml_kit'

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_copy_writes_marker_and_package_files(self):
        self.assertEqual(vendor.copy(self.target), 0)
        marker = json.loads((self.target / 'VENDOR.json').read_text(encoding='utf-8'))
        self.assertEqual(marker['package'], 'office_ooxml_kit')
        self.assertEqual(marker['version'], vendor.manifest()['version'])
        self.assertEqual(sorted(marker['files']), sorted(vendor.manifest()['files']))
        self.assertTrue((self.target / '__init__.py').is_file())
        # 测试与缓存不随副本发布。
        self.assertFalse(list(self.target.glob('test_*.py')))

    def test_check_accepts_fresh_copy(self):
        vendor.copy(self.target)
        self.assertEqual(vendor.check(self.target), 0)

    def test_check_detects_hand_edit(self):
        vendor.copy(self.target)
        victim = self.target / 'ns.py'
        victim.write_text(victim.read_text(encoding='utf-8') + '\n# 本地改动\n', encoding='utf-8')
        self.assertEqual(vendor.check(self.target), 1)

    def test_check_detects_source_newer_than_copy(self):
        vendor.copy(self.target)
        marker_path = self.target / 'VENDOR.json'
        marker = json.loads(marker_path.read_text(encoding='utf-8'))
        # 删掉一个已记录的条目，模拟源新增了文件而副本没有同步。
        marker['files'].pop('images.py')
        marker_path.write_text(json.dumps(marker, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        self.assertEqual(vendor.check(self.target), 1)

    def test_check_rejects_directory_without_marker(self):
        self.target.mkdir(parents=True)
        self.assertEqual(vendor.check(self.target), 1)

    def test_marker_records_source_path_so_copy_stays_traceable(self):
        vendor.copy(self.target)
        marker = json.loads((self.target / 'VENDOR.json').read_text(encoding='utf-8'))
        self.assertEqual(Path(marker['source']), vendor.SOURCE)
        self.assertIn('syncedAt', marker)


if __name__ == '__main__':
    unittest.main()
