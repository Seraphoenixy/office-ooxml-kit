"""export.py 的测试：可发布形态必须是真的普通目录，且版本必须与仓库声明一致。

这些测试用临时 git 仓库，不碰真实项目，也不联网。
"""
import importlib.util
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location('export', ROOT / 'tools' / 'export.py')
export = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(export)


def git(*args, cwd):
    return subprocess.run(['git', *args], cwd=str(cwd), capture_output=True, text=True,
                          encoding='utf-8', errors='replace')


def commit(repo, message='c'):
    git('add', '-A', cwd=repo)
    return git('-c', 'user.name=t', '-c', 'user.email=t@t', 'commit', '-qm', message, cwd=repo)


@unittest.skipUnless(shutil.which('git'), '需要 git')
class ExportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix='export-test-'))
        self.lib = self.tmp / 'lib'
        self.lib.mkdir()
        git('init', '-q', cwd=self.lib)
        (self.lib / 'hello.py').write_text('VALUE = 1\n', encoding='utf-8')
        commit(self.lib)
        self.consumer = self.tmp / 'consumer'
        self.consumer.mkdir()
        git('init', '-q', cwd=self.consumer)
        (self.consumer / 'readme.txt').write_text('consumer\n', encoding='utf-8')
        commit(self.consumer)
        git('-c', 'protocol.file.allow=always', 'submodule', 'add', str(self.lib), 'vendor/lib',
            cwd=self.consumer)
        commit(self.consumer, 'add submodule')
        self.out = self.tmp / 'out'

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_export_inlines_submodule_and_drops_git(self):
        out, exported = export.run(self.consumer, self.out)
        self.assertEqual(out, self.out.resolve())
        self.assertEqual(len(exported), 1)
        self.assertTrue((self.out / 'readme.txt').is_file())
        # 关键：子模块内容以普通文件形式出现，而不是空目录或 gitlink。
        self.assertTrue((self.out / 'vendor/lib/hello.py').is_file())
        self.assertFalse((self.out / 'vendor/lib/.git').exists())
        self.assertFalse((self.out / '.git').exists())

    def test_export_refuses_drifted_submodule_before_writing(self):
        sub = self.consumer / 'vendor/lib'
        (sub / 'hello.py').write_text('VALUE = 2\n', encoding='utf-8')
        commit(sub, 'drift')
        with self.assertRaises(export.ExportError):
            export.run(self.consumer, self.out)
        # 前置校验失败时不得留下半成品，否则会被当成成功产物发布。
        self.assertFalse(self.out.exists())

    def test_export_refuses_non_empty_output_without_force(self):
        self.out.mkdir()
        (self.out / 'stale.txt').write_text('old\n', encoding='utf-8')
        with self.assertRaises(export.ExportError):
            export.run(self.consumer, self.out)
        with self.assertRaises(export.ExportError):
            export.run(self.tmp / 'nope', self.out)

    def test_export_applies_force_to_wipe_stale_output(self):
        self.out.mkdir()
        (self.out / 'stale.txt').write_text('old\n', encoding='utf-8')
        export.run(self.consumer, self.out, force=True)
        self.assertFalse((self.out / 'stale.txt').exists())
        self.assertTrue((self.out / 'vendor/lib/hello.py').is_file())

    def test_export_records_pinned_commit(self):
        _out, exported = export.run(self.consumer, self.out)
        sub_path, sha, url = exported[0]
        self.assertEqual(sub_path, 'vendor/lib')
        self.assertEqual(sha, git('rev-parse', 'HEAD', cwd=self.lib).stdout.strip())
        self.assertEqual(url, str(self.lib))

    def test_export_without_submodules_is_plain_directory(self):
        plain = self.tmp / 'plain'
        plain.mkdir()
        git('init', '-q', cwd=plain)
        (plain / 'a.txt').write_text('a\n', encoding='utf-8')
        commit(plain)
        out, exported = export.run(plain, self.tmp / 'plain-out')
        self.assertEqual(exported, [])
        self.assertTrue((out / 'a.txt').is_file())


if __name__ == '__main__':
    unittest.main()
