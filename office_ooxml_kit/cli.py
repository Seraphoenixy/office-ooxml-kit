"""命令行入口：`python -m office_ooxml_kit <子命令>`。

Node 项目按既有习惯通过子进程调用（form_generate 就是这样调本组件的）
所以这里只负责参数解析与 JSON 落盘，不做业务判断。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import docx_read, docx_write, verify
from .package import PackageError

__all__ = ['main']


def _write_json(payload, output):
    # 既有流程用 UTF-8 无 BOM；带 BOM 的 JSON 会让 Node 的 readFile+JSON.parse 正常，
    # 但会让 sha256 比对在重新落盘后无故变化。
    text = json.dumps(payload, ensure_ascii=False, indent=1)
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding='utf-8', newline='\n')
    return path


def _load(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))


def main(argv=None):
    parser = argparse.ArgumentParser(prog='office_ooxml_kit', description='Office 文档结构读写组件')
    sub = parser.add_subparsers(dest='command', required=True)

    p_extract = sub.add_parser('extract', help='提取 DOCX 块模型（docx-model-v1）')
    p_extract.add_argument('--path', required=True, help='输入 DOCX')
    p_extract.add_argument('--output', help='输出 JSON；缺省打印到标准输出')
    p_extract.add_argument('--no-text', action='store_true', help='跳过不透明块')

    p_inspect = sub.add_parser('inspect', help='输出包清单、分区几何与外层表格')
    p_inspect.add_argument('--path', required=True)
    p_inspect.add_argument('--output', required=True)

    p_fill = sub.add_parser('fill', help='按 fill.json 填充模板')
    p_fill.add_argument('--path', required=True, help='模板 DOCX')
    p_fill.add_argument('--output', required=True, help='输出 DOCX（已存在则拒写）')
    p_fill.add_argument('--data', required=True, help='fill.json')
    p_fill.add_argument('--overwrite', action='store_true', help='允许覆盖输出')

    p_verify = sub.add_parser('verify', help='结构校验')
    p_verify.add_argument('--path', required=True)
    p_verify.add_argument('--baseline', help='生成前的 inspect.json')
    p_verify.add_argument('--output', help='输出 JSON；缺省打印到标准输出')

    args = parser.parse_args(argv)
    try:
        if args.command == 'extract':
            payload = docx_read.extract(args.path, with_text=not args.no_text)
            if args.output:
                _write_json(payload, args.output)
                print(f'已提取 {len(payload["blocks"])} 个块 → {args.output}')
            else:
                print(json.dumps(payload, ensure_ascii=False, indent=1))
        elif args.command == 'inspect':
            _write_json(verify.inspect(args.path), args.output)
            print(f'已检查 {args.path} → {args.output}')
        elif args.command == 'fill':
            result = docx_write.fill(args.path, args.output, _load(args.data),
                                     overwrite=args.overwrite)
            print(f'已填充 {result["rows"]} 行差异、{result["images"]} 张图片 → {result["output"]}')
        elif args.command == 'verify':
            baseline = _load(args.baseline) if args.baseline else None
            payload = verify.check(args.path, baseline=baseline)
            if args.output:
                _write_json(payload, args.output)
                print(('通过' if payload['passed'] else '未通过') + f' → {args.output}')
            else:
                print(json.dumps(payload, ensure_ascii=False, indent=1))
    except (docx_read.DocxError, docx_write.FillError, verify.VerifyError,
            PackageError) as error:
        print(str(error), file=sys.stderr)
        return 2
    return 0
