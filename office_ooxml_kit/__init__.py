"""通用 Office 文档读写组件。

读写都需要 lxml。读取理论上可用标准库解析，但那样会有两套解析代码，
字段口径迟早分叉——这里宁可统一依赖，也不接受两份实现。

必须先装 lxml：`python -m pip install "lxml>=6"`。缺少它时这里会给出可执行的提示，
而不是抛一个看不出原因的 ModuleNotFoundError——本组件常以子模块或源码副本的形式
内嵌进下游项目，使用者未必知道它依赖什么。
"""
from __future__ import annotations

__version__ = '1.1.0'

_MISSING_LXML = ('office_ooxml_kit 需要 lxml，当前环境没有装。请执行：'
                 'python -m pip install "lxml>=6"')

try:
    from lxml import etree as _lxml_etree  # noqa: F401
except ImportError as _error:  # pragma: no cover - 依赖缺失时的分支
    raise ImportError(_MISSING_LXML) from _error

from . import doc_read  # noqa: E402  —— .doc（Word 97-2003 二进制）直读
from .doc_read import DocError  # noqa: E402
from .docx_read import DocxError, extract, normalize  # noqa: E402
from .package import Package, PackageError  # noqa: E402

__all__ = ['Package', 'PackageError', 'DocxError', 'DocError', 'doc_read',
           'extract', 'normalize', '__version__']
