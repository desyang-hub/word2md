"""``word2md`` 命令行入口。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .converter import convert_docx

__all__ = ["main"]


def build_parser() -> argparse.ArgumentParser:
    """构建命令行参数解析器。"""
    parser = argparse.ArgumentParser(
        prog="word2md",
        description="将 Word 文档转换为 Markdown，并修复 MathType 转换后的 LaTeX 公式。",
    )
    parser.add_argument("input", type=Path, help="待转换的 Word 文档路径")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        metavar="FILE",
        help="输出 Markdown 路径（默认与输入文件同目录、同名的 .md 文件）",
    )
    parser.add_argument(
        "-q", "--quiet", action="store_true", help="静默模式，不输出进度信息"
    )
    parser.add_argument(
        "-V", "--version", action="version", version=f"%(prog)s {__version__}"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """执行命令行程序，返回进程退出码。"""
    args = build_parser().parse_args(argv)
    source: Path = args.input
    target: Path = args.output or source.with_suffix(".md")

    if not args.quiet:
        print(f"[1/2] 正在转换 {source} ...")

    try:
        convert_docx(source, target)
    except FileNotFoundError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 1
    except OSError as exc:
        print(f"错误: 写入 {target} 失败: {exc}", file=sys.stderr)
        return 1

    if not args.quiet:
        print(f"[2/2] 已生成 {target}")
    return 0
