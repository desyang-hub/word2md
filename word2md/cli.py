"""``word2md`` 命令行入口。"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from . import __version__
from .converter import convert_docx
from .mathtype import convert_mathtype

__all__ = ["main"]

TMP_PREFIX = "tmp_"


def build_parser() -> argparse.ArgumentParser:
    """构建命令行参数解析器。"""
    parser = argparse.ArgumentParser(
        prog="word2md",
        description=(
            "将 Word 文档转换为 Markdown：先复制为当前目录下的 tmp_ 前缀副本并批量转换 "
            "MathType 公式，再转换为 Markdown、修复 LaTeX 公式，最后删除副本。"
        ),
    )
    parser.add_argument("input", type=Path, help="待转换的 Word 文档路径")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        metavar="FILE",
        help="输出 Markdown 路径（默认与输入文件名同名的 _final.md）",
    )
    parser.add_argument(
        "--skip-mathtype",
        action="store_true",
        help="跳过 MathType 公式转换（用于非 Windows 环境或公式已转换过的文档）",
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
    target: Path = args.output or Path(source.stem + "_final.md")

    if not source.is_file():
        print(f"错误: 输入文件不存在: {source}", file=sys.stderr)
        return 1

    # 复制为当前目录下的 tmp_ 副本再处理，避免改动原始文档
    working = Path(TMP_PREFIX + source.name)
    temp_is_source = working.resolve() == source.resolve()

    try:
        try:
            if not temp_is_source:
                shutil.copy2(source, working)
        except OSError as exc:
            print(f"错误: 复制到 {working} 失败: {exc}", file=sys.stderr)
            return 1

        if not args.quiet:
            print(f"[1/3] 已复制到 {working}")

        if args.skip_mathtype:
            if not args.quiet:
                print("[2/3] 跳过 MathType 转换")
        else:
            if not args.quiet:
                print("[2/3] MathType 公式转换...")
            try:
                ok, fail = convert_mathtype(working, quiet=args.quiet)
                if not args.quiet:
                    print(f"  成功 {ok} 个，失败 {fail} 个")
            except Exception as exc:  # 公式转换失败时仍继续后续转换
                print(f"警告: MathType 转换失败，继续后续步骤: {exc}", file=sys.stderr)

        if not args.quiet:
            print("[3/3] 转换为 Markdown...")
        try:
            convert_docx(working, target)
        except OSError as exc:
            print(f"错误: 转换为 {target} 失败: {exc}", file=sys.stderr)
            return 1

        if not args.quiet:
            print(f"[完成] {target}")
        return 0
    finally:
        # 转换结束（无论成功与否）删除临时的 tmp_ 副本
        if not temp_is_source:
            try:
                working.unlink(missing_ok=True)
            except OSError as exc:
                print(f"警告: 删除临时文件 {working} 失败: {exc}", file=sys.stderr)
