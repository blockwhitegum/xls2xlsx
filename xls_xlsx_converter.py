#!/usr/bin/env python3
import argparse
import os
from pathlib import Path
import sys


def _require_dependencies(target_ext: str):
    try:
        import pyexcel  # noqa: F401
    except Exception:
        raise RuntimeError(
            "未找到依赖: pyexcel\n请安装: pip install pyexcel pyexcel-xls pyexcel-xlsx"
        )
    # 根据目标格式检查对应插件
    try:
        if target_ext == ".xls":
            import pyexcel_xls  # noqa: F401
        elif target_ext == ".xlsx":
            import pyexcel_xlsx  # noqa: F401
    except Exception:
        need = "pyexcel-xls" if target_ext == ".xls" else "pyexcel-xlsx"
        raise RuntimeError(f"未找到依赖: {need}\n请安装: pip install {need}")


def convert_file(input_path: Path, output_path: Path):
    target_ext = output_path.suffix.lower()
    _require_dependencies(target_ext)
    from pyexcel import get_book

    book = get_book(file_name=str(input_path))
    book.save_as(str(output_path))


def derive_output_path(input_path: Path, target_format: str | None, explicit_output: Path | None) -> Path:
    if explicit_output:
        return explicit_output
    src_ext = input_path.suffix.lower()
    if target_format:
        target_ext = "." + target_format.lower()
    else:
        # 自动推断目标：互转
        target_ext = ".xls" if src_ext == ".xlsx" else ".xlsx"
    return input_path.with_suffix(target_ext)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Excel 转换器：在 xlsx 与 xls 之间互相转换（基于 pyexcel）"
    )
    parser.add_argument("input", help="输入文件路径（.xls 或 .xlsx）")
    parser.add_argument("-o", "--output", help="输出文件路径（不填则与输入同名更换扩展名）")
    parser.add_argument(
        "--to",
        choices=["xls", "xlsx"],
        help="明确指定目标格式（不填则自动与输入相反）",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="若输出文件已存在则覆盖",
    )

    args = parser.parse_args(argv)
    input_path = Path(args.input)

    if not input_path.exists() or not input_path.is_file():
        print(f"错误：输入文件不存在或不是文件: {input_path}", file=sys.stderr)
        return 2

    src_ext = input_path.suffix.lower()
    if src_ext not in {".xls", ".xlsx"}:
        print("错误：仅支持 .xls 或 .xlsx 输入文件", file=sys.stderr)
        return 2

    output_path = derive_output_path(
        input_path,
        args.to,
        Path(args.output) if args.output else None,
    )

    if output_path.exists() and not args.overwrite:
        print(f"错误：输出文件已存在: {output_path}（使用 --overwrite 覆盖）", file=sys.stderr)
        return 2

    try:
        convert_file(input_path, output_path)
    except Exception as e:
        print(f"转换失败：{e}", file=sys.stderr)
        return 1

    print(f"转换成功：{input_path} -> {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
