#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main: CLI tool and entrypoint for generating Word document teaching plans.

Features:
- Multi-subject batch generation from JSON content files or single file rendering.
- Declarative layout template support (-t, --template).
- Smart automatic naming derived from JSON filename (<stem>.docx).
- Configurable output directory (-d, --output-dir) and custom output file (-o, --output).
- Hierarchical format override support (--override).
- Pre-render content validation & font preflight reports.
- Standard CLI compliance (POSIX flags, exit codes, --version, -q/--quiet, -v/--verbose).
"""

import sys
import os
import glob
import argparse
from pathlib import Path

# Connect to skill implementation engine
SKILL_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "skills", "teaching-plan"))
SCRIPTS_DIR = os.path.join(SKILL_ROOT, "scripts")
EXAMPLES_DIR = os.path.join(SKILL_ROOT, "examples")
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from render import (
    generate_single_document,
    generate_document_from_file,
    parse_override_arg,
    DEFAULT_STYLES_PATH,
    DEFAULT_TEMPLATE_PATH
)

__version__ = "2.0.0"
__app_name__ = "teaching-plan-cli"

DEFAULT_LOCAL_CONTENT_DIR = os.path.join(os.path.dirname(__file__), "content")


def generate_document(
    content_path=None,
    template_path=None,
    styles_path=None,
    style_override=None,
    output_filename=None,
    output_dir=None,
    quiet=False
):
    """
    Main generator interface supporting directories, single JSON files, or content dicts.
    """
    target = content_path

    # If no target specified, check local content/ directory
    if not target:
        if os.path.isdir(DEFAULT_LOCAL_CONTENT_DIR) and glob.glob(os.path.join(DEFAULT_LOCAL_CONTENT_DIR, "*.json")):
            target = DEFAULT_LOCAL_CONTENT_DIR
        else:
            if not quiet:
                print("💡 未指定输入内容。请指定待生成的 JSON 文件或目录路径。", file=sys.stderr)
                print(f"👉 示例用法:", file=sys.stderr)
                print(f"   python3 main.py \"skills/teaching-plan/examples/Python程序设计教学设计（示例）.json\"", file=sys.stderr)
                print(f"   python3 main.py --help 查看完整使用帮助", file=sys.stderr)
            return []

    # 1. Content dict passed directly
    if isinstance(target, dict):
        return generate_single_document(
            content_data=target,
            template_path=template_path,
            styles_path=styles_path,
            style_override=style_override,
            output_filename=output_filename,
            output_dir=output_dir,
            default_stem="teaching_plan",
            quiet=quiet
        )

    # 2. Single JSON file passed
    if os.path.isfile(target):
        return generate_document_from_file(
            file_path=target,
            template_path=template_path,
            styles_path=styles_path,
            style_override=style_override,
            output_filename=output_filename,
            output_dir=output_dir,
            quiet=quiet
        )

    # 3. Directory passed: Batch generate all JSON files
    if os.path.isdir(target):
        json_files = sorted(glob.glob(os.path.join(target, "*.json")))
        if not json_files:
            if not quiet:
                print(f"⚠️  警告: 在目录 '{target}' 中未找到任何 .json 文件", file=sys.stderr)
            return []

        if not quiet:
            print(f"📁 扫描到 {len(json_files)} 个科目内容文件 ('{target}')，开始批量生成...")

        results = []
        for i, jf in enumerate(json_files, start=1):
            if not quiet:
                print(f"[{i}/{len(json_files)}] 正在处理: {os.path.basename(jf)}")
            res = generate_document_from_file(
                file_path=jf,
                template_path=template_path,
                styles_path=styles_path,
                style_override=style_override,
                output_filename=output_filename if len(json_files) == 1 else None,
                output_dir=output_dir,
                quiet=quiet
            )
            results.append(res)

        if not quiet:
            dest_info = f" -> {output_dir}" if output_dir else ""
            print(f"✨ 批量生成完成，共成功生成 {len(results)} 份教案文档{dest_info}！")
        return results

    raise FileNotFoundError(f"指定的内容路径不存在: {target}")


def main():
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="教案生成器 CLI：从 JSON 纯文本内容与声明式模板批量/单文件生成排版规范的 Word (.docx) 教学设计方案。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例:
  %(prog)s lesson.json                     # 生成单个科目教案
  %(prog)s skills/teaching-plan/examples/  # 批量生成示例目录下的所有教案
  %(prog)s -t default_vocational           # 指定模板名称
  %(prog)s -d dist/ lesson.json            # 生成并输出到 dist/ 目录
  %(prog)s lesson.json -o custom.docx      # 生成单个文件并重命名
  %(prog)s -q lesson.json                  # 静默模式，只在出错时输出
  %(prog)s --override '{"fonts":{"default_font":"黑体"}}' lesson.json  # 传入行内样式覆盖
"""
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}"
    )
    parser.add_argument(
        "target",
        nargs="?",
        default=None,
        help="内容目录或单文件路径（若未提供，优先尝试本地 content/ 目录）"
    )
    parser.add_argument(
        "-c", "--content",
        dest="content",
        default=None,
        help="显式指定内容目录或文件路径"
    )
    parser.add_argument(
        "-t", "--template",
        dest="template",
        default=None,
        help="声明式文档模板路径或名称 (默认: default_vocational)"
    )
    parser.add_argument(
        "-o", "--output",
        dest="output",
        default=None,
        help="指定输出文件名（仅在生成单个文件时有效）"
    )
    parser.add_argument(
        "-d", "--output-dir",
        dest="output_dir",
        default=None,
        help="指定输出目录（默认保存在当前工作目录）"
    )
    parser.add_argument(
        "-s", "--styles",
        dest="styles",
        default=None,
        help="排版样式配置文件路径 (默认: default_styles.json)"
    )
    parser.add_argument(
        "--override",
        dest="override",
        type=parse_override_arg,
        default=None,
        help="样式覆盖配置（支持 JSON 文件路径或行内 JSON 字符串）"
    )
    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="静默模式，仅输出错误信息"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="显示详细调试信息与错误堆栈"
    )

    args = parser.parse_args()
    content_target = args.content or args.target

    try:
        generate_document(
            content_path=content_target,
            template_path=args.template,
            styles_path=args.styles,
            style_override=args.override,
            output_filename=args.output,
            output_dir=args.output_dir,
            quiet=args.quiet
        )
    except Exception as e:
        if args.verbose:
            import traceback
            traceback.print_exc()
        else:
            print(f"❌ [错误] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
