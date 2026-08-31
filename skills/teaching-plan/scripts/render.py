#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render: 教学设计排版生成器统一入口 CLI (Skill 自包含执行引擎)。
从 JSON 内容文件与声明式模板生成排版规范的 Word (.docx) 文档。
"""

import sys
import os
import json
import glob
import argparse
from pathlib import Path

# 1. 依赖优雅自检
try:
    import docx
except ImportError:
    print("❌ [环境提示] 未检测到 Word 文档处理库 'python-docx'。", file=sys.stderr)
    print("👉 请在当前 Python 环境中运行: pip install python-docx", file=sys.stderr)
    sys.exit(2)

# Ensure sibling modules can be imported
SKILL_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.dirname(__file__))

from style_manager import StyleManager
from template_manager import TemplateManager
from doc_builder import TeachingPlanDocBuilder
from validate_content import validate_content
from font_preflight import build_report_for_style

DEFAULT_STYLES_PATH = os.path.join(SKILL_ROOT, "styles", "default_styles.json")
DEFAULT_TEMPLATE_PATH = os.path.join(SKILL_ROOT, "templates", "default_vocational.json")


class ContentValidationError(ValueError):
    """Raised when content JSON is unsafe or incomplete for rendering."""

    def __init__(self, result):
        self.result = result
        details = "; ".join(
            f"{issue['path']}: {issue['message']}"
            for issue in result.get("errors", [])
        )
        super().__init__(f"内容校验未通过（{len(result.get('errors', []))} 个错误）: {details}")


def _validate_before_render(content_data, quiet=False):
    """Validate once at the rendering boundary and surface advisory warnings."""
    result = validate_content(content_data)
    if not result["valid"]:
        raise ContentValidationError(result)
    if result["warnings"] and not quiet:
        for issue in result["warnings"]:
            print(
                f"  ⚠ 内容警告 [{issue['path']}]: {issue['message']}",
                file=sys.stderr,
            )
    return result


def generate_single_document(
    content_data,
    template_path=None,
    styles_path=None,
    style_override=None,
    output_filename=None,
    output_dir=None,
    default_stem="teaching_plan",
    quiet=False
):
    """
    Generate a single Word document from content dictionary.
    """
    _validate_before_render(content_data, quiet=quiet)

    # 1. Initialize TemplateManager
    t_path = template_path or content_data.get("template") or DEFAULT_TEMPLATE_PATH
    template_mgr = TemplateManager(template_config=t_path)

    # 2. Initialize StyleManager
    s_path = styles_path or DEFAULT_STYLES_PATH
    style_mgr = StyleManager(styles_config=s_path, style_override=style_override)
    substitutions = style_mgr.get_style("fonts", "substitutions", default={}) or {}
    effective_fonts = dict(style_mgr.get_style("fonts", default={}) or {})
    for role, requested in list(effective_fonts.items()):
        if (role.endswith("_font") or role == "font") and isinstance(requested, str):
            effective_fonts[role] = substitutions.get(requested, requested)
    font_report = build_report_for_style({"fonts": effective_fonts}, str(s_path))
    unresolved = [check for check in font_report["font_checks"] if not check["available"]]
    if substitutions and not quiet:
        rendered = ", ".join(f"{source} → {target}" for source, target in substitutions.items())
        print(f"  ✓ 已显式应用字体替代: {rendered}", file=sys.stderr)
    if unresolved and not quiet:
        missing = ", ".join(f"{item['role']}={item['requested']}" for item in unresolved)
        print(
            f"  ⚠ 当前生成环境未发现字体: {missing}。"
            "若目标 Word/WPS 已安装这些字体，应保留原声明并在目标编辑器验收；"
            "仅在目标端也缺失或要求跨平台一致时使用 --override。",
            file=sys.stderr,
        )

    # 3. Build Document
    builder = TeachingPlanDocBuilder(content_data, template=template_mgr, style_manager=style_mgr)
    doc = builder.build_document()

    # 4. Resolve Output Filename
    out_name = (
        output_filename or
        content_data.get("output_filename") or
        content_data.get("document_meta", {}).get("output_filename") or
        content_data.get("document_meta", {}).get("default_output_filename") or
        f"{default_stem}.docx"
    )
    if not out_name.endswith(".docx"):
        out_name += ".docx"

    # 5. Resolve Output Directory
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        final_path = os.path.join(output_dir, os.path.basename(out_name))
    else:
        final_path = out_name

    doc.save(final_path)
    if not quiet:
        print(f"  ✓ 已成功生成教案文档: {final_path}")
    return final_path, doc


def generate_document_from_file(
    file_path,
    template_path=None,
    styles_path=None,
    style_override=None,
    output_filename=None,
    output_dir=None,
    quiet=False
):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"未找到内容文件: {file_path}")

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            content_data = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON 文件解析失败 ({file_path}): {e}")

    stem = Path(file_path).stem
    return generate_single_document(
        content_data=content_data,
        template_path=template_path,
        styles_path=styles_path,
        style_override=style_override,
        output_filename=output_filename,
        output_dir=output_dir,
        default_stem=stem,
        quiet=quiet
    )


def parse_override_arg(override_str):
    if not override_str:
        return None
    if os.path.exists(override_str):
        with open(override_str, "r", encoding="utf-8") as f:
            return json.load(f)
    try:
        return json.loads(override_str)
    except json.JSONDecodeError:
        raise argparse.ArgumentTypeError(
            f"无效的 override 参数: '{override_str}' 既不是现有的 JSON 文件，也不是合法的 JSON 字符串"
        )


def main():
    parser = argparse.ArgumentParser(
        prog="render.py",
        description="教案排版渲染器：从 JSON 纯文本内容与声明式模板编译生成符合排版规范的 Word (.docx) 教学设计方案。",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "target",
        nargs="?",
        default=None,
        help="待生成教案的 JSON 内容文件路径"
    )
    parser.add_argument(
        "-c", "--content",
        dest="content",
        default=None,
        help="显式指定内容 JSON 文件路径"
    )
    parser.add_argument(
        "-t", "--template",
        dest="template",
        default=None,
        help="声明式文档模板路径或名称 (默认: default_vocational)"
    )
    parser.add_argument(
        "-s", "--styles",
        dest="styles",
        default=None,
        help="排版样式配置文件路径 (默认: default_styles.json)"
    )
    parser.add_argument(
        "-o", "--output",
        dest="output",
        default=None,
        help="指定输出文件名 (.docx)"
    )
    parser.add_argument(
        "-d", "--output-dir",
        dest="output_dir",
        default=None,
        help="指定输出目录"
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
        help="静默模式"
    )

    args = parser.parse_args()
    content_target = args.content or args.target

    if not content_target:
        parser.print_help()
        sys.exit(1)

    try:
        generate_document_from_file(
            file_path=content_target,
            template_path=args.template,
            styles_path=args.styles,
            style_override=args.override,
            output_filename=args.output,
            output_dir=args.output_dir,
            quiet=args.quiet
        )
    except Exception as e:
        print(f"❌ [错误] 渲染失败: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
