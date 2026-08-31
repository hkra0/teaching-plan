#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render a DOCX to per-page PNG files for fallback visual inspection.

This script prepares QA evidence; it cannot decide whether the pages look good.
It uses LibreOffice and Poppler and is not a substitute for Word/WPS QA.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def _find_executable(name: str, extra_candidates: tuple[str, ...] = ()) -> str | None:
    found = shutil.which(name)
    if found:
        return found
    for candidate in extra_candidates:
        if Path(candidate).is_file():
            return candidate
    return None


def _natural_page_key(path: Path) -> int:
    try:
        return int(path.stem.rsplit("-", 1)[1])
    except (IndexError, ValueError):
        return 0


def render_pages(docx_path: Path, output_dir: Path) -> dict:
    docx_path = docx_path.expanduser().resolve()
    output_dir = output_dir.expanduser().resolve()
    if not docx_path.is_file():
        raise ValueError(f"找不到 DOCX：{docx_path}")
    if docx_path.suffix.casefold() != ".docx":
        raise ValueError("输入文件必须是 .docx")

    soffice = _find_executable(
        "soffice",
        ("/Applications/LibreOffice.app/Contents/MacOS/soffice",),
    )
    pdftoppm = _find_executable("pdftoppm")
    missing = [name for name, value in (("LibreOffice/soffice", soffice), ("Poppler/pdftoppm", pdftoppm)) if not value]
    if missing:
        raise RuntimeError(f"缺少视觉验收依赖：{', '.join(missing)}")

    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = output_dir / f"{docx_path.stem}.pdf"
    existing_pages = sorted(output_dir.glob("page-*.png"), key=_natural_page_key)
    if pdf_path.exists() or existing_pages:
        raise ValueError(f"QA 目录中已有同名 PDF 或 page-*.png，请使用新的空目录：{output_dir}")

    with tempfile.TemporaryDirectory(prefix="teaching-plan-qa-") as temp_name:
        temp_root = Path(temp_name)
        profile = temp_root / "libreoffice-profile"
        cache = temp_root / "cache"
        profile.mkdir()
        cache.mkdir()
        env = os.environ.copy()
        env.update(
            {
                "TMPDIR": str(temp_root),
                "XDG_CACHE_HOME": str(cache),
            }
        )
        convert = subprocess.run(
            [
                soffice,
                "--headless",
                f"-env:UserInstallation={profile.resolve().as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output_dir),
                str(docx_path),
            ],
            check=False,
            capture_output=True,
            text=True,
            env=env,
            timeout=180,
        )
        if convert.returncode != 0 or not pdf_path.is_file() or pdf_path.stat().st_size == 0:
            detail = (convert.stderr or convert.stdout).strip()
            raise RuntimeError(f"LibreOffice 转换失败（退出码 {convert.returncode}）：{detail}")

        raster = subprocess.run(
            [pdftoppm, "-png", "-r", "144", str(pdf_path), str(output_dir / "page")],
            check=False,
            capture_output=True,
            text=True,
            timeout=180,
        )
        if raster.returncode != 0:
            raise RuntimeError(f"PDF 页面栅格化失败：{raster.stderr.strip()}")

    pages = sorted(output_dir.glob("page-*.png"), key=_natural_page_key)
    invalid = [str(page) for page in pages if page.stat().st_size == 0]
    if not pages or invalid:
        raise RuntimeError("未生成有效的逐页 PNG")
    return {
        "docx": str(docx_path),
        "pdf": str(pdf_path),
        "output_dir": str(output_dir),
        "page_count": len(pages),
        "pages": [str(page) for page in pages],
        "visual_inspection_complete": False,
        "renderer": "LibreOffice",
        "verification_scope": "fallback_regression_only",
        "target_editor_verified": False,
        "warning": "LibreOffice 的字体与分页结果不能代表 Microsoft Word 或 WPS。",
        "required_review": ["缺字", "裁切或重叠", "断表", "异常空白页", "孤立反思页", "页数异常"],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="使用 LibreOffice 将 DOCX 渲染为逐页 PNG，仅供降级回归检查。")
    parser.add_argument("docx", type=Path, help="待检查的 DOCX")
    parser.add_argument("--output-dir", type=Path, required=True, help="新的 QA 输出目录")
    parser.add_argument("--json", action="store_true", help="输出结构化结果")
    args = parser.parse_args(argv)
    try:
        report = render_pages(args.docx, args.output_dir)
    except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
        print(f"视觉验收准备失败：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"已生成 {report['page_count']} 页 LibreOffice 降级预览：{report['output_dir']}")
        print("状态：未经过 Word/WPS 目标编辑器验收，不能据此标记高保真交付完成。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
