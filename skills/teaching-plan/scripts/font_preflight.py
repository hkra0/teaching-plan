#!/usr/bin/env python3
"""Check whether fonts requested by a teaching-plan style are installed.

This utility has no third-party dependencies. It reports fallback suggestions but
never edits the style file or claims that a fallback has been applied.
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import struct
import subprocess
import sys
from pathlib import Path
from typing import Iterable


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_STYLE = SCRIPT_DIR.parent / "styles" / "default_styles.json"

CHINESE_FALLBACKS = {
    "Darwin": {
        "serif": ["Songti SC", "STSong", "PingFang SC", "Arial Unicode MS"],
        "sans": ["PingFang SC", "Heiti SC", "Songti SC", "Arial Unicode MS"],
        "kai": ["Kaiti SC", "STKaiti", "Songti SC", "PingFang SC"],
    },
    "Windows": {
        "serif": ["SimSun", "宋体", "Microsoft YaHei", "微软雅黑"],
        "sans": ["Microsoft YaHei", "微软雅黑", "SimHei", "黑体", "DengXian"],
        "kai": ["KaiTi", "楷体", "KaiTi_GB2312", "楷体_GB2312", "Microsoft YaHei"],
    },
    "Linux": {
        "serif": ["Noto Serif CJK SC", "Source Han Serif SC", "AR PL UMing CN"],
        "sans": ["Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei"],
        "kai": ["LXGW WenKai", "Kaiti SC", "Noto Serif CJK SC"],
    },
}

ASCII_FALLBACKS = {
    "Darwin": ["Times New Roman", "Arial", "Helvetica"],
    "Windows": ["Times New Roman", "Arial", "Calibri"],
    "Linux": ["Liberation Serif", "DejaVu Serif", "Liberation Sans"],
}


def normalize(name: str) -> str:
    return "".join(char.casefold() for char in name if char.isalnum())


def decode_name(raw: bytes, platform_id: int, encoding_id: int) -> str | None:
    encodings = []
    if platform_id in (0, 3):
        encodings.append("utf-16-be")
    elif platform_id == 1:
        encodings.extend(("mac_roman", "utf-8"))
    else:
        encodings.extend(("utf-8", "latin-1"))
    if platform_id == 3 and encoding_id in (2, 3, 4, 5, 6):
        encodings.insert(0, "utf-16-be")
    for encoding in encodings:
        try:
            value = raw.decode(encoding).strip("\x00 ")
        except (LookupError, UnicodeDecodeError):
            continue
        if value:
            return value
    return None


def sfnt_names(path: Path) -> set[str]:
    """Read family/full/PostScript names from TTF, OTF, or TTC metadata."""
    names: set[str] = set()
    try:
        with path.open("rb") as handle:
            signature = handle.read(4)
            is_collection = signature == b"ttcf"
            if is_collection:
                handle.read(4)
                count_raw = handle.read(4)
                if len(count_raw) != 4:
                    return names
                count = min(struct.unpack(">I", count_raw)[0], 128)
                offsets_raw = handle.read(count * 4)
                if len(offsets_raw) != count * 4:
                    return names
                offsets = struct.unpack(f">{count}I", offsets_raw)
            else:
                offsets = (0,)

            for offset in offsets:
                handle.seek(offset + 4)
                header = handle.read(8)
                if len(header) != 8:
                    continue
                num_tables = struct.unpack(">H", header[:2])[0]
                handle.seek(offset + 12)
                name_offset = None
                for _ in range(min(num_tables, 256)):
                    record = handle.read(16)
                    if len(record) != 16:
                        break
                    tag, _checksum, table_offset, _length = struct.unpack(">4sIII", record)
                    if tag == b"name":
                        # TTC table offsets are absolute from the start of the file.
                        name_offset = table_offset if is_collection else offset + table_offset
                if name_offset is None:
                    continue
                handle.seek(name_offset)
                name_header = handle.read(6)
                if len(name_header) != 6:
                    continue
                _fmt, count, string_offset = struct.unpack(">HHH", name_header)
                records = []
                for _ in range(min(count, 4096)):
                    record = handle.read(12)
                    if len(record) != 12:
                        break
                    records.append(struct.unpack(">HHHHHH", record))
                for platform_id, encoding_id, _language_id, name_id, length, rel_offset in records:
                    if name_id not in (1, 4, 6, 16):
                        continue
                    handle.seek(name_offset + string_offset + rel_offset)
                    value = decode_name(handle.read(length), platform_id, encoding_id)
                    if value:
                        names.add(value)
    except (OSError, struct.error, ValueError):
        pass
    return names


def font_directories(system: str) -> list[Path]:
    user_home = Path.home()
    if system == "Darwin":
        return [Path("/System/Library/Fonts"), Path("/Library/Fonts"), user_home / "Library/Fonts"]
    if system == "Windows":
        win_dir = Path(os.environ.get("WINDIR", "C:/Windows"))
        local_data = Path(os.environ.get("LOCALAPPDATA", user_home / "AppData/Local"))
        return [win_dir / "Fonts", local_data / "Microsoft/Windows/Fonts"]
    return [
        Path("/usr/share/fonts"),
        Path("/usr/local/share/fonts"),
        user_home / ".fonts",
        user_home / ".local/share/fonts",
    ]


def names_from_fc_list() -> set[str]:
    executable = shutil.which("fc-list")
    if not executable:
        return set()
    try:
        completed = subprocess.run(
            [executable, "--format", "%{family}\n"],
            check=False,
            capture_output=True,
            text=True,
            timeout=15,
        )
    except (OSError, subprocess.SubprocessError):
        return set()
    result = set()
    for line in completed.stdout.splitlines():
        result.update(name.strip() for name in line.split(",") if name.strip())
    return result


def inventory(system: str) -> tuple[dict[str, str], list[str]]:
    aliases: dict[str, str] = {}
    sources = []
    fc_names = names_from_fc_list()
    if fc_names:
        sources.append("fontconfig")
        for name in fc_names:
            aliases.setdefault(normalize(name), name)

    # fontconfig already reads platform font metadata. Avoid reparsing every font
    # file when it supplied an inventory; this keeps preflight fast in CI.
    if fc_names:
        return aliases, sources

    files_seen = 0
    extensions = {".ttf", ".otf", ".ttc", ".otc"}
    for directory in font_directories(system):
        if not directory.is_dir():
            continue
        sources.append(str(directory))
        try:
            paths: Iterable[Path] = directory.rglob("*")
            for path in paths:
                if path.suffix.casefold() not in extensions or not path.is_file():
                    continue
                files_seen += 1
                for name in sfnt_names(path):
                    aliases.setdefault(normalize(name), name)
                # Filename matching is weak evidence, but helps with malformed legacy fonts.
                aliases.setdefault(normalize(path.stem), path.stem)
        except OSError:
            continue
    if files_seen:
        sources.append(f"sfnt-metadata:{files_seen}-files")
    return aliases, sources


def configured_fonts(style: dict) -> list[tuple[str, str]]:
    fonts = style.get("fonts", {})
    if not isinstance(fonts, dict):
        raise ValueError("样式中的 fonts 必须是对象")
    result = []
    for role, value in fonts.items():
        if (role.endswith("_font") or role == "font") and isinstance(value, str) and value.strip():
            result.append((role, value.strip()))
    return result


def fallback_kind(role: str) -> str:
    lowered = role.casefold()
    if "ascii" in lowered:
        return "ascii"
    if "kaiti" in lowered or "kai" in lowered:
        return "kai"
    if "heading" in lowered or "sans" in lowered:
        return "sans"
    return "serif"


def candidate_names(system: str, role: str) -> list[str]:
    kind = fallback_kind(role)
    if kind == "ascii":
        return ASCII_FALLBACKS.get(system, ASCII_FALLBACKS["Linux"])
    return CHINESE_FALLBACKS.get(system, CHINESE_FALLBACKS["Linux"])[kind]


def build_report_for_style(style: dict, style_label: str = "<memory>") -> dict:
    """Check an already-loaded style dictionary."""
    system = platform.system() or "Unknown"
    available, sources = inventory(system)
    checks = []
    suggestions = {}
    for role, requested in configured_fonts(style):
        matched = available.get(normalize(requested))
        candidates = []
        for candidate in candidate_names(system, role):
            candidate_match = available.get(normalize(candidate))
            candidates.append(
                {"name": candidate, "available": candidate_match is not None, "matched_name": candidate_match}
            )
        suggestion = next((item["matched_name"] for item in candidates if item["available"]), None)
        if not matched:
            suggestions[role] = suggestion
        checks.append(
            {
                "role": role,
                "requested": requested,
                "available": matched is not None,
                "matched_name": matched,
                "fallback_candidates": candidates,
                "suggested_fallback": None if matched else suggestion,
            }
        )

    missing = [item["role"] for item in checks if not item["available"]]
    warnings = []
    if not sources:
        warnings.append("未找到可读取的系统字体目录或 fontconfig，检测结果可能不完整。")
    if missing:
        warnings.append("当前生成环境缺少部分字体；这不代表目标 Word/WPS 环境缺失。")
        warnings.append("若目标 Word/WPS 已安装所需字体，应保留原字体声明并在目标编辑器中验收。")
    warnings.append("字体名称可被发现不代表目标编辑器一定能渲染中文字形；仍须逐页视觉检查。")
    warnings.append("本工具只检测并建议字体；未修改样式，也未在 DOCX 中应用任何替代。")
    return {
        "ok": not missing,
        "platform": system,
        "style_file": style_label,
        "inventory_sources": sources,
        "font_checks": checks,
        "missing_roles": missing,
        "suggestions": suggestions,
        "substitution_applied": False,
        "warnings": warnings,
    }


def build_report(style_path: Path) -> dict:
    try:
        style = json.loads(style_path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ValueError(f"找不到样式文件：{style_path}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法读取样式 JSON：{exc}") from exc
    return build_report_for_style(style, str(style_path.resolve()))


def build_substitution_override(report: dict) -> dict:
    """Build an explicit StyleManager override from available suggestions.

    The returned mapping is inert until the caller passes it to render.py via
    ``--override``. This keeps substitution visible and reproducible.
    """
    substitutions = {}
    for item in report.get("font_checks", []):
        suggestion = item.get("suggested_fallback")
        requested = item.get("requested")
        if requested and suggestion:
            substitutions[requested] = suggestion
    return {"fonts": {"substitutions": substitutions}}


def print_human(report: dict) -> None:
    print(f"字体预检：{'通过' if report['ok'] else '未通过'}")
    print(f"平台：{report['platform']}")
    print(f"样式：{report['style_file']}")
    for item in report["font_checks"]:
        if item["available"]:
            print(f"  [可用] {item['role']}: {item['requested']} -> {item['matched_name']}")
        else:
            suggestion = item["suggested_fallback"] or "未发现已安装候选"
            print(f"  [缺失] {item['role']}: {item['requested']}；建议候选：{suggestion}")
    for warning in report["warnings"]:
        print(f"警告：{warning}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="检测教案样式配置中的字体是否已安装")
    parser.add_argument("--style", type=Path, default=DEFAULT_STYLE, help="样式 JSON 路径")
    parser.add_argument("--json", action="store_true", help="输出结构化 JSON")
    parser.add_argument("--strict", action="store_true", help="存在缺失字体时返回退出码 1")
    parser.add_argument(
        "--write-override",
        type=Path,
        help="将可用候选写成显式字体替代 override JSON；渲染时用 --override 应用",
    )
    args = parser.parse_args(argv)
    try:
        report = build_report(args.style)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    if args.write_override:
        override = build_substitution_override(report)
        substitutions = override["fonts"]["substitutions"]
        unresolved = [
            item["role"]
            for item in report["font_checks"]
            if not item["available"] and not item["suggested_fallback"]
        ]
        if unresolved:
            print(f"无法为以下字体角色找到已安装回退字体：{', '.join(unresolved)}", file=sys.stderr)
            return 1
        try:
            args.write_override.parent.mkdir(parents=True, exist_ok=True)
            args.write_override.write_text(
                json.dumps(override, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except OSError as exc:
            print(f"无法写入字体 override：{exc}", file=sys.stderr)
            return 2
        report["substitution_override_written"] = str(args.write_override.resolve())
        report["planned_substitutions"] = substitutions
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_human(report)
    return 1 if args.strict and not report["ok"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
