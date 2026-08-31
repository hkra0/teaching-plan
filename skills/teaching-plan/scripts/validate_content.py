#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validate teaching-plan content JSON before document rendering.

This module intentionally uses only the Python standard library.  Call
``validate_content(data)`` from other scripts, or execute this file directly.
Validation errors are blocking; warnings are advisory.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path, PurePath
from typing import Any


LEGACY_STAGES = (
    ("stage_intro", "intent_intro"),
    ("stage_analyze", "intent_analyze"),
    ("stage_execute", "intent_execute"),
    ("stage_share", "intent_share"),
    ("stage_evaluate", "intent_evaluate"),
    ("stage_expand", "intent_expand"),
)

LESSON_REQUIRED_TEXT = (
    "lesson_num",
    "title",
    "knowledge_obj",
    "ability_obj",
    "quality_obj",
    "key_points",
    "difficult_points",
    "teaching_methods",
    "learning_methods",
    "env_resources",
)

PLACEHOLDER_PATTERNS = (
    re.compile(r"(?:^|\s)(?:待填写|请填写|待补充|待完善|TBD|TODO|XXX)(?:$|\s)", re.IGNORECASE),
    re.compile(r"(?:\.{3,}|…{2,})"),
    re.compile(r"\{\{[^{}]+\}\}"),
    re.compile(r"<[^<>]*(?:填写|placeholder)[^<>]*>", re.IGNORECASE),
)


def _issue(severity: str, code: str, path: str, message: str) -> dict[str, str]:
    return {"severity": severity, "code": code, "path": path, "message": message}


class _Validator:
    def __init__(self) -> None:
        self.issues: list[dict[str, str]] = []
        self.document_status = "final"
        self.document_profile = "full"

    def error(self, code: str, path: str, message: str) -> None:
        self.issues.append(_issue("error", code, path, message))

    def warning(self, code: str, path: str, message: str) -> None:
        self.issues.append(_issue("warning", code, path, message))

    def require_mapping(self, value: Any, path: str) -> dict[str, Any] | None:
        if not isinstance(value, dict):
            self.error("invalid_type", path, "必须是 JSON 对象")
            return None
        return value

    def require_list(self, value: Any, path: str, *, nonempty: bool = False) -> list[Any] | None:
        if not isinstance(value, list):
            self.error("invalid_type", path, "必须是 JSON 数组")
            return None
        if nonempty and not value:
            self.error("empty_content", path, "不得为空")
        return value

    def require_text(self, mapping: dict[str, Any], key: str, base: str) -> str | None:
        path = f"{base}.{key}" if base else key
        if key not in mapping:
            self.error("required", path, "缺少必填字段")
            return None
        value = mapping[key]
        if not isinstance(value, str):
            self.error("invalid_type", path, "必须是字符串")
            return None
        if not value.strip():
            self.error("empty_content", path, "不得为空")
            return None
        return value

    def validate(self, data: Any) -> None:
        root = self.require_mapping(data, "$")
        if root is None:
            return

        self.document_status = root.get("document_status", "final")
        if self.document_status not in {"draft", "final"}:
            self.error("invalid_document_status", "document_status", "必须是 draft 或 final")
            self.document_status = "final"
        self.document_profile = root.get("document_profile", "full")
        if self.document_profile not in {"full", "single_lesson"}:
            self.error("invalid_document_profile", "document_profile", "必须是 full 或 single_lesson")
            self.document_profile = "full"

        filename = self.require_text(root, "output_filename", "")
        if filename:
            self._validate_filename(filename)

        cover_value = root.get("cover")
        if cover_value is not None:
            cover = self.require_mapping(cover_value, "cover")
            if cover is not None:
                self._validate_cover(cover)
        elif self.document_profile == "full":
            if self.document_status == "draft":
                self.warning("unresolved_section", "cover", "草稿尚未提供封面；不得用示例事实补全")
            else:
                self.error("required", "cover", "完整教案必须提供封面")

        schedule_value = root.get("schedule")
        if schedule_value is not None:
            schedule = self.require_mapping(schedule_value, "schedule")
            if schedule is not None:
                self._validate_schedule(schedule)
        elif self.document_profile == "full":
            if self.document_status == "draft":
                self.warning("unresolved_section", "schedule", "草稿尚未提供进度表；不得虚构整学期数据")
            else:
                self.error("required", "schedule", "完整教案必须提供进度表")

        lessons = self.require_list(root.get("lessons"), "lessons", nonempty=True)
        if lessons is not None:
            for index, lesson in enumerate(lessons):
                self._validate_lesson(lesson, index)

        self._scan_placeholders(root)

    def _validate_filename(self, filename: str) -> None:
        path = PurePath(filename)
        if path.is_absolute() or ".." in path.parts:
            self.error("unsafe_filename", "output_filename", "不得使用绝对路径或 '..' 路径段")
        if "/" in filename or "\\" in filename:
            self.error("unsafe_filename", "output_filename", "只允许文件名，不允许包含目录")
        if not filename.lower().endswith(".docx"):
            self.error("invalid_filename", "output_filename", "输出文件名必须以 .docx 结尾")

    def _validate_cover(self, cover: dict[str, Any]) -> None:
        for key in ("title", "subtitle", "school", "date"):
            path = f"cover.{key}"
            value = cover.get(key)
            if self.document_status == "draft" and (value is None or (isinstance(value, str) and not value.strip())):
                self.warning("unresolved_fact", path, "草稿中该事实尚待确认；应留空，不得从示例补全")
            else:
                self.require_text(cover, key, "cover")

        if self.document_status == "draft" and cover.get("info_items") is None:
            self.warning("unresolved_fact", "cover.info_items", "草稿尚未提供封面信息项")
            return
        items = self.require_list(cover.get("info_items"), "cover.info_items", nonempty=True)
        if items is None:
            return
        for index, item in enumerate(items):
            path = f"cover.info_items[{index}]"
            if not isinstance(item, list) or len(item) != 2:
                self.error("invalid_shape", path, "每项必须是包含名称和值的两列数组")
                continue
            for column, value in enumerate(item):
                if not isinstance(value, str) or not value.strip():
                    item_path = f"{path}[{column}]"
                    if self.document_status == "draft" and column == 1:
                        self.warning("unresolved_fact", item_path, "草稿中的未知值应保持为空")
                    else:
                        self.error("empty_content", item_path, "名称和值均不得为空")

    def _validate_schedule(self, schedule: dict[str, Any]) -> None:
        self.require_text(schedule, "title", "schedule")
        stats = self.require_mapping(schedule.get("stats"), "schedule.stats")
        rows = self.require_list(schedule.get("rows"), "schedule.rows", nonempty=True)

        parsed_stats: dict[str, float] = {}
        if stats is not None:
            for key in (
                "weeks_count",
                "week_hours",
                "total_hours",
                "lecture_hours",
                "lab_hours",
                "exam_hours",
            ):
                path = f"schedule.stats.{key}"
                if key not in stats:
                    self.error("required", path, "缺少必填字段")
                    continue
                value = _parse_hours(stats[key])
                if value is None or value < 0:
                    self.error("invalid_hours", path, "必须包含非负数值，例如 '4/班' 或 '2课时'")
                else:
                    parsed_stats[key] = value

            if all(key in parsed_stats for key in ("weeks_count", "week_hours", "total_hours")):
                expected = parsed_stats["weeks_count"] * parsed_stats["week_hours"]
                self._check_equal(expected, parsed_stats["total_hours"], "schedule.stats.total_hours", "总学时应等于周数 × 周学时")
            components = ("lecture_hours", "lab_hours", "exam_hours")
            if "total_hours" in parsed_stats and all(key in parsed_stats for key in components):
                expected = sum(parsed_stats[key] for key in components)
                self._check_equal(expected, parsed_stats["total_hours"], "schedule.stats.total_hours", "总学时应等于理论、实操和考试学时之和")

        row_hour_total = 0.0
        rows_have_valid_hours = True
        if rows is not None:
            for row_index, row in enumerate(rows):
                path = f"schedule.rows[{row_index}]"
                if not isinstance(row, list) or len(row) != 5:
                    self.error("invalid_schedule_row", path, "每行必须恰好包含 5 列")
                    rows_have_valid_hours = False
                    continue
                for column, value in enumerate(row):
                    if not isinstance(value, (str, int, float)) or isinstance(value, bool) or not str(value).strip():
                        self.error("empty_content", f"{path}[{column}]", "课表单元格不得为空")
                theory = _parse_hours(row[2])
                practical = _parse_hours(row[4])
                if theory is None or practical is None or theory < 0 or practical < 0:
                    self.error("invalid_hours", path, "第 3、5 列必须是非负学时数")
                    rows_have_valid_hours = False
                    continue
                row_total = theory + practical
                row_hour_total += row_total
                if "week_hours" in parsed_stats:
                    self._check_equal(row_total, parsed_stats["week_hours"], path, "每周理论与实操学时之和应等于周学时")

        if rows is not None and "weeks_count" in parsed_stats and not _numbers_equal(len(rows), parsed_stats["weeks_count"]):
            self.warning("schedule_weeks_mismatch", "schedule.rows", "课表行数与授课周数不一致，请确认是否有合并周或缺漏")
        if rows is not None and rows_have_valid_hours and "total_hours" in parsed_stats:
            self._check_equal(row_hour_total, parsed_stats["total_hours"], "schedule.rows", "课表各周学时之和应等于总学时")

    def _check_equal(self, actual: float, expected: float, path: str, message: str) -> None:
        if not _numbers_equal(actual, expected):
            self.error("hours_mismatch", path, f"{message}（计算值 {actual:g}，填写值 {expected:g}）")

    def _validate_lesson(self, lesson: Any, index: int) -> None:
        base = f"lessons[{index}]"
        item = self.require_mapping(lesson, base)
        if item is None:
            return

        for key in LESSON_REQUIRED_TEXT:
            value = self.require_text(item, key, base)
            # Structured objectives are accepted as a compatibility alternative.
            if value is None and key in {"knowledge_obj", "ability_obj", "quality_obj"}:
                objective_key = key.removesuffix("_obj").replace("knowledge", "knowledge").replace("ability", "ability").replace("quality", "quality")
                objectives = item.get("objectives")
                if isinstance(objectives, dict) and isinstance(objectives.get(objective_key), str) and objectives[objective_key].strip():
                    self.issues = [issue for issue in self.issues if not (issue["path"] == f"{base}.{key}" and issue["code"] in {"required", "empty_content"})]

        if "hours" not in item:
            self.error("required", f"{base}.hours", "缺少必填字段")
        else:
            hours = _parse_hours(item["hours"])
            if hours is None or hours <= 0:
                self.error("invalid_hours", f"{base}.hours", "必须包含大于 0 的学时数，例如 '2课时'")

        duration_minutes = None
        if "duration_minutes" in item:
            duration_minutes = _parse_hours(item["duration_minutes"])
            if duration_minutes is None or duration_minutes <= 0:
                self.error("invalid_minutes", f"{base}.duration_minutes", "必须是大于 0 的分钟数")

        stages = item.get("stages")
        if stages is not None:
            self._validate_new_stages(stages, base, duration_minutes=duration_minutes)
        else:
            self._validate_legacy_stages(item, base)

    def _validate_new_stages(self, stages: Any, base: str, *, duration_minutes: float | None = None) -> None:
        items = self.require_list(stages, f"{base}.stages", nonempty=True)
        if items is None:
            return
        stage_minutes: list[float] = []
        timed_stages = 0
        for index, stage in enumerate(items):
            path = f"{base}.stages[{index}]"
            item = self.require_mapping(stage, path)
            if item is None:
                continue
            for key in ("stage_name", "design", "intent"):
                self.require_text(item, key, path)
            if "minutes" in item:
                minutes = _parse_hours(item["minutes"])
                if minutes is None or minutes <= 0:
                    self.error("invalid_minutes", f"{path}.minutes", "必须是大于 0 的分钟数")
                else:
                    timed_stages += 1
                    stage_minutes.append(minutes)
            if "objective_refs" in item:
                refs = item["objective_refs"]
                if not isinstance(refs, list) or not refs or any(not isinstance(ref, str) or not ref.strip() for ref in refs):
                    self.error("invalid_objective_refs", f"{path}.objective_refs", "必须是非空目标编号字符串数组")
            for key in ("evidence", "criterion", "support", "extension"):
                if key in item and (not isinstance(item[key], str) or not item[key].strip()):
                    self.error("invalid_quality_evidence", f"{path}.{key}", "提供时必须是非空字符串")

        if timed_stages and timed_stages != len(items):
            self.warning("partial_stage_minutes", f"{base}.stages", "只标注了部分环节时间，无法核算整课时长")
        if duration_minutes is not None and timed_stages == len(items):
            self._check_equal(
                sum(stage_minutes),
                duration_minutes,
                f"{base}.stages",
                "各环节分钟数之和应等于本课 duration_minutes",
            )

    def _validate_legacy_stages(self, lesson: dict[str, Any], base: str) -> None:
        present = [(stage, intent) for stage, intent in LEGACY_STAGES if stage in lesson or intent in lesson]
        if not present:
            self.error("required", f"{base}.stages", "必须提供 stages[] 或旧版 stage_* / intent_* 环节字段")
            return
        for stage, intent in present:
            self.require_text(lesson, stage, base)
            self.require_text(lesson, intent, base)

    def _scan_placeholders(self, value: Any, path: str = "$") -> None:
        if isinstance(value, dict):
            for key, child in value.items():
                self._scan_placeholders(child, key if path == "$" else f"{path}.{key}")
        elif isinstance(value, list):
            for index, child in enumerate(value):
                self._scan_placeholders(child, f"{path}[{index}]")
        elif isinstance(value, str) and value.strip():
            if any(pattern.search(value) for pattern in PLACEHOLDER_PATTERNS):
                self.warning("placeholder", path, "发现疑似占位符，请在交付前确认或替换")


def _parse_hours(value: Any) -> float | None:
    """Extract the first finite number from values such as ``4/班`` or ``2课时``."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None
    if not isinstance(value, str):
        return None
    match = re.search(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", value.strip())
    if not match:
        return None
    number = float(match.group(0))
    return number if math.isfinite(number) else None


def _numbers_equal(left: float, right: float) -> bool:
    return math.isclose(left, right, rel_tol=1e-9, abs_tol=1e-9)


def validate_content(data: Any) -> dict[str, Any]:
    """Return a structured validation result for decoded content JSON.

    The result contains ``valid``, ``errors``, ``warnings`` and the combined
    ``issues`` list.  Consumers must stop rendering when ``valid`` is false.
    """
    validator = _Validator()
    validator.validate(data)
    errors = [issue for issue in validator.issues if issue["severity"] == "error"]
    warnings = [issue for issue in validator.issues if issue["severity"] == "warning"]
    return {
        "valid": not errors,
        "errors": errors,
        "warnings": warnings,
        "issues": validator.issues,
        "summary": {"errors": len(errors), "warnings": len(warnings)},
    }


def _load_json(source: str) -> Any:
    if source == "-":
        return json.load(sys.stdin)
    with Path(source).open("r", encoding="utf-8") as handle:
        return json.load(handle)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="校验教案内容 JSON；错误阻断，警告可继续。")
    parser.add_argument("input", help="JSON 文件路径；使用 - 从标准输入读取")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出完整校验结果")
    args = parser.parse_args(argv)

    try:
        data = _load_json(args.input)
    except (OSError, json.JSONDecodeError) as exc:
        print(f"无法读取 JSON：{exc}", file=sys.stderr)
        return 2

    result = validate_content(data)
    if args.json:
        json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
        print()
    else:
        for issue in result["issues"]:
            label = "错误" if issue["severity"] == "error" else "警告"
            print(f"[{label}] {issue['path']}: {issue['message']} ({issue['code']})")
        summary = result["summary"]
        print(f"校验完成：{summary['errors']} 个错误，{summary['warnings']} 个警告")
    return 0 if result["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
