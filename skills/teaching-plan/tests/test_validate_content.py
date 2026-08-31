# -*- coding: utf-8 -*-

import importlib.util
import json
import subprocess
import sys
import unittest
from copy import deepcopy
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPT = SKILL_DIR / "scripts" / "validate_content.py"
SPEC = importlib.util.spec_from_file_location("validate_content", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


def minimal_valid_content():
    return {
        "output_filename": "测试教案.docx",
        "cover": {
            "title": "教学设计",
            "subtitle": "第一学期",
            "info_items": [["课程名称", "系统管理"]],
            "school": "示范学校",
            "date": "示例日期",
        },
        "schedule": {
            "title": "教学进度表",
            "stats": {
                "weeks_count": "1",
                "week_hours": "4/班",
                "total_hours": "4课时",
                "lecture_hours": "2",
                "lab_hours": "2",
                "exam_hours": "0",
            },
            "rows": [["1", "系统概述", "2课时", "操作练习", "2"]],
        },
        "lessons": [
            {
                "lesson_num": "第1、2课时",
                "title": "系统管理入门",
                "hours": "2课时",
                "knowledge_obj": "说明系统管理的基本概念。",
                "ability_obj": "能完成基础配置。",
                "quality_obj": "养成规范记录的习惯。",
                "key_points": "基础配置流程",
                "difficult_points": "配置参数辨析",
                "teaching_methods": "任务驱动法",
                "learning_methods": "小组协作法",
                "env_resources": "实训机房与任务单",
                "stages": [
                    {
                        "stage_name": "任务导入",
                        "design": "教师展示任务，学生分析要求。",
                        "intent": "建立任务认知。",
                    }
                ],
            }
        ],
    }


class ValidateContentTests(unittest.TestCase):
    def test_bundled_examples_are_valid(self):
        for path in sorted((SKILL_DIR / "examples").glob("*.json")):
            with self.subTest(path=path.name):
                data = json.loads(path.read_text(encoding="utf-8"))
                result = MODULE.validate_content(data)
                self.assertTrue(result["valid"], result["errors"])

    def test_new_stage_schema_and_units_are_accepted(self):
        result = MODULE.validate_content(minimal_valid_content())
        self.assertTrue(result["valid"], result["issues"])
        self.assertEqual(result["summary"], {"errors": 0, "warnings": 0})

    def test_legacy_stage_schema_is_accepted(self):
        data = minimal_valid_content()
        lesson = data["lessons"][0]
        lesson.pop("stages")
        lesson["stage_intro"] = "教师提出任务，学生记录要求。"
        lesson["intent_intro"] = "激活已有经验。"
        result = MODULE.validate_content(data)
        self.assertTrue(result["valid"], result["issues"])

    def test_invalid_top_level_and_empty_content_are_errors(self):
        result = MODULE.validate_content([])
        self.assertFalse(result["valid"])
        self.assertEqual(result["errors"][0]["code"], "invalid_type")

        data = minimal_valid_content()
        data["cover"]["school"] = "   "
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertTrue(any(issue["path"] == "cover.school" for issue in result["errors"]))

    def test_schedule_rows_must_have_exactly_five_columns(self):
        data = minimal_valid_content()
        data["schedule"]["rows"][0].append("多余列")
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertTrue(any(issue["code"] == "invalid_schedule_row" for issue in result["errors"]))

    def test_all_hours_equations_are_blocking(self):
        data = minimal_valid_content()
        data["schedule"]["stats"]["total_hours"] = "5"
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertGreaterEqual(sum(issue["code"] == "hours_mismatch" for issue in result["errors"]), 2)

        data = minimal_valid_content()
        data["schedule"]["rows"][0][4] = "1"
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertTrue(any(issue["path"] == "schedule.rows[0]" for issue in result["errors"]))

    def test_missing_or_empty_stages_are_errors(self):
        data = minimal_valid_content()
        data["lessons"][0]["stages"] = []
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertTrue(any(issue["path"] == "lessons[0].stages" for issue in result["errors"]))

    def test_placeholder_is_warning_only(self):
        data = minimal_valid_content()
        data["lessons"][0]["key_points"] = "待填写 "
        result = MODULE.validate_content(data)
        self.assertTrue(result["valid"])
        self.assertEqual(result["summary"]["warnings"], 1)
        self.assertEqual(result["warnings"][0]["code"], "placeholder")

    def test_unsafe_output_filename_is_blocking(self):
        data = minimal_valid_content()
        data["output_filename"] = "../测试教案.docx"
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertTrue(any(issue["code"] == "unsafe_filename" for issue in result["errors"]))

    def test_single_lesson_profile_does_not_require_cover_or_schedule(self):
        data = minimal_valid_content()
        data["document_profile"] = "single_lesson"
        data.pop("cover")
        data.pop("schedule")
        result = MODULE.validate_content(data)
        self.assertTrue(result["valid"], result["issues"])
        self.assertFalse(any(issue["path"] in {"cover", "schedule"} for issue in result["issues"]))

    def test_draft_unknown_cover_facts_warn_without_inviting_fabrication(self):
        data = minimal_valid_content()
        data["document_status"] = "draft"
        data["cover"]["school"] = ""
        data["cover"]["info_items"][0][1] = ""
        result = MODULE.validate_content(data)
        self.assertTrue(result["valid"], result["issues"])
        self.assertEqual(
            {issue["path"] for issue in result["warnings"] if issue["code"] == "unresolved_fact"},
            {"cover.school", "cover.info_items[0][1]"},
        )

    def test_final_full_profile_still_requires_cover_and_schedule(self):
        data = minimal_valid_content()
        data.pop("cover")
        data.pop("schedule")
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertEqual(
            {issue["path"] for issue in result["errors"] if issue["code"] == "required"},
            {"cover", "schedule"},
        )

    def test_invalid_document_profile_and_status_are_blocking(self):
        data = minimal_valid_content()
        data["document_profile"] = "brief"
        data["document_status"] = "ready"
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertTrue(any(issue["code"] == "invalid_document_profile" for issue in result["errors"]))
        self.assertTrue(any(issue["code"] == "invalid_document_status" for issue in result["errors"]))

    def test_explicit_stage_minutes_must_match_lesson_duration(self):
        data = minimal_valid_content()
        lesson = data["lessons"][0]
        lesson["duration_minutes"] = 45
        lesson["stages"] = [
            {
                "stage_name": "示范",
                "design": "教师示范，学生记录。",
                "intent": "建立方法。",
                "minutes": 10,
            },
            {
                "stage_name": "练习",
                "design": "教师检查，学生完成练习。",
                "intent": "形成技能。",
                "minutes": 20,
            },
        ]
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertTrue(any(issue["path"] == "lessons[0].stages" and issue["code"] == "hours_mismatch" for issue in result["errors"]))

        lesson["stages"][1]["minutes"] = 35
        result = MODULE.validate_content(data)
        self.assertTrue(result["valid"], result["issues"])

    def test_partial_stage_minutes_warn_and_quality_fields_are_typed(self):
        data = minimal_valid_content()
        lesson = data["lessons"][0]
        lesson["duration_minutes"] = 45
        lesson["stages"] = [
            {
                "stage_name": "示范",
                "design": "教师示范，学生记录。",
                "intent": "建立方法。",
                "minutes": 10,
                "objective_refs": ["A1"],
                "evidence": "练习单",
                "criterion": "三项中至少两项正确",
                "support": "提供分步提示",
                "extension": "增加一项变式",
            },
            {
                "stage_name": "练习",
                "design": "教师检查，学生完成练习。",
                "intent": "形成技能。",
            },
        ]
        result = MODULE.validate_content(data)
        self.assertTrue(result["valid"], result["issues"])
        self.assertTrue(any(issue["code"] == "partial_stage_minutes" for issue in result["warnings"]))

        lesson["stages"][0]["objective_refs"] = "A1"
        result = MODULE.validate_content(data)
        self.assertFalse(result["valid"])
        self.assertTrue(any(issue["code"] == "invalid_objective_refs" for issue in result["errors"]))

    def test_cli_exit_codes_follow_error_severity(self):
        example = next((SKILL_DIR / "examples").glob("*.json"))
        valid_run = subprocess.run(
            [sys.executable, str(SCRIPT), str(example), "--json"],
            check=False,
            capture_output=True,
            text=True,
        )
        self.assertEqual(valid_run.returncode, 0, valid_run.stderr)
        self.assertTrue(json.loads(valid_run.stdout)["valid"])

        invalid_run = subprocess.run(
            [sys.executable, str(SCRIPT), "-", "--json"],
            input="{}",
            check=False,
            capture_output=True,
            text=True,
        )
        self.assertEqual(invalid_run.returncode, 1)
        self.assertFalse(json.loads(invalid_run.stdout)["valid"])


if __name__ == "__main__":
    unittest.main()
