# -*- coding: utf-8 -*-
"""
DataNormalizer: Normalizes raw teaching plan content dictionaries into
standardized semantic structures while preserving 100% backward compatibility
with legacy flattened JSON schemas.
"""

import copy

LEGACY_STAGE_DEFINITIONS = [
    ("stage_intro", "intent_intro", "（一）", "任务初探"),
    ("stage_analyze", "intent_analyze", "（二）", "分析任务"),
    ("stage_execute", "intent_execute", "（三）", "任务执行"),
    ("stage_share", "intent_share", "（四）", "任务分享"),
    ("stage_evaluate", "intent_evaluate", "（五）", "任务评价"),
    ("stage_expand", "intent_expand", "（六）", "任务拓展"),
]

class DataNormalizer:
    @staticmethod
    def normalize_content(content_dict):
        """
        Normalize full document content dictionary.
        Returns a clean deep-copied dictionary with standardized semantic fields.
        """
        if not content_dict:
            return {}

        normalized = copy.deepcopy(content_dict)

        # 1. Normalize lessons
        if "lessons" in normalized and isinstance(normalized["lessons"], list):
            normalized["lessons"] = [
                DataNormalizer.normalize_lesson(lesson, idx)
                for idx, lesson in enumerate(normalized["lessons"])
            ]

        return normalized

    @staticmethod
    def normalize_lesson(lesson_dict, lesson_idx=0):
        """
        Normalize an individual lesson plan item.
        Ensures 'stages', 'objectives', and standard fields are structured and accessible.
        """
        if not isinstance(lesson_dict, dict):
            return {}

        lesson = copy.deepcopy(lesson_dict)

        # Normalize Objectives
        knowledge = lesson.get("knowledge_obj") or lesson.get("objectives", {}).get("knowledge", "")
        ability = lesson.get("ability_obj") or lesson.get("objectives", {}).get("ability", "")
        quality = lesson.get("quality_obj") or lesson.get("objectives", {}).get("quality", "")

        lesson["knowledge_obj"] = knowledge
        lesson["ability_obj"] = ability
        lesson["quality_obj"] = quality
        lesson["objectives"] = {
            "knowledge": knowledge,
            "ability": ability,
            "quality": quality,
        }

        # Normalize Stages
        if "stages" in lesson and isinstance(lesson["stages"], list) and len(lesson["stages"]) > 0:
            normalized_stages = []
            for i, st in enumerate(lesson["stages"]):
                if isinstance(st, dict):
                    st_copy = copy.deepcopy(st)
                    if "stage_index" not in st_copy:
                        st_copy["stage_index"] = f"（{DataNormalizer._num_to_chinese(i + 1)}）"
                    if "design" not in st_copy:
                        st_copy["design"] = st_copy.get("content", st_copy.get("text", ""))
                    if "intent" not in st_copy:
                        st_copy["intent"] = st_copy.get("purpose", "")
                    normalized_stages.append(st_copy)
                else:
                    normalized_stages.append({
                        "stage_index": f"（{DataNormalizer._num_to_chinese(i + 1)}）",
                        "stage_name": str(st),
                        "design": "",
                        "intent": ""
                    })
            lesson["stages"] = normalized_stages
        else:
            # Reconstruct stages from legacy flat keys
            stages = []
            has_legacy_keys = any(k in lesson for k, _, _, _ in LEGACY_STAGE_DEFINITIONS)
            if has_legacy_keys:
                for stage_k, intent_k, idx_str, name_str in LEGACY_STAGE_DEFINITIONS:
                    stages.append({
                        "stage_index": idx_str,
                        "stage_name": name_str,
                        "design": lesson.get(stage_k, ""),
                        "intent": lesson.get(intent_k, ""),
                    })
            lesson["stages"] = stages

        # Set default values for other standard fields if absent
        lesson.setdefault("hours", "2课时")
        lesson.setdefault("env_resources", "")
        lesson.setdefault("teaching_methods", "")
        lesson.setdefault("learning_methods", "")
        lesson.setdefault("key_points", "")
        lesson.setdefault("difficult_points", "")
        lesson.setdefault("reflection", "")

        return lesson

    @staticmethod
    def _num_to_chinese(n):
        chinese_digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
        if 1 <= n <= 10:
            return chinese_digits[n]
        elif 11 <= n <= 19:
            return f"十{chinese_digits[n - 10]}"
        elif n == 20:
            return "二十"
        return str(n)
