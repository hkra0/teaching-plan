# -*- coding: utf-8 -*-

import importlib.util
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "evaluate_quality.py"
SPEC = importlib.util.spec_from_file_location("evaluate_quality", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


def evaluation(case_id="case", reviewer_id="reviewer", scores=None, blockers=None):
    scores = scores or {f"D{i}": 3 for i in range(1, 10)}
    return {
        "case_id": case_id,
        "reviewer_id": reviewer_id,
        "blockers": blockers or [],
        "dimensions": [
            {
                "id": dim,
                "score": score,
                "evidence_paths": ["lessons[0]"],
                "reason": "有可核验的路径和理由。",
            }
            for dim, score in scores.items()
        ],
    }


class QualityEvaluationTests(unittest.TestCase):
    def test_complete_record_is_scored(self):
        result = MODULE.summarize({"evaluations": [evaluation()]})
        self.assertTrue(result["valid"], result["errors"])
        self.assertEqual(result["evaluations"][0]["total"], 27)
        self.assertEqual(result["evaluations"][0]["verdict"], "pass")

    def test_blocker_overrides_high_score(self):
        scores = {f"D{i}": 4 for i in range(1, 10)}
        result = MODULE.summarize({"evaluations": [evaluation(scores=scores, blockers=["B01"])]})
        self.assertEqual(result["evaluations"][0]["verdict"], "blocked")

    def test_zero_dimension_and_low_total_fail(self):
        scores = {f"D{i}": 3 for i in range(1, 10)}
        scores["D9"] = 0
        result = MODULE.summarize({"evaluations": [evaluation(scores=scores)]})
        self.assertEqual(result["evaluations"][0]["verdict"], "fail")

    def test_excellent_requires_four_core_fours(self):
        scores = {f"D{i}": 3 for i in range(1, 10)}
        for dim in ("D3", "D4", "D5", "D6", "D7"):
            scores[dim] = 4
        result = MODULE.summarize({"evaluations": [evaluation(scores=scores)]})
        self.assertEqual(result["evaluations"][0]["verdict"], "excellent")

    def test_missing_dimension_and_evidence_are_invalid(self):
        item = evaluation()
        item["dimensions"].pop()
        item["dimensions"][0]["evidence_paths"] = []
        result = MODULE.summarize({"evaluations": [item]})
        self.assertFalse(result["valid"])
        self.assertTrue(any("缺少 D9" in error for error in result["errors"]))
        self.assertTrue(any("evidence_paths" in error for error in result["errors"]))

    def test_two_point_reviewer_gap_requires_arbitration(self):
        first = evaluation(case_id="same", reviewer_id="a")
        scores = {f"D{i}": 3 for i in range(1, 10)}
        scores["D7"] = 1
        second = evaluation(case_id="same", reviewer_id="b", scores=scores)
        result = MODULE.summarize({"evaluations": [first, second]})
        self.assertEqual(result["summary"]["arbitrations"], 1)
        self.assertEqual(result["disagreements"][0]["dimension"], "D7")

    def test_required_cases_need_two_distinct_reviewers(self):
        result = MODULE.summarize(
            {"evaluations": [evaluation(case_id="TC-01", reviewer_id="a")]},
            required_case_ids={"TC-01", "TC-02"},
            min_reviews_per_case=2,
        )
        self.assertFalse(result["valid"])
        self.assertTrue(any("TC-01 只有 1 份评审" in error for error in result["errors"]))
        self.assertTrue(any("缺少必评用例：TC-02" in error for error in result["errors"]))

    def test_duplicate_reviewer_cannot_count_as_independent(self):
        result = MODULE.summarize(
            {
                "evaluations": [
                    evaluation(case_id="TC-01", reviewer_id="same"),
                    evaluation(case_id="TC-01", reviewer_id="same"),
                ]
            },
            required_case_ids={"TC-01"},
            min_reviews_per_case=2,
        )
        self.assertFalse(result["valid"])
        self.assertTrue(any("重复 reviewer_id" in error for error in result["errors"]))


if __name__ == "__main__":
    unittest.main()
