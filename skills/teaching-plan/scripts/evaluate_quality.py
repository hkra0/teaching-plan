#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validate and summarize human/model teaching-quality evaluation records.

This script does not infer pedagogical quality from prose. It verifies that a
reviewer supplied complete D1-D9 scores with evidence, applies the published
thresholds, and reports disagreements that require arbitration.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any

DIMENSIONS = tuple(f"D{i}" for i in range(1, 10))


def verdict_for(scores: dict[str, int], blockers: list[str]) -> str:
    if blockers:
        return "blocked"
    total = sum(scores.values())
    if (
        total >= 32
        and all(scores[dim] >= 3 for dim in DIMENSIONS)
        and sum(scores[dim] == 4 for dim in ("D3", "D4", "D5", "D6", "D7")) >= 4
    ):
        return "excellent"
    if (
        total >= 27
        and scores["D1"] >= 3
        and scores["D2"] >= 3
        and all(scores[dim] >= 2 for dim in ("D3", "D4", "D5", "D6", "D7", "D8"))
        and all(scores[dim] > 0 for dim in DIMENSIONS)
    ):
        return "pass"
    return "fail"


def validate_evaluation(item: Any, index: int) -> tuple[dict[str, Any] | None, list[str]]:
    errors: list[str] = []
    path = f"evaluations[{index}]"
    if not isinstance(item, dict):
        return None, [f"{path} 必须是对象"]

    for key in ("case_id", "reviewer_id"):
        if not isinstance(item.get(key), str) or not item[key].strip():
            errors.append(f"{path}.{key} 必须是非空字符串")

    blockers = item.get("blockers", [])
    if not isinstance(blockers, list) or any(not isinstance(value, str) or not value.strip() for value in blockers):
        errors.append(f"{path}.blockers 必须是字符串数组")
        blockers = []

    dimensions = item.get("dimensions")
    if not isinstance(dimensions, list):
        return None, errors + [f"{path}.dimensions 必须是数组"]

    by_id: dict[str, dict[str, Any]] = {}
    for dim_index, dim in enumerate(dimensions):
        dim_path = f"{path}.dimensions[{dim_index}]"
        if not isinstance(dim, dict):
            errors.append(f"{dim_path} 必须是对象")
            continue
        dim_id = dim.get("id")
        if dim_id not in DIMENSIONS:
            errors.append(f"{dim_path}.id 必须是 D1-D9")
            continue
        if dim_id in by_id:
            errors.append(f"{dim_path}.id 重复")
            continue
        score = dim.get("score")
        if not isinstance(score, int) or isinstance(score, bool) or not 0 <= score <= 4:
            errors.append(f"{dim_path}.score 必须是 0-4 整数")
        evidence = dim.get("evidence_paths")
        if not isinstance(evidence, list) or not evidence or any(not isinstance(value, str) or not value.strip() for value in evidence):
            errors.append(f"{dim_path}.evidence_paths 必须是非空字符串数组")
        if not isinstance(dim.get("reason"), str) or not dim["reason"].strip():
            errors.append(f"{dim_path}.reason 必须是非空字符串")
        by_id[dim_id] = dim

    missing = [dim for dim in DIMENSIONS if dim not in by_id]
    if missing:
        errors.append(f"{path}.dimensions 缺少 {', '.join(missing)}")
    if errors:
        return None, errors

    scores = {dim: by_id[dim]["score"] for dim in DIMENSIONS}
    normalized = dict(item)
    normalized["blockers"] = blockers
    normalized["scores"] = scores
    normalized["total"] = sum(scores.values())
    normalized["verdict"] = verdict_for(scores, blockers)
    return normalized, []


def summarize(
    payload: Any,
    required_case_ids: set[str] | None = None,
    min_reviews_per_case: int = 0,
) -> dict[str, Any]:
    errors: list[str] = []
    if not isinstance(payload, dict):
        return {"valid": False, "errors": ["根对象必须是对象"]}
    raw = payload.get("evaluations")
    if not isinstance(raw, list) or not raw:
        return {"valid": False, "errors": ["evaluations 必须是非空数组"]}

    evaluations: list[dict[str, Any]] = []
    for index, item in enumerate(raw):
        normalized, item_errors = validate_evaluation(item, index)
        errors.extend(item_errors)
        if normalized is not None:
            evaluations.append(normalized)

    disagreements: list[dict[str, Any]] = []
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in evaluations:
        grouped[item["case_id"]].append(item)

    for case_id, reviews in grouped.items():
        reviewer_ids = [review["reviewer_id"] for review in reviews]
        if len(reviewer_ids) != len(set(reviewer_ids)):
            errors.append(f"{case_id} 存在重复 reviewer_id，不能计为独立评审")

    if required_case_ids:
        seen_case_ids = set(grouped)
        missing_cases = sorted(required_case_ids - seen_case_ids)
        unexpected_cases = sorted(seen_case_ids - required_case_ids)
        if missing_cases:
            errors.append(f"缺少必评用例：{', '.join(missing_cases)}")
        if unexpected_cases:
            errors.append(f"出现未声明用例：{', '.join(unexpected_cases)}")
        for case_id in sorted(required_case_ids):
            review_count = len(grouped.get(case_id, []))
            if review_count < min_reviews_per_case:
                errors.append(
                    f"{case_id} 只有 {review_count} 份评审，少于要求的 {min_reviews_per_case} 份"
                )
    for case_id, reviews in grouped.items():
        if len(reviews) < 2:
            continue
        for dim in DIMENSIONS:
            values = [review["scores"][dim] for review in reviews]
            if max(values) - min(values) >= 2:
                disagreements.append({
                    "case_id": case_id,
                    "dimension": dim,
                    "scores": {review["reviewer_id"]: review["scores"][dim] for review in reviews},
                    "requires_arbitration": True,
                })

    return {
        "valid": not errors,
        "errors": errors,
        "evaluations": evaluations,
        "disagreements": disagreements,
        "summary": {
            "count": len(evaluations),
            "blocked": sum(item["verdict"] == "blocked" for item in evaluations),
            "fail": sum(item["verdict"] == "fail" for item in evaluations),
            "pass": sum(item["verdict"] == "pass" for item in evaluations),
            "excellent": sum(item["verdict"] == "excellent" for item in evaluations),
            "arbitrations": len(disagreements),
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="校验并汇总 P1-2 教学质量评分记录。")
    parser.add_argument("input", help="评分记录 JSON")
    parser.add_argument("--json", action="store_true", help="输出完整 JSON")
    parser.add_argument(
        "--require-case",
        action="append",
        default=[],
        help="要求出现的 case_id；可重复传入",
    )
    parser.add_argument(
        "--min-reviews",
        type=int,
        default=0,
        help="每个 --require-case 所需的最少独立评审数",
    )
    args = parser.parse_args(argv)

    try:
        payload = json.loads(Path(args.input).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"无法读取评分记录：{exc}", file=sys.stderr)
        return 2

    if args.min_reviews < 0:
        parser.error("--min-reviews 不能为负数")
    result = summarize(
        payload,
        required_case_ids=set(args.require_case) or None,
        min_reviews_per_case=args.min_reviews,
    )
    if args.json:
        json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
        print()
    else:
        for item in result.get("evaluations", []):
            print(f"{item['case_id']} / {item['reviewer_id']}: {item['total']}/36 {item['verdict']}")
        for issue in result.get("errors", []):
            print(f"[错误] {issue}")
        print(json.dumps(result.get("summary", {}), ensure_ascii=False))
    return 0 if result["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
