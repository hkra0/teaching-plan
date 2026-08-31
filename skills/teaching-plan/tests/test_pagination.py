"""Pure, lossless pagination-block tests; no document runtime required."""

import itertools
import sys
import unittest
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from pagination import split_paragraph_blocks  # noqa: E402


def is_heading(line):
    return line.strip().startswith(("活动", "教师活动", "学生活动"))


class ParagraphBlockTests(unittest.TestCase):
    def assert_blocks(self, text, expected):
        actual = split_paragraph_blocks(text, is_heading)
        self.assertEqual(actual, expected)
        self.assertEqual("\n".join(actual), text)
        self.assertTrue(actual)

    def test_empty_string(self):
        self.assert_blocks("", [""])

    def test_only_blank_lines(self):
        self.assert_blocks("\n\n", ["\n\n"])
        self.assert_blocks(" \n\t\n", [" \n\t\n"])

    def test_body_paragraphs_are_separate(self):
        self.assert_blocks("正文一\n正文二", ["正文一", "正文二"])

    def test_heading_chain_keeps_first_body(self):
        self.assert_blocks(
            "活动1\n教师活动\n1. 观察作品\n2. 比较作品\n学生活动\n1. 记录结果",
            ["活动1\n教师活动\n1. 观察作品", "2. 比较作品", "学生活动\n1. 记录结果"],
        )

    def test_blank_lines_do_not_break_heading_chain(self):
        self.assert_blocks(
            "活动1\n\n教师活动\n \n1. 观察作品\n\n2. 比较作品",
            ["活动1\n\n教师活动\n \n1. 观察作品", "\n2. 比较作品"],
        )

    def test_leading_and_trailing_blank_lines_are_retained(self):
        self.assert_blocks("\n正文\n\n", ["\n正文", "\n"])
        self.assert_blocks("正文\n", ["正文", ""])

    def test_trailing_heading_is_not_lost_or_joined_backwards(self):
        self.assert_blocks("正文\n学生活动", ["正文", "学生活动"])
        self.assert_blocks("正文\n活动2\n\n教师活动\n", ["正文", "活动2\n\n教师活动\n"])

    def test_heading_only_input_is_retained(self):
        self.assert_blocks("活动1\n教师活动", ["活动1\n教师活动"])

    def test_classifier_receives_original_nonblank_lines(self):
        seen = []

        def classify(line):
            seen.append(line)
            return line.strip() == "标题"

        text = " \n 标题 \n\t\n 正文 \n"
        blocks = split_paragraph_blocks(text, classify)
        self.assertEqual(seen, [" 标题 ", " 正文 "])
        self.assertEqual(blocks, [" \n 标题 \n\t\n 正文 ", ""])
        self.assertEqual("\n".join(blocks), text)

    def test_crlf_and_unicode_are_preserved(self):
        text = "活动1\r\n教师活动\r\n1. 中文、English 与 emoji 🙂\r\n"
        blocks = split_paragraph_blocks(text, is_heading)
        self.assertEqual("\n".join(blocks), text)

    def test_lossless_for_all_short_line_combinations(self):
        choices = ["", " ", "活动1", "教师活动", "正文"]
        for length in range(1, 5):
            for lines in itertools.product(choices, repeat=length):
                text = "\n".join(lines)
                with self.subTest(text=text):
                    blocks = split_paragraph_blocks(text, is_heading)
                    self.assertTrue(blocks)
                    self.assertEqual("\n".join(blocks), text)

    def test_invalid_arguments_fail_explicitly(self):
        with self.assertRaises(TypeError):
            split_paragraph_blocks(None, is_heading)
        with self.assertRaises(TypeError):
            split_paragraph_blocks("正文", None)


if __name__ == "__main__":
    unittest.main()
