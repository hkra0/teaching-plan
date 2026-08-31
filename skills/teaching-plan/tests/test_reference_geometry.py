"""A small geometry fixture sampled read-only from a real school reference.

The values were extracted from the reference teaching plan document.
The test does not read that file, copy its contents, or claim template fidelity.
Only supported page size/margins, table grids/indent and cell margins are covered.
Header/footer distances, original auto table widths, and visual pagination are
intentionally outside this fixture's scope.
"""

import json
import sys
import unittest
from pathlib import Path

from docx.oxml.ns import qn

SKILL_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_DIR / "scripts"))

from doc_builder import TeachingPlanDocBuilder
from style_manager import StyleManager


class SchoolReferenceGeometryTests(unittest.TestCase):
    def test_custom_geometry_from_school_reference_survives_balanced_table_splitting(self):
        page_twips = (11906, 16838)
        margin_twips = 1134
        schedule_grid = [496, 512, 409, 1040, 1740, 496, 1041, 1393, 899, 496]
        lesson_grid = [1438, 1153, 4235, 1696]
        cell_margins = {"top": 0, "bottom": 0, "left": 108, "right": 108}
        template = json.loads((SKILL_DIR / "templates" / "default_vocational.json").read_text(encoding="utf-8"))
        for key, grid, indent in (("schedule", schedule_grid, 0), ("lesson_table", lesson_grid, -102)):
            template[key]["table"].update({"grid_cols_dxa": grid,
                                              "cell_margins_dxa": cell_margins,
                                              "table_indent_dxa": indent})
        manager = StyleManager(style_override={"page": {
            "page_width_pt": page_twips[0] / 20, "page_height_pt": page_twips[1] / 20,
            "top_margin_pt": margin_twips / 20, "bottom_margin_pt": margin_twips / 20,
            "left_margin_pt": margin_twips / 20, "right_margin_pt": margin_twips / 20,
        }})
        content = {
            "schedule": {"rows": [["1", "理论测试", "2", "实作测试", "2"]]},
            "lessons": [{"lesson_num": "第1课时", "title": "几何兼容测试",
                         "stages": [{"stage_name": "任务执行", "design": "教师活动\n观察甲\n观察乙\n学生活动\n记录丙",
                                     "intent": "意图甲\n意图乙"}]}],
        }
        document = TeachingPlanDocBuilder(content, template=template, style_manager=manager).build_document()
        section = document.sections[0]
        self.assertEqual((section.page_width.twips, section.page_height.twips), page_twips)
        for side in ("top", "bottom", "left", "right"):
            self.assertEqual(getattr(section, side + "_margin").twips, margin_twips)
        self.assertEqual(len(document.tables), 4, "Schedule plus independent metadata/process/reflection tables")
        for index, table in enumerate(document.tables):
            expected_grid = schedule_grid if index == 0 else lesson_grid
            self.assertEqual([int(node.get(qn("w:w"))) for node in table._tbl.xpath("w:tblGrid/w:gridCol")], expected_grid)
            self.assertEqual(int(table._tbl.xpath("w:tblPr/w:tblW")[0].get(qn("w:w"))), sum(expected_grid))
            self.assertEqual(int(table._tbl.xpath("w:tblPr/w:tblInd")[0].get(qn("w:w"))), 0 if index == 0 else -102)
            for side, value in cell_margins.items():
                node = table._tbl.xpath("w:tblPr/w:tblCellMar/w:" + side)[0]
                self.assertEqual(int(node.get(qn("w:w"))), value)
            for row in table.rows:
                self.assertEqual(sum(int(node.get(qn("w:w"))) for node in row._tr.xpath("w:tc/w:tcPr/w:tcW")), sum(expected_grid))


if __name__ == "__main__":
    unittest.main()
