import copy
import io
import sys
import unittest
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

try:
    import exporting
except ModuleNotFoundError:
    exporting = None


FIELDS = ["商品名称", "SKU编码", "规格", "售价", "库存"]


def issue(field, message, level="待核对", code="price_conflict"):
    return {"code": code, "field": field, "level": level,
            "message": message, "suggestion": "请核对原始资料"}


def sample_report():
    values = [
        ["棉袜", "000087", "米白/均码", "19.90", "0", "保留备注"],
        ['=HYPERLINK("https://example.invalid")', "=2+2", "<script>alert(1)</script>", "-1", "-3", ""],
        ["棉袜", "000087", "米白/均码", "29.90", "6", "重复行仍然保留"],
    ]
    columns = FIELDS + ["原列·备注"]
    rows = []
    for index, values_row in enumerate(values, 2):
        rows.append({
            "row": index,
            "raw": dict(zip(FIELDS + ["备注"], values_row)),
            "data": dict(zip(columns, values_row)),
            "issues": ([issue("售价", "同一SKU出现不同售价")] if index != 3 else
                       [issue("售价", "售价必须大于零", "错误", "invalid_price"),
                        issue("库存", "库存不能为负数", "错误", "invalid_stock")]),
            "changes": ([{"field": "售价", "before": "19.9元", "after": "19.90"}] if index == 2 else []),
            "status": "需修正" if index == 3 else "待核对",
        })
    rows[0]["raw"]["售价"] = "19.9元"
    return {"source": "虚构演示资料.xlsx", "sheet": "商品资料", "generated_at": "2026-11-13 14:03:27 +08:00",
            "mapping": dict(zip(FIELDS, FIELDS)), "columns": columns, "rows": rows,
            "file_issues": [], "notices": ["本文件为虚构演示资料"],
            "summary": {"total_rows": 3, "passed_rows": 0, "issue_rows": 3,
                        "error_rows": 1, "review_rows": 2, "issue_count": 4, "normalized_cells": 1}}


class ExportingTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(exporting, "结果导出功能尚未实现")

    def workbook(self, report=None):
        return load_workbook(io.BytesIO(exporting.export_xlsx(report or sample_report())), data_only=False)

    def test_workbook_has_five_named_sheets_and_keeps_all_rows(self):
        book = self.workbook()
        self.assertEqual(book.sheetnames, ["整理结果", "问题清单", "修改记录", "原始资料", "说明"])
        self.assertEqual(book["整理结果"].max_row, 4)
        self.assertEqual(book["原始资料"].max_row, 4)
        self.assertEqual(book["整理结果"]["A2"].value, 2)
        self.assertEqual(book["整理结果"]["A4"].value, 4)

    def test_sku_leading_zero_and_text_formulas_survive_excel_roundtrip(self):
        book = self.workbook()
        normalized = book["整理结果"]
        header = {cell.value: cell.column for cell in normalized[1]}
        sku_cell = normalized.cell(2, header["SKU编码"])
        self.assertEqual(sku_cell.value, "000087")
        self.assertEqual(sku_cell.data_type, "s")
        self.assertEqual(sku_cell.number_format, "@")
        for sheet in book:
            for row in sheet:
                for cell in row:
                    self.assertNotEqual(cell.data_type, "f", f"{sheet.title}!{cell.coordinate}不可执行公式")
        self.assertEqual(normalized.cell(3, header["SKU编码"]).value, "=2+2")

    def test_valid_numbers_export_as_numbers_but_errors_remain_source_text(self):
        sheet = self.workbook()["整理结果"]
        header = {cell.value: cell.column for cell in sheet[1]}
        self.assertEqual(sheet.cell(2, header["售价"]).value, 19.9)
        self.assertEqual(sheet.cell(2, header["售价"]).number_format, "0.00")
        self.assertEqual(sheet.cell(2, header["库存"]).value, 0)
        self.assertEqual(sheet.cell(3, header["库存"]).value, "-3")
        self.assertEqual(sheet.cell(3, header["售价"]).value, "-1")

    def test_problem_and_change_sheets_include_location_and_values(self):
        book = self.workbook()
        issues = list(book["问题清单"].values)
        self.assertEqual(len(issues), 5)
        self.assertEqual(issues[1][0], 2)
        self.assertIn("同一SKU出现不同售价", str(issues))
        changes = list(book["修改记录"].values)
        self.assertEqual(changes[1], (2, "000087", "售价", "19.9元", "19.90"))

    def test_raw_source_and_extra_columns_are_preserved(self):
        sheet = self.workbook()["原始资料"]
        rows = list(sheet.values)
        headers = list(rows[0])
        self.assertEqual(rows[1][headers.index("售价")], "19.9元")
        self.assertEqual(rows[1][headers.index("备注")], "保留备注")
        self.assertEqual(rows[2][headers.index("商品名称")], '=HYPERLINK("https://example.invalid")')

    def test_excel_layout_and_problem_highlights_are_present(self):
        book = self.workbook()
        for name in ["整理结果", "问题清单", "修改记录", "原始资料"]:
            self.assertEqual(book[name].freeze_panes, "A2")
            self.assertTrue(book[name].auto_filter.ref)
            self.assertGreater(book[name].column_dimensions["A"].width, 8)
        sheet = book["整理结果"]
        headers = {cell.value: cell.column for cell in sheet[1]}
        self.assertNotEqual(sheet.cell(3, headers["库存"]).fill.fgColor.rgb, "00000000")

    def test_entire_table_issues_are_included(self):
        report = sample_report()
        report["file_issues"] = [issue("规格", "未选择规格列", "错误", "missing_column")]
        report["summary"]["issue_count"] += 1
        book = self.workbook(report)
        self.assertIn("整表", str(list(book["问题清单"].values)))
        self.assertIn("未选择规格列", str(list(book["问题清单"].values)))

    def test_html_escapes_every_source_value_and_works_without_javascript(self):
        report = sample_report()
        report["source"] = '<img src=x onerror="alert(1)">'
        report["notices"] = ["<script>source_notice()</script>"]
        report["rows"][0]["issues"][0]["message"] = "<script>issue()</script>"
        html = exporting.export_html(report).decode("utf-8")
        self.assertIn("<!doctype html>", html.lower())
        self.assertIn("charset=\"utf-8\"", html.lower())
        self.assertNotIn("<script", html.lower())
        self.assertNotIn("<img src=x", html)
        self.assertIn("&lt;script&gt;", html)
        self.assertIn("000087", html)
        self.assertIn("原始资料", html)
        self.assertIn("修改记录", html)

    def test_template_has_required_headers_and_separate_instructions(self):
        book = load_workbook(io.BytesIO(exporting.template_xlsx()), data_only=False)
        self.assertEqual(book.sheetnames, ["商品资料", "填写说明"])
        self.assertEqual([cell.value for cell in book["商品资料"][1]], FIELDS)
        self.assertEqual(book["商品资料"].max_row, 1)
        instructions = str(list(book["填写说明"].values))
        self.assertIn("SKU", instructions)
        self.assertIn("文本", instructions)
        self.assertIn("价格", instructions)

    def test_demo_export_keeps_original_headers_and_values(self):
        table = {"source": "虚构商品资料", "sheet": "演示资料", "headers": ["商品标题", "sku_id", "售价", "库存"],
                 "rows": [["棉袜", "000087", "19.9元", "0"], ["=1+1", "=2+2", "-3", "-1"]],
                 "row_numbers": [2, 3]}
        book = load_workbook(io.BytesIO(exporting.table_xlsx(table)), data_only=False)
        rows = list(book.active.values)
        self.assertEqual(list(rows[0]), table["headers"])
        self.assertEqual(list(rows[1]), table["rows"][0])
        self.assertEqual(book.active["A3"].data_type, "s")
        self.assertEqual(book.active["B2"].number_format, "@")

    def test_export_does_not_mutate_input_report(self):
        report = sample_report()
        before = copy.deepcopy(report)
        exporting.export_xlsx(report)
        exporting.export_html(report)
        self.assertEqual(report, before)

    def test_values_beyond_excel_numeric_precision_remain_exact_text(self):
        report = sample_report()
        row = report["rows"][0]
        row["data"]["售价"] = "1234567890123456.79"
        row["data"]["库存"] = "12345678901234567"
        row["issues"] = []
        sheet = self.workbook(report)["整理结果"]
        headers = {cell.value: cell.column for cell in sheet[1]}
        for field in ["售价", "库存"]:
            cell = sheet.cell(2, headers[field])
            self.assertEqual(cell.value, row["data"][field])
            self.assertEqual(cell.data_type, "s")


if __name__ == "__main__":
    unittest.main()
