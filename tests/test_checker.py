"""Business-boundary tests for product-table ingestion and checking."""

import io
import sys
import unittest
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
try:
    import checker
except ModuleNotFoundError:
    checker = None


FIELDS = ["商品名称", "SKU编码", "规格", "售价", "库存"]


def table(rows, headers=None, **extra):
    result = {
        "source": "测试.csv", "sheet": "CSV", "sheet_names": [],
        "headers": headers or FIELDS,
        "rows": rows, "row_numbers": list(range(2, len(rows) + 2)),
        "formulas": [], "numeric_cells": [], "notices": [],
    }
    result.update(extra)
    return result


def workbook_bytes(rows, extra_sheet=None):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "商品资料"
    for row in rows:
        sheet.append(row)
    if extra_sheet:
        second = workbook.create_sheet("第二批")
        for row in extra_sheet:
            second.append(row)
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class CheckerTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(checker, "商品资料读取与检查核心尚未实现")

    def analyze(self, rows, headers=None, **extra):
        return checker.analyze_table(table(rows, headers, **extra))

    def test_valid_zero_stock_and_text_identifier(self):
        report = self.analyze([["棉袜", "000087", "米白/均码", "19.9", "0"]])
        row = report["rows"][0]
        self.assertEqual(row["status"], "通过")
        self.assertEqual(row["data"]["SKU编码"], "000087")
        self.assertEqual(row["data"]["库存"], "0")
        self.assertEqual(row["data"]["售价"], "19.90")

    def test_normalization_keeps_original_values_and_additional_columns(self):
        headers = FIELDS + ["材质", "备注"]
        report = self.analyze([[" 短袖 ", " A01 ", "白/M", "￥１,２９９.９０元", "３.０", " 棉 ", "需人工复查"]], headers)
        row = report["rows"][0]
        self.assertEqual(row["raw"]["商品名称"], " 短袖 ")
        self.assertEqual(row["data"]["商品名称"], "短袖")
        self.assertEqual(row["data"]["售价"], "1299.90")
        self.assertEqual(row["data"]["库存"], "3")
        self.assertEqual(row["data"]["原列·材质"], "棉")
        self.assertEqual(row["data"]["原列·备注"], "需人工复查")
        self.assertEqual(report["summary"]["normalized_cells"], len(row["changes"]))

    def test_missing_values_are_marked_without_guessing(self):
        report = self.analyze([["未知", "A01", "", "待确认", ""]])
        row = report["rows"][0]
        self.assertEqual(row["status"], "需修正")
        self.assertEqual({i["field"] for i in row["issues"]}, {"商品名称", "规格", "售价", "库存"})
        self.assertEqual(row["data"]["售价"], "待确认")

    def test_aliases_are_recognized_but_price_types_not_guessed(self):
        headers = ["商品标题", "sku_id", "规格组合", "单价", "可售库存"]
        report = self.analyze([["袜", "A01", "均码", "19", "0"]], headers)
        self.assertEqual(report["summary"]["passed_rows"], 1)
        for price_header in ["供货价", "拼单价", "单买价", "建议零售价"]:
            with self.subTest(price_header=price_header):
                mapping = checker.guess_mapping(["商品名称", "SKU编码", "规格", price_header, "库存"])
                self.assertIsNone(mapping["售价"])

    def test_ambiguous_aliases_require_manual_mapping(self):
        headers = FIELDS + ["单价"]
        self.assertIsNone(checker.guess_mapping(headers)["售价"])
        original = table([["袜", "A01", "均码", "19", "0", "18"]], headers)
        report = checker.analyze_table(original)
        self.assertEqual(report["summary"]["passed_rows"], 0)
        mapping = {field: field for field in FIELDS}
        selected = checker.analyze_table(original, mapping)
        self.assertEqual(selected["summary"]["passed_rows"], 1)
        self.assertEqual(selected["rows"][0]["data"]["原列·单价"], "18")

    def test_invalid_or_reused_manual_mapping_is_rejected(self):
        original = table([["袜", "A01", "均码", "19", "0"]])
        for mapping in [{"商品名称": "SKU编码", "SKU编码": "SKU编码"}, {"商品名称": "不存在的列"}]:
            with self.subTest(mapping=mapping), self.assertRaises(ValueError):
                checker.analyze_table(original, mapping)

    def test_missing_column_cannot_produce_false_pass(self):
        report = self.analyze([["袜", "A01", "19", "0"]], ["商品名称", "SKU编码", "售价", "库存"])
        self.assertTrue(any(i["field"] == "规格" for i in report["file_issues"]))
        self.assertEqual(report["summary"]["passed_rows"], 0)
        self.assertEqual(report["rows"][0]["data"]["规格"], "")

    def test_duplicate_sku_retains_both_rows_for_review(self):
        rows = [["袜", "A01", "均码", "19", "2"], ["袜", "A01", "均码", "19.00", "2"]]
        report = self.analyze(rows)
        self.assertEqual(len(report["rows"]), 2)
        self.assertEqual(report["summary"]["review_rows"], 2)
        self.assertTrue(all(any(i["code"] == "duplicate_sku" for i in row["issues"]) for row in report["rows"]))
        self.assertFalse(any(i["code"] == "price_conflict" for row in report["rows"] for i in row["issues"]))

    def test_price_conflict_is_review_and_does_not_pick_a_price(self):
        report = self.analyze([["袜", "A01", "均码", "19", "2"], ["袜", "A01", "均码", "29", "2"]])
        self.assertEqual([r["data"]["售价"] for r in report["rows"]], ["19.00", "29.00"])
        for row in report["rows"]:
            conflicts = [i for i in row["issues"] if i["code"] == "price_conflict"]
            self.assertEqual(len(conflicts), 1)
            self.assertEqual(conflicts[0]["level"], "待核对")

    def test_invalid_prices_are_neither_accepted_nor_rounded(self):
        for price in ["0", "-1", "12.345", "NaN", "Infinity", "1e3", "12,34", "$19.00", "免费"]:
            with self.subTest(price=price):
                row = self.analyze([["袜", "A01", "均码", price, "0"]])["rows"][0]
                self.assertTrue(any(i["field"] == "售价" and i["level"] == "错误" for i in row["issues"]))
                self.assertEqual(row["data"]["售价"], price)

    def test_long_price_does_not_hide_extra_decimal_places(self):
        price = "123456789123456789123456789.123"
        row = self.analyze([["袜", "A01", "均码", price, "0"]])["rows"][0]
        self.assertEqual(row["status"], "需修正")
        self.assertEqual(row["data"]["售价"], price)

    def test_csv_illegal_control_characters_are_rejected(self):
        for character in ["\x01", "\x0b"]:
            with self.subTest(character=repr(character)), self.assertRaises(ValueError):
                self.analyze([["袜" + character, "A01", "均码", "19", "0"]])

    def test_valid_price_formats_are_equivalent(self):
        for price in ["19", "19.00", "１９．００", "￥19.00", "19元"]:
            with self.subTest(price=price):
                row = self.analyze([["袜", "A01", "均码", price, "0"]])["rows"][0]
                self.assertEqual(row["data"]["售价"], "19.00")
                self.assertEqual(row["status"], "通过")

    def test_negative_fractional_or_text_stock_does_not_pass(self):
        for stock in ["-1", "2.5", "NaN", "十件", "1e3", "10件"]:
            with self.subTest(stock=stock):
                row = self.analyze([["袜", "A01", "均码", "19", stock]])["rows"][0]
                self.assertEqual(row["status"], "需修正")
                self.assertEqual(row["data"]["库存"], stock)

    def test_csv_encodings_and_empty_line_row_numbers(self):
        text = "商品名称,SKU编码,规格,售价,库存\r\n\r\n棉袜,000087,均码,19.9,0\r\n"
        for encoding in ["utf-8-sig", "gb18030"]:
            with self.subTest(encoding=encoding):
                data = checker.load_upload("资料.csv", text.encode(encoding))
                self.assertEqual(data["row_numbers"], [3])
                self.assertEqual(data["rows"][0][1], "000087")
                self.assertEqual(checker.analyze_table(data)["summary"]["passed_rows"], 1)

    def test_csv_delimiter_detection(self):
        text = "商品名称\tSKU编码\t规格\t售价\t库存\n袜\t000087\t均码\t19\t0\n"
        data = checker.load_upload("资料.csv", text.encode())
        self.assertEqual(data["headers"], FIELDS)

    def test_empty_bad_or_unsupported_files_are_rejected(self):
        for name, data in [("空.csv", b""), ("文件.txt", b"abc"), ("坏.xlsx", b"abc"), ("旧.xls", b"abc"), ("空.csv", "商品名称,售价\n".encode()), ("空.csv", b"\x00abc")]:
            with self.subTest(name=name, data=data), self.assertRaises(ValueError):
                checker.load_upload(name, data)

    def test_malformed_csv_is_rejected(self):
        with self.assertRaises(ValueError):
            checker.load_upload("坏.csv", '商品名称,SKU编码,规格,售价,库存\n"袜,A01,均码,19,0'.encode())

    def test_extra_cells_duplicate_or_blank_headers_are_rejected(self):
        files = ["商品名称,售价\n袜,19,额外值\n", "商品名称,商品名称\n袜,19\n", "商品名称,,售价\n袜,额外值,19\n"]
        for text in files:
            with self.subTest(text=text), self.assertRaises(ValueError):
                checker.load_upload("坏表.csv", text.encode())

    def test_xlsx_preserves_text_identifier_and_zero_inventory(self):
        data = checker.load_upload("资料.xlsx", workbook_bytes([FIELDS, ["袜", "000087", "均码", 19.9, 0]]))
        report = checker.analyze_table(data)
        self.assertEqual(data["sheet_names"], ["商品资料"])
        self.assertEqual(report["sheet"], "商品资料")
        self.assertEqual(report["summary"]["passed_rows"], 1)
        self.assertEqual(report["rows"][0]["data"]["SKU编码"], "000087")

    def test_excel_numeric_sku_requires_human_review(self):
        data = checker.load_upload("资料.xlsx", workbook_bytes([FIELDS, ["袜", 87, "均码", 19.9, 0]]))
        row = checker.analyze_table(data)["rows"][0]
        self.assertEqual(row["status"], "待核对")
        self.assertTrue(any(i["code"] == "numeric_sku" for i in row["issues"]))

    def test_excel_formula_is_not_treated_as_verified_data(self):
        data = checker.load_upload("资料.xlsx", workbook_bytes([FIELDS, ["袜", "A01", "均码", "=10+9.9", 0]]))
        row = checker.analyze_table(data)["rows"][0]
        self.assertNotEqual(row["status"], "通过")
        self.assertEqual(row["raw"]["售价"], "=10+9.9")
        self.assertTrue(any(i["code"] == "formula" for i in row["issues"]))

    def test_formula_in_extra_column_is_also_disclosed(self):
        data = checker.load_upload("资料.xlsx", workbook_bytes([FIELDS + ["备注"], ["袜", "A01", "均码", 19.9, 0, "=1+1"]]))
        row = checker.analyze_table(data)["rows"][0]
        self.assertEqual(row["status"], "待核对")
        self.assertTrue(any(i["field"] == "原列·备注" and i["code"] == "formula" for i in row["issues"]))

    def test_excel_sheet_can_be_selected(self):
        payload = workbook_bytes([FIELDS, ["袜", "A01", "均码", 19.9, 0]], [FIELDS, ["短袖", "B01", "M", 39, 2]])
        data = checker.load_upload("资料.xlsx", payload, "第二批")
        self.assertEqual(data["sheet_names"], ["商品资料", "第二批"])
        self.assertEqual(data["rows"][0][0], "短袖")
        with self.assertRaises(ValueError):
            checker.load_upload("资料.xlsx", payload, "不存在")

    def test_empty_first_sheet_does_not_hide_later_products(self):
        for first_sheet in [[], [FIELDS]]:
            with self.subTest(first_sheet=first_sheet):
                payload = workbook_bytes(first_sheet, [FIELDS, ["短袖", "B01", "M", 39, 2]])
                data = checker.load_upload("资料.xlsx", payload)
                self.assertEqual(data["sheet"], "第二批")
                self.assertEqual(data["sheet_names"], ["商品资料", "第二批"])

    def test_hidden_sheets_are_not_offered(self):
        workbook = Workbook()
        hidden = workbook.active
        hidden.title = "隐藏资料"
        hidden.sheet_state = "hidden"
        visible = workbook.create_sheet("公开资料")
        visible.append(FIELDS)
        visible.append(["袜", "A01", "均码", 19.9, 0])
        output = io.BytesIO()
        workbook.save(output)
        workbook.close()
        data = checker.load_upload("资料.xlsx", output.getvalue())
        self.assertEqual(data["sheet_names"], ["公开资料"])

    def test_sheet_metadata_does_not_parse_an_invalid_cover(self):
        self.assertTrue(callable(getattr(checker, "list_workbook_sheets", None)), "尚未提供独立工作表元数据接口")
        payload = workbook_bytes([[""], ["这是封面说明，不是商品表"]], [FIELDS, ["袜", "A01", "均码", 19, 0]])
        from unittest.mock import patch
        with patch.object(checker, "_worksheet_rows", side_effect=AssertionError("元数据接口不应解析商品行")):
            self.assertEqual(checker.list_workbook_sheets(payload), ["商品资料", "第二批"])
        with self.assertRaises(ValueError):
            checker.load_upload("封面.xlsx", payload)
        self.assertEqual(checker.load_upload("封面.xlsx", payload, "第二批")["rows"][0][0], "袜")

    def test_sheet_metadata_validates_file_and_decompression_limits(self):
        self.assertTrue(callable(getattr(checker, "list_workbook_sheets", None)), "尚未提供独立工作表元数据接口")
        for payload in [b"", b"not an xlsx", b"x" * (checker.MAX_BYTES + 1)]:
            with self.subTest(size=len(payload)), self.assertRaises(ValueError):
                checker.list_workbook_sheets(payload)
        from unittest.mock import patch
        payload = workbook_bytes([FIELDS, ["袜", "A01", "均码", 19, 0]])
        with patch.object(checker, "MAX_UNCOMPRESSED_BYTES", 100), self.assertRaises(ValueError):
            checker.list_workbook_sheets(payload)

    def test_sheet_metadata_requires_visible_sheet(self):
        self.assertTrue(callable(getattr(checker, "list_workbook_sheets", None)), "尚未提供独立工作表元数据接口")
        import zipfile
        payload = workbook_bytes([FIELDS, ["袜", "A01", "均码", 19, 0]])
        output = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(payload)) as source, zipfile.ZipFile(output, "w") as target:
            for entry in source.infolist():
                value = source.read(entry.filename)
                if entry.filename == "xl/workbook.xml":
                    value = value.replace(b'state="visible"', b'state="hidden"')
                target.writestr(entry, value)
        with self.assertRaisesRegex(ValueError, "可见工作表"):
            checker.list_workbook_sheets(output.getvalue())

    def test_size_row_column_and_cell_limits_are_enforced(self):
        with self.assertRaises(ValueError):
            checker.load_upload("大.csv", b"a" * (checker.MAX_BYTES + 1))
        with self.assertRaises(ValueError):
            self.analyze([["袜", str(i), "均码", "19", "0"] for i in range(checker.MAX_ROWS + 1)])
        with self.assertRaises(ValueError):
            self.analyze([["x"] * 81], [str(i) for i in range(81)])
        with self.assertRaises(ValueError):
            self.analyze([["袜" * 32768, "A01", "均码", "19", "0"]])

    def test_wide_excel_row_is_rejected_before_reading_later_rows(self):
        from types import SimpleNamespace

        class GuardedWorksheet:
            def reset_dimensions(self):
                pass

            def iter_rows(self):
                yield tuple(SimpleNamespace(value=value, data_type="s") for value in FIELDS)
                wide = ["袜", "A01", "均码", "19", "0"] + [None] * (checker.MAX_COLS - len(FIELDS)) + ["远列实际数据"]
                yield tuple(SimpleNamespace(value=value, data_type="s") for value in wide)
                raise AssertionError("遇到超宽有效数据后不应继续读取后续商品行")

        with self.assertRaisesRegex(ValueError, "80列|第2行|列数"):
            checker._worksheet_rows(GuardedWorksheet())

    def test_far_excel_formatting_without_values_keeps_real_columns(self):
        from openpyxl.styles import Font
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(FIELDS + ["品牌"])
        sheet.append(["袜", "000087", "均码", 19, 0, "虚构品牌"])
        sheet["XFD2"].font = Font(bold=True)
        output = io.BytesIO()
        workbook.save(output)
        workbook.close()
        data = checker.load_upload("远列格式.xlsx", output.getvalue())
        self.assertEqual(data["headers"], FIELDS + ["品牌"])
        self.assertEqual(len(data["rows"][0]), 6)
        row = checker.analyze_table(data)["rows"][0]
        self.assertEqual(row["status"], "通过")
        self.assertEqual(row["data"]["原列·品牌"], "虚构品牌")

    def test_demo_has_ten_rows_and_auditable_summary(self):
        data = checker.demo_table()
        self.assertEqual(len(data["rows"]), 10)
        self.assertIn("虚构", data["source"])
        self.assertTrue(any("000087" in row for row in data["rows"]))
        report = checker.analyze_table(data)
        summary = report["summary"]
        self.assertEqual(summary["total_rows"], 10)
        self.assertEqual(summary["passed_rows"] + summary["issue_rows"], 10)
        self.assertEqual(summary["error_rows"] + summary["review_rows"], summary["issue_rows"])
        self.assertEqual(summary["issue_count"], len(report["file_issues"]) + sum(len(r["issues"]) for r in report["rows"]))
        self.assertEqual(summary["normalized_cells"], sum(len(r["changes"]) for r in report["rows"]))

    def test_analysis_does_not_mutate_the_input(self):
        original = table([[" 袜 ", " A01 ", "均码", "19", "0"]])
        import copy
        before = copy.deepcopy(original)
        checker.analyze_table(original)
        self.assertEqual(original, before)


if __name__ == "__main__":
    unittest.main()
