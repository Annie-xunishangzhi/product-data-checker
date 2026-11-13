"""Streamlit真实页面状态测试；文件上传与下载另由浏览器验收。"""
import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parents[1] / "app.py"


class AppTests(unittest.TestCase):
    def setUp(self):
        self.assertTrue(APP.is_file(), "商品资料操作页面尚未实现")

    def create_app(self):
        app = AppTest.from_file(str(APP), default_timeout=20).run()
        self.assertEqual(len(app.exception), 0)
        return app

    def checked_demo(self):
        app = self.create_app()
        app.button(key="load_demo").click().run()
        self.assertEqual(len(app.exception), 0)
        app.button(key="run_check").click().run()
        self.assertEqual(len(app.exception), 0)
        return app

    def test_initial_view_has_demo_and_no_old_results(self):
        app = self.create_app()
        self.assertEqual(app.button(key="load_demo").label, "载入演示数据")
        self.assertEqual(len(app.metric), 0)

    def test_demo_can_be_checked_and_counts_match_actual_report(self):
        app = self.checked_demo()
        report = app.session_state["report"]
        summary = report["summary"]
        self.assertEqual(summary["total_rows"], 10)
        self.assertGreater(summary["issue_rows"], 0)
        self.assertEqual(summary["passed_rows"] + summary["issue_rows"], 10)
        self.assertEqual(len(app.metric), 4)
        self.assertEqual(int(app.metric[0].value), 10)
        self.assertEqual(len(app.get("download_button")), 4)

    def test_mapping_change_hides_stale_report_and_exports(self):
        app = self.checked_demo()
        sku_mapping = next(box for box in app.selectbox if box.label == "SKU编码对应哪一列")
        sku_mapping.set_value("（未对应）").run()
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.metric), 0)
        self.assertTrue(any("重新检查" in item.value for item in app.info))

    def test_clear_returns_to_empty_state(self):
        app = self.checked_demo()
        app.button(key="clear_input").click().run()
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.metric), 0)
        self.assertNotIn("report", app.session_state)


if __name__ == "__main__":
    unittest.main()
