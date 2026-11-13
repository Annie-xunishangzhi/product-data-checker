"""对正在运行的本地工具做浏览器验收，使用系统Python的Playwright。"""
import argparse
import io
import json
from pathlib import Path
import sys

from openpyxl import Workbook, load_workbook
from playwright.sync_api import expect, sync_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "artifacts"


def download(page, name, target):
    with page.expect_download() as pending:
        page.get_by_role("button", name=name, exact=True).click()
    pending.value.save_as(target)
    assert target.stat().st_size > 0
    return target


def check(page):
    page.get_by_role("button", name="开始整理与检查", exact=True).click()
    page.get_by_role("button", name="下载整理结果 Excel", exact=True).wait_for()
    expect(page.locator('[data-testid="stException"]')).to_have_count(0)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default="http://127.0.0.1:8501")
    args = parser.parse_args()
    OUT.mkdir(exist_ok=True)
    results = {}
    errors = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(channel="chrome", headless=True)
        context = browser.new_context(viewport={"width": 1440, "height": 1060}, accept_downloads=True)
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.set_default_timeout(15000)
        page.goto(args.url, wait_until="networkidle")
        page.get_by_role("button", name="载入演示数据", exact=True).wait_for()
        expect(page.locator('[data-testid="stException"]')).to_have_count(0)
        page.screenshot(path=str(OUT / "01_初始页面.png"))

        template_path = download(page, "下载空白Excel模板", OUT / "验收_空白模板.xlsx")
        template = load_workbook(template_path)
        assert template.sheetnames == ["商品资料", "填写说明"]
        assert list(template["商品资料"].values)[0] == ("商品名称", "SKU编码", "规格", "售价", "库存")
        template.close()
        demo_path = download(page, "下载演示Excel", OUT / "验收_演示资料.xlsx")
        page.get_by_role("button", name="载入演示数据", exact=True).click()
        check(page)
        metrics = page.locator('[data-testid="stMetricValue"]').all_text_contents()
        assert metrics == ["10", "3", "5", "2"], metrics
        results["demo_counts"] = metrics
        print("PASS: 演示资料及统计", flush=True)
        page.get_by_text("显示 10 / 10 行。需要补充的信息可在导出的表格中修改，再上传复查。", exact=True).scroll_into_view_if_needed()
        page.screenshot(path=str(OUT / "02_演示检查结果.png"))
        status_box = page.get_by_role("combobox", name="按状态筛选", exact=True)
        status_box.locator("..").get_by_role("button", name="Open", exact=True).click()
        page.get_by_role("option", name="待核对", exact=True).click()
        page.get_by_text("显示 2 / 10 行。需要补充的信息可在导出的表格中修改，再上传复查。", exact=True).wait_for()
        status_box.locator("..").get_by_role("button", name="Open", exact=True).click()
        page.get_by_role("option", name="全部记录", exact=True).click()
        page.get_by_role("textbox", name="搜索商品名称或SKU", exact=True).fill("000087")
        page.get_by_role("textbox", name="搜索商品名称或SKU", exact=True).press("Enter")
        page.get_by_text("显示 1 / 10 行。需要补充的信息可在导出的表格中修改，再上传复查。", exact=True).wait_for()
        page.get_by_role("textbox", name="搜索商品名称或SKU", exact=True).fill("")
        page.get_by_role("textbox", name="搜索商品名称或SKU", exact=True).press("Enter")
        page.get_by_text("显示 10 / 10 行。需要补充的信息可在导出的表格中修改，再上传复查。", exact=True).wait_for()
        page.get_by_role("tab", name="问题清单", exact=True).click()
        page.get_by_role("radiogroup", name="问题级别", exact=True).get_by_text("待核对", exact=True).click()
        expect(page.get_by_role("radio", name="待核对", exact=True)).to_be_checked()
        page.get_by_role("tab", name="修改记录", exact=True).click()
        page.get_by_role("tab", name="原始资料", exact=True).click()
        page.get_by_role("tab", name="整理后的资料", exact=True).click()
        results["result_filters_and_tabs"] = "通过"
        print("PASS: 结果筛选、搜索和报告页签", flush=True)
        export_path = download(page, "下载整理结果 Excel", OUT / "验收_整理结果.xlsx")
        html_path = download(page, "下载检查报告 HTML", OUT / "验收_检查报告.html")
        report_book = load_workbook(export_path, data_only=False)
        assert report_book.sheetnames == ["整理结果", "问题清单", "修改记录", "原始资料", "说明"]
        normalized = list(report_book["整理结果"].values)
        assert len(normalized) == 11
        sku_index = normalized[0].index("SKU编码")
        assert normalized[-1][sku_index] == "000087"
        assert "原始资料" in html_path.read_text(encoding="utf-8")
        report_book.close()
        results["downloads"] = "Excel五页、HTML、模板及示例已下载并重读"
        print("PASS: 四种下载及Excel重读", flush=True)

        # 用户确实需要将导出表修正后重新上传。该路径不能因来源列重名而报错。
        page.locator('input[type="file"]').set_input_files(export_path)
        page.get_by_role("button", name="开始整理与检查", exact=True).wait_for()
        expect(page.locator('[data-testid="stException"]')).to_have_count(0)
        expect(page.get_by_role("button", name="下载整理结果 Excel", exact=True)).to_have_count(0)
        check(page)
        assert page.locator('[data-testid="stMetricValue"]').all_text_contents() == ["10", "3", "5", "2"]
        results["export_reupload"] = "通过"
        print("PASS: 整理结果重新上传复查", flush=True)

        page.get_by_text("核对列名对应与检查规则", exact=True).wait_for()
        sku_box = page.locator('[data-testid="stSelectbox"]').filter(has=page.get_by_text("SKU编码对应哪一列", exact=True)).get_by_role("combobox")
        sku_box.scroll_into_view_if_needed()
        sku_box.locator("..").get_by_role("button", name="Open", exact=True).click()
        page.get_by_role("option", name="（未对应）", exact=True).click()
        page.get_by_text("资料或列名对应已变化，请重新检查后查看和导出结果。", exact=True).wait_for()
        expect(page.get_by_role("button", name="下载整理结果 Excel", exact=True)).to_have_count(0)
        expect(page.locator('[data-testid="stMetricValue"]')).to_have_count(0)
        results["stale_report_hidden"] = "通过"
        print("PASS: 映射变更隐藏旧结果", flush=True)

        # 首表是有问题的封面时，也必须先显示工作表选择，让用户选择商品页。
        cover_book = Workbook()
        cover_book.active.title = "说明封面"
        cover_book.active["A2"] = "填写说明，非商品资料"
        products = cover_book.create_sheet("商品资料")
        products.append(["商品名称", "SKU编码", "规格", "售价", "库存"])
        products.append(["棉质长袖", "001000", "奶白/M", "59.9", "0"])
        cover_buffer = io.BytesIO()
        cover_book.save(cover_buffer)
        cover_book.close()
        page.locator('input[type="file"]').set_input_files({"name": "带封面的商品资料.xlsx", "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "buffer": cover_buffer.getvalue()})
        sheet_box = page.get_by_role("combobox", name="选择工作表", exact=True)
        sheet_box.wait_for()
        sheet_box.locator("..").get_by_role("button", name="Open", exact=True).click()
        page.get_by_role("option", name="商品资料", exact=True).click()
        page.get_by_role("button", name="开始整理与检查", exact=True).wait_for()
        check(page)
        assert page.locator('[data-testid="stMetricValue"]').all_text_contents() == ["1", "1", "0", "0"]
        results["invalid_cover_sheet_selection"] = "通过"
        print("PASS: 坏封面后的工作表选择", flush=True)

        page.locator('input[type="file"]').set_input_files({"name": "损坏文件.xlsx", "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "buffer": b"not an Excel file"})
        page.get_by_text("无法读取Excel，请确认文件未加密且为有效的.xlsx格式。", exact=True).wait_for()
        expect(page.locator('[data-testid="stException"]')).to_have_count(0)
        csv_data = "商品名称,SKU编码,规格,售价,库存,备注\n棉质长袖,000123,奶白/M,59.9,0,保留备注\n短袜,SOCK-02,灰/均码,19,8,普通商品\n"
        page.locator('input[type="file"]').set_input_files({"name": "实际上传测试.csv", "mimeType": "text/csv", "buffer": csv_data.encode("utf-8-sig")})
        page.get_by_role("button", name="开始整理与检查", exact=True).wait_for()
        check(page)
        assert page.locator('[data-testid="stMetricValue"]').all_text_contents() == ["2", "2", "0", "0"]
        results["bad_file_recovery_csv_zero_stock"] = "通过"
        print("PASS: 损坏文件恢复、CSV与零库存", flush=True)

        csv_export = download(page, "下载整理结果 Excel", OUT / "验收_CSV整理结果.xlsx")
        csv_book = load_workbook(csv_export)
        csv_rows = list(csv_book["整理结果"].values)
        assert csv_rows[1][csv_rows[0].index("SKU编码")] == "000123"
        csv_book.close()

        page.set_viewport_size({"width": 375, "height": 900})
        page.get_by_role("heading", name="商品资料整理与检查", exact=True).scroll_into_view_if_needed()
        page.screenshot(path=str(OUT / "03_手机宽度.png"))
        overflow = page.evaluate("({page:document.documentElement.scrollWidth,viewport:innerWidth})")
        assert overflow["page"] <= overflow["viewport"] + 1, overflow
        results["mobile_375px"] = overflow

        page.get_by_role("button", name="清空资料", exact=True).click()
        page.get_by_text("准备好一张商品表，就能开始。", exact=True).wait_for()
        expect(page.locator('[data-testid="stMetricValue"]')).to_have_count(0)
        expect(page.locator('[data-testid="stException"]')).to_have_count(0)
        results["clear"] = "通过"
        assert not errors, errors
        results["javascript_errors"] = errors
        browser.close()
    (OUT / "浏览器验收.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
