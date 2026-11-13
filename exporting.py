"""应用运行时的商品资料导出，不写入或覆盖任何源文件。"""

from __future__ import annotations

import io
import re
from decimal import Decimal, InvalidOperation
from html import escape

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


FIELDS = ["商品名称", "SKU编码", "规格", "售价", "库存"]
GREEN = "28634C"
PALE_GREEN = "E9F1EA"
PALE_RED = "FCE8E6"
PALE_AMBER = "FFF1D6"
INK = "20342A"
DECIMAL_NUMBER = re.compile(r"^\d+(?:\.\d{1,2})?$")


def _text(value) -> str:
    return "" if value is None else str(value)


def _excel_literal(value: str) -> str:
    # OOXML cannot hold these characters. A visible escape retains their identity.
    return ILLEGAL_CHARACTERS_RE.sub(lambda match: f"\\u{ord(match.group()):04x}", value)


def _put(cell, value, *, text=False):
    if text or isinstance(value, str):
        cell.value = _excel_literal(_text(value))
        # Assigning '=...' normally creates a formula, so explicitly retain a string.
        cell.data_type = "s"
        cell.number_format = "@"
    else:
        cell.value = value
    cell.alignment = Alignment(vertical="top", wrap_text=True)


def _new_sheet(book, name, headers, rows):
    sheet = book.create_sheet(name)
    for column, header in enumerate(headers, 1):
        cell = sheet.cell(1, column)
        _put(cell, header, text=True)
        cell.fill = PatternFill("solid", fgColor=GREEN)
        cell.font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    for row_number, values in enumerate(rows, 2):
        for column, value in enumerate(values, 1):
            cell = sheet.cell(row_number, column)
            _put(cell, value)
            cell.font = Font(name="微软雅黑", size=10, color=INK)
            if row_number % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F3F6F1")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(max(1, len(headers)))}{max(1, sheet.max_row)}"
    sheet.row_dimensions[1].height = 28
    sheet.sheet_view.showGridLines = False
    for column, header in enumerate(headers, 1):
        width = 16
        if any(word in str(header) for word in ["名称", "标题", "原因", "问题", "建议", "说明", "原值", "修改", "内容", "项目"]):
            width = 34
        if header in ["原表行号", "级别", "检查状态", "库存"]:
            width = 13
        if header == "SKU编码":
            width = 23
        sheet.column_dimensions[get_column_letter(column)].width = width
    return sheet


def _bytes(book) -> bytes:
    output = io.BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def _empty_book():
    book = Workbook()
    book.remove(book.active)
    book.properties.creator = "商品资料检查"
    book.properties.title = "商品资料整理与检查"
    book.calculation.fullCalcOnLoad = False
    return book


def _raw_headers(report):
    return list(dict.fromkeys(key for row in report["rows"] for key in row["raw"]))


def _issues(report):
    for issue in report.get("file_issues", []):
        yield ["整表", "", issue.get("field", ""), issue.get("level", "错误"),
               issue.get("message", ""), issue.get("suggestion", "")]
    for row in report["rows"]:
        for issue in row.get("issues", []):
            yield [row["row"], row["data"].get("SKU编码", ""), issue.get("field", ""),
                   issue.get("level", "错误"), issue.get("message", ""), issue.get("suggestion", "")]


def _changes(report):
    for row in report["rows"]:
        for change in row.get("changes", []):
            yield [row["row"], row["data"].get("SKU编码", ""), change.get("field", ""),
                   change.get("before", ""), change.get("after", "")]


def _issue_text(row):
    return "；".join(f"{item.get('field', '')}：{item.get('message', '')}" for item in row.get("issues", []))


def _numeric_value(row, field):
    value = row["data"].get(field, "")
    if any(item.get("field") == field and item.get("level") == "错误" for item in row.get("issues", [])):
        return value
    if field not in {"售价", "库存"} or not DECIMAL_NUMBER.fullmatch(_text(value)):
        return value
    # Excel only guarantees fifteen significant digits for a numeric cell.
    # Preserve unusually large values as text rather than round their source.
    if len(_text(value).replace(".", "").lstrip("0")) > 15:
        return value
    try:
        number = Decimal(str(value))
    except InvalidOperation:
        return value
    if field == "库存":
        return int(number) if number >= 0 and number == number.to_integral_value() else value
    return number if number > 0 else value


def export_xlsx(report: dict) -> bytes:
    """Return a styled five-sheet review workbook, preserving source literals."""
    book = _empty_book()
    columns = list(report["columns"])
    normalized_rows = [
        [row["row"], row["status"]] + [_numeric_value(row, field) for field in columns] + [_issue_text(row)]
        for row in report["rows"]
    ]
    normalized = _new_sheet(book, "整理结果", ["原表行号", "检查状态"] + columns + ["问题说明"], normalized_rows)
    column_indexes = {field: index + 3 for index, field in enumerate(columns)}
    for excel_row, row in enumerate(report["rows"], 2):
        status_cell = normalized.cell(excel_row, 2)
        status_cell.fill = PatternFill("solid", fgColor={"通过": PALE_GREEN, "需修正": PALE_RED, "待核对": PALE_AMBER}.get(row["status"], PALE_AMBER))
        status_cell.font = Font(name="微软雅黑", size=10, bold=True, color=INK)
        for issue in row.get("issues", []):
            if issue.get("field") in column_indexes:
                cell = normalized.cell(excel_row, column_indexes[issue["field"]])
                color = PALE_RED if issue.get("level") == "错误" else PALE_AMBER
                # An error highlight takes precedence over a review highlight.
                if cell.fill.fgColor.rgb != f"00{PALE_RED}":
                    cell.fill = PatternFill("solid", fgColor=color)
        for field in ["售价", "库存"]:
            if field in column_indexes:
                cell = normalized.cell(excel_row, column_indexes[field])
                if cell.data_type == "n":
                    cell.number_format = "0.00" if field == "售价" else "0"

    issue_rows = list(_issues(report))
    issue_sheet = _new_sheet(book, "问题清单", ["原表行号", "SKU编码", "字段", "级别", "问题说明", "处理建议"], issue_rows)
    for row_number, row in enumerate(issue_rows, 2):
        issue_sheet.cell(row_number, 4).fill = PatternFill("solid", fgColor=PALE_RED if row[3] == "错误" else PALE_AMBER)
    _new_sheet(book, "修改记录", ["原表行号", "SKU编码", "字段", "原值", "整理后"], list(_changes(report)))
    raw_headers = _raw_headers(report)
    _new_sheet(book, "原始资料", ["原表行号"] + raw_headers,
               [[row["row"]] + [row["raw"].get(header, "") for header in raw_headers] for row in report["rows"]])
    summary = report["summary"]
    explanations = [
        ["源文件", report.get("source", "")],
        ["工作表", report.get("sheet", "")],
        ["生成时间（北京时间）", report.get("generated_at", "")],
        ["商品行数", summary["total_rows"]],
        ["通过行数", summary["passed_rows"]],
        ["需处理行数", summary["issue_rows"]],
        ["其中需修正", summary["error_rows"]],
        ["其中待核对", summary["review_rows"]],
        ["问题条目", summary["issue_count"]],
        ["格式整理单元格", summary["normalized_cells"]],
        ["数据口径", "一行一个SKU，同一店铺、同一时间和价格口径。供货价、单买价、拼单价分别保留。"],
        ["价格与库存", "售价必须大于零且最多两位小数；库存必须为非负整数，零库存允许。"],
        ["重复与冲突", "所有重复记录均保留。重复SKU及价格冲突交由人工核对，不自动选取价格或删除记录。"],
        ["源值与公式", "源字符串按文本保存，公式不会执行；SKU以文本保存。Excel不能存储的控制字符以\\uXXXX显示。"],
        ["数值精度", "超过Excel数值精度的金额或库存保留为文本，避免导出时产生舍入。"],
        ["应用范围", "本报告仅做通用资料检查，不代表平台类目、资质或发布审核已通过。"],
        ["后续操作", "核对问题清单，按来源补齐或修正资料，再上传检查。源文件未被覆盖。"],
    ]
    explanations += [["字段对应", f"{field} ← {source or '未选择'}"] for field, source in report.get("mapping", {}).items()]
    explanations += [["来源提示", notice] for notice in report.get("notices", [])]
    _new_sheet(book, "说明", ["项目", "内容"], explanations)
    book["说明"].column_dimensions["B"].width = 84
    return _bytes(book)


def _html_table(headers, rows):
    header = "".join(f"<th scope=\"col\">{escape(_text(value), quote=True)}</th>" for value in headers)
    body = "".join("<tr>" + "".join(f"<td>{escape(_text(value), quote=True)}</td>" for value in row) + "</tr>" for row in rows)
    if not body:
        body = f'<tr><td colspan="{len(headers)}" class="empty">暂无记录</td></tr>'
    return f'<div class="table-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>'


def export_html(report: dict) -> bytes:
    """Return an escaped, self-contained HTML report suitable for offline sharing."""
    summary = report["summary"]
    columns = report["columns"]
    stats = "".join(f'<div><span>{label}</span><strong>{int(summary[key])}</strong></div>' for label, key in [
        ("商品行数", "total_rows"), ("通过", "passed_rows"), ("需处理", "issue_rows"),
        ("问题条目", "issue_count"), ("格式整理单元格", "normalized_cells")])
    normalized = _html_table(["原表行号", "检查状态"] + columns + ["问题说明"],
                             [[row["row"], row["status"]] + [row["data"].get(column, "") for column in columns] + [_issue_text(row)] for row in report["rows"]])
    problems = _html_table(["原表行号", "SKU编码", "字段", "级别", "问题说明", "处理建议"], list(_issues(report)))
    modifications = _html_table(["原表行号", "SKU编码", "字段", "原值", "整理后"], list(_changes(report)))
    raw_headers = _raw_headers(report)
    raw = _html_table(["原表行号"] + raw_headers,
                      [[row["row"]] + [row["raw"].get(header, "") for header in raw_headers] for row in report["rows"]])
    mapping = "；".join(f"{field} ← {source or '未选择'}" for field, source in report.get("mapping", {}).items())
    notices = "".join(f"<li>{escape(_text(notice), quote=True)}</li>" for notice in report.get("notices", []))
    html = f'''<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'">
<title>商品资料检查报告</title><style>
:root{{color-scheme:light}}*{{box-sizing:border-box}}body{{margin:0;background:#f6f7f3;color:#20342a;font:14px/1.7 "Microsoft YaHei",sans-serif}}
main{{max-width:1250px;margin:36px auto;padding:0 24px 40px}}header{{border-bottom:3px solid #28634c;padding-bottom:18px}}h1{{font-size:28px;line-height:1.3;margin:8px 0}}h2{{font-size:19px;margin:30px 0 12px}}
.eyebrow{{color:#28634c;font-weight:700;letter-spacing:.12em;font-size:12px}}.meta,.note{{color:#526258;overflow-wrap:anywhere}}.stats{{display:flex;flex-wrap:wrap;gap:24px;padding:22px 0;border-bottom:1px solid #d5dfd5}}
.stats div{{min-width:100px}}.stats span{{display:block;color:#526258;font-size:12px}}.stats strong{{font:600 30px/1.4 Georgia,serif;font-variant-numeric:tabular-nums}}
.table-wrap{{overflow-x:auto;background:#fff;border:1px solid #d5dfd5;border-radius:6px}}table{{width:100%;border-collapse:collapse;min-width:650px}}th,td{{text-align:left;padding:10px 12px;border-bottom:1px solid #e1e8df;vertical-align:top;overflow-wrap:anywhere;max-width:320px}}th{{background:#e9eee6;font-weight:600;white-space:nowrap}}tr:last-child td{{border-bottom:0}}.empty{{color:#68736c}}details{{margin-top:26px}}summary{{font-size:18px;font-weight:600;cursor:pointer}}footer{{margin-top:30px;padding-top:18px;border-top:1px solid #d5dfd5;font-size:12px;color:#526258}}@media(max-width:500px){{main{{padding:0 14px 30px;margin-top:20px}}h1{{font-size:24px}}.stats{{gap:14px}}.stats div{{min-width:90px}}}}@media print{{body{{background:#fff}}main{{max-width:none;margin:0;padding:0}}.table-wrap{{overflow:visible}}table{{min-width:0;font-size:10px}}th,td{{padding:5px}}details{{display:block}}thead{{display:table-header-group}}}}
</style></head><body><main><header><div class="eyebrow">商品基础资料 · 自动整理与检查</div><h1>商品资料检查报告</h1>
<div class="meta">源文件：{escape(_text(report.get("source")), quote=True)}<br>工作表：{escape(_text(report.get("sheet")), quote=True)}<br>生成时间（北京时间）：{escape(_text(report.get("generated_at")), quote=True)}</div></header>
<section class="stats" aria-label="检查统计">{stats}</section>
<p class="note">所有原始记录均保留。需修正 {int(summary['error_rows'])} 行，待核对 {int(summary['review_rows'])} 行；“通过”仅表示本次通用字段规则检查未发现问题。</p>
{f'<ul class="note">{notices}</ul>' if notices else ''}
<section><h2>问题清单</h2>{problems}</section><section><h2>整理结果</h2>{normalized}</section>
<section><h2>修改记录</h2>{modifications}</section><details><summary>原始资料</summary>{raw}</details>
<footer><p>字段对应：{escape(mapping, quote=True)}</p><p>请按资料来源核对问题，补齐或修正后重新检查。本报告不代表平台类目、资质或发布审核已通过。</p></footer>
</main></body></html>'''
    return html.encode("utf-8")


def template_xlsx() -> bytes:
    """Return an empty five-field template plus concise filling instructions."""
    book = _empty_book()
    sheet = _new_sheet(book, "商品资料", FIELDS, [])
    sheet.column_dimensions["B"].number_format = "@"
    instructions = [
        ["使用范围", "同一店铺、同一时间、同一价格口径；一行一个SKU。先填写商品资料页，再上传检查。"],
        ["商品名称", "必填。填写能够核对的实际商品名称。"],
        ["SKU编码", "必填。请先将SKU列设为文本，再输入00123等编码，避免前导零丢失。"],
        ["规格", "必填。例如米白/M、500ml/单瓶，单位写清楚。"],
        ["售价", "必填。价格为大于零的人民币金额，最多两位小数。供货价、单买价、拼单价含义不同，需要明确价格口径。"],
        ["库存", "必填。非负整数，零库存允许。不要将未知库存填成零。"],
        ["其他属性", "可以增加品牌、材质、货号等列，整理时会保留原列。"],
        ["表格格式", "第一行为唯一表头，不合并单元格；最多5000行、80列、5MB。不要使用公式代替已核对的数据。"],
        ["检查后操作", "查看问题清单，按来源补齐或修正，再次上传核对。源文件不会被覆盖。"],
    ]
    _new_sheet(book, "填写说明", ["字段/项目", "填写要求"], instructions)
    book["填写说明"].column_dimensions["B"].width = 84
    return _bytes(book)


def table_xlsx(table: dict) -> bytes:
    """Return source-like sample data, without normalization or formula execution."""
    book = _empty_book()
    sheet_name = re.sub(r"[\\/*?:\[\]]", "_", _text(table.get("sheet") or "商品资料"))[:31] or "商品资料"
    sheet = _new_sheet(book, sheet_name, table["headers"], table["rows"])
    # All source values stay strings, including values that look numeric or formulas.
    for row_number, values in enumerate(table["rows"], 2):
        for column, value in enumerate(values, 1):
            _put(sheet.cell(row_number, column), value, text=True)
    return _bytes(book)
