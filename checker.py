"""Read product tables and apply explicit, reversible maintenance rules.

No store connection or model call is performed. All business values remain
traceable to the supplied table; uncertain values are reported for review.
"""

from __future__ import annotations

import csv
import io
import re
import unicodedata
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path

FIELDS = ["商品名称", "SKU编码", "规格", "售价", "库存"]
MAX_BYTES = 5 * 1024 * 1024
MAX_ROWS = 5000
MAX_COLS = 80
MAX_CELL_CHARS = 32767
MAX_SCAN_ROWS = 50000
MAX_UNCOMPRESSED_BYTES = 100 * 1024 * 1024

_ALIASES = {
    "商品名称": {"商品名称", "商品标题", "标题", "品名", "名称", "product_name", "name", "title"},
    "SKU编码": {"SKU编码", "SKU编号", "SKU", "SKU_ID", "规格编码", "商家编码", "sku_code"},
    "规格": {"规格", "商品规格", "规格组合", "SKU规格", "型号规格", "spec", "specification"},
    "售价": {"售价", "销售价", "销售价格", "单价", "价格", "price", "selling_price"},
    "库存": {"库存", "库存数量", "可售库存", "真实可售库存", "stock", "inventory", "quantity"},
}
_MISSING = {"", "未知", "待补", "待确认", "未填写", "未提供"}
_DECIMAL_PATTERN = re.compile(r"[+-]?(?:\d+(?:\.\d+)?|\.\d+)\Z", re.ASCII)
_GROUPED_PATTERN = re.compile(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?\Z", re.ASCII)
_ILLEGAL_XML_CHARACTERS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _text(value) -> str:
    return "" if value is None else str(value)


def _header_key(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", str(value)).lstrip("\ufeff").casefold()
    return re.sub(r"[\s_\-]+", "", normalized)


def guess_mapping(headers: list[str]) -> dict[str, str | None]:
    """Match only unambiguous aliases; never guess between price types."""
    result = {}
    for field in FIELDS:
        aliases = {_header_key(alias) for alias in _ALIASES[field]}
        candidates = [header for header in headers if _header_key(header) in aliases]
        result[field] = candidates[0] if len(candidates) == 1 else None
    return result


def _build_table(
    source: str, sheet: str, sheet_names: list[str], headers: list,
    rows: list[list], row_numbers: list[int] | None = None,
    formulas: list[dict] | None = None, numeric_cells: list[dict] | None = None,
    notices: list[str] | None = None,
) -> dict:
    if not headers:
        raise ValueError("没有找到表头，请将字段名放在工作表第一行。")
    headers = [_text(value) for value in headers]
    rows = [[_text(value) for value in row] for row in rows]
    if row_numbers is None:
        row_numbers = list(range(2, len(rows) + 2))
    if len(row_numbers) != len(rows):
        raise ValueError("来源行号与数据行不一致，请重新读取原始文件。")
    while headers and not headers[-1].strip() and all(len(row) < len(headers) or not row[len(headers) - 1].strip() for row in rows):
        headers.pop()
    if not headers:
        raise ValueError("第一行没有字段名，请使用商品资料模板填写表头。")
    if len(headers) > MAX_COLS:
        raise ValueError(f"表格超过{MAX_COLS}列，请拆分后再检查。")
    normalized_headers = [_header_key(header) for header in headers]
    if any(not key for key in normalized_headers):
        raise ValueError("表头中有空白列，请补齐字段名或删除无用空列。")
    if len(set(normalized_headers)) != len(normalized_headers):
        raise ValueError("存在重复或仅空格/大小写不同的表头，请先为每列设置唯一名称。")
    if any(len(header) > MAX_CELL_CHARS for header in headers):
        raise ValueError("表头文字超过Excel单元格长度限制，请缩短后重试。")
    if any(_ILLEGAL_XML_CHARACTERS.search(header) for header in headers):
        raise ValueError("表头包含无法写入Excel的控制字符，请清理后重新上传。")
    kept_rows, kept_numbers = [], []
    for number, row in zip(row_numbers, rows):
        if any(_ILLEGAL_XML_CHARACTERS.search(value) for value in row):
            raise ValueError(f"第{number}行包含无法写入Excel的控制字符，请清理后重新上传。")
        if not any(value.strip() for value in row):
            continue
        if len(row) > len(headers) and any(value.strip() for value in row[len(headers):]):
            raise ValueError(f"第{number}行的数据列多于表头，请补齐表头，避免遗漏数据。")
        if any(len(value) > MAX_CELL_CHARS for value in row):
            raise ValueError(f"第{number}行包含超过32767字符的单元格，请缩短后重试。")
        kept_rows.append((row + [""] * len(headers))[:len(headers)])
        kept_numbers.append(number)
        if len(kept_rows) > MAX_ROWS:
            raise ValueError(f"商品数据超过{MAX_ROWS}行，请分批检查。")
    if not kept_rows:
        raise ValueError("表格中没有商品数据，请在表头下面填写至少一行商品。")
    return {
        "source": source, "sheet": sheet, "sheet_names": list(sheet_names),
        "headers": headers, "rows": kept_rows, "row_numbers": kept_numbers,
        "formulas": list(formulas or []), "numeric_cells": list(numeric_cells or []),
        "notices": list(notices or []),
    }


def _load_csv(filename: str, payload: bytes) -> dict:
    text = None
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            text = payload.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("CSV编码无法识别，请另存为UTF-8 CSV后重试。")
    if "\x00" in text:
        raise ValueError("文件包含二进制内容，请使用真正的CSV或.xlsx文件。")
    try:
        try:
            delimiter = csv.Sniffer().sniff(text[:8192], delimiters=",;\t").delimiter
        except csv.Error:
            delimiter = ","
        reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter, strict=True)
        headers = next(reader, [])
        rows, numbers = [], []
        last_line = reader.line_num
        for row in reader:
            start_line = last_line + 1
            last_line = reader.line_num
            if any(value.strip() for value in row):
                rows.append(row)
                numbers.append(start_line)
                if len(rows) > MAX_ROWS:
                    raise ValueError(f"商品数据超过{MAX_ROWS}行，请分批检查。")
        return _build_table(filename, "CSV", [], headers, rows, numbers)
    except csv.Error as exc:
        raise ValueError("CSV引号或分隔符格式有误，请用Excel/WPS重新另存为CSV。") from exc


def _worksheet_rows(worksheet):
    """Stream actual cells, while bounding pathological worksheet dimensions."""
    worksheet.reset_dimensions()
    rows, numbers, formula_indexes, numeric_indexes = [], [], [], []
    headers = None
    for number, cells in enumerate(worksheet.iter_rows(), 1):
        if number > MAX_SCAN_ROWS:
            raise ValueError("工作表有效范围过大，请清除多余空白行/格式后再上传。")
        # Reject an actual far-column value while streaming this row, before
        # accumulating any later rows. Sparse sheets can otherwise expand each
        # record to thousands of cells despite a very small compressed file.
        if len(cells) > MAX_COLS:
            if any(cell.value is not None and str(cell.value).strip() for cell in cells[MAX_COLS:]):
                raise ValueError(f"第{number}行在{MAX_COLS}列之外仍有数据，请拆分或整理表格后再上传。")
            # Formatting-only/blank far cells carry no source value to retain.
            cells = cells[:MAX_COLS]
        values = [cell.value for cell in cells]
        if number == 1:
            if any(getattr(cell, "data_type", None) == "f" for cell in cells):
                raise ValueError("表头不能使用公式，请把第一行改为明确的字段名。")
            headers = values
            continue
        if not any(value is not None and str(value).strip() for value in values):
            continue
        if len(rows) >= MAX_ROWS:
            raise ValueError(f"商品数据超过{MAX_ROWS}行，请分批检查。")
        rows.append(values)
        numbers.append(number)
        for column, cell in enumerate(cells):
            if cell.value is None:
                continue
            if getattr(cell, "data_type", None) == "f":
                formula_indexes.append((number, column))
            elif getattr(cell, "data_type", None) == "n" and not isinstance(cell.value, bool):
                numeric_indexes.append((number, column))
    return headers or [], rows, numbers, formula_indexes, numeric_indexes


def _validate_payload(payload: bytes) -> None:
    if not isinstance(payload, bytes):
        raise ValueError("文件内容无效，请重新选择上传文件。")
    if not payload:
        raise ValueError("上传文件为空，请选择含有商品资料的表格。")
    if len(payload) > MAX_BYTES:
        raise ValueError("文件大小超过5MB，请拆分后上传。")


def _open_xlsx(payload: bytes):
    """Validate archive bounds before opening the workbook's metadata."""
    from openpyxl import load_workbook

    _validate_payload(payload)
    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            entries = archive.infolist()
            if len(entries) > 5000 or sum(info.file_size for info in entries) > MAX_UNCOMPRESSED_BYTES:
                raise ValueError("Excel解压后的内容过大，请拆分文件再上传。")
        return load_workbook(io.BytesIO(payload), read_only=True, data_only=False, keep_links=False)
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("无法读取Excel，请确认文件未加密且为有效的.xlsx格式。") from exc


def _visible_sheet_names(workbook) -> list[str]:
    visible = [sheet.title for sheet in workbook.worksheets if sheet.sheet_state == "visible"]
    if not visible:
        raise ValueError("Excel没有可见工作表，请先取消目标工作表的隐藏状态。")
    return visible


def list_workbook_sheets(payload: bytes) -> list[str]:
    """List visible sheets without parsing headers or product rows.

    This remains available when the first visible sheet is a cover or has bad
    headers, allowing the user to explicitly choose a usable product sheet.
    """
    workbook = _open_xlsx(payload)
    try:
        return _visible_sheet_names(workbook)
    finally:
        workbook.close()


def _load_xlsx(filename: str, payload: bytes, sheet_name: str | None) -> dict:
    workbook = _open_xlsx(payload)
    try:
        visible = _visible_sheet_names(workbook)
        if sheet_name is not None and sheet_name not in visible:
            raise ValueError("所选工作表不存在或已隐藏，请重新选择可见工作表。")
        candidates = [sheet_name] if sheet_name else visible
        for name in candidates:
            headers, rows, numbers, formula_indexes, numeric_indexes = _worksheet_rows(workbook[name])
            if not rows and sheet_name is None:
                continue
            # Metadata uses exact source headers, so manual mappings remain reliable.
            def metadata(indexes):
                return [{"row": row, "column": _text(headers[column])} for row, column in indexes if column < len(headers)]

            notices = []
            if len(visible) > 1:
                notices.append(f"本次仅检查工作表“{name}”，其余工作表可切换后分别检查。")
            return _build_table(filename, name, visible, headers, rows, numbers, metadata(formula_indexes), metadata(numeric_indexes), notices)
        raise ValueError("Excel中没有商品数据，请在表头下面填写至少一行商品。")
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("工作表数据无法完整读取，请用Excel/WPS修复并另存为.xlsx后重试。") from exc
    finally:
        workbook.close()


def load_upload(filename: str, payload: bytes, sheet_name: str | None = None) -> dict:
    """Read one explicitly bounded upload into string values and provenance."""
    _validate_payload(payload)
    filename = str(filename or "商品资料")
    extension = Path(filename).suffix.casefold()
    if extension == ".csv":
        return _load_csv(filename, payload)
    if extension == ".xlsx":
        return _load_xlsx(filename, payload, sheet_name)
    raise ValueError("仅支持.xlsx或.csv；旧版.xls请先用Excel/WPS另存为.xlsx。")


def _issue(code: str, field: str, level: str, message: str, suggestion: str) -> dict:
    return {"code": code, "field": field, "level": level, "message": message, "suggestion": suggestion}


def _number(value: str, price: bool = False) -> Decimal:
    candidate = unicodedata.normalize("NFKC", value).strip()
    if price:
        if candidate.startswith("¥"):
            candidate = candidate[1:].strip()
        if candidate.endswith("元"):
            candidate = candidate[:-1].strip()
    if "," in candidate:
        if not _GROUPED_PATTERN.fullmatch(candidate):
            raise ValueError("数字的千位分隔格式有误")
        candidate = candidate.replace(",", "")
    if len(candidate) > 40 or not _DECIMAL_PATTERN.fullmatch(candidate):
        raise ValueError("不是明确的十进制数字")
    try:
        result = Decimal(candidate)
    except InvalidOperation as exc:
        raise ValueError("不是有效数字") from exc
    if not result.is_finite():
        raise ValueError("不是有限数字")
    return result


def analyze_table(table: dict, mapping: dict | None = None) -> dict:
    """Normalize known formats and expose errors or questions without guessing."""
    validated = _build_table(
        table.get("source", "商品资料"), table.get("sheet", ""), table.get("sheet_names", []),
        table.get("headers", []), table.get("rows", []), table.get("row_numbers"),
        table.get("formulas", []), table.get("numeric_cells", []), table.get("notices", []),
    )
    headers = validated["headers"]
    proposed = guess_mapping(headers) if mapping is None else {field: mapping.get(field) or None for field in FIELDS}
    selected = [source for source in proposed.values() if source is not None]
    if any(source not in headers for source in selected):
        raise ValueError("字段对应包含不存在的原表列，请重新选择。")
    if len(selected) != len(set(selected)):
        raise ValueError("同一原表列不能同时对应多个标准字段，请重新选择。")
    unselected = [header for header in headers if header not in selected]
    columns = FIELDS + [f"原列·{header}" for header in unselected]
    file_issues = [
        _issue("missing_mapping", field, "错误", f"未找到或未明确对应“{field}”列。", "请在字段对应中选择准确的原表列；原表缺列时补齐后重新上传。")
        for field in FIELDS if proposed[field] is None
    ]
    formula_cells = {(item["row"], item["column"]) for item in validated["formulas"]}
    numeric_cells = {(item["row"], item["column"]) for item in validated["numeric_cells"]}
    output_rows = []
    valid_prices = {}
    by_sku = defaultdict(list)
    for row_number, values in zip(validated["row_numbers"], validated["rows"]):
        raw = dict(zip(headers, values))
        data = {field: raw.get(proposed[field], "").strip() for field in FIELDS}
        data.update({f"原列·{header}": raw[header].strip() for header in unselected})
        issues = []
        for field in FIELDS:
            value = data[field]
            source = proposed[field]
            if value in _MISSING:
                issues.append(_issue("missing_value", field, "错误", f"{field}未填写或仍为待确认内容。", "根据真实商品资料补齐，不要根据其他商品猜测。"))
                continue
            if (row_number, source) in formula_cells:
                issues.append(_issue("formula", field, "错误", "该字段来自Excel公式，尚未核实为确定值。", "核对公式结果后将实际值粘贴为数值/文本，再重新检查。"))
                continue
            if field in ("售价", "库存"):
                try:
                    number = _number(value, price=field == "售价")
                    if field == "售价":
                        # The accepted input can exceed Decimal's default 28
                        # digits. Validate at sufficient precision so rounding
                        # cannot conceal a third fractional digit.
                        with localcontext() as context:
                            context.prec = 80
                            if number <= 0 or number * 100 != (number * 100).to_integral_value():
                                raise ValueError("售价必须大于0，且最多保留两位有效小数")
                        data[field] = format(number, ".2f")
                        valid_prices[row_number] = number
                    else:
                        if number < 0 or number != number.to_integral_value():
                            raise ValueError("库存必须为非负整数，0库存允许")
                        data[field] = str(int(number))
                except (ValueError, InvalidOperation):
                    instruction = "填入大于0、最多两位小数的实际售价，并确认价格口径。" if field == "售价" else "填入大于或等于0的整数，核对真实可售库存。"
                    message = "售价不是有效的正数金额，或超过两位有效小数。" if field == "售价" else "库存不是有效的非负整数。"
                    issues.append(_issue("invalid_price" if field == "售价" else "invalid_stock", field, "错误", message, instruction))
            if field == "SKU编码" and (row_number, source) in numeric_cells:
                issues.append(_issue("numeric_sku", field, "待核对", "原Excel中的SKU为数字类型，前导零或长编码可能已经丢失。", "对照原始SKU核实；将源单元格设为文本，无法从数字自动恢复丢失位数。"))
        for header in unselected:
            if (row_number, header) in formula_cells:
                issues.append(_issue("formula", f"原列·{header}", "待核对", "原始附加列含公式，仅保留表达式。", "如该列参与业务判断，请核实计算结果并粘贴确定值。"))
        changes = []
        for field in columns:
            source = proposed.get(field) if field in FIELDS else field.removeprefix("原列·")
            before = raw.get(source, "")
            if data[field] != before:
                changes.append({"field": field, "before": before, "after": data[field]})
        record = {"row": row_number, "raw": raw, "data": data, "issues": issues, "changes": changes, "status": "通过"}
        output_rows.append(record)
        if data["SKU编码"] not in _MISSING and (row_number, proposed["SKU编码"]) not in formula_cells:
            by_sku[data["SKU编码"]].append(record)
    for sku, records in by_sku.items():
        if len(records) < 2:
            continue
        row_labels = "、".join(str(record["row"]) for record in records[:12])
        if len(records) > 12:
            row_labels += f"等{len(records)}行"
        different_prices = {valid_prices[record["row"]] for record in records if record["row"] in valid_prices}
        for record in records:
            record["issues"].append(_issue("duplicate_sku", "SKU编码", "待核对", f"同一SKU出现在第{row_labels}行，全部原记录已保留。", "确认是否重复录入，或是否混入不同店铺/时间的记录，再人工处理。"))
            if len(different_prices) > 1:
                record["issues"].append(_issue("price_conflict", "售价", "待核对", "同一SKU存在不同售价，需核对价格口径。", "确认记录时间、活动及价格类型；工具不选择或覆盖其中任一价格。"))
    for row in output_rows:
        row["status"] = "需修正" if any(issue["level"] == "错误" for issue in row["issues"]) else ("待核对" if row["issues"] else "通过")
    summary = {
        "total_rows": len(output_rows),
        "passed_rows": sum(row["status"] == "通过" for row in output_rows),
        "issue_rows": sum(row["status"] != "通过" for row in output_rows),
        "error_rows": sum(row["status"] == "需修正" for row in output_rows),
        "review_rows": sum(row["status"] == "待核对" for row in output_rows),
        "issue_count": len(file_issues) + sum(len(row["issues"]) for row in output_rows),
        "normalized_cells": sum(len(row["changes"]) for row in output_rows),
    }
    return {
        "source": validated["source"], "sheet": validated["sheet"],
        "generated_at": datetime.now(timezone(timedelta(hours=8))).isoformat(timespec="seconds"),
        "mapping": proposed, "columns": columns, "rows": output_rows,
        "file_issues": file_issues, "notices": validated["notices"], "summary": summary,
    }


def demo_table() -> dict:
    """A small, explicitly fictional apparel table with purposeful issues."""
    headers = ["商品标题", "SKU编码", "规格组合", "售价", "库存", "品牌", "材质"]
    rows = [
        [" 棉质圆领短袖 ", "NL2401-CRM-S", "米白/S", "￥89.00", "18", "示例衣橱（虚构）", "棉"],
        ["棉质圆领短袖", " NL2401-CRM-M ", "米白/M", "89", "12", "示例衣橱（虚构）", "棉"],
        ["棉质圆领短袖", "NL2401-BLK-S", "黑色/S", "89", "-3", "示例衣橱（虚构）", "棉"],
        ["水洗牛仔衬衫", "NL2402-BLU-M", "浅蓝/M", "129", "7", "示例衣橱（虚构）", "棉"],
        ["水洗牛仔衬衫", "NL2402-BLU-M", "浅蓝/M", "139", "7", "示例衣橱（虚构）", "棉"],
        ["水洗牛仔衬衫", "NL2402-BLU-L", "", "129", "9", "示例衣橱（虚构）", "棉"],
        ["", "NL2403-GRY-M", "灰色/M", "76", "11", "示例衣橱（虚构）", "棉/聚酯纤维"],
        ["休闲短裤", "NL2403-GRY-L", "灰色/L", "待确认", "6", "示例衣橱（虚构）", "棉/聚酯纤维"],
        ["帆布购物袋", "NL2404-WHT-ONE", "米白/均码", "39.9", "十件", "示例衣橱（虚构）", "棉"],
        ["基础棉袜", "000087", "米白/均码", "19.9", "0", "示例衣橱（虚构）", "棉/氨纶"],
    ]
    return _build_table("演示商品资料（虚构）", "演示资料", ["演示资料"], headers, rows, notices=["本表为虚构演示资料，含故意设置的错误，不代表真实商品或经营结果。"])
