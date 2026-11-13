"""本地商品资料工作台。用户文件仅保存在当前会话内存中。"""
import hashlib
import json
import time

import streamlit as st

from checker import FIELDS, analyze_table, demo_table, guess_mapping, list_workbook_sheets, load_upload
from exporting import export_html, export_xlsx, table_xlsx, template_xlsx
from ui import install_style, render_report, source_frame


def reset_input():
    st.session_state["upload_epoch"] = st.session_state.get("upload_epoch", 0) + 1
    st.session_state["source_mode"] = "none"
    for key in ("report", "report_signature", "report_files", "elapsed_ms", "loaded_id", "loaded_table", "sheet_names_id", "sheet_names"):
        st.session_state.pop(key, None)


def select_demo():
    st.session_state["source_mode"] = "demo"
    for key in ("report", "report_signature", "report_files"):
        st.session_state.pop(key, None)


def select_upload():
    st.session_state["source_mode"] = "upload"
    for key in ("report", "report_signature", "report_files", "loaded_id", "loaded_table", "sheet_names_id", "sheet_names"):
        st.session_state.pop(key, None)


@st.cache_data(show_spinner=False)
def download_template():
    return template_xlsx()


@st.cache_data(show_spinner=False)
def download_demo():
    return table_xlsx(demo_table())


def current_table(upload):
    if st.session_state.get("source_mode") == "demo":
        return demo_table(), "demo-v1"
    if upload is None:
        return None, ""
    data = upload.getvalue()
    token = hashlib.sha256(upload.name.encode("utf-8") + data).hexdigest()[:16]
    sheet_name = None
    if upload.name.lower().endswith(".xlsx"):
        if st.session_state.get("sheet_names_id") != token:
            st.session_state["sheet_names"] = list_workbook_sheets(data)
            st.session_state["sheet_names_id"] = token
        sheet_name = st.selectbox("选择工作表", st.session_state["sheet_names"], key=f"sheet_{token}")
    selected_id = f"{token}:{sheet_name}"
    if st.session_state.get("loaded_id") != selected_id:
        st.session_state["loaded_table"] = load_upload(upload.name, data, sheet_name=sheet_name)
        st.session_state["loaded_id"] = selected_id
    table = st.session_state["loaded_table"]
    return table, f"{token}-{table['sheet']}"


def mapping_controls(table, source_id):
    guesses = guess_mapping(table["headers"])
    token = hashlib.sha256(source_id.encode()).hexdigest()[:12]
    with st.container(border=True):
        st.markdown("**核对列名对应与检查规则**")
        st.caption("五个标准字段各选一列。售价请选同一价格口径，供货价、单买价和拼单价分别保留。")
        columns = st.columns(len(FIELDS))
        mapping = {}
        missing_label = "（未对应）"
        while missing_label in table["headers"]:
            missing_label += "·"
        choices = [missing_label] + table["headers"]
        for index, field in enumerate(FIELDS):
            with columns[index]:
                choice = st.selectbox(
                    f"{field}对应哪一列", choices,
                    index=choices.index(guesses.get(field) or missing_label),
                    key=f"map_{token}_{field}",
                )
                mapping[field] = None if choice == missing_label else choice
        st.markdown(
            "**检查规则：** 五项必填；售价为正数且最多两位小数；库存为非负整数。"
            "同一SKU重复或售价不一致标为待核对，保留全部记录。"
        )
        st.caption("原表应来自同一店铺、同一时间及价格口径。数字型SKU和Excel公式会提示核对。")
    return mapping


def main():
    st.set_page_config(page_title="商品资料工作台", page_icon="📋", layout="wide", initial_sidebar_state="collapsed")
    install_style()
    st.markdown('<div class="workspace-label">运营自动化 <span>/</span> 商品基础资料 <span class="local-label">本机运行</span></div>', unsafe_allow_html=True)
    st.title("商品资料整理与检查")
    st.caption("统一商品表的格式，找出缺项、重复和异常，留下可核对的处理记录。")

    with st.container(border=True):
        left, right = st.columns([1.5, 1], gap="large")
        with left:
            st.subheader("01 选择商品资料", anchor=False)
            st.caption("Excel（.xlsx）或CSV，每行一个SKU，第一行是列名。最多5MB、5000行、80列。")
            epoch = st.session_state.get("upload_epoch", 0)
            upload = st.file_uploader(
                "上传商品表", type=["xlsx", "csv"], key=f"upload_{epoch}",
                on_change=select_upload, label_visibility="collapsed",
            )
            st.download_button("下载空白Excel模板", download_template(), "商品资料模板.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore")
        with right:
            st.subheader("先用样例体验", anchor=False)
            st.write("内置10行虚构服装资料，包含缺项、重复SKU、价格差异和异常库存。")
            st.button("载入演示数据", key="load_demo", on_click=select_demo, width="stretch")
            st.download_button("下载演示Excel", download_demo(), "演示商品资料_虚构.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore", width="stretch")

    if upload is None and st.session_state.get("source_mode") != "demo":
        st.markdown('<div class="empty-note"><b>准备好一张商品表，就能开始。</b><br>必填字段：商品名称、SKU编码、规格、售价、库存。其他列也会保留。</div>', unsafe_allow_html=True)
        return

    try:
        table, source_id = current_table(upload)
    except ValueError as error:
        st.error(str(error))
        st.caption("可以重新选择文件，或先载入演示数据确认操作流程。")
        return
    if table is None:
        return

    st.subheader("02 确认并开始检查", anchor=False)
    st.text(f"资料：{table['source']}  ·  {len(table['rows'])}行商品记录")
    for notice in table.get("notices", []):
        st.info(notice)
    mapping = mapping_controls(table, source_id)
    with st.expander("查看导入的原始资料"):
        frame = source_frame([dict(zip(table["headers"], row)) for row in table["rows"]], table["row_numbers"])
        st.dataframe(frame, hide_index=True, width="stretch")

    signature = hashlib.sha256(json.dumps({"source": source_id, "mapping": mapping}, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    run_column, clear_column, _ = st.columns([2, 1, 3])
    with run_column:
        run = st.button("开始整理与检查", key="run_check", type="primary", width="stretch")
    with clear_column:
        st.button("清空资料", key="clear_input", on_click=reset_input, width="stretch")

    if run:
        try:
            with st.spinner("正在整理资料并检查规则…"):
                start = time.perf_counter()
                report = analyze_table(table, mapping=mapping)
                elapsed_ms = (time.perf_counter() - start) * 1000
                files = {"xlsx": export_xlsx(report), "html": export_html(report)}
            st.session_state["report"] = report
            st.session_state["report_signature"] = signature
            st.session_state["report_files"] = files
            st.session_state["elapsed_ms"] = elapsed_ms
            st.session_state["result_status"] = "全部记录"
            st.session_state["result_search"] = ""
            st.session_state["issue_level"] = "全部问题"
        except ValueError as error:
            st.error(str(error))
            return

    if "report" not in st.session_state:
        st.caption("只自动整理明确的格式问题。缺失的价格、库存及矛盾资料会保留并提示补充。")
        return
    if signature != st.session_state.get("report_signature"):
        st.info("资料或列名对应已变化，请重新检查后查看和导出结果。")
        return
    render_report(st.session_state["report"], st.session_state["report_files"], st.session_state["elapsed_ms"])


if __name__ == "__main__":
    main()
