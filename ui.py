"""操作界面的样式和报告视图，不参与业务判断。"""
import pandas as pd
import streamlit as st


def install_style():
    st.markdown("""<style>
    .stApp {font-family:'Microsoft YaHei UI','Microsoft YaHei',sans-serif;}
    .block-container {max-width:1250px;padding-top:4.4rem;padding-bottom:3rem;}
    h1 {font-size:2rem!important;letter-spacing:-.03em;line-height:1.35!important;}
    h3 {font-size:1.05rem!important;line-height:1.5!important;}
    .workspace-label {font-size:.8rem;color:#52675b;margin-bottom:1.1rem;}
    .workspace-label>span {margin:0 .7rem;color:#7a897e;}
    .workspace-label .local-label {float:right;color:#28634c;border-left:2px solid #28634c;padding-left:.6rem;}
    [data-testid="stVerticalBlockBorderWrapper"] {border-radius:10px;}
    [data-testid="stMetric"] {background:#e9eee6;padding:1rem 1.2rem;border-radius:8px;}
    [data-testid="stMetricValue"] {font-variant-numeric:tabular-nums;font-size:1.9rem;}
    [data-testid="stMetricLabel"] {font-size:.85rem;}
    .empty-note {padding:1.3rem 0;color:#52675b;line-height:2;font-size:.9rem;}
    .empty-note b {color:#20342a;font-weight:500;}
    button {touch-action:manipulation;transition:transform .12s ease,background-color .12s ease!important;}
    button:active {transform:scale(.98);}
    button:focus-visible {outline:3px solid #678f76!important;outline-offset:3px;}
    @media(max-width:640px) {
      .block-container {padding:4rem 1rem 2rem;}
      h1 {font-size:1.55rem!important;}
      [data-testid="stHorizontalBlock"] {flex-wrap:wrap;}
      [data-testid="stColumn"] {min-width:min(100%,150px)!important;flex:1 1 150px!important;}
      .workspace-label .local-label {float:none;display:inline-block;margin-left:0;}
    }
    @media(prefers-reduced-motion:reduce) {button {transition:none!important;}button:active {transform:none;}}
    </style>""", unsafe_allow_html=True)


def issue_records(report):
    items = []
    for issue in report["file_issues"]:
        items.append({"原表行号": None, "SKU编码": "整张表", "字段": issue["field"], "级别": issue["level"], "问题": issue["message"], "建议": issue["suggestion"]})
    for row in report["rows"]:
        for issue in row["issues"]:
            items.append({"原表行号": row["row"], "SKU编码": row["data"].get("SKU编码", ""), "字段": issue["field"], "级别": issue["level"], "问题": issue["message"], "建议": issue["suggestion"]})
    return items


def source_frame(records, row_numbers):
    """添加本次来源行号，不覆盖上传表中可能同名的业务或历史字段。"""
    frame = pd.DataFrame(records)
    label = "原表行号"
    while label in frame.columns:
        label += "（本次）"
    frame.insert(0, label, row_numbers)
    return frame


def render_report(report, files, elapsed_ms):
    st.divider()
    st.subheader("03 查看检查结果", anchor=False)
    summary = report["summary"]
    if summary["issue_rows"]:
        st.warning(f"检查完成：{summary['issue_rows']}行资料需要处理。请先修正错误，再核对重复、价格差异等提示。")
    else:
        st.success("检查完成：本次未发现规则内的问题。")
    columns = st.columns(4)
    for column, label, key in zip(columns, ["商品记录", "通过", "需修正", "待核对"], ["total_rows", "passed_rows", "error_rows", "review_rows"]):
        column.metric(label, summary[key])
    st.caption(f"自动规范 {summary['normalized_cells']} 处格式 · 问题条目 {summary['issue_count']} 条 · 检查程序耗时 {elapsed_ms:.1f} ms（不含人工核对）")

    download_columns = st.columns(2)
    with download_columns[0]:
        st.download_button("下载整理结果 Excel", files["xlsx"], "商品资料_整理结果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore", type="primary", width="stretch")
    with download_columns[1]:
        st.download_button("下载检查报告 HTML", files["html"], "商品资料_检查报告.html", mime="text/html", on_click="ignore", width="stretch")
    st.caption("Excel包含整理结果、问题清单、修改记录、原始资料及说明。保留全部记录，原文件不变。")

    result_tab, issues_tab, changes_tab, raw_tab = st.tabs(["整理后的资料", "问题清单", "修改记录", "原始资料"])
    with result_tab:
        filter_column, search_column = st.columns([1, 2])
        with filter_column:
            status = st.selectbox("按状态筛选", ["全部记录", "需修正", "待核对", "通过"], key="result_status")
        with search_column:
            search = st.text_input("搜索商品名称或SKU", key="result_search", placeholder="输入商品名称或编码")
        selected = [row for row in report["rows"] if (status == "全部记录" or row["status"] == status) and (not search or search.casefold() in (row["data"].get("商品名称", "") + " " + row["data"].get("SKU编码", "")).casefold())]
        records = [{"原表行号": row["row"], "检查状态": row["status"], **row["data"], "问题说明": "；".join(i["message"] for i in row["issues"])} for row in selected]
        if not records:
            st.info("没有符合筛选条件的记录，可以调整状态或搜索内容。")
        else:
            frame = pd.DataFrame(records)
            styled = frame.style.map(lambda value: "background-color:#fbe6df;color:#8e3022" if value == "需修正" else "background-color:#fff0ce;color:#765610" if value == "待核对" else "background-color:#deecdf;color:#285b3c", subset=["检查状态"])
            st.dataframe(styled, hide_index=True, width="stretch", height=min(470, 40 + 36 * len(frame)))
        st.caption(f"显示 {len(selected)} / {summary['total_rows']} 行。需要补充的信息可在导出的表格中修改，再上传复查。")
    with issues_tab:
        issues = issue_records(report)
        if issues:
            level = st.radio("问题级别", ["全部问题", "错误", "待核对"], horizontal=True, key="issue_level")
            filtered = [item for item in issues if level == "全部问题" or item["级别"] == level]
            if filtered:
                st.dataframe(pd.DataFrame(filtered), hide_index=True, width="stretch", column_config={"原表行号": st.column_config.NumberColumn(format="%d")})
            else:
                st.info("该级别没有问题。")
        else:
            st.success("未发现规则内的问题。")
    with changes_tab:
        changes = [{"原表行号": row["row"], "SKU编码": row["data"].get("SKU编码", ""), "字段": change["field"], "原值": change["before"], "整理后": change["after"]} for row in report["rows"] for change in row["changes"]]
        if changes:
            st.dataframe(pd.DataFrame(changes), hide_index=True, width="stretch")
        else:
            st.info("资料格式已规范，本次没有自动修改。")
    with raw_tab:
        raw = source_frame([row["raw"] for row in report["rows"]], [row["row"] for row in report["rows"]])
        st.dataframe(raw, hide_index=True, width="stretch")
        st.caption(f"来源：{report['source']} · 工作表：{report['sheet'] or 'CSV'} · 检查时间：{report['generated_at']}")
