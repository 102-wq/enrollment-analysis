import io
import zipfile
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 1. 页面基本配置
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="招生数据动态管理与多维分析系统",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 招生数据动态管理与多维分析系统")
st.caption("2026年9月数据 - 已同步最新招生明细与全量图表看板")
st.markdown("---")

# -----------------------------------------------------------------------------
# 2. 侧边栏：纯动物图像脱敏配置
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ 数据管理与设置")

RAW_PERSONS = ["覃小燕", "左丹丹", "梁书华", "古晨晓", "周欢喜"]
ANIMALS = ["🦊", "🐼", "🦁", "🐰", "🐯"]

enable_anonymize = st.sidebar.checkbox("开启数据脱敏 / 纯图像模式", value=True)

if enable_anonymize:
    alias_map = {p: ANIMALS[i] for i, p in enumerate(RAW_PERSONS)}
else:
    alias_map = {p: p for p in RAW_PERSONS}

PERSONS = [alias_map[p] for p in RAW_PERSONS]

if enable_anonymize:
    with st.sidebar.expander("👁️ 视角对照图表（仅管理者可见）", expanded=False):
        mapping_df = pd.DataFrame({
            "真实姓名": RAW_PERSONS,
            "展示图像": PERSONS
        })
        st.dataframe(mapping_df, hide_index=True, use_container_width=True)

# -----------------------------------------------------------------------------
# 3. 基础数据定义
# -----------------------------------------------------------------------------
MAJORS = [
    ("给排水专业", 19, [4, 3, 6, 3, 3]),
    ("发输电专业", 10, [2, 3, 2, 1, 2]),
    ("供配电专业", 11, [3, 4, 2, 1, 1]),
    ("环保专业", 13, [4, 2, 2, 3, 2]),
    ("环评专业", 9, [2, 2, 1, 3, 1]),
    ("岩土专业", 8, [2, 2, 2, 1, 1]),
    ("暖通专业", 16, [3, 3, 3, 2, 5]),
    ("结构专业", 20, [4, 4, 4, 5, 3]),
    ("道路专业", 5, [1, 1, 1, 1, 1]),
    ("233网校", 5, [1, 1, 1, 1, 1]),
    ("电气基础", 53, [11, 16, 11, 8, 7]),
    ("环保基础", 30, [9, 6, 6, 5, 4]),
    ("岩土基础", 36, [9, 8, 7, 6, 6]),
    ("水基础", 17, [4, 4, 5, 2, 2]),
    ("暖通基础", 27, [5, 5, 4, 4, 9]),
    ("结构基础", 9, [2, 2, 1, 3, 1]),
    ("公共基础", 4, [2, 1, 1, 0, 0]),
    ("道路基础", 6, [1, 1, 1, 1, 2]),
    ("水利水电基础", 10, [2, 2, 3, 2, 1]),
]

DATES = ["9月1日", "9月2日", "9月3日", "9月4日", "9月5日", "9月6日", "9月7日"]
WEEKDAYS = ["星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "星期一"]

def init_default_data():
    base_data = []
    for idx, (name, total_target, person_tgts) in enumerate(MAJORS, 1):
        row = {"序号": idx, "专业/基础名称": name, "目标人数": total_target}
        for p, tgt in zip(RAW_PERSONS, person_tgts):
            row[f"{p}_目标"] = tgt
            row[f"{p}_实际"] = 0
        row["其他人员_实际"] = 0
        base_data.append(row)
    st.session_state["base_targets"] = pd.DataFrame(base_data)

    st.session_state["daily_deltas"] = {
        "9月1日": [("电气基础", "覃小燕_实际", 1)],
        "9月2日": [
            ("环保专业", "覃小燕_实际", 1),
            ("环保专业", "左丹丹_实际", 1),
            ("电气基础", "覃小燕_实际", 1),
            ("发输电专业", "其他人员_实际", 1),
            ("233网校", "其他人员_实际", 1),
        ],
        "9月3日": [
            ("环评专业", "左丹丹_实际", 1),
            ("暖通专业", "周欢喜_实际", 1),
            ("环保基础", "左丹丹_实际", 1),
            ("环保基础", "其他人员_实际", 1),
        ],
        "9月4日": [
            ("电气基础", "覃小燕_实际", 1),
            ("电气基础", "左丹丹_实际", 1),
            ("环保基础", "覃小燕_实际", 1),
            ("水利水电基础", "其他人员_实际", 1),
        ],
        "9月5日": [
            ("给排水专业", "梁书华_实际", 1),
            ("暖通专业", "周欢喜_实际", 1),
            ("岩土基础", "梁书华_实际", 1),
            ("暖通基础", "周欢喜_实际", 1),
        ],
        "9月6日": [
            ("给排水专业", "梁书华_实际", 1),
            ("环保专业", "其他人员_实际", 1),
            ("岩土专业", "梁书华_实际", 1),
        ],
        "9月7日": [
            ("电气基础", "覃小燕_实际", 3),
            ("环保基础", "覃小燕_实际", 1),
            ("水基础", "覃小燕_实际", 1),
            ("环保基础", "左丹丹_实际", 1),
            ("暖通基础", "左丹丹_实际", 1),
            ("暖通专业", "梁书华_实际", 1),
            ("岩土基础", "梁书华_实际", 1),
            ("暖通基础", "梁书华_实际", 1),
            ("公共基础", "古晨晓_实际", 1),
            ("环保基础", "古晨晓_实际", 1),
        ],
    }

if "base_targets" not in st.session_state or "daily_deltas" not in st.session_state:
    init_default_data()

# -----------------------------------------------------------------------------
# 4. 录入表单
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("➕ 快捷录入新增招生")

with st.sidebar.form("add_delta_form"):
    input_date = st.selectbox("选择日期", DATES)
    input_major = st.selectbox("选择专业/基础", [m[0] for m in MAJORS])
    input_person_disp = st.selectbox("选择归属人员", PERSONS + ["其他人员"])
    input_val = st.number_input("新增人数", min_value=1, value=1, step=1)

    submit_btn = st.form_submit_button("确认录入数据")
    if submit_btn:
        inv_alias_map = {v: k for k, v in alias_map.items()}
        raw_person_name = inv_alias_map.get(input_person_disp, input_person_disp)
        
        target_col = f"{raw_person_name}_实际"
        if input_date not in st.session_state["daily_deltas"]:
            st.session_state["daily_deltas"][input_date] = []
        st.session_state["daily_deltas"][input_date].append(
            (input_major, target_col, int(input_val))
        )
        st.sidebar.success(f"已成功添加：{input_date} {input_major} - {input_person_disp} +{input_val}人")

if st.sidebar.button("🔄 重置为默认演示数据"):
    init_default_data()
    st.sidebar.info("数据已重置！")

# -----------------------------------------------------------------------------
# 5. 时间汇总粒度切片筛选
# -----------------------------------------------------------------------------
st.subheader("🗓️ 时间汇总粒度筛选")

f_col1, f_col2 = st.columns(2)

with f_col1:
    time_granularity_type = st.selectbox(
        "选择时间汇总粒度：",
        ["按日（单日切片）", "按周（周度汇总）", "按月（月度全量）"]
    )

with f_col2:
    if time_granularity_type == "按日（单日切片）":
        selected_time_range = st.selectbox("选择具体日期：", DATES, index=len(DATES)-1)
        selected_dates_list = [selected_time_range]
    elif time_granularity_type == "按周（周度汇总）":
        selected_time_range = st.selectbox("选择具体周：", ["2026年第36周 (9月1日-9月7日)"])
        selected_dates_list = DATES  # 包含当周的所有日期
    else:
        selected_time_range = st.selectbox("选择具体月份：", ["2026年9月全月"])
        selected_dates_list = DATES  # 包含当月的所有日期

def get_processed_df_by_dates(dates_list):
    df_result = st.session_state["base_targets"].copy()
    for d in dates_list:
        for major, col, val in st.session_state["daily_deltas"].get(d, []):
            df_result.loc[df_result["专业/基础名称"] == major, col] += val
    return df_result

calc_df = get_processed_df_by_dates(selected_dates_list)

act_cols = [c for c in calc_df.columns if c.endswith("_实际")]
calc_df["实际完成"] = calc_df[act_cols].sum(axis=1)
calc_df["与目标之差"] = calc_df["实际完成"] - calc_df["目标人数"]

num_cols = [c for c in calc_df.columns if c not in ["序号", "专业/基础名称"]]
sum_row = {"专业/基础名称": "合计"}
for c in num_cols:
    sum_row[c] = int(calc_df[c].sum())

diff_row = {"专业/基础名称": "与目标之差"}
for p in RAW_PERSONS:
    diff_row[f"{p}_目标"] = ""
    diff_row[f"{p}_实际"] = sum_row[f"{p}_实际"] - sum_row[f"{p}_目标"]
diff_row["其他人员_实际"] = ""
diff_row["目标人数"] = ""
diff_row["实际完成"] = ""
diff_row["与目标之差"] = sum_row["与目标之差"]

total_target_cum = sum_row["目标人数"]
total_actual_cum = sum_row["实际完成"]
cum_rate_val = (total_actual_cum / total_target_cum * 100) if total_target_cum > 0 else 0

rate_row = {"专业/基础名称": "目标完成率"}
for c in num_cols:
    rate_row[c] = ""
rate_row["实际完成"] = f"{cum_rate_val:.2f}%"

# -----------------------------------------------------------------------------
# 6. HTML 数据表格
# -----------------------------------------------------------------------------
def build_html_document(df, sum_r, diff_r, rate_r, p_names, raw_p_names, range_title):
    rows_html = ""
    for idx, row in df.iterrows():
        person_cells = ""
        for rp in raw_p_names:
            tgt_val = row[f"{rp}_目标"]
            act_val = row[f"{rp}_实际"]
            person_cells += f'<td class="num">{tgt_val}</td><td class="num">{act_val if act_val > 0 else ""}</td>'

        rows_html += f"""
        <tr>
            <td class="num">{row['序号']}</td>
            <td class="zh">{row['专业/基础名称']}</td>
            <td class="num">{row['目标人数']}</td>
            {person_cells}
            <td class="num">{row['other_act'] if row['other_act'] > 0 else ''}</td>
            <td class="num">{row['目标人数']}</td>
            <td class="num">{row['实际完成']}</td>
            <td class="num">{row['与目标之差']}</td>
        </tr>
        """

    person_headers = "".join([f'<th colspan="2" class="bg-person">{p}</th>' for p in p_names])
    sub_headers = '<th class="bg-header">目标</th><th class="bg-header">实际</th>' * len(p_names)

    sum_person_cells = ""
    diff_person_cells = ""
    for rp in raw_p_names:
        sum_person_cells += f'<td class="bg-total num">{sum_r[f"{rp}_目标"]}</td><td class="bg-total num">{sum_r[f"{rp}_实际"]}</td>'
        diff_person_cells += f'<td class="bg-total"></td><td class="bg-total num">{diff_r[f"{rp}_实际"]}</td>'

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ margin: 0; padding: 0; font-family: SimSun, "Times New Roman", serif; background-color: #ffffff; }}
        .table-container {{ width: 100%; overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; text-align: center; border: 1px solid #7F7F7F; }}
        th, td {{ border: 1px solid #7F7F7F; padding: 6px 4px; font-weight: normal; color: #000000; }}
        .bg-title {{ background-color: #D9EAD3; font-size: 15px; font-weight: bold; }}
        .bg-header {{ background-color: #F2F2F2; font-weight: bold; }}
        .bg-person {{ background-color: #D0E0E3; font-weight: bold; font-size: 16px; }}
        .bg-total {{ background-color: #FFF2CC; font-weight: bold; }}
        .num {{ font-family: "Times New Roman", serif; }}
        .zh {{ font-family: SimSun, serif; }}
    </style>
    </head>
    <body>
    <div class="table-container">
    <table>
        <tr><td colspan="17" class="bg-title">2026年招生数据动态表 ({range_title})</td></tr>
        <tr>
            <th rowspan="2" class="bg-header">序号</th>
            <th rowspan="2" class="bg-header">专业/基础名称</th>
            <th rowspan="2" class="bg-header">目标人数</th>
            {person_headers}
            <th rowspan="2" class="bg-header">其他人员</th>
            <th rowspan="2" class="bg-header">目标人数</th>
            <th rowspan="2" class="bg-header">实际完成</th>
            <th rowspan="2" class="bg-header">与目标之差</th>
        </tr>
        <tr>
            {sub_headers}
        </tr>
        {rows_html}
        <tr>
            <td class="bg-total"></td><td class="bg-total zh">{sum_r['专业/基础名称']}</td><td class="bg-total num">{sum_r['目标人数']}</td>
            {sum_person_cells}
            <td class="bg-total num">{sum_r['其他人员_实际']}</td>
            <td class="bg-total num">{sum_r['目标人数']}</td>
            <td class="bg-total num">{sum_r['实际完成']}</td>
            <td class="bg-total num">{sum_r['与目标之差']}</td>
        </tr>
        <tr>
            <td class="bg-total"></td><td class="bg-total zh">{diff_r['专业/基础名称']}</td><td class="bg-total"></td>
            {diff_person_cells}
            <td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td>
            <td class="bg-total num">{diff_r['与目标之差']}</td>
        </tr>
        <tr>
            <td class="bg-total"></td><td class="bg-total zh">{rate_r['专业/基础名称']}</td><td class="bg-total"></td>
            <td class="bg-total" colspan="{len(p_names)*2}"></td>
            <td class="bg-total"></td><td class="bg-total"></td>
            <td class="bg-total num">{rate_r['实际完成']}</td>
            <td class="bg-total"></td>
        </tr>
    </table>
    </div>
    </body>
    </html>
    """
    return full_html

st.markdown(f"### 📝 招生数据动态明细 （区间：{selected_time_range}）")
calc_df['other_act'] = calc_df['其他人员_实际']
html_code = build_html_document(calc_df, sum_row, diff_row, rate_row, PERSONS, RAW_PERSONS, selected_time_range)
st.components.v1.html(html_code, height=680, scrolling=True)

# -----------------------------------------------------------------------------
# 7. 可视化图表展示
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("### 📊 基础指标分析")

c1, c2 = st.columns(2)

with c1:
    person_target_vals = [sum_row[f"{p}_目标"] for p in RAW_PERSONS]
    person_actual_vals = [sum_row[f"{p}_实际"] for p in RAW_PERSONS]

    fig_person = go.Figure(
        data=[
            go.Bar(
                name="目标人数", x=PERSONS, y=person_target_vals, 
                marker_color="#0B3C5D", text=person_target_vals, textposition="outside",
                textfont=dict(size=12, color="#000000")
            ),
            go.Bar(
                name="实际完成", x=PERSONS, y=person_actual_vals, 
                marker_color="#FF3D00", text=person_actual_vals, textposition="outside",
                textfont=dict(size=12, color="#000000")
            ),
        ]
    )
    fig_person.update_layout(
        title="<b>各成员目标 vs 实际完成对比</b>",
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        margin=dict(l=20, r=20, t=50, b=20),
        xaxis=dict(tickfont=dict(size=16)),
        yaxis=dict(gridcolor="#E0E0E0", showline=True, linewidth=1, linecolor="#000000")
    )
    st.plotly_chart(fig_person, use_container_width=True)

with c2:
    top_majors = calc_df.sort_values(by="实际完成", ascending=False).head(8)
    
    fig_major = px.bar(
        top_majors, x="实际完成", y="专业/基础名称", orientation="h",
        title="<b>招生完成人数 Top 8 专业/基础</b>", text="实际完成",
        color="专业/基础名称",
        color_discrete_sequence=px.colors.qualitative.Bold
    )
    fig_major.update_layout(
        yaxis={"categoryorder": "total ascending"},
        showlegend=False,
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        margin=dict(l=20, r=20, t=50, b=20),
        xaxis=dict(gridcolor="#E0E0E0", showline=True, linewidth=1, linecolor="#000000")
    )
    fig_major.update_traces(
        textposition="outside", 
        textfont=dict(size=12, color="#000000")
    )
    st.plotly_chart(fig_major, use_container_width=True)

# -----------------------------------------------------------------------------
# 8. 时间维度趋势分析
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("### 🔄 动态趋势与人员贡献构成分析")

ctrl_c1, ctrl_c2 = st.columns(2)

with ctrl_c1:
    person_mode = st.radio("选择分析视角：", ["全体人员", "单人独立分析", "多人员对比分析"], horizontal=True)

with ctrl_c2:
    if person_mode == "单人独立分析":
        selected_person_disp = st.selectbox("选择分析成员：", PERSONS + ["其他人员"])
    elif person_mode == "多人员对比分析":
        selected_persons_disp = st.multiselect("选择对比成员：", PERSONS + ["其他人员"], default=PERSONS[:3])
    else:
        selected_person_disp = "全体人员"

time_records = []
for d_idx, d in enumerate(DATES):
    w = WEEKDAYS[d_idx]
    is_weekend = "周末" if w in ["星期六", "星期日"] else "工作日"
    d_dict = {p: 0 for p in RAW_PERSONS + ["其他人员"]}
    for major, col, val in st.session_state["daily_deltas"].get(d, []):
        p_name = col.replace("_实际", "")
        d_dict[p_name] += val
    
    for p_name, val in d_dict.items():
        disp_p_name = alias_map.get(p_name, p_name)
        time_records.append({"日期": d, "星期": w, "类型": is_weekend, "月份": "2026年9月", "人员": disp_p_name, "新增报名数": val})

df_time_series = pd.DataFrame(time_records)

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    if person_mode == "全体人员":
        df_chart_line = df_time_series.groupby("日期", as_index=False)["新增报名数"].sum()
        fig_line = px.line(df_chart_line, x="日期", y="新增报名数", markers=True, title="📈 <b>全体人员招生趋势 (按日明细)</b>", text="新增报名数")
        fig_line.update_traces(textposition="top center", line_color="#D50000", line_width=2, marker=dict(size=6, color="#D50000"))
    elif person_mode == "单人独立分析":
        df_sub = df_time_series[df_time_series["人员"] == selected_person_disp]
        df_chart_line = df_sub.groupby("日期", as_index=False)["新增报名数"].sum()
        fig_line = px.line(df_chart_line, x="日期", y="新增报名数", markers=True, title=f"📈 <b>【{selected_person_disp}】趋势 (按日明细)</b>", text="新增报名数")
        fig_line.update_traces(textposition="top center", line_color="#2962FF", line_width=2, marker=dict(size=6, color="#2962FF"))
    else:
        df_sub = df_time_series[df_time_series["人员"].isin(selected_persons_disp)]
        df_chart_line = df_sub.groupby(["日期", "人员"], as_index=False)["新增报名数"].sum()
        fig_line = px.line(
            df_chart_line, x="日期", y="新增报名数", color="人员", markers=True, 
            title="📈 <b>多人招生趋势对比 (按日明细)</b>",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_line.update_traces(line_width=2, marker=dict(size=6))

    fig_line.update_xaxes(categoryorder="array", categoryarray=DATES)
    fig_line.update_layout(
        yaxis_title="新增报名人数", xaxis_title="日期", hovermode="x unified",
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        margin=dict(l=20, r=20, t=50, b=20),
        yaxis=dict(gridcolor="#E0E0E0")
    )
    st.plotly_chart(fig_line, use_container_width=True)

with chart_col2:
    if person_mode == "全体人员":
        df_pie = df_time_series.groupby("人员", as_index=False)["新增报名数"].sum()
        fig_pie = px.pie(
            df_pie, values="新增报名数", names="人员", title="🍩 <b>全体人员招生贡献占比</b>", 
            hole=0.4, color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_pie.update_traces(textinfo="label+percent+value", textposition="inside")
    elif person_mode == "单人独立分析":
        inv_map = {v: k for k, v in alias_map.items()}
        raw_sel_p = inv_map.get(selected_person_disp, selected_person_disp)
        major_records = []
        for d in DATES:
            for major, col, val in st.session_state["daily_deltas"].get(d, []):
                p_name = col.replace("_实际", "")
                if p_name == raw_sel_p:
                    major_records.append({"专业/基础": major, "新增人数": val})
        df_major_pie = pd.DataFrame(major_records)
        if not df_major_pie.empty:
            df_major_pie = df_major_pie.groupby("专业/基础", as_index=False)["新增人数"].sum()
            fig_pie = px.pie(
                df_major_pie, values="新增人数", names="专业/基础", 
                title=f"🍩 <b>【{selected_person_disp}】专业成交构成</b>", 
                hole=0.4, color_discrete_sequence=px.colors.qualitative.Set1
            )
            fig_pie.update_traces(textinfo="label+percent+value", textposition="inside", marker=dict(line=dict(color="#FFFFFF", width=2)))
        else:
            fig_pie = go.Figure()
            fig_pie.update_layout(title=f"【{selected_person_disp}】暂无增量数据")
    else:
        df_pie = df_time_series[df_time_series["人员"].isin(selected_persons_disp)].groupby("人员", as_index=False)["新增报名数"].sum()
        fig_pie = px.pie(
            df_pie, values="新增报名数", names="人员", title=f"🍩 <b>对比成员占比构成</b>", 
            hole=0.4, color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_pie.update_traces(textinfo="label+percent+value", textposition="inside")

    fig_pie.update_layout(paper_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20))
    st.plotly_chart(fig_pie, use_container_width=True)

# -----------------------------------------------------------------------------
# 9. 导出 Excel / ZIP 打包
# -----------------------------------------------------------------------------
def export_color_excel(calc_df, sum_row, diff_row, rate_row, persons_disp, raw_persons):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "招生数据动态表"
    ws.views.sheetView[0].showGridLines = True

    TITLE_FILL = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    HEADER_BG = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    PERSON_BG = PatternFill(start_color="D0E0E3", end_color="D0E0E3", fill_type="solid")
    TOTAL_ROW_BG = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    FONT_TITLE = Font(name="SimSun", size=14, bold=True)
    FONT_HEADER = Font(name="SimSun", size=10, bold=True)
    FONT_ANIMAL = Font(size=14, bold=True)
    FONT_BODY_NUM = Font(name="Times New Roman", size=10)
    FONT_BODY_ZH = Font(name="SimSun", size=10)

    BORDER_THIN = Border(
        left=Side(style="thin", color="7F7F7F"),
        right=Side(style="thin", color="7F7F7F"),
        top=Side(style="thin", color="7F7F7F"),
        bottom=Side(style="thin", color="7F7F7F"),
    )
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:Q1")
    t_cell = ws["A1"]
    t_cell.value = f"2026年招生数据动态表 ({selected_time_range})"
    t_cell.font = FONT_TITLE
    t_cell.fill = TITLE_FILL
    t_cell.alignment = ALIGN_CENTER

    ws.merge_cells("A2:A3")
    ws["A2"] = "序号"
    ws.merge_cells("B2:B3")
    ws["B2"] = "专业/基础名称"
    ws.merge_cells("C2:C3")
    ws["C2"] = "目标人数"

    col_idx = 4
    for p in persons_disp:
        ws.merge_cells(start_row=2, start_column=col_idx, end_row=2, end_column=col_idx + 1)
        p_cell = ws.cell(row=2, column=col_idx, value=p)
        p_cell.fill = PERSON_BG
        p_cell.font = FONT_ANIMAL
        ws.cell(row=3, column=col_idx, value="目标")
        ws.cell(row=3, column=col_idx + 1, value="实际")
        col_idx += 2

    ws.merge_cells("N2:N3")
    ws["N2"] = "其他人员"
    ws.merge_cells("O2:O3")
    ws["O2"] = "目标人数"
    ws.merge_cells("P2:P3")
    ws["P2"] = "实际完成"
    ws.merge_cells("Q2:Q3")
    ws["Q2"] = "与目标之差"

    for r in range(2, 4):
        for c in range(1, 18):
            cell = ws.cell(row=r, column=c)
            if not cell.fill.start_color.rgb:
                cell.fill = HEADER_BG
            if not (r == 2 and c in range(4, 14, 2)):
                cell.font = FONT_HEADER
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN

    curr_r = 4
    for idx, row in calc_df.iterrows():
        ws.cell(row=curr_r, column=1, value=row["序号"]).font = FONT_BODY_NUM
        ws.cell(row=curr_r, column=2, value=row["专业/基础名称"]).font = FONT_BODY_ZH
        ws.cell(row=curr_r, column=3, value=row["目标人数"]).font = FONT_BODY_NUM

        c_offset = 4
        for rp in raw_persons:
            ws.cell(row=curr_r, column=c_offset, value=row[f"{rp}_目标"]).font = FONT_BODY_NUM
            ws.cell(row=curr_r, column=c_offset + 1, value=row[f"{rp}_实际"] if row[f"{rp}_实际"] > 0 else "").font = FONT_BODY_NUM
            c_offset += 2

        ws.cell(row=curr_r, column=c_offset, value=row.get("其他人员_实际", 0) if row.get("其他人员_实际", 0) > 0 else "").font = FONT_BODY_NUM
        ws.cell(row=curr_r, column=c_offset + 1, value=row["目标人数"]).font = FONT_BODY_NUM
        ws.cell(row=curr_r, column=c_offset + 2, value=row["实际完成"]).font = FONT_BODY_NUM
        ws.cell(row=curr_r, column=c_offset + 3, value=row["与目标之差"]).font = FONT_BODY_NUM

        for c in range(1, 18):
            cell = ws.cell(row=curr_r, column=c)
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN
        curr_r += 1

    for r_data in [sum_row, diff_row, rate_row]:
        ws.cell(row=curr_r, column=2, value=r_data["专业/基础名称"]).font = FONT_HEADER
        ws.cell(row=curr_r, column=3, value=r_data.get("目标人数", "")).font = FONT_HEADER

        c_offset = 4
        for rp in raw_persons:
            ws.cell(row=curr_r, column=c_offset, value=r_data.get(f"{rp}_目标", "")).font = FONT_HEADER
            ws.cell(row=curr_r, column=c_offset + 1, value=r_data.get(f"{rp}_实际", "")).font = FONT_HEADER
            c_offset += 2

        ws.cell(row=curr_r, column=c_offset, value=r_data.get("其他人员_实际", "")).font = FONT_HEADER
        ws.cell(row=curr_r, column=c_offset + 1, value=r_data.get("目标人数", "")).font = FONT_HEADER
        ws.cell(row=curr_r, column=c_offset + 2, value=r_data.get("实际完成", "")).font = FONT_HEADER
        ws.cell(row=curr_r, column=c_offset + 3, value=r_data.get("与目标之差", "")).font = FONT_HEADER

        for c in range(1, 18):
            cell = ws.cell(row=curr_r, column=c)
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN
            cell.fill = TOTAL_ROW_BG
        curr_r += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def build_full_export_pack(excel_bytes, fig_dict):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"招生数据动态表_{selected_time_range}.xlsx", excel_bytes)
        for fig_name, fig_obj in fig_dict.items():
            try:
                img_bytes = fig_obj.to_image(format="png", width=1200, height=700, scale=2)
                zip_file.writestr(f"导出图表_{fig_name}.png", img_bytes)
            except Exception:
                zip_file.writestr("图表导出提示.txt", "生成高分辨率PNG图片需要 kaleido 依赖包 (pip install kaleido)")

    return zip_buffer.getvalue()

st.markdown("---")
st.markdown("### 📥 结果导出与报表生成")

excel_data = export_color_excel(calc_df, sum_row, diff_row, rate_row, PERSONS, RAW_PERSONS)

fig_collection = {
    "1_各人员目标实际对比柱状图": fig_person,
    "2_Top8专业完成横柱图": fig_major,
    "3_招生动态趋势折线图": fig_line,
    "4_多维构成占比饼图": fig_pie,
}

zip_data = build_full_export_pack(excel_data, fig_collection)

col_d1, col_d2 = st.columns(2)

with col_d1:
    st.download_button(
        label="📊 导出 Excel 动态明细表 (.xlsx)",
        data=excel_data,
        file_name=f"2026年招生数据动态表_{selected_time_range}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_d2:
    st.download_button(
        label="📦 一键打包导出全量图表与 Excel 压缩包 (.zip)",
        data=zip_data,
        file_name=f"2026年招生看板及图表全量导出包_{selected_time_range}.zip",
        mime="application/zip",
        use_container_width=True,
    )
