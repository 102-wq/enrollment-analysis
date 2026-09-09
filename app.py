import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
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
st.caption("2026年9月数据 - 包含多维人员/时间粒度折线图与饼图")
st.markdown("---")

# -----------------------------------------------------------------------------
# 2. 基础数据定义
# -----------------------------------------------------------------------------
PERSONS = ["人员A", "人员B", "人员C", "人员D", "人员E"]

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
        for p, tgt in zip(PERSONS, person_tgts):
            row[f"{p}_目标"] = tgt
            row[f"{p}_实际"] = 0
        row["其他人员_实际"] = 0
        base_data.append(row)
    st.session_state["base_targets"] = pd.DataFrame(base_data)

    st.session_state["daily_deltas"] = {
        "9月1日": [("电气基础", "人员A_实际", 1)],
        "9月2日": [
            ("环保专业", "人员A_实际", 1),
            ("环保专业", "人员B_实际", 1),
            ("电气基础", "人员A_实际", 1),
            ("发输电专业", "其他人员_实际", 1),
            ("233网校", "其他人员_实际", 1),
        ],
        "9月3日": [
            ("环评专业", "人员B_实际", 1),
            ("暖通专业", "人员E_实际", 1),
            ("环保基础", "人员B_实际", 1),
            ("环保基础", "其他人员_实际", 1),
        ],
        "9月4日": [
            ("电气基础", "人员A_实际", 1),
            ("电气基础", "人员B_实际", 1),
            ("环保基础", "人员A_实际", 1),
            ("水利水电基础", "其他人员_实际", 1),
        ],
        "9月5日": [
            ("给排水专业", "人员C_实际", 1),
            ("暖通专业", "人员E_实际", 1),
            ("岩土基础", "人员C_实际", 1),
            ("暖通基础", "人员E_实际", 1),
        ],
        "9月6日": [
            ("给排水专业", "人员C_实际", 1),
            ("环保专业", "其他人员_实际", 1),
            ("岩土专业", "其他人员_实际", 1),
        ],
        "9月7日": [
            ("电气基础", "人员A_实际", 2),
            ("环保基础", "人员A_实际", 1),
            ("水基础", "人员A_实际", 1),
            ("环保基础", "人员B_实际", 1),
            ("暖通基础", "人员B_实际", 1),
            ("岩土基础", "人员C_实际", 1),
            ("暖通专业", "人员C_实际", 1),
            ("公共基础", "人员D_实际", 1),
            ("环保基础", "人员D_实际", 1),
        ],
    }

if "base_targets" not in st.session_state or "daily_deltas" not in st.session_state:
    init_default_data()

# -----------------------------------------------------------------------------
# 3. 侧边栏设置与录入
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ 数据管理与设置")
st.sidebar.subheader("➕ 录入单日新增招生")

with st.sidebar.form("add_delta_form"):
    input_date = st.selectbox("选择日期", DATES)
    input_major = st.selectbox("选择专业/基础", [m[0] for m in MAJORS])
    input_person = st.selectbox("选择归属人员", PERSONS + ["其他人员"])
    input_val = st.number_input("新增人数", min_value=1, value=1, step=1)

    submit_btn = st.form_submit_button("提交录入")
    if submit_btn:
        target_col = f"{input_person}_实际"
        if input_date not in st.session_state["daily_deltas"]:
            st.session_state["daily_deltas"][input_date] = []
        st.session_state["daily_deltas"][input_date].append(
            (input_major, target_col, int(input_val))
        )
        st.sidebar.success(f"已成功添加：{input_date} {input_major} - {input_person} +{input_val}人")

if st.sidebar.button("🔄 重置为默认数据"):
    init_default_data()
    st.sidebar.info("数据已重置！")

# -----------------------------------------------------------------------------
# 4. 数据视图与计算
# -----------------------------------------------------------------------------
def get_processed_df(selected_date):
    df_result = st.session_state["base_targets"].copy()
    if selected_date == "📅 当月累计数据（截至9月7日）":
        for d in DATES:
            for major, col, val in st.session_state["daily_deltas"].get(d, []):
                df_result.loc[df_result["专业/基础名称"] == major, col] += val
    else:
        for major, col, val in st.session_state["daily_deltas"].get(selected_date, []):
            df_result.loc[df_result["专业/基础名称"] == major, col] += val
    return df_result

st.subheader("🗓️ 表格视图选择")
view_options = [f"{d}单日" for d in DATES] + ["📅 当月累计数据（截至9月7日）"]
date_option = st.radio("切换时间范围：", view_options, index=len(view_options) - 1, horizontal=True)

raw_date_name = date_option.replace("单日", "")
calc_df = get_processed_df(raw_date_name)

act_cols = [c for c in calc_df.columns if c.endswith("_实际")]
calc_df["实际完成"] = calc_df[act_cols].sum(axis=1)
calc_df["与目标之差"] = calc_df["实际完成"] - calc_df["目标人数"]

# 底部汇总逻辑
num_cols = [c for c in calc_df.columns if c not in ["序号", "专业/基础名称"]]
sum_row = {"专业/基础名称": "合计"}
for c in num_cols:
    sum_row[c] = int(calc_df[c].sum())

diff_row = {"专业/基础名称": "与目标之差"}
for p in PERSONS:
    diff_row[f"{p}_目标"] = ""
    diff_row[f"{p}_实际"] = sum_row[f"{p}_实际"] - sum_row[f"{p}_目标"]
diff_row["其他人员_实际"] = ""
diff_row["目标人数"] = ""
diff_row["实际完成"] = ""
diff_row["与目标之差"] = sum_row["与目标之差"]

total_target_cum = sum_row["目标人数"]
total_actual_cum = sum_row["实际完成"]
cum_rate_val = (total_actual_cum / total_target_cum * 100) if total_target_cum > 0 else 0

rate_row = {"专业/基础名称": "2026年9月目标人数完成比例"}
for c in num_cols:
    rate_row[c] = ""
rate_row["实际完成"] = f"{cum_rate_val:.2f}%"

# -----------------------------------------------------------------------------
# 5. 原生 HTML 表格渲染
# -----------------------------------------------------------------------------
def build_html_document(df, sum_r, diff_r, rate_r):
    rows_html = ""
    for idx, row in df.iterrows():
        rows_html += f"""
        <tr>
            <td class="num">{row['序号']}</td>
            <td class="zh">{row['专业/基础名称']}</td>
            <td class="num">{row['目标人数']}</td>
            <td class="num">{row['人员A_目标']}</td><td class="num">{row['人员A_实际']}</td>
            <td class="num">{row['人员B_目标']}</td><td class="num">{row['人员B_实际']}</td>
            <td class="num">{row['人员C_目标']}</td><td class="num">{row['人员C_实际']}</td>
            <td class="num">{row['人员D_目标']}</td><td class="num">{row['人员D_实际']}</td>
            <td class="num">{row['人员E_目标']}</td><td class="num">{row['人员E_实际']}</td>
            <td class="num">{row['其他人员_实际']}</td>
            <td class="num">{row['目标人数']}</td>
            <td class="num">{row['实际完成']}</td>
            <td class="num">{row['与目标之差']}</td>
        </tr>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ margin: 0; padding: 0; font-family: SimSun, "Times New Roman", serif; background-color: #ffffff; }}
        .table-container {{ width: 100%; overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; text-align: center; border: 1px solid #A6A6A6; }}
        th, td {{ border: 1px solid #A6A6A6; padding: 6px 4px; font-weight: normal; color: #000000; }}
        .bg-title {{ background-color: #D9EAD3; font-size: 15px; font-weight: normal; }}
        .bg-header {{ background-color: #EFEFEF; }}
        .bg-person {{ background-color: #D0E0E3; }}
        .bg-total {{ background-color: #FFF2CC; }}
        .num {{ font-family: "Times New Roman", serif; }}
        .zh {{ font-family: SimSun, serif; }}
    </style>
    </head>
    <body>
    <div class="table-container">
    <table>
        <tr><td colspan="17" class="bg-title">2026年9月招生数据动态表(2026年9月1日-9月7日)</td></tr>
        <tr>
            <th rowspan="2" class="bg-header">序号</th>
            <th rowspan="2" class="bg-header">专业/基础名称</th>
            <th rowspan="2" class="bg-header">目标人数</th>
            <th colspan="2" class="bg-person">人员A</th>
            <th colspan="2" class="bg-person">人员B</th>
            <th colspan="2" class="bg-person">人员C</th>
            <th colspan="2" class="bg-person">人员D</th>
            <th colspan="2" class="bg-person">人员E</th>
            <th rowspan="2" class="bg-header">其他人员</th>
            <th rowspan="2" class="bg-header">目标人数</th>
            <th rowspan="2" class="bg-header">实际完成</th>
            <th rowspan="2" class="bg-header">与目标之差</th>
        </tr>
        <tr>
            <th class="bg-header">目标人数</th><th class="bg-header">实际完成</th>
            <th class="bg-header">目标人数</th><th class="bg-header">实际完成</th>
            <th class="bg-header">目标人数</th><th class="bg-header">实际完成</th>
            <th class="bg-header">目标人数</th><th class="bg-header">实际完成</th>
            <th class="bg-header">目标人数</th><th class="bg-header">实际完成</th>
        </tr>
        {rows_html}
        <tr>
            <td class="bg-total"></td><td class="bg-total zh">{sum_r['专业/基础名称']}</td><td class="bg-total num">{sum_r['目标人数']}</td>
            <td class="bg-total num">{sum_r['人员A_目标']}</td><td class="bg-total num">{sum_r['人员A_实际']}</td>
            <td class="bg-total num">{sum_r['人员B_目标']}</td><td class="bg-total num">{sum_r['人员B_实际']}</td>
            <td class="bg-total num">{sum_r['人员C_目标']}</td><td class="bg-total num">{sum_r['人员C_实际']}</td>
            <td class="bg-total num">{sum_r['人员D_目标']}</td><td class="bg-total num">{sum_r['人员D_实际']}</td>
            <td class="bg-total num">{sum_r['人员E_目标']}</td><td class="bg-total num">{sum_r['人员E_实际']}</td>
            <td class="bg-total num">{sum_r['其他人员_实际']}</td>
            <td class="bg-total num">{sum_r['目标人数']}</td>
            <td class="bg-total num">{sum_r['实际完成']}</td>
            <td class="bg-total num">{sum_r['与目标之差']}</td>
        </tr>
        <tr>
            <td class="bg-total"></td><td class="bg-total zh">{diff_r['专业/基础名称']}</td><td class="bg-total"></td>
            <td class="bg-total"></td><td class="bg-total num">{diff_r['人员A_实际']}</td>
            <td class="bg-total"></td><td class="bg-total num">{diff_r['人员B_实际']}</td>
            <td class="bg-total"></td><td class="bg-total num">{diff_r['人员C_实际']}</td>
            <td class="bg-total"></td><td class="bg-total num">{diff_r['人员D_实际']}</td>
            <td class="bg-total"></td><td class="bg-total num">{diff_r['人员E_实际']}</td>
            <td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td>
            <td class="bg-total num">{diff_r['与目标之差']}</td>
        </tr>
        <tr>
            <td class="bg-total"></td><td class="bg-total zh">{rate_r['专业/基础名称']}</td><td class="bg-total"></td>
            <td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td><td class="bg-total"></td>
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

st.markdown("### 📝 2026年9月招生数据动态表")
html_code = build_html_document(calc_df, sum_row, diff_row, rate_row)
components.html(html_code, height=680, scrolling=True)

# -----------------------------------------------------------------------------
# 6. 🔥 强效升级：多维动态可视化分析（全体/个人/多人 x 日/星期/月 x 饼图/折线图）
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("### 📈 深度多维动态图表分析系统")

# 1. 交互控制面板
ctrl_c1, ctrl_c2, ctrl_c3 = st.columns(3)

with ctrl_c1:
    person_mode = st.radio("选择分析人员范围：", ["全体人员", "单人独立分析", "多人员对比分析"], horizontal=True)

with ctrl_c2:
    if person_mode == "单人独立分析":
        selected_person = st.selectbox("选择具体人员：", PERSONS + ["其他人员"])
    elif person_mode == "多人员对比分析":
        selected_persons = st.multiselect("选择要对比的人员：", PERSONS + ["其他人员"], default=["人员A", "人员B", "人员C"])
    else:
        selected_person = "全体人员"

with ctrl_c3:
    time_granularity = st.selectbox("选择时间维度：", ["按日期 (9月1日-9月7日)", "按星期 (星期二-星期一)", "按月份视角 (历史与当月趋势)"])

# 2. 构建明细时间序列数据 DataFrame
time_records = []
for d_idx, d in enumerate(DATES):
    w = WEEKDAYS[d_idx]
    # 初始化当日每人数据
    d_dict = {p: 0 for p in PERSONS + ["其他人员"]}
    for major, col, val in st.session_state["daily_deltas"].get(d, []):
        p_name = col.replace("_实际", "")
        d_dict[p_name] += val
    
    for p_name, val in d_dict.items():
        time_records.append({
            "日期": d,
            "星期": w,
            "月份": "2026年9月",
            "人员": p_name,
            "新增报名数": val
        })

df_time_series = pd.DataFrame(time_records)

# 3. 动态绘制折线图与饼图
chart_col1, chart_col2 = st.columns(2)

# --- 左侧：趋势折线图 ---
with chart_col1:
    x_col = "日期" if "日期" in time_granularity else ("星期" if "星期" in time_granularity else "月份")
    
    if person_mode == "全体人员":
        df_chart_line = df_time_series.groupby(x_col)["新增报名数"].sum().reset_index()
        fig_line = px.line(
            df_chart_line, x=x_col, y="新增报名数", markers=True,
            title=f"【全体人员】报名趋势折线图 ({time_granularity})",
            text="新增报名数"
        )
        fig_line.update_traces(textposition="top center", line_color="#2E8B57", line_width=3)
        
    elif person_mode == "单人独立分析":
        df_chart_line = df_time_series[df_time_series["人员"] == selected_person].groupby(x_col)["新增报名数"].sum().reset_index()
        fig_line = px.line(
            df_chart_line, x=x_col, y="新增报名数", markers=True,
            title=f"【{selected_person}】个人报名趋势折线图 ({time_granularity})",
            text="新增报名数"
        )
        fig_line.update_traces(textposition="top center", line_color="#1F77B4", line_width=3)
        
    else: # 多人员对比
        df_chart_line = df_time_series[df_time_series["人员"].isin(selected_persons)].groupby([x_col, "人员"])["新增报名数"].sum().reset_index()
        fig_line = px.line(
            df_chart_line, x=x_col, y="新增报名数", color="人员", markers=True,
            title=f"【多人员对比】报名趋势折线图 ({time_granularity})"
        )
        fig_line.update_traces(line_width=2.5)

    fig_line.update_layout(yaxis_title="新增报名人数", xaxis_title=x_col)
    st.plotly_chart(fig_line, use_container_width=True)

# --- 右侧：比例/构成饼状图 ---
with chart_col2:
    if person_mode == "全体人员":
        # 显示全体人员业绩贡献占比
        df_pie = df_time_series.groupby("人员")["新增报名数"].sum().reset_index()
        fig_pie = px.pie(
            df_pie, values="新增报名数", names="人员",
            title="【全体人员】累计招生贡献占比饼图",
            hole=0.3, color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_pie.update_traces(textinfo="label+percent+value")
        
    elif person_mode == "单人独立分析":
        # 显示该人员在各个专业/基础上的分布占比
        major_records = []
        for d in DATES:
            for major, col, val in st.session_state["daily_deltas"].get(d, []):
                p_name = col.replace("_实际", "")
                if p_name == selected_person:
                    major_records.append({"专业/基础": major, "新增人数": val})
        
        df_major_pie = pd.DataFrame(major_records)
        if not df_major_pie.empty:
            df_major_pie = df_major_pie.groupby("专业/基础")["新增人数"].sum().reset_index()
            fig_pie = px.pie(
                df_major_pie, values="新增人数", names="专业/基础",
                title=f"【{selected_person}】招生专业分布构成饼图",
                hole=0.3, color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig_pie.update_traces(textinfo="label+percent+value")
        else:
            fig_pie = go.Figure()
            fig_pie.update_layout(title=f"【{selected_person}】暂无招生数据录入")
            
    else: # 多人员对比
        # 显示选中人员的总招生量占比
        df_pie = df_time_series[df_time_series["人员"].isin(selected_persons)].groupby("人员")["新增报名数"].sum().reset_index()
        fig_pie = px.pie(
            df_pie, values="新增报名数", names="人员",
            title=f"【选中多人员】业绩总量构成占比饼图",
            hole=0.3, color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig_pie.update_traces(textinfo="label+percent+value")

    st.plotly_chart(fig_pie, use_container_width=True)

# -----------------------------------------------------------------------------
# 7. Excel 导出功能
# -----------------------------------------------------------------------------
def export_color_excel(calc_df, sum_row, diff_row, rate_row, persons):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "招生数据动态表"
    ws.views.sheetView[0].showGridLines = True

    TITLE_FILL = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    HEADER_BG = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    PERSON_BG = PatternFill(start_color="D0E0E3", end_color="D0E0E3", fill_type="solid")
    TOTAL_ROW_BG = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    FONT_TITLE = Font(name="SimSun", size=14)
    FONT_HEADER = Font(name="SimSun", size=10)
    FONT_BODY_NUM = Font(name="Times New Roman", size=10)
    FONT_BODY_ZH = Font(name="SimSun", size=10)

    BORDER_THIN = Border(
        left=Side(style="thin", color="A6A6A6"),
        right=Side(style="thin", color="A6A6A6"),
        top=Side(style="thin", color="A6A6A6"),
        bottom=Side(style="thin", color="A6A6A6"),
    )
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:Q1")
    t_cell = ws["A1"]
    t_cell.value = "2026年9月招生数据动态表(2026年9月1日-9月7日)"
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
    for p in persons:
        ws.merge_cells(start_row=2, start_column=col_idx, end_row=2, end_column=col_idx + 1)
        p_cell = ws.cell(row=2, column=col_idx, value=p)
        p_cell.fill = PERSON_BG
        ws.cell(row=3, column=col_idx, value="目标人数")
        ws.cell(row=3, column=col_idx + 1, value="实际完成")
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
            cell.font = FONT_HEADER
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN

    curr_r = 4
    for idx, row in calc_df.iterrows():
        ws.cell(row=curr_r, column=1, value=row["序号"]).font = FONT_BODY_NUM
        ws.cell(row=curr_r, column=2, value=row["专业/基础名称"]).font = FONT_BODY_ZH
        ws.cell(row=curr_r, column=3, value=row["目标人数"]).font = FONT_BODY_NUM

        c_offset = 4
        for p in persons:
            ws.cell(row=curr_r, column=c_offset, value=row[f"{p}_目标"]).font = FONT_BODY_NUM
            ws.cell(row=curr_r, column=c_offset + 1, value=row[f"{p}_实际"]).font = FONT_BODY_NUM
            c_offset += 2

        ws.cell(row=curr_r, column=c_offset, value=row.get("避免人员_实际", row.get("其他人员_实际", 0))).font = FONT_BODY_NUM
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
        for p in persons:
            ws.cell(row=curr_r, column=c_offset, value=r_data.get(f"{p}_目标", "")).font = FONT_HEADER
            ws.cell(row=curr_r, column=c_offset + 1, value=r_data.get(f"{p}_实际", "")).font = FONT_HEADER
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

st.markdown("---")
excel_bytes = export_color_excel(calc_df, sum_row, diff_row, rate_row, PERSONS)

st.download_button(
    label="📊 导出为原版颜色 Excel 表格 (.xlsx)",
    data=excel_bytes,
    file_name="2026年9月招生数据动态表_带颜色版.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
