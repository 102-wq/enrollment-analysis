import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# ==============================================================================
# 1. 页面基础配置 (Page Configuration)
# ==============================================================================
st.set_page_config(
    page_title="招生数据智能分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 全局 CSS 样式注入
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: bold;
        color: #1E293B;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .custom-table {
        width: 100%;
        border-collapse: collapse;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        font-size: 14px;
        margin-top: 10px;
        margin-bottom: 20px;
    }
    .custom-table th, .custom-table td {
        border: 1px solid #CBD5E1;
        padding: 10px 12px;
        text-align: center;
    }
    .custom-table th {
        background-color: #F1F5F9;
        color: #334155;
        font-weight: 600;
    }
    .custom-table tr:nth-child(even) {
        background-color: #F8FAFC;
    }
    .highlight-row {
        background-color: #E2E8F0 !important;
        font-weight: bold;
        color: #0F172A;
    }
    .rate-row {
        background-color: #EFF6FF !important;
        color: #1D4ED8;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. 默认常量与 Session State 初始化
# ==============================================================================
RAW_PERSONS = ["覃小燕", "张三", "李四", "王五"]
MAJORS = ["环境影响评价工程师", "一级建造师", "注册安全工程师", "消防工程师"]
DATES = ["9月1日", "9月2日", "9月3日", "9月4日", "9月5日", "9月6日", "9月7日"]

# 侧边栏：脱敏与配置
st.sidebar.title("⚙️ 系统配置")
enable_anonymize = st.sidebar.checkbox("开启数据脱敏 / 匿名模式", value=False)

# 动态构建人员别名映射表
alias_map = {}
for i, name in enumerate(RAW_PERSONS):
    alias_map[name] = f"咨询顾问 {chr(65 + i)}" if enable_anonymize else name

DISPLAY_PERSONS = [alias_map[p] for p in RAW_PERSONS]

# 初始化基础数据结构（存储原始列名，保持底层数据稳固）
if "base_targets" not in st.session_state:
    data = []
    for major in MAJORS:
        row = {"专业/基础名称": major, "总目标": 100}
        for p in RAW_PERSONS:
            row[f"{p}_目标"] = 25
            row[f"{p}_实际"] = 0
        row["其他人员_实际"] = 0
        data.append(row)
    st.session_state["base_targets"] = pd.DataFrame(data)

# 初始化每日增量记录字典
if "daily_deltas" not in st.session_state:
    st.session_state["daily_deltas"] = {d: [] for d in DATES}

# ==============================================================================
# 3. 核心计算函数 (动态计算与完全容错)
# ==============================================================================
def get_calculated_df(selected_date):
    """根据选定日期动态叠加每日增量，并安全处理列名脱敏映射"""
    df_calc = st.session_state["base_targets"].copy()
    
    # 确保存储的基础列都存在（防 Key 缺失）
    for p in RAW_PERSONS:
        if f"{p}_目标" not in df_calc.columns:
            df_calc[f"{p}_目标"] = 0
        if f"{p}_实际" not in df_calc.columns:
            df_calc[f"{p}_实际"] = 0
    if "其他人员_实际" not in df_calc.columns:
        df_calc["其他人员_实际"] = 0

    # 叠加增量数据
    target_dates = DATES if selected_date == "📅 当月累计数据（截至9月7日）" else [selected_date]
    for d in target_dates:
        for major, col_raw, val in st.session_state["daily_deltas"].get(d, []):
            if col_raw in df_calc.columns:
                df_calc.loc[df_calc["专业/基础名称"] == major, col_raw] += val

    # 构建展示用 DataFrame（安全映射脱敏列名）
    df_display = pd.DataFrame()
    df_display["专业/基础名称"] = df_calc["专业/基础名称"]
    df_display["总目标"] = df_calc["总目标"]

    for raw_p in RAW_PERSONS:
        disp_p = alias_map[raw_p]
        df_display[f"{disp_p}_目标"] = df_calc.get(f"{raw_p}_目标", 0)
        df_display[f"{disp_p}_实际"] = df_calc.get(f"{raw_p}_实际", 0)

    df_display["其他人员_实际"] = df_calc.get("其他人员_实际", 0)
    return df_display

def build_custom_html_table(df, sum_row, rate_row):
    """渲染原生 HTML 表格，保证表格视觉统一与无缝响应"""
    html = """<table class="custom-table"><thead><tr>
    <th>专业/基础名称</th><th>总目标</th>"""
    
    for p in DISPLAY_PERSONS:
        html += f"<th>{p} (目标)</th><th>{p} (实际)</th>"
    html += "<th>其他人员 (实际)</th></tr></thead><tbody>"

    for _, row in df.iterrows():
        html += f"<tr><td>{row['专业/基础名称']}</td><td>{row['总目标']}</td>"
        for p in DISPLAY_PERSONS:
            html += f"<td>{row[f'{p}_目标']}</td><td>{row[f'{p}_实际']}</td>"
        html += f"<td>{row['其他人员_实际']}</td></tr>"

    # 合计行
    html += f'<tr class="highlight-row"><td>{sum_row["专业/基础名称"]}</td><td>{sum_row["总目标"]}</td>'
    for p in DISPLAY_PERSONS:
        html += f'<td>{sum_row[f"{p}_目标"]}</td><td>{sum_row[f"{p}_实际"]}</td>'
    html += f'<td>{sum_row["其他人员_实际"]}</td></tr>'

    # 达成率行
    html += f'<tr class="rate-row"><td>{rate_row["专业/基础名称"]}</td><td>{rate_row["总目标"]}</td>'
    for p in DISPLAY_PERSONS:
        html += f'<td>-</td><td>{rate_row[f"{p}_实际"]}</td>'
    html += f'<td>-</td></tr>'

    html += "</tbody></table>"
    return html

# ==============================================================================
# 4. 主界面渲染与交互
# ==============================================================================
st.markdown('<div class="main-header">📊 招生数据智能分析系统</div>', unsafe_allow_html=True)

# 时间节点选择器
date_options = ["📅 当月累计数据（截至9月7日）"] + DATES
selected_date = st.selectbox("📅 选择查看的时间节点：", date_options)

# 计算当前视图数据
current_df = get_calculated_df(selected_date)

# 计算汇总与达成率指标
sum_series = {"专业/基础名称": "合计", "总目标": int(current_df["总目标"].sum())}
rate_series = {"专业/基础名称": "达成率", "总目标": "100%"}

grand_total_actual = 0
for p in DISPLAY_PERSONS:
    t_val = int(current_df[f"{p}_目标"].sum())
    a_val = int(current_df[f"{p}_实际"].sum())
    sum_series[f"{p}_目标"] = t_val
    sum_series[f"{p}_实际"] = a_val
    rate_series[f"{p}_实际"] = f"{(a_val / t_val * 100):.1f}%" if t_val > 0 else "0.0%"
    grand_total_actual += a_val

other_actual = int(current_df["其他人员_实际"].sum())
sum_series["其他人员_实际"] = other_actual
grand_total_actual += other_actual

# 1. 顶部核心 KPI 指标卡
kpi1, kpi2, kpi3 = st.columns(3)
with kpi1:
    st.metric(label="月度总目标", value=f"{sum_series['总目标']} 人")
with kpi2:
    st.metric(label="当前累计完成", value=f"{grand_total_actual} 人")
with kpi3:
    overall_rate = (grand_total_actual / sum_series['总目标'] * 100) if sum_series['总目标'] > 0 else 0
    st.metric(label="总体完成率", value=f"{overall_rate:.1f}%")

st.markdown("---")

# 2. 核心数据表格
st.markdown("### 📝 招生数据动态汇总表")
html_code = build_custom_html_table(current_df, sum_series, rate_series)
st.html(html_code)

# 3. 统计图表展现
st.markdown("### 📈 招生进度图表分析")
chart_tab1, chart_tab2 = st.tabs(["人员完成度对比", "专业目标分布"])

with chart_tab1:
    person_chart_data = []
    for p in DISPLAY_PERSONS:
        person_chart_data.append({"顾问": p, "指标类型": "目标", "人数": sum_series[f"{p}_目标"]})
        person_chart_data.append({"顾问": p, "指标类型": "实际", "人数": sum_series[f"{p}_实际"]})
    df_person_chart = pd.DataFrame(person_chart_data)
    
    fig_person = px.bar(
        df_person_chart, 
        x="顾问", 
        y="人数", 
        color="指标类型", 
        barmode="group",
        text_auto=True,
        color_discrete_map={"目标": "#94A3B8", "实际": "#2563EB"}
    )
    fig_person.update_layout(margin=dict(l=20, r=20, t=30, b=20), height=380)
    st.plotly_chart(fig_person, width="stretch")

with chart_tab2:
    fig_major = px.bar(
        current_df, 
        x="专业/基础名称", 
        y="总目标", 
        title="各专业月度目标分布",
        text_auto=True,
        color_discrete_sequence=["#0EA5E9"]
    )
    fig_major.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=380)
    st.plotly_chart(fig_major, width="stretch")

st.markdown("---")

# 4. 增量录入与基础配置修改
st.markdown("### ✏️ 快速录入日增量数据")
with st.expander("点击展开增量数据录入窗口", expanded=False):
    col1, col2, col3 = st.columns(3)
    with col1:
        entry_date = st.selectbox("录入日期", DATES)
    with col2:
        entry_major = st.selectbox("选择专业", MAJORS)
    with col3:
        entry_person = st.selectbox("责任顾问", DISPLAY_PERSONS)
        
    entry_val = st.number_input("今日新增招生人数", min_value=0, value=1, step=1)
    
    if st.button("提交增量数据", width="stretch"):
        raw_person_name = RAW_PERSONS[DISPLAY_PERSONS.index(entry_person)]
        target_col_raw = f"{raw_person_name}_实际"
        
        st.session_state["daily_deltas"][entry_date].append((entry_major, target_col_raw, entry_val))
        st.success(f"已成功录入：{entry_date} | {entry_major} | {entry_person} +{entry_val}人！")
        st.rerun()

st.markdown("### ⚙️ 基础目标配置与修改")
edited_df = st.data_editor(
    st.session_state["base_targets"],
    key="target_editor",
    width="stretch"
)

if st.button("保存目标配置修改", width="stretch"):
    st.session_state["base_targets"] = edited_df
    st.success("基础目标配置已成功更新！")
    st.rerun()
