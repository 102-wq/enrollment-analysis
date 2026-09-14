import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 1. 页面基本配置与全局 UI 样式
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="招生数据动态管理与多维分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .kpi-card {
        background-color: #F8F9FA;
        border: 1px solid #E9ECEF;
        border-radius: 10px;
        padding: 16px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    }
    .stProgress > div > div > div > div {
        background-color: #2E7D32;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 招生数据动态管理与多维分析系统")
st.caption("2026年9月1日 - 9月13日 全量精准数据对齐版")
st.markdown("---")

# -----------------------------------------------------------------------------
# 2. 基础数据定义与 Session State 初始化
# -----------------------------------------------------------------------------
DEFAULT_MAJORS = [
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
    ("结构基础", 9, [2, 2, 1, 2, 1]),
    ("公共基础", 4, [2, 1, 1, 0, 1]),
    ("道路基础", 6, [1, 1, 1, 1, 2]),
    ("水利水电基础", 10, [2, 2, 3, 2, 1]),
]

DATES = [
    "9月1日", "9月2日", "9月3日", "9月4日", "9月5日", "9月6日", "9月7日",
    "9月8日", "9月9日", "9月10日", "9月11日", "9月12日", "9月13日"
]

def init_default_data():
    st.session_state["raw_persons"] = ["覃小燕", "左丹丹", "梁书华", "古晨晓", "周欢喜"]
    st.session_state["person_animals"] = {
        "覃小燕": "🦊", "左丹丹": "🐼", "梁书华": "🦁", "古晨晓": "🐰", "周欢喜": "🐯"
    }

    base_data = []
    for idx, (name, total_target, person_tgts) in enumerate(DEFAULT_MAJORS, 1):
        row = {"序号": idx, "专业/基础名称": name, "目标人数": total_target}
        for p, tgt in zip(st.session_state["raw_persons"], person_tgts):
            row[f"{p}_目标"] = tgt
            row[f"{p}_实际"] = 0
        row["其他人员_实际"] = 0
        base_data.append(row)
    st.session_state["base_targets"] = pd.DataFrame(base_data)

    # 1-13号全量增量明细数据
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
            ("结构基础", "其他人员_实际", 2),
        ],
        "9月8日": [
            ("给排水专业", "梁书华_实际", 1),
        ],
        "9月9日": [
            ("暖通专业", "左丹丹_实际", 1),
            ("233网校", "左丹丹_实际", 1),
            ("水基础", "周欢喜_实际", 1),
            ("给排水专业", "其他人员_实际", 1),
        ],
        "9月10日": [
            ("电气基础", "覃小燕_实际", 2),
            ("环保基础", "覃小燕_实际", 1),
            ("暖通专业", "周欢喜_实际", 1),
            ("电气基础", "周欢喜_实际", 1),
            ("暖通基础", "周欢喜_实际", 1),
            ("岩土基础", "古晨晓_实际", 1),
            ("水利水电基础", "左丹丹_实际", 1),
            ("岩土基础", "梁书华_实际", 1),
        ],
        "9月11日": [
            ("电气基础", "覃小燕_实际", 1),
            ("道路基础", "覃小燕_实际", 1),
            ("岩土基础", "其他人员_实际", 1),
            ("电气基础", "梁书华_实际", 1),
        ],
        "9月12日": [
            ("电气基础", "覃小燕_实际", 1),
        ],
        "9月13日": [
            ("给排水专业", "周欢喜_实际", 1),
            ("结构专业", "其他人员_实际", 1),
            ("电气基础", "覃小燕_实际", 1),
            ("电气基础", "左丹丹_实际", 1),
            ("水基础", "覃小燕_实际", 1),
            ("水基础", "梁书华_实际", 1),
            ("暖通基础", "覃小燕_实际", 1),
            ("暖通基础", "周欢喜_实际", 1),
            ("水利水电基础", "梁书华_实际", 1),
            ("水利水电基础", "其他人员_实际", 1),
        ],
    }

if "base_targets" not in st.session_state or "raw_persons" not in st.session_state:
    init_default_data()

RAW_PERSONS = st.session_state["raw_persons"]

# -----------------------------------------------------------------------------
# 3. 侧边栏：脱敏与数据录入管理
# -----------------------------------------------------------------------------
st.sidebar.title("⚙️ 系统设置与功能菜单")

enable_anonymize = st.sidebar.checkbox("开启数据脱敏 (显示动物图标)", value=True)

alias_map = {}
for p in RAW_PERSONS:
    if enable_anonymize:
        alias_map[p] = st.session_state["person_animals"].get(p, "🐱")
    else:
        alias_map[p] = p

if enable_anonymize:
    with st.sidebar.expander("👁️ 人员与代称对照表", expanded=False):
        mapping_df = pd.DataFrame({
            "真实姓名": RAW_PERSONS,
            "代称图标": [st.session_state["person_animals"].get(p, "🐱") for p in RAW_PERSONS]
        })
        st.dataframe(mapping_df, hide_index=True, use_container_width=True)

# 快捷新增数据表单
st.sidebar.markdown("---")
st.sidebar.subheader("➕ 快捷录入每日增量")
with st.sidebar.form("add_delta_form"):
    input_date = st.selectbox("选择日期", DATES)
    input_major = st.selectbox("选择专业", st.session_state["base_targets"]["专业/基础名称"].tolist())
    input_person = st.selectbox("选择负责人", RAW_PERSONS + ["其他人员"])
    input_val = st.number_input("新增实际完成数", min_value=1, value=1, step=1)
    
    submit_btn = st.form_submit_button("提交录入")
    if submit_btn:
        col_key = f"{input_person}_实际" if input_person in RAW_PERSONS else "其他人员_实际"
        if input_date not in st.session_state["daily_deltas"]:
            st.session_state["daily_deltas"][input_date] = []
        st.session_state["daily_deltas"][input_date].append((input_major, col_key, input_val))
        st.success(f"已成功录入：{input_date} {input_major} - {input_person} +{input_val}")

# -----------------------------------------------------------------------------
# 4. 数据计算函数
# -----------------------------------------------------------------------------
def get_processed_df_by_dates(dates_list):
    df_result = st.session_state["base_targets"].copy()
    act_cols = [c for c in df_result.columns if c.endswith("_实际")]
    for col in act_cols:
        df_result[col] = 0

    for d in dates_list:
        for item in st.session_state["daily_deltas"].get(d, []):
            if len(item) == 3:
                major, col, val = item
                if col in df_result.columns:
                    df_result.loc[df_result["专业/基础名称"] == major, col] += val

    return df_result

# 日期选择筛选器
st.subheader("📅 数据时间区间选择")
selected_dates = st.multiselect("选择要查看的日期范围（可多选）", DATES, default=DATES)
if not selected_dates:
    selected_dates = DATES

calc_df = get_processed_df_by_dates(selected_dates)
act_cols = [c for c in calc_df.columns if c.endswith("_实际")]
calc_df["实际完成"] = calc_df[act_cols].sum(axis=1)
calc_df["与目标之差"] = calc_df["实际完成"] - calc_df["目标人数"]

# 重命名显示列（脱敏处理）
display_df = calc_df.copy()
for p in RAW_PERSONS:
    if f"{p}_目标" in display_df.columns:
        display_df.rename(columns={
            f"{p}_目标": f"{alias_map[p]}_目标",
            f"{p}_实际": f"{alias_map[p]}_实际"
        }, inplace=True)

# -----------------------------------------------------------------------------
# 5. KPI 核心指标卡片
# -----------------------------------------------------------------------------
total_target = calc_df["目标人数"].sum()
total_actual = calc_df["实际完成"].sum()
completion_rate = (total_actual / total_target * 100) if total_target > 0 else 0

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("🎯 总目标人数", f"{total_target} 人")
with c2:
    st.metric("✅ 实际完成人数", f"{total_actual} 人", delta=f"{total_actual - total_target} 人")
with c3:
    st.metric("📈 目标完成率", f"{completion_rate:.2f}%")
with c4:
    progress_val = min(completion_rate / 100, 1.0)
    st.write("整体进度追踪")
    st.progress(progress_val)

st.markdown("---")

# -----------------------------------------------------------------------------
# 6. 多维分析标签页
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📋 数据明细表", "📊 可视化图表", "📥 数据导出"])

with tab1:
    st.subheader("📋 招生目标与实际完成明细表")
    st.dataframe(display_df, use_container_width=True)

with tab2:
    st.subheader("📊 招生数据多维分析图表")
    
    col_chart1, col_chart2 = st.columns(2)
    
    with col_chart1:
        # 1. 各专业目标与实际对比
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(x=calc_df["专业/基础名称"], y=calc_df["目标人数"], name="目标人数", marker_color="#9E9E9E"))
        fig_bar.add_trace(go.Bar(x=calc_df["专业/基础名称"], y=calc_df["实际完成"], name="实际完成", marker_color="#2E7D32"))
        fig_bar.update_layout(title="各专业/基础 目标 vs 实际完成人数", barmode="group", xaxis_tickangle=-45)
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_chart2:
        # 2. 团队成员贡献占比
        person_contrib = {}
        for p in RAW_PERSONS:
            person_contrib[alias_map[p]] = calc_df[f"{p}_实际"].sum()
        person_contrib["其他人员"] = calc_df["其他人员_实际"].sum()
        
        contrib_df = pd.DataFrame(list(person_contrib.items()), columns=["人员/代称", "完成人数"])
        fig_pie = px.pie(contrib_df, names="人员/代称", values="完成人数", title="团队成员完成人数贡献占比", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

    # 3. 每日新增走势
    daily_totals = []
    for d in DATES:
        day_df = get_processed_df_by_dates([d])
        day_act_cols = [c for c in day_df.columns if c.endswith("_实际")]
        daily_totals.append({"日期": d, "单日新增": day_df[day_act_cols].sum().sum()})
    
    trend_df = pd.DataFrame(daily_totals)
    fig_line = px.line(trend_df, x="日期", y="单日新增", title="9月1日-13日 每日招生新增走势图", markers=True)
    st.plotly_chart(fig_line, use_container_width=True)

with tab3:
    st.subheader("📥 导出分析报告")
    st.write("点击下方按钮下载包含全量数据的 Excel 报表：")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        display_df.to_excel(writer, sheet_name="数据汇总表", index=False)
        
    st.download_button(
        label="📥 下载 Excel 汇总表格",
        data=buffer.getvalue(),
        file_name="2026年9月招生数据汇总表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
