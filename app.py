import streamlit as st
import pandas as pd
import plotly.express as px

# -----------------------------------------------------------------------------
# 1. 页面基本配置
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="招生数据趋势分析看板",
    page_icon="📊",
    layout="wide"
)

# 定义预设的时间顺序列表（用于约束 Plotly 绘图的横坐标时间轴顺序）
DATES = [
    "9月1日", "9月2日", "9月3日", "9月4日", "9月5日", "9月6日",
    "9月7日", "9月8日", "9月9日", "9月10日", "9月11日", "9月12日", "9月13日"
]

# -----------------------------------------------------------------------------
# 2. 模拟/初始化数据源
# -----------------------------------------------------------------------------
@st.cache_data
def get_raw_data():
    return {
        "9月1日": [("电气基础", "张工", 1)],
        "9月2日": [("电气基础", "张工", 2), ("环评实战", "李工", 3)],
        "9月3日": [("环评实战", "李工", 4)],
        "9月4日": [("电气基础", "张工", 2), ("消防攻坚", "王工", 2)],
        "9月5日": [("环评实战", "李工", 4)],
        "9月6日": [("电气基础", "张工", 1), ("消防攻坚", "王工", 2)],
        "9月7日": [("电气基础", "张工", 5), ("环评实战", "李工", 7)],
        "9月8日": [("消防攻坚", "王工", 1)],
        "9月9日": [("电气基础", "张工", 2), ("环评实战", "覃小燕", 2)],
        "9月10日": [("环评实战", "覃小燕", 9)],
        "9月11日": [("电气基础", "张工", 4)],
        "9月12日": [("电气基础", "覃小燕", 1)],
        "9月13日": [("环评实战", "覃小燕", 6), ("消防攻坚", "王工", 4)]
    }

# 解析数据字典转换为标准 DataFrame
rows = []
for date, items in get_raw_data().items():
    for item in items:
        if len(item) == 3:
            course, person, count = item
            rows.append({"日期": date, "课程": course, "人员": person, "新增报名数": count})

df_time_series = pd.DataFrame(rows)

# -----------------------------------------------------------------------------
# 3. 页面标题与侧边栏控制
# -----------------------------------------------------------------------------
st.title("📊 全体人员招生趋势分析看板")
st.markdown("---")

st.sidebar.header("🔍 分析维度与筛选")

# 日期筛选框
selected_dates = st.sidebar.multiselect(
    "选择显示日期",
    options=DATES,
    default=DATES
)

# 人员模式选择
person_mode = st.sidebar.radio(
    "人员分析模式",
    options=["全体人员", "单人独立分析", "多人招生趋势对比"]
)

all_persons = sorted(df_time_series["人员"].unique().tolist())
selected_person_disp = None
selected_persons_disp = []

if person_mode == "单人独立分析":
    selected_person_disp = st.sidebar.selectbox("选择分析人员", options=all_persons)
elif person_mode == "多人招生趋势对比":
    selected_persons_disp = st.sidebar.multiselect("选择对比人员", options=all_persons, default=all_persons[:2])

# 根据侧边栏日期过滤数据
df_filtered = df_time_series[df_time_series["日期"].isin(selected_dates)].copy()

# -----------------------------------------------------------------------------
# 4. KPI 核心指标显示
# -----------------------------------------------------------------------------
st.subheader("📌 核心数据概览")
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

total_enrollment = df_filtered["新增报名数"].sum()
active_days = df_filtered["日期"].nunique()
avg_per_day = round(total_enrollment / active_days, 1) if active_days > 0 else 0
top_person = df_filtered.groupby("人员")["新增报名数"].sum().idxmax() if not df_filtered.empty else "无"

kpi1.metric("总新增报名人数", f"{total_enrollment} 人")
kpi2.metric("统计有效天数", f"{active_days} 天")
kpi3.metric("日均新增报名", f"{avg_per_day} 人")
kpi4.metric("最佳表现人员", top_person)

st.markdown("---")

# -----------------------------------------------------------------------------
# 5. 图表呈现区域
# -----------------------------------------------------------------------------
chart_col1, chart_col2 = st.columns([3, 2])

with chart_col1:
    if person_mode == "全体人员":
        # 按日汇总
        df_chart_line = df_filtered.groupby("日期", as_index=False)["新增报名数"].sum()
        
        # 【核心修复】：转换为 Categorical 字段并按 DATES 自然时间序列严格排序，彻底解决折线首尾穿梭/连线错乱问题
        df_chart_line["日期"] = pd.Categorical(df_chart_line["日期"], categories=DATES, ordered=True)
        df_chart_line = df_chart_line.sort_values("日期").reset_index(drop=True)
        
        fig_line = px.line(
            df_chart_line, x="日期", y="新增报名数", markers=True, 
            title="📈 <b>全体人员招生趋势 (按日明细)</b>", text="新增报名数"
        )
        fig_line.update_traces(
            textposition="top center", 
            line_color="#D50000", 
            line_width=2.5, 
            marker=dict(size=7, color="#D50000")
        )

    elif person_mode == "单人独立分析":
        df_sub = df_filtered[df_filtered["人员"] == selected_person_disp]
        df_chart_line = df_sub.groupby("日期", as_index=False)["新增报名数"].sum()
        
        df_chart_line["日期"] = pd.Categorical(df_chart_line["日期"], categories=DATES, ordered=True)
        df_chart_line = df_chart_line.sort_values("日期").reset_index(drop=True)
        
        fig_line = px.line(
            df_chart_line, x="日期", y="新增报名数", markers=True, 
            title=f"📈 <b>【{selected_person_disp}】招生趋势 (按日明细)</b>", text="新增报名数"
        )
        fig_line.update_traces(
            textposition="top center", 
            line_color="#2962FF", 
            line_width=2.5, 
            marker=dict(size=7, color="#2962FF")
        )

    else:
        df_sub = df_filtered[df_filtered["人员"].isin(selected_persons_disp)]
        df_chart_line = df_sub.groupby(["日期", "人员"], as_index=False)["新增报名数"].sum()
        
        df_chart_line["日期"] = pd.Categorical(df_chart_line["日期"], categories=DATES, ordered=True)
        df_chart_line = df_chart_line.sort_values(["日期", "人员"]).reset_index(drop=True)
        
        fig_line = px.line(
            df_chart_line, x="日期", y="新增报名数", color="人员", markers=True, 
            title="📈 <b>多人招生趋势对比 (按日明细)</b>",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_line.update_traces(line_width=2.5, marker=dict(size=7))

    fig_line.update_layout(
        yaxis_title="新增报名人数", 
        xaxis_title="日期", 
        hovermode="x unified",
        plot_bgcolor="#FFFFFF", 
        paper_bgcolor="#FFFFFF",
        margin=dict(l=20, r=20, t=50, b=20),
        yaxis=dict(gridcolor="#E0E0E0")
    )
    st.plotly_chart(fig_line, use_container_width=True)

with chart_col2:
    # 课程转化柱状图
    df_course = df_filtered.groupby("课程", as_index=False)["新增报名数"].sum().sort_values("新增报名数", ascending=True)
    fig_bar = px.bar(
        df_course, x="新增报名数", y="课程", orientation='h',
        title="📚 <b>课程招生分布明细</b>", text="新增报名数",
        color="新增报名数", color_continuous_scale="Reds"
    )
    fig_bar.update_traces(textposition="outside")
    fig_bar.update_layout(
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        margin=dict(l=20, r=20, t=50, b=20),
        xaxis=dict(gridcolor="#E0E0E0")
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# -----------------------------------------------------------------------------
# 6. 明细数据表与导出功能
# -----------------------------------------------------------------------------
st.markdown("---")
st.subheader("📋 详细招生明细数据表")
st.dataframe(df_filtered, use_container_width=True, hide_index=True)

# 导出 CSV
csv = df_filtered.to_csv(index=False, encoding='utf-8-sig')
st.download_button(
    label="📥 下载当前视图数据 (CSV)",
    data=csv,
    file_name="招生明细数据.csv",
    mime="text/csv"
)
