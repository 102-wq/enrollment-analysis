import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

# -----------------------------------------------------------------------------
# 1. 页面基本配置
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="招生数据智能分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义 CSS 样式提升美观度
st.markdown("""
    <style>
    .main { padding: 1rem 2rem; }
    .stMetric {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. 数据加载与缓存
# -----------------------------------------------------------------------------
@st.cache_data(ttl=600)
def generate_sample_data():
    """生成示例数据（若用户未上传数据时使用）"""
    np.random.seed(42)
    dates = pd.date_range(end=datetime.today(), periods=90, freq='D')
    campuses = ['城北校区', '高新校区', '锦江校区']
    channels = ['线上广告', '社群裂变', '线下转介绍', '门店自然到访']
    courses = ['环评工程师通关班', '环境数据分析班', '职业技能提升班']
    
    data = []
    for d in dates:
        for c in campuses:
            consults = np.random.randint(10, 50)
            trials = int(consults * np.random.uniform(0.4, 0.7))
            enrolls = int(trials * np.random.uniform(0.3, 0.6))
            revenue = enrolls * np.random.choice([3800, 4500, 5800])
            
            data.append({
                '日期': d,
                '校区': c,
                '渠道': np.random.choice(channels),
                '课程': np.random.choice(courses),
                '咨询量': consults,
                '试听量': trials,
                '报名人数': enrolls,
                '总业绩(元)': revenue
            })
    df = pd.DataFrame(data)
    return df

def load_data():
    if 'raw_data' not in st.session_state:
        st.session_state['raw_data'] = generate_sample_data()
    return st.session_state['raw_data']

df_raw = load_data()

# -----------------------------------------------------------------------------
# 3. 侧边栏（Sidebar）与全局筛选器
# -----------------------------------------------------------------------------
st.sidebar.title("🔍 筛选与设置")

# 文件上传组件
uploaded_file = st.sidebar.file_uploader("上传您的 Excel/CSV 招生数据", type=['xlsx', 'csv'])
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_uploaded = pd.read_csv(uploaded_file)
        else:
            df_uploaded = pd.read_excel(uploaded_file)
        
        # 转换日期列
        if '日期' in df_uploaded.columns:
            df_uploaded['日期'] = pd.to_datetime(df_uploaded['日期'])
        st.session_state['raw_data'] = df_uploaded
        df_raw = df_uploaded
        st.sidebar.success("数据上传成功！")
    except Exception as e:
        st.sidebar.error(f"文件读取失败: {e}")

# 筛选维度设置
all_campuses = df_raw['校区'].unique().tolist() if '校区' in df_raw.columns else []
all_channels = df_raw['渠道'].unique().tolist() if '渠道' in df_raw.columns else []

selected_campuses = st.sidebar.multiselect("选择校区", options=all_campuses, default=all_campuses)
selected_channels = st.sidebar.multiselect("选择渠道", options=all_channels, default=all_channels)

# 日期筛选
if '日期' in df_raw.columns and not df_raw.empty:
    min_date = df_raw['日期'].min().date()
    max_date = df_raw['日期'].max().date()
    date_range = st.sidebar.date_input("日期范围", value=(min_date, max_date), min_value=min_date, max_value=max_date)
else:
    date_range = None

# 数据过滤逻辑
df_filtered = df_raw.copy()
if selected_campuses:
    df_filtered = df_filtered[df_filtered['校区'].isin(selected_campuses)]
if selected_channels:
    df_filtered = df_filtered[df_filtered['渠道'].isin(selected_channels)]
if date_range and len(date_range) == 2:
    start_d, end_d = date_range
    df_filtered = df_filtered[(df_filtered['日期'].dt.date >= start_d) & (df_filtered['日期'].dt.date <= end_d)]

# -----------------------------------------------------------------------------
# 4. 主界面：看板标题与核心 KPI
# -----------------------------------------------------------------------------
st.title("📈 招生数据智能分析与管理系统")
st.caption("实时监控招生转化、业绩指标及渠道效果")

# 关键指标卡片 (Metrics)
col1, col2, col3, col4, col5 = st.columns(5)

total_consults = df_filtered['咨询量'].sum() if '咨询量' in df_filtered.columns else 0
total_trials = df_filtered['试听量'].sum() if '试听量' in df_filtered.columns else 0
total_enrolls = df_filtered['报名人数'].sum() if '报名人数' in df_filtered.columns else 0
total_revenue = df_filtered['总业绩(元)'].sum() if '总业绩(元)' in df_filtered.columns else 0

conv_rate = (total_enrolls / total_consults * 100) if total_consults > 0 else 0

col1.metric("总咨询量", f"{total_consults:,} 人")
col2.metric("试听人数", f"{total_trials:,} 人")
col3.metric("总报名人数", f"{total_enrolls:,} 人")
col4.metric("总招生业绩", f"￥{total_revenue:,.2f}")
col5.metric("总体转化率", f"{conv_rate:.1f}%")

st.markdown("---")

# -----------------------------------------------------------------------------
# 5. 图表分析大盘 (Tab 分页)
# -----------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(["📊 招生趋势与对比", "🌪️ 转化漏斗分析", "🎯 渠道与校区分布", "📝 数据管理与编辑"])

with tab1:
    st.subheader("招生业绩与人数趋势")
    if '日期' in df_filtered.columns and not df_filtered.empty:
        # 按日期聚合
        df_trend = df_filtered.groupby('日期').agg({
            '报名人数': 'sum',
            '总业绩(元)': 'sum',
            '咨询量': 'sum'
        }).reset_index()
        
        fig_trend = px.line(
            df_trend, x='日期', y=['报名人数', '咨询量'],
            title="每日咨询与报名趋势变化",
            markers=True,
            color_discrete_sequence=['#1f77b4', '#ff7f0e']
        )
        fig_trend.update_layout(xaxis_title="日期", yaxis_title="人数", hovermode="x unified")
        st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.info("暂无趋势数据")

with tab2:
    st.subheader("招生全流程转化漏斗")
    funnel_data = dict(
        number=[total_consults, total_trials, total_enrolls],
        stage=["1. 咨询量", "2. 试听量", "3. 最终报名"]
    )
    fig_funnel = go.Figure(go.Funnel(
        y=funnel_data['stage'],
        x=funnel_data['number'],
        textinfo="value+percent initial",
        marker={"color": ["#636EFA", "#EF553B", "#00CC96"]}
    ))
    fig_funnel.update_layout(title_text="全流程转化漏斗分析")
    st.plotly_chart(fig_funnel, use_container_width=True)

with tab3:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("各渠道报名贡献")
        if '渠道' in df_filtered.columns and not df_filtered.empty:
            df_channel = df_filtered.groupby('渠道')['报名人数'].sum().reset_index()
            fig_pie = px.pie(df_channel, names='渠道', values='报名人数', hole=0.4, title="渠道报名人数占比")
            st.plotly_chart(fig_pie, use_container_width=True)
    
    with c2:
        st.subheader("各校区业绩对比")
        if '校区' in df_filtered.columns and not df_filtered.empty:
            df_campus = df_filtered.groupby('校区')['总业绩(元)'].sum().reset_index()
            fig_bar = px.bar(df_campus, x='校区', y='总业绩(元)', color='校区', title="各校区总业绩比拼", text_auto='.2s')
            st.plotly_chart(fig_bar, use_container_width=True)

with tab4:
    st.subheader("交互式数据查看与修改")
    st.write("您可以在下方直接编辑修改数据，修改后可直接下载最新的 Excel 表格。")
    
    # 交互式数据编辑表格
    edited_df = st.data_editor(
        df_filtered,
        num_rows="dynamic",
        use_container_width=True,
        key="data_editor"
    )
    
    col_dl1, col_dl2 = st.columns([1, 4])
    with col_dl1:
        # 下载 CSV 格式
        csv_data = edited_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 导出当前数据为 CSV",
            data=csv_data,
            file_name=f"招生分析数据_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
