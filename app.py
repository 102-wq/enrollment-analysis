import streamlit as st
import pandas as pd
import sqlite3
import json
from datetime import datetime, date

# ==================== 页面与样式配置 ====================
st.set_page_config(
    page_title="招生数据分析与管理系统",
    page_icon="📊",
    layout="wide"
)

# 注入 CSS 实现移动端响应式适配与卡片样式
st.markdown("""
<style>
    .metric-card {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    @media (max-width: 768px) {
        .stButton>button {
            width: 100%;
        }
    }
</style>
""", unsafe_allow_html=True)

# ==================== 数据库初始化 (SQLite 本地持久化) ====================
DB_FILE = "enrollment_data.db"

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS enrollment (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_date TEXT NOT NULL,
            channel TEXT NOT NULL,
            leads_count INTEGER NOT NULL,
            enrolled_count INTEGER NOT NULL,
            revenue REAL NOT NULL
        )
    """)
    conn.commit()
    conn.close()

init_db()

# ==================== 数据操作函数 ====================
def load_data():
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM enrollment", conn)
    conn.close()
    if not df.empty:
        df['record_date'] = pd.to_datetime(df['record_date'])
        # 严格按日期升序排列，确保折线图趋势不乱序
        df = df.sort_values(by='record_date', ascending=True).reset_index(drop=True)
    return df

def insert_record(record_date, channel, leads_count, enrolled_count, revenue):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO enrollment (record_date, channel, leads_count, enrolled_count, revenue)
        VALUES (?, ?, ?, ?, ?)
    """, (str(record_date), channel, leads_count, enrolled_count, revenue))
    conn.commit()
    conn.close()

def delete_record(record_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM enrollment WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()

# ==================== 侧边栏：数据录入与备份恢复 ====================
st.sidebar.header("📝 数据录入与管理")

with st.sidebar.expander("新增招生记录", expanded=True):
    with st.form("entry_form", clear_on_submit=True):
        input_date = st.date_input("记录日期", value=date.today())
        input_channel = st.selectbox("来源渠道", ["线上广告", "社群转化", "线下转介绍", "自然流量", "其他"])
        input_leads = st.number_input("线索数量", min_value=0, step=1, value=10)
        input_enrolled = st.number_input("报名人数", min_value=0, step=1, value=1)
        input_revenue = st.number_input("成交金额 (元)", min_value=0.0, step=100.0, value=1000.0)
        
        submitted = st.form_submit_button("提交数据")
        if submitted:
            insert_record(input_date, input_channel, input_leads, input_enrolled, input_revenue)
            st.success("数据提交成功！")
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("💾 数据备份与恢复")

# JSON 导出
df_current = load_data()
if not df_current.empty:
    json_str = df_current.to_json(orient="records", date_format="iso")
    st.sidebar.download_button(
        label="📥 导出为 JSON 备份",
        data=json_str,
        file_name=f"enrollment_backup_{datetime.now().strftime('%Y%m%d')}.json",
        mime="application/json"
    )

# JSON 导入恢复
uploaded_file = st.sidebar.file_uploader("📤 导入 JSON 恢复数据", type=["json"])
if uploaded_file is not None:
    try:
        data = json.load(uploaded_file)
        conn = get_db_connection()
        cursor = conn.cursor()
        for row in data:
            cursor.execute("""
                INSERT INTO enrollment (record_date, channel, leads_count, enrolled_count, revenue)
                VALUES (?, ?, ?, ?, ?)
            """, (row['record_date'][:10], row['channel'], row['leads_count'], row['enrolled_count'], row['revenue']))
        conn.commit()
        conn.close()
        st.sidebar.success("数据成功导入并恢复！")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"导入失败，请检查文件格式: {e}")

# ==================== 主界面：数据展示与分析 ====================
st.title("📊 招生数据分析与管理看板")

df = load_data()

if df.empty:
    st.info("当前数据库暂无数据，请通过左侧边栏录入或导入数据。")
else:
    # 顶部时间范围过滤（按天切片分析）
    col_f1, col_f2 = st.columns(2)
    min_d = df['record_date'].min().date()
    max_d = df['record_date'].max().date()
    
    with col_f1:
        start_date = st.date_input("开始日期", value=min_d, min_value=min_d, max_value=max_d)
    with col_f2:
        end_date = st.date_input("结束日期", value=max_d, min_value=min_d, max_value=max_d)
    
    # 筛选数据
    mask = (df['record_date'].dt.date >= start_date) & (df['record_date'].dt.date <= end_date)
    filtered_df = df.loc[mask]

    # 核心指标卡展示
    total_leads = filtered_df['leads_count'].sum()
    total_enrolled = filtered_df['enrolled_count'].sum()
    total_revenue = filtered_df['revenue'].sum()
    conversion_rate = (total_enrolled / total_leads * 100) if total_leads > 0 else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("总线索量", f"{total_leads:,}")
    m2.metric("总报名人数", f"{total_enrolled:,}")
    m3.metric("总成交金额", f"¥{total_revenue:,.2f}")
    m4.metric("整体转化率", f"{conversion_rate:.2f}%")

    st.markdown("---")

    # 趋势分析与渠道分布图表
    t1, t2 = st.tabs(["📈 每日趋势变化", "📊 渠道数据对比"])

    with t1:
        st.subheader("每日招生与转化趋势")
        # 按天聚合，保证按日期升序
        daily_df = filtered_df.groupby('record_date')[['leads_count', 'enrolled_count', 'revenue']].sum().reset_index()
        daily_df = daily_df.sort_values(by='record_date', ascending=True)
        
        st.line_chart(daily_df, x='record_date', y=['leads_count', 'enrolled_count'])
        st.caption("注：折线图已配置严格按时间升序渲染。")

    with t2:
        st.subheader("不同渠道表现对比")
        channel_df = filtered_df.groupby('channel')[['leads_count', 'enrolled_count', 'revenue']].sum().reset_index()
        st.bar_chart(channel_df, x='channel', y='revenue')

    # 数据明细与删除管理
    st.markdown("---")
    st.subheader("📋 详细数据列表")
    
    # 展示数据表
    st.dataframe(filtered_df, use_container_width=True)

    with st.expander("🗑️ 管理/删除记录"):
        record_to_delete = st.selectbox("选择要删除的记录 ID", filtered_df['id'].tolist())
        if st.button("确认删除记录", type="primary"):
            delete_record(record_to_delete)
            st.success(f"ID 为 {record_to_delete} 的记录已安全删除！")
            st.rerun()
