import io
import json
import sqlite3
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 1. 数据库初始化与持久化操作
# -----------------------------------------------------------------------------
DB_FILE = "enrollment_data.db"

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        c = conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT)")
        c.execute("""
            CREATE TABLE IF NOT EXISTS daily_deltas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date_str TEXT NOT NULL,
                major TEXT NOT NULL,
                target_col TEXT NOT NULL,
                val INTEGER NOT NULL
            )
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_date ON daily_deltas(date_str)")
        conn.commit()

def load_data_from_db():
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT value FROM config WHERE key = 'raw_persons'")
        row_p = c.fetchone()
        raw_persons = json.loads(row_p['value']) if row_p else None
        
        c.execute("SELECT value FROM config WHERE key = 'person_animals'")
        row_a = c.fetchone()
        person_animals = json.loads(row_a['value']) if row_a else None

        c.execute("SELECT date_str, major, target_col, val FROM daily_deltas")
        rows = c.fetchall()
        
        daily_deltas = {}
        for r in rows:
            d, m, col, v = r['date_str'], r['major'], r['target_col'], r['val']
            daily_deltas.setdefault(d, []).append((m, col, v))
            
        return raw_persons, person_animals, daily_deltas

def save_all_to_db(raw_persons, person_animals, daily_deltas):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("INSERT OR REPLACE INTO config (key, value) VALUES ('raw_persons', ?)", (json.dumps(raw_persons, ensure_ascii=False),))
        c.execute("INSERT OR REPLACE INTO config (key, value) VALUES ('person_animals', ?)", (json.dumps(person_animals, ensure_ascii=False),))
        
        c.execute("DELETE FROM daily_deltas")
        for d, items in daily_deltas.items():
            for item in items:
                m, col, v = item if len(item) == 3 else ("电气基础", item[0], item[1])
                c.execute("INSERT INTO daily_deltas (date_str, major, target_col, val) VALUES (?, ?, ?, ?)", (d, m, col, v))
        conn.commit()

init_db()

# -----------------------------------------------------------------------------
# 2. 页面基本配置与 CSS 注入
# -----------------------------------------------------------------------------
st.set_page_config(page_title="招生数据动态管理与多维分析系统", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @media (max-width: 768px) {
        .kpi-card { height: auto !important; padding: 10px !important; margin-bottom: 8px; }
        .kpi-value { font-size: 20px !important; }
        .kpi-title { font-size: 12px !important; }
    }
    .kpi-card {
        background-color: #F8F9FA; border: 1px solid #E9ECEF; border-radius: 10px;
        padding: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.03); height: 105px;
        display: flex; flex-direction: column; justify-content: space-between; box-sizing: border-box;
    }
    .kpi-title { font-size: 14px; color: #31333F; font-weight: 500; display: flex; align-items: center; gap: 6px; }
    .kpi-body { display: flex; align-items: baseline; gap: 10px; }
    .kpi-value { font-size: 28px; font-weight: 700; color: #0E1117; font-family: "Source Sans Pro", sans-serif; line-height: 1; }
    .kpi-delta { font-size: 13px; font-weight: 600; color: #FF2B2B; background-color: #FFE6E6; padding: 2px 8px; border-radius: 12px; display: inline-flex; align-items: center; white-space: nowrap; }
    .table-container { width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; border-radius: 8px; border: 1px solid #7F7F7F; }
    div[data-testid="stForm"] { border-radius: 10px; background-color: #FAFAFA; }
</style>
""", unsafe_allow_html=True)

st.title("📊 招生数据动态管理与多维分析系统")
st.caption("2026年9月数据 - 默认动物头像全脱敏模式 | 本地 SQLite 自动实时存储")
st.markdown("---")

# -----------------------------------------------------------------------------
# 3. 基础数据定义（仅填入 9月1日 真实数据，清空其他日期）
# -----------------------------------------------------------------------------
DEFAULT_MAJORS = [
    ("给排水专业", 19, [4, 3, 6, 3, 3]), ("发输电专业", 10, [2, 3, 2, 1, 2]),
    ("供配电专业", 11, [3, 4, 2, 1, 1]), ("环保专业", 13, [4, 2, 2, 3, 2]),
    ("环评专业", 9, [2, 2, 1, 3, 1]), ("岩土专业", 8, [2, 2, 2, 1, 1]),
    ("暖通专业", 16, [3, 3, 3, 2, 5]), ("结构专业", 20, [4, 4, 4, 5, 3]),
    ("道路专业", 5, [1, 1, 1, 1, 1]), ("233网校", 5, [1, 1, 1, 1, 1]),
    ("电气基础", 53, [11, 16, 11, 8, 7]), ("环保基础", 30, [9, 6, 6, 5, 4]),
    ("岩土基础", 36, [9, 8, 7, 6, 6]), ("水基础", 17, [4, 4, 5, 2, 2]),
    ("暖通基础", 27, [5, 5, 4, 4, 9]), ("结构基础", 9, [2, 2, 1, 2, 1]),
    ("公共基础", 4, [2, 1, 1, 0, 1]), ("道路基础", 6, [1, 1, 1, 1, 2]),
    ("水利水电基础", 10, [2, 2, 3, 2, 1]),
]

DATES = [f"9月{i}日" for i in range(1, 31)]

def build_base_targets_df(persons):
    base_data = []
    for idx, (name, total_target, person_tgts) in enumerate(DEFAULT_MAJORS, 1):
        row = {"序号": idx, "专业/基础名称": name, "目标人数": total_target}
        for p_idx, p in enumerate(persons):
            row[f"{p}_目标"] = person_tgts[p_idx] if p_idx < len(person_tgts) else 0
            row[f"{p}_实际"] = 0
        row["其他人员_目标"] = 0
        row["其他人员_实际"] = 0
        base_data.append(row)
    return pd.DataFrame(base_data)

def reset_to_sep1_only():
    raw_persons = ["覃小燕", "左丹丹", "梁书华", "古晨晓", "周欢喜"]
    person_animals = {"覃小燕": "🦊", "左丹丹": "🐼", "梁书华": "🦁", "古晨晓": "🐰", "周欢喜": "🐯"}
    
    # 清空所有旧数据，仅保留 9月1日 的 9 笔真实成交流水
    daily_deltas = {
        "9月1日": [
            ("环保专业", "覃小燕_实际", 1),
            ("电气基础", "覃小燕_实际", 1),
            ("环保专业", "左丹丹_实际", 1),
            ("环保基础", "左丹丹_实际", 2),
            ("岩土基础", "梁书华_实际", 1),
            ("暖通基础", "周欢喜_实际", 1),
            ("233网校", "其他人员_实际", 1),
            ("电气基础", "其他人员_实际", 1)
        ]
    }
    save_all_to_db(raw_persons, person_animals, daily_deltas)
    return raw_persons, person_animals, daily_deltas

db_raw_persons, db_person_animals, db_daily_deltas = reset_to_sep1_only()
st.session_state["raw_persons"] = db_raw_persons
st.session_state["person_animals"] = db_person_animals
st.session_state["daily_deltas"] = db_daily_deltas

st.session_state["base_targets"] = build_base_targets_df(st.session_state["raw_persons"])
RAW_PERSONS = st.session_state["raw_persons"]

# -----------------------------------------------------------------------------
# 4. 侧边栏控制
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ 数据管理与设置")
enable_anonymize = st.sidebar.checkbox("开启数据脱敏 / 纯动物符号模式", value=True)

alias_map = {p: (st.session_state["person_animals"].get(p, "🐱") if enable_anonymize else p) for p in RAW_PERSONS}
PERSONS = [alias_map[p] for p in RAW_PERSONS]

if enable_anonymize:
    with st.sidebar.expander("👁️ 视角对照表（管理者隐私预览）", expanded=False):
        st.dataframe(pd.DataFrame({"真实姓名": RAW_PERSONS, "代称动物": PERSONS}), hide_index=True, use_container_width=True)

st.sidebar.markdown("---")
st.sidebar.subheader("💾 数据重置")
if st.sidebar.button("💥 重置（仅保留9月1日数据）", use_container_width=True):
    r_p, p_a, d_d = reset_to_sep1_only()
    st.session_state["raw_persons"], st.session_state["person_animals"], st.session_state["daily_deltas"] = r_p, p_a, d_d
    st.sidebar.success("已清空旧数据，仅保留 9月1日 数据！")
    st.rerun()

# -----------------------------------------------------------------------------
# 5. 时间汇总与 KPI 展示
# -----------------------------------------------------------------------------
st.subheader("🗓️ 时间汇总粒度筛选")
f_col1, f_col2 = st.columns(2)

with f_col1:
    time_granularity_type = st.selectbox("选择时间汇总粒度：", ["按月（月度全量）", "按日（单日切片）"])

with f_col2:
    if time_granularity_type == "按日（单日切片）":
        selected_time_range = st.selectbox("选择具体日期：", DATES, index=0)
        selected_dates_list = [selected_time_range]
    else:
        selected_time_range = st.selectbox("选择具体月份：", ["2026年9月"])
        selected_dates_list = DATES

def get_processed_df_by_dates(dates_list):
    df_result = build_base_targets_df(st.session_state["raw_persons"])
    act_cols = [c for c in df_result.columns if c.endswith("_实际")]
    for col in act_cols: df_result[col] = 0

    for d in dates_list:
        for item in st.session_state["daily_deltas"].get(d, []):
            major, col, val = item if len(item) == 3 else ("电气基础", item[0], item[1])
            if col in df_result.columns:
                df_result.loc[df_result["专业/基础名称"] == major, col] += val
    return df_result

calc_df = get_processed_df_by_dates(selected_dates_list)
act_cols = [c for c in calc_df.columns if c.endswith("_实际")]
calc_df["实际完成"] = calc_df[act_cols].sum(axis=1)
calc_df["与目标之差"] = calc_df["实际完成"] - calc_df["目标人数"]

num_cols = [c for c in calc_df.columns if c not in ["序号", "专业/基础名称"]]
sum_row = {"专业/基础名称": "合计"}
for c in num_cols: sum_row[c] = int(calc_df[c].sum())

total_target_cum = sum_row["目标人数"]
total_actual_cum = sum_row["实际完成"]
cum_rate_val = (total_actual_cum / total_target_cum * 100) if total_target_cum > 0 else 0

m_col1, m_col2, m_col3, m_col4 = st.columns(4)
m_col1.markdown(f'<div class="kpi-card"><div class="kpi-title">🎯 总目标人数</div><div class="kpi-body"><div class="kpi-value">{total_target_cum} 人</div></div></div>', unsafe_allow_html=True)
m_col2.markdown(f'<div class="kpi-card"><div class="kpi-title">✅ 实际完成人数</div><div class="kpi-body"><div class="kpi-value">{total_actual_cum} 人</div><div class="kpi-delta">↓ {sum_row["与目标之差"]} 人</div></div></div>', unsafe_allow_html=True)
m_col3.markdown(f'<div class="kpi-card"><div class="kpi-title">📈 目标完成比例</div><div class="kpi-body"><div class="kpi-value">{cum_rate_val:.2f}%</div></div></div>', unsafe_allow_html=True)
m_col4.markdown(f'<div class="kpi-card"><div class="kpi-title">📅 当前切片完成人数</div><div class="kpi-body"><div class="kpi-value">{total_actual_cum} 人</div></div></div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 6. 图表可视化展示
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("### 📊 图表可视化分析")
c1, c2 = st.columns(2)

with c1:
    fig_person = go.Figure(data=[
        go.Bar(name="目标人数", x=PERSONS, y=[sum_row[f"{p}_目标"] for p in RAW_PERSONS], marker_color="#0B3C5D", text=[sum_row[f"{p}_目标"] for p in RAW_PERSONS], textposition="outside"),
        go.Bar(name="实际完成", x=PERSONS, y=[sum_row[f"{p}_实际"] for p in RAW_PERSONS], marker_color="#FF3D00", text=[sum_row[f"{p}_实际"] for p in RAW_PERSONS], textposition="outside")
    ])
    fig_person.update_layout(title="<b>各成员目标 vs 实际完成对比</b>", barmode="group", plot_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20), yaxis=dict(gridcolor="#E0E0E0"))
    st.plotly_chart(fig_person, use_container_width=True)

with c2:
    active_majors = calc_df[calc_df["实际完成"] > 0].sort_values(by="实际完成", ascending=False)
    if not active_majors.empty:
        fig_major = px.bar(active_majors, x="实际完成", y="专业/基础名称", orientation="h", title="<b>当前切片有成交的专业排行</b>", text="实际完成", color="专业/基础名称")
        fig_major.update_traces(textposition="outside")
    else:
        fig_major = px.bar(x=[0], y=["无成交记录"], orientation="h", title="<b>当前切片无成交记录</b>")
    fig_major.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False, plot_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20), xaxis=dict(gridcolor="#E0E0E0"))
    st.plotly_chart(fig_major, use_container_width=True)
