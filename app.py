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
# 1. 資料庫初始化與持久化操作
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
                m, col, v = item if len(item) == 3 else ("電氣基礎", item[0], item[1])
                c.execute("INSERT INTO daily_deltas (date_str, major, target_col, val) VALUES (?, ?, ?, ?)", (d, m, col, v))
        conn.commit()

init_db()

# -----------------------------------------------------------------------------
# 2. 頁面基本配置與 CSS 注入
# -----------------------------------------------------------------------------
st.set_page_config(page_title="招生數據動態管理與多維分析系統", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

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

st.title("📊 招生數據動態管理與多維分析系統")
st.caption("2026年9月數據 - 預設動物頭像全脫敏模式 | 本地 SQLite 自動實時儲存")
st.markdown("---")

# -----------------------------------------------------------------------------
# 3. 基礎數據定義（更新 9月1日 精確真實數據）
# -----------------------------------------------------------------------------
DEFAULT_MAJORS = [
    ("給排水專業", 19, [4, 3, 6, 3, 3]), ("發輸電專業", 10, [2, 3, 2, 1, 2]),
    ("供配電專業", 11, [3, 4, 2, 1, 1]), ("環保專業", 13, [4, 2, 2, 3, 2]),
    ("環評專業", 9, [2, 2, 1, 3, 1]), ("岩土專業", 8, [2, 2, 2, 1, 1]),
    ("暖通專業", 16, [3, 3, 3, 2, 5]), ("結構專業", 20, [4, 4, 4, 5, 3]),
    ("道路專業", 5, [1, 1, 1, 1, 1]), ("233網校", 5, [1, 1, 1, 1, 1]),
    ("電氣基礎", 53, [11, 16, 11, 8, 7]), ("環保基礎", 30, [9, 6, 6, 5, 4]),
    ("岩土基礎", 36, [9, 8, 7, 6, 6]), ("水基礎", 17, [4, 4, 5, 2, 2]),
    ("暖通基礎", 27, [5, 5, 4, 4, 9]), ("結構基礎", 9, [2, 2, 1, 2, 1]),
    ("公共基礎", 4, [2, 1, 1, 0, 1]), ("道路基礎", 6, [1, 1, 1, 1, 2]),
    ("水利水電基礎", 10, [2, 2, 3, 2, 1]),
]

DATES = [f"9月{i}日" for i in range(1, 21)]
WEEKDAYS = ["星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
AVAL_ANIMALS = ["🦊", "🐼", "🦁", "🐰", "🐯", "🐱", "🐶", "🐻", "🐨", "🐮", "🐵", "🐥"]

def build_base_targets_df(persons):
    base_data = []
    for idx, (name, total_target, person_tgts) in enumerate(DEFAULT_MAJORS, 1):
        row = {"序號": idx, "專業/基礎名稱": name, "目標人數": total_target}
        for p_idx, p in enumerate(persons):
            row[f"{p}_目標"] = person_tgts[p_idx] if p_idx < len(person_tgts) else 0
            row[f"{p}_實際"] = 0
        row["其他人員_目標"] = 0
        row["其他人員_實際"] = 0
        base_data.append(row)
    return pd.DataFrame(base_data)

def reset_to_default_mock():
    raw_persons = ["覃小燕", "左丹丹", "梁書華", "古晨曉", "周歡喜"]
    person_animals = {"覃小燕": "🦊", "左丹丹": "🐼", "梁書華": "🦁", "古晨曉": "🐰", "周歡喜": "🐯"}
    
    # 根據圖片精確錄入 9月1日 流水 (合計 9 人)
    daily_deltas = {
        "9月1日": [
            ("環保專業", "覃小燕_實際", 1),
            ("電氣基礎", "覃小燕_實際", 1),
            ("環保專業", "左丹丹_實際", 1),
            ("環保基礎", "左丹丹_實際", 2),
            ("岩土基礎", "梁書華_實際", 1),
            ("暖通基礎", "周歡喜_實際", 1),
            ("233網校", "其他人員_實際", 1),
            ("電氣基礎", "其他人員_實際", 1)
        ],
        "9月2日": [("環保專業", "覃小燕_實際", 1), ("環保專業", "左丹丹_實際", 1), ("電氣基礎", "覃小燕_實際", 1), ("發輸電專業", "其他人員_實際", 1)],
        "9月3日": [("環評專業", "左丹丹_實際", 1), ("暖通專業", "周歡喜_實際", 1), ("環保基礎", "左丹丹_實際", 1)],
        "9月4日": [("電氣基礎", "覃小燕_實際", 1), ("電氣基礎", "左丹丹_實際", 1), ("水利水電基礎", "其他人員_實際", 1)],
        "9月5日": [("給排水專業", "梁書華_實際", 1), ("暖通專業", "周歡喜_實際", 1), ("暖通基礎", "周歡喜_實際", 1)],
        "9月6日": [("給排水專業", "梁書華_實際", 1), ("岩土專業", "梁書華_實際", 1)],
        "9月7日": [("電氣基礎", "覃小燕_實際", 3), ("水基礎", "覃小燕_實際", 1), ("暖通專業", "梁書華_實際", 1)],
        "9月8日": [("給排水專業", "梁書華_實際", 1)],
        "9月9日": [("暖通專業", "左丹丹_實際", 1), ("水基礎", "周歡喜_實際", 1)],
        "9月10日": [("電氣基礎", "覃小燕_實際", 2), ("暖通專業", "周歡喜_實際", 1)],
        "9月11日": [("電氣基礎", "覃小燕_實際", 1), ("道路基礎", "覃小燕_實際", 1)],
        "9月12日": [("電氣基礎", "覃小燕_實際", 1)],
        "9月13日": [("給排水專業", "周歡喜_實際", 1), ("結構專業", "其他人員_實際", 1)],
        "9月14日": [("電氣基礎", "覃小燕_實際", 1), ("結構專業", "梁書華_實際", 1)],
        "9月15日": [("暖通專業", "古晨曉_實際", 1), ("環保基礎", "左丹丹_實際", 1)],
        "9月16日": [("發輸電專業", "左丹丹_實際", 1), ("岩土基礎", "梁書華_實際", 1)],
        "9月17日": [("電氣基礎", "覃小燕_實際", 2), ("供配電專業", "周歡喜_實際", 1)],
        "9月18日": [("給排水專業", "梁書華_實際", 1), ("水基礎", "古晨曉_實際", 1)],
        "9月19日": [("暖通基礎", "周歡喜_實際", 1), ("環保專業", "其他人員_實際", 1)],
        "9月20日": [("電氣基礎", "覃小燕_實際", 1), ("岩土專業", "古晨曉_實際", 1)]
    }
    save_all_to_db(raw_persons, person_animals, daily_deltas)
    return raw_persons, person_animals, daily_deltas

db_raw_persons, db_person_animals, db_daily_deltas = reset_to_default_mock()
st.session_state["raw_persons"] = db_raw_persons
st.session_state["person_animals"] = db_person_animals
st.session_state["daily_deltas"] = db_daily_deltas

st.session_state["base_targets"] = build_base_targets_df(st.session_state["raw_persons"])
RAW_PERSONS = st.session_state["raw_persons"]

# -----------------------------------------------------------------------------
# 4. 側邊欄控制
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ 資料管理與設定")
enable_anonymize = st.sidebar.checkbox("開啟數據脫敏 / 純動物符號模式", value=True)

alias_map = {p: (st.session_state["person_animals"].get(p, "🐱") if enable_anonymize else p) for p in RAW_PERSONS}
PERSONS = [alias_map[p] for p in RAW_PERSONS]

if enable_anonymize:
    with st.sidebar.expander("👁️ 視角對照表（管理者隱私預覽）", expanded=False):
        st.dataframe(pd.DataFrame({"真實姓名": RAW_PERSONS, "代稱動物": PERSONS}), hide_index=True, use_container_width=True)

st.sidebar.markdown("---")
st.sidebar.subheader("💾 資料重置")
if st.sidebar.button("💥 強制重置為圖片中 9月1日 真實數據", use_container_width=True):
    r_p, p_a, d_d = reset_to_default_mock()
    st.session_state["raw_persons"], st.session_state["person_animals"], st.session_state["daily_deltas"] = r_p, p_a, d_d
    st.sidebar.info("9月1日真實數據已重置成功！")
    st.rerun()

# -----------------------------------------------------------------------------
# 5. 時間彙總粒度篩選與 KPI 渲染
# -----------------------------------------------------------------------------
st.subheader("🗓️ 時間彙總粒度篩選")
f_col1, f_col2 = st.columns(2)

with f_col1:
    time_granularity_type = st.selectbox("選擇時間彙總粒度：", ["按月（月度全量）", "按周（週度彙總）", "按日（單日切片）"])

with f_col2:
    if time_granularity_type == "按日（單日切片）":
        selected_time_range = st.selectbox("選擇具體日期：", DATES, index=0)
        selected_dates_list = [selected_time_range]
    elif time_granularity_type == "按周（週度彙總）":
        selected_time_range = st.selectbox("選擇具體周：", ["2026年第36-37周 (9月1日-9月20日)"])
        selected_dates_list = DATES
    else:
        selected_time_range = st.selectbox("選擇具體月份：", ["2026年9月1日-9月20日"])
        selected_dates_list = DATES

def get_processed_df_by_dates(dates_list):
    df_result = build_base_targets_df(st.session_state["raw_persons"])
    act_cols = [c for c in df_result.columns if c.endswith("_實際")]
    for col in act_cols: df_result[col] = 0

    for d in dates_list:
        for item in st.session_state["daily_deltas"].get(d, []):
            major, col, val = item if len(item) == 3 else ("電氣基礎", item[0], item[1])
            if col in df_result.columns:
                df_result.loc[df_result["專業/基礎名稱"] == major, col] += val
    return df_result

calc_df = get_processed_df_by_dates(selected_dates_list)
act_cols = [c for c in calc_df.columns if c.endswith("_實際")]
calc_df["實際完成"] = calc_df[act_cols].sum(axis=1)
calc_df["與目標之差"] = calc_df["實際完成"] - calc_df["目標人數"]

num_cols = [c for c in calc_df.columns if c not in ["序號", "專業/基礎名稱"]]
sum_row = {"專業/基礎名稱": "合計"}
for c in num_cols: sum_row[c] = int(calc_df[c].sum())

diff_row = {"專業/基礎名稱": "與目標之差"}
for p in RAW_PERSONS:
    diff_row[f"{p}_目標"] = ""
    diff_row[f"{p}_實際"] = sum_row[f"{p}_實際"] - sum_row[f"{p}_目標"]
diff_row["其他人員_目標"] = ""
diff_row["其他人員_實際"] = sum_row["其他人員_實際"]
diff_row.update({"目標人數": "", "實際完成": "", "與目標之差": sum_row["與目標之差"]})

total_target_cum = sum_row["目標人數"]
total_actual_cum = sum_row["實際完成"]
cum_rate_val = (total_actual_cum / total_target_cum * 100) if total_target_cum > 0 else 0
rate_row = {c: "" for c in num_cols}
rate_row.update({"專業/基礎名稱": "目標人數完成比例", "實際完成": f"{cum_rate_val:.2f}%"})

m_col1, m_col2, m_col3, m_col4 = st.columns(4)
m_col1.markdown(f'<div class="kpi-card"><div class="kpi-title">🎯 總目標人數</div><div class="kpi-body"><div class="kpi-value">{total_target_cum} 人</div></div></div>', unsafe_allow_html=True)
m_col2.markdown(f'<div class="kpi-card"><div class="kpi-title">✅ 實際完成人數</div><div class="kpi-body"><div class="kpi-value">{total_actual_cum} 人</div><div class="kpi-delta">↓ {sum_row["與目標之差"]} 人</div></div></div>', unsafe_allow_html=True)
m_col3.markdown(f'<div class="kpi-card"><div class="kpi-title">📈 目標完成比例</div><div class="kpi-body"><div class="kpi-value">{cum_rate_val:.2f}%</div></div></div>', unsafe_allow_html=True)
m_col4.markdown(f'<div class="kpi-card"><div class="kpi-title">📅 當前切片完成人數</div><div class="kpi-body"><div class="kpi-value">{total_actual_cum} 人</div></div></div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 6. HTML 表格與圖表展示
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("### 📊 圖表可視化分析")
c1, c2 = st.columns(2)

with c1:
    fig_person = go.Figure(data=[
        go.Bar(name="目標人數", x=PERSONS, y=[sum_row[f"{p}_目標"] for p in RAW_PERSONS], marker_color="#0B3C5D", text=[sum_row[f"{p}_目標"] for p in RAW_PERSONS], textposition="outside"),
        go.Bar(name="實際完成", x=PERSONS, y=[sum_row[f"{p}_實際"] for p in RAW_PERSONS], marker_color="#FF3D00", text=[sum_row[f"{p}_實際"] for p in RAW_PERSONS], textposition="outside")
    ])
    fig_person.update_layout(title="<b>各成員目標 vs 實際完成對比</b>", barmode="group", plot_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20), yaxis=dict(gridcolor="#E0E0E0"))
    st.plotly_chart(fig_person, use_container_width=True)

with c2:
    active_majors = calc_df[calc_df["實際完成"] > 0].sort_values(by="實際完成", ascending=False)
    if not active_majors.empty:
        fig_major = px.bar(active_majors, x="實際完成", y="專業/基礎名稱", orientation="h", title="<b>當前切片有成交的專業排行</b>", text="實際完成", color="專業/基礎名稱")
        fig_major.update_traces(textposition="outside")
    else:
        fig_major = px.bar(x=[0], y=["無成交記錄"], orientation="h", title="<b>當前切片無成交記錄</b>")
    fig_major.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False, plot_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20), xaxis=dict(gridcolor="#E0E0E0"))
    st.plotly_chart(fig_major, use_container_width=True)
