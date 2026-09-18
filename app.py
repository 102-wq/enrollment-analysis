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
# 1. 數據庫初始化與持久化操作 (SQLite 安全上下文優化)
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
# 2. 頁面基本配置與全局響應式 CSS 注入
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
st.caption("2026年9月數據 - 預設動物頭像全脫敏模式 | 本地 SQLite 自動實時存儲")
st.markdown("---")

# -----------------------------------------------------------------------------
# 3. 基礎數據定義與 Session State / SQLite 同步
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

DATES = [f"9月{i}日" for i in range(1, 14)]
WEEKDAYS = ["星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
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
    daily_deltas = {
        "9月1日": [("電氣基礎", "覃小燕_實際", 1)],
        "9月2日": [("環保專業", "覃小燕_實際", 1), ("環保專業", "左丹丹_實際", 1), ("電氣基礎", "覃小燕_實際", 1), ("發輸電專業", "其他人員_實際", 1), ("233網校", "其他人員_實際", 1)],
        "9月3日": [("環評專業", "左丹丹_實際", 1), ("暖通專業", "周歡喜_實際", 1), ("環保基礎", "左丹丹_實際", 1), ("環保基礎", "其他人員_實際", 1)],
        "9月4日": [("電氣基礎", "覃小燕_實際", 1), ("電氣基礎", "左丹丹_實際", 1), ("環保基礎", "覃小燕_實際", 1), ("水利水電基礎", "其他人員_實際", 1)],
        "9月5日": [("給排水專業", "梁書華_實際", 1), ("暖通專業", "周歡喜_實際", 1), ("岩土基礎", "梁書華_實際", 1), ("暖通基礎", "周歡喜_實際", 1)],
        "9月6日": [("給排水專業", "梁書華_實際", 1), ("環保專業", "其他人員_實際", 1), ("岩土專業", "梁書華_實際", 1)],
        "9月7日": [("電氣基礎", "覃小燕_實際", 3), ("環保基礎", "覃小燕_實際", 1), ("水基礎", "覃小燕_實際", 1), ("環保基礎", "左丹丹_實際", 1), ("暖通基礎", "左丹丹_實際", 1), ("暖通專業", "梁書華_實際", 1), ("岩土基礎", "梁書華_實際", 1), ("暖通基礎", "梁書華_實際", 1), ("公共基礎", "古晨曉_實際", 1), ("環保基礎", "古晨曉_實際", 1), ("結構基礎", "其他人員_實際", 2)],
        "9月8日": [("給排水專業", "梁書華_實際", 1)],
        "9月9日": [("暖通專業", "左丹丹_實際", 1), ("233網校", "左丹丹_實際", 1), ("水基礎", "周歡喜_實際", 1), ("給排水專業", "其他人員_實際", 1)],
        "9月10日": [("電氣基礎", "覃小燕_實際", 2), ("環保基礎", "覃小燕_實際", 1), ("暖通專業", "周歡喜_實際", 1), ("電氣基礎", "周歡喜_實際", 1), ("暖通基礎", "周歡喜_實際", 1), ("岩土基礎", "古晨曉_實際", 1), ("水利水電基礎", "左丹丹_實際", 1), ("岩土基礎", "梁書華_實際", 1)],
        "9月11日": [("電氣基礎", "覃小燕_實際", 1), ("道路基礎", "覃小燕_實際", 1), ("岩土基礎", "其他人員_實際", 1), ("電氣基礎", "梁書華_實際", 1)],
        "9月12日": [("電氣基礎", "覃小燕_實際", 1)],
        "9月13日": [("給排水專業", "周歡喜_實際", 1), ("結構專業", "其他人員_實際", 1), ("電氣基礎", "覃小燕_實際", 1), ("電氣基礎", "左丹丹_實際", 1), ("水基礎", "覃小燕_實際", 1), ("水基礎", "梁書華_實際", 1), ("暖通基礎", "覃小燕_實際", 1), ("暖通基礎", "周歡喜_實際", 1), ("水利水電基礎", "梁書華_實際", 1), ("水利水電基礎", "其他人員_實際", 1)]
    }
    save_all_to_db(raw_persons, person_animals, daily_deltas)
    return raw_persons, person_animals, daily_deltas

if "raw_persons" not in st.session_state:
    db_raw_persons, db_person_animals, db_daily_deltas = load_data_from_db()
    if db_raw_persons is None:
        db_raw_persons, db_person_animals, db_daily_deltas = reset_to_default_mock()
    st.session_state["raw_persons"] = db_raw_persons
    st.session_state["person_animals"] = db_person_animals
    st.session_state["daily_deltas"] = db_daily_deltas

st.session_state["base_targets"] = build_base_targets_df(st.session_state["raw_persons"])
RAW_PERSONS = st.session_state["raw_persons"]

# -----------------------------------------------------------------------------
# 4. 側邊欄：脫敏管理 & 人員增刪 & 數據備份恢復
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ 數據管理與設置")
enable_anonymize = st.sidebar.checkbox("開啟數據脫敏 / 純動物符號模式", value=True)

alias_map = {p: (st.session_state["person_animals"].get(p, "🐱") if enable_anonymize else p) for p in RAW_PERSONS}
PERSONS = [alias_map[p] for p in RAW_PERSONS]

if enable_anonymize:
    with st.sidebar.expander("👁️ 視角對照表（管理者隱私預覽）", expanded=False):
        st.dataframe(pd.DataFrame({"真實姓名": RAW_PERSONS, "代稱動物": PERSONS}), hide_index=True, use_container_width=True)

with st.sidebar.expander("👥 人員增刪管理"):
    new_p_name = st.text_input("姓名", placeholder="例如：張三", key="new_person_name_input")
    new_p_emoji = st.selectbox("分配動物標誌", AVAL_ANIMALS, key="new_person_emoji_input")
    if st.button("➕ 確認新增", use_container_width=True):
        if new_p_name and new_p_name not in RAW_PERSONS:
            st.session_state["raw_persons"].append(new_p_name)
            st.session_state["person_animals"][new_p_name] = new_p_emoji
            save_all_to_db(st.session_state["raw_persons"], st.session_state["person_animals"], st.session_state["daily_deltas"])
            st.sidebar.success(f"已添加：{new_p_name} ({new_p_emoji})")
            st.rerun()

# -----------------------------------------------------------------------------
# 5. 快捷錄入與刪減招生數據
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("✏️ 招生人數 增加 / 刪減")
op_mode = st.sidebar.radio("操作模式：", ["➕ 新增完成人數", "➖ 刪減完成人數"], horizontal=True)

with st.sidebar.form("add_delta_form", clear_on_submit=True):
    input_date = st.selectbox("日期", DATES)
    input_major = st.selectbox("專業/基礎", list(st.session_state["base_targets"]["專業/基礎名稱"]))
    input_person_disp = st.selectbox("歸屬人員", PERSONS + ["其他人員"])
    input_val = st.number_input("變動人數", min_value=1, value=1, step=1)

    if st.form_submit_button("確認提交修改", use_container_width=True):
        inv_alias_map = {v: k for k, v in alias_map.items()}
        raw_person_name = inv_alias_map.get(input_person_disp, input_person_disp)
        target_col = f"{raw_person_name}_實際" if raw_person_name != "其他人員" else "其他人員_實際"
        actual_change = int(input_val) if op_mode == "➕ 新增完成人數" else -int(input_val)
        
        st.session_state["daily_deltas"].setdefault(input_date, []).append((input_major, target_col, actual_change))
        save_all_to_db(st.session_state["raw_persons"], st.session_state["person_animals"], st.session_state["daily_deltas"])
        st.sidebar.success(f"已更新：{input_date} {input_major} - {input_person_disp}")
        st.rerun()

with st.sidebar.expander("🗑️ 招生流水明細與單條刪除"):
    del_date = st.selectbox("選擇要查驗的日期：", DATES, key="del_date_sel")
    day_records = st.session_state["daily_deltas"].get(del_date, [])
    if not day_records:
        st.info("該日期暫無記錄")
    else:
        for r_idx, item in enumerate(day_records):
            m_name, p_col, val_num = item if len(item) == 3 else ("電氣基礎", item[0], item[1])
            p_raw = p_col.replace("_實際", "")
            c_lbl, c_btn = st.columns([3, 1])
            c_lbl.caption(f"{m_name} | {alias_map.get(p_raw, p_raw)} | {'+' if val_num>0 else ''}{val_num}人")
            if c_btn.button("刪除", key=f"del_{del_date}_{r_idx}"):
                st.session_state["daily_deltas"][del_date].pop(r_idx)
                if not st.session_state["daily_deltas"][del_date]:
                    del st.session_state["daily_deltas"][del_date]
                save_all_to_db(st.session_state["raw_persons"], st.session_state["person_animals"], st.session_state["daily_deltas"])
                st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("💾 數據備份與恢復")
json_str = json.dumps({
    "raw_persons": st.session_state["raw_persons"],
    "person_animals": st.session_state["person_animals"],
    "daily_deltas": st.session_state["daily_deltas"]
}, ensure_ascii=False, indent=2)

st.sidebar.download_button("📤 匯出 JSON 備份文件", data=json_str, file_name="招生數據備份.json", mime="application/json", use_container_width=True)

uploaded_file = st.sidebar.file_uploader("📥 匯入 JSON 恢復數據", type=["json"])
if uploaded_file:
    try:
        data = json.load(uploaded_file)
        st.session_state["raw_persons"], st.session_state["person_animals"], st.session_state["daily_deltas"] = data["raw_persons"], data["person_animals"], data["daily_deltas"]
        save_all_to_db(st.session_state["raw_persons"], st.session_state["person_animals"], st.session_state["daily_deltas"])
        st.sidebar.success("數據恢復成功！")
        st.rerun()
    except Exception:
        st.sidebar.error("備份文件格式不正確")

if st.sidebar.button("🔄 重置全表為初始狀態", use_container_width=True):
    r_p, p_a, d_d = reset_to_default_mock()
    st.session_state["raw_persons"], st.session_state["person_animals"], st.session_state["daily_deltas"] = r_p, p_a, d_d
    st.sidebar.info("數據已重置！")
    st.rerun()

# -----------------------------------------------------------------------------
# 6. 時間彙總粒度篩選與 KPI 渲染
# -----------------------------------------------------------------------------
st.subheader("🗓️ 時間彙總粒度篩選")
f_col1, f_col2 = st.columns(2)

with f_col1:
    time_granularity_type = st.selectbox("選擇時間彙總粒度：", ["按月（月度全量）", "按周（周度彙總）", "按日（單日切片）"])

with f_col2:
    if time_granularity_type == "按日（單日切片）":
        selected_time_range = st.selectbox("選擇具體日期：", DATES, index=len(DATES)-1)
        selected_dates_list = [selected_time_range]
    elif time_granularity_type == "按周（周度彙總）":
        selected_time_range = st.selectbox("選擇具體周：", ["2026年第36-37周 (9月1日-9月13日)"])
        selected_dates_list = DATES
    else:
        selected_time_range = st.selectbox("選擇具體月份：", ["2026年9月1日-9月13日"])
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

avg_per_day = total_actual_cum / len(selected_dates_list) if len(selected_dates_list) > 0 else 0

m_col1, m_col2, m_col3, m_col4 = st.columns(4)
m_col1.markdown(f'<div class="kpi-card"><div class="kpi-title">🎯 總目標人數</div><div class="kpi-body"><div class="kpi-value">{total_target_cum} 人</div></div></div>', unsafe_allow_html=True)
m_col2.markdown(f'<div class="kpi-card"><div class="kpi-title">✅ 實際完成人數</div><div class="kpi-body"><div class="kpi-value">{total_actual_cum} 人</div><div class="kpi-delta">↓ {sum_row["與目標之差"]} 人</div></div></div>', unsafe_allow_html=True)
m_col3.markdown(f'<div class="kpi-card"><div class="kpi-title">📈 目標完成比例</div><div class="kpi-body"><div class="kpi-value">{cum_rate_val:.2f}%</div></div></div>', unsafe_allow_html=True)
m_col4.markdown(f'<div class="kpi-card"><div class="kpi-title">📅 選定區間日均新增</div><div class="kpi-body"><div class="kpi-value">{avg_per_day:.1f} 人/天</div></div></div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 7. HTML 數據表格渲染 (對齊其他人員與目標彙總)
# -----------------------------------------------------------------------------
def build_html_document(df, sum_r, diff_r, rate_r, p_names, raw_p_names, range_title):
    rows_html = ""
    for idx, row in df.iterrows():
        person_cells = "".join([f'<td class="num">{row[f"{rp}_目標"]}</td><td class="num">{row[f"{rp}_實際"] or ""}</td>' for rp in raw_p_names])
        rows_html += f'<tr><td class="num">{row["序號"]}</td><td class="zh">{row["專業/基礎名稱"]}</td><td class="num">{row["目標人數"]}</td>{person_cells}<td class="num">{row["其他人員_目標"] or ""}</td><td class="num">{row["其他人員_實際"] or ""}</td><td class="num">{row["目標人數"]}</td><td class="num">{row["實際完成"]}</td><td class="num">{row["與目標之差"]}</td></tr>'

    person_headers = "".join([f'<th colspan="2" class="bg-person">{p}</th>' for p in p_names])
    sub_headers = '<th class="bg-header">目標</th><th class="bg-header">實際</th>' * (len(p_names) + 1)
    
    sum_person_cells = "".join([f'<td class="bg-total num">{sum_r[f"{rp}_目標"]}</td><td class="bg-total num">{sum_r[f"{rp}_實際"]}</td>' for rp in raw_p_names])
    diff_person_cells = "".join([f'<td class="bg-total" colspan="2"><b class="num">{diff_r[f"{rp}_實際"]}</b></td>' for rp in raw_p_names])
    
    total_colspan = 7 + (len(p_names) + 1) * 2

    return f"""
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
        body {{ margin: 0; padding: 0; font-family: SimSun, "Times New Roman", serif; background-color: #ffffff; }}
        .table-container {{ width: 100%; max-height: 640px; overflow: auto; -webkit-overflow-scrolling: touch; border-radius: 8px; border: 1px solid #7F7F7F; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; text-align: center; }}
        th, td {{ border: 1px solid #7F7F7F; padding: 6px 4px; font-weight: normal; color: #000000; }}
        .bg-title {{ background-color: #D9EAD3; font-size: 16px; font-weight: bold; padding: 8px 0; }}
        .bg-header {{ background-color: #F2F2F2; font-weight: bold; }}
        .bg-person {{ background-color: #D0E0E3; font-weight: bold; font-size: 16px; }}
        .bg-total {{ background-color: #D9EAD3; font-weight: bold; }}
        .num {{ font-family: "Times New Roman", serif; }}
        .zh {{ font-family: SimSun, serif; }}
    </style></head><body><div class="table-container"><table>
        <tr><td colspan="{total_colspan}" class="bg-title">2026年9月招生數據動態表({range_title})</td></tr>
        <tr><th rowspan="2" class="bg-header">序號</th><th rowspan="2" class="bg-header">專業/基礎名稱</th><th rowspan="2" class="bg-header">目標人數</th>{person_headers}<th colspan="2" class="bg-person">其他人員</th><th rowspan="2" class="bg-header">目標人數</th><th rowspan="2" class="bg-header">實際完成</th><th rowspan="2" class="bg-header">與目標之差</th></tr>
        <tr>{sub_headers}</tr>
        {rows_html}
        <tr><td class="bg-total"></td><td class="bg-total zh"><b>{sum_r['專業/基礎名稱']}</b></td><td class="bg-total num"><b>{sum_r['目標人數']}</b></td>{sum_person_cells}<td class="bg-total num"><b>0</b></td><td class="bg-total num"><b>{sum_r['其他人員_實際']}</b></td><td class="bg-total num"><b>{sum_r['目標人數']}</b></td><td class="bg-total num"><b>{sum_r['實際完成']}</b></td><td class="bg-total num"><b>{sum_r['與目標之差']}</b></td></tr>
        <tr><td class="bg-total"></td><td class="bg-total zh"><b>{diff_r['專業/基礎名稱']}</b></td><td class="bg-total"></td>{diff_person_cells}<td class="bg-total" colspan="2"><b class="num">{diff_r['其他人員_實際']}</b></td><td class="bg-total"></td><td class="bg-total"></td><td class="bg-total num"><b>{diff_r['與目標之差']}</b></td></tr>
        <tr><td class="bg-total" colspan="{total_colspan - 2}"></td><td class="bg-total zh"><b>{rate_r['專業/基礎名稱']}</b></td><td class="bg-total num"><b>{rate_r['實際完成']}</b></td></tr>
    </table></div></body></html>
    """

st.markdown(f"### 📝 2026年9月招生數據動態表({selected_time_range})")
st.components.v1.html(build_html_document(calc_df, sum_row, diff_row, rate_row, PERSONS, RAW_PERSONS, selected_time_range), height=680, scrolling=True)

# -----------------------------------------------------------------------------
# 8. 可視化圖表展示
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("### 📊 基礎指標分析")
c1, c2 = st.columns(2)

with c1:
    fig_person = go.Figure(data=[
        go.Bar(name="目標人數", x=PERSONS, y=[sum_row[f"{p}_目標"] for p in RAW_PERSONS], marker_color="#0B3C5D", text=[sum_row[f"{p}_目標"] for p in RAW_PERSONS], textposition="outside"),
        go.Bar(name="實際完成", x=PERSONS, y=[sum_row[f"{p}_實際"] for p in RAW_PERSONS], marker_color="#FF3D00", text=[sum_row[f"{p}_實際"] for p in RAW_PERSONS], textposition="outside")
    ])
    fig_person.update_layout(title="<b>各成員目標 vs 實際完成對比</b>", barmode="group", plot_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20), yaxis=dict(gridcolor="#E0E0E0"))
    st.plotly_chart(fig_person, use_container_width=True)

with c2:
    top_majors = calc_df.sort_values(by="實際完成", ascending=False).head(8)
    fig_major = px.bar(top_majors, x="實際完成", y="專業/基礎名稱", orientation="h", title="<b>招生完成人數 Top 8 專業/基礎</b>", text="實際完成", color="專業/基礎名稱", color_discrete_sequence=px.colors.qualitative.Bold)
    fig_major.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False, plot_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20), xaxis=dict(gridcolor="#E0E0E0"))
    fig_major.update_traces(textposition="outside")
    st.plotly_chart(fig_major, use_container_width=True)

# -----------------------------------------------------------------------------
# 9. 動態趨勢與人員貢獻構成分析 (完全還原)
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("### 🔄 動態趨勢與人員貢獻構成分析")

ctrl_c1, ctrl_c2 = st.columns(2)
with ctrl_c1:
    person_mode = st.radio("選擇分析視角：", ["全體人員", "單人獨立分析", "多人員對比分析"], horizontal=True)

with ctrl_c2:
    if person_mode == "單人獨立分析":
        selected_person_disp = st.selectbox("選擇分析成員：", PERSONS + ["其他人員"])
    elif person_mode == "多人員對比分析":
        selected_persons_disp = st.multiselect("選擇對比成員：", PERSONS + ["其他人員"], default=PERSONS[:3] if len(PERSONS)>=3 else PERSONS)
    else:
        selected_person_disp = "全體人員"

time_records = []
for d_idx, d in enumerate(DATES):
    w = WEEKDAYS[d_idx]
    is_weekend = "週末" if w in ["星期六", "星期日"] else "工作日"
    d_dict = {p: 0 for p in RAW_PERSONS + ["其他人員"]}
    for item in st.session_state["daily_deltas"].get(d, []):
        m, col, val = item if len(item) == 3 else ("電氣基礎", item[0], item[1])
        p_name = col.replace("_實際", "")
        if p_name in d_dict: d_dict[p_name] += val
    
    for p_name, val in d_dict.items():
        time_records.append({"日期": d, "星期": w, "類型": is_weekend, "人員": alias_map.get(p_name, p_name), "新增報名數": val})

df_time_series = pd.DataFrame(time_records)
df_time_series['日期'] = pd.Categorical(df_time_series['日期'], categories=DATES, ordered=True)
df_time_series = df_time_series.sort_values('日期')

if time_granularity_type == "按日（單日切片）":
    df_time_series = df_time_series[df_time_series["日期"] == selected_time_range]

chart_col1, chart_col2 = st.columns(2)
is_single_day = (time_granularity_type == "按日（單日切片）")

with chart_col1:
    if person_mode == "全體人員":
        df_chart_line = df_time_series.groupby("日期", as_index=False, sort=False)["新增報名數"].sum()
        fig_line = px.bar(df_chart_line, x="日期", y="新增報名數", title=f"📊 <b>{selected_time_range} 全體新增總量</b>", text="新增報名數", color_discrete_sequence=["#D50000"]) if is_single_day else px.line(df_chart_line, x="日期", y="新增報名數", markers=True, title="📈 <b>全體人員招生趨勢 (按日明細)</b>", text="新增報名數")
    elif person_mode == "單人獨立分析":
        df_sub = df_time_series[df_time_series["人員"] == selected_person_disp]
        df_chart_line = df_sub.groupby("日期", as_index=False, sort=False)["新增報名數"].sum()
        fig_line = px.bar(df_chart_line, x="日期", y="新增報名數", title=f"📊 <b>【{selected_person_disp}】{selected_time_range} 新增量</b>", text="新增報名數", color_discrete_sequence=["#2962FF"]) if is_single_day else px.line(df_chart_line, x="日期", y="新增報名數", markers=True, title=f"📈 <b>【{selected_person_disp}】趨勢 (按日明細)</b>", text="新增報名數")
    else:
        df_sub = df_time_series[df_time_series["人員"].isin(selected_persons_disp)]
        df_chart_line = df_sub.groupby(["日期", "人員"], as_index=False, sort=False)["新增報名數"].sum()
        fig_line = px.bar(df_chart_line, x="人員", y="新增報名數", color="人員", title=f"📊 <b>{selected_time_range} 多人招生對比</b>", text="新增報名數") if is_single_day else px.line(df_chart_line, x="日期", y="新增報名數", color="人員", markers=True, title="📈 <b>多人招生趨勢對比 (按日明細)</b>")

    fig_line.update_layout(plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20), yaxis=dict(gridcolor="#E0E0E0"), showlegend=(not is_single_day))
    st.plotly_chart(fig_line, use_container_width=True)

with chart_col2:
    if person_mode == "全體人員":
        df_pie = df_time_series.groupby("人員", as_index=False)["新增報名數"].sum()
        fig_pie = px.pie(df_pie[df_pie["新增報名數"] > 0], values="新增報名數", names="人員", title="🍩 <b>全體人員招生貢獻占比</b>", hole=0.4)
    elif person_mode == "單人獨立分析":
        inv_map = {v: k for k, v in alias_map.items()}
        raw_sel_p = inv_map.get(selected_person_disp, selected_person_disp)
        major_records = []
        for d in selected_dates_list:
            for item in st.session_state["daily_deltas"].get(d, []):
                major, col, val = item if len(item) == 3 else ("電氣基礎", item[0], item[1])
                if col.replace("_實際", "") == raw_sel_p: major_records.append({"專業/基礎": major, "新增人數": val})
        df_major_pie = pd.DataFrame(major_records)
        if not df_major_pie.empty:
            df_major_pie = df_major_pie.groupby("專業/基礎", as_index=False)["新增人數"].sum()
            fig_pie = px.pie(df_major_pie[df_major_pie["新增人數"] > 0], values="新增人數", names="專業/基礎", title=f"🍩 <b>【{selected_person_disp}】專業成交構成</b>", hole=0.4)
        else:
            fig_pie = go.Figure()
    else:
        df_pie = df_time_series[df_time_series["人員"].isin(selected_persons_disp)].groupby("人員", as_index=False)["新增報名數"].sum()
        fig_pie = px.pie(df_pie[df_pie["新增報名數"] > 0], values="新增報名數", names="人員", title="🍩 <b>對比成員占比構成</b>", hole=0.4)

    fig_pie.update_layout(paper_bgcolor="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20))
    st.plotly_chart(fig_pie, use_container_width=True)

# -----------------------------------------------------------------------------
# 10. Excel 完整帶樣式導出 (修復匯率與百分比寫入格式)
# -----------------------------------------------------------------------------
def export_color_excel(calc_df, sum_row, diff_row, rate_row, persons_disp, raw_persons):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "招生數據動態表"
    ws.views.sheetView[0].showGridLines = True

    BORDER_THIN = Border(left=Side(style="thin", color="7F7F7F"), right=Side(style="thin", color="7F7F7F"), top=Side(style="thin", color="7F7F7F"), bottom=Side(style="thin", color="7F7F7F"))
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

    total_cols = 7 + (len(raw_persons) + 1) * 2
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    ws["A1"] = f"2026年9月招生數據動態表({selected_time_range})"
    ws["A1"].font = Font(name="SimSun", size=14, bold=True)
    ws["A1"].fill = PatternFill(start_color="D9EAD3", fill_type="solid")
    ws["A1"].alignment = ALIGN_CENTER

    ws.merge_cells("A2:A3"); ws["A2"] = "序號"
    ws.merge_cells("B2:B3"); ws["B2"] = "專業/基礎名稱"
    ws.merge_cells("C2:C3"); ws["C2"] = "目標人數"

    col_idx = 4
    for p in persons_disp + ["其他人員"]:
        ws.merge_cells(start_row=2, start_column=col_idx, end_row=2, end_column=col_idx + 1)
        ws.cell(row=2, column=col_idx, value=p).fill = PatternFill(start_color="D0E0E3", fill_type="solid")
        ws.cell(row=3, column=col_idx, value="目標人數")
        ws.cell(row=3, column=col_idx + 1, value="實際完成")
        col_idx += 2

    other_col = col_idx
    for name, offset in [("目標人數", 0), ("實際完成", 1), ("與目標之差", 2)]:
        ws.merge_cells(start_row=2, start_column=other_col+offset, end_row=3, end_column=other_col+offset)
        ws.cell(row=2, column=other_col+offset, value=name)

    for r in range(2, 4):
        for c in range(1, total_cols + 1):
            cell = ws.cell(row=r, column=c)
            if not cell.fill.start_color.rgb: cell.fill = PatternFill(start_color="EFEFEF", fill_type="solid")
            cell.font = Font(name="SimSun", size=10, bold=True)
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN

    curr_r = 4
    for idx, row in calc_df.iterrows():
        ws.cell(row=curr_r, column=1, value=row["序號"])
        ws.cell(row=curr_r, column=2, value=row["專業/基礎名稱"])
        ws.cell(row=curr_r, column=3, value=row["目標人數"])

        c_offset = 4
        for rp in raw_persons:
            ws.cell(row=curr_r, column=c_offset, value=row[f"{rp}_目標"])
            ws.cell(row=curr_r, column=c_offset + 1, value=row[f"{rp}_實際"] or "")
            c_offset += 2

        ws.cell(row=curr_r, column=c_offset, value=row.get("其他人員_目標", 0) or "")
        ws.cell(row=curr_r, column=c_offset + 1, value=row.get("其他人員_實際", 0) or "")
        ws.cell(row=curr_r, column=c_offset + 2, value=row["目標人數"])
        ws.cell(row=curr_r, column=c_offset + 3, value=row["實際完成"])
        ws.cell(row=curr_r, column=c_offset + 4, value=row["與目標之差"])

        for c in range(1, total_cols + 1):
            cell = ws.cell(row=curr_r, column=c)
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN
        curr_r += 1

    # 渲染底部 3 行彙總
    for r_data in [sum_row, diff_row]:
        ws.cell(row=curr_r, column=2, value=r_data["專業/基礎名稱"])
        ws.cell(row=curr_r, column=3, value=r_data.get("目標人數", ""))

        c_offset = 4
        for rp in raw_persons:
            if r_data["專業/基礎名稱"] == "與目標之差":
                ws.merge_cells(start_row=curr_r, start_column=c_offset, end_row=curr_r, end_column=c_offset+1)
                ws.cell(row=curr_r, column=c_offset, value=r_data.get(f"{rp}_實際", ""))
            else:
                ws.cell(row=curr_r, column=c_offset, value=r_data.get(f"{rp}_目標", ""))
                ws.cell(row=curr_r, column=c_offset + 1, value=r_data.get(f"{rp}_實際", ""))
            c_offset += 2

        if r_data["專業/基礎名稱"] == "與目標之差":
            ws.merge_cells(start_row=curr_r, start_column=c_offset, end_row=curr_r, end_column=c_offset+1)
            ws.cell(row=curr_r, column=c_offset, value=r_data.get("其他人員_實際", ""))
        else:
            ws.cell(row=curr_r, column=c_offset, value=0)
            ws.cell(row=curr_r, column=c_offset + 1, value=r_data.get("其他人員_實際", ""))

        ws.cell(row=curr_r, column=c_offset + 2, value=r_data.get("目標人數", ""))
        ws.cell(row=curr_r, column=c_offset + 3, value=r_data.get("實際完成", ""))
        ws.cell(row=curr_r, column=c_offset + 4, value=r_data.get("與目標之差", ""))

        for c in range(1, total_cols + 1):
            cell = ws.cell(row=curr_r, column=c)
            cell.alignment = ALIGN_CENTER; cell.border = BORDER_THIN; cell.fill = PatternFill(start_color="D9EAD3", fill_type="solid")
        curr_r += 1

    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=total_cols - 2)
    ws.cell(row=curr_r, column=total_cols - 1, value=rate_row["專業/基礎名稱"])
    ws.cell(row=curr_r, column=total_cols, value=rate_row["實際完成"])
    for c in range(1, total_cols + 1):
        cell = ws.cell(row=curr_r, column=c)
        cell.alignment = ALIGN_CENTER; cell.border = BORDER_THIN; cell.fill = PatternFill(start_color="D9EAD3", fill_type="solid")

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.markdown("---")
st.markdown("### 📥 數據導出")
st.download_button(
    label="📥 導出帶樣式的 Excel 報表",
    data=export_color_excel(calc_df, sum_row, diff_row, rate_row, PERSONS, RAW_PERSONS),
    file_name=f"2026年9月招生數據動態表_{selected_time_range}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
