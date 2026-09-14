import io
import zipfile
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 1. 页面基本配置与全局 UI 样式注入
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
        height: 105px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-sizing: border-box;
    }
    
    .kpi-title {
        font-size: 14px;
        color: #31333F;
        font-weight: 500;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    
    .kpi-body {
        display: flex;
        align-items: baseline;
        gap: 10px;
    }
    
    .kpi-value {
        font-size: 28px;
        font-weight: 700;
        color: #0E1117;
        font-family: "Source Sans Pro", sans-serif;
        line-height: 1;
    }
    
    .kpi-delta {
        font-size: 13px;
        font-weight: 600;
        color: #FF2B2B;
        background-color: #FFE6E6;
        padding: 2px 8px;
        border-radius: 12px;
        display: inline-flex;
        align-items: center;
        white-space: nowrap;
    }

    div[data-testid="stForm"] {
        border-radius: 10px;
        background-color: #FAFAFA;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 招生数据动态管理与多维分析系统")
st.caption("2026年9月数据 - 修正对齐版")
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
WEEKDAYS = ["星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "星期一", 
            "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
AVAL_ANIMALS = ["🦊", "🐼", "🦁", "🐰", "🐯", "🐱", "🐶", "🐻", "🐨", "🐮", "🐵", "🐥"]

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

    # 精确匹配全量增量数据
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
# 3. 侧边栏：脱敏管理与转换修复
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ 数据管理与设置")

enable_anonymize = st.sidebar.checkbox("开启数据脱敏 / 纯动物符号模式", value=True)

alias_map = {}
for p in RAW_PERSONS:
    if enable_anonymize:
        alias_map[p] = st.session_state["person_animals"].get(p, "🐱")
    else:
        alias_map[p] = p

PERSONS = [alias_map[p] for p in RAW_PERSONS]

if enable_anonymize:
    with st.sidebar.expander("👁️ 视角对照表（管理者隐私预览）", expanded=False):
        mapping_df = pd.DataFrame({
            "真实姓名": RAW_PERSONS,
            "代称动物": PERSONS
        })
        st.dataframe(mapping_df, hide_index=True, use_container_width=True)

# -----------------------------------------------------------------------------
# 4. 修复后的数据计算逻辑
# -----------------------------------------------------------------------------
def get_processed_df_by_dates(dates_list):
    df_result = st.session_state["base_targets"].copy()
    
    # 强制清零实际列，保证完全由选定区间的增量累加
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

# -----------------------------------------------------------------------------
# 5. 时间汇总与 KPI
# -----------------------------------------------------------------------------
selected_dates_list = DATES
calc_df = get_processed_df_by_dates(selected_dates_list)

act_cols = [c for c in calc_df.columns if c.endswith("_实际")]
calc_df["实际完成"] = calc_df[act_cols].sum(axis=1)
calc_df["与目标之差"] = calc_df["实际完成"] - calc_df["目标人数"]

total_target_cum = calc_df["目标人数"].sum()
total_actual_cum = calc_df["实际完成"].sum()
cum_rate_val = (total_actual_cum / total_target_cum * 100) if total_target_cum > 0 else 0

m_col1, m_col2, m_col3 = st.columns(3)
with m_col1:
    st.metric("🎯 总目标人数", f"{total_target_cum} 人")
with m_col2:
    st.metric("✅ 实际完成人数", f"{total_actual_cum} 人", delta=f"{total_actual_cum - total_target_cum} 人")
with m_col3:
    st.metric("📈 目标完成比例", f"{cum_rate_val:.2f}%")

st.markdown("---")
st.subheader("📝 2026年9月1日-13日全量数据汇总表")

st.dataframe(calc_df, use_container_width=True)
