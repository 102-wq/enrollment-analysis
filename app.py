import streamlit as st
import pandas as pd
import numpy as np

# ==============================================================================
# 1. 页面基础配置
# ==============================================================================
st.set_page_config(
    page_title="招生数据智能分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义 CSS 样式
st.markdown("""
<style>
    .main-header {
        font-size: 24px;
        font-weight: bold;
        color: #1E293B;
        margin-bottom: 20px;
    }
    .stTable {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. 基础数据配置与 Session State 初始化
# ==============================================================================
RAW_PERSONS = ["覃小燕", "张三", "李四", "王五"]
MAJORS = ["环境影响评价工程师", "一级建造师", "注册安全工程师", "消防工程师"]
DATES = ["9月1日", "9月2日", "9月3日", "9月4日", "9月5日", "9月6日", "9月7日"]

# 侧边栏：脱敏设置
st.sidebar.title("⚙️ 系统配置")
enable_anonymize = st.sidebar.checkbox("开启数据脱敏 / 匿名模式", value=False)

# 构建人员映射关系
if enable_anonymize:
    person_alias_map = {name: f"咨询顾问 {chr(65 + i)}" for i, name in enumerate(RAW_PERSONS)}
else:
    person_alias_map = {name: name for name in RAW_PERSONS}

DISPLAY_PERSONS = [person_alias_map[p] for p in RAW_PERSONS]

# 初始化基础目标数据 (Base Targets)
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

# 初始化每日增量记录
if "daily_deltas" not in st.session_state:
    st.session_state["daily_deltas"] = {d: [] for d in DATES}

# ==============================================================================
# 3. 辅助函数 (安全列计算，防 KeyError)
# ==============================================================================
def get_person_col(df: pd.DataFrame, person_raw_name: str, col_type: str) -> str:
    """
    根据原始姓名和列类型 (目标/实际)，安全查找 DataFrame 中对应的真实列名
    """
    raw_col = f"{person_raw_name}_{col_type}"
    if raw_col in df.columns:
        return raw_col
    
    # 尝试查找脱敏后的列名
    alias_name = person_alias_map.get(person_raw_name, person_raw_name)
    alias_col = f"{alias_name}_{col_type}"
    if alias_col in df.columns:
        return alias_col
        
    return None

def compute_processed_data(selected_date):
    """
    根据选定日期合并计算最终的 DataFrame，防止脱敏导致 KeyError
    """
    df_calc = st.session_state["base_targets"].copy()
    
    # 重命名列以符合当前脱敏设置
    rename_dict = {}
    for p in RAW_PERSONS:
        alias = person_alias_map[p]
        rename_dict[f"{p}_目标"] = f"{alias}_目标"
        rename_dict[f"{p}_实际"] = f"{alias}_实际"
    df_calc = df_calc.rename(columns=rename_dict)
    
    # 叠加增量数据
    if selected_date == "📅 当月累计数据（截至9月7日）":
        target_dates = DATES
    else:
        target_dates = [selected_date]
        
    for d in target_dates:
        for major, col_raw, val in st.session_state["daily_deltas"].get(d, []):
            # 将原始列名映射到当前的 DataFrame 列名
            target_col = col_raw
            for raw_p, alias_p in person_alias_map.items():
                if raw_p in col_raw:
                    target_col = col_raw.replace(raw_p, alias_p)
                    break
            
            if target_col in df_calc.columns:
                df_calc.loc[df_calc["专业/基础名称"] == major, target_col] += val
                
    return df_calc

def build_html_document(df, sum_row, diff_row, rate_row):
    """
    构建符合表格展示要求的 HTML 代码
    """
    html = """
    <style>
        .custom-table { width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px; }
        .custom-table th, .custom-table td { border: 1px solid #E2E8F0; padding: 10px; text-align: center; }
        .custom-table th { background-color: #F8FAFC; font-weight: bold; }
        .highlight-row { background-color: #F1F5F9; font-weight: bold; }
        .rate-row { background-color: #EFF6FF; color: #1D4ED8; font-weight: bold; }
    </style>
    <table class="custom-table">
        <thead>
            <tr>
                <th>专业/基础名称</th>
                <th>总目标</th>
    """
    for p in DISPLAY_PERSONS:
        html += f"<th>{p} (目标)</th><th>{p} (实际)</th>"
    html += "<th>其他人员(实际)</th></tr></thead><tbody>"
    
    for _, row in df.iterrows():
        html += f"<tr><td>{row['专业/基础名称']}</td><td>{row['总目标']}</td>"
        for p in DISPLAY_PERSONS:
            html += f"<td>{row.get(f'{p}_目标', 0)}</td><td>{row.get(f'{p}_实际', 0)}</td>"
        html += f"<td>{row.get('其他人员_实际', 0)}</td></tr>"
        
    # 合计行
    html += f"<tr class='highlight-row'><td>{sum_row['专业/基础名称']}</td><td>{sum_row['总目标']}</td>"
    for p in DISPLAY_PERSONS:
        html += f"<td>{sum_row.get(f'{p}_目标', 0)}</td><td>{sum_row.get(f'{p}_实际', 0)}</td>"
    html += f"<td>{sum_row.get('其他人员_实际', 0)}</td></tr>"
    
    # 达成率行
    html += f"<tr class='rate-row'><td>{rate_row['专业/基础名称']}</td><td>{rate_row['总目标']}</td>"
    for p in DISPLAY_PERSONS:
        html += f"<td>-</td><td>{rate_row.get(f'{p}_实际', '0%')}</td>"
    html += f"<td>-</td></tr>"
    
    html += "</tbody></table>"
    return html

# ==============================================================================
# 4. 主界面渲染
# ==============================================================================
st.markdown("<div class='main-header'>📊 招生数据智能分析系统</div>", unsafe_allow_html=True)

# 日期选择器
date_options = ["📅 当月累计数据（截至9月7日）"] + DATES
selected_date = st.selectbox("📅 选择查看的时间节点：", date_options)

# 计算当前合并后的 DataFrame
current_df = compute_processed_data(selected_date)

# 计算汇总指标
sum_series = {"专业/基础名称": "合计", "总目标": current_df["总目标"].sum()}
rate_series = {"专业/基础名称": "达成率", "总目标": "100%"}
diff_series = {"专业/基础名称": "差额", "总目标": 0}

total_actual_sum = 0
for p in DISPLAY_PERSONS:
    t_val = current_df[f"{p}_目标"].sum()
    a_val = current_df[f"{p}_实际"].sum()
    sum_series[f"{p}_目标"] = t_val
    sum_series[f"{p}_实际"] = a_val
    rate_series[f"{p}_实际"] = f"{(a_val / t_val * 100):.1f}%" if t_val > 0 else "0.0%"
    total_actual_sum += a_val

other_act = current_df["其他人员_实际"].sum() if "其他人员_实际" in current_df.columns else 0
sum_series["其他人员_实际"] = other_act
total_actual_sum += other_act

# 顶部 KPI 卡片
kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
with kpi_col1:
    st.metric(label="月度总目标", value=f"{sum_series['总目标']} 人")
with kpi_col2:
    st.metric(label="当前累计完成", value=f"{total_actual_sum} 人")
with kpi_col3:
    overall_rate = (total_actual_sum / sum_series['总目标'] * 100) if sum_series['总目标'] > 0 else 0
    st.metric(label="总体完成率", value=f"{overall_rate:.1f}%")

st.markdown("---")

# 数据表格渲染（改用 Streamlit 1.30+ 推荐的 st.html 替代 components.html）
st.markdown("### 📝 2026年9月招生数据动态表")
html_code = build_html_document(current_df, sum_series, diff_series, rate_series)
st.html(html_code)

st.markdown("---")

# ==============================================================================
# 5. 录入与交互组件（全面更新 width="stretch" 适配新版 API）
# ==============================================================================
st.markdown("### ✏️ 快速录入日增量数据")

with st.expander("点击展开增量数据录入窗口", expanded=False):
    col_input1, col_input2, col_input3 = st.columns(3)
    
    with col_input1:
        entry_date = st.selectbox("录入日期", DATES)
    with col_input2:
        entry_major = st.selectbox("选择专业", MAJORS)
    with col_input3:
        entry_person = st.selectbox("责任顾问", DISPLAY_PERSONS)
        
    entry_val = st.number_input("今日新增招生人数", min_value=0, value=1, step=1)
    
    # 替换废弃参数 use_container_width=True -> width="stretch"
    if st.button("提交增量数据", width="stretch"):
        # 匹配回原始列名保存
        raw_person_name = RAW_PERSONS[DISPLAY_PERSONS.index(entry_person)]
        target_col_raw = f"{raw_person_name}_实际"
        
        st.session_state["daily_deltas"][entry_date].append((entry_major, target_col_raw, entry_val))
        st.success(f"成功为 {entry_date} - {entry_major} - {entry_person} 添加 {entry_val} 人！")
        st.rerun()

# 基础目标可编辑数据表
st.markdown("### ⚙️ 调整基础目标配置")
# 替换废弃参数 use_container_width=True -> width="stretch"
edited_df = st.data_editor(
    st.session_state["base_targets"],
    key="target_editor",
    width="stretch"
)

if st.button("保存目标配置修改", width="stretch"):
    st.session_state["base_targets"] = edited_df
    st.success("基础目标配置已更新！")
    st.rerun()
