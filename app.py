import streamlit as st
import pandas as pd
import plotly.express as px
import json

# 1. 页面基本配置
st.set_page_config(page_title="招生数据动态管理与多维分析系统", layout="wide", initial_sidebar_state="expanded")

st.title("📊 招生数据动态管理与多维分析系统")
st.caption("高级优化版：含核心KPI看板、全员/单人趋势分析、可视化进度图表、动态数据录入及备份还原")
st.markdown("---")

# 定义人员列表与专业列表
PERSONS = ["Person A", "Person B", "Person C", "Person D", "Person E"]
MAJORS = [
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
    ("结构基础", 9, [2, 2, 1, 3, 1]),
    ("公共基础", 4, [2, 1, 1, 0, 0]),
    ("道路基础", 6, [1, 1, 1, 1, 2]),
    ("水利水电基础", 10, [2, 2, 3, 2, 1]),
]

DATES = ["9月1日", "9月2日", "9月3日", "9月4日", "9月5日", "9月6日"]

# 2. 初始数据构造函数（用于初始化和重置）
def init_default_data():
    base_data = []
    for idx, (name, total_target, person_tgts) in enumerate(MAJORS, 1):
        row = {"序号": idx, "专业/基础名称": name, "目标人数": total_target}
        for p, tgt in zip(PERSONS, person_tgts):
            row[f"{p}_目标"] = tgt
            row[f"{p}_实际"] = 0
        row["其他人员_实际"] = 0
        base_data.append(row)
    st.session_state['base_targets'] = pd.DataFrame(base_data)

    st.session_state['daily_deltas'] = {
        "9月1日": [("电气基础", "Person A_实际", 1)],
        "9月2日": [("环保专业", "Person A_实际", 1), ("环保专业", "Person B_实际", 1), ("电气基础", "Person A_实际", 1), ("发输电专业", "其他人员_实际", 1), ("233网校", "其他人员_实际", 1)],
        "9月3日": [("环评专业", "Person B_实际", 1), ("暖通专业", "Person C_实际", 1), ("环保基础", "Person B_实际", 1), ("环保基础", "其他人员_实际", 1)],
        "9月4日": [("电气基础", "Person A_实际", 1), ("电气基础", "Person B_实际", 1), ("环保基础", "Person A_实际", 1), ("水利水电基础", "其他人员_实际", 1)],
        "9月5日": [("给排水专业", "Person C_实际", 1), ("暖通专业", "Person E_实际", 1), ("岩土基础", "Person C_实际", 1), ("暖通基础", "Person E_实际", 1)],
        "9月6日": [("给排水专业", "Person C_实际", 1), ("环保专业", "其他人员_实际", 1), ("岩土专业", "其他人员_实际", 1), ("暖通专业", "Person E_实际", 1), ("环保基础", "Person A_实际", 1), ("环保基础", "其他人员_实际", 1)]
    }

# 初始化 Session State
if 'base_targets' not in st.session_state or 'daily_deltas' not in st.session_state:
    init_default_data()

# 3. ⚙️ 数据管理控制台（数据录入 & 备份还原）
with st.expander("⚙️ 数据管理控制台（录入出单 / 备份与还原）"):
    tab_entry, tab_backup = st.tabs(["➕ 录入单日新增", "💾 数据备份与还原"])
    
    # 子标签页 1：数据录入
    with tab_entry:
        e_col1, e_col2, e_col3, e_col4 = st.columns(4)
        with e_col1:
            in_date = st.selectbox("选择日期", DATES, key="in_date_sel")
        with e_col2:
            in_major = st.selectbox("选择专业", [m[0] for m in MAJORS], key="in_major_sel")
        with e_col3:
            in_person = st.selectbox("归属人员", PERSONS + ["其他人员"], key="in_person_sel")
        with e_col4:
            in_val = st.number_input("新增人数", min_value=1, value=1, step=1, key="in_val_num")
            
        if st.button("🚀 提交新增数据", use_container_width=True):
            col_name = "其他人员_实际" if in_person == "其他人员" else f"{in_person}_实际"
            st.session_state['daily_deltas'][in_date].append((in_major, col_name, in_val))
            st.success(f"成功为 {in_date} 的【{in_major}】增加 {in_person} {in_val} 人！数据已自动更新。")
            st.rerun()

    # 子标签页 2：备份与还原
    with tab_backup:
        b_col1, b_col2, b_col3 = st.columns(3)
        
        # 导出备份
        with b_col1:
            st.markdown("##### 1. 下载备份配置文件")
            backup_json = json.dumps(st.session_state['daily_deltas'], ensure_ascii=False, indent=2)
            st.download_button(
                label="📥 导出当前数据备份 (.json)",
                data=backup_json,
                file_name="enrollment_data_backup.json",
                mime="application/json",
                use_container_width=True
            )
            
        # 导入还原
        with b_col2:
            st.markdown("##### 2. 导入历史备份文件")
            uploaded_file = st.file_uploader("上传 .json 备份文件", type=["json"], label_visibility="collapsed")
            if uploaded_file is not None:
                try:
                    loaded_data = json.load(uploaded_file)
                    # 校验数据结构
                    if isinstance(loaded_data, dict):
                        st.session_state['daily_deltas'] = loaded_data
                        st.success("🎉 数据已成功恢复！")
                        st.rerun()
                    else:
                        st.error("备份文件格式不符合要求！")
                except Exception as e:
                    st.error(f"解析失败: {str(e)}")

        # 重置默认
        with b_col3:
            st.markdown("##### 3. 恢复至系统初始状态")
            if st.button("⚠️ 一键重置所有修改", type="secondary", use_container_width=True):
                init_default_data()
                st.warning("所有数据已恢复至初始默认状态。")
                st.rerun()

st.markdown("---")

# 4. 动态合成数据函数
def get_processed_df(selected_date):
    df_result = st.session_state['base_targets'].copy()
    if selected_date == "📅 当月累计数据（截至9月6日）":
        for d in DATES:
            for major, col, val in st.session_state['daily_deltas'].get(d, []):
                df_result.loc[df_result['专业/基础名称'] == major, col] += val
    else:
        for major, col, val in st.session_state['daily_deltas'].get(selected_date, []):
            df_result.loc[df_result['专业/基础名称'] == major, col] += val
    return df_result

# 5. 顶部视图选择
st.subheader("🗓️ 数据视图切换")
view_options = [f"{d}单日" for d in DATES] + ["📅 当月累计数据（截至9月6日）"]
date_option = st.radio("请选择查看的时间节点：", view_options, horizontal=True)

raw_date_name = date_option.replace("单日", "")
current_df = get_processed_df(raw_date_name)

# 6. 表格计算与底部汇总
calc_df = current_df.copy()
act_cols = [c for c in calc_df.columns if c.endswith("_实际")]
calc_df["实际完成"] = calc_df[act_cols].sum(axis=1)
calc_df["与目标之差"] = calc_df["实际完成"] - calc_df["目标人数"]

ordered_cols = ["序号", "专业/基础名称", "目标人数"]
for p in PERSONS:
    ordered_cols.extend([f"{p}_目标", f"{p}_实际"])
ordered_cols.extend(["其他人员_实际", "目标人数", "实际完成", "与目标之差"])

seen = set()
final_cols = [c for c in ordered_cols if not (c in seen or seen.add(c))]
calc_df = calc_df[final_cols]

num_cols = [c for c in calc_df.columns if c not in ["序号", "专业/基础名称"]]

sum_row = {"序号": "", "专业/基础名称": "合计"}
for c in num_cols:
    sum_row[c] = int(calc_df[c].sum())

diff_row = {"序号": "", "专业/基础名称": "与目标之差"}
for p in PERSONS:
    diff_row[f"{p}_目标"] = ""
    diff_row[f"{p}_实际"] = sum_row[f"{p}_实际"] - sum_row[f"{p}_目标"]

diff_row["其他人员_实际"] = ""
diff_row["目标人数"] = ""
diff_row["实际完成"] = ""
diff_row["与目标之差"] = sum_row["与目标之差"]

total_actual_cum = sum_row["实际完成"]
total_target_cum = sum_row["目标人数"]
rate_val = (total_actual_cum / total_target_cum * 100) if total_target_cum > 0 else 0

rate_row = {"序号": "", "专业/基础名称": "2026年8月目标人数完成比例"}
for c in num_cols:
    rate_row[c] = ""
rate_row["实际完成"] = f"{rate_val:.2f}%"

# 7. 核心指标 KPI 看板
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("🎯 团队目标总人数", f"{total_target_cum} 人")
kpi2.metric("✅ 实际完成人数", f"{total_actual_cum} 人")
kpi3.metric("📉 与目标之差", f"{sum_row['与目标之差']} 人")
kpi4.metric("📊 总体目标完成率", f"{rate_val:.2f}%")

st.markdown("---")

# 8. 在线数据表展示
st.markdown("### 📝 招生数据明细表（平铺横向可滚动）")

display_df = pd.concat([calc_df, pd.DataFrame([sum_row, diff_row, rate_row])], ignore_index=True)
display_df = display_df.astype(str).replace({"nan": "", "None": ""})

st.dataframe(display_df, use_container_width=True, hide_index=True)

st.markdown("---")

# 9. 趋势分析（全员/单人）
st.subheader("📈 每日新增 vs 当月累计趋势分析")

trend_options = ["🌐 全体人员（团队汇总）"] + PERSONS
sel_target = st.selectbox("请选择要分析的对象：", trend_options)

person_daily = []
cum_acc = 0

for d in DATES:
    deltas = st.session_state['daily_deltas'].get(d, [])
    if sel_target == "🌐 全体人员（团队汇总）":
        val = sum(v for m, col, v in deltas)
    else:
        val = sum(v for m, col, v in deltas if col == f"{sel_target}_实际")
        
    cum_acc += val
    person_daily.append({"日期": d, "单日新增": val, "当月累计": cum_acc})

df_p = pd.DataFrame(person_daily)

if sel_target == "🌐 全体人员（团队汇总）":
    p_tgt = st.session_state['base_targets']["目标人数"].sum()
    title_suffix = "全体团队"
else:
    p_tgt = st.session_state['base_targets'][f"{sel_target}_目标"].sum()
    title_suffix = sel_target

col_p1, col_p2 = st.columns(2)
with col_p1:
    fig_p_daily = px.bar(df_p, x="日期", y="单日新增", text_auto=True, title=f"{title_suffix} - 每日新增分布", color_discrete_sequence=['#2b5c8f'])
    st.plotly_chart(fig_p_daily, use_container_width=True)

with col_p2:
    fig_p_cum = px.line(df_p, x="日期", y="当月累计", markers=True, text="当月累计", title=f"{title_suffix} - 累计进度（总目标: {p_tgt} 人）")
    st.plotly_chart(fig_p_cum, use_container_width=True)

st.markdown("---")

# 10. 全员横向对比
st.subheader("👥 人员间横向对比")

comp_list = []
for p in PERSONS + ["其他人员"]:
    row_data = {"人员": p}
    cum_val = 0
    for d in DATES:
        deltas = st.session_state['daily_deltas'].get(d, [])
        val = sum(v for m, col, v in deltas if col == f"{p}_实际")
        row_data[d] = val
        cum_val += val
    row_data["当月累计"] = cum_val
    tgt_col = f"{p}_目标"
    tgt_val = st.session_state['base_targets'][tgt_col].sum() if tgt_col in st.session_state['base_targets'] else 0
    row_data["个人目标"] = tgt_val
    row_data["完成率(%)"] = round((cum_val / tgt_val * 100), 2) if tgt_val > 0 else 0
    comp_list.append(row_data)

df_comp = pd.DataFrame(comp_list)

col_c1, col_c2 = st.columns(2)
with col_c1:
    fig_comp_cum = px.bar(df_comp, x="人员", y=["当月累计", "个人目标"], barmode="group", title="各人员累计完成人数 vs 个人目标")
    st.plotly_chart(fig_comp_cum, use_container_width=True)

with col_c2:
    fig_rate = px.bar(df_comp[df_comp["人员"] != "其他人员"], x="人员", y="完成率(%)", text_auto=True, title="各员工目标完成率 (%)", color="完成率(%)", color_continuous_scale="Blues")
    st.plotly_chart(fig_rate, use_container_width=True)

st.markdown("---")

# 11. 占比分析
st.subheader("🥧 团队贡献 & 专业占比")

col_pie1, col_pie2 = st.columns(2)
with col_pie1:
    fig_pie_person = px.pie(df_comp, values="当月累计", names="人员", title="人员完成贡献占比", hole=0.3)
    st.plotly_chart(fig_pie_person, use_container_width=True)

with col_pie2:
    df_prof_pie = calc_df[calc_df["实际完成"] > 0]
    fig_pie_prof = px.pie(df_prof_pie, values="实际完成", names="专业/基础名称", title="各专业已完成招生分布", hole=0.3)
    st.plotly_chart(fig_pie_prof, use_container_width=True)