import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# 1. 页面基本配置
st.set_page_config(
    page_title="招生数据动态管理与多维分析系统",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 招生数据动态管理与多维分析系统")
st.caption("2026年9月数据 - Edge 高兼容带颜色完美渲染版")
st.markdown("---")

# 2. 基础数据定义
PERSONS = ["人员A", "人员B", "人员C", "人员D", "人员E"]

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

DATES = ["9月1日", "9月2日", "9月3日", "9月4日", "9月5日", "9月6日", "9月7日"]


def init_default_data():
  base_data = []
  for idx, (name, total_target, person_tgts) in enumerate(MAJORS, 1):
    row = {"序号": idx, "专业/基础名称": name, "目标人数": total_target}
    for p, tgt in zip(PERSONS, person_tgts):
      row[f"{p}_目标"] = tgt
      row[f"{p}_实际"] = 0
    row["其他人员_实际"] = 0
    base_data.append(row)
  st.session_state["base_targets"] = pd.DataFrame(base_data)

  st.session_state["daily_deltas"] = {
      "9月1日": [("电气基础", "人员A_实际", 1)],
      "9月2日": [
          ("环保专业", "人员A_实际", 1),
          ("环保专业", "人员B_实际", 1),
          ("电气基础", "人员A_实际", 1),
          ("发输电专业", "其他人员_实际", 1),
          ("233网校", "其他人员_实际", 1),
      ],
      "9月3日": [
          ("环评专业", "人员B_实际", 1),
          ("暖通专业", "人员E_实际", 1),
          ("环保基础", "人员B_实际", 1),
          ("环保基础", "其他人员_实际", 1),
      ],
      "9月4日": [
          ("电气基础", "人员A_实际", 1),
          ("电气基础", "人员B_实际", 1),
          ("环保基础", "人员A_实际", 1),
          ("水利水电基础", "其他人员_实际", 1),
      ],
      "9月5日": [
          ("给排水专业", "人员C_实际", 1),
          ("暖通专业", "人员E_实际", 1),
          ("岩土基础", "人员C_实际", 1),
          ("暖通基础", "人员E_实际", 1),
      ],
      "9月6日": [
          ("给排水专业", "人员C_实际", 1),
          ("环保专业", "其他人员_实际", 1),
          ("岩土专业", "其他人员_实际", 1),
          ("暖通基础", "人员C_实际", 1),
          ("环保基础", "人员A_实际", 1),
          ("环保基础", "其他人员_实际", 1),
      ],
      "9月7日": [
          ("电气基础", "人员A_实际", 2),
          ("环保基础", "人员A_实际", 1),
          ("水基础", "人员A_实际", 1),
          ("环保基础", "人员B_实际", 1),
          ("暖通基础", "人员B_实际", 1),
          ("岩土基础", "人员C_实际", 1),
          ("暖通专业", "人员C_实际", 1),
          ("公共基础", "人员D_实际", 1),
          ("环保基础", "人员D_实际", 1),
      ],
  }


if (
    "base_targets" not in st.session_state
    or "daily_deltas" not in st.session_state
):
  init_default_data()


# 3. 视图切换与计算
def get_processed_df(selected_date):
  df_result = st.session_state["base_targets"].copy()
  if selected_date == "📅 当月累计数据（截至9月7日）":
    for d in DATES:
      for major, col, val in st.session_state["daily_deltas"].get(d, []):
        df_result.loc[df_result["专业/基础名称"] == major, col] += val
  else:
    for major, col, val in st.session_state["daily_deltas"].get(
        selected_date, []
    ):
      df_result.loc[df_result["专业/基础名称"] == major, col] += val
  return df_result


st.subheader("🗓️ 数据视图切换")
view_options = [f"{d}单日" for d in DATES] + [
    "📅 当月累计数据（截至9月7日）"
]
date_option = st.radio(
    "请选择查看的时间节点：",
    view_options,
    index=len(view_options) - 1,
    horizontal=True,
)

raw_date_name = date_option.replace("单日", "")
calc_df = get_processed_df(raw_date_name)

act_cols = [c for c in calc_df.columns if c.endswith("_实际")]
calc_df["实际完成"] = calc_df[act_cols].sum(axis=1)
calc_df["与目标之差"] = calc_df["实际完成"] - calc_df["目标人数"]

# 汇总计算
num_cols = [c for c in calc_df.columns if c not in ["序号", "专业/基础名称"]]
sum_row = {"专业/基础名称": "合计"}
for c in num_cols:
  sum_row[c] = int(calc_df[c].sum())

diff_row = {"专业/基础名称": "与目标之差"}
for p in PERSONS:
  diff_row[f"{p}_目标"] = ""
  diff_row[f"{p}_实际"] = sum_row[f"{p}_实际"] - sum_row[f"{p}_目标"]
diff_row["其他人员_实际"] = ""
diff_row["目标人数"] = ""
diff_row["实际完成"] = ""
diff_row["与目标之差"] = sum_row["与目标之差"]

total_target_cum = sum_row["目标人数"]
total_actual_cum = sum_row["实际完成"]
cum_rate_val = (
    (total_actual_cum / total_target_cum * 100) if total_target_cum > 0 else 0
)

rate_row = {"专业/基础名称": "2026年8月目标人数完成比例"}
for c in num_cols:
  rate_row[c] = ""
rate_row["实际完成"] = f"{cum_rate_val:.2f}%"


# 4. 🔥 强效渲染 HTML 表格 (解决 Edge 颜色不显示)
def render_exact_color_html(df, sum_r, diff_r, rate_r):
  html = """
    <style>
        .custom-table-container {
            overflow-x: auto;
            border: 1px solid #A6A6A6 !important;
            margin-top: 10px;
            background-color: #FFFFFF !important;
        }
        .custom-table {
            width: 100% !important;
            border-collapse: collapse !important;
            font-family: SimSun, "Times New Roman", serif !important;
            font-size: 13px !important;
            text-align: center !important;
            background-color: #FFFFFF !important;
        }
        .custom-table th, .custom-table td {
            border: 1px solid #A6A6A6 !important;
            padding: 6px 4px !important;
            font-weight: normal !important;
            color: #000000 !important;
        }
        /* 1. 大标题：淡绿 */
        .title-cell {
            background-color: #D9EAD3 !important;
            font-size: 15px !important;
        }
        /* 2. 表头：灰色 */
        .header-cell {
            background-color: #EFEFEF !important;
        }
        /* 3. 人员标题：浅蓝灰 */
        .person-cell {
            background-color: #D0E0E3 !important;
        }
        /* 4. 底部合计/汇总：淡黄底色 */
        .total-cell {
            background-color: #FFF2CC !important;
        }
        /* 5. 数字统一 Times New Roman */
        .num-font {
            font-family: "Times New Roman", serif !important;
        }
    </style>
    <div class="custom-table-container">
    <table class="custom-table">
        <!-- 标题行 -->
        <tr>
            <td colspan="17" class="title-cell">2026年9月招生数据动态表(2026年9月1日-9月7日)</td>
        </tr>
        <!-- 表头第 1 层 -->
        <tr>
            <th rowspan="2" class="header-cell">序号</th>
            <th rowspan="2" class="header-cell">专业/基础名称</th>
            <th rowspan="2" class="header-cell">目标人数</th>
            <th colspan="2" class="person-cell">人员A</th>
            <th colspan="2" class="person-cell">人员B</th>
            <th colspan="2" class="person-cell">人员C</th>
            <th colspan="2" class="person-cell">人员D</th>
            <th colspan="2" class="person-cell">人员E</th>
            <th rowspan="2" class="header-cell">其他人员</th>
            <th rowspan="2" class="header-cell">目标人数</th>
            <th rowspan="2" class="header-cell">实际完成</th>
            <th rowspan="2" class="header-cell">与目标之差</th>
        </tr>
        <!-- 表头第 2 层 -->
        <tr>
            <th class="header-cell">目标人数</th><th class="header-cell">实际完成</th>
            <th class="header-cell">目标人数</th><th class="header-cell">实际完成</th>
            <th class="header-cell">目标人数</th><th class="header-cell">实际完成</th>
            <th class="header-cell">目标人数</th><th class="header-cell">实际完成</th>
            <th class="header-cell">目标人数</th><th class="header-cell">实际完成</th>
        </tr>
    """

  # 填充主数据
  for idx, row in df.iterrows():
    html += f"""
        <tr>
            <td class="num-font">{row['序号']}</td>
            <td>{row['专业/基础名称']}</td>
            <td class="num-font">{row['目标人数']}</td>
            <td class="num-font">{row['人员A_目标']}</td><td class="num-font">{row['人员A_实际']}</td>
            <td class="num-font">{row['人员B_目标']}</td><td class="num-font">{row['人员B_实际']}</td>
            <td class="num-font">{row['人员C_目标']}</td><td class="num-font">{row['人员C_实际']}</td>
            <td class="num-font">{row['人员D_目标']}</td><td class="num-font">{row['人员D_实际']}</td>
            <td class="num-font">{row['人员E_目标']}</td><td class="num-font">{row['人员E_实际']}</td>
            <td class="num-font">{row['其他人员_实际']}</td>
            <td class="num-font">{row['目标人数']}</td>
            <td class="num-font">{row['实际完成']}</td>
            <td class="num-font">{row['与目标之差']}</td>
        </tr>
        """

  # 填充底部合计与汇总
  html += f"""
        <tr>
            <td class="total-cell"></td><td class="total-cell">{sum_r['专业/基础名称']}</td><td class="total-cell num-font">{sum_r['目标人数']}</td>
            <td class="total-cell num-font">{sum_r['人员A_目标']}</td><td class="total-cell num-font">{sum_r['人员A_实际']}</td>
            <td class="total-cell num-font">{sum_r['人员B_目标']}</td><td class="total-cell num-font">{sum_r['人员B_实际']}</td>
            <td class="total-cell num-font">{sum_r['人员C_目标']}</td><td class="total-cell num-font">{sum_r['人员C_实际']}</td>
            <td class="total-cell num-font">{sum_r['人员D_目标']}</td><td class="total-cell num-font">{sum_r['人员D_实际']}</td>
            <td class="total-cell num-font">{sum_r['人员E_目标']}</td><td class="total-cell num-font">{sum_r['人员E_实际']}</td>
            <td class="total-cell num-font">{sum_r['其他人员_实际']}</td>
            <td class="total-cell num-font">{sum_r['目标人数']}</td>
            <td class="total-cell num-font">{sum_r['实际完成']}</td>
            <td class="total-cell num-font">{sum_r['与目标之差']}</td>
        </tr>
        <tr>
            <td class="total-cell"></td><td class="total-cell">{diff_r['专业/基础名称']}</td><td class="total-cell"></td>
            <td class="total-cell"></td><td class="total-cell num-font">{diff_r['人员A_实际']}</td>
            <td class="total-cell"></td><td class="total-cell num-font">{diff_r['人员B_实际']}</td>
            <td class="total-cell"></td><td class="total-cell num-font">{diff_r['人员C_实际']}</td>
            <td class="total-cell"></td><td class="total-cell num-font">{diff_r['人员D_实际']}</td>
            <td class="total-cell"></td><td class="total-cell num-font">{diff_r['人员E_实际']}</td>
            <td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td>
            <td class="total-cell num-font">{diff_r['与目标之差']}</td>
        </tr>
        <tr>
            <td class="total-cell"></td><td class="total-cell">{rate_r['专业/基础名称']}</td><td class="total-cell"></td>
            <td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td><td class="total-cell"></td>
            <td class="total-cell"></td><td class="total-cell"></td>
            <td class="total-cell num-font">{rate_r['实际完成']}</td>
            <td class="total-cell"></td>
        </tr>
    </table>
    </div>
    """
  return html


# 渲染网页表格
st.markdown("### 📝 2026年9月招生数据动态表")
st.markdown(
    render_exact_color_html(calc_df, sum_row, diff_row, rate_row),
    unsafe_allow_html=True,
)


# 5. 导出高保真带颜色的 Excel 文件
def export_color_excel(calc_df, sum_row, diff_row, rate_row, persons):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "招生数据动态表"
  ws.views.sheetView[0].showGridLines = True

  TITLE_FILL = PatternFill(
      start_color="D9EAD3", end_color="D9EAD3", fill_type="solid"
  )
  HEADER_BG = PatternFill(
      start_color="EFEFEF", end_color="EFEFEF", fill_type="solid"
  )
  PERSON_BG = PatternFill(
      start_color="D0E0E3", end_color="D0E0E3", fill_type="solid"
  )
  TOTAL_ROW_BG = PatternFill(
      start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
  )

  FONT_TITLE = Font(name="SimSun", size=14)
  FONT_HEADER = Font(name="SimSun", size=10)
  FONT_BODY_NUM = Font(name="Times New Roman", size=10)
  FONT_BODY_ZH = Font(name="SimSun", size=10)

  BORDER_THIN = Border(
      left=Side(style="thin", color="A6A6A6"),
      right=Side(style="thin", color="A6A6A6"),
      top=Side(style="thin", color="A6A6A6"),
      bottom=Side(style="thin", color="A6A6A6"),
  )
  ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

  ws.merge_cells("A1:Q1")
  t_cell = ws["A1"]
  t_cell.value = "2026年9月招生数据动态表(2026年9月1日-9月7日)"
  t_cell.font = FONT_TITLE
  t_cell.fill = TITLE_FILL
  t_cell.alignment = ALIGN_CENTER

  ws.merge_cells("A2:A3")
  ws["A2"] = "序号"
  ws.merge_cells("B2:B3")
  ws["B2"] = "专业/基础名称"
  ws.merge_cells("C2:C3")
  ws["C2"] = "目标人数"

  col_idx = 4
  for p in persons:
    ws.merge_cells(
        start_row=2, start_column=col_idx, end_row=2, end_column=col_idx + 1
    )
    p_cell = ws.cell(row=2, column=col_idx, value=p)
    p_cell.fill = PERSON_BG
    ws.cell(row=3, column=col_idx, value="目标人数")
    ws.cell(row=3, column=col_idx + 1, value="实际完成")
    col_idx += 2

  ws.merge_cells("N2:N3")
  ws["N2"] = "其他人员"
  ws.merge_cells("O2:O3")
  ws["O2"] = "目标人数"
  ws.merge_cells("P2:P3")
  ws["P2"] = "实际完成"
  ws.merge_cells("Q2:Q3")
  ws["Q2"] = "与目标之差"

  for r in range(2, 4):
    for c in range(1, 18):
      cell = ws.cell(row=r, column=c)
      if not cell.fill.start_color.rgb:
        cell.fill = HEADER_BG
      cell.font = FONT_HEADER
      cell.alignment = ALIGN_CENTER
      cell.border = BORDER_THIN

  curr_r = 4
  for idx, row in calc_df.iterrows():
    ws.cell(row=curr_r, column=1, value=row["序号"]).font = FONT_BODY_NUM
    ws.cell(row=curr_r, column=2, value=row["专业/基础名称"]).font = FONT_BODY_ZH
    ws.cell(row=curr_r, column=3, value=row["目标人数"]).font = FONT_BODY_NUM

    c_offset = 4
    for p in persons:
      ws.cell(
          row=curr_r, column=c_offset, value=row[f"{p}_目标"]
      ).font = FONT_BODY_NUM
      ws.cell(
          row=curr_r, column=c_offset + 1, value=row[f"{p}_实际"]
      ).font = FONT_BODY_NUM
      c_offset += 2

    ws.cell(
        row=curr_r, column=c_offset, value=row.get("其他人员_实际", 0)
    ).font = FONT_BODY_NUM
    ws.cell(row=curr_r, column=c_offset + 1, value=row["目标人数"]).font = (
        FONT_BODY_NUM
    )
    ws.cell(row=curr_r, column=c_offset + 2, value=row["实际完成"]).font = (
        FONT_BODY_NUM
    )
    ws.cell(row=curr_r, column=c_offset + 3, value=row["与目标之差"]).font = (
        FONT_BODY_NUM
    )

    for c in range(1, 18):
      cell = ws.cell(row=curr_r, column=c)
      cell.alignment = ALIGN_CENTER
      cell.border = BORDER_THIN
    curr_r += 1

  for r_data in [sum_row, diff_row, rate_row]:
    ws.cell(
        row=curr_r, column=2, value=r_data["专业/基础名称"]
    ).font = FONT_HEADER
    ws.cell(row=curr_r, column=3, value=r_data.get("目标人数", "")).font = (
        FONT_HEADER
    )

    c_offset = 4
    for p in persons:
      ws.cell(
          row=curr_r, column=c_offset, value=r_data.get(f"{p}_目标", "")
      ).font = FONT_HEADER
      ws.cell(
          row=curr_r, column=c_offset + 1, value=r_data.get(f"{p}_实际", "")
      ).font = FONT_HEADER
      c_offset += 2

    ws.cell(
        row=curr_r, column=c_offset, value=r_data.get("其他人员_实际", "")
    ).font = FONT_HEADER
    ws.cell(
        row=curr_r, column=c_offset + 1, value=r_data.get("目标人数", "")
    ).font = FONT_HEADER
    ws.cell(
        row=curr_r, column=c_offset + 2, value=r_data.get("实际完成", "")
    ).font = FONT_HEADER
    ws.cell(
        row=curr_r, column=c_offset + 3, value=r_data.get("与目标之差", "")
    ).font = FONT_HEADER

    for c in range(1, 18):
      cell = ws.cell(row=curr_r, column=c)
      cell.alignment = ALIGN_CENTER
      cell.border = BORDER_THIN
      cell.fill = TOTAL_ROW_BG
    curr_r += 1

  output = io.BytesIO()
  wb.save(output)
  return output.getvalue()


st.markdown("---")
excel_bytes = export_color_excel(
    calc_df, sum_row, diff_row, rate_row, PERSONS
)

st.download_button(
    label="📊 导出为原版颜色 Excel 表格 (.xlsx)",
    data=excel_bytes,
    file_name="2026年9月招生数据动态表_带颜色版.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
