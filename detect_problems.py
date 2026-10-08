#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Jumbo Orient Punch Problem Detector (東淦打卡異常偵測工具)
- 本機離線運行，保障員工考勤數據私隱
- 自動過濾 E 職員與目標業務部門
- 依重要程度排序：全天缺勤 (ABS) 與漏打卡絕對置頂
- 匯出報表新增第 14 欄「異常原因 (Reason)」
"""

import os
import re
import glob
import yaml
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

CONFIG_PATH = "rules.yaml"

def locate_attendance_files():
    """搜尋目錄下最新考勤原檔，排除暫存與輸出檔"""
    files = glob.glob("*Attendance*.xls*") + glob.glob("*attendance*.xls*")
    files = [f for f in files if "problem" not in f.lower() and not os.path.basename(f).startswith("~$")]
    if not files:
        files = [f for f in glob.glob("*.xls*") if "problem" not in f.lower() and not os.path.basename(f).startswith("~$")]
    if not files:
        raise FileNotFoundError("目錄下找不到考勤原始 Excel 檔案！")
    
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return files[0]

def parse_excel_structure(file_path, cutoff_day_cfg=25):
    """自適應引擎、動態定位表頭、保護欄位長度並抽取 Period"""
    engine = "openpyxl" if file_path.endswith(".xlsx") else "xlrd"
    df_raw = pd.read_excel(file_path, header=None, engine=engine)

    # 1. 抽取 Period 檔名與截數截止日
    out_filename = "Problem.xlsx"
    cutoff_date = None
    period_pattern = re.compile(r'Period\s*:\s*:?\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})\s*To\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})', re.IGNORECASE)
    
    for r in range(min(10, len(df_raw))):
        line = " ".join(df_raw.iloc[r].dropna().astype(str).values)
        m = period_pattern.search(line)
        if m:
            d_start, m_start, y_start, d_end, m_end, y_end = m.groups()
            out_filename = f"{int(d_start)}-{int(d_end)}.{int(m_end)}.{y_end}.Problem.xlsx"
            if cutoff_day_cfg:
                cutoff_date = pd.Timestamp(year=int(y_end), month=int(m_end), day=int(cutoff_day_cfg))
            break

    # 2. 定位表頭
    header_idx = None
    for idx, row in df_raw.iterrows():
        if any("Department Code" in str(v) for v in row.values):
            header_idx = idx
            break

    if header_idx is None:
        raise ValueError("無法在原始檔案中定位 'Department Code' 表頭！")

    headers = [
        'Dept_Code', 'Dept_Name', 'Emp_Code', 'Name', 'Designation',
        'Date', 'Day', 'Clk1', 'Clk2', 'Clk3', 'Clk4', 'Time_In', 'Time_Out',
        'Shift', 'Normal', 'Late', 'Early_Out', 'Actual',
        'OT_1_0', 'OT_1_5', 'OT_2_0', 'OT_3_0', 'Flat', 'Reason'
    ]

    # 3. 欄位數安全長度防護
    data_start = header_idx + 2
    df_data = df_raw.iloc[data_start:, :].copy()
    if df_data.shape[1] < len(headers):
        for i in range(len(headers) - df_data.shape[1]):
            df_data[df_data.shape[1]] = 0.0
    
    df_data = df_data.iloc[:, :len(headers)]
    df_data.columns = headers
    return df_data, out_filename, cutoff_date

def load_rules(path=CONFIG_PATH):
    if not os.path.exists(path):
        return {"cutoff_day": 25, "target_departments": ['PCD', 'HOF', 'PMD', 'POD', 'QSD', 'WMD', 'ADD', 'CPD', 'SED']}
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def run_integrity_audit(df, config, cutoff_date=None):
    """執行業務邏輯審計，並按重要級別賦予排序權重"""
    df = df[df['Emp_Code'].notna()].copy()
    df['Emp_Code'] = df['Emp_Code'].astype(str).str.strip().str.upper()
    
    # 僅審計 E-code
    df = df[df['Emp_Code'].str.startswith('E')].copy()
    
    # 部門白名單過濾
    target_depts = config.get('target_departments', ['PCD', 'HOF', 'PMD', 'POD', 'QSD', 'WMD', 'ADD', 'CPD', 'SED'])
    if target_depts:
        df = df[df['Dept_Code'].isin(target_depts)].copy()

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    # 截數日過濾
    if cutoff_date is not None:
        df = df[df['Date'] <= cutoff_date].copy()

    num_cols = ['Clk1', 'Clk2', 'Clk3', 'Clk4', 'Time_In', 'Time_Out', 'Normal', 'Late', 'Early_Out', 'Actual']
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    for c in ['Dept_Code', 'Dept_Name', 'Name', 'Designation', 'Shift', 'Reason']:
        df[c] = df[c].fillna('').astype(str).str.strip().str.upper()

    # 排除休假/公眾假期
    df = df[~((df['Shift'] == 'REST') | (df['Reason'].isin(['RES', 'PH'])))].copy()

    # 計算最後打卡時間
    df['Last_Clk'] = df[['Clk1', 'Clk2', 'Clk3', 'Clk4']].max(axis=1)

    # 判定異常原因與重要級別 (Priority: 越小越排前)
    def evaluate_row(r):
        reason_raw = r['Reason']
        clk1 = r['Clk1']
        last_clk = r['Last_Clk']
        time_out = r['Time_Out']
        day = r['Day']
        normal = r['Normal']

        # 🚨 優先級 1: 全天缺勤 (ABS) -> 置頂
        if reason_raw == 'ABS':
            return (1, "缺勤但有打卡 (ABS with Punch)" if clk1 > 0 else "全天缺勤 (ABS)")

        # ⚠️ 優先級 2: 漏打收工卡 -> 置頂第二順位
        if clk1 > 0 and last_clk == clk1 and (time_out == 0 or last_clk < 13.00):
            return (2, "漏打收工卡 (Missed Checkout)")

        # 優先級 3: 特殊任務無實質打卡
        if clk1 == 0 and last_clk == 0 and reason_raw in ['TRAIN', 'SITE', 'MEET', 'EVENT', 'OTHER']:
            return (3, f"特殊任務無打卡 ({reason_raw})")

        # 優先級 4: 遲到 (遲過 9:00 返工)
        if clk1 > 9.00 and not reason_raw:
            return (4, f"遲過九點 (Late: {clk1:.2f} > 9:00)")

        # 優先級 5: 早退 (星期一至五，標準班次，早過 17:30 收工)
        if day != 'Saturday' and normal >= 7.0 and last_clk > 0 and last_clk < 17.30 and not reason_raw:
            return (5, f"早過五點半 (Early: {last_clk:.2f} < 17:30)")

        return (99, "")

    eval_results = df.apply(evaluate_row, axis=1)
    df['Priority'] = [res[0] for res in eval_results]
    df['Audit_Reason'] = [res[1] for res in eval_results]

    # 只保留有異常的資料
    problem_df = df[df['Priority'] < 99].copy()

    # 打卡全為 0 的記錄，Time IN/OUT 歸零顯示
    zero_clk_mask = (problem_df['Clk1'] == 0) & (problem_df['Clk2'] == 0) & (problem_df['Clk3'] == 0) & (problem_df['Clk4'] == 0)
    problem_df.loc[zero_clk_mask, 'Time_In'] = 0.0
    problem_df.loc[zero_clk_mask, 'Time_Out'] = 0.0

    # 【重要改動】：先按 Priority 排序（缺勤 ABS 最先，漏打卡第二），再按部門與日期排序
    original_dept_order = list(dict.fromkeys(df['Dept_Code'].dropna()))
    dept_cat = pd.CategoricalDtype(categories=original_dept_order, ordered=True)
    problem_df['Dept_Order'] = problem_df['Dept_Code'].astype(dept_cat)

    problem_df = problem_df.sort_values(by=['Priority', 'Dept_Order', 'Emp_Code', 'Date'])

    return problem_df

def export_aligned_excel(df_problems, output_filename):
    """匯出包含第 14 欄 Reason 的標準化 Excel 報表"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = output_filename.replace(".xlsx", ".Log")[:31]

    # 雙層表頭（第 14 欄為 Reason）
    ws.append(['Department Code', 'Department Name', 'Emp Code', 'Name', 'Designation',
               'Date', 'Day', 'Actual Clocking', None, None, None, 'Time', None, 'Reason'])
    ws.append([None, None, None, None, None, None, None, None, None, None, None, 'IN', 'OUT', None])

    ws.merge_cells('H1:K1')
    ws.merge_cells('L1:M1')

    cols_to_export = ['Dept_Code', 'Dept_Name', 'Emp_Code', 'Name', 'Designation',
                      'Date', 'Day', 'Clk1', 'Clk2', 'Clk3', 'Clk4', 'Time_In', 'Time_Out', 'Audit_Reason']
    
    rows_data = df_problems[cols_to_export].values.tolist()
    for row in rows_data:
        if pd.notna(row[5]):
            row[5] = pd.to_datetime(row[5]).to_pydatetime()
        ws.append(row)

    # 樣式定義
    header_font = Font(name="Verdana", size=10, bold=True)
    body_font = Font(name="Verdana", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    dotted_border = Border(
        left=Side(style='dotted', color='A0A0A0'),
        right=Side(style='dotted', color='A0A0A0'),
        top=Side(style='thin', color='A0A0A0'),
        bottom=Side(style='thin', color='A0A0A0')
    )

    # 重要異常 Highlight 樣式
    abs_fill = PatternFill(start_color="FFE4E6", end_color="FFE4E6", fill_type="solid")     # 淺紅底色
    missed_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")  # 淺黃底色
    danger_font = Font(name="Verdana", size=10, bold=True, color="991B1B")                 # 紅色字體
    warning_font = Font(name="Verdana", size=10, bold=True, color="92400E")                # 橙色字體

    # 設定表頭樣式
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=14):
        for cell in row:
            cell.font = header_font
            cell.alignment = center_align

    # 設定資料列樣式與 Highlight
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=14):
        reason_val = str(row[13].value or '')
        is_abs = "缺勤" in reason_val or "ABS" in reason_val
        is_missed = "漏打收工卡" in reason_val

        for col_idx, cell in enumerate(row, start=1):
            cell.font = body_font
            cell.border = dotted_border
            
            # ABS 整列上淡紅，漏打卡上淡黃
            if is_abs:
                cell.fill = abs_fill
            elif is_missed:
                cell.fill = missed_fill

            # 第 6 欄為日期
            if col_idx == 6:
                cell.number_format = 'yyyy-mm-dd hh:mm:ss'
                cell.alignment = center_align
            # 打卡數值格式 0.00
            elif isinstance(cell.value, (int, float)):
                cell.number_format = '0.00'
            # 第 14 欄 Reason 字體加粗突顯
            elif col_idx == 14:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                if is_abs:
                    cell.font = danger_font
                elif is_missed:
                    cell.font = warning_font

    # 自適應欄寬
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(c.value or '')) for c in col[:20])
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output_filename)
    print(f"[✓] 成功產出優先排序報表: {output_filename} (共篩選出 {len(df_problems)} 筆異常記錄，ABS與漏打卡已置頂)")

def main():
    try:
        config = load_rules()
        file_path = locate_attendance_files()
        print(f"[*] 分析考勤原檔: {file_path}")
        
        cutoff_day_cfg = config.get('cutoff_day', 25)
        df_data, out_name, cutoff_date = parse_excel_structure(file_path, cutoff_day_cfg)
        
        if cutoff_date:
            print(f"[*] 生效結算截止日 (Cut-off): <= {cutoff_date.strftime('%Y-%m-%d')}")
            
        problems = run_integrity_audit(df_data, config, cutoff_date)
        export_aligned_excel(problems, out_name)
    except Exception as e:
        print(f"[X] 執行中斷: {e}")

if __name__ == '__main__':
    main()
