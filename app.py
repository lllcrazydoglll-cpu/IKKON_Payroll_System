import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# ==========================================
# 模組一：打卡紀錄清洗
# ==========================================
def clean_ichef_data(file):
    cleaned_data = []
    error_log = []
    raw_data = pd.read_excel(file, header=None)
    current_employee = ""
    current_clock_in = None

    for index, row in raw_data.iterrows():
        action = str(row[0]).strip()
        time_record = str(row[1]).strip()
        system_keywords = ["上班", "下班", "無下班", "無上班", "無下班記錄", "無上班記錄", "無下班紀錄", "無上班紀錄", "結帳收銀", "admin", "nan"]

        is_employee = True
        if action in system_keywords or "總時數" in action:
            is_employee = False
            
        if is_employee and action != "":
            if current_clock_in is not None:
                error_log.append({"員工": current_employee, "異常類型": "換人前無下班紀錄", "打卡時間": current_clock_in})
            current_employee = action
            current_clock_in = None

        elif action == "上班":
            if current_clock_in is not None:
                try:
                    t1 = pd.to_datetime(current_clock_in)
                    t2 = pd.to_datetime(time_record)
                    if abs((t2 - t1).total_seconds()) / 60.0 <= 10:
                        pass 
                    else:
                        error_log.append({"員工": current_employee, "異常類型": "連續上班打卡", "打卡時間": current_clock_in})
                        current_clock_in = time_record
                except:
                    current_clock_in = time_record
            else:
                current_clock_in = time_record

        elif action == "下班":
            if current_clock_in is not None:
                cleaned_data.append({"員工": current_employee, "上班時間": current_clock_in, "下班時間": time_record})
                current_clock_in = None
            else:
                error_log.append({"員工": current_employee, "異常類型": "有下班無上班", "打卡時間": time_record})

        elif "無下班" in action:
            error_log.append({"員工": current_employee, "異常類型": "iCHEF標記無下班", "打卡時間": current_clock_in if current_clock_in else time_record})
            current_clock_in = None
            
        elif "無上班" in action:
            error_log.append({"員工": current_employee, "異常類型": "iCHEF標記無上班", "打卡時間": time_record})
            current_clock_in = None

    return pd.DataFrame(cleaned_data), pd.DataFrame(error_log)

# ==========================================
# 模組二：強固型班表攤平
# ==========================================
def parse_roster_data(file, target_sheet):
    raw_roster = pd.read_excel(file, sheet_name=target_sheet, header=None)
    roster_list = []
    
    title_row_index = -1
    name_row_index = -1
    for index, row in raw_roster.iterrows():
        row_str = str(row.values)
        if "職別" in row_str and title_row_index == -1:
            title_row_index = index
        if "姓名" in row_str and name_row_index == -1:
            name_row_index = index
            break
            
    if name_row_index == -1:
        return None, "找不到「姓名」標籤，請確認班表格式是否正確。"
        
    employee_info = {}
    invalid_names = ["nan", "姓名", "NaT", "None", ""]
    
    for col_idx, val in enumerate(raw_roster.iloc[name_row_index]):
        emp_name = str(val).strip()
        if emp_name not in invalid_names and not pd.isna(val):
            is_pt = False
            if title_row_index != -1:
                title_val = str(raw_roster.iloc[title_row_index, col_idx]).strip()
                if "PT" in title_val.upper() or "兼職" in title_val:
                    is_pt = True
            employee_info[col_idx] = {"name": emp_name, "is_pt": is_pt}
            
    for index in range(name_row_index + 1, len(raw_roster)):
        row = raw_roster.iloc[index]
        date_str = str(row[0]).strip()
        
        if date_str.startswith("202"):
            for col_idx, info in employee_info.items():
                emp_name = info["name"]
                is_pt = info["is_pt"]
                shift_val = str(row[col_idx]).strip()
                
                is_working = False
                shift_string = ""
                
                if shift_val in ["nan", "NaT", "None", ""]:
                    if is_pt:
                        is_working = False
                    else:
                        is_working = True
                        shift_string = "正常班"
                elif any(x in shift_val for x in ["休", "假", "曠"]):
                    is_working = False
                else:
                    is_working = True
                    shift_string = shift_val if "-" in shift_val else "正常班"
                    
                if is_working:
                    roster_list.append({
                        "日期": date_str[:10],
                        "員工": emp_name,
                        "身份": "PT" if is_pt else "正職",
                        "班別字串": shift_string
                    })
                    
    return pd.DataFrame(roster_list), ""

# ==========================================
# 模組三：幹部打卡異常表 (最高權限覆寫解析)
# ==========================================
def parse_anomaly_data(file, target_year):
    if file is None:
        return pd.DataFrame()
        
    try:
        if file.name.endswith('.csv'):
            raw_anomaly = pd.read_csv(file, header=None)
        else:
            raw_anomaly = pd.read_excel(file, header=None)
    except:
        return pd.DataFrame()
        
    anomalies = []
    header_idx = -1
    for i, row in raw_anomaly.iterrows():
        if "姓名" in str(row.values) and "原因" in str(row.values):
            header_idx = i
            break
            
    if header_idx == -1:
        return pd.DataFrame()
        
    for i in range(header_idx + 1, len(raw_anomaly)):
        row = raw_anomaly.iloc[i]
        name = str(row[0]).strip()
        if name in ["nan", "None", "", "姓名"]:
            continue
            
        status = str(row[1]).strip()
        reason = str(row[2]).strip()
        missing_time_str = str(row[3]).strip()
        
        month = None
        day = None
        for col_idx in range(4, len(row)):
            val = str(row[col_idx]).strip()
            if val == "月" and col_idx > 0:
                month = str(row[col_idx-1]).strip()
            if val == "日" and col_idx > 0:
                day = str(row[col_idx-1]).strip()
                
        if not month or not day:
            continue
            
        overtime_val = 0.0
        for col_idx in range(len(row)-1, 4, -1):
            val = str(row[col_idx]).strip()
            if val == "小時" and col_idx > 0:
                ot_str = str(row[col_idx-1]).strip()
                if ":" in ot_str:
                    parts = ot_str.split(":")
                    if len(parts) >= 2:
                        overtime_val = float(parts[0]) + float(parts[1])/60.0
                else:
                    try:
                        overtime_val = float(ot_str)
                    except:
                        pass
                break
                
        try:
            date_str = f"{target_year}-{int(float(month)):02d}-{int(float(day)):02d}"
        except:
            continue
            
        parsed_missing_time = None
        if missing_time_str not in ["nan", "None", ""]:
            parsed_missing_time = str(missing_time_str).replace("時", "").replace("分", "").strip()
            
        anomalies.append({
            "日期": date_str,
            "員工": name,
            "狀態": status,
            "原因": reason,
            "補登時間": parsed_missing_time,
            "核准加班": overtime_val
        })
        
    return pd.DataFrame(anomalies)

# ==========================================
# 核心引擎：薪資工時碰撞與覆寫結算
# ==========================================
def calculate_payroll_hours(df_roster, df_actual, df_anomaly):
    results = []
    audit_logs = []
    
    df_actual['上班時間'] = pd.to_datetime(df_actual['上班時間'])
    df_actual['下班時間'] = pd.to_datetime(df_actual['下班時間'])
    df_actual['日期'] = df_actual['上班時間'].dt.strftime('%Y-%m-%d')
    
    for _, scheduled in df_roster.iterrows():
        date = scheduled['日期']
        emp = scheduled['員工']
        emp_type = scheduled['身份']
        shift_str = scheduled['班別字串']
        
        emp_punches = df_actual[(df_actual['員工'] == emp) & (df_actual['日期'] == date)]
        
        # 異常覆寫機制介入
        has_override = False
        manual_add_ot = 0.0
        override_reason = []
        missing_punch_dts = []
        
        if not df_anomaly.empty:
            emp_anomalies = df_anomaly[(df_anomaly['日期'] == date) & (df_anomaly['員工'] == emp)]
            for _, anom in emp_anomalies.iterrows():
                has_override = True
                override_reason.append(str(anom['原因']))
                if pd.notna(anom['核准加班']) and anom['核准加班'] != 0.0:
                    manual_add_ot += float(anom['核准加班'])
                if pd.notna(anom['補登時間']) and anom['補登時間']:
                    try:
                        time_str = str(anom['補登時間']).strip()
                        if len(time_str) == 5:
                            time_str += ":00"
                        if len(time_str.split(':')) >= 2:
                            dt = pd.to_datetime(f"{date} {time_str}")
                            missing_punch_dts.append(dt)
                    except:
                        pass

        # 彙整所有實際與補登的打卡時間
        all_times = []
        if not emp_punches.empty:
            all_times.extend(emp_punches['上班時間'].dropna().tolist())
            all_times.extend(emp_punches['下班時間'].dropna().tolist())
        all_times.extend(missing_punch_dts)
        
        if not all_times:
            results.append({
                "日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, 
                "遲到(分)": 0, "早退(分)": 0, "加班(時)": 0, "總工時(時)": 0, "狀態": "無打卡紀錄(休假或未核)"
            })
            continue
            
        actual_in = min(all_times)
        actual_out = max(all_times)
        
        # --- PT 計算邏輯 ---
        if emp_type == "PT":
            total_minutes = (actual_out - actual_in).total_seconds() / 60.0
            pt_hours = (total_minutes // 30) * 0.5
            pt_hours += manual_add_ot
            
            final_status = "已套用異常覆寫" if has_override else "PT時數結算"
            results.append({
                "日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, 
                "遲到(分)": 0, "早退(分)": 0, "加班(時)": manual_add_ot, "總工時(時)": pt_hours, "狀態": final_status
            })
            
            if has_override:
                audit_logs.append({
                    "日期": date, "員工": emp, "原始判定": "PT工時結算", 
                    "覆寫內容": f"核准加減班: {manual_add_ot}小時 / 補登時間: {len(missing_punch_dts)}筆",
                    "幹部備註原因": " | ".join(override_reason)
                })
            continue
            
        # --- 正職計算邏輯 ---
        if shift_str == "正常班":
            if actual_in.hour < 13:
                sched_in = pd.to_datetime(f"{date} 11:00:00")
                sched_out = pd.to_datetime(f"{date} 23:00:00")
                is_full_day = True
            else:
                sched_in = pd.to_datetime(f"{date} 15:00:00")
                sched_out = pd.to_datetime(f"{date} 23:00:00")
                is_full_day = False
        else:
            try:
                start_str, end_str = shift_str.split('-')
                start_str = f"{start_str[:2]}:{start_str[2:]}"
                end_str = f"{end_str[:2]}:{end_str[2:]}"
                sched_in = pd.to_datetime(f"{date} {start_str}")
                sched_out = pd.to_datetime(f"{date} {end_str}")
                if sched_out < sched_in:
                    sched_out += timedelta(days=1)
                is_full_day = (sched_out - sched_in).total_seconds() >= 36000 
            except:
                sched_in = actual_in 
                sched_out = actual_out
                is_full_day = False
                
        late_mins = 0
        if actual_in > sched_in:
            late_mins = int((actual_in - sched_in).total_seconds() / 60)
            
        early_leave_mins = 0
        welfare_virtual_hours = 0
        if actual_out < sched_out:
            diff_mins = int((sched_out - actual_out).total_seconds() / 60)
            if diff_mins <= 30:
                early_leave_mins = 0
                welfare_virtual_hours = diff_mins / 60.0
            else:
                early_leave_mins = diff_mins
                
        # 工時運算 (考量補登狀況)
        if missing_punch_dts:
            total_actual_hours = (actual_out - actual_in).total_seconds() / 3600.0
        else:
            total_actual_hours = 0
            for _, punch in emp_punches.iterrows():
                total_actual_hours += (punch['下班時間'] - punch['上班時間']).total_seconds() / 3600.0
                
        final_calculated_hours = total_actual_hours + welfare_virtual_hours
        
        overtime_hours = 0
        if is_full_day:
            overflow = final_calculated_hours - 8.5
        else:
            sched_total = (sched_out - sched_in).total_seconds() / 3600.0
            overflow = final_calculated_hours - sched_total
            
        if overflow > 0:
            overtime_hours = (overflow // 0.5) * 0.5
            
        overtime_hours += manual_add_ot
        final_status = "已套用異常覆寫" if has_override else "正常結算"
            
        results.append({
            "日期": date, "員工": emp, "身份": "正職", "班別": shift_str, 
            "遲到(分)": late_mins, "早退(分)": early_leave_mins, "加班(時)": overtime_hours, "總工時(時)": final_calculated_hours, "狀態": final_status
        })
        
        if has_override:
            audit_logs.append({
                "日期": date, "員工": emp, "原始判定": "異常/正常結算", 
                "覆寫內容": f"核准加減班: {manual_add_ot}小時 / 補登時間: {len(missing_punch_dts)}筆",
                "幹部備註原因": " | ".join(override_reason)
            })

    return pd.DataFrame(results), pd.DataFrame(audit_logs)

# ==========================================
# 介面渲染
# ==========================================
st.set_page_config(page_title="IKKON 薪資結算系統", layout="wide")
st.title("IKKON 薪資自動化結算系統")

st.markdown("### 資料匯入區")
col1, col2, col3 = st.columns(3)
with col1:
    ichef_file = st.file_uploader("1. 上傳 iCHEF 打卡紀錄", type=["xlsx"], key="ichef")
with col2:
    roster_file = st.file_uploader("2. 上傳 店鋪當月班表", type=["xlsx"], key="roster")
    selected_sheet = None
    if roster_file:
        try:
            xls = pd.ExcelFile(roster_file)
            sheet_names = xls.sheet_names
            selected_sheet = st.selectbox("請選擇要結算的班表月份 (工作表)：", sheet_names)
        except Exception as e:
            st.error("讀取班表分頁失敗，請確認檔案格式。")
            
with col3:
    anomaly_file = st.file_uploader("3. 上傳 幹部打卡異常表 (選填)", type=["csv", "xlsx"], key="anomaly")

if ichef_file and roster_file and selected_sheet:
    if st.button("執行結算與稽核比對"):
        with st.spinner('系統運算與權限覆寫中...'):
            df_cleaned, df_error = clean_ichef_data(ichef_file)
            df_roster, error_msg = parse_roster_data(roster_file, selected_sheet)
            
            if error_msg:
                st.error(error_msg)
            else:
                target_year = "2026"
                if not df_roster.empty:
                    target_year = str(df_roster.iloc[0]['日期'])[:4]
                    
                df_anomaly = parse_anomaly_data(anomaly_file, target_year)
                
                df_final_calc, df_audit = calculate_payroll_hours(df_roster, df_cleaned, df_anomaly)
                st.success(f"已成功鎖定工作表：{selected_sheet}，運算與覆寫程序完成。")
                
                tab_main, tab_audit, tab_error, tab_roster = st.tabs([
                    "最終出缺勤結算", "異常表覆寫稽核", "原始打卡異常攔截", "系統攤平班表(除錯)"
                ])
                
                with tab_main:
                    st.dataframe(df_final_calc)
                    
                with tab_audit:
                    if not df_audit.empty:
                        st.dataframe(df_audit)
                    else:
                        st.info("本次結算並未套用任何異常表覆寫紀錄。")
                        
                with tab_error:
                    if not df_error.empty:
                        st.dataframe(df_error)
                    else:
                        st.write("無任何底層異常紀錄。")
                        
                with tab_roster:
                    st.dataframe(df_roster)
