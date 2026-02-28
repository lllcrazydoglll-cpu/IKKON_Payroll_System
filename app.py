import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# ==========================================
# 模組一：打卡紀錄清洗 (保留孤兒紀錄防禦升級)
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
        system_keywords = ["上班", "下班", "無下班", "無上班", "無下班記錄", "無上班記錄", "無下班紀錄", "無上班紀錄", "結帳收銀", "admin", "nan", "總時數：0:00:00"]

        is_employee = True
        if action in system_keywords or "總時數" in action:
            is_employee = False
            
        if is_employee and action != "":
            if current_clock_in is not None:
                error_log.append({"員工": current_employee, "異常類型": "換人前無下班紀錄", "打卡時間": current_clock_in})
                # 防禦升級：保留只有上班的孤兒紀錄，供模組三縫合
                cleaned_data.append({"員工": current_employee, "上班時間": current_clock_in, "下班時間": pd.NaT})
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
                        cleaned_data.append({"員工": current_employee, "上班時間": current_clock_in, "下班時間": pd.NaT})
                        current_clock_in = time_record
                except:
                    cleaned_data.append({"員工": current_employee, "上班時間": current_clock_in, "下班時間": pd.NaT})
                    current_clock_in = time_record
            else:
                current_clock_in = time_record

        elif action == "下班":
            if current_clock_in is not None:
                cleaned_data.append({"員工": current_employee, "上班時間": current_clock_in, "下班時間": time_record})
                current_clock_in = None
            else:
                error_log.append({"員工": current_employee, "異常類型": "有下班無上班", "打卡時間": time_record})
                # 防禦升級：保留只有下班的孤兒紀錄，供模組三縫合
                cleaned_data.append({"員工": current_employee, "上班時間": pd.NaT, "下班時間": time_record})

        elif "無下班" in action:
            error_log.append({"員工": current_employee, "異常類型": "系統標記無下班", "打卡時間": current_clock_in if current_clock_in else time_record})
            if current_clock_in is not None:
                cleaned_data.append({"員工": current_employee, "上班時間": current_clock_in, "下班時間": pd.NaT})
            current_clock_in = None
            
        elif "無上班" in action:
            error_log.append({"員工": current_employee, "異常類型": "系統標記無上班", "打卡時間": time_record})
            cleaned_data.append({"員工": current_employee, "上班時間": pd.NaT, "下班時間": time_record})
            current_clock_in = None

    if current_clock_in is not None:
        error_log.append({"員工": current_employee, "異常類型": "最後一筆無下班", "打卡時間": current_clock_in})
        cleaned_data.append({"員工": current_employee, "上班時間": current_clock_in, "下班時間": pd.NaT})

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
                    shift_string = "休"
                else:
                    is_working = True
                    shift_string = shift_val if "-" in shift_val else "正常班"
                    
                roster_list.append({
                    "日期": date_str[:10],
                    "員工": emp_name,
                    "身份": "PT" if is_pt else "正職",
                    "班別字串": shift_string,
                    "表定上班狀態": is_working
                })
                    
    return pd.DataFrame(roster_list), ""

# ==========================================
# 模組三：標準化異常表解析引擎
# ==========================================
def parse_standard_anomaly_data(file):
    if file is None:
        return pd.DataFrame()
        
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=None)
        else:
            df = pd.read_excel(file, header=None)
            
        anomalies = []
        for index, row in df.iterrows():
            date_val = str(row.iloc[0]).strip()
            if "202" in date_val and len(row) >= 5:
                try:
                    dt = pd.to_datetime(date_val)
                    date_str = dt.strftime('%Y-%m-%d')
                except:
                    continue
                    
                name = str(row.iloc[1]).strip()
                command = str(row.iloc[2]).strip()
                time_val = str(row.iloc[3]).strip()
                hours_val = str(row.iloc[4]).strip()
                reason = str(row.iloc[5]).strip() if len(row) > 5 else ""
                
                try:
                    hours_float = float(hours_val)
                except:
                    hours_float = 0.0
                    
                anomalies.append({
                    "日期": date_str,
                    "員工": name,
                    "指令": command,
                    "時間": time_val if time_val not in ["nan", "None", ""] else None,
                    "時數": hours_float,
                    "原因": reason
                })
        return pd.DataFrame(anomalies)
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 核心引擎：工時碰撞、指令動態覆寫結算
# ==========================================
def calculate_payroll_hours(df_roster, df_actual, df_anomaly):
    results = []
    audit_logs = []
    
    # 確保即使只有單邊打卡的紀錄，也能被精準轉換日期
    df_actual['上班時間'] = pd.to_datetime(df_actual['上班時間'])
    df_actual['下班時間'] = pd.to_datetime(df_actual['下班時間'])
    df_actual['temp_time'] = df_actual['上班時間'].fillna(df_actual['下班時間'])
    df_actual['日期'] = df_actual['temp_time'].dt.strftime('%Y-%m-%d')
    
    for _, scheduled in df_roster.iterrows():
        date = scheduled['日期']
        emp = scheduled['員工']
        emp_type = scheduled['身份']
        original_shift_str = scheduled['班別字串']
        is_working = scheduled['表定上班狀態']
        
        emp_punches = df_actual[(df_actual['員工'] == emp) & (df_actual['日期'] == date)]
        
        shift_str = original_shift_str
        manual_add_ot = 0.0
        missing_punch_dts = []
        override_reasons = []
        has_override = False
        
        # 讀取異常表覆寫指令
        if not df_anomaly.empty:
            emp_anomalies = df_anomaly[(df_anomaly['日期'] == date) & (df_anomaly['員工'] == emp)]
            for _, anom in emp_anomalies.iterrows():
                cmd = anom['指令']
                reason = str(anom['原因'])
                
                if cmd == "變更為排休":
                    shift_str = "休"
                    is_working = False
                    has_override = True
                    override_reasons.append(f"調休變更: {reason}")
                    
                elif cmd == "變更為應勤":
                    shift_str = "正常班"
                    is_working = True
                    has_override = True
                    override_reasons.append(f"調休變更: {reason}")
                    
                elif cmd in ["補登上班", "補登下班", "上班補登", "下班補登"]:
                    if pd.notna(anom['時間']):
                        ts = str(anom['時間']).strip()
                        if len(ts) == 5:
                            ts += ":00"
                        try:
                            dt = pd.to_datetime(f"{date} {ts}")
                            missing_punch_dts.append(dt)
                            has_override = True
                            override_reasons.append(f"{cmd} {ts}: {reason}")
                        except:
                            pass
                            
                elif cmd == "時數增減":
                    if anom['時數'] != 0.0:
                        manual_add_ot += anom['時數']
                        has_override = True
                        override_reasons.append(f"時數增減 {anom['時數']}H: {reason}")

        # 彙整打卡時間 (包含孤兒紀錄與異常表補登時間)
        all_times = []
        for _, punch in emp_punches.iterrows():
            if pd.notna(punch['上班時間']):
                all_times.append(punch['上班時間'])
            if pd.notna(punch['下班時間']):
                all_times.append(punch['下班時間'])
        all_times.extend(missing_punch_dts)
        all_times.sort()
        
        # 防禦升級：處理完全無打卡的覆寫 (例如 -8 小時事假)
        if not is_working and not all_times:
            if has_override and manual_add_ot != 0:
                results.append({
                    "日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, 
                    "遲到(分)": 0, "早退(分)": 0, "加班(時)": manual_add_ot, "總工時(時)": 0, "狀態": "已套用異常覆寫"
                })
                audit_logs.append({
                    "日期": date, "員工": emp, "原始判定": "排休無打卡", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)
                })
            continue
            
        if is_working and not all_times:
            final_status = "已套用異常覆寫" if has_override else "無打卡紀錄(曠職或未核)"
            results.append({
                "日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, 
                "遲到(分)": 0, "早退(分)": 0, "加班(時)": manual_add_ot, "總工時(時)": 0, "狀態": final_status
            })
            if has_override:
                audit_logs.append({
                    "日期": date, "員工": emp, "原始判定": "曠職或未核", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)
                })
            continue

        actual_in = all_times[0]
        actual_out = all_times[-1]
        span_hours = (actual_out - actual_in).total_seconds() / 3600.0

        # 休假日支援判定引擎
        if not is_working and all_times:
            total_actual_hours = 0
            if len(all_times) % 2 == 0:
                for i in range(0, len(all_times), 2):
                    total_actual_hours += (all_times[i+1] - all_times[i]).total_seconds() / 3600.0
            else:
                total_actual_hours = span_hours
                
            if emp_type == "PT":
                support_ot = ((total_actual_hours * 60.0) // 30) * 0.5
            else:
                support_ot = (total_actual_hours // 0.5) * 0.5
                
            support_ot += manual_add_ot
            results.append({
                "日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, 
                "遲到(分)": 0, "早退(分)": 0, "加班(時)": support_ot, "總工時(時)": round(total_actual_hours, 2), "狀態": "休假支援(全額加班)"
            })
            if has_override:
                audit_logs.append({
                    "日期": date, "員工": emp, "原始判定": "休假支援", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)
                })
            continue

        # PT 正常計算邏輯
        if emp_type == "PT":
            total_actual_hours = 0
            if len(all_times) % 2 == 0:
                for i in range(0, len(all_times), 2):
                    total_actual_hours += (all_times[i+1] - all_times[i]).total_seconds() / 3600.0
            else:
                total_actual_hours = span_hours
                
            pt_hours = ((total_actual_hours * 60.0) // 30) * 0.5
            pt_hours += manual_add_ot
            
            final_status = "已套用異常覆寫" if has_override else "PT時數結算"
            results.append({
                "日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, 
                "遲到(分)": 0, "早退(分)": 0, "加班(時)": manual_add_ot, "總工時(時)": pt_hours, "狀態": final_status
            })
            if has_override:
                audit_logs.append({
                    "日期": date, "員工": emp, "原始判定": "PT工時結算", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)
                })
            continue
            
        late_mins = 0
        early_leave_mins = 0
        total_calculated_hours = 0
        
        # 正職正常計算邏輯
        if shift_str == "正常班":
            if actual_in.hour < 13 or len(all_times) >= 4:
                sched1_in = pd.to_datetime(f"{date} 11:00:00")
                sched1_out = pd.to_datetime(f"{date} 14:30:00")
                sched2_in = pd.to_datetime(f"{date} 17:00:00")
                sched2_out = pd.to_datetime(f"{date} 23:00:00")
                
                if all_times[0] > sched1_in:
                    late_mins += int((all_times[0] - sched1_in).total_seconds() / 60)
                if len(all_times) >= 4 and all_times[2] > sched2_in:
                    late_mins += int((all_times[2] - sched2_in).total_seconds() / 60)
                
                s1_in = max(all_times[0], sched1_in)
                s1_out = min(all_times[1] if len(all_times) >= 2 else all_times[0], sched1_out)
                h1 = max(0, (s1_out - s1_in).total_seconds() / 3600.0)
                
                if len(all_times) >= 4:
                    s2_in = max(all_times[2], sched2_in)
                    s2_act_out = all_times[3]
                elif len(all_times) == 2:
                    s2_in = sched2_in
                    s2_act_out = all_times[1]
                else:
                    s2_in = sched2_in
                    s2_act_out = all_times[-1]
                
                if s2_act_out < sched2_out and (sched2_out - s2_act_out).total_seconds() <= 1800:
                    s2_out = sched2_out
                else:
                    s2_out = min(s2_act_out, sched2_out)
                    diff = int((sched2_out - s2_act_out).total_seconds() / 60)
                    if diff > 30:
                        early_leave_mins = diff
                        
                h2 = max(0, (s2_out - s2_in).total_seconds() / 3600.0)
                total_calculated_hours = h1 + h2
                base_hours = 8.5
            else:
                sched_in = pd.to_datetime(f"{date} 15:00:00")
                sched_out = pd.to_datetime(f"{date} 23:00:00")
                
                if all_times[0] > sched_in:
                    late_mins += int((all_times[0] - sched_in).total_seconds() / 60)
                    
                if actual_out < sched_out:
                    diff = int((sched_out - actual_out).total_seconds() / 60)
                    if diff > 30:
                        early_leave_mins = diff
                        valid_out = actual_out
                    else:
                        valid_out = sched_out
                else:
                    valid_out = min(actual_out, sched_out)
                    
                valid_in = max(all_times[0], sched_in)
                total_calculated_hours = max(0, (valid_out - valid_in).total_seconds() / 3600.0)
                base_hours = 8.0
                
        else:
            try:
                start_str, end_str = shift_str.split('-')
                start_str = f"{start_str[:2]}:{start_str[2:]}"
                end_str = f"{end_str[:2]}:{end_str[2:]}"
                sched_in = pd.to_datetime(f"{date} {start_str}")
                sched_out = pd.to_datetime(f"{date} {end_str}")
                if sched_out < sched_in:
                    sched_out += timedelta(days=1)
                    
                if all_times[0] > sched_in:
                    late_mins += int((all_times[0] - sched_in).total_seconds() / 60)
                    
                if actual_out < sched_out:
                    diff = int((sched_out - actual_out).total_seconds() / 60)
                    if diff > 30:
                        early_leave_mins = diff
                        valid_out = actual_out
                    else:
                        valid_out = sched_out
                else:
                    valid_out = min(actual_out, sched_out)
                    
                total_calculated_hours = 0
                if len(all_times) % 2 == 0:
                    for i in range(0, len(all_times), 2):
                        in_time = all_times[i]
                        out_time = all_times[i+1]
                        if i == 0: in_time = max(in_time, sched_in)
                        if i == len(all_times)-2: out_time = valid_out
                        total_calculated_hours += max(0, (out_time - in_time).total_seconds() / 3600.0)
                else:
                    in_time = max(all_times[0], sched_in)
                    total_calculated_hours = max(0, (valid_out - in_time).total_seconds() / 3600.0)
                    
                base_hours = (sched_out - sched_in).total_seconds() / 3600.0
            except:
                base_hours = 8.5
                
        overflow = total_calculated_hours - base_hours
        if overflow > 0:
            overtime_hours = (overflow // 0.5) * 0.5
        else:
            overtime_hours = 0
            
        overtime_hours += manual_add_ot
        final_status = "已套用異常覆寫" if has_override else "正常結算"
            
        results.append({
            "日期": date, "員工": emp, "身份": "正職", "班別": shift_str, 
            "遲到(分)": late_mins, "早退(分)": early_leave_mins, "加班(時)": overtime_hours, "總工時(時)": round(total_calculated_hours, 2), "狀態": final_status
        })
        
        if has_override:
            audit_logs.append({
                "日期": date, "員工": emp, "原始判定": "異常/正常結算", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)
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
    anomaly_file = st.file_uploader("3. 上傳 標準化異常表 (選填)", type=["csv", "xlsx"], key="anomaly")

if ichef_file and roster_file and selected_sheet:
    if st.button("執行結算與稽核比對"):
        with st.spinner('系統運算與權限覆寫中...'):
            df_cleaned, df_error = clean_ichef_data(ichef_file)
            df_roster, error_msg = parse_roster_data(roster_file, selected_sheet)
            
            if error_msg:
                st.error(error_msg)
            else:
                df_anomaly = parse_standard_anomaly_data(anomaly_file)
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
