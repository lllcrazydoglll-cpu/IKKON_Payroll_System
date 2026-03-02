import streamlit as st
import pandas as pd
import math
import io
import zipfile
from datetime import datetime, timedelta

# ==========================================
# 會計級精算：強制傳統四捨五入
# ==========================================
def custom_round(n):
    return int(math.floor(n + 0.5))

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
        system_keywords = ["上班", "下班", "無下班", "無上班", "無下班記錄", "無上班記錄", "無下班紀錄", "無上班紀錄", "結帳收銀", "admin", "nan", "總時數：0:00:00"]

        is_employee = True
        if action in system_keywords or "總時數" in action:
            is_employee = False
            
        if is_employee and action != "":
            if current_clock_in is not None:
                error_log.append({"員工": current_employee, "異常類型": "換人前無下班紀錄", "打卡時間": current_clock_in})
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
# 模組三：表頭智慧追蹤異常表解析引擎
# ==========================================
def parse_standard_anomaly_data(file, sheet_name=None):
    if file is None:
        return pd.DataFrame()
        
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=None)
        else:
            if sheet_name:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
            else:
                df = pd.read_excel(file, header=None)
            
        anomalies = []
        c_date, c_name, c_cmd, c_time, c_range, c_hrs, c_rsn = 0, 1, 2, 3, -1, 4, 5
        
        for index, row in df.iterrows():
            row_vals = [str(x).strip() for x in row.values]
            
            if any("姓名" in v for v in row_vals) and any("指令" in v for v in row_vals):
                for i, v in enumerate(row_vals):
                    if "日期" in v: c_date = i
                    elif "姓名" in v: c_name = i
                    elif "指令" in v: c_cmd = i
                    elif "精確" in v: c_time = i
                    elif v == "時數異動" or "異動脈絡" in v: c_range = i
                    elif "數值" in v or "小時" in v: c_hrs = i
                    elif "事由" in v or "備註" in v: c_rsn = i
                continue
                
            date_val = str(row_vals[c_date]) if c_date < len(row_vals) else ""
            
            if "202" in date_val:
                try:
                    dt = pd.to_datetime(date_val)
                    date_str = dt.strftime('%Y-%m-%d')
                except:
                    continue
                    
                name = row_vals[c_name] if c_name < len(row_vals) and c_name != -1 else ""
                command = row_vals[c_cmd] if c_cmd < len(row_vals) and c_cmd != -1 else ""
                exact_time = row_vals[c_time] if c_time < len(row_vals) and c_time != -1 else ""
                time_range = row_vals[c_range] if c_range < len(row_vals) and c_range != -1 else ""
                hours_val = row_vals[c_hrs] if c_hrs < len(row_vals) and c_hrs != -1 else ""
                reason = row_vals[c_rsn] if c_rsn < len(row_vals) and c_rsn != -1 else ""
                
                if exact_time in ["nan", "None", ""]: exact_time = None
                if time_range in ["nan", "None", ""]: time_range = None
                
                hours_float = 0.0
                if hours_val not in ["nan", "None", ""]:
                    try:
                        hours_float = float(hours_val)
                    except ValueError:
                        hours_float = 0.0
                        
                anomalies.append({
                    "日期": date_str,
                    "員工": name,
                    "指令": command,
                    "精確時間": exact_time,
                    "時數異動脈絡": time_range,
                    "時數": hours_float,
                    "原因": reason if reason not in ["nan", "None"] else ""
                })
        return pd.DataFrame(anomalies)
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 核心引擎：工時碰撞
# ==========================================
def calculate_payroll_hours(df_roster, df_actual, df_anomaly):
    results = []
    audit_logs = []
    
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
        
        if not df_anomaly.empty:
            emp_anomalies = df_anomaly[(df_anomaly['日期'] == date) & (df_anomaly['員工'] == emp)]
            for _, anom in emp_anomalies.iterrows():
                cmd = anom['指令']
                reason = str(anom['原因'])
                exact_time = str(anom['精確時間']).strip() if pd.notna(anom['精確時間']) else ""
                time_range = str(anom['時數異動脈絡']).strip() if pd.notna(anom['時數異動脈絡']) else ""
                
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
                    if exact_time:
                        ts = exact_time
                        if len(ts) == 5: ts += ":00"
                        try:
                            dt = pd.to_datetime(f"{date} {ts}")
                            missing_punch_dts.append(dt)
                            has_override = True
                            override_reasons.append(f"{cmd} {ts}: {reason}")
                        except: pass
                elif cmd == "時數增減":
                    if anom['時數'] != 0.0:
                        manual_add_ot += anom['時數']
                        has_override = True
                        if time_range and time_range.lower() not in ["nan", "none", ""]:
                            override_reasons.append(f"時數增減 {anom['時數']}H [{time_range}]: {reason}")
                        else:
                            override_reasons.append(f"時數增減 {anom['時數']}H: {reason}")

        all_times = []
        for _, punch in emp_punches.iterrows():
            if pd.notna(punch['上班時間']): all_times.append(punch['上班時間'])
            if pd.notna(punch['下班時間']): all_times.append(punch['下班時間'])
        all_times.extend(missing_punch_dts)
        all_times.sort()
        
        if not is_working and not all_times:
            if has_override and manual_add_ot != 0:
                results.append({"日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, "遲到(分)": 0, "早退(分)": 0, "加班(時)": manual_add_ot, "總工時(時)": 0, "狀態": "已套用異常覆寫"})
                audit_logs.append({"日期": date, "員工": emp, "原始判定": "排休無打卡", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)})
            continue
            
        if is_working and not all_times:
            final_status = "已套用異常覆寫" if has_override else "無打卡紀錄(曠職或未核)"
            results.append({"日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, "遲到(分)": 0, "早退(分)": 0, "加班(時)": manual_add_ot, "總工時(時)": 0, "狀態": final_status})
            if has_override:
                audit_logs.append({"日期": date, "員工": emp, "原始判定": "曠職或未核", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)})
            continue

        actual_in = all_times[0]
        actual_out = all_times[-1]
        span_hours = (actual_out - actual_in).total_seconds() / 3600.0

        if not is_working and all_times:
            total_actual_hours = sum([(all_times[i+1] - all_times[i]).total_seconds() / 3600.0 for i in range(0, len(all_times)-1, 2)]) if len(all_times) % 2 == 0 else span_hours
            if emp_type == "PT":
                pt_mins = round(total_actual_hours * 60.0, 2)
                support_ot = (pt_mins // 30) * 0.5
            else:
                support_ot = (total_actual_hours // 0.5) * 0.5
            support_ot += manual_add_ot
            results.append({"日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, "遲到(分)": 0, "早退(分)": 0, "加班(時)": support_ot, "總工時(時)": round(total_actual_hours, 2), "狀態": "休假支援(全額加班)"})
            if has_override: audit_logs.append({"日期": date, "員工": emp, "原始判定": "休假支援", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)})
            continue

        if emp_type == "PT":
            total_actual_hours = 0
            if shift_str == "1100-2200":
                s1_in = pd.to_datetime(f"{date} 11:00:00")
                s2_in = pd.to_datetime(f"{date} 17:00:00")
                if len(all_times) >= 4:
                    in1 = max(all_times[0], s1_in)
                    out1 = all_times[1]
                    in2 = max(all_times[2], s2_in)
                    out2 = all_times[3]
                    total_actual_hours += max(0, (out1 - in1).total_seconds() / 3600.0)
                    total_actual_hours += max(0, (out2 - in2).total_seconds() / 3600.0)
                elif len(all_times) == 2:
                    if all_times[0].hour < 14: in1 = max(all_times[0], s1_in)
                    else: in1 = max(all_times[0], s2_in)
                    out1 = all_times[1]
                    total_actual_hours += max(0, (out1 - in1).total_seconds() / 3600.0)
                else:
                    total_actual_hours = sum([(all_times[i+1] - all_times[i]).total_seconds() / 3600.0 for i in range(0, len(all_times)-1, 2)]) if len(all_times) % 2 == 0 else span_hours
            elif shift_str != "正常班" and "-" in shift_str:
                try:
                    s_str = shift_str.split('-')[0]
                    sched_in = pd.to_datetime(f"{date} {s_str[:2]}:{s_str[2:]}")
                    if len(all_times) % 2 == 0:
                        for i in range(0, len(all_times)-1, 2):
                            in_time = max(all_times[i], sched_in) if i == 0 else all_times[i]
                            out_time = all_times[i+1]
                            total_actual_hours += max(0, (out_time - in_time).total_seconds() / 3600.0)
                    else:
                        total_actual_hours = max(0, (all_times[-1] - max(all_times[0], sched_in)).total_seconds() / 3600.0)
                except:
                    total_actual_hours = sum([(all_times[i+1] - all_times[i]).total_seconds() / 3600.0 for i in range(0, len(all_times)-1, 2)]) if len(all_times) % 2 == 0 else span_hours
            else:
                total_actual_hours = sum([(all_times[i+1] - all_times[i]).total_seconds() / 3600.0 for i in range(0, len(all_times)-1, 2)]) if len(all_times) % 2 == 0 else span_hours

            pt_mins = round(total_actual_hours * 60.0, 2)
            pt_hours = (pt_mins // 30) * 0.5
            pt_hours += manual_add_ot
            
            final_status = "已套用異常覆寫" if has_override else "PT時數結算"
            results.append({"日期": date, "員工": emp, "身份": emp_type, "班別": shift_str, "遲到(分)": 0, "早退(分)": 0, "加班(時)": manual_add_ot, "總工時(時)": pt_hours, "狀態": final_status})
            if has_override: audit_logs.append({"日期": date, "員工": emp, "原始判定": "PT工時結算", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)})
            continue
            
        late_mins = 0
        early_leave_mins = 0
        total_calculated_hours = 0
        
        if shift_str == "正常班":
            if actual_in.hour < 13 or len(all_times) >= 4:
                sched1_in, sched1_out = pd.to_datetime(f"{date} 11:00:00"), pd.to_datetime(f"{date} 14:30:00")
                sched2_in, sched2_out = pd.to_datetime(f"{date} 17:00:00"), pd.to_datetime(f"{date} 23:00:00")
                
                if all_times[0] > sched1_in: late_mins += int((all_times[0] - sched1_in).total_seconds() / 60)
                if len(all_times) >= 4 and all_times[2] > sched2_in: late_mins += int((all_times[2] - sched2_in).total_seconds() / 60)
                
                s1_in = max(all_times[0], sched1_in)
                s1_out = min(all_times[1] if len(all_times) >= 2 else all_times[0], sched1_out)
                h1 = max(0, (s1_out - s1_in).total_seconds() / 3600.0)
                
                if len(all_times) >= 4: s2_in, s2_act_out = max(all_times[2], sched2_in), all_times[3]
                elif len(all_times) == 2: s2_in, s2_act_out = sched2_in, all_times[1]
                else: s2_in, s2_act_out = sched2_in, all_times[-1]
                
                if s2_act_out < sched2_out and (sched2_out - s2_act_out).total_seconds() <= 1800: s2_out = sched2_out
                else:
                    s2_out = min(s2_act_out, sched2_out)
                    diff = int((sched2_out - s2_act_out).total_seconds() / 60)
                    if diff > 30: early_leave_mins = diff
                        
                total_calculated_hours = h1 + max(0, (s2_out - s2_in).total_seconds() / 3600.0)
                base_hours = 8.5
            else:
                sched_in, sched_out = pd.to_datetime(f"{date} 15:00:00"), pd.to_datetime(f"{date} 23:00:00")
                if all_times[0] > sched_in: late_mins += int((all_times[0] - sched_in).total_seconds() / 60)
                if actual_out < sched_out:
                    diff = int((sched_out - actual_out).total_seconds() / 60)
                    if diff > 30: early_leave_mins, valid_out = diff, actual_out
                    else: valid_out = sched_out
                else: valid_out = min(actual_out, sched_out)
                valid_in = max(all_times[0], sched_in)
                total_calculated_hours = max(0, (valid_out - valid_in).total_seconds() / 3600.0)
                base_hours = 8.0
        else:
            try:
                s_str, e_str = shift_str.split('-')
                sched_in = pd.to_datetime(f"{date} {s_str[:2]}:{s_str[2:]}")
                sched_out = pd.to_datetime(f"{date} {e_str[:2]}:{e_str[2:]}")
                if sched_out < sched_in: sched_out += timedelta(days=1)
                if all_times[0] > sched_in: late_mins += int((all_times[0] - sched_in).total_seconds() / 60)
                if actual_out < sched_out:
                    diff = int((sched_out - actual_out).total_seconds() / 60)
                    if diff > 30: early_leave_mins, valid_out = diff, actual_out
                    else: valid_out = sched_out
                else: valid_out = min(actual_out, sched_out)
                
                total_calculated_hours = 0
                if len(all_times) % 2 == 0:
                    for i in range(0, len(all_times), 2):
                        i_in, i_out = all_times[i], all_times[i+1]
                        if i == 0: i_in = max(i_in, sched_in)
                        if i == len(all_times)-2: i_out = valid_out
                        total_calculated_hours += max(0, (i_out - i_in).total_seconds() / 3600.0)
                else:
                    total_calculated_hours = max(0, (valid_out - max(all_times[0], sched_in)).total_seconds() / 3600.0)
                base_hours = (sched_out - sched_in).total_seconds() / 3600.0
            except: base_hours = 8.5
                
        overflow = total_calculated_hours - base_hours
        overtime_hours = (overflow // 0.5) * 0.5 if overflow > 0 else 0
        overtime_hours += manual_add_ot
        final_status = "已套用異常覆寫" if has_override else "正常結算"
            
        results.append({"日期": date, "員工": emp, "身份": "正職", "班別": shift_str, "遲到(分)": late_mins, "早退(分)": early_leave_mins, "加班(時)": overtime_hours, "總工時(時)": round(total_calculated_hours, 2), "狀態": final_status})
        if has_override: audit_logs.append({"日期": date, "員工": emp, "原始判定": "異常/正常結算", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)})

    return pd.DataFrame(results), pd.DataFrame(audit_logs)

# ==========================================
# 模組四：最終薪資報表產出引擎 (全動態獎金解析)
# ==========================================
def parse_salary_params(file):
    try:
        df_fixed = pd.read_excel(file, sheet_name="固定參數")
        df_var = pd.read_excel(file, sheet_name="本月浮動獎金")
        
        df_fixed.columns = df_fixed.columns.str.strip()
        df_var.columns = df_var.columns.str.strip()
        
        num_cols_fixed = ['本薪或時薪', '勞保扣款', '健保扣款', '租屋補助']
        for col in num_cols_fixed:
            if col in df_fixed.columns:
                df_fixed[col] = pd.to_numeric(df_fixed[col], errors='coerce').fillna(0)
                
        # 動態掃描所有浮動獎金欄位
        exclude_cols = ['員工姓名', '特殊節日加給(時數)']
        dynamic_bonus_cols = [c for c in df_var.columns if c not in exclude_cols]
        
        for col in dynamic_bonus_cols + (['特殊節日加給(時數)'] if '特殊節日加給(時數)' in df_var.columns else []):
            df_var[col] = pd.to_numeric(df_var[col], errors='coerce').fillna(0)
                
        return df_fixed, df_var, dynamic_bonus_cols, ""
    except Exception as e:
        return None, None, None, "薪資與獎金設定表讀取失敗，請確認檔案包含「固定參數」與「本月浮動獎金」兩個工作表。"

def generate_final_payslip(df_calc, df_fixed, df_var, dynamic_bonus_cols):
    summary = df_calc.groupby('員工').agg({
        '遲到(分)': 'sum',
        '早退(分)': 'sum',
        '加班(時)': 'sum',
        '總工時(時)': 'sum',
        '身份': 'first'
    }).reset_index()
    
    payslip_data = []
    
    for _, emp_data in summary.iterrows():
        emp_name = emp_data['員工']
        emp_type = emp_data['身份']
        
        fixed_record = df_fixed[df_fixed['員工姓名'] == emp_name] if not df_fixed.empty and '員工姓名' in df_fixed.columns else pd.DataFrame()
        var_record = df_var[df_var['員工姓名'] == emp_name] if not df_var.empty and '員工姓名' in df_var.columns else pd.DataFrame()
        
        base_salary_or_hourly = fixed_record['本薪或時薪'].values[0] if not fixed_record.empty and '本薪或時薪' in fixed_record.columns else 0
        labor_ins = fixed_record['勞保扣款'].values[0] if not fixed_record.empty and '勞保扣款' in fixed_record.columns else 0
        health_ins = fixed_record['健保扣款'].values[0] if not fixed_record.empty and '健保扣款' in fixed_record.columns else 0
        rent_subsidy = fixed_record['租屋補助'].values[0] if not fixed_record.empty and '租屋補助' in fixed_record.columns else 0
        
        exact_hourly_rate = base_salary_or_hourly / 240.0 if emp_type == "正職" and base_salary_or_hourly > 0 else base_salary_or_hourly
        display_hourly_rate = custom_round(exact_hourly_rate)
        
        earned_bonuses = {}
        total_variable_bonus = 0
        
        if not var_record.empty:
            for col in dynamic_bonus_cols:
                val = var_record[col].values[0]
                if val > 0:
                    earned_bonuses[col] = int(val)
                    total_variable_bonus += val
                    
            if '特殊節日加給(時數)' in var_record.columns:
                sh_hours = var_record['特殊節日加給(時數)'].values[0]
                if sh_hours > 0:
                    special_val = custom_round(exact_hourly_rate * sh_hours * 1.5)
                    if special_val > 0:
                        earned_bonuses['特殊節日加成'] = special_val
                        total_variable_bonus += special_val

        if emp_type == "PT":
            base_pay = 0
            time_deduction = 0
            work_pay = custom_round(emp_data['總工時(時)'] * exact_hourly_rate)
            ot_pay = custom_round(emp_data['加班(時)'] * exact_hourly_rate) 
            gross_pay = work_pay + ot_pay + rent_subsidy + total_variable_bonus
        else:
            base_pay = base_salary_or_hourly
            total_penalty_mins = emp_data['遲到(分)'] + emp_data['早退(分)']
            time_deduction = custom_round(total_penalty_mins * (exact_hourly_rate / 60.0))
            ot_pay = custom_round(emp_data['加班(時)'] * exact_hourly_rate)
            gross_pay = base_pay + ot_pay + rent_subsidy + total_variable_bonus - time_deduction
            
        net_pay = gross_pay - labor_ins - health_ins
        
        record = {
            "員工姓名": emp_name,
            "身份": emp_type,
            "精算時薪": display_hourly_rate,
            "本薪/PT基礎薪": base_pay if emp_type == "正職" else work_pay,
            "總工時": emp_data['總工時(時)'],
            "加班時數": emp_data['加班(時)'],
            "加班加給": ot_pay,
            "動態獎金明細": earned_bonuses,
            "遲到早退合計(分)": emp_data['遲到(分)'] + emp_data['早退(分)'],
            "出勤扣款": -time_deduction if time_deduction > 0 else 0,
            "各項獎金總計": total_variable_bonus,
            "租屋補助": rent_subsidy,
            "勞健保扣款": -(labor_ins + health_ins) if (labor_ins + health_ins) > 0 else 0,
            "本月實領薪資": net_pay
        }
        payslip_data.append(record)
        
    return payslip_data

# ==========================================
# 模組五：個人薪資單渲染與打包引擎
# ==========================================
def create_payslip_text(record, month_str):
    lines = []
    lines.append("=======================================")
    lines.append("           IKKON 薪資明細表")
    lines.append("=======================================")
    lines.append(f"發放月份：{month_str}")
    lines.append(f"員工姓名：{record['員工姓名']} ({record['身份']})")
    lines.append("---------------------------------------")
    lines.append("【基本薪資】")
    
    if record['身份'] == "正職":
        lines.append(f"本薪 / 基礎薪： {int(record['本薪/PT基礎薪']):>12,}")
    else:
        lines.append(f"出勤薪資({record['總工時']}H)：{int(record['本薪/PT基礎薪']):>11,}")
        
    lines.append(f"精算時薪：      {int(record['精算時薪']):>12,}")
    lines.append("")
    lines.append("【加項與獎金】")
    
    has_bonus = False
    if record['加班時數'] > 0:
        lines.append(f"加班加給({record['加班時數']}H)： {int(record['加班加給']):>11,}")
        has_bonus = True
        
    if record['租屋補助'] > 0:
        lines.append(f"租屋補助：      {int(record['租屋補助']):>12,}")
        has_bonus = True

    for b_name, b_val in record['動態獎金明細'].items():
        lines.append(f"{b_name}： {b_val:>12,}")
        has_bonus = True
        
    if not has_bonus:
        lines.append("無")
        
    lines.append("---------------------------------------")
    total_adds = record['各項獎金總計'] + record['加班加給'] + record['租屋補助']
    lines.append(f"加項與獎金總計：{int(total_adds):>12,}")
    lines.append("")
    
    lines.append("【扣項】")
    lines.append(f"出勤扣款({record['遲到早退合計(分)']}分)： {int(record['出勤扣款']):>11,}")
    lines.append(f"勞健保扣款：    {int(record['勞健保扣款']):>12,}")
    
    lines.append("=======================================")
    lines.append(f"本月實領薪資：  {int(record['本月實領薪資']):>12,}")
    lines.append("=======================================")
    lines.append("辛苦了，謝謝你本月的付出！")
    
    return "\n".join(lines)

def create_zip_archive(payslips, month_str):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for p in payslips:
            text_content = create_payslip_text(p, month_str)
            zip_file.writestr(f"{p['員工姓名']}_{month_str}薪資單.txt", text_content.encode('utf-8'))
    return zip_buffer.getvalue()

# ==========================================
# 介面渲染：兩階段防禦性解耦架構
# ==========================================
st.set_page_config(page_title="IKKON 薪資自動化結算系統", layout="wide")
st.title("IKKON 薪資自動化結算系統")

st.markdown("---")
st.markdown("### 階段一：出缺勤診斷與異常覆寫")

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
            selected_sheet = st.selectbox("請選擇班表月份 (工作表)：", sheet_names)
        except Exception as e:
            st.error("讀取班表失敗。")
with col3:
    anomaly_file = st.file_uploader("3. 上傳 異常表 (若無可略過)", type=["csv", "xlsx"], key="anomaly")
    anomaly_selected_sheet = None
    if anomaly_file and anomaly_file.name.endswith('.xlsx'):
        try:
            xls_anomaly = pd.ExcelFile(anomaly_file)
            anomaly_sheet_names = xls_anomaly.sheet_names
            anomaly_selected_sheet = st.selectbox("請選擇異常表月份 (工作表)：", anomaly_sheet_names)
        except Exception as e:
            st.error("讀取異常表分頁失敗。")

if 'df_final_calc' not in st.session_state:
    st.session_state.df_final_calc = pd.DataFrame()

if ichef_file and roster_file and selected_sheet:
    if st.button("執行第一階段：出缺勤試算"):
        with st.spinner('進行時間碰撞與異常診斷中...'):
            df_cleaned, df_error = clean_ichef_data(ichef_file)
            df_roster, error_msg = parse_roster_data(roster_file, selected_sheet)
            
            if error_msg:
                st.error(error_msg)
            else:
                df_anomaly = parse_standard_anomaly_data(anomaly_file, anomaly_selected_sheet)
                df_final_calc, df_audit = calculate_payroll_hours(df_roster, df_cleaned, df_anomaly)
                
                st.session_state.df_final_calc = df_final_calc
                st.success("第一階段運算完成。請於下方報表查閱異常攔截紀錄與每日出缺勤明細。")
                
                tab_main, tab_audit, tab_error = st.tabs([
                    "每日出缺勤明細 (試算結果)", "異常表覆寫稽核", "原始打卡異常攔截 (需人工查核)"
                ])
                
                with tab_main: 
                    st.dataframe(df_final_calc)
                    
                with tab_audit: 
                    if not df_audit.empty:
                        st.dataframe(df_audit)
                    else:
                        st.info("本次無覆寫紀錄。")
                        
                with tab_error: 
                    if not df_error.empty:
                        st.dataframe(df_error)
                    else:
                        st.write("無異常紀錄。")

st.markdown("---")
st.markdown("### 階段二：最終薪資發放與個人薪資單產出")

salary_param_file = st.file_uploader("4. 上傳 薪資與獎金設定表 (包含全動態獎金)", type=["xlsx"], key="salary")

if salary_param_file and not st.session_state.df_final_calc.empty:
    if st.button("執行第二階段：產出個人薪資單"):
        with st.spinner('結合薪資基準與浮動獎金結算中...'):
            df_fixed, df_var, dyn_cols, err = parse_salary_params(salary_param_file)
            if err:
                st.error(err)
            else:
                payslip_records = generate_final_payslip(st.session_state.df_final_calc, df_fixed, df_var, dyn_cols)
                
                # 產出 ZIP 壓縮檔
                zip_data = create_zip_archive(payslip_records, selected_sheet)
                
                st.success("結算完成！你可以直接預覽下方卡片（可截圖），或點擊下方按鈕下載全體員工的薪資單文字檔。")
                st.download_button(
                    label="📥 下載全體員工薪資單 (ZIP 壓縮檔)",
                    data=zip_data,
                    file_name=f"IKKON_薪資單_{selected_sheet}.zip",
                    mime="application/zip"
                )
                
                # 在介面上動態產生卡片供預覽或截圖
                st.markdown("#### 薪資單預覽區 (可直接截圖或複製文字)")
                cols = st.columns(3)
                for idx, record in enumerate(payslip_records):
                    with cols[idx % 3]:
                        slip_text = create_payslip_text(record, selected_sheet)
                        st.code(slip_text, language='text')

elif salary_param_file and st.session_state.df_final_calc.empty:
    st.warning("請先完成「第一階段：出缺勤試算」，再執行薪資發放。")
