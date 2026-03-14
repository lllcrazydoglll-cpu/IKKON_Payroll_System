import streamlit as st
import pandas as pd
import math
import io
import zipfile
import os
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime, timedelta

# ==========================================
# 會計級精算引擎
# ==========================================
def custom_round(n):
    return int(math.floor(n + 0.5))

def custom_round_2(n):
    return math.floor(n * 100 + 0.5) / 100.0

def fmt(val):
    s = f"{val:,.2f}"
    if s.endswith(".00"):
        return s[:-3]
    if s.endswith("0"):
        return s[:-1]
    return s

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
        
        waive_penalty = False 
        
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
                    waive_penalty = True
                    override_reasons.append(f"調休變更: {reason}")
                elif cmd == "變更為應勤":
                    shift_str = "正常班"
                    is_working = True
                    has_override = True
                    waive_penalty = True
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
                        waive_penalty = True 
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

        if waive_penalty:
            late_mins = 0
            early_leave_mins = 0
                
        overflow = total_calculated_hours - base_hours
        overtime_hours = (overflow // 0.5) * 0.5 if overflow > 0 else 0
        overtime_hours += manual_add_ot
        final_status = "已套用異常覆寫" if has_override else "正常結算"
            
        results.append({"日期": date, "員工": emp, "身份": "正職", "班別": shift_str, "遲到(分)": late_mins, "早退(分)": early_leave_mins, "加班(時)": overtime_hours, "總工時(時)": round(total_calculated_hours, 2), "狀態": final_status})
        if has_override: audit_logs.append({"日期": date, "員工": emp, "原始判定": "異常/正常結算", "覆寫內容": "已執行上述指令", "幹部備註原因": " | ".join(override_reasons)})

    return pd.DataFrame(results), pd.DataFrame(audit_logs)

# ==========================================
# 模組四：人事資料庫主導之最終薪資引擎
# ==========================================
def parse_salary_params(file):
    try:
        df_fixed = pd.read_excel(file, sheet_name="固定參數")
        df_var = pd.read_excel(file, sheet_name="本月浮動獎金")
        
        try:
            df_hr_reward = pd.read_excel(file, sheet_name="時數獎勵")
            df_hr_reward.columns = df_hr_reward.columns.str.strip()
        except Exception:
            df_hr_reward = pd.DataFrame()
            
        df_fixed.columns = df_fixed.columns.str.strip()
        df_var.columns = df_var.columns.str.strip()
        
        core_fixed_cols = ['部門', '員工姓名', '身份(正職或PT)', '本薪或時薪', '勞保扣款', '健保扣款']
        dynamic_fixed_cols = [c for c in df_fixed.columns if c not in core_fixed_cols]
        for col in dynamic_fixed_cols:
            df_fixed[col] = pd.to_numeric(df_fixed[col], errors='coerce').fillna(0)
            
        exclude_var_cols = ['部門', '員工姓名', '特殊節日加給(時數)']
        dynamic_bonus_cols = [c for c in df_var.columns if c not in exclude_var_cols]
        
        for col in dynamic_bonus_cols + (['特殊節日加給(時數)'] if '特殊節日加給(時數)' in df_var.columns else []):
            df_var[col] = pd.to_numeric(df_var[col], errors='coerce').fillna(0)
            
        hr_reward_pairs = []
        if not df_hr_reward.empty:
            for col in df_hr_reward.columns:
                if str(col).endswith('(時數)'):
                    base_name = col.replace('(時數)', '')
                    mult_col = f"{base_name}(倍數)"
                    
                    df_hr_reward[col] = pd.to_numeric(df_hr_reward[col], errors='coerce').fillna(0)
                    
                    if mult_col in df_hr_reward.columns:
                        df_hr_reward[mult_col] = pd.to_numeric(df_hr_reward[mult_col], errors='coerce').fillna(1.0)
                    else:
                        df_hr_reward[mult_col] = 1.0
                        
                    hr_reward_pairs.append((col, mult_col, base_name))
                
        return df_fixed, df_var, dynamic_bonus_cols, dynamic_fixed_cols, df_hr_reward, hr_reward_pairs, ""
    except Exception as e:
        return None, None, None, None, None, None, "薪資與獎金設定表讀取失敗，請確認檔案結構。"

def generate_final_payslip(df_calc, df_fixed, df_var, dynamic_bonus_cols, dynamic_fixed_cols, df_hr_reward, hr_reward_pairs):
    if not df_calc.empty:
        summary = df_calc.groupby('員工').agg({
            '遲到(分)': 'sum',
            '早退(分)': 'sum',
            '加班(時)': 'sum',
            '總工時(時)': 'sum',
            '身份': 'first'
        }).reset_index()
    else:
        return []
        
    payslip_data = []
    
    for _, emp_data in summary.iterrows():
        emp_name = emp_data['員工']
        emp_type = emp_data['身份']
        
        fixed_record = df_fixed[df_fixed['員工姓名'] == emp_name] if not df_fixed.empty and '員工姓名' in df_fixed.columns else pd.DataFrame()
        var_record = df_var[df_var['員工姓名'] == emp_name] if not df_var.empty and '員工姓名' in df_var.columns else pd.DataFrame()
        hr_record = df_hr_reward[df_hr_reward['員工姓名'] == emp_name] if not df_hr_reward.empty and '員工姓名' in df_hr_reward.columns else pd.DataFrame()
        
        base_salary_or_hourly = float(fixed_record['本薪或時薪'].values[0]) if not fixed_record.empty and pd.notna(fixed_record['本薪或時薪'].values[0]) else 0.0
        
        # 【修正 BUG】在此處移除提早取整，讓 exact_hourly_rate 保持無限精度浮點數（例如 208.333333...），以符合財務標準。
        exact_hourly_rate = float(base_salary_or_hourly / 240.0) if emp_type == "正職" and base_salary_or_hourly > 0 else float(base_salary_or_hourly)
        
        labor_ins = float(fixed_record['勞保扣款'].values[0]) if not fixed_record.empty and '勞保扣款' in df_fixed.columns and pd.notna(fixed_record['勞保扣款'].values[0]) else 0.0
        health_ins = float(fixed_record['健保扣款'].values[0]) if not fixed_record.empty and '健保扣款' in df_fixed.columns and pd.notna(fixed_record['健保扣款'].values[0]) else 0.0
        
        earned_bonuses = {}
        deductions = {}
        total_variable_bonus = 0.0
        special_holiday_bonus = 0.0
        total_other_deductions = 0.0
        
        if not fixed_record.empty:
            for col in dynamic_fixed_cols:
                val = float(fixed_record[col].values[0])
                if val > 0:
                    earned_bonuses[col] = val
                    total_variable_bonus += val
                elif val < 0:
                    deductions[col] = abs(val)
                    total_other_deductions += abs(val)

        if not var_record.empty:
            for col in dynamic_bonus_cols:
                val = float(var_record[col].values[0])
                if val > 0:
                    earned_bonuses[col] = val
                    total_variable_bonus += val
                elif val < 0:
                    deductions[col] = abs(val)
                    total_other_deductions += abs(val)
                    
            if '特殊節日加給(時數)' in var_record.columns:
                sh_hours = float(var_record['特殊節日加給(時數)'].values[0])
                if sh_hours > 0:
                    special_val = custom_round_2(exact_hourly_rate * sh_hours * 1.5)
                    if special_val > 0:
                        earned_bonuses['特殊節日加成(1.5倍)'] = special_val
                        total_variable_bonus += special_val
                        special_holiday_bonus += special_val

        if not hr_record.empty:
            for hr_col, mult_col, base_name in hr_reward_pairs:
                h_val = float(hr_record[hr_col].values[0])
                m_val = float(hr_record[mult_col].values[0])
                if h_val > 0:
                    # 使用無限精度時薪計算後，將個別獎金結果四捨五入到小數點後 2 位
                    calculated_val = custom_round_2(exact_hourly_rate * h_val * m_val)
                    if calculated_val > 0:
                        display_name = f"{base_name}({m_val}倍)" if m_val != 1.0 else base_name
                        earned_bonuses[display_name] = calculated_val
                        total_variable_bonus += calculated_val
                        special_holiday_bonus += calculated_val

        if emp_type == "PT":
            base_pay = 0.0
            time_deduction = 0.0
            work_pay = custom_round_2(emp_data['總工時(時)'] * exact_hourly_rate)
            ot_pay = custom_round_2(emp_data['加班(時)'] * exact_hourly_rate) 
            gross_pay = work_pay + ot_pay + total_variable_bonus
        else:
            base_pay = base_salary_or_hourly
            total_penalty_mins = emp_data['遲到(分)'] + emp_data['早退(分)']
            time_deduction = custom_round_2(total_penalty_mins * (exact_hourly_rate / 60.0))
            ot_pay = custom_round_2(emp_data['加班(時)'] * exact_hourly_rate)
            gross_pay = base_pay + ot_pay + total_variable_bonus - time_deduction
            
        net_pay = gross_pay - total_other_deductions - labor_ins - health_ins
        
        record = {
            "員工姓名": emp_name,
            "身份": emp_type,
            "精算時薪": custom_round_2(exact_hourly_rate), # 僅用於顯示排版與 Excel 留底
            "本薪/PT基礎薪": base_pay if emp_type == "正職" else work_pay,
            "總工時": emp_data['總工時(時)'],
            "加班時數": emp_data['加班(時)'],
            "加班加給": ot_pay,
            "動態加項明細": earned_bonuses,
            "動態扣項明細": deductions,
            "特殊節日加成金額": special_holiday_bonus,
            "遲到早退合計(分)": emp_data['遲到(分)'] + emp_data['早退(分)'],
            "出勤扣款": time_deduction,
            "各項獎金與津貼總計": total_variable_bonus,
            "各項扣款總計": total_other_deductions,
            "應發薪資(毛額)": gross_pay,
            "勞健保扣款": -(labor_ins + health_ins) if (labor_ins + health_ins) > 0 else 0.0,
            "本月實領薪資": custom_round(net_pay)
        }
        payslip_data.append(record)
        
    return payslip_data

# ==========================================
# 模組五：會計統計總表產生器
# ==========================================
def generate_accounting_excel(payslip_records, revenue):
    df = pd.DataFrame(payslip_records)
    if df.empty: return io.BytesIO().getvalue()
    
    ft_total = df[df['身份'] == '正職']['應發薪資(毛額)'].sum()
    pt_total = df[df['身份'] == 'PT']['應發薪資(毛額)'].sum()
    ot_total = df['加班加給'].sum()
    bonus_total = df['各項獎金與津貼總計'].sum()
    
    total_cost = ft_total + pt_total
    cost_ratio = f"{(total_cost / revenue * 100):.2f}%" if revenue > 0 else "0%"
    
    summary_data = {
        "統計項目": ["本月營業總額", "總人事成本 (正職+兼職)", "人事成本佔比", "正職薪資合計 (含獎金加班)", "兼職(PT)薪資合計 (含獎金)", "全體加班費合計", "全體各項獎金合計"],
        "金額 / 數據": [f"{int(revenue):,}", fmt(total_cost), cost_ratio, fmt(ft_total), fmt(pt_total), fmt(ot_total), fmt(bonus_total)]
    }
    df_summary = pd.DataFrame(summary_data)
    df_detailed = df.drop(columns=['動態加項明細', '動態扣項明細'])
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_summary.to_excel(writer, sheet_name='會計統計報表', index=False)
        df_detailed.to_excel(writer, sheet_name='員工薪資明細', index=False)
        
        worksheet1 = writer.sheets['會計統計報表']
        worksheet1.set_column('A:A', 30)
        worksheet1.set_column('B:B', 20)
        
    return output.getvalue()

# ==========================================
# 模組六：絕對防禦 JPG 薪資圖檔生成引擎
# ==========================================
def get_text_width(draw, text, font):
    try:
        return draw.textlength(text, font=font)
    except AttributeError:
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            return bbox[2] - bbox[0]
        except AttributeError:
            return draw.textsize(text, font=font)[0]

def split_text_into_lines(text, max_chars_per_line=22):
    lines = []
    while len(text) > max_chars_per_line:
        lines.append(text[:max_chars_per_line])
        text = text[max_chars_per_line:]
    if text:
        lines.append(text)
    return lines

def create_payslip_image(record, month_str, custom_msg):
    font_path = "NotoSansTC-Regular.ttf"
    try:
        font = ImageFont.truetype(font_path, 20)
        font_title = ImageFont.truetype(font_path, 26)
        font_bold = ImageFont.truetype(font_path, 22)
    except OSError:
        font = ImageFont.load_default()
        font_title = font
        font_bold = font

    msg_lines = []
    if custom_msg:
        raw_lines = custom_msg.split('\n')
        for raw_l in raw_lines:
            msg_lines.extend(split_text_into_lines(raw_l, 24))

    base_h = 700
    bonus_count = len(record['動態加項明細'])
    deduction_count = len(record['動態扣項明細'])
    msg_count = len(msg_lines)
    img_h = base_h + (bonus_count * 35) + (deduction_count * 35) + (msg_count * 35)

    img = Image.new('RGB', (550, img_h), color='#FFFFFF')
    draw = ImageDraw.Draw(img)

    y = 30
    margin = 40
    right = 510

    def line_light():
        nonlocal y
        draw.line([(margin, y), (right, y)], fill="#CCCCCC", width=1)
        y += 20

    def text_center(text, f=font):
        nonlocal y
        w = get_text_width(draw, text, f)
        draw.text(((550 - w) / 2, y), text, font=f, fill="#000000")
        y += 35

    def text_left(text, f=font):
        nonlocal y
        draw.text((margin, y), text, font=f, fill="#000000")
        y += 35

    def text_row(label, val, f=font):
        nonlocal y
        draw.text((margin, y), label, font=f, fill="#000000")
        w = get_text_width(draw, str(val), f)
        draw.text((right - w, y), str(val), font=f, fill="#000000")
        y += 35

    draw.line([(margin, y), (right, y)], fill="#000000", width=3)
    y += 8 
    w = get_text_width(draw, "IKKON 薪資明細表", font_title)
    draw.text(((550 - w) / 2, y), "IKKON 薪資明細表", font=font_title, fill="#000000")
    y += 44 
    draw.line([(margin, y), (right, y)], fill="#000000", width=3)
    y += 25 

    text_left(f"發放月份：{month_str}", f=font_bold)
    text_left(f"員工姓名：{record['員工姓名']} ({record['身份']})", f=font_bold)
    line_light()

    text_left("【基本薪資】", f=font_bold)
    if record['身份'] == "正職":
        text_row("本薪 / 基礎薪：", fmt(record['本薪/PT基礎薪']))
    else:
        text_row(f"出勤薪資({record['總工時']}H)：", fmt(record['本薪/PT基礎薪']))
    text_row("精算時薪：", fmt(record['精算時薪']))
    y += 10

    text_left("【加項與獎金】", f=font_bold)
    has_bonus = False
    if record['加班時數'] > 0:
        text_row(f"加班加給({record['加班時數']}H)：", fmt(record['加班加給']))
        has_bonus = True

    for b_name, b_val in record['動態加項明細'].items():
        text_row(f"{b_name}：", fmt(b_val))
        has_bonus = True

    if not has_bonus:
        text_row("無：", "0")

    y += 5
    line_light()
    total_adds = record['各項獎金與津貼總計'] + record['加班加給']
    text_row("加項與獎金總計：", fmt(total_adds), f=font_bold)
    y += 10

    text_left("【扣項】", f=font_bold)
    text_row(f"出勤扣款({record['遲到早退合計(分)']}分)：", f"-{fmt(record['出勤扣款'])}" if record['出勤扣款'] > 0 else "0")
    if record['勞健保扣款'] < 0:
        text_row("勞健保扣款：", fmt(record['勞健保扣款']))
        
    for d_name, d_val in record['動態扣項明細'].items():
        text_row(f"{d_name}：", f"-{fmt(d_val)}")
        
    y += 15

    draw.line([(margin, y), (right, y)], fill="#000000", width=3)
    y += 8 
    draw.text((margin, y), "本月實領薪資：", font=font_title, fill="#000000")
    val_str = f"{record['本月實領薪資']:,}"
    w = get_text_width(draw, val_str, font_title)
    draw.text((right - w, y), val_str, font=font_title, fill="#000000")
    y += 44 
    draw.line([(margin, y), (right, y)], fill="#000000", width=3)
    y += 20 

    if msg_lines:
        y += 10
        for line in msg_lines:
            text_center(line, f=font_bold)

    img = img.crop((0, 0, 550, y + 20))
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=95)
    return img_byte_arr.getvalue()

def create_zip_archive_images(payslips, month_str, custom_msg):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for p in payslips:
            img_bytes = create_payslip_image(p, month_str, custom_msg)
            zip_file.writestr(f"{p['員工姓名']}_{month_str}薪資單.jpg", img_bytes)
    return zip_buffer.getvalue()

# ==========================================
# 介面渲染：兩階段防禦性解耦架構 (Session State 保護)
# ==========================================
st.set_page_config(page_title="IKKON 薪資自動化結算系統", layout="wide")
st.title("IKKON 薪資自動化結算系統")

if not os.path.exists("NotoSansTC-Regular.ttf"):
    st.error("系統警告：尚未偵測到中文字體檔 (NotoSansTC-Regular.ttf)。請將該檔案上傳至 GitHub，否則產出的薪資圖檔將會顯示為亂碼。")

if 'df_final_calc' not in st.session_state:
    st.session_state.df_final_calc = pd.DataFrame()
if 'stage2_done' not in st.session_state:
    st.session_state.stage2_done = False
if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None

st.markdown("---")
st.markdown("### 階段一：出缺勤診斷與異常覆寫")

col1, col2, col3 = st.columns(3)
with col1:
    ichef_file = st.file_uploader("1. 上傳 iCHEF 打卡紀錄", type=["xlsx"], key="ichef")
with col2:
    roster_file = st.file_uploader("2. 上傳 店鋪當月班表 (總部點名單)", type=["xlsx"], key="roster")
    selected_sheet = None
    if roster_file:
        try:
            xls = pd.ExcelFile(roster_file)
            sheet_names = xls.sheet_names
            selected_sheet = st.selectbox("請選擇班表月份 (工作表)：", sheet_names)
        except Exception as e:
            st.error("讀取班表失敗。")
with col3:
    anomaly_file = st.file_uploader("3. 上傳 7欄位異常表 (若無可略過)", type=["csv", "xlsx"], key="anomaly")
    anomaly_selected_sheet = None
    if anomaly_file and anomaly_file.name.endswith('.xlsx'):
        try:
            xls_anomaly = pd.ExcelFile(anomaly_file)
            anomaly_sheet_names = xls_anomaly.sheet_names
            anomaly_selected_sheet = st.selectbox("請選擇異常表月份 (工作表)：", anomaly_sheet_names)
        except Exception as e:
            st.error("讀取異常表分頁失敗。")

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
                st.session_state.stage2_done = False
                st.session_state.zip_data = None
                st.session_state.excel_data = None
                
                st.success("第一階段運算完成。請於下方報表查閱異常攔截紀錄與每日出缺勤明細。")
                
                tab_main, tab_audit, tab_error = st.tabs([
                    "每日出缺勤明細 (試算結果)", "異常表覆寫稽核", "原始打卡異常攔截 (需人工查核)"
                ])
                
                with tab_main: 
                    st.dataframe(df_final_calc)
                with tab_audit: 
                    if not df_audit.empty: st.dataframe(df_audit)
                    else: st.info("本次無覆寫紀錄。")
                with tab_error: 
                    if not df_error.empty: st.dataframe(df_error)
                    else: st.write("無異常紀錄。")

st.markdown("---")
st.markdown("### 階段二：圖形化薪資單產出與會計報表")

col_a, col_b = st.columns(2)
with col_a:
    st.markdown("##### 1. 會計核算參數")
    revenue_input = st.number_input("請輸入本月營業總額 (供計算人事成本佔比)：", min_value=0, value=0, step=1000)
    salary_param_file = st.file_uploader("4. 上傳 薪資與獎金設定表 (支援無限欄位擴充)", type=["xlsx"], key="salary")
    
with col_b:
    st.markdown("##### 2. 薪資單發放設定")
    custom_msg = st.text_area("給同仁的當月結語 (將印在圖檔最下方)：", value="辛苦了，謝謝你本月的付出！", height=120)

if salary_param_file and not st.session_state.df_final_calc.empty:
    if st.button("執行第二階段：產出 JPG 薪資單與會計報表"):
        with st.spinner('結合薪資基準繪製圖檔與結算會計報表中...'):
            df_fixed, df_var, dyn_cols, dyn_fixed_cols, df_hr_reward, hr_pairs, err = parse_salary_params(salary_param_file)
            if err:
                st.error(err)
            else:
                payslip_records = generate_final_payslip(st.session_state.df_final_calc, df_fixed, df_var, dyn_cols, dyn_fixed_cols, df_hr_reward, hr_pairs)
                
                st.session_state.zip_data = create_zip_archive_images(payslip_records, selected_sheet, custom_msg)
                st.session_state.excel_data = generate_accounting_excel(payslip_records, revenue_input)
                st.session_state.stage2_done = True

    if st.session_state.get('stage2_done') and st.session_state.get('zip_data') and st.session_state.get('excel_data'):
        st.success("結算與繪製完成！請點擊下方按鈕下載檔案。")
        
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label="📥 下載全體員工 JPG 薪資圖檔 (ZIP)",
                data=st.session_state.zip_data,
                file_name=f"IKKON_薪資圖檔_{selected_sheet}.zip",
                mime="application/zip"
            )
        with dl_col2:
            st.download_button(
                label="📊 下載會計結算總表 (Excel)",
                data=st.session_state.excel_data,
                file_name=f"IKKON_會計結算總表_{selected_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

elif salary_param_file and st.session_state.df_final_calc.empty:
    st.warning("請先完成「第一階段：出缺勤試算」，再執行薪資發放。")
