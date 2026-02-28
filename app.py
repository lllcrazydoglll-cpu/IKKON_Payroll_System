import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# ==========================================
# 模組一：打卡紀錄清洗 (維持原有的高容錯邏輯)
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
                        pass # 10分鐘內容錯，保留第一筆
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
# 模組二 (核心A)：攤平二維班表 (視角轉換)
# ==========================================
def parse_roster_data(file):
    raw_roster = pd.read_excel(file, header=None)
    roster_list = []
    
    # 尋找「姓名」所在的列數，建立欄位與員工的對應表
    name_row_index = -1
    for index, row in raw_roster.iterrows():
        if "姓名" in str(row.values):
            name_row_index = index
            break
            
    if name_row_index == -1:
        return None, "找不到「姓名」標籤，請確認班表格式是否正確。"
        
    # 紀錄哪個直行是哪位員工 (防禦：排除空白與標題)
    name_map = {}
    for col_idx, val in enumerate(raw_roster.iloc[name_row_index]):
        val_str = str(val).strip()
        if val_str and val_str not in ["nan", "姓名"]:
            name_map[col_idx] = val_str
            
    # 開始往下讀取每日排班資料
    for index in range(name_row_index + 1, len(raw_roster)):
        row = raw_roster.iloc[index]
        date_str = str(row[0]).strip()
        
        # 防禦機制：只抓取開頭是 202 的日期列 (例如 2026-01-01)
        if date_str.startswith("202"):
            # 遍歷有員工名字的直行，抓取當日班別
            for col_idx, employee_name in name_map.items():
                shift_val = str(row[col_idx]).strip()
                
                # 若儲存格內包含 "-" 代表有排定時間 (例如 1100-2200)
                if shift_val and "-" in shift_val:
                    roster_list.append({
                        "日期": date_str[:10],
                        "員工": employee_name,
                        "班別字串": shift_val
                    })
                    
    return pd.DataFrame(roster_list), ""

# ==========================================
# 模組二 (核心B)：商業邏輯運算 (遲到、早退福利、加班)
# ==========================================
def calculate_payroll_hours(df_roster, df_actual):
    results = []
    
    # 將實際打卡紀錄轉換為時間格式，方便後續運算
    df_actual['上班時間'] = pd.to_datetime(df_actual['上班時間'])
    df_actual['下班時間'] = pd.to_datetime(df_actual['下班時間'])
    df_actual['日期'] = df_actual['上班時間'].dt.strftime('%Y-%m-%d')
    
    # 逐筆檢視排班表，去跟實際打卡碰撞
    for _, scheduled in df_roster.iterrows():
        date = scheduled['日期']
        emp = scheduled['員工']
        shift_str = scheduled['班別字串'] # 例如 "1100-2200"
        
        # 拆解預定上下班時間字串
        try:
            start_str, end_str = shift_str.split('-')
            # 補齊格式 (1100 -> 11:00)
            start_str = f"{start_str[:2]}:{start_str[2:]}"
            end_str = f"{end_str[:2]}:{end_str[2:]}"
            
            sched_in = pd.to_datetime(f"{date} {start_str}")
            sched_out = pd.to_datetime(f"{date} {end_str}")
            
            # 【防禦機制】燒肉店跨夜處理
            if sched_out < sched_in:
                sched_out += timedelta(days=1)
                
        except:
            continue # 若班別格式錯誤則跳過
            
        # 篩選該員工當日的實際打卡紀錄
        emp_punches = df_actual[(df_actual['員工'] == emp) & (df_actual['日期'] == date)]
        
        if emp_punches.empty:
            results.append({"日期": date, "員工": emp, "班別": shift_str, "遲到(分)": 0, "早退(分)": 0, "加班(時)": 0, "狀態": "無打卡紀錄(休假或曠職)"})
            continue
            
        # 取得當日「最早的上班」與「最晚的下班」
        actual_in = emp_punches['上班時間'].min()
        actual_out = emp_punches['下班時間'].max()
        
        # 1. 計算遲到 (大於預定時間才算)
        late_mins = 0
        if actual_in > sched_in:
            late_mins = int((actual_in - sched_in).total_seconds() / 60)
            
        # 2. 計算早退與【選項B：福利虛擬工時】
        early_leave_mins = 0
        welfare_virtual_hours = 0
        
        if actual_out < sched_out:
            diff_mins = int((sched_out - actual_out).total_seconds() / 60)
            if diff_mins <= 30:
                # 觸發福利：早退歸零，並把這段時間轉換為虛擬工時
                early_leave_mins = 0
                welfare_virtual_hours = diff_mins / 60.0
            else:
                early_leave_mins = diff_mins
                
        # 3. 計算實際總待店工時 (加總當天所有打卡區間，精準扣除空班)
        total_actual_hours = 0
        for _, punch in emp_punches.iterrows():
            total_actual_hours += (punch['下班時間'] - punch['上班時間']).total_seconds() / 3600.0
            
        # 注入福利虛擬工時
        final_calculated_hours = total_actual_hours + welfare_virtual_hours
        
        # 4. 加班費計算邏輯
        overtime_hours = 0
        if "1100" in start_str and ("2200" in end_str or "2300" in end_str):
            # 兩頭全天班，基準為 8.5 小時
            overflow = final_calculated_hours - 8.5
        else:
            # 單班，基準為表定時數
            sched_total = (sched_out - sched_in).total_seconds() / 3600.0
            overflow = final_calculated_hours - sched_total
            
        # 防禦：以 0.5 小時為單位向下取整
        if overflow > 0:
            overtime_hours = (overflow // 0.5) * 0.5
            
        results.append({
            "日期": date, 
            "員工": emp, 
            "班別": shift_str, 
            "遲到(分)": late_mins, 
            "早退(分)": early_leave_mins, 
            "加班(時)": overtime_hours, 
            "狀態": "正常結算"
        })
        
    return pd.DataFrame(results)

# ==========================================
# 系統介面 (UI) 設計
# ==========================================
st.set_page_config(page_title="IKKON 薪資結算系統", layout="wide")
st.title("IKKON 薪資自動化結算系統")

st.markdown("### 步驟一：上傳原始資料")
col_upload1, col_upload2 = st.columns(2)
with col_upload1:
    ichef_file = st.file_uploader("1. 請上傳 iCHEF 打卡紀錄 (.xlsx)", type=["xlsx"], key="ichef")
with col_upload2:
    roster_file = st.file_uploader("2. 請上傳 店鋪當月班表 (.xlsx)", type=["xlsx"], key="roster")

if ichef_file and roster_file:
    if st.button("執行自動化結算"):
        with st.spinner('系統運算中 (資料清洗 ➔ 班表攤平 ➔ 邏輯碰撞)...'):
            
            # 執行模組一
            df_cleaned, df_error = clean_ichef_data(ichef_file)
            
            # 執行模組二
            df_roster, error_msg = parse_roster_data(roster_file)
            
            if error_msg:
                st.error(error_msg)
            else:
                # 執行薪資工時碰撞計算
                df_final_calc = calculate_payroll_hours(df_roster, df_cleaned)
                
                st.success("✅ 運算完成！")
                
                tab1, tab2, tab3 = st.tabs(["📊 最終出缺勤結算", "⚠️ 需人工確認之異常打卡", "🔍 系統攤平後之班表 (除錯用)"])
                
                with tab1:
                    st.markdown("#### 自動計算結果 (含跨夜判定、早退福利、精準加班)")
                    st.dataframe(df_final_calc)
                    
                with tab2:
                    st.markdown("#### 異常打卡紀錄攔截")
                    if not df_error.empty:
                        st.warning("請經理確認以下紀錄是否需補登工時")
                        st.dataframe(df_error)
                    else:
                        st.write("完美！無任何異常紀錄。")
                        
                with tab3:
                    st.markdown("#### 這是電腦眼中看懂的班表 (視角轉換結果)")
                    st.dataframe(df_roster)
