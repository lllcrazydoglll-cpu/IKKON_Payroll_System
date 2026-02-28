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
# 模組二：強固型班表攤平 (解決 NaT 問題)
# ==========================================
def parse_roster_data(file):
    raw_roster = pd.read_excel(file, header=None)
    roster_list = []
    
    # 強制定位姓名列
    name_row_index = -1
    for index, row in raw_roster.iterrows():
        if "姓名" in str(row.values):
            name_row_index = index
            break
            
    if name_row_index == -1:
        return None, "找不到「姓名」標籤，請確認班表格式是否正確。"
        
    # 地毯式掃描：精準建立直行與員工的對應，強制排除空白與無效字眼
    name_map = {}
    invalid_names = ["nan", "姓名", "NaT", "None", ""]
    for col_idx, val in enumerate(raw_roster.iloc[name_row_index]):
        val_str = str(val).strip()
        if val_str not in invalid_names and not pd.isna(val):
            name_map[col_idx] = val_str
            
    # 開始讀取日期與排班
    for index in range(name_row_index + 1, len(raw_roster)):
        row = raw_roster.iloc[index]
        date_str = str(row[0]).strip()
        
        if date_str.startswith("202"):
            for col_idx, employee_name in name_map.items():
                shift_val = str(row[col_idx]).strip()
                if shift_val and "-" in shift_val and shift_val not in ["nan", "NaT"]:
                    roster_list.append({
                        "日期": date_str[:10],
                        "員工": employee_name,
                        "班別字串": shift_val
                    })
                    
    return pd.DataFrame(roster_list), ""

# ==========================================
# 模組三：系統介面與稽核架構
# ==========================================
st.set_page_config(page_title="IKKON 薪資結算系統", layout="wide")
st.title("IKKON 薪資自動化結算系統")

st.markdown("### 資料匯入區")
col1, col2, col3 = st.columns(3)
with col1:
    ichef_file = st.file_uploader("1. 上傳 iCHEF 打卡紀錄", type=["xlsx"], key="ichef")
with col2:
    roster_file = st.file_uploader("2. 上傳 店鋪當月班表", type=["xlsx"], key="roster")
with col3:
    anomaly_file = st.file_uploader("3. 上傳 幹部打卡異常表", type=["csv", "xlsx"], key="anomaly")

if ichef_file and roster_file:
    if st.button("執行結算與稽核比對"):
        with st.spinner('系統運算中...'):
            
            df_cleaned, df_error = clean_ichef_data(ichef_file)
            df_roster, error_msg = parse_roster_data(roster_file)
            
            if error_msg:
                st.error(error_msg)
            else:
                st.success("基礎資料解析成功。")
                
                # 建立沒有表情符號的標籤頁
                tab_main, tab_audit, tab_error, tab_roster = st.tabs([
                    "最終出缺勤結算", 
                    "異常表覆寫稽核", 
                    "原始打卡異常攔截", 
                    "系統攤平班表(除錯)"
                ])
                
                with tab_main:
                    st.markdown("#### 核心運算結果")
                    st.info("模組二工時碰撞與異常表覆寫邏輯即將實作於此。")
                    
                with tab_audit:
                    st.markdown("#### 主管手動覆寫稽核軌跡")
                    if anomaly_file:
                        st.info("已接收異常表，將於此處列出所有被強制覆寫的工時與主管備註，供經理二次核實。")
                    else:
                        st.warning("本次結算未上傳異常表，無覆寫紀錄。")
                        
                with tab_error:
                    st.markdown("#### 需人工確認之異常打卡")
                    if not df_error.empty:
                        st.dataframe(df_error)
                    else:
                        st.write("無任何底層異常紀錄。")
                        
                with tab_roster:
                    st.markdown("#### 系統解析之標準化班表")
                    st.dataframe(df_roster)
