import streamlit as st
import pandas as pd

def clean_ichef_data(file):
    cleaned_data = []
    error_log = []

    raw_data = pd.read_excel(file, header=None)

    current_employee = ""
    current_clock_in = None

    for index, row in raw_data.iterrows():
        action = str(row[0]).strip()
        time_record = str(row[1]).strip()
        
        # 定義 iCHEF 系統產生的所有「非員工姓名」關鍵字，建立絕對防禦網
        system_keywords = ["上班", "下班", "無下班", "無上班", "無下班記錄", "無上班記錄", "無下班紀錄", "無上班紀錄", "結帳收銀", "admin", "nan"]

        # 邏輯區塊一：精準辨識員工姓名
        is_employee = True
        if action in system_keywords or "總時數" in action:
            is_employee = False
            
        if is_employee and action != "":
            # 【防禦機制加強】換人前，檢查上一個人是不是有忘記下班的紀錄卡在暫存區
            if current_clock_in is not None:
                error_log.append({
                    "員工": current_employee,
                    "異常類型": "換人前無下班紀錄 (系統強制攔截)",
                    "打卡時間": current_clock_in
                })
            current_employee = action
            current_clock_in = None

        # 邏輯區塊二：處理「上班」
        elif action == "上班":
            if current_clock_in is not None:
                error_log.append({
                    "員工": current_employee,
                    "異常類型": "連續兩次上班打卡，無下班紀錄",
                    "打卡時間": current_clock_in
                })
            current_clock_in = time_record

        # 邏輯區塊三：處理「下班」
        elif action == "下班":
            if current_clock_in is not None:
                cleaned_data.append({
                    "員工": current_employee,
                    "上班時間": current_clock_in,
                    "下班時間": time_record
                })
                current_clock_in = None
            else:
                error_log.append({
                    "員工": current_employee,
                    "異常類型": "有下班打卡，但無上班紀錄",
                    "打卡時間": time_record
                })

        # 邏輯區塊四：處理 iCHEF 標記的各種無上下班異常
        elif "無下班" in action:
            error_log.append({
                "員工": current_employee,
                "異常類型": "iCHEF標記無下班，需人工確認",
                "打卡時間": current_clock_in if current_clock_in else time_record
            })
            current_clock_in = None
            
        elif "無上班" in action:
            error_log.append({
                "員工": current_employee,
                "異常類型": "iCHEF標記無上班，需人工確認",
                "打卡時間": time_record
            })
            current_clock_in = None

    return cleaned_data, error_log

# --- 系統介面 (UI) 設計 ---
st.set_page_config(page_title="IKKON 薪資結算系統", layout="wide")
st.title("IKKON 薪資自動化結算系統")
st.markdown("### 模組一：打卡紀錄清洗與異常攔截")

uploaded_file = st.file_uploader("請上傳 iCHEF 打卡紀錄 (Excel格式 .xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("執行資料清洗"):
        with st.spinner('系統處理中...'):
            cleaned_data, error_log = clean_ichef_data(uploaded_file)

            df_cleaned = pd.DataFrame(cleaned_data)
            df_error = pd.DataFrame(error_log)

            st.markdown("**處理完成！請核對以下資料：**")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("#### 正常打卡紀錄 (準備進入模組二)")
                if not df_cleaned.empty:
                    st.dataframe(df_cleaned)
                else:
                    st.write("無正常打卡紀錄")

            with col2:
                st.markdown("#### 異常紀錄 (觸發防禦機制)")
                if not df_error.empty:
                    st.markdown("**發現異常打卡，請人工確認以下紀錄。**")
                    st.dataframe(df_error)
                else:
                    st.write("完美！無任何異常紀錄。")
