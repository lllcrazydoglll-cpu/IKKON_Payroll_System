import streamlit as st
import pandas as pd

def clean_ichef_data(file):
    cleaned_data = []
    error_log = []

    # 讀取原始檔案，設定 header=None 是防禦機制，因為 iCHEF 原始檔沒有標準的標題列
    raw_data = pd.read_csv(file, header=None)

    current_employee = ""
    current_clock_in = None

    for index, row in raw_data.iterrows():
        action = str(row[0]).strip()
        time_record = str(row[1]).strip()

        # 邏輯區塊一：辨識員工姓名與排除雜訊
        # 增加防禦：排除 iCHEF 報表常出現的「總時數：...」等無關字眼
        if action not in ["上班", "下班", "無下班", "nan", "結帳收銀", "admin"]:
            if "總時數" not in action:
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

        # 邏輯區塊四：處理「無下班」
        elif "無下班" in action:
            error_log.append({
                "員工": current_employee,
                "異常類型": "系統標記無下班，需人工確認下班時間",
                "打卡時間": current_clock_in
            })
            current_clock_in = None

    return cleaned_data, error_log

# --- 系統介面 (UI) 設計 ---
st.set_page_config(page_title="IKKON 薪資結算系統", layout="wide")
st.title("IKKON 薪資自動化結算系統")
st.markdown("### 模組一：打卡紀錄清洗與異常攔截")

# 建立檔案上傳區塊
uploaded_file = st.file_uploader("請上傳 iCHEF 打卡紀錄 (CSV格式)", type=["csv"])

if uploaded_file is not None:
    # 建立執行按鈕
    if st.button("執行資料清洗"):
        with st.spinner('系統處理中...'):
            # 呼叫上方定義的邏輯函數
            cleaned_data, error_log = clean_ichef_data(uploaded_file)

            # 將結果轉換為資料表格式以便在網頁顯示
            df_cleaned = pd.DataFrame(cleaned_data)
            df_error = pd.DataFrame(error_log)

            st.markdown("**處理完成！請核對以下資料：**")

            # 將畫面切割為左右兩塊，便於對照
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
