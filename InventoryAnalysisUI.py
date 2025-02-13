import streamlit as st
import pandas as pd
import numpy as np


def process_inventory(file, keyword):
    # 讀取 Excel 檔案的所有工作表
    excel_data = pd.read_excel(file, sheet_name=None)

    # 確保檔案有至少一個工作表
    if not excel_data:
        raise ValueError("上傳的 Excel 檔案沒有任何工作表，請檢查檔案內容。")

    # 取得第一個工作表名稱並讀取
    first_sheet = list(excel_data.keys())[0]
    df = excel_data[first_sheet]

    # 確保 Excel 內容不為空
    if df.empty:
        raise ValueError("Excel 檔案沒有可讀取的數據，請確認檔案內容。")

    # 過濾品牌關鍵字（假設品牌名稱在 '商品名稱' 欄位內）
    df = df[df['商品名稱'].astype(str).str.contains(keyword, na=False, case=False)]

    # 只保留指定欄位
    keep_columns = ['商品名稱', '商品款式', '商品原價', '商品成本', '庫存總量', '總成本']
    df = df[keep_columns]

    # 移除庫存為 0 或負數的商品
    df = df[df['庫存總量'] > 0]

    # Highlight 商品成本為 NaN 的商品
    df['標記'] = np.where(df['商品成本'].isna(), '缺少成本', '')

    # 轉換數字欄位格式
    numeric_columns = ['商品原價', '商品成本', '庫存總量', '總成本']
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')

    # 計算總庫存數量與成本總額
    total_inventory = df['庫存總量'].sum()
    total_cost = df['總成本'].sum()

    # 建立報告 DataFrame
    summary_df = pd.DataFrame({
        '項目': ['總庫存數量', '總成本額'],
        '數值': [total_inventory, total_cost]
    })

    return df, summary_df


# Streamlit UI 設計
st.title("庫存分析工具")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])
keyword = st.text_input("請輸入品牌關鍵字")

if uploaded_file:
    if not uploaded_file.name.endswith(".xlsx"):
        st.error("請上傳有效的 Excel (.xlsx) 檔案！")
    else:
        if keyword:
            try:
                df_processed, summary_df = process_inventory(uploaded_file, keyword)

                # 顯示處理後的資料
                st.write("### 處理後的庫存數據")
                st.dataframe(df_processed)

                # 顯示總庫存與總成本
                st.write("### 庫存與成本總結")
                st.dataframe(summary_df)
                cd / Users / alyssawu / PycharmProjects / PythonProject2  # 進入你的專案資料夾
                git
                init  # 初始化 Git 儲存庫
                git
                add.  # 添加所有檔案
                git
                commit - m
                "Initial commit"  # 提交變更
                git
                branch - M
                main  # 設定 main 為主分支
                git
                remote
                add
                origin
                https: // github.com / 你的GitHub帳號 / inventory - analysis.git  # 替換為你的 GitHub 儲存庫網址
                git
                push - u
                origin
                main  # 上傳程式碼

                # 匯出 Excel
                with pd.ExcelWriter("processed_inventory.xlsx", engine='xlsxwriter') as writer:
                    df_processed.to_excel(writer, sheet_name='整理後的資料', index=False)
                    summary_df.to_excel(writer, sheet_name='總庫存與總成本', index=False)

                with open("processed_inventory.xlsx", "rb") as file:
                    st.download_button("下載處理後的 Excel", file, file_name="庫存分析.xlsx")
            except Exception as e:
                st.error(f"發生錯誤：{e}")



