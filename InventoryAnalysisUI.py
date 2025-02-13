import streamlit as st
import pandas as pd
import numpy as np


def process_inventory(file, keyword):
    try:
        # 嘗試讀取 Excel 檔案
        excel_data = pd.ExcelFile(file)

        # 檢查是否有任何工作表
        if len(excel_data.sheet_names) == 0:
            raise ValueError("Excel 檔案沒有任何工作表，請確認內容。")

        # 顯示 Excel 的工作表名稱（幫助 Debug）
        st.write(f"📂 Excel 檔案包含的工作表: {excel_data.sheet_names}")

        # 讀取第一個工作表
        df = pd.read_excel(file, sheet_name=0)

        # 確保 Excel 內容不為空
        if df.empty:
            raise ValueError("Excel 檔案沒有可讀取的數據，請確認檔案內容。")

        # 確保包含所需欄位
        required_columns = ['商品名稱', '商品款式', '商品原價', '商品成本', '庫存總量', '總成本']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Excel 檔案缺少以下欄位: {missing_columns}，請確認格式。")

        # 過濾品牌關鍵字
        df = df[df['商品名稱'].astype(str).str.contains(keyword, na=False, case=False)]

        # 只保留指定欄位
        df = df[required_columns]

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

    except Exception as e:
        raise ValueError(f"讀取 Excel 時發生錯誤: {e}")


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

                # 匯出 Excel
                with pd.ExcelWriter("processed_inventory.xlsx", engine='xlsxwriter') as writer:
                    df_processed.to_excel(writer, sheet_name='整理後的資料', index=False)
                    summary_df.to_excel(writer, sheet_name='總庫存與總成本', index=False)

                with open("processed_inventory.xlsx", "rb") as file:
                    st.download_button("下載處理後的 Excel", file, file_name="InventoryAnalysisUI.xlsx")
            except Exception as e:
                st.error(f"發生錯誤：{e}")




