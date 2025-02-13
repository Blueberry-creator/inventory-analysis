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

        # 轉換數值欄位，確保所有數字欄位都是 float
        numeric_columns = ['商品原價', '商品成本', '庫存總量', '總成本']
        for col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')  # 轉換成數字，無法轉換的變 NaN

        # 檢查是否所有數值欄位都轉換成功
        if df[numeric_columns].isna().all().all():
            raise ValueError("數值欄位格式錯誤，請確認 Excel 檔案內容。")

        # 過濾品牌關鍵字
        df = df[df['商品名稱'].astype(str).str.contains(keyword, na=False, case=False)]

        # 只保留指定欄位
        df = df[required_columns]

        # 移除庫存為 0 或負數的商品
        df = df[df['庫存總量'] > 0]

        # Highlight 商品成本為 NaN 的商品
        df['標記'] = np.where(df['商品成本'].isna(), '缺少成本', '')

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
