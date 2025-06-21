# tools/momo_processor.py

import pandas as pd
import msoffcrypto # 注意：這個庫在 PyPI 上是 msoffcrypto-tool
import io
import os # 雖然大部分檔案操作在雲端會改變，但 os.path.join 等還是常用
from datetime import datetime
import streamlit as st # 引入 Streamlit

# 移除 Tkinter 和 subprocess 相關的 import
# import tkinter as tk
# from tkinter import filedialog, messagebox
# import subprocess

# 您的密碼，請根據實際情況替換
# 考慮將其作為 Streamlit 的 secret 變數，以增強安全性
# https://docs.streamlit.io/deploy/streamlit-community-cloud/secrets-management
PASSWORD = "50916648"

# === 解密 Excel 檔案（支援加密 .xlsx） ===
# 這個函式現在接收 Streamlit 的 uploaded_file 物件，並返回解密後的 BytesIO
def decrypt_excel_streamlit(uploaded_file_stream, password):
    decrypted = io.BytesIO()
    try:
        # msoffcrypto.OfficeFile 可以直接從 BytesIO 對象讀取
        office_file = msoffcrypto.OfficeFile(uploaded_file_stream)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted
    except Exception as e:
        st.error(f"❌ 解密 Excel 檔案失敗，請檢查密碼是否正確或檔案是否損壞：\n{str(e)}")
        return None

# === 主轉檔流程函式 ===
# 這個函式現在接收 Streamlit 的 uploaded_file 物件，並返回處理後的 DataFrame
def process_momo_excel_logic(uploaded_file):
    try:
        st.info("🔄 正在解密 Excel 檔案...")
        decrypted_file_stream = decrypt_excel_streamlit(uploaded_file, PASSWORD)
        if decrypted_file_stream is None:
            return None # 解密失敗，直接退出

        st.info("🔄 正在讀取解密後的 Excel 檔案...")
        # sheet_name=0 表示讀取第一個工作表
        df = pd.read_excel(decrypted_file_stream, sheet_name=0, engine="openpyxl")
        st.success("✔ Excel 檔案讀取成功！")

        # 固定欄位名稱（請再次確認這些是您的 Excel T 欄和 W 欄的實際標題）
        col_product_name = '商品名稱' # 請根據您的 Excel T 欄實際標題修改
        col_order_amount = '訂單金額' # 請根據您的 Excel W 欄實際標題修改

        # 檢查關鍵欄位是否存在
        if col_product_name not in df.columns:
            st.error(f"Excel 中缺少必要欄位：'{col_product_name}'。請檢查欄位名稱是否正確。")
            return None
        if col_order_amount not in df.columns:
            st.error(f"Excel 中缺少必要欄位：'{col_order_amount}'。請檢查欄位名稱是否正確。")
            return None

        # 篩選出符合條件的列 (商品名稱為「運費」且訂單金額為 0)
        st.info("🔄 正在篩選並刪除 '商品名稱為運費且訂單金額為0' 的資料...")
        initial_rows = len(df)
        df = df[~((df[col_product_name] == '運費') & (df[col_order_amount] == 0))]
        deleted_rows = initial_rows - len(df)
        st.success(f"✔ 已刪除 {deleted_rows} 列符合條件的資料。")

        st.success("✅ 所有處理步驟完成！")
        return df # 返回處理後的 DataFrame

    except Exception as e:
        st.error(f"❌ 處理失敗：\n{str(e)}")
        return None

# === Streamlit 介面函式 ===
def momo_excel_app():
    st.header("🛍️ Momo 訂單 Excel 處理工具")
    st.markdown("這個工具可以解密 Momo 訂單 Excel，並刪除商品名稱為「運費」且訂單金額為 0 的資料。")

    # Streamlit 的檔案上傳器
    uploaded_file = st.file_uploader("請上傳加密的 Momo 訂單 Excel 檔案 (.xlsx 或 .xlsm)", type=["xlsx", "xlsm"])

    if uploaded_file is not None:
        if st.button("🚀 開始處理"):
            with st.spinner("檔案正在處理中，請稍候..."):
                processed_df = process_momo_excel_logic(uploaded_file)

            if processed_df is not None:
                st.success("✅ 檔案處理完成！您可以下載結果檔案。")

                today_str = datetime.now().strftime('%Y%m%d')
                output_filename = f"{today_str}momo.xlsx"

                # 將處理後的 DataFrame 保存到 BytesIO，以便 Streamlit 提供下載
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    processed_df.to_excel(writer, index=False, sheet_name='工作表1')
                output_buffer.seek(0)

                st.download_button(
                    label=f"💾 下載 {output_filename}",
                    data=output_buffer,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("❗ 處理失敗，請檢查錯誤訊息。")
    st.markdown("---")
    st.markdown("如有任何問題，可能沒有人可以修XD(再看看)")