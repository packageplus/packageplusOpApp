import pandas as pd
import msoffcrypto # 注意：這個庫在 PyPI 上是 msoffcrypto-tool
import io
import os
from datetime import datetime
import streamlit as st # 引入 Streamlit

# 您的密碼，您可以考慮將其作為 Streamlit 的 secret 變數
# 但為了簡潔，這裡暫時直接寫入
PASSWORD = "533793"

# === 解密 Excel 檔案（支援加密 .xlsx） ===
# 這個函式現在接收 BytesIO 物件，並返回解密後的 BytesIO
def decrypt_excel_streamlit(uploaded_file, password):
    decrypted = io.BytesIO()
    try:
        # msoffcrypto.OfficeFile 可以直接從 BytesIO 對象讀取
        office_file = msoffcrypto.OfficeFile(uploaded_file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted
    except Exception as e:
        st.error(f"❌ 解密 Excel 檔案失敗，請檢查密碼或檔案是否損壞：\n{str(e)}")
        return None

# === 拆解 AJ 欄內容 ===
# 這個函式保持不變，因為它是純數據處理邏輯
def split_aj(value):
    if pd.isna(value):
        return pd.NA, pd.NA
    parts = str(value).split(',')
    if len(parts) == 2:
        return parts[0].strip(), "轉接碼：" + parts[1].strip()
    return value, pd.NA

# === 主轉檔流程函式 ===
# 這個函式現在接收 Streamlit 的 uploaded_file 物件，並返回處理後的 DataFrame
def process_shopee_excel_logic(uploaded_file):
    try:
        st.info("🔄 正在解密 Excel 檔案...")
        decrypted_file_stream = decrypt_excel_streamlit(uploaded_file, PASSWORD)
        if decrypted_file_stream is None:
            return None # 解密失敗，直接退出

        st.info("🔄 正在讀取解密後的 Excel 檔案...")
        df = pd.read_excel(decrypted_file_stream, sheet_name=0, engine="openpyxl")
        st.success("✔ Excel 檔案讀取成功！")

        # 固定欄位名稱
        aj_col = "蝦皮專線和包裹查詢碼 \n(請複製下方完整編號提供給您配合的物流商當做聯絡電話)"
        ai_col = "收件者電話\n(若您是自行配送請使用後方蝦皮專線和包裹查詢碼聯繫買家)"
        az_col = "備註"

        st.info("🔄 正在拆解蝦皮專線欄位...")
        # 拆解 AJ → 寫入 AI 與 AZ
        # 使用 .copy() 避免 SettingWithCopyWarning
        temp_df = df[aj_col].apply(lambda x: pd.Series(split_aj(x))).copy()
        df[ai_col] = temp_df[0]
        df[az_col] = temp_df[1]
        st.success("✔ 蝦皮專線欄位拆解完成！")


        st.info("🔄 正在分攤賣場優惠券金額...")
        # 分攤金額（O欄），根據訂單號（A欄）
        # 這裡假設 "訂單編號" 是正確的列名
        if '訂單編號' in df.columns:
            df['賣場優惠券'] = df.groupby('訂單編號')['賣場優惠券'].transform(lambda x: x / len(x))
            st.success("✔ 賣場優惠券金額分攤完成！")
        else:
            st.warning("⚠️ 未找到 '訂單編號' 欄位，跳過賣場優惠券分攤。")


        st.success("✅ 所有處理步驟完成！")
        return df # 返回處理後的 DataFrame

    except Exception as e:
        st.error(f"❌ 處理失敗：\n{str(e)}")
        return None

# === Streamlit 介面函式 ===
def shopee_excel_app():
    st.header("🦐 蝦皮訂單 Excel 處理工具")
    st.markdown("這個工具可以解密蝦皮訂單 Excel，拆解聯絡資訊並分攤優惠券金額。")

    # Streamlit 的檔案上傳器
    uploaded_file = st.file_uploader("請上傳加密的蝦皮訂單 Excel 檔案 (.xlsx 或 .xlsm)", type=["xlsx", "xlsm"])

    if uploaded_file is not None:
        if st.button("🚀 開始處理"):
            with st.spinner("檔案正在處理中，請稍候..."):
                processed_df = process_shopee_excel_logic(uploaded_file)

            if processed_df is not None:
                st.success("✅ 檔案處理完成！您可以下載結果檔案。")

                today_str = datetime.now().strftime('%Y%m%d')
                output_filename = f"{today_str}蝦皮.xlsx"

                # 將處理後的 DataFrame 保存到 BytesIO，以便 Streamlit 提供下載
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    processed_df.to_excel(writer, index=False, sheet_name='orders')
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
    st.markdown("如有任何問題，請聯繫 IT 部門。")