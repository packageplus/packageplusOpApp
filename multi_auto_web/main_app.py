# main_app.py

import streamlit as st
import pandas as pd # 這些基礎模組通常可以放在 main_app 頂部，因為 Streamlit 元件可能用到
import datetime
import openpyxl
import re
from io import BytesIO

# 從 tools 資料夾引入各個工具的入口函式
from tools.shopline_processor import shopline_excel_app
from tools.shopee_processor import shopee_excel_app
from tools.momo_processor import momo_excel_app # <-- 新增這一行！

st.set_page_config(layout="wide", page_title="綜合 Excel 自動化工具")
st.sidebar.title("🛠️ 工具選單")

selected_tool = st.sidebar.radio(
    "請選擇您要使用的工具：",
    ("Shopline 訂單處理", "蝦皮訂單處理", "Momo 訂單處理") # <-- 更新選單選項！
)

st.title("💡 綜合 Excel 自動化處理平台")
st.markdown("歡迎使用我們的綜合工具！請從左側選單選擇您需要的功能。")
st.markdown("---")

if selected_tool == "Shopline 訂單處理":
    shopline_excel_app()
elif selected_tool == "蝦皮訂單處理":
    shopee_excel_app()
elif selected_tool == "Momo 訂單處理": # <-- 新增這個條件
    momo_excel_app()
# 您可以根據需要繼續添加更多 elif 條件來支持其他工具

st.markdown("---")
st.markdown("由營運部Intern製作～祝大家用的開心")