import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options 
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
from datetime import datetime
import os # 引入 os 用於環境變數
import json # 引入 json 用於憑證
import pytz # 引入 pytz用來設定時間

# === 設定 ChromeDriver 選項 ===
chrome_options = Options()
chrome_options.add_argument("--headless")  # 啟用無頭模式
chrome_options.add_argument("--no-sandbox") # 在 Docker 環境中需要
chrome_options.add_argument("--disable-dev-shm-usage") # 避免 /dev/shm 空間不足
chrome_options.add_argument("--window-size=1920,1080") # 設定視窗大小，模擬正常瀏覽器

# 小白加入
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--disable-extensions')
chrome_options.add_argument('--disable-plugins')
chrome_options.add_argument('--disable-images') 

try:
    # 使用 webdriver-manager 自動下載匹配的 ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print("WebDriver 初始化成功")
except Exception as e:
    print(f"WebDriver 初始化失敗: {e}")
    # 如果 webdriver-manager 失敗，嘗試使用系統中的 ChromeDriver
    try:
        driver = webdriver.Chrome(options=chrome_options)
        print("使用系統 ChromeDriver 成功")
    except Exception as e2:
        print(f"所有 WebDriver 初始化方法都失敗: {e2}")
        raise e2
# -------

# GitHub Actions 會自動安裝 ChromeDriver，所以不需要指定路徑
# 在 Actions 上直接用 WebDriver.Chrome(options=chrome_options) 即可
driver = webdriver.Chrome(options=chrome_options)

# === 登入後台 ===
driver.get("https://taichen.ibiza.com.tw/users/sign_in")
time.sleep(1)

# 從環境變數獲取登入資訊
user_email = os.environ.get("IBIZA_EMAIL")
user_password = os.environ.get("IBIZA_PASSWORD")

if not user_email or not user_password:
    print("錯誤：IBIZA_EMAIL 或 IBIZA_PASSWORD 環境變數未設定。")
    driver.quit()
    exit()

driver.find_element(By.ID, "user_email").send_keys(user_email)
driver.find_element(By.ID, "user_password").send_keys(user_password)
driver.find_element(By.CLASS_NAME, "btn-primary").click()
time.sleep(3)

print("✅ 成功登入！")

# === 點擊進入即時庫存頁面 ===
inventory_button = driver.find_element(By.XPATH, "//li[contains(@class, 'px-nav-dropdown')]/a[span[text()='庫存']]")
inventory_button.click()
time.sleep(1)
realtime_stock = driver.find_element(By.XPATH, "//a[span[text()='即時庫存']]")
realtime_stock.click()
time.sleep(1)
print("✅ 成功進入庫存頁面！")

# === 設定顯示 100 筆資料 ===
# select_element = driver.find_element(By.NAME, "stock-table_length")
# select = Select(select_element)
# select.select_by_value("100")
select_element = driver.find_element(By.CSS_SELECTOR, "select.form-control")
select = Select(select_element)

from selenium.common.exceptions import NoSuchElementException, TimeoutException

try:
    select.select_by_value("100")
    print("✅ 成功選到 value=100")
except NoSuchElementException:
    print("❌ 找不到 value=100 的選項，請檢查 HTML")
    
time.sleep(1)

# === 抓取庫存資料 ===
inventory_data = []

# rows = driver.find_elements(By.CSS_SELECTOR, "#stock-table tbody tr")
rows = driver.find_elements(By.CSS_SELECTOR, ".rdt_TableBody .rdt_TableRow")

for row in rows:
    # columns = row.find_elements(By.TAG_NAME, "td")
    columns = row.find_elements(By.CSS_SELECTOR, ".rdt_TableCell")
    
    if len(columns) > 6:
        # product_name = columns[0].text.strip()
        # quality_status = columns[5].text.strip()
        # available_stock = columns[6].text.strip()
        product_name = columns[1].text.strip()
        quality_status = columns[7].text.strip()
        available_stock = columns[10].text.strip()
        product_code = product_name.split(" ")[0].strip()
        # === 處理特例：商品編號修正 ===
        if product_code.startswith("DDA00000001"):
            suffix = product_code[13:]  # 取 -2 類尾碼
            product_code = "DDA0000001" + (f"-{suffix}" if suffix else "")
        # === 防盜貼紙：不篩選良品，強制統一命名 ===
        if "防盜貼紙" in product_name or "RP-ANS1" in product_name:
            product_code = "RP-ANS1 R膠防盜貼紙"
            # print(f"[防盜貼紙] {product_name}, 狀態: {quality_status}, 庫存: {available_stock}")
        elif "良品" not in quality_status:
            continue
        inventory_data.append({
            "商品編號": product_code,
            "商品名稱": product_name,
            "庫存": int(available_stock) if available_stock.isdigit() else 0
        })

df = pd.DataFrame(inventory_data)


# === 建立原始編號欄（取前 10 碼）===
df["原始編號"] = df["商品編號"].apply(lambda x: x[:10] if "防盜貼紙" not in x else x)

# === 自定義排序順序（含防盜貼紙作為最後一項）===
custom_order = [
    "ECA0000005", "ECA0000001", "ECA0000002", "ECA0000006", "ECA0000003", "ECA0000004", "ECA0000009",
    "DEA0000001", "DEA0000000", "DDA0000000", "DDA0000001", "DDB0000000", "EBA0000000", "EBB0000000", "DCA0000000",
    "DCB0000000", "DBA0000000", "DBB0000000", "FAA0000000", "FBA0000000", "FBB0000000", "RP-ANS1 R膠防盜貼紙"
]

# === 統計表：每個原始編號的總庫存 vs 不含 R 的庫存 ===
summary_df = pd.DataFrame()
summary_df["原始編號"] = df["原始編號"].unique()

# 現有總庫存量：所有前 10 碼相同的都算
summary_df["現有總庫存量"] = summary_df["原始編號"].apply(
    lambda code: df[df["原始編號"] == code]["庫存"].sum()
)

# 不包含 R 的數量：只統計 base 或 -數字，不統計 -R 類
def is_not_R_variant(row):
    code = row["商品編號"]
    base = row["原始編號"]
    if base == "RP-ANS1 R膠防盜貼紙":  # 防盜貼紙全算為一種，不列入 R 判斷
        return True
    if code == base:
        return True
    elif re.match(rf"^{re.escape(base)}-\d+$", code):  # 如 ECA0000002-01
        return True
    else:
        return False

df["是否計入不含R"] = df.apply(is_not_R_variant, axis=1)

summary_df["不包含R的數量"] = summary_df["原始編號"].apply(
    lambda code: df[(df["原始編號"] == code) & (df["是否計入不含R"])]["庫存"].sum()
)

# === 只保留在 custom_order 中的原始編號，並依順序排序 ===
summary_df = summary_df[summary_df["原始編號"].isin(custom_order)]
summary_df["排序"] = summary_df["原始編號"].apply(lambda x: custom_order.index(x))
summary_df = summary_df.sort_values(by="排序").drop(columns=["排序"]).reset_index(drop=True)

# === 自動化 Google Sheets ===
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 從環境變數獲取 Google 憑證
try:
    google_credentials_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not google_credentials_json:
        raise ValueError("GOOGLE_CREDENTIALS 環境變數未設定。")
    
    # 將 JSON 字串轉換為 Python 字典
    google_creds_dict = json.loads(google_credentials_json)

    # 授權
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(google_creds_dict, scope)
    client = gspread.authorize(creds)

except Exception as e:
    print(f"與 Google Sheets 進行身份驗證時發生錯誤: {e}")
    driver.quit()
    exit()

# === 使用 Google Sheet ID 開啟 ===
sheet_id = "1U6F3hvj76YGOSWodEkfZmkSS31F1QmNndI16igBkgGE"  # Sheet ID
# sheet = client.open_by_key(sheet_id).worksheet("3. 庫存管理表（自動）")
sheet = client.open_by_key(sheet_id).worksheet("3. 庫存管理表（自動）表頭名稱請勿更動!!")

# === 更新每筆庫存資料到 Google Sheet (G欄 與 Q欄) ===
updates = []
for idx, row in summary_df.iterrows():
    q_value = row['不包含R的數量']
    g_value = row['現有總庫存量']
    
    # debug用：爬蟲更新的值
    #print(f"DEBUG_DATA: 準備更新 第 {idx + 2} 行, 第 7 欄 (Q) 資料為: {q_value}")
    #print(f"DEBUG_DATA: 準備更新 第 {idx + 2} 行, 第 17 欄 (G) 資料為: {g_value}")

    updates.append(gspread.Cell(row=idx + 2, col=7, value=q_value))
    updates.append(gspread.Cell(row=idx + 2, col=17, value=g_value))

if updates:
    try:
        sheet.update_cells(updates)
        print(f"DEBUG_DATA: 成功向 Google Sheet API 發送了 {len(updates)} 個單元格更新請求。")
    except Exception as update_error:
        print(f"ERROR_DATA_UPDATE: 更新 Google Sheet 資料單元格失敗: {update_error}")
        exit(1)

# === 新增最後更新時間 ===
# 定義台灣時區
taiwan_timezone = pytz.timezone('Asia/Taipei')
# 獲取當前台灣時間
current_time_taiwan = datetime.now(taiwan_timezone).strftime("%Y-%m-%d %H:%M:%S")

last_row_calculated = len(summary_df) + 2
sheet.update_cell(row=last_row_calculated, col=7, value=f"最後更新時間：{current_time_taiwan}")

print("✅ 已成功同步至 Google Sheet！")

# === 關閉瀏覽器 ===
driver.quit()






