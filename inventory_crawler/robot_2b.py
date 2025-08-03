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

# === 點擊「庫存」按鈕展開選單 === 
inventory_button = driver.find_element(By.XPATH, "//li[contains(@class, 'px-nav-dropdown')]/a[span[text()='庫存']]")
inventory_button.click()
time.sleep(1)  # 等待選單展開

# === 點擊「即時庫存」 === 
realtime_stock = driver.find_element(By.XPATH, "//a[span[text()='即時庫存']]")
realtime_stock.click()
time.sleep(1)  # 等待頁面加載

print("✅ 成功進入庫存頁面！")

#  === 選擇顯示 100 項結果 === 
#select_element = driver.find_element(By.NAME, "stock-table_length")
#select = Select(select_element)
#select.select_by_value("100")  # 設定為 100 項

select = Select(driver.find_element(By.XPATH, "//div[contains(@class, 'dataTable-container')]//select"))
select.select_by_visible_text("100")

time.sleep(1)  # 等待頁面更新

#  === 定義允許「移倉不盤點」的商品 === 
allow_extra_status = {"RP-SIZESS", "RP-SIZESM", "RP-SIZESL", "RP-SIZESXL", "rp-strap200cm"}

#  === 找到商品行 === 
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

wait = WebDriverWait(driver, 10)  # 最多等10秒

inventory_data = []

while True:
    rows = driver.find_elements(By.CSS_SELECTOR, "#stock-table tbody tr")

    # print 頁碼和第一列商品，方便 debug
    try:
        current_page = driver.find_element(By.CSS_SELECTOR, "li.paginate_button.active").text
    except Exception:
        current_page = "未知"
    if rows:
        print(f"頁碼: {current_page}, 第一列: {rows[0].text}")
    else:
        print(f"頁碼: {current_page}, 沒有資料")

    for row in rows:
        columns = row.find_elements(By.TAG_NAME, "td")
        if len(columns) > 6:
            product_name = columns[0].text.strip()
            quality_status = columns[5].text.strip()
            available_stock = columns[6].text.strip()
            product_code = product_name.split(" ")[0].strip()

            if product_code in allow_extra_status:
                if "良品" not in quality_status and "移倉不盤點" not in quality_status:
                    continue
            else:
                if "良品" not in quality_status:
                    continue

            inventory_data.append({
                "商品編號": product_code,
                "商品名稱": product_name,
                "庫存": available_stock
            })

    # 換頁準備：記下這一頁第一列
    if rows:
        first_row_text = rows[0].text
    else:
        first_row_text = None

    # 尋找「可點」的下一頁 <a>
    try:
        next_a = driver.find_element(By.CSS_SELECTOR, "li.paginate_button.next:not(.disabled) > a")
    except NoSuchElementException:
        break

    # 點擊「下一頁」
    next_a.click()

    # 等到表格內容真的換新
    try:
        wait.until(
            lambda d: d.find_elements(By.CSS_SELECTOR, "#stock-table tbody tr")
                      and d.find_elements(By.CSS_SELECTOR, "#stock-table tbody tr")[0].text != first_row_text
        )
    except TimeoutException:
        print("⚠️ 頁面未成功刷新，停止爬取")
        break


df = pd.DataFrame(inventory_data)

#  === 合併相同商品編號的庫存數 === 
df = df.groupby(["商品編號"], as_index=False).agg({
    "商品名稱": "first",  # 保留第一個名稱
    "庫存": "sum"  # 合併庫存數量
})

#  === 自定義排序規則 === 
custom_order = [
    "EBEA0000000", "EBFA0000000", "EBGA0000000", "EBHA0000000", "EBJA0000000", "EBMA0000000",
    "EBCA0000000", "EBDA0000000", "EBA0000000", "EBB0000000",
    "TSMCAA", "TSMCRA",
    "RP-COLLECTBA", "RP-COLLECTBB", "RP-COLLECTBC",
    "RP-SIZESS", "RP-SIZESM", "RP-SIZESL", "RP-SIZESXL",
    "rp-strap200cm"
]

#  === 確保所有商品都顯示，即使沒抓到 === 
all_items = pd.DataFrame({"商品編號": custom_order})
df = pd.merge(all_items, df, on="商品編號", how="left")

#  === 確保商品編號正確 === 
df["排序"] = df["商品編號"].apply(lambda x: custom_order.index(x) if x in custom_order else len(custom_order))  

#  === 依照自訂順序排序 === 
df = df.sort_values(by="排序").drop(columns=["排序"])

# === 寫入 Google Sheet：Ｆ欄位（第 6 欄） ===
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


sheet_id = "1qB5xe4inx4spFXPxOdSytL_BUElR6Bd0aVZ2TTqrAuY"
sheet = client.open_by_key(sheet_id).worksheet("庫存管理表")

# 寫入 F 欄（從第 2 列開始）
for idx, row in df.iterrows():
    quantity = row["庫存"]
    # 把 NaN 或缺值變成空字串或 0（視需求）
    if pd.isna(quantity):
        quantity = ""
    sheet.update_cell(idx + 2, 6, quantity)

# === 新增最後更新時間 ===
# 定義台灣時區
taiwan_timezone = pytz.timezone('Asia/Taipei')

# 獲取當前台灣時間
current_time_taiwan = datetime.now(taiwan_timezone).strftime("%Y-%m-%d %H:%M:%S")

# 最後一列寫入更新時間（E欄），基於 df 的長度計算行數
last_row = len(df) + 2 
sheet.update_cell(row=last_row, col=6, value=f"最後更新時間：{current_time_taiwan}") 
print("✅ 數據已寫入 F 欄並更新時間！")

# **關閉瀏覽器**
driver.quit()

