import pandas as pd
import msoffcrypto
import io
from tkinter import Tk, filedialog
from datetime import datetime
import os

PASSWORD = "533793"

# === 解密 Excel 檔案（支援加密 .xlsx） ===
def decrypt_excel(file_path, password):
    decrypted = io.BytesIO()
    with open(file_path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted

# === 選擇檔案 ===
def select_file():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title='選擇加密的 Excel 檔案',
        filetypes=[('Excel Files', '*.xlsx *.xlsm')]
    )

# === 拆解 AJ 欄內容 ===
def split_aj(value):
    if pd.isna(value):
        return pd.NA, pd.NA
    parts = str(value).split(',')
    if len(parts) == 2:
        return parts[0].strip(), "轉接碼：" + parts[1].strip()
    return value, pd.NA

# === 主轉檔流程 ===
def process_file(file_path):
    decrypted_file = decrypt_excel(file_path, PASSWORD)
    df = pd.read_excel(decrypted_file, sheet_name=0, engine="openpyxl")

    # 固定欄位名稱（你提供的）
    aj_col = "蝦皮專線和包裹查詢碼 \n(請複製下方完整編號提供給您配合的物流商當做聯絡電話)"
    ai_col = "收件者電話\n(若您是自行配送請使用後方蝦皮專線和包裹查詢碼聯繫買家)"
    az_col = "備註"

    # 拆解 AJ → 寫入 AI 與 AZ
    df[[ai_col, az_col]] = df[aj_col].apply(lambda x: pd.Series(split_aj(x)))

    # 分攤金額（O欄），根據訂單號（A欄）
    df['賣場優惠券'] = df.groupby('訂單編號')['賣場優惠券'].transform(lambda x: x / len(x))

    # 儲存路徑：日期 + 蝦皮.xlsm
    today_str = datetime.now().strftime('%Y%m%d')
    output_path = os.path.join(os.path.dirname(file_path), f"{today_str}蝦皮.xlsm")
    df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"✅ 檔案轉檔完成，儲存為：{output_path}")

# === 程式進入點 ===
if __name__ == '__main__':
    file_path = select_file()
    if file_path:
        try:
            process_file(file_path)
        except Exception as e:
            print(f"❌ 錯誤：{e}")
    else:
        print("⚠️ 未選擇任何檔案。")
