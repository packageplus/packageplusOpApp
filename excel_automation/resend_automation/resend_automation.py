import os
import pandas as pd
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import xlwings as xw

def read_excel_with_xlwings(file_path):
    """ 🚀 使用 xlwings 讀取 Excel 內部數據（完全不受格式影響） """
    try:
        print(f"🔄 使用 xlwings 讀取 Excel：{file_path}")

        # 開啟 Excel 檔案
        app = xw.App(visible=False)  # 隱藏 Excel
        wb = app.books.open(file_path)
        sheet = wb.sheets[0]  # 只取第一個工作表

        # 讀取純數據
        data = sheet.used_range.value  # 取得資料（不帶格式）
        wb.close()
        app.quit()

        # 轉換成 DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])  # 第一列作為標題
        print("✔ Excel 讀取成功！")
        return df

    except Exception as e:
        print(f"❌ 無法讀取 Excel：{e}")
        return None

def select_files():
    """ 讓使用者手動選擇 3 個 Excel 檔案，並顯示明確標示 """
    file_paths = []
    instructions = [
        "請選擇【訂單匯出.xlsx】（VLOOKUP 來源）",
        "請選擇【禾洛出貨通知.xlsx】（主檔案）",
        "請選擇【範例檔的 Excel 檔案】"
    ]
    
    for i in range(3):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], title=instructions[i])
        if not file_path:
            messagebox.showwarning("警告", f"您未選擇 {instructions[i]}，程式即將終止。")
            return None
        file_paths.append(file_path)
    
    print("✔ 已選擇檔案：", file_paths)  # Debug
    return file_paths

def process_excel(file1, file2, file3):
    """ 讀取 Excel，處理數據，並儲存結果 """
    try:
        print("\n=== 1️⃣ 使用 xlwings 讀取 Excel ===")

        # **使用 xlwings 讀取 Excel**
        df1 = pd.read_excel(file1, engine="openpyxl", dtype=str)  # 訂單匯出
        df2 = read_excel_with_xlwings(file2)  # 禾洛出貨通知（使用 xlwings 避免格式錯誤）

        if df2 is None:
            raise Exception("禾洛出貨通知讀取失敗！")

        print("\n=== 2️⃣ 處理數據 ===")

        # **步驟 1：刪除 禾洛出貨通知.xlsx（主檔案）品名包含特定關鍵字的行**
        keywords = ["紙箱", "防盜貼紙", "第三代"]
        if "品名" in df2.columns:
            df2 = df2[~df2["品名"].str.contains("|".join(keywords), na=False)]
            print("✔ 已刪除指定品名的行")

        # **步驟 2：將 訂單編號* 欄去掉 `#` 並新增新欄位**
        if "訂單編號*" in df2.columns:
            df2.insert(0, "訂單編號_純數字", df2["訂單編號*"].str.replace("#", "", regex=False))
            print("✔ 已新增純數字訂單編號欄位")

        # **步驟 3：確保 訂單編號 欄位名稱一致**
        df1 = df1.rename(columns={"訂單編號": "訂單編號*"})  # 訂單匯出的訂單欄位

        # **步驟 4：確保 VLOOKUP 不會發生索引重複**
        if df1.duplicated(subset=["訂單編號*"], keep="first").sum() > 0:
            print("⚠ 發現重複訂單，正在處理...")
            df1 = df1.drop_duplicates(subset=["訂單編號*"], keep="first")  # **只保留第一筆**
            print("✔ 已刪除重複訂單，保留第一筆")

        # **步驟 5：建立 VLOOKUP 查詢字典**
        lookup_dict = df1.set_index("訂單編號*")[["出倉日", "物流追蹤碼", "物流類型"]].to_dict("index")

        # **步驟 6：將 VLOOKUP 結果貼入 禾洛出貨通知**
        df2["出貨日期"] = df2["訂單編號_純數字"].map(lambda x: lookup_dict.get(x, {}).get("出倉日", ""))
        df2["宅配編號"] = df2["訂單編號_純數字"].map(lambda x: lookup_dict.get(x, {}).get("物流追蹤碼", ""))
        df2["貨運公司"] = df2["訂單編號_純數字"].map(lambda x: lookup_dict.get(x, {}).get("物流類型", ""))
        print("✔ VLOOKUP 匹配完成")

        # **步驟 7：獨立剪下貨運公司為空的資料**
        df_missing = df2[df2["貨運公司"].isna() | (df2["貨運公司"] == "")].copy()  # 建立完全獨立副本
        df2 = df2.loc[~df2.index.isin(df_missing.index)].copy()  # 只刪除異常資料，確保 df2 其他欄位不受影響
        print(f"✔ 已剪下 {len(df_missing)} 筆異常資料")

        # **步驟 8：格式轉換**
        df_missing.loc[:, "訂單編號*"] = df_missing["訂單編號*"].str.replace("#", "", regex=False)

        # **僅對 df_missing 內的 '物流類型' 進行替換**
        df_missing.loc[:, "物流類型"] = df_missing["物流類型"].replace({
         "【7-11】取貨不付款": "無串接-Shopline 7-11已付",
         "【全家】取貨不付款": "無串接-Shopline全家已付"
        })

        df_missing.loc[:, "倉別"] = "A1"
        df_missing.loc[:, "數量*"] = pd.to_numeric(df_missing["數量*"], errors='coerce').fillna(0).astype(int)
        df_missing.loc[:, "預約出貨日"] = pd.to_datetime(df_missing["預約出貨日"], errors='coerce').dt.strftime("%Y/%m/%d")

        # **清空指定欄位**
        clear_columns = ["城市", "國家", "識別碼", "幣別"]
        df_missing.loc[:, clear_columns] = ""

        # **確保 '訂單編號_純數字' 不會被貼上**
        df_missing = df_missing.drop(columns=["訂單編號_純數字"], errors="ignore")
        
        # **步驟 9：確保 df2 貨運公司欄位未被影響**
        df2["貨運公司"] = df2["物流類型"]

        #步驟10 優化清除三欄
        columns_to_clear = ["出貨號碼", "訂單號碼", "行號碼"]
        for col in columns_to_clear:
         if col in df_missing.columns:
             df_missing[col] = ""
        print("✔ 已清除 禾洛回傳 中的出貨號碼、訂單號碼、行號碼")
        

        # **儲存 Excel**
        today_date = datetime.datetime.now().strftime("%m%d")
        folder_path = os.path.dirname(file2)

        output_path1 = os.path.join(folder_path, f"{today_date}_ERP回傳.xlsx")
        df2.to_excel(output_path1, index=False, engine="openpyxl", sheet_name="資料")
        print(f"✔ ERP回傳已存儲：{output_path1}")

        output_path2 = os.path.join(folder_path, f"{today_date}_禾洛回傳.xlsx")
        df_missing.to_excel(output_path2, index=False, engine="openpyxl")
        print(f"✔ 禾洛回傳已存儲：{output_path2}")

        messagebox.showinfo("完成", f"處理完成！\n\n已儲存：\n{output_path1}\n{output_path2}")

    except Exception as e:
        print("❌ 錯誤：", str(e))
        messagebox.showerror("錯誤", f"處理 Excel 時發生錯誤：\n{str(e)}")

root = tk.Tk()
root.withdraw()
file_paths = select_files()
if file_paths:
    process_excel(file_paths[0], file_paths[1], file_paths[2])

