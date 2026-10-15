import os
import pandas as pd

# 設定資料夾路徑
SOURCE_FOLDER = r"C:\Users\User\Desktop"
OUTPUT_FILE = "matched_results.xlsx"

# 搜尋的關鍵字
TARGET_KEYWORDS = ["google chrome", "microsoft edge"]

def search_files_and_extract(source_dir):
    all_matches = []

    # 遞迴走訪資料夾中的所有子資料夾與檔案
    for root, _, files in os.walk(source_dir):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            file_path = os.path.join(root, file)

            # 只處理 Excel 與 CSV 檔案，並跳過 Excel 的暫存檔 (以 ~$ 開頭)
            if ext in ['.csv', '.xlsx', '.xls'] and not file.startswith('~$'):
                try:
                    # 讀取檔案
                    if ext == '.csv':
                        # 嘗試讀取 CSV（優先 utf-8，失敗則嘗試 utf-8-sig 或 cp950/big5）
                        try:
                            df = pd.read_csv(file_path, dtype=str)
                        except UnicodeDecodeError:
                            df = pd.read_csv(file_path, dtype=str, encoding='cp950')
                    else:
                        # 讀取 Excel 檔案的所有工作表 (Sheet)
                        excel_file = pd.ExcelFile(file_path)
                        sheet_dfs = []
                        for sheet_name in excel_file.sheet_names:
                            temp_df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
                            temp_df['來源工作表'] = sheet_name
                            sheet_dfs.append(temp_df)
                        df = pd.concat(sheet_dfs, ignore_index=True) if sheet_dfs else pd.DataFrame()

                    if df.empty:
                        continue

                    # 將整行的內容轉為字串並轉小寫，檢查是否包含任一關鍵字
                    row_as_string = df.astype(str).apply(lambda row: ' '.join(row).lower(), axis=1)
                    
                    pattern = '|'.join(TARGET_KEYWORDS)
                    mask = row_as_string.str.contains(pattern, na=False)

                    # 篩選出符合條件的整行資料
                    matched_df = df[mask].copy()

                    if not matched_df.empty:
                        # 在第一欄插入「檔案路徑」與「檔案名稱」資訊
                        matched_df.insert(0, '檔案路徑', file_path)
                        matched_df.insert(1, '檔案名稱', file)
                        all_matches.append(matched_df)

                        print(f"[找到符合項目] {file_path} ({len(matched_df)} 筆)")

                except Exception as e:
                    print(f"[讀取失敗] {file_path}: {e}")

    # 合併所有搜尋結果並匯出
    if all_matches:
        final_df = pd.concat(all_matches, ignore_index=True)
        
        if OUTPUT_FILE.endswith('.xlsx'):
            final_df.to_excel(OUTPUT_FILE, index=False)
        else:
            final_df.to_csv(OUTPUT_FILE, index=False, encoding='utf-8-sig')
            
        print(f"\n搜尋完成！結果已成功寫入：{OUTPUT_FILE}")
    else:
        print("\n未找到包含指定關鍵字的資料。")

if __name__ == "__main__":
    search_files_and_extract(SOURCE_FOLDER)