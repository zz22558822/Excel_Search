# Excel Search (Excel 關鍵字搜尋工具)

基於 Python + PyQt6 開發的工具，用於在大量 Excel (`.xlsx`, `.xls`) 與 CSV (`.csv`) 檔案中搜尋特定關鍵字。

## ✨ 功能特色

- **支援多種格式**：可讀取並搜尋 `.xlsx`、`.xls` 以及 `.csv` 檔案。
- **多層資料搜尋**：自動掃描指定資料夾及其所有子資料夾下的檔案。
- **智慧編碼處理**：讀取 CSV 時會優先嘗試 UTF-8，若失敗則自動切換為 CP950/Big5，避免亂碼問題。
- **完整結果匯出**：將所有包含關鍵字的資料列整理合併，並在第一欄與第二欄自動加入「檔案路徑」與「檔案名稱」，最後匯出成單一 Excel 或 CSV 檔案，方便後續查閱。
- **雙版本支援**：
  - **CLI 版本**：輕量、快速，適合需要自動化排程或習慣終端機操作的開發者。
  - **GUI 版本**：使用 PyQt6 打造的現代化介面，提供直覺的操作體驗、進度條顯示及主題切換。

## 📦 安裝依賴套件

請確保您的環境已安裝 Python 3.8 或以上版本。接著，在終端機中執行以下指令來安裝必要的套件：

```bash
pip install PyQt6 pandas numpy openpyxl xlrd
```

## 🚀 使用說明

### 1. CLI 版本 (`Excel_Search_CLI.py`)

CLI 版本透過直接修改腳本內的變數來進行設定，適合快速執行故無封裝 exe。

1. 使用文字編輯器打開 `Excel_Search_CLI.py`。
2. 根據需求修改檔案上方的設定變數：
   - `SOURCE_FOLDER`: 要搜尋的目標資料夾路徑 (例如: `r"C:\Users\YourName\Documents"`)。
   - `OUTPUT_FILE`: 搜尋結果的匯出檔名與路徑 (例如: `"Excel_Search_Results.xlsx"`)。
   - `TARGET_KEYWORDS`: 欲搜尋的關鍵字清單 (例如: `["Google Chrome", "Microsoft Edge"]`)。
3. 在終端機中執行腳本：
   ```bash
   python Excel_Search_CLI.py
   ```
4. 程式會將搜尋結果與進度印出在終端機，完成後即可至設定的匯出路徑查看結果。

### 2. GUI 版本 (`Excel_Search_GUI.py`)

GUI 版本提供完整的使用者介面，無需修改程式碼即可進行所有設定。

1. 在終端機中執行腳本啟動程式：
   ```bash
   python Excel_Search_GUI.py
   ```
2. **選擇來源目錄**：透過介面選取您要搜尋的資料夾。
3. **設定匯出路徑**：指定結果檔案要儲存的位置。
4. **設定關鍵字**：輸入您要尋找的關鍵字，GUI 版本還支援「別名設定」（如 `版本=軟體版本`），讓搜尋更加彈性。
5. **開始搜尋**：點擊搜尋按鈕，下方清單與進度條會即時顯示尋找狀況與匹配筆數。

### 3. 使用打包exe

[Releases 下載](https://github.com/zz22558822/Excel_Search/releases/latest)
1. 基於 C 編譯 Excel_Search_GUI_C.exe
2. 基於 Pyinstaller 打包 Excel_Search_GUI_P.exe


## 🛠 打包為執行檔 (.exe)

若更改原碼需將 GUI 版本分享給沒有安裝 Python 環境的使用者，
可以參考 打包指令.txt 自行封裝打包。


## 📄 授權條款

MIT License