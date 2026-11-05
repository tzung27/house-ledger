# 房屋收支帳

把 `HOUSE.xlsx` 的購屋收支呈現成網頁，並可從網頁新增收支、直接寫回 Excel。

## 使用方式

- **新增收支**：雙擊 `開啟收支帳.bat`（或執行 `python server.py`），瀏覽器會開啟 `http://127.0.0.1:8765`。寫入前請先關閉 Excel 中的 `HOUSE.xlsx`；每次寫入前會自動備份到 `backups/`。
- **只瀏覽**：直接開啟 `index.html`。
- **手動改了 Excel 後**：執行 `python update_data.py` 更新 `data.js`。

金額單位為萬元。`lib/` 內附 openpyxl，不需另外安裝套件。
