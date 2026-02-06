# CSV Dashboard Desktop (PySide6)

**重點**
- 桌面 GUI，不使用瀏覽器、不需要 localhost
- CSV 路徑讀取 / 上傳
- 篩選器、欄位架、多圖表儀表板
- 產出 `data/latest.csv` 與每日歷史檔

**本機執行**
1. `python -m venv .venv`
2. `source .venv/bin/activate`（Windows 用 `\.venv\Scripts\activate`）
3. `pip install -r requirements.txt`
4. `python app.py`

**測試資料**
- 可用 `sample_data.csv`

**資料儲存**
- `data/latest.csv`
- `data/history/history_YYYY-MM-DD.csv`

**打包 Windows EXE（PyInstaller）**
1. `pip install pyinstaller`
2. `pyinstaller --onefile --noconsole --name CsvDashboardDesktop app.py`

> 若要更穩定，可用資料夾模式：
> `pyinstaller --noconsole --name CsvDashboardDesktop app.py`
