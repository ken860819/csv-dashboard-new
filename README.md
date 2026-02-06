# CSV Dashboard Desktop (PySide6)

**重點**
- 桌面 GUI，不使用瀏覽器、不需要 localhost
- CSV 路徑讀取（支援日期樣板）
- 篩選器、欄位架、多圖表儀表板
- 產出 `data/latest.csv` 與每日歷史檔

**CSV 讀取方式（操作步驟）**
1. 開啟程式後，在「資料來源」輸入 CSV 檔案完整路徑  
   例：`C:\Users\name\Desktop\data.csv`
2. 按「從路徑更新」  
   - 或直接按「選檔並更新」，選檔後會立即讀取  
3. 資料會載入到預覽表格，並寫入：  
   - `data/latest.csv`
   - `data/history/history_YYYY-MM-DD.csv`（若勾選「保留每日歷史檔」）
4. 若要重新讀取最新檔案，按「載入最新」

**日期樣板（每日自動讀檔）**
- 勾選「使用日期樣板」
- 路徑可用：`{date}` 或 `{date:%m%d}`  
  例：`C:\Reports\\daily_{date:%m%d}.csv`
- 若路徑沒有 `{date}`，程式會嘗試替換檔名最後一段數字  
  例：`daily_0206.csv` → `daily_0207.csv`
- 勾選「啟動時自動更新」可在開啟程式時自動讀今日檔案

**篩選器視窗**
- 主視窗點「開啟篩選器視窗」
- 篩選結果會立即影響圖表與儀表板

**樞紐（Rows / Columns / Values）**
1. 先用「篩選器」設定要的資料範圍  
2. 左側「樞紐」區塊選擇列（Rows）  
3. 左側「樞紐」區塊選擇欄（Columns）  
4. 左側「樞紐」區塊選擇值（Values，可多選）  
5. 左側「樞紐」區塊選擇彙總（sum / count / mean / ...）  
6. 按「套用樞紐」  
7. 右側「樞紐」分頁會顯示樞紐表與樞紐圖表

**樞紐圖表類型**
- 長條
- 折線
- 長條+折線（多欄位時折線會畫「總和」）

**完整標籤**
- 在「欄位架 / 圖表」勾選「完整標籤」
- X 軸會換行並完整顯示中文（不再縮成 ...）

**統計模板**
- 點「新增模板」可保存目前圖表與篩選設定
- 模板會自動命名為「模板1 / 模板2 / ...」
- 選擇模板後按「套用模板」即可還原配置

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
- `data/app.log`（程式日誌）
- `data/update.log`（更新/排程日誌）

**打包 Windows EXE（PyInstaller / onedir）**
1. `pip install pyinstaller`
2. `pyinstaller --onedir --noconsole --name CsvDashboardDesktop --collect-all PySide6 --collect-all pandas --collect-all numpy app.py`
3. `pyinstaller --onedir --noconsole --name CsvDashboardUpdater --collect-all PySide6 --collect-all pandas --collect-all numpy updater.py`

**Windows 工作排程（每天 12:00 自動更新）**
1. 把 `windows-bundle` 解壓縮  
2. 進入資料夾後執行 `schedule_task.bat`  
3. 排程會每天 12:00 執行 `CsvDashboardUpdater`
4. 若要移除排程，執行 `remove_task.bat`

**Nuitka（可選）**
```
python -m nuitka --standalone --enable-plugin=pyside6 --output-dir=nuitka_dist --output-filename=CsvDashboardDesktop.exe app.py
```

> EXE 不需要安裝 Python 或編譯工具即可執行  
> 若遇到閃退，請確認已安裝 Microsoft Visual C++ Redistributable (x64)
