
---

## 輸入資料規格

必要欄位（欄名完全一致）：
- `Customer Name` — 客戶名稱
- `Ship Date` — 出貨日期（容錯解析，NaT 自動跳過）
- `QTY` — 數量
- `SALES Total AMT` — 銷售額（TWD）
- `final GP(NTD,data from Financial Report)` — 毛利
- `Part Number` — 料號
- `Category` — 分類

選用欄位：
- `DES` — 用於 DES 關鍵字分類（若無此欄，DES 分類停用）

---

## Category 分類邏輯

優先順序：
1. Category 欄直接比對（Tablet / CDR / Tablet ACC / CDR ACC，大小寫不敏感）
2. DES 欄關鍵字比對（DES_RULES 字典，substring contains）
3. Fallback → Others

有效 Category：Tablet / CDR / Tablet ACC / CDR ACC / AI_SW / Others

DES_RULES（修改時需同步更新 app.py 頂部字典）：
- CDR ACC: cdr, gemini, evo, sprint, sd card, panic button, iosix, uvc camera,
           k220, k245, k265, smart link dongle, safetycam
- Tablet ACC: tablet, prometheus, chiron, hera, phaeton, surfing pro, cradle,
              f840, ulmo, fleet cable
- AI_SW: visionmax

---

## FCST 資料整合

### 檔案位置
`data/FCST/*.xlsx`（自動選最新修改的檔案）

### 支援 Sheets
`Div.1&2_All`、`VT`、`Signify`（`FCST_SHEETS` 常數）

### Blend 邏輯
- `month < current_month` → Actual（Shipping Record）
- `month >= current_month` → Forecast（FCST 檔案）
- 當月一律用 Forecast，不做 Actual 優先判斷

### 單位轉換
FCST 的 AMT / GP 是千元，`_parse_sheet()` 在建立 record 時自動 ×1,000。QTY 不轉換。

### Customer Name Mapping
`FCST_TO_CANONICAL`（在 `fcst_loader.py` 頂部）：FCST 檔案名稱 → Performance Report 正規化名稱。

查找順序（`normalize_fcst_customer()`）：
1. `FCST_TO_CANONICAL` 精確比對（先 exact，再 case-insensitive）
2. `aliases.json` 的 customer section（與 Shipping Record 共用）
3. Fallback → `"{sheet_name}_Others"`（例如 `Div.1&2_All_Others`）

新增客戶 mapping 時：只需在 `FCST_TO_CANONICAL` 加 entry，不需改其他地方。
多個 FCST 名稱可對應同一個正規化名稱（如 Zonar-CDR + Zonar-Tablet → Zonar System Inc.）。

### Sidebar 選項
`All Sheets`（預設）/ `Div.1&2_All` / `VT` / `Signify`
`All Sheets` 時傳 `sheet_name=None` 給 `get_fcst_for_dashboard()`，自動合併全部 sheets。

---

## 關鍵設計決策

- **overrides.json**: Key 為 (Customer Name, Part Number, Month, DES) 的複合
  key，避免 Excel 更新後 index 偏移。跨 session / 重啟保留。
- **Cache busting**: DES_RULES 變更時透過 _rules_key() 自動使
  @st.cache_data 失效。
- **--server.headless true**: launcher.py 控制開瀏覽器時機（偵測 port 就緒
  再開），不依賴 Streamlit 預設行為。
- **更新 app.py 不需重新打包**: 直接替換 dist/SalesReportTool/app/ 下的檔案即可。
  適用：app.py、charts.py、fcst_loader.py（小修正不需重新 build）。
- **FCST aliases cache**: `_load_fcst_customer_aliases()` 使用 module-level
  `_ALIASES_CACHE`，每個 process 只讀一次 aliases.json。

---

## 目前版本

v3.3（最新）— FCST 整合完成。

### 核心模組
| 檔案 | 職責 |
|------|------|
| `app.py` | Streamlit UI、tab 邏輯、FCST blend 觸發 |
| `utils.py` | 資料載入、Category 分類、KPI 計算、圖表資料準備 |
| `charts.py` | Altair 圖表函式（含 blended trend charts） |
| `fcst_loader.py` | FCST Excel 解析、blend、customer name mapping |

---

## 常見工作模式

- 修改分類規則 → 編輯 `app.py` 內的 `DES_RULES`，並同步更新 Notion 對照表
- 新增 FCST 客戶 mapping → 編輯 `fcst_loader.py` 的 `FCST_TO_CANONICAL`
- 新增 FCST Sheet → `FCST_SHEETS` 加 entry + app.py sidebar radio 加選項
- 新功能開發 → `py -m streamlit run app/app.py`
- 出貨給使用者 → 執行 `build.bat` 重新打包
- 小修正（只改 app 層檔案）→ 直接替換 `dist/SalesReportTool/app/` 下的對應檔案
