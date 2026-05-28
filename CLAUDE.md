
---

## 資料夾結構

```
data/
├── Over the Years/
│   └── historical.csv      # 所有歷史年份合併（UTF-8-BOM）
├── Current Year/
│   └── *.xlsx              # 當年度出貨記錄（Actual sheet）
└── FCST/
    └── *.xlsx              # 最新 FCST 檔（依 mtime 自動選擇）
```

### 初次遷移
執行 `python scripts/merge_historical.py` 將舊版 `data/{year}/` 年份資料夾合併為 `historical.csv`，
再將當年度 xlsx 移至 `data/Current Year/`。

### 讀取邏輯（`utils.py`）
| 常數 | 路徑 |
|------|------|
| `HISTORICAL_CSV` | `data/Over the Years/historical.csv` |
| `CURRENT_YEAR_DIR` | `data/Current Year/` |
| `HISTORICAL_DIR` | `data/Over the Years/` |

- `load_historical_csv(file_path, rules_key)` — 讀 CSV，套用與 `load_single_file` 相同清洗流程
- `scan_current_year_folder()` — 回傳 `Current Year/` 內最新 xlsx，或 `None`
- 年份選擇器從合併後的 df 的 `Ship Date` 年份自動推導（不再依賴資料夾名稱）

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
1. Customer Name 比對（`CUSTOMER_CATEGORY_MAP`，優先於所有其他規則）
2. Category 欄直接比對（Tablet / CDR / Tablet ACC / CDR ACC / Signify，大小寫不敏感）
3. DES 欄關鍵字比對（DES_RULES 字典，substring contains）
4. Fallback → Others

有效 Category：Tablet / CDR / Tablet ACC / CDR ACC / AI_SW / Signify / Others

CUSTOMER_CATEGORY_MAP（`utils.py` load_single_file 內）：
- SIGNIFY → Signify

DES_RULES（修改時需同步更新 `utils.py` 頂部字典）：
- CDR ACC: cdr, gemini, evo, sprint, sd card, panic button, iosix, uvc camera,
           k220, k245, k265, smart link dongle, safetycam
- Tablet ACC: tablet, prometheus, chiron, hera, phaeton, surfing pro, cradle,
              f840, ulmo, fleet cable
- AI_SW: visionmax
- Signify: signify

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
`aliases.json` 的 "fcst_customer" section：FCST 檔案名稱 → Performance Report 正規化名稱。

查找順序（`normalize_fcst_customer()`）：
1. `aliases.json` "fcst_customer" 精確比對（先 exact，再 case-insensitive）
2. `aliases.json` "customer" section（與 Shipping Record 共用）
3. Fallback → `"{sheet_name}_Others"`（例如 `Div.1&2_All_Others`）+ 收集未匹配客戶

新增客戶 mapping 時：只需在 `aliases.json` 的 "fcst_customer" section 加 entry，不需改程式碼。
多個 FCST 名稱可對應同一個正規化名稱（如 Zonar-CDR + Zonar-Tablet → Zonar System Inc.）。

未匹配客戶會在 sidebar System Info 中顯示警告，提示更新 aliases.json。

### Sidebar 選項
`All Sheets`（預設）/ `Div.1&2_All` / `VT` / `Signify`
`All Sheets` 時傳 `sheet_name=None` 給 `get_fcst_for_dashboard()`，自動合併全部 sheets。

### Budget 整合
`agg_budget_monthly(fcst_df)` 讀取 FCST 的 `AMT_Budget`/`GP_Budget`/`QTY_Budget` 欄，
輸出與 `agg_blended_monthly()` 相同格式（`Source = "Budget"`）。

Dashboard 行為：
- KPI 列新增 **Budget Achievement%**（YTD Actual / FY Budget Revenue）與 **FY Budget Revenue**
- 月趨勢圖將 Budget 以灰色虛線疊加在 Actual / Forecast 之上（`_SOURCE_COLOR` / `_SOURCE_DASH`）
- `chart_gp_trend_blended` 中 Budget bar 以低透明度 (0.3) 獨立渲染，不與 Actual/Forecast 疊加

### Customer Drill-Down FCST 整合
當 FCST 資料可用且當年度被選取時，Customer Drill-Down 會：
- 篩選 `_fcst_raw` 只留選取客戶
- 重新執行 `blend_actual_fcst()` + `agg_blended_monthly()` + `agg_budget_monthly()`
- 顯示 **FY Forecast KPIs** 列（FY Forecast Revenue、GP、Budget Achievement%、FY Budget Revenue）
- 月收入圖改用 `chart_revenue_trend_blended`（Actual + Forecast + Budget）

### 未匹配客戶警告位置
`get_unmatched_customers()` 在 FCST 載入後由 Company Dashboard 呼叫，
警告訊息直接顯示在 Dashboard 頁面頂部（非 sidebar）。

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

v3.6（最新）— 資料夾結構重構（Over the Years / Current Year）。

v3.5 — Budget 整合 + Customer Drill-Down FCST + Signify 獨立分類。

### 核心模組
| 檔案 | 職責 |
|------|------|
| `app.py` | Streamlit UI、tab 邏輯、FCST/Budget blend 觸發、Customer Drill-Down FCST |
| `utils.py` | 資料載入、Category 分類（含 CUSTOMER_CATEGORY_MAP）、KPI 計算、圖表資料準備 |
| `charts.py` | Altair 圖表函式（Actual / Forecast / Budget 三線並呈） |
| `fcst_loader.py` | FCST Excel 解析、blend、Budget aggregation、customer name mapping |

---

## 常見工作模式

- 修改分類規則 → 編輯 `utils.py` 內的 `DES_RULES`，並同步更新 Notion 對照表
- 新增 FCST 客戶 mapping → 編輯 `aliases.json` 的 "fcst_customer" section
- 新增 FCST Sheet → `FCST_SHEETS` 加 entry + app.py sidebar radio 加選項
- 新功能開發 → `py -m streamlit run app/app.py`
- 出貨給使用者 → 執行 `build.bat` 重新打包
- 小修正（只改 app 層檔案）→ 直接替換 `dist/SalesReportTool/app/` 下的對應檔案
- 年度結算（新年開始）→ 執行 `python scripts/merge_historical.py` 將舊當年度合併入 `historical.csv`，
  再將新年度 xlsx 放入 `data/Current Year/`
