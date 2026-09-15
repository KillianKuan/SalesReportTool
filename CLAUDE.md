
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

### 資料夾實際位置（打包後）
| 平台 | data/ 根目錄 |
|------|-------------|
| Windows `.exe` | exe 同層的 `data\`（未變動） |
| macOS `.app` | `~/Library/Application Support/SalesReportTool/data/` |
| 原始碼開發 | repo 內的 `data/` |

`.app` bundle 為唯讀，launcher 每次啟動會把 bundle 內的 `app/` 鏡射到
`~/Library/Application Support/SalesReportTool/app`（保留 `overrides.json` 與
`settings.json`，見 `launcher.py` 的 `USER_STATE_FILES`），並從該副本啟動
Streamlit。因此 `utils.py` 既有的 `DATA_DIR = APP_DIR.parent / "data"` 與
`OVERRIDES_FILE` / `SETTINGS_FILE` 不需任何修改就會落在可寫路徑。

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

## 使用者設定（⚙️ Settings tab / settings.json）

第 4 個主 tab（順序固定在 Company Dashboard 之後），讓一般使用者透過 UI 修改設定，不需編輯
程式碼或 JSON。三個子區塊用 `st.radio`（非 `st.tabs`）切換——因為每次存檔都呼叫
`st.rerun()`，`st.tabs()` 在 programmatic rerun 後會重置回第一個分頁，改用帶 `key` 的
`st.radio` 才能保留使用者所在的子區塊。

### `app/settings.json`
與 `overrides.json` 同等級的使用者可寫檔案：macOS 打包版的 `app/` 鏡射步驟會保留它
（`launcher.py` 的 `USER_STATE_FILES`），**永遠不會寫回 `aliases.json`**（後者維持
git-tracked 的出貨預設值；settings.json 在讀取時疊加在其上）。

```json
{
  "theme": "system",
  "ignored_customers": ["MITAC COMPUTERKUNSHAN COLTD"],
  "customer_aliases": {},
  "fcst_customer_aliases": {},
  "custom_groups": []
}
```

- `utils.load_settings()` / `save_settings()`：容錯讀取（檔案不存在或格式壞掉 → 回傳
  schema 預設值，絕不 crash），寫入時與現有檔案內容 merge。
- `utils.DEFAULT_SETTINGS`：schema 預設值來源；`ignored_customers` 的預設值就是舊有的
  `EXCLUDED_CUSTOMERS`（現在只作為預設值來源，不再是獨立的第二條過濾路徑）。
- **Cache busting**：`utils._settings_hash()` 對 `ignored_customers` /
  `customer_aliases` / `fcst_customer_aliases` 做 hash，`_rules_key()` 回傳
  `(DES_RULES tuple, settings_hash)`；`fcst_loader.load_fcst()` 也多帶一個
  `settings_key` 參數。任何一項設定變更都會讓 `load_single_file()` /
  `load_historical_csv()` / `load_fcst()`（原本 300 秒 TTL）的 `@st.cache_data`
  立即失效重新載入，不用等 TTL 過期或重啟。

### Section A — Theme
Light / Dark / System（預設 System）。`utils.inject_theme_css()`：Light/Dark 透過
`st.markdown()` 注入 CSS 覆蓋 `.stApp` 背景／文字色，當下 session 立即生效，並顯示一行
提示「重新啟動可套用完整原生主題」；System 不注入任何 CSS。

### Section B — Customer Ignore List
- 只用 Customer Name（normalized：去標點、大寫）比對，不支援 Part Number / sheet 層級。
- **單一過濾路徑**：實際過濾邏輯是 `utils._apply_ignore_list()`，在
  `load_single_file()` / `load_historical_csv()` 回傳前執行；FCST 端在
  `fcst_loader._parse_sheet()` 內對 `normalize_fcst_customer()` 解析後的名稱比對。
  app.py 不再對 `all_df` 做任何 `EXCLUDED_CUSTOMERS` 二次過濾。
- 兩邊都會統計被排除的列數（`ignored_count` 回傳值／`fcst_loader.get_ignored_row_count()`），
  UI 顯示合計，方便使用者 sanity-check。

### Section C — Account Match（FCST ↔ Performance Report）
Performance Report 客戶名稱為 source of truth。`fcst_loader.normalize_fcst_customer()`
查找順序（**settings.json 疊加在 aliases.json 之上，settings 優先**）：
1. `aliases.json` "fcst_customer" + `settings.json` "fcst_customer_aliases" —— 先 exact，
   再 case-insensitive
2. `aliases.json` "customer" + `settings.json` "customer_aliases"（與 Shipping Record
   共用的 alias section）
3. Fallback → `"{sheet}_Others"`，並記錄未匹配客戶（含 FCST sheet 名稱與 Forecast 金額，
   供 Needs Mapping 表格使用）

- **Needs Mapping** 表格：列出未匹配的 FCST 客戶名稱／Sheet／Forecast 金額，每列一個
  inline selectbox，可指定既有 Performance Report 客戶名稱或 Custom Group，寫回
  `settings.json` 的 `fcst_customer_aliases`。
- **Custom Groups**（`settings.json` 的 `custom_groups`）：可建立額外「Others」式分類
  （例如 `Others - EMEA Distributors`），可改名（連動更新所有指向該 group 的 mapping）
  與刪除（刪除後受影響的 mapping 還原為預設 `{sheet}_Others`）。
- 也可在此編輯 Shipping Record 的 "customer" alias section（不只 fcst_customer）。
- 驗證（`utils.validate_alias_mappings()`）：self-reference（key 正規化後等於 value）、
  正規化後 key 衝突、target 在目前資料／custom groups 中都找不到——以 warning 顯示，
  不會阻擋儲存。

### Reset
每個 section 都有各自的「Reset to Default」按鈕；頁面底部另有「Reset ALL Settings」，
需二次確認（`st.session_state["settings_confirm_reset_all"]`）。

### 新增客戶 mapping / ignore 名單的優先順序
- **一般使用者**：一律透過 UI（⚙️ Settings 頁面），寫入 `settings.json`，立即生效、
  跨重啟與版本升級保留。
- **開發者調整出貨預設值**：`aliases.json`（fcst_customer / customer section）或
  `utils.DEFAULT_SETTINGS["ignored_customers"]`——這些是新安裝或使用者尚未覆寫時的初始值，
  改完需要重新打包／發版才會生效，且不會覆寫使用者既有的 `settings.json`。

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
`aliases.json` 的 "fcst_customer" section（出貨預設值）+ `settings.json` 的
`fcst_customer_aliases`（使用者透過 UI 新增／覆寫）：FCST 檔案名稱 → Performance Report
正規化名稱。完整查找順序、Needs Mapping UI、Custom Groups 見上方
「使用者設定（⚙️ Settings tab / settings.json）」章節。

多個 FCST 名稱可對應同一個正規化名稱（如 Zonar-CDR + Zonar-Tablet → Zonar System Inc.）。

### Sidebar 選項
`All Sheets`（預設）/ `Div.1&2_All` / `VT` / `Signify`
`All Sheets` 時傳 `sheet_name=None` 給 `get_fcst_for_dashboard()`，自動合併全部 sheets。

### Budget 整合
`agg_budget_monthly(fcst_df)` 讀取 FCST 的 `AMT_Budget`/`GP_Budget`/`QTY_Budget` 欄，
輸出與 `agg_blended_monthly()` 相同格式（`Source = "Budget"`）。

Dashboard 行為：
- KPI 列新增 **Budget Achievement%**（YTD Actual / FY Budget Revenue）與 **FY Budget Revenue**
- 月越勢圖將 Budget 以灰色虛線疊加在 Actual / Forecast 之上（`_SOURCE_COLOR` / `_SOURCE_DASH`）
- `chart_gp_trend_blended` 中 Budget bar 以低透明度 (0.3) 獨立渲染，不與 Actual/Forecast 疊加

### Customer Drill-Down FCST 整合
當 FCST 資料可用且當年度被選取時，Customer Drill-Down 會：
- 篩選 `_fcst_raw` 只留選取客戶
- 重新執行 `blend_actual_fcst()` + `agg_blended_monthly()` + `agg_budget_monthly()`
- 顯示 **FY Forecast KPIs** 列（FY Forecast Revenue、GP、Budget Achievement%、FY Budget Revenue）
- 月收入圖改用 `chart_revenue_trend_blended`（Actual + Forecast + Budget）

### 未匹配客戶警告位置
`get_unmatched_customers()` 在 FCST 載入後由 Company Dashboard 呼叫，
警告訊息直接顯示在 Dashboard 頁面頂部（非 sidebar），並附連結指向
**⚙️ Settings ▸ Account Match** 的 Needs Mapping 表格進行指派。

---

## 跨平台打包與發版（CI）

開發環境為 macOS（Apple Silicon），同事在 Windows 執行工具。**PyInstaller 無法跨平台編譯**，
所以兩個平台各自在對應的 runner 上打包，並把產物掛到**同一個** GitHub Release。

### 關鍵檔案
| 檔案 | 用途 |
|------|------|
| `.github/workflows/build-windows.yml` | CI：`windows-latest` + Python 3.11，產生 `SalesReportTool-windows.zip` |
| `.github/workflows/build-macos.yml` | CI：`macos-14`（arm64）+ Python 3.11，呼叫 `./build-mac.sh --skip-deps`，用 `ditto` 產生 `SalesReportTool-macOS-arm64.zip` |
| `build.bat` | Windows / VM 本地打包 fallback（未變動） |
| `build-mac.sh` | macOS arm64 `.app` 打包；**macOS PyInstaller flags 的唯一來源**，CI 也呼叫它 |
| `run.sh` | macOS/Linux 開發 server（`streamlit run app/app.py`） |

### 平台差異一覽
| 項目 | Windows `.exe` | macOS `.app` |
|------|----------------|--------------|
| Log | exe 同層 `salesreport.log` | `~/Library/Logs/SalesReportTool/salesreport.log` |
| Streamlit app | exe 同層 `app/` | `~/Library/Application Support/SalesReportTool/app`（鏡射） |
| data/ | exe 同層 `data\` | `~/Library/Application Support/SalesReportTool/data` |
| 圖示 | `assets/app.ico` | `assets/app.icns`（由 `.ico` 於 build 時生成，git-ignored） |
| 致命錯誤對話框 | `MessageBoxW` | `osascript display dialog` |
| PyInstaller | `--noconsole` | `--windowed` + `--osx-bundle-identifier` + `--target-architecture arm64` |

> **原則**：`app/` 底下的程式碼完全平台無關、不得 fork；所有 `sys.platform` 分歧只寫在
> `launcher.py`。共用的 hidden-import / collect flags 需在 `build.bat`、`build-windows.yml`、
> `build-mac.sh` 三處同步更新。

### 發版流程（推薦）
1. push 你的變更。
2. 打 tag 並 push：`git tag v3.7.0 && git push origin v3.7.0`。
3. 兩個 workflow 同時執行，分別把 `SalesReportTool-windows.zip` 與
   `SalesReportTool-macOS-arm64.zip` 附到同一個 GitHub Release。
4. 使用者從 Release 下載對應平台的 zip。

也可從 **Actions** tab 手動觸發（`workflow_dispatch`），`.zip` 會以 workflow artifact 提供下載。

### 簽章 / 公證
目前不做：`.app` 未簽章，首次開啟需右鍵 → **Open**（或
`xattr -dr com.apple.quarantine`）。`build-macos.yml` 內有標記好的 TODO 位置，取得
Developer ID 憑證後在 zip 步驟前加入 `codesign` + `xcrun notarytool`。

---

## 關鍵設計決策

- **overrides.json**: Key 為 (Customer Name, Part Number, Month, DES) 的複合
  key，避免 Excel 更新後 index 偏移。跨 session / 重啟保留。
  macOS 打包版寫入 Application Support 內的鏡射副本，升級 `.app` 時會被保留（其餘 `app/`
  檔案每次啟動從 bundle 更新）。
- **Cache busting**: DES_RULES 變更時透過 `_rules_key()` 自動使 `@st.cache_data` 失效
  （settings 相關的 cache busting 見下方「Settings cache busting」）。
- **launcher 架構（v3.6+）**: 父 process 持有 pystray system tray / menu bar icon（主執行緒
  blocking），子 process 執行 Streamlit，stdout/stderr 導入 log 檔。
  browser open + server ready 偵測在 daemon thread 執行。
- **平台分歧集中於 launcher.py**: log 路徑、`app/` 鏡射、data 資料夾建立與 seed、圖示格式、
  錯誤對話框、開啟 log 的方式皆由 `IS_WINDOWS` / `IS_MACOS` 與 `macos_bundle_dir()` 判斷。
- **Single-instance 保護**: 啟動時檢查 `TEMP/salesreport.lock`（JSON 含 PID + port）。
  以 port 是否仍在服務判斷存活（`os.kill(pid, 0)` 在 Windows 不可靠）：port 有回應 →
  開瀏覽器到已執行的 instance + sys.exit；否則刪除 stale lock。此邏輯 Windows/macOS 共用。
  atexit + SIGTERM/SIGINT handler 確保 lock 在正常/異常結束時都會清除。
- **System tray**: pystray + Pillow；macOS 優先載入 `assets/app.icns`，Windows 優先 `app.ico`，
  fallback 為程式生成的藍色長條圖圖示。選單：Open Browser / Open Log / Quit。
- **--server.headless true**: launcher.py 控制開瀏覽器時機（偵測 port 就緒
  再開），不依賴 Streamlit 預設行為。
- **更新 app.py 不需重新打包**: Windows 直接替換 `dist/SalesReportTool/app/` 下的檔案；
  macOS 可替換 `~/Library/Application Support/SalesReportTool/app/` 下的檔案（注意下次啟動
  會被 bundle 內容覆蓋）。
- **FCST aliases cache**: `_load_fcst_customer_aliases()` 使用 module-level
  `_ALIASES_CACHE`，每個 process 只讀一次 aliases.json；`settings.json` 的覆蓋值則每次
  即時讀取（不快取），確保 UI 存檔後立即生效。
- **settings.json**: 與 `overrides.json` 同機制的使用者設定檔（theme / ignore list /
  account match aliases / custom groups）。schema 預設值定義在 `utils.DEFAULT_SETTINGS`，
  容錯讀取（缺檔／壞檔 → 預設值，不 crash），永遠不寫回 `aliases.json`。macOS 打包版由
  launcher 的 `USER_STATE_FILES` 保留。
- **Settings cache busting**: `_settings_hash()` 併入 `_rules_key()`，
  `fcst_loader.load_fcst()` 多帶一個 `settings_key` 參數，讓 ignore list / alias
  mapping 的變更立即讓 Performance Report、Historical、FCST 三邊的 cache 失效。
- **Settings 子導覽用 st.radio 而非 st.tabs**: 每次存檔都呼叫 `st.rerun()`，
  `st.tabs()` 在 programmatic rerun 後會跳回第一個分頁；改用帶 `key` 的 `st.radio`
  才能保留使用者所在的子區塊。

---

## 目前版本

v3.9（最新，已發版）— Windows 打包硬化，降低防毒軟體誤判（見 README Change Log）。

已合併（待下個 tag 發版）：⚙️ Settings tab（Theme / Customer Ignore List /
Account Match）+ `settings.json` 持久化設定，settings hash 併入快取失效機制；
macOS（Apple Silicon）打包與發版流程 — `build-mac.sh` 產出 arm64 `.app`、
新增 `build-macos.yml`，一個 tag 同時產出雙平台 Release 產物。

v3.6 — 資料夾結構重構（Over the Years / Current Year）。

v3.5 — Budget 整合 + Customer Drill-Down FCST + Signify 獨立分類。

### 核心模組
| 檔案 | 職責 |
|------|------|
| `app.py` | Streamlit UI、tab 邏輯（含 ⚙️ Settings）、FCST/Budget blend 觸發、Customer Drill-Down FCST |
| `utils.py` | 資料載入、Category 分類（含 CUSTOMER_CATEGORY_MAP）、KPI 計算、圖表資料準備、Settings 讀寫（`load_settings`/`save_settings`） |
| `charts.py` | Altair 圖表函式（Actual / Forecast / Budget 三線並呈） |
| `fcst_loader.py` | FCST Excel 解析、blend、Budget aggregation、customer name mapping（aliases.json + settings.json 疊加） |
| `launcher.py` | 打包後入口；tray icon、單一實例、log、Windows/macOS 路徑分歧、`overrides.json`/`settings.json` 保留 |

---

## 常見工作模式

- 修改分類規則 → 編輯 `utils.py` 內的 `DES_RULES`，並同步更新 Notion 對照表（DES_RULES
  不在 Settings UI 範圍內，仍為程式碼層級設定）
- 忽略特定客戶 → **一般使用者**透過 UI ⚙️ Settings ▸ Customer Ignore List；開發者調整
  出貨預設值則編輯 `utils.py` 的 `DEFAULT_SETTINGS["ignored_customers"]`（= 原
  `EXCLUDED_CUSTOMERS`）
- 新增 FCST 客戶 mapping → **一般使用者**透過 UI ⚙️ Settings ▸ Account Match（Needs
  Mapping 表格 inline 指派，寫入 `settings.json`，立即生效）；開發者調整出貨預設值則編輯
  `aliases.json` 的 "fcst_customer" section
- 新增 FCST Sheet → `FCST_SHEETS` 加 entry + app.py sidebar radio 加選項
- 新功能開發（macOS/Linux）→ `./run.sh`（或 `py -m streamlit run app/app.py`）
- 出貨給使用者 → push `vX.Y.Z` tag，GitHub Actions 同時產生 Windows `.exe` zip 與
  macOS arm64 `.app` zip 並附到同一個 GitHub Release；本地 `build.bat` / `./build-mac.sh` 為 fallback
- 本地打包 macOS 版 → `./build-mac.sh`（產 `dist/SalesReportTool.app`，未簽章）
- 修改打包參數 → `build.bat`、`build-windows.yml`、`build-mac.sh` 三處同步
- 小修正（只改 app 層檔案）→ 直接替換 `dist/SalesReportTool/app/` 下的對應檔案
- 年度結算（新年開始）→ 執行 `python scripts/merge_historical.py` 將舊當年度合併入 `historical.csv`，
  再將新年度 xlsx 放入 `data/Current Year/`
