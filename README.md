# 互動式極速問答挑戰網頁 (v2.0)

本資料夾包含依據系統設計規格書所實作的 Google Apps Script (GAS) + Google Sheets 互動式極速問答挑戰網頁專案。

## 檔案清單

- `Code.gs`：GAS 後端程式，負責 `doGet(e)` 伺服端資料注入 (`initialData`)、讀取 `Settings` / `Questions` 工作表、JSON 容錯與預設測試題庫（Fallback Mock Data），並附帶試算表一鍵初始化選單。
- `index.html`：前端單頁式應用程式 (SPA)，包含 Tailwind CSS 介面、Google Drive 圖片直連轉換與預載、10 秒倒數計時器、選擇題／連連看互動邏輯、Web Audio API 合成音效引擎與 `canvas-confetti` 紙花特效。
- `appsscript.json`：GAS 專案設定檔（時區 `Asia/Taipei`、Web App 權限設定）。

## 快速部署步驟

1. **建立 Google 試算表**：前往 Google Sheets 新增一份試算表。
2. **開啟 Apps Script**：點選試算表上方選單 `擴充功能` → `Apps Script`。
3. **貼上程式碼**：
   - 將 `Code.gs` 內容貼入編輯器中的 `Code.gs`（或 `程式碼.gs`）。
   - 點擊左側 `+` 新增 HTML 檔案，命名為 `index`（即 `index.html`），將本資料夾 `index.html` 內容完整貼入。
4. **初始化工作表（選用）**：
   - 重新整理 Google 試算表頁面後，上方會出現 `⚡ 極速問答後台` 選單。
   - 點擊 `初始化 Settings 與 Questions 工作表`，即可自動建立標準欄位與 10 題示範題庫。
5. **部署為 Web App**：
   - 回到 Apps Script 編輯器，點擊右上角 `部署` → `新增部署作業`。
   - 類型選擇「網頁應用程式 (Web app)」。
   - **執行身分**：選擇 `我 (Me)`。
   - **誰可以存取**：選擇 `所有人 (Anyone)`。
   - 點擊 `部署` 並完成授權，即可取得公開挑戰網址。

## 注意事項

- **Google Drive 圖片權限**：放入 `Settings` 或 `Questions` 的 Google Drive 圖片，共用設定必須為「知道連結的任何人 → 檢視者」。
- **更新版本規範**：每次修改 `Code.gs` 或 `index.html` 後，需點擊「管理部署作業 → 編輯 → 版本選擇『**新版本**』→ 部署」，正式網址才會套用最新程式碼。
