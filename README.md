# Trip Conflict Checker — 出差衝突檢查器（Outlook 版）

> **v1.0 個人版** · 部署於 https://bono519.github.io/trip-checker-outlook
> 帳號：c1552@csd.org.tw（Microsoft 365）

---

## 與 Google 版差異對照

| 項目 | Google 版 | Outlook 版（本程式）|
|-----|----------|-------------------|
| 帳號 | csd1552@gmail.com | c1552@csd.org.tw |
| 認證方式 | Google OAuth 2.0 + GIS | MSAL.js 2.x（Microsoft）|
| 日曆 API | Google Calendar API v3 | Microsoft Graph API v1.0 |
| 需要 API Key | ✅ 是 | ❌ 否（MSAL 不需要）|
| 需要 Client ID | ✅ 是 | ✅ 是（Azure App Registration）|
| 部署網址 | /trip-checker | /trip-checker-outlook |
| 功能完整性 | 完整 | 完整（相同功能）|
| 月費 | NT$0 | NT$0 |

---

## 功能說明

| 功能 | 說明 |
|-----|------|
| Microsoft 帳號登入 | MSAL.js，讀寫 Outlook 日曆 |
| 出差資訊輸入 | 地點、目的、日期時間、前後緩衝時間 |
| 行程分類引擎 | 不可移動 / 灰色地帶（預設不可移動）/ 可移動 |
| 不可移動關鍵字 | 長官職稱、重要事件、授課演講（可自訂）|
| 視覺化時間軸 | 顏色區分各類行程與空檔 |
| 空檔推薦 | 完整天 > 上午空 > 下午空 |
| 行程寫入 | PATCH 修改既有行程，POST 新增出差行程 |
| 匯出 | 複製文字摘要 / PDF 列印 |
| PWA | 可安裝至手機桌面 |

---

## 版本紀錄與擴充路線

| 版本 | 說明 |
|-----|------|
| v1.0 | 個人單機版。僅讀寫 c1552@csd.org.tw 的 Outlook 日曆。|

### 未來擴充路線（預留介面）
- **v1.1**：同時讀取 Google 版與 Outlook 版，合併衝突判斷
- **v2.0**：啟用 Microsoft 365 群組共用日曆
  - `app.js` CONFIG → `SHARED_CALENDAR_ENABLED: true`、填入群組 ID
- **v2.1**：Teams 整合，衝突確認後自動發 Teams 訊息

---

## 部署步驟

### 第一步：Azure Portal 應用程式註冊（約 15 分鐘）

1. 前往 https://portal.azure.com（用 c1552@csd.org.tw 登入）
2. 搜尋「Azure Active Directory」→ 應用程式註冊 → 新增註冊
3. 填入：
   - 名稱：`trip-checker-outlook`
   - 支援的帳戶類型：**此組織目錄中的帳戶**（單一租用戶）
   - 重新導向 URI：選「單頁應用程式（SPA）」→ 填入 `https://bono519.github.io/trip-checker-outlook`
4. 點擊「註冊」→ 複製 **應用程式（用戶端）識別碼**（即 CLIENT_ID）
5. 左側「API 權限」→ 新增權限 → Microsoft Graph → 委派的權限
   - 勾選：`Calendars.ReadWrite`、`User.Read`
   - 點擊「新增權限」
   - （不需要管理員同意，因為是委派權限）

> ⚠️ **注意**：若 c1552@csd.org.tw 的 Microsoft 365 是由 IT 管理，需請 IT 在 Azure AD 允許此應用程式，或由 IT 代為完成應用程式註冊。

### 第二步：取得 Tenant ID（選用）

若註冊時選「此組織目錄中的帳戶」，需填入 Tenant ID：

1. Azure Active Directory → 概觀 → 複製「目錄（租用戶）識別碼」
2. 或直接填 `'organizations'` 讓 MSAL 自動判斷

### 第三步：填入金鑰

開啟 `app.js`，找到頂部 CONFIG 區塊：

```javascript
const CONFIG = {
  CLIENT_ID: 'YOUR_CLIENT_ID_HERE',  // ← 貼上應用程式 ID
  TENANT_ID: 'organizations',        // ← 或填實際 Tenant ID
  ...
};
```

> Google 版需要 API Key，Outlook 版**不需要**，只需要 CLIENT_ID。

### 第四步：本機測試

```bash
cd trip-checker-outlook
python3 -m http.server 8080
# 瀏覽 http://localhost:8080
# 需在 Azure Portal 的重新導向 URI 加入 http://localhost:8080
```

### 第五步：部署至 GitHub Pages

```bash
# 在 GitHub 建立 repo：bono519/trip-checker-outlook

git clone https://github.com/bono519/trip-checker-outlook.git
# 複製所有檔案進 repo
git add .
git commit -m "v1.0 Outlook personal edition"
git push origin main

# GitHub → Settings → Pages → Branch: main → Save
# 部署網址：https://bono519.github.io/trip-checker-outlook
```

---

## 檔案結構

```
trip-checker-outlook/
├── index.html      主程式介面
├── style.css       樣式表（Microsoft 藍主題）
├── app.js          核心邏輯（MSAL + Graph API）
├── manifest.json   PWA 設定
├── sw.js           Service Worker
└── README.md       本文件
```

---

## 費用說明

- Microsoft Graph API：完全免費
- MSAL.js：完全免費（MIT 授權開源）
- Azure App Registration：完全免費
- GitHub Pages：完全免費
- **月費：NT$0**
