MaSaLu GAS 完整檔案 — 2026-11-01

覆蓋原 Apps Script 專案內「程式碼.gs」與「CalendarApi.gs」的完整內容。
請勿另建同名功能檔案，以免重複宣告 doPost / doGet。

本包新增：LINE access token 驗證、綁定 A01/A02/A03 管理權限、managerSession 查詢。
保留既有管理金鑰備援、排班通知、版本衝突檢查與資料保護。

重要：這是 GAS 後端修改包，尚不是完整免密碼上線包。
Worker 需把 LIFF 的 access token 放在 calendar_api JSON 頂層 lineAccessToken 並轉送。
GAS_SHARED_SECRET 必須繼續由 Worker 端加入，絕不可放在前端。
前端需完成 LIFF 登入、傳送憑證與免密碼管理介面，才會實際生效。
預設 LINE Login Channel ID：2010020470（來自先前 LIFF 頁面資訊，需核對）。
若不同，請設定指令碼屬性 MASALU_LINE_LOGIN_CHANNEL_ID。
LINE Login 與 Messaging API 必須在相同 Provider 下，UID 才可與員工綁定表對應。
未附前端、Worker、金鑰或 appsscript.json；請保留原設定。

完成覆蓋後，更新既有網頁應用程式部署至新版本。
本包僅完成本機模擬與語法檢查，未進行真實 LINE 帳號端對端測試。
