# MaSaLu 第二 LINE OA 推播接線

此分支讓同一個 GAS 排班程式在主 OA 的推播 API 明確回覆月額度用盡時，將該筆通知送往副 OA。排班與通知紀錄仍由同一個 GAS 專案及試算表處理。

## 啟用條件

1. 第二 OA 的 Messaging API channel 與目前 LIFF／主 OA 位於同一 LINE provider。若不同，員工 UID 不同，本版不得填入好友清單。
2. 收件人已加入副 OA，且實際驗證副 OA 能向該 UID 推播。僅將已驗證的 UID 加入好友清單。
3. 在現有 GAS 專案的 Script Properties 新增 `MASALU_BACKUP_LINE_CHANNEL_ACCESS_TOKEN`，值為副 OA 的 channel access token。
4. 新增 `MASALU_BACKUP_FRIEND_UIDS`，值為已驗證好友的 UID，以逗號或換行分隔。不要把 token 或 UID 名單提交至 GitHub。
5. 第二個 OA 必須擁有**自己的** Messaging API Webhook。部署 `wrangler.test.jsonc` 指向的測試 Worker，分別設定第二 OA 的 `LINE_CHANNEL_SECRET`、`LINE_CHANNEL_ACCESS_TOKEN`，以及同一個 GAS 的 `GAS_WEB_APP_URL`、`GAS_SHARED_SECRET`。LINE Developers 第二 OA Webhook URL 指向該測試 Worker 的 `/webhook`。
6. 先從第二 OA 向測試好友發送互動排班卡，再實際按「同意／不同意」，確認該第二 OA Webhook 收到事件、原 GAS 試算表回寫成功、LINE 回覆成功，且權限隔離正常。
7. **只有上述實機測試成功後**，在 GAS Script Properties 新增 `MASALU_BACKUP_OA_ENABLED=true`。預設值為停用，僅有 Access Token 與好友 UID 不會自動啟用備援。
8. 先在測試人員身上模擬月額度錯誤並驗證副 OA 收訊，再更新既有 GAS Web App 部署。更新正式程式前請確認 GitHub Actions 的自動部署設定，先採人工控制。

只有 HTTP 429 且 LINE 回覆 `You have reached your monthly limit.` 時嘗試副 OA。其他 429、未列入好友名單、缺少副 token 時，保留原失敗結果以供重試。推播成功的帳號寫入現有發送紀錄。LINE 回覆的即時訊息仍由收到 Webhook 的 OA 發出。

`wrangler.test.jsonc` 是第二 OA 獨立測試 Worker 的部署目標，供先驗證第二 OA Webhook 與憑證。該 Worker 的環境變數必須個別設定，不能直接沿用主 OA 的 channel secret 和 access token。這個設定檔本身不會建立 LINE channel 或部署 Worker。

## 第二 OA 圖文選單

啟用後，`masaluPublishCalendarRichMenu` 與定時更新會用兩個 OA 各自的 Channel Access Token 建立預設月曆圖文選單。兩邊的 Rich Menu ID 與 alias 分屬各自的 LINE channel。副 OA 只對 `MASALU_BACKUP_FRIEND_UIDS` 中已驗證的團員建立個人月曆，並使用獨立的 `MASALU_BACKUP_PERSONAL_MENU_` 狀態，避免覆蓋主 OA 的選單紀錄。第二 OA 的 Worker 將 `MASALU_CHANNEL=backup` 傳到 GAS，點選「更新個人班表」時由副 OA 建立選單並在該聊天室回覆。

部署 GAS 與第二 OA Worker 後，先從 Apps Script 執行 `masaluPublishCalendarRichMenu`，確認兩個 OA 都出現當週預設月曆。再由已加入第二 OA 且列在好友清單的測試員傳「更新個人班表」，確認第二 OA 個人月曆顯示正確，並點選日期驗證回覆。若兩個 OA 位於不同 provider，必須先擴充員工名單保存副 OA 的 UID，不能沿用主 OA UID。LINE 對尚未加好友的使用者可能接受選單連結 API 請求卻不顯示選單，因此需在第二 OA 聊天室實測。

## 發布與回滾安全

- 本整合先進入隔離審查分支，絕不在未完成第二 OA Webhook 實測前部署 `main`。
- 只對 LINE API 回覆 HTTP 429 且訊息明確為 `You have reached your monthly limit.` 的推播嘗試副 OA；其他錯誤不切換。
- 即使啟用開關，僅限 `MASALU_BACKUP_FRIEND_UIDS` 所列且已在副 OA 實測可收訊的 UID。
- 驗證失敗時可立刻將 `MASALU_BACKUP_OA_ENABLED=false`，不需刪除 Token 或重綁員工。
- 備援訊息由第二 OA 發出，其回覆/按鈕會送至第二 OA Webhook；不能只設定第一 OA 的 Webhook。
