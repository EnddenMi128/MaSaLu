# MaSaLu 第二 LINE OA 推播接線

此分支讓同一個 GAS 排班程式在主 OA 的推播 API 明確回覆月額度用盡時，將該筆通知送往副 OA。排班與通知紀錄仍由同一個 GAS 專案及試算表處理。

## 啟用條件

1. 第二 OA 的 Messaging API channel 與目前 LIFF／主 OA 位於同一 LINE provider。若不同，員工 UID 不同，本版不得填入好友清單。
2. 收件人已加入副 OA，且實際驗證副 OA 能向該 UID 推播。僅將已驗證的 UID 加入好友清單。
3. 在現有 GAS 專案的 Script Properties 新增 `MASALU_BACKUP_LINE_CHANNEL_ACCESS_TOKEN`，值為副 OA 的 channel access token。
4. 新增 `MASALU_BACKUP_FRIEND_UIDS`，值為已驗證好友的 UID，以逗號或換行分隔。不要把 token 或 UID 名單提交至 GitHub。
5. 先在測試人員身上驗證月額度錯誤及副 OA 收訊，再更新既有 GAS Web App 部署。

只有 HTTP 429 且 LINE 回覆 `You have reached your monthly limit.` 時嘗試副 OA。其他 429、未列入好友名單、缺少副 token 時，保留原失敗結果以供重試。推播成功的帳號寫入現有發送紀錄。LINE 回覆的即時訊息仍由收到 Webhook 的 OA 發出。

`wrangler.test.jsonc` 是第二 OA 獨立測試 Worker 的部署目標，供先驗證第二 OA Webhook 與憑證。該 Worker 的環境變數必須個別設定，不能直接沿用主 OA 的 channel secret 和 access token。這個設定檔本身不會建立 LINE channel 或部署 Worker。
