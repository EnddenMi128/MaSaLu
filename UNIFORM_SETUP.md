# MaSaLu 團服確認（2026-11-05）

本更新保留原有排班與 LINE 同意／不同意流程，新增每日期團服：

- 白色團服（`white`）
- 黑色團服（`black`）
- 暫不選（預設，空字串）

## 行為

1. 排班管理員在 HTML 編輯日期時可選團服。
2. 若當天有實際工作班次而團服為「暫不選」，LINE 立即通知已綁定的 A02，提供白／黑按鈕；新排班通知暫緩五分鐘。
3. A02 在五分鐘內選擇，該日期團服立即保存，未發送的新班次隨即發送。
4. 五分鐘屆滿仍未選擇，發送標示「未擇衣」的排班確認通知。計時檢查由 Cloudflare Cron 每分鐘執行；實際發送可能因平台排程或網路略晚於五分鐘。
5. 上班當天 **12:00 Asia/Taipei** 若仍未選擇，系統再通知 A02 一次。成功通知後，同日不會因 Cron 重跑而重複提醒。
6. 逾時後 A02 仍可選擇，系統會向原已收到排班的人員發送團服更新；原本的排班同意狀態及通知版本維持不變。
7. 完全清空排班或僅有休假班次時不需選團服，不執行待選計時。

## 部署前檢查

- GAS 的 `LINE_CHANNEL_ACCESS_TOKEN`、`SPREADSHEET_ID`、`GAS_SHARED_SECRET` 正常設定。
- 員工名單中 A02 已綁定有效 LINE_UID。
- 保持 `GAS/UniformNotifications.js` 隨其他 GAS 程式一起發布。
- Cloudflare Worker 更新 `scheduled()` 與 `wrangler.jsonc` 的 `* * * * *` Cron。Cron 採 UTC 排程，GAS 以台灣時間判斷當日 12:00。
- 首次讀取會增建 `MaSaLu團服` 工作表，**不更動**原有 `排班表` 17 欄。
- 使用原專案 `.github/workflows/masalu-deploy.yml` 的正式部署保護流程；請先核對線上 GAS 版本，再發佈。

## 自動測試

```bash
node --test tests/*.mjs
```

注意：這些為離線測試。正式 LINE 推播、Cloudflare Cron 部署與 LIFF 畫面需要正式環境另外驗證。
