# 寶光崇正 2027 白陽慈航成績回報系統（第一版）

## 檔案
- `index.html`：上傳 GitHub Pages 的網頁
- `Code.gs`：貼到 Google Apps Script 的後端

## 安裝
1. 建立 Google 試算表，從網址取得試算表 ID。
2. 建立 Google Drive 資料夾，從網址取得資料夾 ID。**只讓管理者存取資料夾**。
3. 開啟 [script.google.com](https://script.google.com) 建立 Apps Script 專案，把 `Code.gs` 內容貼入。
4. 修改 `CONFIG.SPREADSHEET_ID`、`CONFIG.DRIVE_FOLDER_ID`。
5. 在 Apps Script 編輯器手動執行 `setupSheets` 並完成授權。此動作會建立「選項設定」及六個回報分頁。
6. 在「選項設定」從第 2 列起輸入國別、單位、區別，每列一個組合，例如 `台灣 | 北部 | 中和`。即使某類回報不需要區別，也要在選項表填寫區別以供連動；正式回報只檢查該類別必要的欄位。
7. 部署 → 新增部署 → 網頁應用程式；執行身分選「我」，存取權選「所有人」（匿名回報需要）。複製 `/exec` 網址。
8. 在 `index.html` 把 `API_URL` 改成 `/exec` 網址。
9. 建立 GitHub repository，上傳 `index.html`，Settings → Pages → Deploy from a branch → main / root。
10. 用手機與電腦測試六種回報、照片連結與各工作表寫入。

## 注意事項
- 此版本是**匿名公開回報的原型**，沒有登入、身分驗證、限流或防惡意灌單。正式對外開放前建議加上回報者驗證、管理員審核、資料保留及刪除政策。
- 求道卡與出席表可能包含個人資料；照片預設不公開，Google Drive 資料夾請勿開啟「知道連結的任何人」。請確認已取得適當的蒐集與上傳授權。
- GitHub Pages 與 Apps Script 跨來源提交可能受到瀏覽器 CORS、重導向或部署權限影響。請以正式部署網址測試；若遇到限制，建議改為 Apps Script HTML Service 同網域表單，或使用具備 CORS 支援的後端。
- 單張照片上限 5MB，每次最多 5 張；Apps Script 整體請求與執行時間另有限制，大量照片可能失敗。
- 此版本只保存原始回報，不自動計算頒獎、結算歸零或核驗辦道簿。
- 目前無法由伺服器可靠阻擋同一人重複回報；僅防止使用者在同一頁面連續點擊提交。
