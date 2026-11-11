# 寶光崇正｜2027 年白陽慈航盛會成績回報系統

> **版本狀態：開發測試中，尚未完成實際回報與照片上傳驗證，請勿直接作為正式成績收件系統。**

本專案以 **GitHub Pages + HTML + Google Apps Script + Google 試算表 + Google Drive** 建置線上成績回報頁面。填報者可依回報類別填寫資料，部分選單從試算表讀取；後端預計將成績寫入不同工作表，並將驗證照片儲存在 Google Drive。

- [前端網站（GitHub Pages；是否啟用需以實際頁面為準）](https://chy0205.github.io/ikt1114/)
- [Google 試算表：白陽慈航成績回報](https://docs.google.com/spreadsheets/d/1ZglfWKES6sHWcn2vi9JVHXHVJImPqzcWGb_rYHaj0Uk/edit)
- [照片儲存資料夾：白陽慈航成績回報](https://drive.google.com/drive/folders/1KlTFhCQhSX40oqEps8cP-BVXRNxhglUM)

## 回報類別

| 類別 | 填寫項目摘要 | 照片 |
| --- | --- | --- |
| 個人渡眾 | 國別、單位、區別、引師姓名、辦道日期、渡眾人數 | 求道卡（必填） |
| 家堂渡眾 | 國別、單位、區別、佛堂堂號、壇主姓名、辦道日期、渡眾人數 | 結算時另提供辦道簿核實 |
| 公堂渡眾 | 國別、單位、公堂堂號、辦道日期、渡眾人數 | 結算時另提供辦道簿核實 |
| 個人成全進階班 | 國別、單位、區別、進階班類別、成全者、班員、上課日期、第 1～24 次上課 | 出席表（必填） |
| 公堂進階班 | 國別、單位、公堂堂號、進階班類別、達標班員人數 | 出席表（必填） |
| 成全安堂完成 | 國別、單位、區別、成全者、壇主、正式安堂日期 | 當天佛堂入鏡照片（必填） |

進階班類別為 **禮行、義理、仁德**。

**成績結算規則（尚未程式化）**：第一梯次截止 2027/05/23 23:59；第二梯次截止 2027/11/07 23:59。第一梯次已獲獎的獎項重新累計，未達標成績得延續至第二梯次。系統目前只記錄原始回報，不會自動歸零、核獎或排名。

## 專案檔案

| 檔案 | 用途 |
| --- | --- |
| `index.html` | 六類回報表單、選單連動、照片預覽與前端資料送出 |
| `Code.gs` | Apps Script API、選項讀取、資料驗證、工作表寫入、照片儲存 |
| `README.md` | 使用方式、設定狀態、更新紀錄 |

## 安裝與部署

1. 開啟 [Google 試算表](https://docs.google.com/spreadsheets/d/1ZglfWKES6sHWcn2vi9JVHXHVJImPqzcWGb_rYHaj0Uk/edit)，確認「選項設定」分頁的 A～C 欄依序為 `國別`、`單位`、`區別`，從第 2 列開始填入選項。
2. 將 GitHub 最新版 `Code.gs` **手動同步**到原 Google Apps Script 專案。GitHub 的修改不會自動更新已部署的 Apps Script。
3. 確認 `CONFIG.SPREADSHEET_ID` 與 `CONFIG.DRIVE_FOLDER_ID` 正確，執行 `setupSheets()` 並授權。它會建立六個回報工作表（如不存在），不應刪除既有資料。
4. 重新部署 Apps Script 的「網頁應用程式」版本，並確認 `index.html` 的 `API_URL` 對應目前有效的 `/exec` 網址。
5. GitHub 儲存庫 → **Settings → Pages** → **Deploy from a branch** → 選 `main` / `/(root)` → Save。
6. 使用非敏感的測試資料，依序測試選項讀取、六種回報寫入、照片存檔、手機介面與失敗提示；測試完成後刪除測試資料。

## 開發／部署狀態

| 項目 | 狀態 |
| --- | --- |
| 六類動態表單與前端照片選擇 | ✅ 程式已建立 |
| Google 試算表 ID 設定至 GitHub `Code.gs` | ✅ 已設定 |
| Google Drive 資料夾 ID 設定至 GitHub `Code.gs` | ✅ 已設定 |
| Apps Script `/exec` 網址設定至 GitHub `index.html` | ✅ 已設定 |
| 試算表「選項設定」資料讀取檢視 | ✅ 已檢視（現有國別、單位、區別資料） |
| 六個回報工作表建立 | ⏳ 未確認 |
| 已部署 Apps Script 同步最新版後端 | ⏳ 未確認 |
| GitHub Pages 成功公開 | ⏳ 未確認 |
| 跨來源 API 提交、照片上傳、寫入端對端測試 | ⏳ 尚未驗證 |
| 權限管控、反灌單、審核與統計結算 | ⏳ 待開發 |

## 更新紀錄（Changelog）

> 時間均為 **台灣時間（UTC+8）**，依 GitHub Commit 的建立時間記錄。每項「已更新」僅表示程式碼已提交，不等於 Apps Script 已重新部署或功能已完成實際測試。往後新增功能或修正問題時，請在此表最上方新增紀錄並附上 Commit 連結。

| 更新時間（台灣時間） | 更新內容 | 影響檔案／紀錄 |
| --- | --- | --- |
| 2026/11/12 06:31:44 | 設定照片上傳使用的 Google Drive 資料夾 ID。 | [`Code.gs`](https://github.com/chy0205/ikt1114/commit/ce4c8d30cfd72f28e1eea1398e1dd0108ce3ff40) |
| 2026/11/12 06:29:49 | 設定成績回報使用的 Google 試算表 ID。 | [`Code.gs`](https://github.com/chy0205/ikt1114/commit/a023ac7b02e55c79004b1e0d9426995c573cb6fd) |
| 2026/11/12 06:28:28 | 將現有 Google Apps Script `/exec` 網址填入前端 API 設定。 | [`index.html`](https://github.com/chy0205/ikt1114/commit/573f44fbd5add294195bb2ac4b56cbb9c8a6b8c9) |
| 2026/11/12 06:27:42 | 建立第一版專案說明與部署步驟。 | [`README.md`](https://github.com/chy0205/ikt1114/commit/1ed2cd93cf94d7dec3c91ce6297b4dc541fea2c3) |
| 2026/11/12 06:27:39 | 建立 Apps Script 後端：選項讀取、六類工作表寫入、照片儲存與初始化功能。 | [`Code.gs`](https://github.com/chy0205/ikt1114/commit/deedbe1ff890a8ba9c701f31a6c3e5bc397fdb7e) |
| 2026/11/12 06:27:37 | 建立第一版響應式 HTML 表單與六種回報欄位。 | [`index.html`](https://github.com/chy0205/ikt1114/commit/748c4e0937222a0aaa8547ae76d69e232fe9acde) |

完整版本歷史可至 [GitHub Commits](https://github.com/chy0205/ikt1114/commits/main/) 檢視；這裡以容易閱讀的方式說明每次修改的目的。

## 注意事項

- **個資保護**：求道卡、出席表可能含個人資料，Google Drive 資料夾目前曾檢查到「知道連結的任何人可檢視」權限，正式使用前應改成「受限制」，只開放必要管理人員。請確認上傳者已取得適當授權。
- **公開回報風險**：此版本沒有登入認證、可靠的防重複回報或防惡意灌單機制，不宜直接對外正式開放。
- **跨網域注意**：GitHub Pages 向 Apps Script 發送資料可能遇到 CORS 或重新導向問題，目前尚未通過端對端測試。
- **照片限制**：最多 5 張、每張上限 5MB；大量上傳仍可能受 Apps Script 請求大小與執行時間影響。
- **維護原則**：不要將求道卡照片、帳號密碼或其他個資放入公開的 GitHub 儲存庫。
