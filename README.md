# 土石流潛勢溪流現勘名單篩選

依「本計畫調查範圍初步篩選流程」從候選溪流試算表篩出年度現地調查建議名單，並匯出格式與原 Word 清冊相同的
「新增及調整土石流潛勢溪流調查評估之○條建議現勘清單.docx」。

純前端網頁，不需登入、不需伺服器；試算表只在使用者的瀏覽器中處理，不會上傳。

**線上使用**：`https://<帳號>.github.io/<倉庫名稱>/`（啟用 GitHub Pages 後的網址）

![篩選流程](docs/調查範圍篩選流程.jpg)

## 功能

| 步驟 | 內容 |
|---|---|
| 1 匯入資料 | 匯入 .xlsx／.xls／.csv，或貼上 Google 試算表、OneDrive 分享連結；自動找出標題列與最適合的工作表 |
| 2 欄位對應 | 依範例欄名自動預選，可手動調整；崩塌率自動辨識小數或百分比格式 |
| 3 篩選設定 | 調查數量、調查年度、類別優先順序（拖曳）、各類上限數量、類內排序依據、優先納入／遞補；篩選類別空白者可依流程圖自動判定第 1~4 類；新增業主指定類別（重大災例、陳情案等） |
| 4 篩選清冊 | 即時預覽、搜尋、依類別過濾；縣市／鄉鎮市區／村里／辦理原因可直接修改；移出溪流後自動遞補 |
| 5 匯出 Word | 套用原 Word 範本（A4 直式、表頭跨頁重複、頁尾「附件-頁碼」），中文標楷體、英文數字 Times New Roman；標題與註記依名單自動產生，可修改 |
| 變更紀錄 | 每項更動記錄時間、版本、更動位置、更動項目、原值與新值；可匯出 CSV 與設定檔 JSON |

## 版本紀錄與追溯

本工具有兩層紀錄，詳見 [CONTRIBUTING.md](CONTRIBUTING.md)：

1. **程式的更動**：由 Git 記錄。每次提交都能以 `git log -p`、`git blame` 查到改了哪個檔案、哪一行、誰改、為什麼改；
   各版重點整理在 [CHANGELOG.md](CHANGELOG.md)，正式版以 Git 標籤（例如 `v1.0.0`）標記。
2. **名單的更動**：由網頁「變更紀錄」頁記錄使用者的每項操作（例如「3 篩選設定／第1類｜上限數量｜不限 → 20」）。
   按「匯出設定檔」取得 JSON，提交到倉庫（建議放在 `records/` 資料夾），即可在 Git 中留存每一版名單的設定、入選清單與操作紀錄。

## 本機使用

直接用瀏覽器開啟 `index.html` 即可（所有套件都在 `vendor/`，可離線使用）。

若要讀取雲端連結或測試完整功能，建議用本機伺服器開啟：

```bash
npm run serve        # 開啟 http://localhost:8080
```

## 部署到 GitHub Pages

1. 在 GitHub 建立新倉庫（Public 才能讓所有人檢視；Private 倉庫的 Pages 需付費方案）。
2. 推送本專案：
   ```bash
   git remote add origin https://github.com/<帳號>/<倉庫名稱>.git
   git push -u origin main --tags
   ```
3. 倉庫 **Settings → Pages → Build and deployment → Source** 選 **GitHub Actions**。
4. 之後每次推送到 `main`，`.github/workflows/pages.yml` 會先跑測試，通過後自動部署。

## 專案結構

```
index.html                 網頁
css/style.css              樣式
js/core.js                 篩選邏輯與 Word 匯出（可在 Node 測試）
js/app.js                  介面、變更紀錄、設定檔匯出入
js/version.js              版本號
vendor/                    前端套件（SheetJS 0.18.5、JSZip 3.10.1）與授權、SHA256SUMS
template/template.docx     Word 範本（已移除資料列與作者資訊）
template/row.xml           資料列範本
template/template.js       由上兩者產生的內嵌範本（勿手動修改）
examples/                  虛構範例試算表
scripts/                   範本產生、套件更新與檢查、版本檢查
test/                      自動測試
.github/workflows/         CI 檢查與 Pages 部署
docs/                      篩選流程圖
```

## 開發指令

```bash
npm ci                                   # 安裝測試用套件
npm test                                 # 執行測試
npm run version:check                    # 檢查三處版本號一致
npm run vendor:verify                    # 檢查 vendor/ 套件未被竄改
npm run vendor:update                    # 升級套件後更新 vendor/ 與 SHA256SUMS
npm run template:build -- 新範本.docx    # 由新的 Word 清冊重建範本
npm run example                          # 重新產生範例試算表
```

## 資料注意事項

- 公開倉庫請勿提交實際業務資料（含現勘人員姓名等欄位的試算表）；`.gitignore` 已排除 `data/`、`private/`。
- 提交設定檔 JSON 前請確認內容可公開；設定檔含入選溪流編號、自訂類別與操作紀錄，但不含整份試算表。
- 「變更紀錄」存在各使用者瀏覽器，不同電腦之間不共用；需要共同留存時，請匯出設定檔提交到倉庫。

## 授權

程式採 MIT 授權（見 [LICENSE](LICENSE)）；SheetJS 為 Apache-2.0、JSZip 為 MIT／GPLv3 雙授權，授權文件見 `vendor/`。
