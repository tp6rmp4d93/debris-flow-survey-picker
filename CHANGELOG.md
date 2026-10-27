# 版本沿革

格式依循 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.1.0/)，版本號依循 [語意化版本](https://semver.org/lang/zh-TW/)：

- **主版號**：篩選規則或 Word 格式有不相容的改變（同一份資料會得到不同名單）
- **次版號**：新增功能，原有結果不變
- **修訂號**：修正錯誤

每筆紀錄標示更動位置（檔案）。完整逐行差異請用 `git log -p` 或 GitHub 的 Compare 頁面查看。

## [未發佈]

## [1.0.0] - 2026-10-26

GitHub 公開版首次發佈。

### 新增
- 免登入、可離線的獨立網頁，所有套件內含於 `vendor/`（`index.html`、`vendor/`）
- 變更紀錄頁：記錄每項更動的時間、版本、位置、項目、原值與新值，可搜尋、匯出 CSV（`js/app.js`）
- 設定檔匯出／匯入（JSON），含欄位對應、篩選設定、手動修改、入選清單與變更紀錄，可提交 Git 留存（`js/app.js`）
- 匯入試算表時記錄檔名、列數與 SHA-256 指紋，供追溯資料來源（`js/app.js`）
- 頁尾顯示版本號（`js/version.js`）
- 虛構範例試算表（`examples/`、`scripts/make-example.js`）
- 自動測試、CI 檢查、GitHub Pages 自動部署（`test/`、`.github/workflows/`）
- 版本號一致性與前端套件完整性檢查（`scripts/check-version.js`、`scripts/verify-vendor.js`）
- 範本產生腳本，並清除 Word 範本中的作者、公司資訊（`scripts/build-template.py`、`template/`）

### 變更
- 字型改用系統字型（微軟正黑體／蘋方），不再從 Google Fonts 載入（`css/style.css`）
- Word 匯出改用瀏覽器下載，不再依賴 claude.ai 下載功能（`js/app.js`）

## [0.1.0] - 2026-10-26

claude.ai 上的原型（未納入本倉庫 Git 歷史）。

### 新增
- 匯入試算表、自動欄位對應、類別優先順序與數量設定、業主指定類別、篩選清冊預覽與修改、依原 Word 範本匯出

[未發佈]: https://github.com/OWNER/REPO/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/OWNER/REPO/releases/tag/v1.0.0
