# 更動與版本紀錄規範

目標：**任何一次更動，都能查到改了哪裡、改了什麼、為什麼改、誰改的。**

## 一、程式更動（Git）

### 1. 以分支 + Pull Request 修改

```bash
git switch -c fix/崩塌率格式
# 修改檔案…
npm test
git add -A
git commit -m "fix(core): 崩塌率小於 0.01% 時顯示三位小數"
git push -u origin fix/崩塌率格式
```

在 GitHub 開 Pull Request，依範本勾選更動位置；CI 通過後再合併到 `main`。
建議在 **Settings → Branches** 為 `main` 開啟保護規則（需 PR、需 CI 通過），避免未留紀錄的直接修改。

### 2. 提交訊息格式

```
<類型>(<位置>): <更動項目>

<原因或補充說明（選填）>
```

| 類型 | 用途 | 位置 | 對應檔案 |
|---|---|---|---|
| `feat` | 新功能 | `core` | `js/core.js` 篩選邏輯、Word 匯出 |
| `fix` | 修正錯誤 | `ui` | `index.html`、`js/app.js`、`css/style.css` |
| `rule` | 篩選規則改變 | `template` | `template/`、`scripts/build-template.py` |
| `docs` | 文件 | `vendor` | `vendor/` |
| `chore` | 套件、設定、工具 | `ci` | `.github/` |
| `records` | 提交名單設定檔 | `records` | `records/*.json` |

### 3. 記錄到 CHANGELOG.md

每個 PR 都在 `## [未發佈]` 下加一行，並在括號標示檔案，例如：

```
### 修正
- 崩塌率小於 0.01% 時顯示三位小數（`js/core.js`）
```

### 4. 發佈新版

1. 把 `[未發佈]` 的內容移到新版本標題下，例如 `## [1.1.0] - 2026-11-03`
2. 同步修改 `package.json` 的 `version` 與 `js/version.js` 的 `APP_VERSION`
3. `npm run version:check` 確認三處一致
4. 提交並加標籤：
   ```bash
   git commit -am "chore: 發佈 v1.1.0"
   git tag -a v1.1.0 -m "v1.1.0"
   git push && git push --tags
   ```
5. 在 GitHub **Releases** 依標籤建立版本說明（可貼上 CHANGELOG 該段）

### 5. 如何追溯

> 檔名含中文時，先執行一次 `git config --global core.quotepath false`，`git log --stat` 才會正常顯示中文檔名。

| 想知道 | 指令／位置 |
|---|---|
| 每次提交改了哪些檔案 | `git log --stat` |
| 某檔案每行最後由誰、哪次提交修改 | `git blame js/core.js`，或 GitHub 檔案頁的 **Blame** |
| 某檔案的完整修改歷史與內容差異 | `git log -p -- js/core.js` |
| 兩個版本間的所有差異 | `git diff v1.0.0 v1.1.0`，或 GitHub `compare/v1.0.0...v1.1.0` |
| 某段文字何時被加入或刪除 | `git log -S "第3類" -p` |
| 網頁目前是哪一版 | 頁尾版本號；匯出的設定檔與變更紀錄每筆都記有版本 |

## 二、名單更動（網頁變更紀錄 + 設定檔）

網頁的「變更紀錄」頁自動記錄每項操作：

| 欄位 | 說明 | 範例 |
|---|---|---|
| 時間 | 民國年／月／日 時:分:秒 | 115/04/23 14:05:12 |
| 版本 | 操作當時的程式版本 | 1.0.0 |
| 更動位置 | 步驟與區塊（清冊修改會帶溪流編號） | 4 篩選清冊／投縣DF187 |
| 更動項目 | 欄位或動作 | 初步篩選辦理原因 |
| 原值 → 新值 | 更動前後內容 | 114年度颱風豪雨事件致災 → 陳情案 |

每次名單定稿（例如送審、期中、期末）時：

1. 在「變更紀錄」頁按 **匯出設定檔**，存到 `records/`，檔名建議 `records/115_期中_篩選設定.json`
2. 提交：`git commit -m "records: 115年期中現勘名單 150 條"`
3. 之後可用 `git diff` 比較兩版設定檔，看出名單、數量、順序與手動修改的差異；
   需要重現當時結果時，在網頁按 **匯入設定檔** 並匯入同一份試算表即可。

> 設定檔不含整份試算表，只含檔名、列數與 SHA-256 指紋。若要確認重現時用的是同一份資料，比對指紋即可。
