# 曝光報告外掛（Lark 多維表格）

在 Lark 多維表格「TW Marketing Promotion Hub」裡，選一個專案，按一下就產出這個專案的曝光報告 PPT，並存回專案那一列。

- 不需要 Lark 自建應用、不需要管理員審核、不用 AI 額度
- 程式在使用者自己的瀏覽器裡執行，用使用者自己的 Lark 權限讀資料
- 資料只在 Lark 和使用者瀏覽器之間傳，不會經過 GitHub 或任何伺服器

## 報告內容

1. 封面：專案名稱、曝光數、已有截圖數、已有連結數、POC、優惠期間
2. 曝光總覽：所有曝光的渠道、時間、狀態、截圖和連結是否已提供
3. 每一則勾選「要截圖」的曝光一頁：渠道、時間、狀態、備註、截圖、連結、渠道 Owner
   還沒有截圖或連結的，會先標示「待補」

## 檔案說明

| 檔案 | 用途 | 要不要改 |
|---|---|---|
| `index.html` | 外掛本體，Lark 讀的就是這個 | 不要直接改，由 build 產生 |
| `config.json` | 欄位名稱、報告名稱、每頁截圖數等設定 | **可以直接在 GitHub 網頁上改** |
| `src/main.js` | 程式原始碼 | 要改功能時才動 |
| `src/shell.html` | 外掛畫面的版型和樣式 | 要改外觀時才動 |
| `build.mjs`、`package.json` | 把原始碼打包成 `index.html` 的工具 | 不用改 |

## 常見修改：改 config.json（不用寫程式）

表格欄位改名、想換報告名稱、想調整每頁截圖數時：

1. 在 GitHub 打開 `config.json`，點右上角的鉛筆圖示
2. 只改冒號右邊、引號裡的文字，例如把 `"曝光報告"` 改成 `"交功課報告"`
3. 按「Commit changes」
4. 約 1 分鐘後，在 Lark 重新整理頁面，再打開外掛就生效

注意：引號、逗號、冒號都要保留。改壞了外掛會自動改用預設值，不會壞掉，但設定不會生效。

## 改功能：改原始碼

請用 AI（例如 Claude）幫忙：

1. 把 `src/main.js`、`src/shell.html` 和你想要的改動告訴 AI，請它修改
2. 在電腦上安裝 Node.js，在這個資料夾執行：
   ```
   npm install
   npm run build
   ```
3. 會產生新的 `index.html`，連同改過的 `src` 檔案一起上傳到 GitHub 取代舊的

## 搬到另一個 GitHub 帳號

1. 新帳號建立一個 Public repository，例如 `exposure-report`
2. 把這個 repository 的所有檔案上傳過去（Add file → Upload files，可以整個資料夾拖進去）
3. Settings → Pages → Branch 選 `main`、資料夾選 `/ (root)` → Save
4. 新網址會是 `https://新帳號名稱.github.io/exposure-report/`
5. 在 Lark 多維表格：拼圖圖示 → 自訂外掛程式 → 編輯原本的外掛，把服務 URL 換成新網址
6. 確認新網址能用之後，舊帳號的 repository 再刪除

## 安全

- 請替擁有這個 repository 的 GitHub 帳號開啟兩步驟驗證
- 不要隨意把別人加成協作者：能修改這裡的人，就能改變外掛在每個使用者 Lark 裡做的事
