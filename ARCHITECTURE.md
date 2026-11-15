# 系統架構：移次搞定＋移件翻

整個服務由兩個 repo 組成：

- **chung38/translate**（這個）負責網站與翻譯 API，放在 Cloudflare Pages。
- **chung38/migrant-poster-hub** 負責爬蟲與資料，放在 GitHub Actions 加 GitHub Pages。

```
                   使用者打開 https://translate-1up.pages.dev
                                   │
┌──────────────────────────────────┴───────────────────────────────────┐
│ Cloudflare Pages（chung38/translate）                                 │
│                                                                      │
│  /               移次搞定外殼 index.html                              │
│                  首頁、文宣、聘僱、新聞、工具、搜尋                    │
│                  └─ 首頁中間用 iframe 嵌入 /translate/                │
│  /translate/     移件翻（React，程式在 src/）                         │
│  /view.html      文宣檢視頁（public/view.html）                       │
│  /api/translate  唯一的後端（functions/api/translate.ts）             │
│                  先驗證 Firebase 登入，再轉給 DeepSeek 翻譯            │
└───────┬───────────────────────────┬──────────────────────┬───────────┘
        │ 讀 JSON 與縮圖              │ 登入、翻譯紀錄          │ 翻譯
        ▼                           ▼                      ▼
 chung38.github.io/            Firebase                 DeepSeek API
 migrant-poster-hub/           Auth + Firestore
 （GitHub Actions 每天爬）
```

## 為什麼分成兩個 repo

- **省 Cloudflare 建置次數。** 爬蟲一天會推資料約 10 次。放在同一個 repo 的話，每次推送都會觸發 Cloudflare 重新建置，一個月約 300 次，免費方案上限是 500 次。
- **爬蟲免費。** 資料 repo 是公開的，GitHub Actions 在公開 repo 不收費。
- **改資料不用重新部署。** 網站每次開啟都直接讀 GitHub Pages 上最新的 JSON。

## 想改什麼，就改哪個檔

| 想改的東西 | 檔案 |
|---|---|
| 首頁、文宣、聘僱、新聞、工具、搜尋各分頁的畫面 | `index.html`（全部在這一個檔） |
| 薪資／加班費計算機的算法 | `index.html`；費率在資料 repo 的 `docs/info.json` → `payroll` |
| 聘僱指南內容、申請應備文件 | 資料 repo 的 `docs/guides.json` |
| 緊急電話、各國辦事處、節日 | 資料 repo 的 `docs/info.json` |
| 文宣要抓哪些政府網站 | 資料 repo 的 `sources.json` |
| 文宣分類與關鍵字 | 資料 repo 的 `categories.json` |
| 翻譯頁畫面（三步驟） | `src/App.tsx`、`src/index.css` |
| Word／Excel／PPT／PDF 怎麼拆字、怎麼把譯文放回去 | `src/utils/documentProcessors.ts` |
| 給 AI 的翻譯指令、一次送幾句、譯文標籤修復 | `src/hooks/useTranslation.ts` |
| 後端：登入驗證、呼叫 DeepSeek | `functions/api/translate.ts` |
| 登入、管理員、歷史紀錄等視窗 | `src/components/` |
| Firebase 連線設定 | `firebase-applet-config.json`、`src/firebase.ts` |
| Firestore 權限規則 | `firestore.rules`：改完要到 Firebase Console → Firestore → 規則，貼上後按「發布」，不會自動生效 |

## 金鑰與設定放在哪

| 項目 | 放在哪 | 說明 |
|---|---|---|
| `DEEPSEEK_API_KEY` | Cloudflare → Workers & Pages → translate → Settings → Variables and Secrets（舊介面叫 Environment variables） | 翻譯用的金鑰，請設成加密（Secret） |
| `FIREBASE_PROJECT_ID` | 同上 | `ai-studio-applet-webapp-5cf4b`，後端用它驗證登入 |
| `ALLOWED_ORIGIN` | 同上，選填 | 只允許這個網址呼叫 API。只設在 Production，不要設在 Preview |
| `DAILY_LIMIT` | 同上，選填 | 每人每天最多呼叫翻譯 API 幾次，預設 2000 |
| D1 資料庫 `DB` | Cloudflare → translate → Settings → Bindings → Add → D1 database，變數名稱填 `DB`（舊介面在 Settings → Functions） | 記錄每人每天的次數。沒接的話不限制次數，其他功能照常 |
| Firebase 前端設定（apiKey 等） | `firebase-applet-config.json` | 這是公開設定，放在程式裡沒關係。資料安全靠 `firestore.rules` |
| 誰是管理員 | Firebase Console → Firestore → `users/{uid}` 的 `role` 欄位設成 `admin` | 前端不能自己改成管理員，規則會擋 |
| 爬蟲推資料的權限 | 資料 repo → Settings → Actions → General → Workflow permissions 選 Read and write | 用 GitHub 內建的權杖，不用另外建 token |

## 上線方式

- **這個 repo**：推到 `main`，Cloudflare 會自動執行 `npm run build`，把 `dist/` 上線，約 1–2 分鐘完成。推到其他分支會產生預覽網址；預覽環境沒有設金鑰，所以預覽站的翻譯會失敗，這是正常的。
- **資料 repo**：Actions 依排程自動跑。資料推上去後，GitHub Pages 約 1 分鐘更新，網站下次開啟就會讀到，Cloudflare 不需要重新建置。

## 本機測試（選用）

```bash
npm install
npm test           # 自動測試（45 項：翻譯邏輯＋API 的額度與次數上限）
npm run lint       # 型別檢查
cp .dev.vars.example .dev.vars   # 填入 DEEPSEEK_API_KEY
npm run preview    # 跟正式站一樣（含翻譯 API）：http://localhost:8788
npm run dev        # 只改畫面、要即時更新時用：http://localhost:5173
                   # /api 會轉給 8788，所以要同時開著 npm run preview
```

`npm run preview` 用的是 Cloudflare 官方工具 wrangler，跑的就是 `functions/` 裡同一份後端，不用另外維護一份本機伺服器。

## 自動檢查

- 爬蟲每次跑完，`crawler/check-data.js` 會檢查資料格式。不合格就保留上一版，網站不會出現空白；該次 Action 會標成失敗，GitHub 會寄信通知。
- 在 GitHub 網頁上修改 `docs/guides.json`、`docs/info.json`、`sources.json` 或 `categories.json`，存檔後 `check.yml` 會自動檢查。少了逗號、引號這類錯誤會變紅色並寄信。

## 翻譯 API 的保護

每次呼叫 `/api/translate` 依序檢查：

1. **有登入、Email 已驗證**，否則拒絕。
2. **帳號沒被停用、檔案額度還沒用完**。後端用使用者自己的登入 token 讀他在 Firestore 的資料（`role`、`quota`、`isPendingDeletion`、翻譯紀錄筆數），規則跟網頁上一樣，管理員不受額度限制。所以就算繞過網頁直接呼叫 API，也一樣會被擋。讀不到資料時放行，不影響正常使用。
3. **每人每天的次數上限**（預設 2000 次，台灣時間 0 點重新計算，管理員不限）。一份文件會分成很多批送出，每批算 1 次：一般 Word 或 PPT 約 10～50 次，56 頁的掃描 PDF 約 100～150 次。這是防止有人寫程式一直呼叫、讓 DeepSeek 帳單暴增的最後一道防線。次數記在 D1 資料庫的 `api_usage` 資料表，90 天前的紀錄會自動清掉。

想看每個人的用量：Cloudflare 左側選單的 D1 → 點你的資料庫 → Console，輸入
`SELECT * FROM api_usage ORDER BY day DESC, n DESC LIMIT 50;`

> 第 2 項的額度能不能被繞過，取決於 Firebase 上實際生效的 Firestore 規則：規則必須禁止使用者自己改 `quota`、`role`（repo 裡的 `firestore.rules` 有寫，但要在 Firebase Console 發布才算數）。第 3 項存在 Cloudflare，使用者碰不到。

## 目前已知的缺口

- **刪除使用者是「標記後下次登入時刪除」。** 管理員按「刪除」會先把帳號標記為待刪除，翻譯 API 立刻拒絕他；對方下次登入時，網頁會自動刪掉他的登入帳號與資料。如果對方再也不登入，Firebase 的登入帳號會一直留著，要清掉請到 Firebase Console → Authentication 手動刪。
- **付費升級沒有後端。** LINE Pay 和藍新金流只有前端程式（`UpgradeModal.tsx`、`App.tsx`），`/api/linepay/*`、`/api/newebpay/*` 都不存在，畫面上也沒有按鈕會打開它。要收費時再補。

## 檔案地圖

```
index.html                       移次搞定外殼（所有分頁、計算機、搜尋）
translate/index.html             移件翻入口 → src/main.tsx
public/view.html                 文宣檢視頁
src/App.tsx                      翻譯頁畫面
src/index.css                    樣式
src/hooks/useTranslation.ts      呼叫 /api/translate、分批、標籤修復
src/utils/documentProcessors.ts  各種檔案格式的處理
src/utils/__tests__/             自動測試
src/components/                  登入、管理員、歷史紀錄等視窗
src/firebase.ts                  Firebase 初始化
functions/api/translate.ts       後端（唯一一份）
firestore.rules                  Firestore 權限規則
firebase-blueprint.json          Firestore 資料欄位說明
tools/opc_check.py               檢查輸出的 Office 檔會不會被說「需要修復」
types/pages-function.d.ts        Cloudflare 型別宣告（只給 npm run lint 用）
.dev.vars.example                本機測試用的金鑰範本
```
