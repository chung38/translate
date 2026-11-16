# 系統架構：移次搞定＋移件翻

整個服務由兩個 repo 組成：

- **chung38/translate**（這個）負責網站與翻譯 API，放在 Cloudflare Pages。
- **chung38/migrant-poster-hub** 負責爬蟲與資料，放在 GitHub Actions 加 GitHub Pages。

```
                        使用者打開 https://yicipass.com
                                   │
┌──────────────────────────────────┴───────────────────────────────────┐
│ Cloudflare Pages（chung38/translate）                                 │
│                                                                      │
│  /               移次搞定外殼 index.html                              │
│                  首頁、文宣、聘僱、新聞、工具、搜尋                    │
│                  └─ 首頁中間用 iframe 嵌入 /translate/                │
│  /translate/     移件翻（React，程式在 src/）                         │
│  /view.html      文宣檢視頁（public/view.html）                       │
│  /api/*          後端（functions/api/，共用程式在 functions/_lib/）   │
│                  先驗證 Firebase 登入，再轉給 OpenAI 翻譯              │
└───────┬───────────────────────────┬──────────────────────┬───────────┘
        │ 讀 JSON 與縮圖              │ 登入、翻譯紀錄          │ 翻譯
        ▼                           ▼                      ▼
 chung38.github.io/            Firebase                 OpenAI API
 migrant-poster-hub/           Auth + Firestore         （gpt-6-luna）
 （GitHub Actions 每天爬）
```

## 網址

- **正式網址 `yicipass.com`**：網域買在 Cloudflare，在 translate → Custom domains 接到這個網站。
- **自動轉址**：`www.yicipass.com` 和舊網址 `translate-1up.pages.dev` 會自動轉到 `yicipass.com`（`index.html`、`public/view.html` 開頭的轉址程式）。舊網址會先確認新網址連得上才轉。
- **Google 登入**：網址要列在 Firebase Console → Authentication → Settings → Authorized domains，換網址或加網址時記得一起加，不然新網址登入會失敗。
- **之後如果再換網址**：要改的地方有 `index.html`、`public/view.html` 的轉址程式，以及資料 repo 的 `docs/index.html`、`docs/view.html`。

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
| 翻譯 API：登入驗證、檢查工作單、呼叫翻譯模型 | `functions/api/translate.ts` |
| 額度：開始翻譯扣額度、查剩餘、管理員調整與記錄付款 | `functions/api/jobs.ts`、`me.ts`、`admin/credits.ts`，規則在 `functions/_lib/credits.ts` |
| 翻譯模型與參數 | `functions/_lib/model.ts` |
| 服務條款、隱私權政策 | `public/terms.html`、`public/privacy.html`（網址是 /terms、/privacy） |
| 網站瀏覽人數 | `functions/api/visit.ts`（後端）、`index.html` 的「瀏覽人數」段落（顯示） |
| 登入、管理員、歷史紀錄等視窗 | `src/components/` |
| Firebase 連線設定 | `firebase-applet-config.json`、`src/firebase.ts` |
| Firestore 權限規則 | `firestore.rules`：改完要到 Firebase Console → Firestore → 規則，貼上後按「發布」，不會自動生效 |

## 金鑰與設定放在哪

| 項目 | 放在哪 | 說明 |
|---|---|---|
| `OPENAI_API_KEY` | Cloudflare → Workers & Pages → translate → Settings → Variables and Secrets（舊介面叫 Environment variables） | 翻譯用的金鑰（OpenAI），請設成加密（Secret）。換翻譯服務商時要一起改隱私權政策 |
| `OPENAI_MODEL` | 同上，選填 | 換別的 OpenAI 模型，預設 `gpt-6-luna` |
| `FIREBASE_PROJECT_ID` | 同上 | `ai-studio-applet-webapp-5cf4b`，後端用它驗證登入 |
| `ALLOWED_ORIGIN` | 同上，選填 | 允許哪些「其他網站」呼叫 API，多個用逗號分隔。自己網站（yicipass.com、舊網址）一律允許，所以一般不用設 |
| `DAILY_LIMIT` | 同上，選填 | 每人每天最多呼叫翻譯 API 幾次，預設 2000 |
| D1 資料庫 `DB` | Cloudflare → translate → Settings → Bindings → Add → D1 database，變數名稱填 `DB`（舊介面在 Settings → Functions） | 記錄額度、付款紀錄、每日翻譯次數、瀏覽人數。沒接的話額度退回舊方式（Firestore）、不限次數、不顯示人數 |
| `ADMIN_EMAILS` | Variables and Secrets | 管理員的 Email，多個用逗號分隔。後端只認這裡；沒設時暫時沿用 Firestore 的角色（使用者有機會自己改，請務必設定） |
| `FREE_QUOTA` | 同上，選填 | 新帳號免費幾份，預設 2。改了要一起改 `public/terms.html` 第六點 |
| `VISIT_SALT` | Variables and Secrets，選填 | 計算瀏覽人數時打亂訪客資料用的隨機字串，隨便打一串英數字即可 |
| Firebase 前端設定（apiKey 等） | `firebase-applet-config.json` | 這是公開設定，放在程式裡沒關係。資料安全靠 `firestore.rules` |
| 誰是管理員 | Firebase Console → Firestore → `users/{uid}` 的 `role` 欄位設成 `admin` | 前端不能自己改成管理員，規則會擋 |
| 爬蟲推資料的權限 | 資料 repo → Settings → Actions → General → Workflow permissions 選 Read and write | 用 GitHub 內建的權杖，不用另外建 token |

## 上線方式

- **這個 repo**：推到 `main`，Cloudflare 會自動執行 `npm run build`，把 `dist/` 上線，約 1–2 分鐘完成。推到其他分支會產生預覽網址；預覽環境沒有設金鑰，所以預覽站的翻譯會失敗，這是正常的。
- **資料 repo**：Actions 依排程自動跑。資料推上去後，GitHub Pages 約 1 分鐘更新，網站下次開啟就會讀到，Cloudflare 不需要重新建置。

## 本機測試（選用）

```bash
npm install
npm test           # 自動測試（61 項：翻譯邏輯、額度與工作單、次數上限、瀏覽人數）
npm run lint       # 型別檢查
cp .dev.vars.example .dev.vars   # 填入 OPENAI_API_KEY
npm run preview    # 跟正式站一樣（含翻譯 API）：http://localhost:8788
npm run dev        # 只改畫面、要即時更新時用：http://localhost:5173
                   # /api 會轉給 8788，所以要同時開著 npm run preview
```

`npm run preview` 用的是 Cloudflare 官方工具 wrangler，跑的就是 `functions/` 裡同一份後端，不用另外維護一份本機伺服器。

## 自動檢查

- 爬蟲每次跑完，`crawler/check-data.js` 會檢查資料格式。不合格就保留上一版，網站不會出現空白；該次 Action 會標成失敗，GitHub 會寄信通知。
- 在 GitHub 網頁上修改 `docs/guides.json`、`docs/info.json`、`sources.json` 或 `categories.json`，存檔後 `check.yml` 會自動檢查。少了逗號、引號這類錯誤會變紅色並寄信。

## 翻譯模型

目前用 OpenAI 的 **gpt-6-luna**，設定在 `functions/_lib/model.ts` 的 `pickProvider`：

- **推理關掉**（`reasoning_effort: none`）：翻譯不需要模型「思考」，開著會變慢，推理用的 token 也照輸出計費。
- **溫度 0.3**：讓同一句話每次翻得比較一致。
- **只回 JSON**、單次最多輸出 8192 tokens。
- 如果 OpenAI 回報不接受某個參數（例如換了不支援溫度的模型），後端會自動拿掉那個參數重送，並記住，之後不再重試。
- 想看實際用的是哪個模型：瀏覽器開發者工具 → Network → `translate` → Response Headers 的 `X-Translate-Model`。
- 費用：每次呼叫約台幣 0.02 元，56 頁的掃描 PDF 約 3 元。用量在 OpenAI 後台（platform.openai.com → Usage）查，建議在 Limits 設每月上限。

## 額度與付款紀錄

額度記在 Cloudflare D1，使用者碰不到：

- **計費單位**：上傳 1 份檔案算 1 份，不管選幾種語言、輸出成幾個檔案。新帳號免費 2 份（`FREE_QUOTA`）。
- **扣額度的時機**：每份檔案按下「開始翻譯」時，網頁先呼叫 `/api/jobs`，後端扣 1 份並發一張「工作單」。這份檔案之後的每一批翻譯都要帶著工作單，後端才會翻；所以繞過網頁直接呼叫翻譯 API 沒有用。
- **工作單限制**：只限本人、12 小時內、最多 800 次呼叫（56 頁掃描 PDF 約 100～150 次）。翻譯中途失敗時，重試同一份檔案沿用同一張工作單，不會重複扣。
- **管理員**：`ADMIN_EMAILS` 裡的人翻譯不扣額度、不受每日上限。
- **舊帳號搬家**：每個人第一次使用新版時，自動用他在 Firestore 的舊額度（`quota`）和翻譯紀錄筆數建立帳戶。上線後請到管理後台看一次大家的額度是否正確。
- **管理後台**：每位用戶旁邊的「調整」可以「記錄付款」（份數＋金額＋備註）、「加額度」或「設定總額度」，下面會列出這個人所有的額度變動紀錄。付款功能串好之後，付款成功也會寫進同一份紀錄。
- **資料表**：`accounts`（每人的總額度與已用份數）、`ledger`（每一筆變動）、`jobs`（工作單，7 天後自動清掉）。在 D1 的 Console 可以直接查，例如：
  `SELECT a.email, a.quota, a.used FROM accounts a ORDER BY a.updated_at DESC;`

## 翻譯 API 的保護

每次呼叫 `/api/translate` 依序檢查：

1. **有登入、Email 已驗證**。
2. **帳號沒被停用**（管理員按「刪除」會先停用）。
3. **帶著有效的工作單**（見上一段）。沒接 D1 時退回舊方式：比對 Firestore 的額度與翻譯紀錄筆數。
4. **每人每天的次數上限**（預設 2000 次，台灣時間 0 點重新計算，管理員不限）。這是防止有人寫程式一直呼叫、讓翻譯費用暴增的最後一道防線。次數記在 D1 的 `api_usage` 資料表，90 天前的紀錄會自動清掉。

想看每個人的每日用量：D1 → 你的資料庫 → Console，輸入
`SELECT * FROM api_usage ORDER BY day DESC, n DESC LIMIT 50;`

## 服務條款與隱私權政策

- 網址：`/terms`、`/privacy`，檔案在 `public/terms.html`、`public/privacy.html`。
- 連結出現在：翻譯頁「開始翻譯」下方、登入視窗、工具頁最下面。
- 聯絡信箱寫的是 `service@yicipass.com`：請在 Cloudflare → yicipass.com → Email → Email Routing 把它轉寄到你的信箱。
- 內容有改（例如免費份數、退費規則、換翻譯服務商）時，要記得一起更新這兩頁和上面的「最後更新」日期。

## 網站瀏覽人數

標頭右上角顯示「👀 今日 N 人・累計 M 人次」，由 `functions/api/visit.ts` 計算，存在 D1 資料庫：

- **今日人數**：同一天、同一個網路（IP）＋同一個瀏覽器只算 1 人，台灣時間 0 點歸零。重新整理、切換分頁都不會重複算。
- **累計人次**：每人每天算一次，一直往上加。
- 首頁和文宣檢視頁（`view.html`）共用同一個計數；首頁中間的翻譯頁不另外算。
- 搜尋引擎、社群連結預覽等機器人不算；從別的網站直接呼叫 API 也不算。
- 不存 IP、不用 cookie：只存「日期＋IP＋瀏覽器」打亂後的代碼，用來判斷今天是否來過，隔天就刪。
- 同一個辦公室或宿舍共用網路、又用同款手機的人，可能被算成同一人，所以數字會略低於實際。

想看每天的明細：D1 → 你的資料庫 → Console，輸入

```sql
SELECT substr(k, 10) AS 日期, n AS 人數 FROM site_stats WHERE k LIKE 'visitors:2%' ORDER BY k DESC LIMIT 30;
SELECT substr(k, 7) AS 日期, n AS 開啟次數 FROM site_stats WHERE k LIKE 'views:2%' ORDER BY k DESC LIMIT 30;
```

不想公開顯示、只要自己看：把 `index.html` 裡「瀏覽人數」那段的 `el.hidden = false;` 刪掉即可，後台照樣記錄。

## 目前已知的缺口

- **刪除使用者是「標記後下次登入時刪除」。** 管理員按「刪除」會先把帳號標記為待刪除，翻譯 API 立刻拒絕他；對方下次登入時，網頁會自動刪掉他的登入帳號與資料。如果對方再也不登入，Firebase 的登入帳號會一直留著，要清掉請到 Firebase Console → Authentication 手動刪。
- **線上付款還沒串。** 目前收款後由管理員在後台「記錄付款」加額度。LINE Pay、藍新金流只有舊的前端程式（`UpgradeModal.tsx`），沒有後端，畫面上也沒有按鈕會打開它；要串金流時，付款成功的通知要由後端寫進 `ledger`，並在結帳前讓使用者勾選同意（見服務條款第六點）。

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
functions/api/translate.ts       翻譯 API
functions/api/jobs.ts            開始翻譯一份檔案（扣額度、發工作單）
functions/api/me.ts              目前登入者的額度
functions/api/admin/credits.ts   管理員：額度、付款紀錄
functions/api/visit.ts           網站瀏覽人數 API
functions/_lib/                  後端共用程式（登入驗證、額度、模型、次數上限）
public/terms.html、privacy.html  服務條款、隱私權政策
src/components/CreditControls.tsx 管理後台的額度調整
firestore.rules                  Firestore 權限規則
firebase-blueprint.json          Firestore 資料欄位說明
tools/opc_check.py               檢查輸出的 Office 檔會不會被說「需要修復」
types/pages-function.d.ts        Cloudflare 型別宣告（只給 npm run lint 用）
.dev.vars.example                本機測試用的金鑰範本
```
