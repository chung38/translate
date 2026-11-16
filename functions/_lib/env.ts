// Cloudflare 上的設定（Workers & Pages → translate → Settings → Variables and Secrets / Bindings）
export interface Env {
  FIREBASE_PROJECT_ID: string;   // 驗證登入用
  OPENAI_API_KEY?: string;       // 翻譯用（OpenAI，預設模型 gpt-6-luna）
  OPENAI_MODEL?: string;         // 選填：換別的 OpenAI 模型
  ALLOWED_ORIGIN?: string;       // 選填：允許哪些其他網站呼叫（逗號分隔）；自己網站一律允許
  ADMIN_EMAILS?: string;         // 管理員 Email（逗號分隔）。後端只認這裡，不認使用者自己能改的資料
  FREE_QUOTA?: string;           // 新帳號免費幾份檔案，預設 2
  DB?: D1Database;               // D1 資料庫（Bindings，名稱 DB）：額度、付款紀錄、每日次數、瀏覽人數
  DAILY_LIMIT?: string;          // 選填：每人每天最多呼叫翻譯幾次，預設 2000
  VISIT_SALT?: string;           // 選填：瀏覽人數用的隨機字串
}
