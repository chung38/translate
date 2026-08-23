# 檔案放置說明

解開之後的資料夾結構，就是 repo 的結構。整包對著 repo 根目錄覆蓋即可。

```
translate-更新檔/
├── .gitignore                                    → .gitignore
├── firestore.rules                               → firestore.rules
├── index.html                                    → index.html
├── src/
│   ├── App.tsx                                   → src/App.tsx
│   ├── index.css                                 → src/index.css
│   ├── components/
│   │   ├── OutputPreview.tsx                     → src/components/OutputPreview.tsx   ★新檔案
│   │   ├── AdminPanel.tsx                        → src/components/AdminPanel.tsx
│   │   ├── AuthModal.tsx                         → src/components/AuthModal.tsx
│   │   ├── DeletedModal.tsx                      → src/components/DeletedModal.tsx
│   │   ├── HistoryPanel.tsx                      → src/components/HistoryPanel.tsx
│   │   ├── TranslationHistory.tsx                → src/components/TranslationHistory.tsx
│   │   ├── TranslationProgress.tsx               → src/components/TranslationProgress.tsx
│   │   └── UpgradeModal.tsx                      → src/components/UpgradeModal.tsx
│   └── utils/
│       ├── documentProcessors.ts                 → src/utils/documentProcessors.ts
│       └── __tests__/
│           └── documentProcessors.test.ts        → src/utils/__tests__/documentProcessors.test.ts   ★新檔案
└── tools/
    └── opc_check.py                              → tools/opc_check.py   ★新檔案（檢查工具，不影響建置）
```

---

## 覆蓋前要先刪一個檔案

```
src/utils/documentProcessors.test.ts     ← 刪掉這個
```

這一份是 `documentProcessors.ts` 的舊版複本（1879 行），檔名被存成 `.test.ts`。
vitest 會把它當測試檔跑，然後失敗：

```
FAIL  src/utils/documentProcessors.test.ts
Error: No test suite found in file
```

真正的測試檔在 `src/utils/__tests__/documentProcessors.test.ts`。

---

## 這幾個檔案不用動（你已經更新過了）

- `package.json`、`package-lock.json`
- `server.ts`
- `functions/api/translate.ts`
- `src/firebase.ts`
- `src/hooks/useTranslation.ts`

---

## firestore.rules 要另外部署

這個檔案放進 repo 不會自動生效，要另外推到 Firebase：

```bash
firebase deploy --only firestore:rules
```

或到 Firebase Console → Firestore → 規則貼上發布。

**發布前務必先做這兩件事**，否則你會失去管理員權限：

1. 到 Console 把 `users/{你的uid}` 的 `role` 手動設成 `admin`
2. 用 Console 的「規則 Playground」試一次讀寫

檔案末尾的註解列出套用後會受影響的三處程式碼，記得一起看。

---

## 覆蓋完的驗證步驟

```bash
npm install        # 沒動 package.json，通常不需要
npm test           # 應該是 24 passed
npm run lint       # tsc，剩 3 個錯誤都是原本就有的
npx vite build     # 應該成功
```

剩下那 3 個既有錯誤：兩個是 `PagesFunction` 缺型別定義（Cloudflare 的全域型別），
一個是 pdf.js 的 `RenderParameters`。都不影響建置。
