import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from 'uuid';
import fs from 'fs';
import PQueue from 'p-queue';
import rateLimit from 'express-rate-limit';
import { initializeApp } from 'firebase/app';
import { 
  getFirestore, 
  collection, 
  doc, 
  setDoc, 
  getDoc, 
  updateDoc, 
  runTransaction, 
  serverTimestamp 
} from 'firebase/firestore';
import { initializeApp as initializeAdminApp } from 'firebase-admin/app';
import { getAuth as getAdminAuth } from 'firebase-admin/auth';
import { getFirestore as getAdminFirestore } from 'firebase-admin/firestore';

console.log("SERVER STARTING UP...");

// Load Firebase config
const firebaseConfig = JSON.parse(fs.readFileSync(path.join(process.cwd(), 'firebase-applet-config.json'), 'utf8'));
console.log("Firebase Config Loaded:", { ...firebaseConfig, apiKey: "REDACTED" });

// Initialize Firebase Client SDK on server
const firebaseApp = initializeApp(firebaseConfig);
const db = getFirestore(firebaseApp, firebaseConfig.firestoreDatabaseId);

// Initialize Firebase Admin
let adminApp;
let adminDb: any;
try {
  adminApp = initializeAdminApp({
    projectId: firebaseConfig.projectId,
  });
  adminDb = getAdminFirestore(adminApp, firebaseConfig.firestoreDatabaseId);
  console.log("Firebase Admin initialized");
} catch (error) {
  console.error("Firebase Admin initialization error:", error);
}

// Test Firestore connection on startup
async function testFirestore() {
  try {
    console.log("Testing Firestore connection to database:", firebaseConfig.firestoreDatabaseId);
    const testRef = doc(db, 'server_status', 'last_start');
    await setDoc(testRef, {
      timestamp: serverTimestamp(),
      message: "Server started",
      projectId: firebaseConfig.projectId
    });
    console.log("Firestore connection test successful");
  } catch (err) {
    console.error("Firestore connection test failed:", err);
  }
}
testFirestore();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  
  // Trust the first proxy (e.g., Google Cloud Run / Nginx) to correctly populate req.ip from X-Forwarded-For
  app.set('trust proxy', 1);
  
  const PORT = Number(process.env.PORT) || 3000;

  // 限制請求的 payload 大小，避免傳送過大的檔案或內容
  app.use(express.json({ limit: '100mb' }));
  app.use(express.urlencoded({ extended: true, limit: '100mb' }));

  // 設定 API 請求次數限制 (Rate Limiting)
  // 限制每個 IP 在 15 分鐘內最多只能發送 100 次請求
  // ── 登入驗證 middleware ────────────────────────────────────────────────
  // 沒有這一段，任何人都能直接 POST /api/translate，用你的 DeepSeek 金鑰。
  const requireAuth = async (req: any, res: any, next: any) => {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: { message: '請先登入後再使用翻譯功能。' } });
    }
    try {
      const decoded = await getAdminAuth().verifyIdToken(authHeader.split('Bearer ')[1]);
      if (!decoded.email_verified) {
        return res.status(403).json({ error: { message: '請先完成 Email 驗證。' } });
      }
      req.user = decoded;
      next();
    } catch {
      return res.status(401).json({ error: { message: '登入狀態已失效，請重新登入。' } });
    }
  };

  const apiLimiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 100,
    message: { error: { message: '請求次數過多，請稍後再試。' } },
    standardHeaders: true,
    legacyHeaders: false,
    keyGenerator: (req: any) => {
      // 已登入就用 uid 當 key（同一間工廠常常共用同一個對外 IP）
      if (req.user?.uid) return `uid:${req.user.uid}`;
      // 處理 Forwarded header 警告
      const forwarded = req.headers['forwarded'];
      if (forwarded && typeof forwarded === 'string') {
        const match = forwarded.match(/for="?([^;"]+)"?/);
        if (match) return match[1];
      }
      // Fallback to Express's req.ip (which uses X-Forwarded-For because of trust proxy)
      return req.ip || req.socket.remoteAddress || 'unknown';
    }
  });

  // Request logger
  app.use((req, res, next) => {
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
    next();
  });

  // Catch-all API logger
  app.all("/api/*", (req, res, next) => {
    console.log(`API Request: ${req.method} ${req.url}`);
    next();
  });

  // ─── DeepSeek 全域並發控制 ────────────────────────────────────────────────
  // concurrency: 3   → 同時最多 3 個請求打向 DeepSeek（防止 429）
  // intervalCap: 6   → 每秒最多觸發 6 次（3 concurrency × 2，保守設定）
  // interval: 1000   → 計算窗口 1 秒
  // MAX_QUEUE_SIZE   → queue 超過此值直接回 503，讓前端 retry backoff
  //                    避免 queue 無限堆積、記憶體暴漲
  const MAX_QUEUE_SIZE = 20;
  const translationQueue = new PQueue({ concurrency: 3, intervalCap: 6, interval: 1000 });

  // Admin API: Delete user immediately
  app.delete('/api/admin/users/:uid', async (req, res) => {
    try {
      const { uid } = req.params;
      const authHeader = req.headers.authorization;
      
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({ error: 'Unauthorized' });
      }
      
      const idToken = authHeader.split('Bearer ')[1];
      const decodedToken = await getAdminAuth().verifyIdToken(idToken);
      
      // Check if the requester is an admin
      let isAdmin = false;
      const isDefaultAdmin = decodedToken.email === 'chen.chung.shih@gmail.com';
      
      try {
        const adminDoc = await adminDb.collection('users').doc(decodedToken.uid).get();
        isAdmin = adminDoc.exists && adminDoc.data()?.role === 'admin';
      } catch (adminCheckError: any) {
        if (!adminCheckError.message?.includes('PERMISSION_DENIED')) {
          console.warn(`Failed to check admin status in Firestore: ${adminCheckError.message}`);
        }
        // If we can't check Firestore (e.g. PERMISSION_DENIED), we rely solely on isDefaultAdmin
      }
      
      if (!isAdmin && !isDefaultAdmin) {
        return res.status(403).json({ error: 'Forbidden: Admin access required' });
      }

      console.log(`Admin ${decodedToken.email} is deleting user ${uid}`);

      // Delete from Firebase Auth
      let authDeleted = false;
      try {
        await getAdminAuth().deleteUser(uid);
        console.log(`Successfully deleted user ${uid} from Firebase Auth`);
        authDeleted = true;
      } catch (authError: any) {
        if (authError.code === 'auth/user-not-found') {
          console.log(`User ${uid} not found in Firebase Auth, proceeding to delete from Firestore`);
          authDeleted = true;
        } else if (authError.message?.includes('Identity Toolkit API has not been used')) {
          console.log(`[Preview Environment] Skipping backend Auth deletion due to missing API permissions. Proceeding with soft delete.`);
        } else {
          console.warn(`Failed to delete user from Firebase Auth: ${authError.message}. Proceeding with soft delete.`);
        }
      }

      let firestoreDeleted = false;
      // Delete from Firestore
      try {
        const userRef = adminDb.collection('users').doc(uid);
        
        if (authDeleted) {
          // Hard delete history subcollection
          const historySnapshot = await userRef.collection('history').get();
          if (!historySnapshot.empty) {
            const batch = adminDb.batch();
            historySnapshot.docs.forEach((doc: any) => {
              batch.delete(doc.ref);
            });
            await batch.commit();
          }
          
          // Hard delete user document
          await userRef.delete();
          console.log(`Successfully deleted user ${uid} and their history from Firestore`);
        } else {
          // Soft delete user document so client SDK can delete Auth on next login
          await userRef.set({ isPendingDeletion: true }, { merge: true });
          console.log(`Soft deleted user ${uid} in Firestore`);
        }
        firestoreDeleted = true;
      } catch (firestoreError: any) {
        if (firestoreError.code === 5 || firestoreError.message?.includes('NOT_FOUND')) {
          console.log(`User ${uid} or history not found in Firestore, proceeding`);
          firestoreDeleted = true;
        } else if (firestoreError.code === 7 || firestoreError.message?.includes('PERMISSION_DENIED')) {
          console.log(`[Preview Environment] Skipping backend Firestore deletion due to missing permissions. Proceeding with client SDK fallback.`);
        } else {
          throw firestoreError;
        }
      }

      res.json({ success: true, firestoreDeleted });
    } catch (error: any) {
      console.error('Error deleting user:', error);
      res.status(500).json({ error: error.message || 'Failed to delete user' });
    }
  });

  // DeepSeek Proxy API
  app.post("/api/translate", requireAuth, apiLimiter, async (req, res) => {
    const { prompt } = req.body;
    const apiKey = process.env.DEEPSEEK_API_KEY;

    // 驗證 prompt 長度，避免過大的文本導致後端或 API 崩潰
    if (typeof prompt !== 'string') {
      return res.status(400).json({ error: { message: "無效的請求內容" } });
    }
    if (prompt.length > 200000) { // 限制約 20 萬字元
      return res.status(400).json({ error: { message: "文本內容過長，超過系統單次處理限制。" } });
    }

    if (!apiKey) {
      return res.status(500).json({ 
        error: { 
          message: "請先在 AI Studio 的 Settings 選單中配置 DEEPSEEK_API_KEY" 
        } 
      });
    }

    // ── Queue 滿載保護 ───────────────────────────────────────────────────────
    // size  = 等待中（尚未開始執行）的請求數
    // pending = 正在執行中的請求數
    // 當 waiting 超過 MAX_QUEUE_SIZE，拒絕加入，讓前端 retry backoff
    const waitingCount = translationQueue.size;
    if (waitingCount >= MAX_QUEUE_SIZE) {
      console.warn(`[Queue] OVERLOADED. Waiting: ${waitingCount}, Pending: ${translationQueue.pending}. Returning 503.`);
      return res.status(503).json({
        error: {
          message: `伺服器目前繁忙（佇列已滿 ${waitingCount}/${MAX_QUEUE_SIZE}），請稍後自動重試。`
        }
      });
    }

    console.log(`[Queue] Enqueued. Waiting: ${waitingCount + 1}, Pending: ${translationQueue.pending}`);

    try {
      // Add the request to the queue
      const data = await translationQueue.add(async () => {
        console.log(`[Queue] Starting translation request. Queue size: ${translationQueue.size}, Pending: ${translationQueue.pending}`);
        
        let retries = 3;
        let delay = 1000;
        
        while (retries > 0) {
          try {
            const response = await fetch("https://api.deepseek.com/v1/chat/completions", {
              method: "POST",
              headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
              },
              body: JSON.stringify({
                model: "deepseek-chat",
                messages: [
                  { role: "system", content: "You are a helpful assistant that translates text into multiple languages and outputs only structured JSON." },
                  { role: "user", content: prompt }
                ],
                response_format: { type: "json_object" },
                temperature: 0.3,
                max_tokens: 8192
              })
            });

            const responseText = await response.text();
            
            if (!response.ok) {
              if ((response.status === 429 || response.status >= 500) && retries > 1) {
                console.log(`[Queue] DeepSeek API Error (${response.status}). Retrying in ${delay}ms...`);
                await new Promise(resolve => setTimeout(resolve, delay));
                retries--;
                delay *= 2;
                continue;
              }
              
              let errorData;
              try {
                errorData = JSON.parse(responseText);
              } catch (e) {
                errorData = { error: { message: `DeepSeek API Error (${response.status}): ${responseText.substring(0, 200)}` } };
              }
              throw { status: response.status, data: errorData };
            }

            try {
              return JSON.parse(responseText);
            } catch (e) {
              throw new Error(`DeepSeek API returned invalid JSON: ${responseText.substring(0, 200)}`);
            }
          } catch (fetchError: any) {
            if (retries > 1 && !fetchError.status) {
              console.log(`[Queue] Network Error: ${fetchError.message}. Retrying in ${delay}ms...`);
              await new Promise(resolve => setTimeout(resolve, delay));
              retries--;
              delay *= 2;
              continue;
            }
            throw fetchError;
          }
        }
      });

      res.json(data);
    } catch (err: any) {
      console.error("DeepSeek Proxy Error:", err);
      if (err.status && err.data) {
        return res.status(err.status).json(err.data);
      }
      res.status(500).json({ error: { message: err.message || "Internal Server Error" } });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on port ${PORT}`);
  });
}

startServer();
