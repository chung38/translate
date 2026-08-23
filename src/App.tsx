/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { saveAs } from 'file-saver';
import { 
  Upload, 
  FileText, 
  Languages, 
  Download, 
  Loader2, 
  RefreshCw,
  CheckCircle2, 
  AlertCircle,
  ArrowRight,
  FileSpreadsheet,
  File as FileIcon,
  Presentation,
  History,
  LogIn,
  LogOut,
  User as UserIcon,
  Clock,
  Shield,
  Search,
  X,
  Settings
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { TranslationProgress } from './components/TranslationProgress';
import { HistoryPanel } from './components/HistoryPanel';
import { AuthModal } from './components/AuthModal';
import { AdminPanel } from './components/AdminPanel';
import { UpgradeModal } from './components/UpgradeModal';
import { DeletedModal } from './components/DeletedModal';
import { UserProfile } from './types';
import { useTranslation } from './hooks/useTranslation';
import { processDocx, processExcel, processPdf, processPptx } from './utils/documentProcessors';
import type { PptxLayoutMode } from './utils/documentProcessors';
import { OutputPreview } from './components/OutputPreview';
import { 
  auth, 
  db, 
  googleProvider, 
  signInWithPopup,
  getRedirectResult,
  onAuthStateChanged, 
  doc, 
  getDoc, 
  setDoc, 
  deleteDoc,
  collection, 
  addDoc, 
  query, 
  orderBy, 
  onSnapshot, 
  getDocs,
  handleFirestoreError,
  OperationType,
  Timestamp,
  where,
  limit,
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  updateProfile,
  sendEmailVerification,
  deleteUser as deleteAuthUser
} from './firebase';
import type { User } from './firebase';

// Error Boundary Component (Placeholder for functional compatibility)
const ErrorBoundary: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  return <>{children}</>;
};

// 標題的譯文（跟語言選擇連動，讓首屏本身就是一個雙語範例）
const HERO_SUBTITLE: Record<string, string> = {
  '越南文': 'Trạm dịch tài liệu',
  '泰文': 'สถานีแปลเอกสาร',
  '印尼文': 'Stasiun terjemahan dokumen',
  '英文': 'Document translation workbench',
  '繁體中文': '文件翻譯工作台',
};

const AVAILABLE_LANGUAGES = [
  { id: 'zh-TW', name: '繁體中文', label: 'Traditional Chinese', flag: '🇹🇼' },
  { id: 'th', name: '泰文', label: 'Thai', flag: '🇹🇭' },
  { id: 'id', name: '印尼文', label: 'Indonesian', flag: '🇮🇩' },
  { id: 'vi', name: '越南文', label: 'Vietnamese', flag: '🇻🇳' },
  { id: 'en', name: '英文', label: 'English', flag: '🇺🇸' },
];

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [userProfile, setUserProfile] = useState<UserProfile | null>(null);
  const [isAuthReady, setIsAuthReady] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [showAdminPanel, setShowAdminPanel] = useState(false);
  const [showUpgradeModal, setShowUpgradeModal] = useState(false);
  const [showAuthModal, setShowAuthModal] = useState(false);
  const [showDeletedModal, setShowDeletedModal] = useState(false);
  const [isPaying, setIsPaying] = useState(false);
  const [pendingOrderId, setPendingOrderId] = useState<string | null>(null);

  const [dbHistory, setDbHistory] = useState<any[]>([]);
  const [files, setFiles] = useState<File[]>([]);
  const [processingFiles, setProcessingFiles] = useState<File[]>([]);
  const [selectedLanguages, setSelectedLanguages] = useState<string[]>(['英文']);
  const [outputMode, setOutputMode] = useState<'combined' | 'separate'>('combined');
  // PPTX 版面模式：append = 同一頁雙語對照；duplicate-slide = 另外插一頁純譯文
  const [pptxLayoutMode, setPptxLayoutMode] = useState<PptxLayoutMode>('append');
  const [industry, setIndustry] = useState('');
  const [uploadProgress, setUploadProgress] = useState<Record<string, number>>({});
  const [uploadStatus, setUploadStatus] = useState<Record<string, 'uploading' | 'success' | 'error'>>({});
  const fileInputRef = useRef<HTMLInputElement>(null);

  const {
    status,
    setStatus,
    statusMessage,
    setStatusMessage,
    progress,
    setProgress,
    fileProgress,
    setFileProgress,
    error,
    setError,
    isCancelledRef,
    saveToFirestore,
    translateBatch,
    cancelTranslation
  } = useTranslation(user);

  const [reloadCounter, setReloadCounter] = useState(0);

  // 手機 Google redirect 登入結果處理（必須在最頂層，獨立於 AuthModal）
  useEffect(() => {
    const handleRedirectResult = async () => {
      try {
        const result = await getRedirectResult(auth);
        if (result?.user) {
          // redirect 登入成功，onAuthStateChanged 會自動觸發更新 user 狀態
          setShowAuthModal(false);
          console.log('Mobile Google redirect login success:', result.user.email);
        }
      } catch (err: any) {
        if (err.code === 'auth/popup-closed-by-user' || err.code === 'auth/cancelled-popup-request') {
          return;
        }
        console.error('Google redirect result error:', err);
        if (err.code === 'auth/unauthorized-domain') {
          setError('此網域未授權 Google 登入，請聯繫管理員在 Firebase Console 新增授權網域');
        } else {
          setError('Google 登入失敗，請再試一次');
        }
      }
    };
    handleRedirectResult();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Firebase Auth Listener
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      if (currentUser && !currentUser.emailVerified) {
        try {
          await currentUser.reload();
          setReloadCounter(c => c + 1);
        } catch (e) {
          console.error("Failed to reload user", e);
        }
      }
      setUser(auth.currentUser);
      if (!auth.currentUser) {
        setUserProfile(null);
        setIsAuthReady(true);
      }
    });
    return () => unsubscribe();
  }, []);

  // Poll for email verification status
  useEffect(() => {
    let intervalId: NodeJS.Timeout;
    if (user && !user.emailVerified) {
      intervalId = setInterval(async () => {
        try {
          await user.reload();
          if (auth.currentUser?.emailVerified) {
            setReloadCounter(c => c + 1);
            clearInterval(intervalId);
          }
        } catch (e) {
          console.error("Failed to poll user reload", e);
        }
      }, 3000);
    }
    return () => {
      if (intervalId) clearInterval(intervalId);
    };
  }, [user, reloadCounter]);

  // Profile Listener
  useEffect(() => {
    if (!user) return;

    const userRef = doc(db, 'users', user.uid);
    let unsubscribeSnapshot: (() => void) | null = null;

    const initProfile = async () => {
      try {
        const userSnap = await getDoc(userRef);
        // 管理員身分改由 Firestore 的 users/{uid}.role 決定，不再用寫死的 email 判斷

        if (!userSnap.exists()) {
          const q = query(collection(db, 'users'), where('email', '==', user.email), limit(1));
          const emailQuerySnap = await getDocs(q);
          
          let existingData: Partial<UserProfile> = {};
          if (!emailQuerySnap.empty && user.emailVerified) {
            const oldDoc = emailQuerySnap.docs[0];
            const oldData = oldDoc.data() as UserProfile;
            
            if (oldData.role !== undefined) existingData.role = oldData.role;
            if (oldData.isPaid !== undefined) existingData.isPaid = oldData.isPaid;
            if (oldData.quota !== undefined) existingData.quota = oldData.quota;
            if (oldData.displayName || user.displayName) existingData.displayName = oldData.displayName || user.displayName || undefined;
            if (oldData.photoURL || user.photoURL) existingData.photoURL = oldData.photoURL || user.photoURL || undefined;
            
            if (oldDoc.id !== user.uid) {
              try {
                await deleteDoc(doc(db, 'users', oldDoc.id));
              } catch (e) {
                console.error("Failed to delete old profile:", e);
              }
            }
          }

          // 注意：...existingData 必須放在最前面。
          // 原本放在最後，會把下面算好的 role / isPaid / quota / displayName
          // 又整個蓋回舊值，導致管理員自動升級對搬移過來的帳號永遠失效。
          const newProfile: UserProfile = {
            ...existingData,
            uid: user.uid,
            email: user.email || null,
            emailVerified: user.emailVerified,
            displayName: existingData.displayName || user.displayName || (user.email ? user.email.split('@')[0] : '未命名用戶'),
            photoURL: existingData.photoURL || user.photoURL || null,
            createdAt: Timestamp.now(),
            role: existingData.role || 'user',
            isPaid: existingData.isPaid || false,
            quota: existingData.quota !== undefined ? existingData.quota : 2,
          };
          
          Object.keys(newProfile).forEach(key => {
            if (newProfile[key as keyof UserProfile] === undefined) {
              delete newProfile[key as keyof UserProfile];
            }
          });

          await setDoc(userRef, newProfile);
          setUserProfile(newProfile);
        } else {
          const currentData = userSnap.data() as UserProfile;
          
          if (currentData.isPendingDeletion) {
            try {
              if (auth.currentUser) {
                await deleteAuthUser(auth.currentUser);
              }
              await deleteDoc(userRef);
            } catch (e) {
              console.error("Failed to delete user auth record:", e);
              await auth.signOut();
            }
            setShowDeletedModal(true);
            setUser(null);
            setUserProfile(null);
            setIsAuthReady(true);
            return;
          }
          
          let updatedData = { ...currentData };
          let needsUpdate = false;

          if (currentData.emailVerified !== user.emailVerified) {
            updatedData.emailVerified = user.emailVerified;
            needsUpdate = true;
          }

          if (needsUpdate) {
            await setDoc(userRef, updatedData, { merge: true });
          }
          
          setUserProfile(updatedData);
          // 管理員權限一律由 Firebase Console 或後端設定，前端不再自行寫入 role。
        }

        unsubscribeSnapshot = onSnapshot(userRef, async (doc) => {
          if (doc.exists()) {
            const data = doc.data() as UserProfile;
            
            if (data.isPendingDeletion) {
              try {
                if (auth.currentUser) {
                  await deleteAuthUser(auth.currentUser);
                }
                await deleteDoc(userRef);
              } catch (e) {
                console.error("Failed to delete user auth record:", e);
                await auth.signOut();
              }
              setShowDeletedModal(true);
              setUser(null);
              setUserProfile(null);
              return;
            }

            // 未完成 Email 驗證時，畫面上不給管理員權限（真正的權限以規則為準）
            if (data.role === 'admin' && !user.emailVerified) {
              setUserProfile({ ...data, role: 'user' });
            } else {
              setUserProfile(data);
            }
          } else {
            try {
              if (auth.currentUser) {
                await deleteAuthUser(auth.currentUser);
              }
            } catch (e) {
              console.error("Failed to delete user auth record:", e);
              await auth.signOut();
            }
            setShowDeletedModal(true);
            setUser(null);
            setUserProfile(null);
          }
        });

      } catch (error) {
        console.error("Profile initialization error:", error);
      }
      setIsAuthReady(true);
    };

    initProfile();

    return () => {
      if (unsubscribeSnapshot) unsubscribeSnapshot();
    };
  }, [user, reloadCounter]);

  // Firestore History Listener
  useEffect(() => {
    if (!user || !isAuthReady) {
      setDbHistory([]);
      return;
    }

    const historyRef = collection(db, 'users', user.uid, 'history');
    const q = query(historyRef, orderBy('timestamp', 'desc'));

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const items = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      setDbHistory(items);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, `users/${user.uid}/history`, user);
    });

    return () => unsubscribe();
  }, [user, isAuthReady]);

  // Listen for payment updates
  useEffect(() => {
    if (!user || !isPaying) return;
    
    if (pendingOrderId) {
      const paymentRef = doc(db, 'payments', pendingOrderId);
      const unsubscribe = onSnapshot(paymentRef, (doc) => {
        if (doc.exists() && doc.data().status === 'completed') {
          setPendingOrderId(null);
          setIsPaying(false);
          setStatus('completed');
          setStatusMessage('支付成功！額度已更新。');
        }
      }, (error) => {
        console.error("Specific payment listener error:", error);
      });
      return () => unsubscribe();
    }

    const fiveMinutesAgo = new Date(Date.now() - 5 * 60 * 1000);
    const paymentsRef = collection(db, 'payments');
    const q = query(
      paymentsRef, 
      where('userId', '==', user.uid), 
      where('status', '==', 'completed'),
      where('createdAt', '>=', Timestamp.fromDate(fiveMinutesAgo)),
      limit(1)
    );
    
    const unsubscribe = onSnapshot(q, (snapshot) => {
      if (!snapshot.empty) {
        setIsPaying(false);
        setStatus('completed');
        setStatusMessage('支付成功！額度已更新。');
      }
    }, (error) => {
      console.error("Fallback payment listener error:", error);
    });
    
    return () => unsubscribe();
  }, [user, isPaying, pendingOrderId]);

  const handleLogin = async () => {
    setError(null);
    setShowAuthModal(true);
  };

  const handleLogout = async () => {
    setError(null);
    await auth.signOut();
  };

  // Check for payment confirmation on mount
  useEffect(() => {
    const urlParams = new URLSearchParams(window.location.search);
    const transactionId = urlParams.get('transactionId');
    const orderId = urlParams.get('orderId');
    const statusParam = urlParams.get('status');
    const messageParam = urlParams.get('message');
    
    if (statusParam === 'success') {
      setStatus('completed');
      setStatusMessage('支付成功！額度已更新。');
      window.history.replaceState({}, document.title, window.location.pathname);
    } else if (statusParam === 'error') {
      setError(messageParam || '支付失敗');
      setStatus('error');
      window.history.replaceState({}, document.title, window.location.pathname);
    } else if (transactionId && orderId) {
      const confirmPayment = async () => {
        setStatus('processing');
        setStatusMessage('正在確認支付狀態...');
        try {
          const response = await fetch('/api/linepay/confirm', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ transactionId, orderId }),
          });
          const data = await response.json();
          if (data.success) {
            setStatus('completed');
            setStatusMessage('支付成功！額度已更新。');
            window.history.replaceState({}, document.title, window.location.pathname);
          } else {
            throw new Error(data.error || '支付確認失敗');
          }
        } catch (err: any) {
          setError(err.message);
          setStatus('error');
        }
      };
      confirmPayment();
    }
  }, []);

  const MAX_FILE_SIZE = 250 * 1024 * 1024;
  const MAX_FILE_COUNT = 10;

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(e.target.files || []) as File[];
    if (selectedFiles.length > 0) {
      if (files.length + selectedFiles.length > MAX_FILE_COUNT) {
        setError(`一次最多只能處理 ${MAX_FILE_COUNT} 個檔案`);
        e.target.value = '';
        return;
      }

      const validExtensions = ['docx', 'xlsx', 'pdf', 'pptx'];
      const legacyExtensions = ['doc', 'xls', 'ppt'];
      const newFiles: File[] = [];
      let hasInvalid = false;
      let hasLegacy = false;
      let hasOversized = false;

      selectedFiles.forEach(f => {
        const extension = f.name.split('.').pop()?.toLowerCase();
        if (legacyExtensions.includes(extension || '')) {
          hasLegacy = true;
        } else if (!validExtensions.includes(extension || '')) {
          hasInvalid = true;
        } else if (f.size > MAX_FILE_SIZE) {
          hasOversized = true;
        } else {
          newFiles.push(f);
        }
      });

      if (newFiles.length > 0) {
        setFiles(prev => {
          const existingNames = new Set(prev.map(f => f.name));
          const uniqueNewFiles = newFiles.filter(f => !existingNames.has(f.name));
          return [...prev, ...uniqueNewFiles];
        });
        
        if (hasLegacy) {
          setError('系統不支援舊版 .doc, .xls, .ppt 格式，請先使用 Office 另存為 .docx, .xlsx, .pptx 後再上傳。');
        } else if (hasOversized) {
          setError('部分檔案超過 250MB 限制，已跳過');
        } else if (hasInvalid) {
          setError('部分檔案格式不支援，已跳過');
        } else {
          setError(null);
        }
        
        setStatus('idle');

        newFiles.forEach(file => {
          setUploadStatus(prev => ({ ...prev, [file.name]: 'uploading' }));
          setUploadProgress(prev => ({ ...prev, [file.name]: 0 }));
          
          let progress = 0;
          const interval = setInterval(() => {
            progress += Math.random() * 30 + 10;
            if (progress >= 100) {
              progress = 100;
              clearInterval(interval);
              setUploadStatus(prev => ({ ...prev, [file.name]: 'success' }));
            }
            setUploadProgress(prev => ({ ...prev, [file.name]: progress }));
          }, 200);
        });

      } else if (hasLegacy) {
        setError('系統不支援舊版 .doc, .xls, .ppt 格式，請先使用 Office 另存為 .docx, .xlsx, .pptx 後再上傳。');
      } else if (hasOversized) {
        setError('檔案大小不能超過 250MB');
      } else if (hasInvalid) {
        setError('請上傳有效的 .docx, .xlsx, .pdf 或 .pptx 檔案');
      }
      e.target.value = '';
    }
  };

  const removeFile = (index: number) => {
    const fileToRemove = files[index];
    setFiles(prev => prev.filter((_, i) => i !== index));
    if (fileToRemove) {
      setUploadStatus(prev => {
        const next = { ...prev };
        delete next[fileToRemove.name];
        return next;
      });
      setUploadProgress(prev => {
        const next = { ...prev };
        delete next[fileToRemove.name];
        return next;
      });
    }
    if (files.length === 1) {
      setStatus('idle');
      setProgress(0);
    }
  };

  const toggleLanguage = (langName: string) => {
    setSelectedLanguages(prev => 
      prev.includes(langName) 
        ? prev.filter(l => l !== langName)
        : [...prev, langName]
    );
  };

  const processFiles = async () => {
    if (files.length === 0 || selectedLanguages.length === 0) return;
    
    if (!user) {
      setError('請先登入以使用翻譯功能');
      return;
    }

    try {
      await user.reload();
    } catch (e) {
      console.error("Failed to reload user", e);
    }

    if (!auth.currentUser?.emailVerified && !userProfile?.isManuallyAdded) {
      setError('請先至您的信笱收取驗證信並完成驗證，才可使用翻譯功能。系統會自動偵測您的驗證狀態。');
      return;
    }

    const isAdmin = userProfile?.role === 'admin';
    const quota = userProfile?.quota || 2;

    if (!isAdmin) {
      if (dbHistory.length >= quota) {
        setError(`額度已達上限 (${quota}份檔案)，請升級以繼續使用`);
        return;
      }
      if (dbHistory.length + files.length > quota) {
        setError(`剩餘額度 ${quota - dbHistory.length} 份，請減少上傳檔案數量或升級額度`);
        return;
      }
    }

    try {
      setStatus('processing');
      setError(null);
      setProgress(0);
      setFileProgress({});
      isCancelledRef.current = false;

      const filesToProcess = [...files];
      setProcessingFiles(filesToProcess);
      for (let i = 0; i < filesToProcess.length; i++) {
        if (isCancelledRef.current) break;
        
        const currentFile = filesToProcess[i];
        const extension = currentFile.name.split('.').pop()?.toLowerCase();
        
        setStatusMessage(`正在處理第 ${i + 1} / ${filesToProcess.length} 份文件: ${currentFile.name}`);
        
        const baseProgress = (i / filesToProcess.length) * 100;
        const fileWeight = 100 / filesToProcess.length;
        
        const updateFileProgress = (p: number, currentStatus?: TranslationStatus) => {
          setFileProgress(prev => ({ ...prev, [currentFile.name]: p }));
          setProgress(Math.round(baseProgress + (p * fileWeight / 100)));
          if (currentStatus) setStatus(currentStatus);
        };

        try {
          let results: { blob: Blob, name: string }[] = [];
          switch (extension) {
            case 'docx':
              results = await processDocx(currentFile, selectedLanguages, industry, translateBatch, updateFileProgress, isCancelledRef, outputMode);
              break;
            case 'xlsx':
              results = await processExcel(currentFile, selectedLanguages, industry, translateBatch, updateFileProgress, isCancelledRef, outputMode);
              break;
            case 'pdf':
              results = await processPdf(currentFile, selectedLanguages, industry, translateBatch, updateFileProgress, isCancelledRef, outputMode);
              break;
            case 'pptx':
              results = await processPptx(currentFile, selectedLanguages, industry, translateBatch, updateFileProgress, isCancelledRef, outputMode, pptxLayoutMode);
              break;
            default:
              console.warn(`不支援的檔案格式: ${currentFile.name}`);
          }
          
          if (results && results.length > 0) {
            for (const result of results) {
              saveAs(result.blob, result.name);
              await saveToFirestore(currentFile.name, result.name, extension || 'unknown', selectedLanguages, industry);
            }
          }

          setFiles(prev => prev.filter(f => f !== currentFile));
        } catch (fileErr) {
          console.error(`處理檔案 ${currentFile.name} 時發生錯誤:`, fileErr);
          setError(`處理 ${currentFile.name} 時發生錯誤: ${fileErr instanceof Error ? fileErr.message : '未知錯誤'}`);
        }
      }
      
      if (!isCancelledRef.current) {
        setStatus('completed');
        setStatusMessage('所有翻譯已完成！');
        setProgress(100);
      }
    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : '處理檔案時發生錯誤');
      setStatus('error');
      setStatusMessage('發生錯誤');
    }
  };

  const isUploading = files.some(f => uploadStatus[f.name] === 'uploading');

  return (
    <ErrorBoundary>
      <div className="app-shell h-screen overflow-auto font-sans">

        {/* Deleted Account Modal */}
        <DeletedModal 
          isOpen={showDeletedModal} 
          onClose={() => setShowDeletedModal(false)} 
        />

        {/* Auth Modal */}
        <AuthModal 
          isOpen={showAuthModal} 
          onClose={() => setShowAuthModal(false)} 
          onSuccess={(msg) => {
            if (msg) setError(msg);
          }}
        />

        {/* History Panel Overlay */}
        <HistoryPanel 
          isOpen={showHistory} 
          onClose={() => setShowHistory(false)} 
          dbHistory={dbHistory} 
        />

        {/* Upgrade Modal */}
        <UpgradeModal 
          isOpen={showUpgradeModal} 
          onClose={() => setShowUpgradeModal(false)} 
          user={user} 
          setStatus={setStatus} 
          setStatusMessage={setStatusMessage} 
          setError={setError} 
        />

        {/* Admin Panel */}
        <AdminPanel 
          isOpen={showAdminPanel} 
          onClose={() => setShowAdminPanel(false)} 
          user={user}
          userProfile={userProfile}
        />

        <div className="w-full max-w-[95vw] sm:max-w-2xl lg:max-w-6xl mx-auto p-3 sm:p-5 md:p-6 md:pt-10">
          {/* Header */}
          <header className="mb-7 md:mb-9">
            <div className="flex items-center justify-between gap-4 pb-3 mb-6 border-b border-[var(--rule)]">
              <div className="flex items-center gap-2.5 min-w-0">
                <span className="w-7 h-7 rounded-[7px] bg-[var(--ink)] text-white flex items-center justify-center text-[13px] font-bold leading-none shrink-0">文</span>
                <span className="text-[13px] font-semibold text-[var(--ink)] truncate">文件翻譯工作台</span>
              </div>

            {/* Auth Bar — 標題上方，非左對齊中間內容區塊 */}
            {isAuthReady && (
              <motion.div
                initial={{ opacity: 0, y: -6 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.05, duration: 0.4 }}
                className="flex items-center gap-2"
              >
                {user ? (
                  <div className="flex items-center gap-2">
                    <button 
                      onClick={() => setShowHistory(!showHistory)}
                      className="iconbtn"
                      title="翻譯紀錄"
                    >
                      <Clock className="w-5 h-5" />
                    </button>
                    {userProfile?.role === 'admin' && (
                      <button 
                        onClick={() => setShowAdminPanel(true)}
                        className="iconbtn"
                        title="管理後台"
                      >
                        <Shield className="w-5 h-5" />
                      </button>
                    )}
                    <div className="flex items-center gap-2 bg-[var(--card)] border border-[var(--rule)] px-3 py-1.5 rounded-[8px]">
                      {user.photoURL ? (
                        <img src={user.photoURL} alt={user.displayName || userProfile?.displayName || ""} className="w-6 h-6 rounded-full" referrerPolicy="no-referrer" />
                      ) : (
                        <UserIcon className="w-4 h-4 text-[var(--muted)]" />
                      )}
                      <span className="text-sm font-medium text-[var(--ink)] inline-block max-w-[120px] truncate sm:max-w-[200px]">
                        {user.displayName || userProfile?.displayName || user.email?.split('@')[0]}
                      </span>
                      {userProfile && (
                        <div className="flex items-center gap-1.5 ml-1">
                          {userProfile.role === 'admin' ? (
                            <span className="tag tag--signal">Admin</span>
                          ) : userProfile.isPaid ? (
                            <div className="flex items-center gap-1">
                              <span className="tag tag--signal">Pro</span>
                              <span className="num text-[10px] text-[var(--signal-ink)]">{userProfile.quota}次</span>
                            </div>
                          ) : (
                            <div className="flex items-center gap-1">
                              <span className="tag">
                                {userProfile.isPaid ? 'Pro' : 'Free'}
                              </span>
                              <span className="num text-[10px] text-[var(--muted)]">({dbHistory.length}/{userProfile.quota})</span>
                            </div>
                          )}
                        </div>
                      )}
                      {!user.emailVerified && !userProfile?.isManuallyAdded && (
                        <button 
                          onClick={async () => {
                            try {
                              await sendEmailVerification(user);
                              setError('驗證信已重新發送，請檢查您的信笱。');
                            } catch (e) {
                              setError('發送驗證信失敗，請稍後再試。');
                            }
                          }}
                          className="tag tag--brass ml-2 hover:brightness-95"
                        >
                          重發驗證信
                        </button>
                      )}
                      <button onClick={handleLogout} className="ml-2 text-[var(--muted)] hover:text-[var(--alert)] transition-colors">
                        <LogOut className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                ) : (
                  <button 
                    onClick={handleLogin}
                    className="flex items-center gap-2 bg-[var(--card)] border border-[var(--rule)] px-4 py-2 rounded-[8px] text-[var(--ink)] font-medium hover:border-[var(--signal)] transition-colors"
                  >
                    <LogIn className="w-4 h-4 text-[var(--signal)]" />
                    <span>登入</span>
                  </button>
                )}
              </motion.div>
            )}
            </div>

            {/* Title — 標題本身就用這個工具的輸出格式排：原文在上，譯文在下 */}
            <div>
              <motion.h1
                initial={{ opacity: 0, y: 8 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.4 }}
                className="bititle__src"
              >
                文件翻譯工作台
              </motion.h1>
              <div className="bititle__rule" />
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                transition={{ delay: 0.12, duration: 0.4 }}
                className="bititle__tgt"
              >
                {(selectedLanguages.length ? selectedLanguages : ['越南文']).map(l => (
                  <span key={l}>{HERO_SUBTITLE[l] ?? l}</span>
                ))}
              </motion.div>
              <p className="mt-3.5 text-sm text-[var(--muted)] max-w-lg leading-relaxed">
                譯文直接排進原檔，版面與表格保留。原檔已有的語言不會再翻一次。
              </p>
            </div>
          </header>

        {/* 工作台：左邊設定、右邊即時預覽 */}
        <div className="grid lg:grid-cols-[minmax(0,1fr)_360px] gap-5 items-start">
        <motion.div 
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.15 }}
          className="panel"
        >
          <div className="p-4 sm:p-6 md:p-7">
            {/* Error Message — 顯示在頂部 */}
            <AnimatePresence>
            {error && (
              <motion.div
                key="error-top"
                initial={{ opacity: 0, y: -8 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -8 }}
                className="note note--alert mb-4"
              >
                <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                <div className="flex-1">
                  <p className="font-medium">發生錯誤</p>
                  <p className="mt-0.5">{error}</p>
                </div>
                <button onClick={() => setError(null)} className="opacity-60 hover:opacity-100 transition-opacity ml-2 shrink-0">
                  <X className="w-4 h-4" />
                </button>
              </motion.div>
            )}
            </AnimatePresence>

            {/* Upload Section */}
            <div className="step">
              <span className="step__n">01</span>
              <span className="step__label">要翻譯的檔案</span>
            </div>
            <div 
              onClick={() => fileInputRef.current?.click()}
              className={`group ${files.length > 0 ? 'dropzone dropzone--filled' : 'dropzone'}`}
            >
              <input 
                type="file" 
                ref={fileInputRef}
                onChange={handleFileChange}
                accept=".docx,.xlsx,.pdf,.pptx,.doc,.xls,.ppt"
                multiple
                className="hidden"
              />
              
              <div className="flex items-center gap-4">
                <div className="filerow__icon">
                  <Upload className="w-5 h-5" />
                </div>
                <div>
                  <p className="text-sm font-medium text-[var(--ink)]">選擇檔案，或拖進來</p>
                  <p className="text-xs text-[var(--muted)] mt-0.5">Word、Excel、PowerPoint、PDF，可一次選多份</p>
                </div>
              </div>
            </div>

            {/* File List */}
            <AnimatePresence>
              {files.length > 0 && (
                <motion.div 
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="mt-6 space-y-2"
                >
                  {files.some(f => f.name.toLowerCase().endsWith('.pdf')) && (
                    <div className="note note--brass mb-4">
                      <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                      <div>
                        <p className="font-bold mb-1">PDF 翻譯注意事項</p>
                        <p>系統目前僅能擷取 PDF 中的「純文字」進行翻譯，將會**遺失原有的表格與排版**。若為揃描檔或圖片 PDF，系統會自動啟用 OCR (光學字元辨識) 進行處理，但辨識可能需要較長時間。若您的 PDF 包含表格或特殊字體（可能導致亂碼），強烈建議您先將其轉換為 Word 或 Excel 檔案後再進行翻譯。</p>
                      </div>
                    </div>
                  )}
                  <div className="flex justify-between items-center px-1 mb-2">
                    <span className="text-[13px] font-semibold text-[var(--ink)]">待處理檔案 ({files.length})</span>
                    {status === 'idle' && (
                      <button 
                        onClick={() => setFiles([])}
                        className="text-xs text-[var(--alert)] hover:underline"
                      >
                        全部清除
                      </button>
                    )}
                  </div>
                  {files.map((f, idx) => (
                    <motion.div 
                      key={`${f.name}-${idx}`}
                      initial={{ opacity: 0, x: -10 }}
                      animate={{ opacity: 1, x: 0 }}
                      exit={{ opacity: 0, x: 10 }}
                      className="filerow"
                    >
                      <div className="flex items-center gap-4 min-w-0">
                        <div className="filerow__icon">
                          {f.name.endsWith('.docx') && <FileText className="w-5 h-5" />}
                          {f.name.endsWith('.xlsx') && <FileSpreadsheet className="w-5 h-5" />}
                          {f.name.endsWith('.pdf') && <FileIcon className="w-5 h-5" />}
                          {f.name.endsWith('.pptx') && <Presentation className="w-5 h-5" />}
                        </div>
                        <div className="min-w-0">
                          <p className="text-sm font-medium text-[var(--ink)] truncate">{f.name}</p>
                          <p className="num text-[10px] text-[var(--muted)] mt-0.5">{(f.size / 1024).toFixed(1)} KB</p>
                        </div>
                      </div>
                      
                      <div className="flex items-center gap-4">
                        {status !== 'idle' && status !== 'error' && (
                          <div className="flex flex-col items-end gap-1.5">
                            <span className="num text-[10px] text-[var(--signal)]">
                              {fileProgress[f.name] || 0}%
                            </span>
                            <div className="meter">
                              <div 
                                className="meter__fill" 
                                style={{ width: `${fileProgress[f.name] || 0}%` }}
                              />
                            </div>
                          </div>
                        )}
                        
                        {status === 'idle' && uploadStatus[f.name] === 'uploading' && (
                          <div className="flex flex-col items-end gap-1.5">
                            <span className="num text-[10px] text-[var(--muted)]">
                              {Math.round(uploadProgress[f.name] || 0)}%
                            </span>
                            <div className="meter">
                              <div 
                                className="meter__fill" 
                                style={{ width: `${uploadProgress[f.name] || 0}%` }}
                              />
                            </div>
                          </div>
                        )}

                        {status === 'idle' && uploadStatus[f.name] === 'success' && (
                          <div className="flex items-center gap-2">
                            <span className="tag tag--signal flex items-center gap-1">
                              <CheckCircle2 className="w-3 h-3" />
                              上傳成功
                            </span>
                            <button 
                              onClick={() => removeFile(idx)}
                              className="p-1.5 rounded-[6px] text-[var(--muted)] hover:text-[var(--alert)] hover:bg-[var(--alert-dim)] transition-all"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          </div>
                        )}

                        {status === 'idle' && !uploadStatus[f.name] && (
                          <button 
                            onClick={() => removeFile(idx)}
                            className="p-1.5 rounded-[6px] text-[var(--muted)] hover:text-[var(--alert)] hover:bg-[var(--alert-dim)] transition-all"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        )}

                        {status === 'completed' && fileProgress[f.name] === 100 && (
                          <div className="w-7 h-7 rounded-full bg-[var(--signal-dim)] flex items-center justify-center">
                            <CheckCircle2 className="w-4 h-4 text-[var(--signal)]" />
                          </div>
                        )}
                      </div>
                    </motion.div>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>

            {/* Settings */}
            <div className="mt-8 space-y-7">
              <div>
                <div className="step">
                  <span className="step__n">02</span>
                  <span className="step__label">產業別<span className="step__hint">　選填，用來校準專有名詞</span></span>
                </div>
                <div className="relative">
                  <input
                    type="text"
                    value={industry}
                    onChange={(e) => setIndustry(e.target.value)}
                    placeholder="例如：電子、紡織、汽車..."
                    className="input"
                  />
                </div>
              </div>

              <div>
                <div className="step">
                  <span className="step__n">03</span>
                  <span className="step__label">翻成哪些語言<span className="step__hint">　可複選</span></span>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  {AVAILABLE_LANGUAGES.map((lang) => (
                    <button
                      key={lang.id}
                      onClick={() => toggleLanguage(lang.name)}
                      className={selectedLanguages.includes(lang.name) ? 'chip chip--on' : 'chip'}
                    >
                      <span className="flex items-center gap-2.5">
                        <span className="text-base">{lang.flag}</span>
                        <span>{lang.name}</span>
                      </span>
                      {selectedLanguages.includes(lang.name) && (
                        <CheckCircle2 className="w-4 h-4 text-[var(--signal)]" />
                      )}
                    </button>
                  ))}
                </div>
              </div>

              {selectedLanguages.length > 1 && (
                <div>
                  <div className="step">
                    <span className="step__n">04</span>
                    <span className="step__label">輸出成幾個檔案</span>
                  </div>
                  <div className="flex flex-col sm:flex-row gap-3">
                    <label className={outputMode === 'combined' ? 'optcard optcard--on' : 'optcard'}>
                      <input 
                        type="radio" 
                        name="outputMode" 
                        value="combined" 
                        checked={outputMode === 'combined'} 
                        onChange={() => setOutputMode('combined')}
                        className="w-4 h-4"
                      />
                      <span className="optcard__title">合併為單一檔案</span>
                    </label>
                    <label className={outputMode === 'separate' ? 'optcard optcard--on' : 'optcard'}>
                      <input 
                        type="radio" 
                        name="outputMode" 
                        value="separate" 
                        checked={outputMode === 'separate'} 
                        onChange={() => setOutputMode('separate')}
                        className="w-4 h-4"
                      />
                      <span className="optcard__title">分開為多個檔案</span>
                    </label>
                  </div>
                </div>
              )}

              {files.some(f => f.name.toLowerCase().endsWith('.pptx')) && (
                <div>
                  <div className="step">
                    <span className="step__n">05</span>
                    <span className="step__label">簡報版面<span className="step__hint">　只影響 PowerPoint</span></span>
                  </div>
                  <div className="flex flex-col sm:flex-row gap-3">
                    <label className={pptxLayoutMode === 'append' ? 'optcard optcard--on' : 'optcard'}>
                      <input
                        type="radio"
                        name="pptxLayoutMode"
                        value="append"
                        checked={pptxLayoutMode === 'append'}
                        onChange={() => setPptxLayoutMode('append')}
                        className="w-4 h-4"
                      />
                      <span className="optcard__title">
                        同頁對照
                        <span className="optcard__desc">
                          譯文接在原文下方，頁數不變。文字較多的頁面會自動縮小字級。
                        </span>
                      </span>
                    </label>
                    <label className={pptxLayoutMode === 'duplicate-slide' ? 'optcard optcard--on' : 'optcard'}>
                      <input
                        type="radio"
                        name="pptxLayoutMode"
                        value="duplicate-slide"
                        checked={pptxLayoutMode === 'duplicate-slide'}
                        onChange={() => setPptxLayoutMode('duplicate-slide')}
                        className="w-4 h-4"
                      />
                      <span className="optcard__title">
                        另加譯文頁
                        <span className="optcard__desc">
                          原稿頁保持不動，後面插入一頁純譯文。版面不變形，頁數會變兩倍。
                        </span>
                      </span>
                    </label>
                  </div>
                </div>
              )}
              
              <div className="flex flex-col gap-3 pt-6">
                {!user ? (
                  <button
                    onClick={handleLogin}
                    className="btn btn--primary"
                  >
                    <LogIn className="w-5 h-5 shrink-0" />
                    <span className="truncate">請先登入以開始翻譯</span>
                  </button>
                ) : (!user.emailVerified && !userProfile?.isManuallyAdded) ? (
                  <button
                    onClick={async () => {
                      try {
                        await user.reload();
                        if (auth.currentUser?.emailVerified) {
                          setReloadCounter(c => c + 1);
                        } else {
                          setError('請先至您的信笱收取驗證信並完成驗證，才可使用翻譯功能。系統會自動偵測您的驗證狀態。');
                        }
                      } catch (e) {
                        setError('請先至您的信笱收取驗證信並完成驗證，才可使用翻譯功能。');
                      }
                    }}
                    className="btn btn--brass"
                  >
                    <AlertCircle className="w-5 h-5 shrink-0" />
                    <span className="truncate">請先驗證 Email 以開始翻譯</span>
                  </button>
                ) : (
                  <button
                    disabled={files.length === 0 || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating' || isUploading}
                    onClick={processFiles}
                    className="btn btn--primary"
                  >
                    {status === 'idle' && (
                      <>
                        <span className="truncate">
                          {isUploading ? '檔案上傳中...' : `開始翻譯 ${files.length > 0 ? `${files.length} 份` : ''} ${selectedLanguages.length > 0 ? `(${selectedLanguages.length} 種語言)` : ''}`}
                        </span>
                        {!isUploading && <ArrowRight className="w-4 h-4 shrink-0" />}
                      </>
                    )}
                    {(status === 'processing' || status === 'translating' || status === 'generating') && (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin shrink-0" />
                        <span className="truncate">正在處理 {Object.keys(fileProgress).length} / {files.length}...</span>
                      </>
                    )}
                    {status === 'completed' && (
                      <>
                        <CheckCircle2 className="w-4 h-4" />
                        <span>全部翻譯完成</span>
                      </>
                    )}
                    {status === 'error' && (
                      <>
                        <AlertCircle className="w-4 h-4" />
                        <span>重試</span>
                      </>
                    )}
                  </button>
                )}
                
                {(status === 'processing' || status === 'translating' || status === 'generating') && (
                  <button
                    onClick={cancelTranslation}
                    className="btn btn--quiet"
                  >
                    取消翻譯
                  </button>
                )}
              </div>
            </div>

            {/* Progress & Status */}
            <TranslationProgress 
              status={status} 
              statusMessage={statusMessage} 
              progress={progress} 
              files={processingFiles}
              fileProgress={fileProgress}
            />

          </div>
        </motion.div>

        {/* 預覽欄：捲動時固定，讓選項一改就看得到結果 */}
        <motion.aside
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.22 }}
          className="lg:sticky lg:top-6 space-y-3"
        >
          <OutputPreview
            languages={selectedLanguages}
            layoutMode={pptxLayoutMode}
            hasPptx={files.some(f => f.name.toLowerCase().endsWith('.pptx'))}
          />
          <p className="text-xs text-[var(--muted)] leading-relaxed px-1">
            示意圖，實際字級會依原檔的版面自動調整。表格內的譯文會縮小，避免撐開欄高。
          </p>
        </motion.aside>
        </div>
      </div>
    </div>
    </ErrorBoundary>
  );
}
