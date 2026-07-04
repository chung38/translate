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
        const isAdminEmail = user.email === 'chen.chung.shih@gmail.com';

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

          const newProfile: UserProfile = {
            uid: user.uid,
            email: user.email || null,
            emailVerified: user.emailVerified,
            displayName: user.displayName || (user.email ? user.email.split('@')[0] : '未命名用戶'),
            photoURL: user.photoURL || null,
            createdAt: Timestamp.now(),
            role: (isAdminEmail && user.emailVerified) ? 'admin' : (existingData.role || 'user'),
            isPaid: existingData.isPaid || false,
            quota: existingData.quota !== undefined ? existingData.quota : 2,
            ...existingData
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

          if (isAdminEmail && !user.emailVerified && currentData.role === 'admin') {
            updatedData.role = 'user';
            needsUpdate = true;
          }
          
          if (currentData.emailVerified !== user.emailVerified) {
            updatedData.emailVerified = user.emailVerified;
            needsUpdate = true;
          }

          if (needsUpdate) {
            await setDoc(userRef, updatedData, { merge: true });
          }
          
          setUserProfile(updatedData);
          
          if (isAdminEmail && user.emailVerified && currentData.role !== 'admin') {
            await setDoc(userRef, { role: 'admin' }, { merge: true });
          }
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

            if (user.email === 'chen.chung.shih@gmail.com') {
              if (user.emailVerified) {
                setUserProfile(data);
              } else {
                setUserProfile({ ...data, role: 'user' });
              }
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

    const isAdmin = userProfile?.role === 'admin' || user?.email === 'chen.chung.shih@gmail.com';
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
              results = await processPptx(currentFile, selectedLanguages, industry, translateBatch, updateFileProgress, isCancelledRef, outputMode);
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
      <div className="min-h-screen h-screen overflow-auto bg-slate-50 font-sans text-slate-800 relative">

        {/* Background Accents */}
        <div className="absolute inset-0 bg-[radial-gradient(ellipse_at_top_right,_var(--tw-gradient-stops))] from-indigo-100/40 via-slate-50 to-slate-50 -z-10" />
        <div className="absolute top-[-10%] right-[-5%] w-[500px] h-[500px] bg-indigo-400/10 blur-[100px] rounded-full -z-10 pointer-events-none" />
        <div className="absolute bottom-[-10%] left-[-5%] w-[500px] h-[500px] bg-violet-400/10 blur-[100px] rounded-full -z-10 pointer-events-none" />

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

        <div className="w-full max-w-[95vw] sm:max-w-2xl lg:max-w-4xl mx-auto p-3 sm:p-5 md:p-6 md:pt-10 relative z-10">
          {/* Header */}
          <header className="mb-8 md:mb-12">

            {/* Auth Bar — 標題上方，非左對齊中間內容區塊 */}
            {isAuthReady && (
              <motion.div
                initial={{ opacity: 0, y: -6 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.05, duration: 0.4 }}
                className="flex justify-start mb-3"
              >
                {user ? (
                  <div className="flex items-center gap-2">
                    <button 
                      onClick={() => setShowHistory(!showHistory)}
                      className="p-2 bg-white/70 backdrop-blur-md rounded-full shadow-sm border border-white text-slate-500 hover:text-indigo-600 transition-colors"
                      title="翻譯紀錄"
                    >
                      <Clock className="w-5 h-5" />
                    </button>
                    {userProfile?.role === 'admin' && (
                      <button 
                        onClick={() => setShowAdminPanel(true)}
                        className="p-2 bg-white/70 backdrop-blur-md rounded-full shadow-sm border border-white text-slate-500 hover:text-violet-600 transition-colors"
                        title="管理後台"
                      >
                        <Shield className="w-5 h-5" />
                      </button>
                    )}
                    <div className="flex items-center gap-2 bg-white/70 backdrop-blur-md px-3 py-1.5 rounded-full shadow-sm border border-white">
                      {user.photoURL ? (
                        <img src={user.photoURL} alt={user.displayName || userProfile?.displayName || ""} className="w-6 h-6 rounded-full" referrerPolicy="no-referrer" />
                      ) : (
                        <UserIcon className="w-4 h-4 text-slate-400" />
                      )}
                      <span className="text-sm font-medium text-slate-800 inline-block max-w-[120px] truncate sm:max-w-[200px]">
                        {user.displayName || userProfile?.displayName || user.email?.split('@')[0]}
                      </span>
                      {userProfile && (
                        <div className="flex items-center gap-1.5 ml-1">
                          {(userProfile.role === 'admin' || user?.email === 'chen.chung.shih@gmail.com') ? (
                            <span className="px-2 py-0.5 bg-violet-100 border border-violet-200 text-violet-700 text-[10px] font-bold rounded-full uppercase tracking-wider">Admin</span>
                          ) : userProfile.isPaid ? (
                            <div className="flex items-center gap-1">
                              <span className="px-2 py-0.5 bg-indigo-100 border border-indigo-200 text-indigo-700 text-[10px] font-bold rounded-full uppercase tracking-wider">Pro</span>
                              <span className="text-[10px] text-indigo-700 font-bold">{userProfile.quota}次</span>
                            </div>
                          ) : (
                            <div className="flex items-center gap-1">
                              <span className="px-2 py-0.5 bg-slate-100 border border-slate-200 text-slate-500 text-[10px] font-bold rounded-full uppercase tracking-wider">
                                {userProfile.isPaid ? 'Pro' : 'Free'}
                              </span>
                              <span className="text-[10px] text-slate-500 font-medium">({dbHistory.length}/{userProfile.quota})</span>
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
                          className="ml-2 px-2 py-0.5 bg-amber-100 border border-amber-200 text-amber-700 text-[9px] font-bold rounded-full hover:bg-amber-200 transition-colors"
                        >
                          重發驗證信
                        </button>
                      )}
                      <button onClick={handleLogout} className="ml-2 text-slate-400 hover:text-red-500 transition-colors">
                        <LogOut className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                ) : (
                  <button 
                    onClick={handleLogin}
                    className="flex items-center gap-2 bg-white/70 backdrop-blur-md px-4 py-2 rounded-full shadow-sm border border-white text-slate-800 font-medium hover:bg-white transition-colors"
                  >
                    <LogIn className="w-4 h-4 text-indigo-600" />
                    <span>登入</span>
                  </button>
                )}
              </motion.div>
            )}

            {/* Title */}
            <div className="space-y-2 text-center">
              <motion.h1 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.1, duration: 0.5 }}
                className="text-3xl md:text-5xl font-bold tracking-tight text-slate-900 flex items-center justify-center gap-4"
              >
                <div className="relative flex-shrink-0">
                  <div className="absolute inset-0 bg-indigo-200 blur-xl rounded-full" />
                  <div className="relative w-12 h-12 md:w-14 md:h-14 bg-gradient-to-tr from-indigo-600 to-violet-600 rounded-2xl shadow-[0_8px_20px_-6px_rgba(79,70,229,0.5)] flex items-center justify-center transform -rotate-3 hover:rotate-0 transition-all duration-300">
                    <Languages className="w-7 h-7 md:w-8 md:h-8 text-white" />
                  </div>
                </div>
                <span className="font-tech">全能文件<span className="text-transparent bg-clip-text bg-gradient-to-r from-indigo-600 to-violet-600">多語翻譯器</span></span>
              </motion.h1>

              <motion.p
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                transition={{ delay: 0.2, duration: 0.5 }}
                className="text-indigo-600/80 text-sm md:text-base font-medium tracking-wide uppercase"
              >
                Professional AI-powered document translation
              </motion.p>
            </div>
          </header>

        {/* Main Card */}
        <motion.div 
          initial={{ opacity: 0, y: 30 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.3 }}
          className="bg-white/90 backdrop-blur-2xl rounded-3xl shadow-xl shadow-slate-200/50 border border-white overflow-hidden"
        >
          <div className="p-4 sm:p-6 md:p-8">
            {/* Error Message — 顯示在頂部 */}
            <AnimatePresence>
            {error && (
              <motion.div
                key="error-top"
                initial={{ opacity: 0, y: -8 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -8 }}
                className="mb-4 p-3 bg-red-50 border border-red-200 rounded-xl flex items-start gap-3 text-red-600 text-sm"
              >
                <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                <div className="flex-1">
                  <p className="font-medium">發生錯誤</p>
                  <p className="text-red-500 mt-0.5">{error}</p>
                </div>
                <button onClick={() => setError(null)} className="text-red-400 hover:text-red-600 transition-colors ml-2 flex-shrink-0">
                  <X className="w-4 h-4" />
                </button>
              </motion.div>
            )}
            </AnimatePresence>

            {/* Upload Section — 縮小版 */}
            <div 
              onClick={() => fileInputRef.current?.click()}
              className={`
                relative group cursor-pointer border-2 border-dashed rounded-2xl p-5 md:p-6 transition-all duration-300
                ${files.length > 0 ? 'border-indigo-300 bg-indigo-50/50' : 'border-slate-200 hover:border-indigo-400 hover:bg-slate-50/50'}
              `}
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
                <div className={`
                  w-10 h-10 rounded-full flex-shrink-0 flex items-center justify-center transition-all duration-500 group-hover:scale-110
                  ${files.length > 0 ? 'bg-indigo-100 text-indigo-600 shadow-[0_0_12px_rgba(79,70,229,0.2)]' : 'bg-slate-100 text-slate-400'}
                `}>
                  <Upload className="w-5 h-5" />
                </div>
                <div>
                  <p className="text-sm font-medium text-slate-800">點擊或拖拽上傳文件</p>
                  <p className="text-xs text-slate-400 font-light mt-0.5">支援 .docx, .xlsx, .pdf, .pptx（可多選）</p>
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
                    <div className="p-4 bg-amber-50 border border-amber-200 rounded-xl flex items-start gap-3 mb-4">
                      <AlertCircle className="w-5 h-5 text-amber-500 flex-shrink-0 mt-0.5" />
                      <div className="text-sm text-amber-800">
                        <p className="font-bold mb-1">PDF 翻譯注意事項</p>
                        <p>系統目前僅能擷取 PDF 中的「純文字」進行翻譯，將會**遺失原有的表格與排版**。若為揃描檔或圖片 PDF，系統會自動啟用 OCR (光學字元辨識) 進行處理，但辨識可能需要較長時間。若您的 PDF 包含表格或特殊字體（可能導致亂碼），強烈建議您先將其轉換為 Word 或 Excel 檔案後再進行翻譯。</p>
                      </div>
                    </div>
                  )}
                  <div className="flex justify-between items-center px-1 mb-2">
                    <span className="text-[11px] uppercase tracking-wider font-semibold text-slate-400">待處理檔案 ({files.length})</span>
                    {status === 'idle' && (
                      <button 
                        onClick={() => setFiles([])}
                        className="text-[10px] text-red-500 hover:underline"
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
                      className="flex items-center justify-between p-4 bg-white/60 rounded-2xl border border-white hover:bg-white hover:shadow-sm transition-all"
                    >
                      <div className="flex items-center gap-4 min-w-0">
                        <div className="w-10 h-10 rounded-xl bg-indigo-50 flex items-center justify-center text-indigo-600 shadow-sm">
                          {f.name.endsWith('.docx') && <FileText className="w-5 h-5" />}
                          {f.name.endsWith('.xlsx') && <FileSpreadsheet className="w-5 h-5" />}
                          {f.name.endsWith('.pdf') && <FileIcon className="w-5 h-5" />}
                          {f.name.endsWith('.pptx') && <Presentation className="w-5 h-5" />}
                        </div>
                        <div className="min-w-0">
                          <p className="text-sm font-medium text-slate-800 truncate">{f.name}</p>
                          <p className="text-[10px] text-slate-500 font-light">{(f.size / 1024).toFixed(1)} KB</p>
                        </div>
                      </div>
                      
                      <div className="flex items-center gap-4">
                        {status !== 'idle' && status !== 'error' && (
                          <div className="flex flex-col items-end gap-1.5">
                            <span className="text-[10px] font-mono font-bold text-indigo-600">
                              {fileProgress[f.name] || 0}%
                            </span>
                            <div className="w-16 h-1 bg-slate-200 rounded-full overflow-hidden">
                              <div 
                                className="h-full bg-indigo-500 transition-all duration-500 shadow-[0_0_10px_rgba(99,102,241,0.5)]" 
                                style={{ width: `${fileProgress[f.name] || 0}%` }}
                              />
                            </div>
                          </div>
                        )}
                        
                        {status === 'idle' && uploadStatus[f.name] === 'uploading' && (
                          <div className="flex flex-col items-end gap-1.5">
                            <span className="text-[10px] font-mono font-bold text-blue-500">
                              {Math.round(uploadProgress[f.name] || 0)}%
                            </span>
                            <div className="w-16 h-1 bg-slate-200 rounded-full overflow-hidden">
                              <div 
                                className="h-full bg-blue-500 transition-all duration-200 shadow-[0_0_10px_rgba(59,130,246,0.5)]" 
                                style={{ width: `${uploadProgress[f.name] || 0}%` }}
                              />
                            </div>
                          </div>
                        )}

                        {status === 'idle' && uploadStatus[f.name] === 'success' && (
                          <div className="flex items-center gap-2">
                            <span className="text-[10px] text-emerald-600 font-medium bg-emerald-50 px-2 py-1 rounded-full flex items-center gap-1">
                              <CheckCircle2 className="w-3 h-3" />
                              上傳成功
                            </span>
                            <button 
                              onClick={() => removeFile(idx)}
                              className="p-2 hover:bg-red-50 rounded-full text-slate-400 hover:text-red-500 transition-all"
                            >
                              <AlertCircle className="w-4 h-4 rotate-45" />
                            </button>
                          </div>
                        )}

                        {status === 'idle' && !uploadStatus[f.name] && (
                          <button 
                            onClick={() => removeFile(idx)}
                            className="p-2 hover:bg-red-50 rounded-full text-slate-400 hover:text-red-500 transition-all"
                          >
                            <AlertCircle className="w-4 h-4 rotate-45" />
                          </button>
                        )}

                        {status === 'completed' && fileProgress[f.name] === 100 && (
                          <div className="w-8 h-8 rounded-full bg-emerald-50 flex items-center justify-center">
                            <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                          </div>
                        )}
                      </div>
                    </motion.div>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>

            {/* Settings */}
            <div className="mt-10 space-y-8">
              <div>
                <label className="block text-xs uppercase tracking-widest font-bold text-slate-500 mb-3 ml-1">
                  工廠行業 <span className="text-slate-400 font-normal normal-case tracking-normal">(可選，使翻譯更精準)</span>
                </label>
                <div className="relative">
                  <input
                    type="text"
                    value={industry}
                    onChange={(e) => setIndustry(e.target.value)}
                    placeholder="例如：電子、紡織、汽車..."
                    className="w-full px-5 py-3.5 rounded-xl border border-slate-200 bg-white/50 focus:bg-white focus:border-indigo-400 focus:ring-4 focus:ring-indigo-400/10 outline-none transition-all text-sm text-slate-800 placeholder:text-slate-400"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs uppercase tracking-widest font-bold text-slate-500 mb-4 ml-1">
                  選擇目標語言 <span className="text-slate-400 font-normal normal-case tracking-normal">(可多選)</span>
                </label>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  {AVAILABLE_LANGUAGES.map((lang) => (
                    <button
                      key={lang.id}
                      onClick={() => toggleLanguage(lang.name)}
                      className={`
                        group relative flex items-center justify-between px-4 py-3.5 rounded-xl border transition-all duration-300
                        ${selectedLanguages.includes(lang.name) 
                          ? 'border-indigo-400 bg-indigo-50/80 text-indigo-700 shadow-sm' 
                          : 'border-slate-200 bg-white/50 hover:border-indigo-300 hover:bg-indigo-50/30 text-slate-600'}
                      `}
                    >
                      <div className="flex items-center gap-3">
                        <span className="text-lg">{lang.flag}</span>
                        <span className="text-xs font-medium tracking-wide">{lang.name}</span>
                      </div>
                      {selectedLanguages.includes(lang.name) && (
                        <CheckCircle2 className="w-4 h-4 text-indigo-500" />
                      )}
                    </button>
                  ))}
                </div>
              </div>

              {selectedLanguages.length > 1 && (
                <div>
                  <label className="block text-xs uppercase tracking-widest font-bold text-slate-500 mb-4 ml-1">
                    輸出檔案模式
                  </label>
                  <div className="flex flex-col sm:flex-row gap-3">
                    <label className={`flex items-center gap-3 px-4 py-3.5 rounded-xl border cursor-pointer transition-all duration-300 flex-1 ${outputMode === 'combined' ? 'border-indigo-400 bg-indigo-50/80' : 'border-slate-200 bg-white/50 hover:border-indigo-300 hover:bg-indigo-50/30'}`}>
                      <input 
                        type="radio" 
                        name="outputMode" 
                        value="combined" 
                        checked={outputMode === 'combined'} 
                        onChange={() => setOutputMode('combined')}
                        className="w-4 h-4 text-indigo-600 border-slate-300 bg-white focus:ring-indigo-500"
                      />
                      <span className="text-sm font-medium text-slate-700">合併為單一檔案</span>
                    </label>
                    <label className={`flex items-center gap-3 px-4 py-3.5 rounded-xl border cursor-pointer transition-all duration-300 flex-1 ${outputMode === 'separate' ? 'border-indigo-400 bg-indigo-50/80' : 'border-slate-200 bg-white/50 hover:border-indigo-300 hover:bg-indigo-50/30'}`}>
                      <input 
                        type="radio" 
                        name="outputMode" 
                        value="separate" 
                        checked={outputMode === 'separate'} 
                        onChange={() => setOutputMode('separate')}
                        className="w-4 h-4 text-indigo-600 border-slate-300 bg-white focus:ring-indigo-500"
                      />
                      <span className="text-sm font-medium text-slate-700">分開為多個檔案</span>
                    </label>
                  </div>
                </div>
              )}
              
              <div className="flex flex-col gap-3 pt-6">
                {!user ? (
                  <button
                    onClick={handleLogin}
                    className="w-full h-[56px] md:h-[64px] rounded-2xl font-semibold text-sm tracking-widest uppercase transition-all duration-500 flex items-center justify-center gap-3 bg-gradient-to-r from-indigo-600 to-violet-600 text-white hover:shadow-[0_8px_20px_rgba(99,102,241,0.3)] hover:-translate-y-0.5 active:scale-[0.98]"
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
                    className="w-full h-[56px] md:h-[64px] rounded-2xl font-semibold text-sm tracking-widest uppercase transition-all duration-500 flex items-center justify-center gap-3 bg-gradient-to-r from-amber-500 to-orange-500 text-white hover:shadow-[0_8px_20px_rgba(245,158,11,0.3)] hover:-translate-y-0.5 active:scale-[0.98]"
                  >
                    <AlertCircle className="w-5 h-5 shrink-0" />
                    <span className="truncate">請先驗證 Email 以開始翻譯</span>
                  </button>
                ) : (
                  <button
                    disabled={files.length === 0 || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating' || isUploading}
                    onClick={processFiles}
                    className={`
                      w-full h-[56px] md:h-[64px] rounded-2xl font-semibold text-sm tracking-widest uppercase transition-all duration-500 flex items-center justify-center gap-3
                      ${files.length === 0 || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating' || isUploading
                        ? 'bg-slate-100 text-slate-400 cursor-not-allowed border border-slate-200'
                        : 'bg-gradient-to-r from-indigo-600 to-violet-600 text-white hover:shadow-[0_8px_20px_rgba(99,102,241,0.3)] hover:-translate-y-0.5 active:scale-[0.98]'}
                    `}
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
                    className="w-full h-[48px] rounded-[20px] font-medium text-red-500 hover:bg-red-50 transition-all text-sm border border-red-200"
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
      </div>
    </div>
    </ErrorBoundary>
  );
}
