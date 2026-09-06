import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { X, User as UserIcon, AlertCircle, Loader2, ExternalLink } from 'lucide-react';
import { 
  auth, 
  db, 
  googleProvider, 
  signInWithPopup,
  signInWithRedirect,
  getRedirectResult,
  // isMobileDevice 已不再需要（改用 popup，不分裝置）
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  updateProfile,
  sendEmailVerification,
  sendPasswordResetEmail,
  doc, 
  setDoc, 
  Timestamp
} from '../firebase';

interface AuthModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSuccess?: (message?: string) => void;
}

/** 偵測是否為 WebView（Line、Facebook、Instagram 等 App 內建瀏覽器） */
function isWebView(): boolean {
  const ua = navigator.userAgent || '';
  return (
    /FBAN|FBAV|Instagram|Line\/|MicroMessenger|WebView|wv/.test(ua) ||
    ((/iPhone|iPod|iPad/.test(ua)) && !/Safari\//.test(ua)) ||
    (/Android/.test(ua) && /Version\/\d/.test(ua) && /Chrome\/\d/.test(ua) && /Mobile Safari\/\d/.test(ua) === false)
  );
}

export const AuthModal: React.FC<AuthModalProps> = ({ isOpen, onClose, onSuccess }) => {
  const [authMode, setAuthMode] = useState<'login' | 'register' | 'reset'>('login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [displayName, setDisplayName] = useState('');
  const [authLoading, setAuthLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [resetSent, setResetSent] = useState(false);
  const [inWebView, setInWebView] = useState(false);

  useEffect(() => {
    if (isOpen) {
      setInWebView(isWebView());
    }
  }, [isOpen]);

  const handleResetPassword = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email) {
      setError('請輸入電子郵件');
      return;
    }
    setAuthLoading(true);
    setError(null);
    try {
      await sendPasswordResetEmail(auth, email);
      setResetSent(true);
      setError('密碼重設信已發送，請檢查您的信箱。');
    } catch (err: any) {
      let msg = '發送失敗，請檢查您的輸入';
      if (err.code === 'auth/user-not-found') msg = '找不到此電子郵件對應的帳號';
      else if (err.code === 'auth/invalid-email') msg = '無效的電子郵件格式';
      setError(msg);
    } finally {
      setAuthLoading(false);
    }
  };

  const handleEmailAuth = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email || !password) {
      setError('請輸入電子郵件和密碼');
      return;
    }
    if (authMode === 'register' && !displayName) {
      setError('請輸入帳戶名稱');
      return;
    }
    setAuthLoading(true);
    setError(null);
    try {
      if (authMode === 'register') {
        const userCredential = await createUserWithEmailAndPassword(auth, email, password);
        await updateProfile(userCredential.user, { displayName });
        
        try {
          await sendEmailVerification(userCredential.user);
          console.log("Verification email sent.");
        } catch (e: any) {
          console.error("Failed to send verification email:", e);
          if (e.code === 'auth/too-many-requests') {
            setError('發送驗證信次數過多，請稍後再試。');
          } else {
            setError('註冊成功，但驗證信發送失敗，請稍後登入並點擊「重發驗證信」。');
          }
        }
        
        const userRef = doc(db, 'users', userCredential.user.uid);
        await setDoc(userRef, {
          uid: userCredential.user.uid,
          email: userCredential.user.email,
          emailVerified: false,
          displayName: displayName,
          photoURL: null,
          createdAt: Timestamp.now(),
          role: 'user',
          isPaid: false,
          quota: 2
        }, { merge: true });
      } else {
        await signInWithEmailAndPassword(auth, email, password);
      }
      
      const successMsg = authMode === 'register' ? (error || '註冊成功！請檢查您的電子郵件並點擊驗證連結以啟用完整權限。') : undefined;
      
      setEmail('');
      setPassword('');
      setDisplayName('');
      onClose();
      if (onSuccess) onSuccess(successMsg);
      
    } catch (err: any) {
      let msg = '認證失敗，請檢查您的輸入';
      
      const errorCode = err.code;
      if (errorCode === 'auth/email-already-in-use') msg = '此電子郵件已被使用，請直接登入';
      else if (errorCode === 'auth/invalid-email') msg = '無效的電子郵件格式';
      else if (errorCode === 'auth/weak-password') msg = '密碼強度不足（至少需要6位字元）';
      else if (errorCode === 'auth/operation-not-allowed') msg = '尚未在 Firebase 啟用此登入方式，請聯繫管理員。';
      else if (errorCode === 'auth/user-not-found' || errorCode === 'auth/wrong-password' || errorCode === 'auth/invalid-credential') {
        msg = '帳號或密碼錯誤。若您尚未設定密碼，請先進行「註冊」。';
      } else if (errorCode === 'auth/too-many-requests') {
        msg = '嘗試次數過多，帳號已被暫時鎖定，請稍後再試';
      } else {
        console.error('Auth error:', err);
      }
      
      setError(msg);
    } finally {
      setAuthLoading(false);
    }
  };

  const handleGoogleLogin = async () => {
    if (authLoading) return;

    // WebView 環境下 Google 登入會被封鎖，提示使用者改用系統瀏覽器
    if (inWebView) {
      setError('目前在 App 內建瀏覽器中，Google 登入不被允許。請點擊下方按鈕，用系統預設瀏覽器開啟本頁面後再登入。');
      return;
    }

    setAuthLoading(true);
    setError(null);
    try {
      // 手機不要再用 signInWithRedirect。
      // Firebase 的 redirect 流程靠一個連到 <專案>.firebaseapp.com 的跨來源 iframe
      // 把登入結果傳回來，但本站在 pages.dev、authDomain 在 firebaseapp.com，不同源。
      // Safari 16.1+、Chrome 115+、Firefox 109+ 會擋掉這種第三方儲存存取，
      // 症狀就是：Google 那邊帳號選完、跳回網站卻仍然是未登入狀態。
      // popup 走 window.postMessage，不受這個限制，現在的手機瀏覽器都支援。
      await signInWithPopup(auth, googleProvider);
      onClose();
      if (onSuccess) onSuccess();
    } catch (err: any) {
      if (err.code === 'auth/popup-closed-by-user') return;
      if (err.code === 'auth/cancelled-popup-request') return;

      // 少數瀏覽器會擋 popup，這時才退回 redirect
      if (err.code === 'auth/popup-blocked' || err.code === 'auth/operation-not-supported-in-this-environment') {
        try {
          await signInWithRedirect(auth, googleProvider);
          return;
        } catch {
          setError('這個瀏覽器擋住了 Google 登入視窗。請改用帳號密碼登入，或換 Chrome / Safari 開啟本頁。');
          return;
        }
      }

      console.error('Google login error:', err);
      let msg = 'Google 登入失敗';
      if (err.code === 'auth/unauthorized-domain') msg = '此網域未授權，請聯繫管理員';
      else if (err.code === 'auth/disallowed-useragent' || (err.message && err.message.includes('disallowed_useragent'))) {
        msg = '您的瀏覽器不支援 Google 登入，請改用系統預設瀏覽器開啟本頁面。';
        setInWebView(true);
      }
      setError(msg);
    } finally {
      setAuthLoading(false);
    }
  };

  const handleOpenInBrowser = () => {
    const url = window.location.href;
    // 嘗試透過 location.href 跳轉（部分 App 會自動用系統瀏覽器開啟）
    window.location.href = url;
  };

  return (
    <AnimatePresence>
      {isOpen && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            exit={{ opacity: 0, scale: 0.95 }}
            className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden relative"
          >
            <button 
              onClick={onClose}
              className="absolute top-4 right-4 text-[var(--muted)] hover:text-slate-600 transition-colors"
            >
              <X className="w-5 h-5" />
            </button>

            <div className="p-8">
              <div className="text-center mb-8">
                <div className="w-12 h-12 bg-emerald-100 rounded-xl flex items-center justify-center mx-auto mb-4">
                  <UserIcon className="w-6 h-6 text-emerald-600" />
                </div>
                <h2 className="text-2xl font-bold text-[var(--ink)]">
                  {authMode === 'login' ? '歡迎回來' : authMode === 'register' ? '建立帳號' : '重設密碼'}
                </h2>
                <p className="text-sm text-[var(--muted)] mt-1">
                  {authMode === 'login' ? '請登入您的帳號以繼續' : authMode === 'register' ? '註冊一個新帳號開始使用' : '輸入您的電子郵件，我們將發送密碼重設連結給您'}
                </p>
              </div>

              {/* WebView 警告橫幅 */}
              {inWebView && (
                <div className="mb-6 p-4 bg-amber-50 border border-amber-200 rounded-xl">
                  <div className="flex items-start gap-3">
                    <AlertCircle className="w-5 h-5 text-amber-500 flex-shrink-0 mt-0.5" />
                    <div className="text-sm text-amber-800">
                      <p className="font-bold mb-1">偵測到 App 內建瀏覽器</p>
                      <p className="mb-3">您目前在 Line / Facebook 等 App 的內建瀏覽器中，Google 登入不支援此環境。</p>
                      <p className="mb-3 font-medium">請改用<strong>帳號密碼登入</strong>，或點擊下方按鈕用系統瀏覽器開啟：</p>
                      <button
                        onClick={handleOpenInBrowser}
                        className="w-full flex items-center justify-center gap-2 py-2.5 px-4 bg-amber-500 text-white font-bold rounded-lg hover:bg-amber-600 transition-colors text-sm"
                      >
                        <ExternalLink className="w-4 h-4" />
                        在系統瀏覽器中開啟
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {authMode === 'reset' ? (
                <form onSubmit={handleResetPassword} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-[var(--muted)] uppercase tracking-wider mb-1 ml-1">電子郵件</label>
                    <input 
                      type="email" 
                      value={email}
                      onChange={(e) => setEmail(e.target.value)}
                      placeholder="your@email.com"
                      className="w-full px-4 py-3 bg-[var(--paper)] border border-[var(--rule)] rounded-xl focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-[var(--ink)]"
                      required
                    />
                  </div>
                  {error && (
                    <div className={`flex items-center gap-2 p-3 rounded-lg text-sm ${resetSent ? 'bg-emerald-50 text-emerald-600' : 'bg-red-50 text-red-600'}`}>
                      <AlertCircle className="w-4 h-4 flex-shrink-0" />
                      <p>{error}</p>
                    </div>
                  )}
                  <button 
                    type="submit"
                    disabled={authLoading}
                    className="w-full py-3 bg-emerald-600 text-white font-bold rounded-xl hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-600/20 flex items-center justify-center gap-2 disabled:opacity-70"
                  >
                    {authLoading ? (
                      <Loader2 className="w-5 h-5 animate-spin" />
                    ) : (
                      '發送重設密碼信'
                    )}
                  </button>
                  <div className="mt-8 text-center">
                    <button 
                      type="button"
                      onClick={() => { setAuthMode('login'); setError(null); setResetSent(false); }}
                      className="text-sm text-[var(--muted)] hover:underline"
                    >
                      返回登入
                    </button>
                  </div>
                </form>
              ) : (
                <>
                  <form onSubmit={handleEmailAuth} className="space-y-4">
                    {authMode === 'register' && (
                      <div>
                        <label className="block text-xs font-bold text-[var(--muted)] uppercase tracking-wider mb-1 ml-1">帳戶名稱</label>
                        <input 
                          type="text" 
                          value={displayName}
                          onChange={(e) => setDisplayName(e.target.value)}
                          placeholder="您的名字"
                          className="w-full px-4 py-3 bg-[var(--paper)] border border-[var(--rule)] rounded-xl focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-[var(--ink)]"
                          required
                        />
                      </div>
                    )}
                    <div>
                      <label className="block text-xs font-bold text-[var(--muted)] uppercase tracking-wider mb-1 ml-1">電子郵件</label>
                      <input 
                        type="email" 
                        value={email}
                        onChange={(e) => setEmail(e.target.value)}
                        placeholder="your@email.com"
                        className="w-full px-4 py-3 bg-[var(--paper)] border border-[var(--rule)] rounded-xl focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-[var(--ink)]"
                        required
                      />
                    </div>
                    <div>
                      <div className="flex items-center justify-between mb-1 ml-1">
                        <label className="block text-xs font-bold text-[var(--muted)] uppercase tracking-wider">密碼</label>
                        {authMode === 'login' && (
                          <button 
                            type="button" 
                            onClick={() => { setAuthMode('reset'); setError(null); setResetSent(false); }} 
                            className="text-xs text-emerald-600 hover:underline mr-1"
                          >
                            忘記密碼？
                          </button>
                        )}
                      </div>
                      <input 
                        type="password" 
                        value={password}
                        onChange={(e) => setPassword(e.target.value)}
                        placeholder="••••••••"
                        className="w-full px-4 py-3 bg-[var(--paper)] border border-[var(--rule)] rounded-xl focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-[var(--ink)]"
                        required
                      />
                    </div>

                    {error && (
                      <div className="flex items-center gap-2 p-3 bg-red-50 text-red-600 rounded-lg text-sm">
                        <AlertCircle className="w-4 h-4 flex-shrink-0" />
                        <p>{error}</p>
                      </div>
                    )}

                    <button 
                      type="submit"
                      disabled={authLoading}
                      className="w-full py-3 bg-emerald-600 text-white font-bold rounded-xl hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-600/20 flex items-center justify-center gap-2 disabled:opacity-70"
                    >
                      {authLoading ? (
                        <Loader2 className="w-5 h-5 animate-spin" />
                      ) : (
                        authMode === 'login' ? '登入' : '註冊'
                      )}
                    </button>
                  </form>

                  <div className="relative my-8">
                    <div className="absolute inset-0 flex items-center">
                      <div className="w-full border-t border-[var(--rule)]"></div>
                    </div>
                    <div className="relative flex justify-center text-xs uppercase">
                      <span className="bg-white px-4 text-[var(--muted)] font-medium">或使用</span>
                    </div>
                  </div>

                  <button 
                    onClick={handleGoogleLogin}
                    disabled={authLoading}
                    className={`w-full py-3 bg-white border border-[var(--rule)] text-slate-700 font-bold rounded-xl transition-all flex items-center justify-center gap-3 ${
                      inWebView 
                        ? 'opacity-40 cursor-not-allowed' 
                        : authLoading 
                          ? 'opacity-50 cursor-not-allowed' 
                          : 'hover:bg-[var(--paper)]'
                    }`}
                  >
                    {authLoading ? (
                      <Loader2 className="w-5 h-5 animate-spin text-[var(--muted)]" />
                    ) : (
                      <svg className="w-5 h-5" viewBox="0 0 24 24">
                        <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>
                        <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>
                        <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l3.66-2.84z"/>
                        <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>
                      </svg>
                    )}
                    {inWebView ? 'Google 登入（不支援此環境）' : authLoading ? '登入中...' : 'Google 登入'}
                  </button>

                  <div className="mt-8 text-center">
                    <button 
                      onClick={() => { setAuthMode(authMode === 'login' ? 'register' : 'login'); setError(null); }}
                      className="text-sm text-emerald-600 font-bold hover:underline"
                    >
                      {authMode === 'login' ? '還沒有帳號？立即註冊' : '已有帳號？返回登入'}
                    </button>
                  </div>
                </>
              )}
            </div>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );
};
