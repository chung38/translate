import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Shield, LogIn, X, Search, File as FileIcon, Clock, User as UserIcon } from 'lucide-react';
import { initializeApp, deleteApp } from 'firebase/app';
import { getAuth as getAuthSecondary, createUserWithEmailAndPassword as createUserSecondary, updateProfile as updateProfileSecondary } from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';
import { 
  db, 
  doc, 
  setDoc, 
  deleteDoc,
  collection, 
  query, 
  getDocs,
  handleFirestoreError,
  OperationType,
  Timestamp,
  where,
  limit
} from '../firebase';
import { UserProfile } from '../types';
import { User } from 'firebase/auth';

interface AdminPanelProps {
  isOpen: boolean;
  onClose: () => void;
  userProfile: UserProfile | null;
  user: User | null;
}

export const AdminPanel: React.FC<AdminPanelProps> = ({ isOpen, onClose, userProfile, user }) => {
  const [allUsers, setAllUsers] = useState<UserProfile[]>([]);
  const [allHistory, setAllHistory] = useState<any[]>([]);
  const [adminTab, setAdminTab] = useState<'users' | 'history'>('users');
  const [showAddUserForm, setShowAddUserForm] = useState(false);
  const [newUserEmail, setNewUserEmail] = useState('');
  const [newUserPassword, setNewUserPassword] = useState('');
  const [newUserDisplayName, setNewUserDisplayName] = useState('');
  const [deletingUserId, setDeletingUserId] = useState<string | null>(null);
  const [adminSearch, setAdminSearch] = useState('');
  const [addUserMessage, setAddUserMessage] = useState<{ type: 'success' | 'error', text: string } | null>(null);

  // Fetch all users for admin panel
  const fetchAllUsers = async () => {
    if (!userProfile || userProfile.role !== 'admin') return;
    try {
      const usersSnap = await getDocs(collection(db, 'users'));
      const rawUsers = usersSnap.docs
        .map(doc => doc.data() as UserProfile)
        .filter(u => !u.isPendingDeletion);
      
      // Deduplicate by Email (Prioritize Email)
      const emailMap = new Map<string, UserProfile>();
      rawUsers.forEach(u => {
        const email = u.email?.toLowerCase() || u.uid; // Fallback to UID if email is missing
        
        const existing = emailMap.get(email);
        if (!existing) {
          emailMap.set(email, u);
        } else {
          // Keep the one with more privileges (Admin > User, Paid > Free)
          const isExistingAdmin = existing.role === 'admin';
          const isNewAdmin = u.role === 'admin';
          const isExistingPaid = existing.isPaid;
          const isNewPaid = u.isPaid;
          
          if ((isNewAdmin && !isExistingAdmin) || (isNewPaid && !isExistingPaid)) {
            emailMap.set(email, u);
          }
        }
      });
      
      setAllUsers(Array.from(emailMap.values()));
    } catch (error) {
      handleFirestoreError(error, OperationType.GET, 'users', user);
    }
  };

  // Fetch all translation history for admin panel
  const fetchAllHistory = async () => {
    if (!userProfile || userProfile.role !== 'admin') return;
    try {
      const usersSnap = await getDocs(collection(db, 'users'));
      const historyPromises = usersSnap.docs.map(async (userDoc) => {
        const historySnap = await getDocs(collection(db, 'users', userDoc.id, 'history'));
        return historySnap.docs.map(hDoc => {
          const data = hDoc.data();
          return {
            id: `${userDoc.id}-${hDoc.id}`,
            userEmail: userDoc.data().email,
            userDisplayName: userDoc.data().displayName,
            fileName: data.fileName,
            timestamp: data.timestamp,
            targetLanguages: data.targetLanguages,
            ...data
          };
        });
      });
      const historyResults = await Promise.all(historyPromises);
      const flattenedHistory = historyResults.flat().sort((a, b) => 
        (b.timestamp?.toMillis() || 0) - (a.timestamp?.toMillis() || 0)
      );
      setAllHistory(flattenedHistory);
    } catch (error) {
      console.error("Fetch all history error:", error);
    }
  };

  const deleteUser = async (userId: string) => {
    try {
      if (!user) return;
      const idToken = await user.getIdToken();
      
      let backendSuccess = false;
      try {
        const response = await fetch(`/api/admin/users/${userId}`, {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${idToken}`
          }
        });
        
        if (response.ok) {
          const data = await response.json();
          if (data.firestoreDeleted) {
            backendSuccess = true;
          }
        }
      } catch (e) {
        console.warn("Backend delete failed, falling back to client SDK", e);
      }

      if (!backendSuccess) {
        // Fallback: Soft delete from Firestore using client SDK
        const userRef = doc(db, 'users', userId);
        await setDoc(userRef, { isPendingDeletion: true }, { merge: true });
      }
      
      setAllUsers(prev => prev.filter(u => u.uid !== userId));
      setDeletingUserId(null);
      setAddUserMessage({ type: 'success', text: '用戶已成功刪除！' });
      setTimeout(() => setAddUserMessage(null), 3000);
    } catch (error: any) {
      console.error("Delete user error:", error);
      setAddUserMessage({ type: 'error', text: `刪除使用者失敗: ${error.message}` });
      setDeletingUserId(null);
    }
  };

  const manualAddUser = async () => {
    if (!newUserEmail || !newUserDisplayName || !newUserPassword) return;
    
    setAddUserMessage(null);
    
    try {
      // Initialize a secondary Firebase app to create the user without logging out the admin
      let secondaryApp;
      try {
        secondaryApp = initializeApp(firebaseConfig, "SecondaryApp");
      } catch (e: any) {
        if (e.code === 'app/duplicate-app') {
          // If it already exists, delete it first then re-initialize
          const { getApp, deleteApp } = await import('firebase/app');
          await deleteApp(getApp("SecondaryApp"));
          secondaryApp = initializeApp(firebaseConfig, "SecondaryApp");
        } else {
          throw e;
        }
      }
      
      const secondaryAuth = getAuthSecondary(secondaryApp);
      const { setPersistence, inMemoryPersistence } = await import('firebase/auth');
      await setPersistence(secondaryAuth, inMemoryPersistence);
      
      let newUid = '';
      try {
        const userCredential = await createUserSecondary(secondaryAuth, newUserEmail, newUserPassword);
        newUid = userCredential.user.uid;
        await updateProfileSecondary(userCredential.user, { displayName: newUserDisplayName });
        await deleteApp(secondaryApp);
      } catch (authError: any) {
        await deleteApp(secondaryApp);
        if (authError.code === 'auth/email-already-in-use') {
          // Check if the user exists in Firestore and is pending deletion
          const q = query(collection(db, 'users'), where('email', '==', newUserEmail), limit(1));
          const emailQuerySnap = await getDocs(q);
          if (!emailQuerySnap.empty) {
            const existingDoc = emailQuerySnap.docs[0];
            const existingData = existingDoc.data() as UserProfile;
            if (existingData.isPendingDeletion) {
              // Restore the user
              const userRef = doc(db, 'users', existingDoc.id);
              await setDoc(userRef, { isPendingDeletion: false, isManuallyAdded: true }, { merge: true });
              setAddUserMessage({ type: 'success', text: '此帳號曾被刪除，已為您成功恢復！請用戶使用原密碼登入。' });
              fetchAllUsers(); // Refresh the list
              setNewUserEmail('');
              setNewUserPassword('');
              setNewUserDisplayName('');
              setShowAddUserForm(false);
              setTimeout(() => setAddUserMessage(null), 5000);
              return;
            }
          }
          setAddUserMessage({ type: 'error', text: '此電子郵件已被註冊。若下方列表未顯示該用戶，請請該用戶直接登入，系統將自動建立其資料。' });
          return;
        } else if (authError.code === 'auth/weak-password') {
          setAddUserMessage({ type: 'error', text: '密碼強度不足（至少需要6位字元）' });
          return;
        }
        throw authError; // Re-throw if it's another error
      }

      // Check if email already exists in Firestore (just in case)
      const q = query(collection(db, 'users'), where('email', '==', newUserEmail), limit(1));
      const emailQuerySnap = await getDocs(q);
      
      if (!emailQuerySnap.empty) {
        const existingDoc = emailQuerySnap.docs[0];
        const existingData = existingDoc.data() as UserProfile;
        
        // Update the existing document
        const userRef = doc(db, 'users', existingDoc.id);
        const newProfile: UserProfile = {
          ...existingData,
          uid: newUid,
          displayName: newUserDisplayName,
          email: newUserEmail
        };
        await setDoc(userRef, newProfile);
        
        setAllUsers(prev => prev.map(u => u.email === newUserEmail ? newProfile : u));
      } else {
        const userRef = doc(db, 'users', newUid);
        const newProfile: UserProfile = {
          uid: newUid,
          email: newUserEmail,
          displayName: newUserDisplayName,
          photoURL: null,
          createdAt: Timestamp.now(),
          role: 'user',
          isPaid: false,
          quota: 2,
          isManuallyAdded: true
        };
        await setDoc(userRef, newProfile);
        setAllUsers(prev => [...prev, newProfile]);
      }
      
      setShowAddUserForm(false);
      setNewUserEmail('');
      setNewUserPassword('');
      setNewUserDisplayName('');
      setAddUserMessage({ type: 'success', text: '用戶新增成功！' });
      
      // Clear success message after 3 seconds
      setTimeout(() => setAddUserMessage(null), 3000);
    } catch (error: any) {
      console.error("Manual add user error:", error);
      setAddUserMessage({ type: 'error', text: '新增用戶時發生錯誤: ' + (error.message || '未知錯誤') });
    }
  };

  useEffect(() => {
    if (isOpen) {
      fetchAllUsers();
    }
  }, [isOpen]);

  const updateUserPermission = async (userId: string, updates: Partial<UserProfile>) => {
    try {
      const userRef = doc(db, 'users', userId);
      await setDoc(userRef, updates, { merge: true });
      // Update local state
      setAllUsers(prev => prev.map(u => u.uid === userId ? { ...u, ...updates } : u));
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `users/${userId}`, user);
    }
  };

  return (
    <AnimatePresence>
      {isOpen && (
        <motion.div 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[100] flex items-center justify-center p-4"
        >
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="bg-white  w-full max-w-4xl max-h-[90vh] rounded-[32px] shadow-[0_8px_30px_rgb(0,0,0,0.12)] border border-white overflow-hidden flex flex-col"
          >
            <div className="p-6 border-b border-[var(--rule)] flex items-center justify-between bg-white/50">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-[var(--signal-dim)] rounded-xl border border-[var(--rule)]">
                  <Shield className="w-6 h-6 text-[var(--signal)]" />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-[var(--ink)]">管理後台</h2>
                  <div className="flex gap-4 mt-1">
                    <button 
                      onClick={() => setAdminTab('users')}
                      className={`text-xs font-bold uppercase tracking-wider transition-colors ${adminTab === 'users' ? 'text-[var(--signal)]' : 'text-[var(--muted)] hover:text-slate-700'}`}
                    >
                      用戶管理
                    </button>
                    <button 
                      onClick={() => {
                        setAdminTab('history');
                        fetchAllHistory();
                      }}
                      className={`text-xs font-bold uppercase tracking-wider transition-colors ${adminTab === 'history' ? 'text-[var(--signal)]' : 'text-[var(--muted)] hover:text-slate-700'}`}
                    >
                      全域紀錄
                    </button>
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2">
                {adminTab === 'users' && (
                  <button 
                    onClick={() => setShowAddUserForm(!showAddUserForm)}
                    className="flex items-center gap-1 px-3 py-1.5 bg-[var(--signal-dim)] text-[var(--signal-ink)] border border-[var(--rule)] text-xs font-bold rounded-xl hover:bg-violet-200 transition-all shadow-sm"
                  >
                    <LogIn className="w-3.5 h-3.5" />
                    {showAddUserForm ? '取消新增' : '新增用戶'}
                  </button>
                )}
                <button 
                  onClick={onClose}
                  className="p-2 hover:bg-slate-100 rounded-full transition-colors"
                >
                  <X className="w-6 h-6 text-[var(--muted)] hover:text-slate-600" />
                </button>
              </div>
            </div>

            {showAddUserForm && adminTab === 'users' && (
              <div className="p-6 bg-[var(--signal-dim)] border-b border-violet-100 flex flex-col gap-3">
                <div className="flex flex-col sm:flex-row gap-3 items-end">
                  <div className="flex-1 space-y-1">
                    <label className="text-[10px] font-bold text-[var(--signal)] uppercase tracking-wider">帳戶名稱</label>
                    <input 
                      type="text" 
                      placeholder="用戶名稱"
                      value={newUserDisplayName}
                      onChange={(e) => setNewUserDisplayName(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-[var(--rule)] rounded-xl text-sm text-[var(--ink)] focus:outline-none focus:ring-2 focus:ring-violet-500/50"
                    />
                  </div>
                  <div className="flex-1 space-y-1">
                    <label className="text-[10px] font-bold text-[var(--signal)] uppercase tracking-wider">用戶信箱</label>
                    <input 
                      type="email" 
                      placeholder="example@gmail.com"
                      value={newUserEmail}
                      onChange={(e) => setNewUserEmail(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-[var(--rule)] rounded-xl text-sm text-[var(--ink)] focus:outline-none focus:ring-2 focus:ring-violet-500/50"
                    />
                  </div>
                  <div className="flex-1 space-y-1">
                    <label className="text-[10px] font-bold text-[var(--signal)] uppercase tracking-wider">登入密碼</label>
                    <input 
                      type="password" 
                      placeholder="至少 6 位字元"
                      value={newUserPassword}
                      onChange={(e) => setNewUserPassword(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-[var(--rule)] rounded-xl text-sm text-[var(--ink)] focus:outline-none focus:ring-2 focus:ring-violet-500/50"
                    />
                  </div>
                  <button 
                    onClick={manualAddUser}
                    disabled={!newUserEmail || !newUserDisplayName || !newUserPassword || newUserPassword.length < 6}
                    className="px-6 py-2 bg-violet-600 text-white text-sm font-bold rounded-xl hover:bg-violet-700 transition-all disabled:opacity-50 border border-violet-500"
                  >
                    確認新增
                  </button>
                </div>
              </div>
            )}

            {addUserMessage && (
              <div className={`mx-6 mt-6 p-3 rounded-xl text-sm font-medium ${addUserMessage.type === 'success' ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' : 'bg-red-50 text-red-700 border border-red-200'}`}>
                {addUserMessage.text}
              </div>
            )}

            <div className="p-6 bg-[var(--paper)] border-b border-[var(--rule)]">
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[var(--muted)]" />
                <input 
                  type="text" 
                  placeholder={adminTab === 'users' ? "搜尋用戶信箱或名稱..." : "搜尋檔案名稱或用戶..."}
                  value={adminSearch}
                  onChange={(e) => setAdminSearch(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 bg-white border border-[var(--rule)] rounded-xl text-sm text-[var(--ink)] focus:outline-none focus:ring-2 focus:ring-violet-500/50 focus:border-violet-500/50 transition-all placeholder:text-[var(--muted)]"
                />
              </div>
            </div>

            <div className="flex-1 overflow-y-auto p-6">
              {adminTab === 'users' ? (
                <div className="grid grid-cols-1 gap-4">
                  {allUsers
                    .filter(u => {
                      const searchLower = adminSearch.toLowerCase();
                      const emailMatch = u.email ? u.email.toLowerCase().includes(searchLower) : false;
                      const nameMatch = u.displayName ? u.displayName.toLowerCase().includes(searchLower) : false;
                      return emailMatch || nameMatch;
                    })
                    .map((u) => (
                    <React.Fragment key={u.uid}>
                      <div className="p-4 rounded-2xl border border-[var(--rule)] hover:border-violet-300 hover:bg-[var(--signal-dim)]/50 transition-all flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-white">
                        <div className="flex items-center gap-3 min-w-0">
                          {u.photoURL ? (
                            <img src={u.photoURL} alt="" className="w-10 h-10 rounded-full border border-[var(--rule)]" referrerPolicy="no-referrer" />
                          ) : (
                            <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center border border-[var(--rule)]">
                              <UserIcon className="w-5 h-5 text-[var(--muted)]" />
                            </div>
                          )}
                          <div className="min-w-0">
                            <div className="flex items-center gap-2">
                              <p className="text-sm font-bold text-[var(--ink)] truncate">{u.displayName || '未命名用戶'}</p>
                            </div>
                            <div className="flex items-center gap-1.5">
                              <p className="text-xs text-[var(--muted)] truncate">{u.email}</p>
                              {u.emailVerified !== undefined && (
                                <span className={`px-1.5 py-0.5 text-[9px] font-bold rounded uppercase ${u.emailVerified ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'}`}>
                                  {u.emailVerified ? '已驗證' : '未驗證'}
                                </span>
                              )}
                            </div>
                          </div>
                        </div>

                        <div className="flex flex-wrap items-center gap-4">
                          <div className="flex flex-col gap-1">
                            <span className="text-[10px] uppercase tracking-wider text-[var(--muted)] font-bold">角色權限</span>
                            <select 
                              value={u.role}
                              onChange={(e) => updateUserPermission(u.uid, { role: e.target.value as 'user' | 'admin' })}
                              className="text-xs font-medium bg-white border border-[var(--rule)] rounded-lg px-2 py-1 text-[var(--ink)] focus:outline-none focus:ring-2 focus:ring-violet-500/50"
                            >
                              <option value="user">一般用戶</option>
                              <option value="admin">管理員</option>
                            </select>
                          </div>

                          <div className="flex flex-col gap-1">
                            <span className="text-[10px] uppercase tracking-wider text-[var(--muted)] font-bold">付費狀態</span>
                            <button 
                              onClick={() => updateUserPermission(u.uid, { isPaid: !u.isPaid })}
                              className={`
                                px-3 py-1 rounded-lg text-xs font-bold transition-all
                                ${u.isPaid 
                                  ? 'bg-emerald-100 text-emerald-700 hover:bg-emerald-200' 
                                  : 'bg-slate-100 text-[var(--muted)] hover:bg-slate-200'}
                              `}
                            >
                              {u.isPaid ? 'PRO (已付費)' : 'FREE (未付費)'}
                            </button>
                          </div>

                          <div className="flex flex-col gap-1">
                            <span className="text-[10px] uppercase tracking-wider text-[var(--muted)] font-bold">剩餘額度</span>
                            <input 
                              type="number" 
                              value={u.quota}
                              onChange={(e) => updateUserPermission(u.uid, { quota: parseInt(e.target.value) || 0 })}
                              className="w-16 text-xs font-medium bg-white border border-[var(--rule)] rounded-lg px-2 py-1 text-[var(--ink)] focus:outline-none focus:ring-2 focus:ring-violet-500/50"
                            />
                          </div>

                          <button 
                            onClick={() => setDeletingUserId(u.uid)}
                            className="p-2 rounded-lg transition-all text-red-500 hover:text-red-700 hover:bg-red-50"
                            title="刪除帳號"
                          >
                            <X className="w-5 h-5" />
                          </button>
                        </div>
                      </div>
                      
                      {deletingUserId === u.uid && (
                        <div className="mt-2 p-3 bg-red-50 border border-red-200 rounded-xl flex items-center justify-between">
                          <p className="text-xs text-red-600 font-medium">確定要刪除此用戶嗎？該用戶將會被立即刪除。</p>
                          <div className="flex gap-2">
                            <button 
                              onClick={() => setDeletingUserId(null)}
                              className="px-3 py-1 text-xs font-bold text-[var(--muted)] hover:text-[var(--ink)]"
                            >
                              取消
                            </button>
                            <button 
                              onClick={() => deleteUser(u.uid)}
                              className="px-3 py-1 bg-red-600 text-white text-xs font-bold rounded-lg hover:bg-red-700 border border-red-500"
                            >
                              確認刪除
                            </button>
                          </div>
                        </div>
                      )}
                    </React.Fragment>
                  ))}
                </div>
              ) : (
                <div className="space-y-3">
                  {allHistory
                    .filter(h => {
                      const searchLower = adminSearch.toLowerCase();
                      const fileMatch = h.fileName ? h.fileName.toLowerCase().includes(searchLower) : false;
                      const emailMatch = h.userEmail ? h.userEmail.toLowerCase().includes(searchLower) : false;
                      return fileMatch || emailMatch;
                    })
                    .map((item) => (
                    <div key={item.id} className="p-4 rounded-2xl border border-[var(--rule)] bg-white flex items-center justify-between gap-4">
                      <div className="flex items-center gap-3 min-w-0">
                        <div className="p-2 bg-[var(--signal-dim)] rounded-lg border border-indigo-100">
                          <FileIcon className="w-4 h-4 text-[var(--signal)]" />
                        </div>
                        <div className="min-w-0">
                          <p className="text-sm font-bold text-slate-700 truncate">{item.fileName}</p>
                          <div className="flex items-center gap-2 mt-0.5">
                            <span className="text-[10px] text-[var(--muted)] font-medium">{item.userEmail}</span>
                            <span className="text-[10px] text-slate-300">•</span>
                            <span className="text-[10px] text-[var(--muted)]">{item.timestamp?.toDate().toLocaleString()}</span>
                          </div>
                        </div>
                      </div>
                      <div className="flex gap-1">
                        {item.targetLanguages?.map((l: string) => (
                          <span key={l} className="px-1.5 py-0.5 bg-[var(--signal-dim)] text-[var(--signal-ink)] text-[9px] font-bold rounded uppercase">
                            {l}
                          </span>
                        ))}
                      </div>
                    </div>
                  ))}
                  {allHistory.length === 0 && (
                    <div className="text-center py-12">
                      <Clock className="w-12 h-12 text-slate-200 mx-auto mb-4" />
                      <p className="text-[var(--muted)] text-sm">尚無全域翻譯紀錄</p>
                    </div>
                  )}
                </div>
              )}
            </div>
          </motion.div>
        </motion.div>
      )}
    </AnimatePresence>
  );
};
