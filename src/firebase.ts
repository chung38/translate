import { initializeApp } from 'firebase/app';
import { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, getRedirectResult, onAuthStateChanged, User, createUserWithEmailAndPassword, signInWithEmailAndPassword, deleteUser, updateProfile, sendEmailVerification, sendPasswordResetEmail } from 'firebase/auth';
import { getFirestore, doc, getDoc, setDoc, deleteDoc, collection, addDoc, query, orderBy, onSnapshot, getDocs, Timestamp, where, limit } from 'firebase/firestore';

// Import the Firebase configuration
import firebaseConfig from '../firebase-applet-config.json';

// Initialize Firebase SDK
const app = initializeApp(firebaseConfig);
// Use getFirestore with the database ID
export const db = getFirestore(app, firebaseConfig.firestoreDatabaseId);
export const auth = getAuth();
export const googleProvider = new GoogleAuthProvider();
googleProvider.setCustomParameters({
  prompt: 'select_account'
});

// Detect mobile environment
export const isMobileDevice = (): boolean => {
  return /Android|iPhone|iPad|iPod|Mobile|webOS|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
};

// Error Handling Spec for Firestore Operations
export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

export interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

export function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null, user?: User | null) {
  const currentUser = user || auth.currentUser;
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: currentUser?.uid,
      email: currentUser?.email,
      emailVerified: currentUser?.emailVerified,
      isAnonymous: currentUser?.isAnonymous,
      tenantId: currentUser?.tenantId,
      providerInfo: currentUser?.providerData.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  }
  // 完整資訊只留在 console；丟出去的訊息不要夾帶 uid / email，
  // 否則這包 JSON 會直接顯示在畫面上或被前端錯誤追蹤服務收走。
  console.error('Firestore Error: ', errInfo);
  const raw = error instanceof Error ? error.message : String(error);
  throw new Error(`資料庫操作失敗（${operationType}）：${raw}`);
}

export { 
  signInWithPopup,
  signInWithRedirect,
  getRedirectResult,
  onAuthStateChanged, 
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  updateProfile,
  sendEmailVerification,
  sendPasswordResetEmail,
  deleteUser, 
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
  Timestamp,
  where,
  limit
};
export type { User };
