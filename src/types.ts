export interface UserProfile {
  uid: string;
  email: string | null;
  emailVerified?: boolean;
  displayName: string | null;
  photoURL: string | null;
  createdAt: any;
  role: 'user' | 'admin';
  isPaid: boolean;
  quota: number;
  isPendingDeletion?: boolean;
  isManuallyAdded?: boolean;
}
