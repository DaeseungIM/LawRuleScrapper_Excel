import { initializeApp } from 'firebase/app';
import {
  getAuth,
  signInWithPopup,
  GoogleAuthProvider,
  onAuthStateChanged,
  signOut,
  User
} from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);

const provider = new GoogleAuthProvider();
// Add required Google Workspace scopes for Sheets and Drive
provider.addScope('https://www.googleapis.com/auth/spreadsheets');
provider.addScope('https://www.googleapis.com/auth/drive.file');
provider.addScope('https://www.googleapis.com/auth/drive');
provider.setCustomParameters({
  prompt: 'select_account consent',
});

let isSigningIn = false;
let cachedAccessToken: string | null = null;

// Check if stored access token is still within the 55-minute validity window
export const isAccessTokenValid = (): boolean => {
  if (typeof window === 'undefined') return false;
  const token = localStorage.getItem('google_access_token');
  const timeStr = localStorage.getItem('google_access_token_time');
  if (!token || !timeStr) return false;
  const timestamp = parseInt(timeStr, 10);
  if (isNaN(timestamp)) return false;
  const elapsed = Date.now() - timestamp;
  // Google access tokens expire after 60 minutes (3600s). We consider it expired after 50 minutes.
  return elapsed < 50 * 60 * 1000;
};

export const clearAuthToken = () => {
  cachedAccessToken = null;
  if (typeof window !== 'undefined') {
    localStorage.removeItem('google_access_token');
    localStorage.removeItem('google_access_token_time');
  }
};

export const initAuth = (
  onAuthSuccess?: (user: User, token: string) => void,
  onAuthFailure?: () => void
) => {
  return onAuthStateChanged(auth, async (user: User | null) => {
    if (user) {
      if (isAccessTokenValid()) {
        const token = localStorage.getItem('google_access_token');
        cachedAccessToken = token;
        if (token && onAuthSuccess) {
          onAuthSuccess(user, token);
          return;
        }
      }
      // Token is expired or missing. Clear it and trigger failure so user re-authenticates fresh token.
      clearAuthToken();
      if (onAuthFailure) onAuthFailure();
    } else {
      clearAuthToken();
      if (onAuthFailure) onAuthFailure();
    }
  });
};

export const googleSignIn = async (): Promise<{ user: User; accessToken: string } | null> => {
  try {
    isSigningIn = true;
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Google OAuth 인증 토큰을 받지 못했습니다. 팝업에서 드라이브 및 스프레드시트 권한을 허용해 주세요.');
    }

    cachedAccessToken = credential.accessToken;
    if (typeof window !== 'undefined') {
      localStorage.setItem('google_access_token', cachedAccessToken);
      localStorage.setItem('google_access_token_time', String(Date.now()));
    }
    return { user: result.user, accessToken: cachedAccessToken };
  } catch (error: any) {
    console.error('Google Sign In Error:', error);
    throw error;
  } finally {
    isSigningIn = false;
  }
};

export const getAccessToken = (): string | null => {
  if (typeof window === 'undefined') return null;
  if (!isAccessTokenValid()) {
    clearAuthToken();
    return null;
  }
  if (!cachedAccessToken) {
    cachedAccessToken = localStorage.getItem('google_access_token');
  }
  return cachedAccessToken;
};

export const ensureFreshToken = async (): Promise<string> => {
  const current = getAccessToken();
  if (current) return current;
  const result = await googleSignIn();
  if (!result?.accessToken) {
    throw new Error('Google 로그인 및 인증이 필요합니다.');
  }
  return result.accessToken;
};

export const setCachedToken = (token: string) => {
  cachedAccessToken = token;
  if (typeof window !== 'undefined') {
    localStorage.setItem('google_access_token', token);
    localStorage.setItem('google_access_token_time', String(Date.now()));
  }
};

export const logout = async () => {
  await signOut(auth);
  clearAuthToken();
};

