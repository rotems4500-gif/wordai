import { getApp, getApps, initializeApp } from "firebase/app";

let cachedConfig = null;

const REQUIRED_ENV_MAP = {
  apiKey: "VITE_FIREBASE_API_KEY",
  authDomain: "VITE_FIREBASE_AUTH_DOMAIN",
  projectId: "VITE_FIREBASE_PROJECT_ID",
  storageBucket: "VITE_FIREBASE_STORAGE_BUCKET",
  appId: "VITE_FIREBASE_APP_ID",
};

function readFirebaseConfig() {
  if (cachedConfig) return cachedConfig;

  const cfg = {
    apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
    authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
    projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
    storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
    messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
    appId: import.meta.env.VITE_FIREBASE_APP_ID,
  };

  const requiredKeys = Object.keys(REQUIRED_ENV_MAP);
  const missing = requiredKeys.filter((key) => !String(cfg[key] || "").trim());
  if (missing.length > 0) {
    cachedConfig = null;
    return null;
  }

  cachedConfig = cfg;
  return cachedConfig;
}

export function hasFirebaseConfig() {
  return !!readFirebaseConfig();
}

export function getMissingFirebaseEnvVars() {
  const cfg = {
    apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
    authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
    projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
    storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
    appId: import.meta.env.VITE_FIREBASE_APP_ID,
  };
  return Object.entries(REQUIRED_ENV_MAP)
    .filter(([key]) => !String(cfg[key] || "").trim())
    .map(([, envName]) => envName);
}

export function getFirebaseApp() {
  const cfg = readFirebaseConfig();
  if (!cfg) {
    throw new Error("Firebase config is missing. Add VITE_FIREBASE_* values to your environment.");
  }
  return getApps().length ? getApp() : initializeApp(cfg);
}
