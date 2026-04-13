import { getAuth, GoogleAuthProvider, signInWithPopup } from "firebase/auth";
import { getFirebaseApp, getMissingFirebaseEnvVars, hasFirebaseConfig } from "./firebase/config";

function setStatus(message) {
    const el = document.getElementById("status");
    if (el) el.innerText = message;
}

function messageParent(payload) {
    try {
        Office.context.ui.messageParent(JSON.stringify(payload));
    } catch (e) {
        setStatus("לא ניתן לשלוח תשובה ל-Taskpane: " + (e?.message || String(e)));
    }
}

function messageFromError(error, fallback) {
    if (!error) return fallback;
    const text = String(error.message || fallback || "");
    if (text.includes("auth/popup-blocked")) {
        return "חלון ההתחברות נחסם. אשר Pop-ups עבור Word/localhost ונסה שוב.";
    }
    if (text.includes("missing initial state")) {
        return "נחסם אחסון דפדפן (sessionStorage/cookies). אשר Cookies ונסה שוב.";
    }
    return text || fallback;
}

function buildGoogleProvider() {
    const provider = new GoogleAuthProvider();
    provider.setCustomParameters({ prompt: "select_account" });
    return provider;
}

function sendAuthSuccess(result) {
    const googleCredential = GoogleAuthProvider.credentialFromResult(result);
    const googleIdToken = googleCredential?.idToken || "";
    const googleAccessToken = googleCredential?.accessToken || "";

    messageParent({
        type: "google-auth-success",
        googleIdToken,
        googleAccessToken,
        hasFirebaseSession: !!result?.user,
        email: result.user.email || "",
        uid: result.user.uid,
    });
}

async function runGoogleSignInInDialog() {
    if (!hasFirebaseConfig()) {
        const missing = getMissingFirebaseEnvVars();
        const missingText = missing.length ? `Missing env vars: ${missing.join(", ")}` : "Firebase config is missing";
        setStatus("Firebase לא מוגדר בסביבה. ודא שקיים .env.local עם VITE_FIREBASE_* והפעל מחדש את npm run dev.");
        messageParent({ type: "google-auth-error", message: missingText });
        return;
    }

    try {
        const auth = getAuth(getFirebaseApp());
        const provider = buildGoogleProvider();

        // עדיף Popup קודם כי הוא לא תלוי ב-sessionStorage של Redirect בכל סביבה.
        setStatus("פותח חלון התחברות Google...");
        const popupResult = await signInWithPopup(auth, provider);
        setStatus("התחברות הצליחה. חוזר לתוסף...");
        sendAuthSuccess(popupResult);
    } catch (e) {
        const message = messageFromError(e, "Google sign-in failed");
        setStatus("שגיאה בהתחברות: " + message);
        messageParent({ type: "google-auth-error", message });
    }
}

Office.onReady(() => {
    const startBtn = document.getElementById("startGoogleAuth");
    setStatus("לחץ על הכפתור כדי להתחבר עם Google.");
    if (startBtn) {
        startBtn.onclick = async () => {
            startBtn.disabled = true;
            await runGoogleSignInInDialog();
            startBtn.disabled = false;
        };
    }
});
