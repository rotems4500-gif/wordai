import { getAuth, GoogleAuthProvider, onAuthStateChanged, signInWithCredential, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { collection, getDocs, limit, query, where } from "firebase/firestore";
import { getBlob, getDownloadURL, getStorage, ref } from "firebase/storage";
import { getFirestore } from "firebase/firestore";
import { getFirebaseApp, hasFirebaseConfig } from "./config";

let auth = null;
let db = null;
let storage = null;

function ensureFirebaseClients() {
    if (!hasFirebaseConfig()) return null;
    if (auth && db && storage) return { auth, db, storage };

    const app = getFirebaseApp();
    auth = getAuth(app);
    db = getFirestore(app);
    storage = getStorage(app);
    return { auth, db, storage };
}

function normalizeCourse(docSnap) {
    const data = docSnap.data() || {};
    return {
        id: docSnap.id,
        title: data.title || data.name || `קורס ${docSnap.id}`,
        ...data,
    };
}

function normalizeMaterial(docSnap, courseId = "") {
    const data = docSnap.data() || {};
    return {
        id: docSnap.id,
        courseId: data.courseId || courseId || "",
        title: data.title || data.name || data.fileName || docSnap.id,
        type: String(data.type || "").toLowerCase(),
        storagePath: data.storagePath || data.path || "",
        storageUrl: data.storageUrl || data.url || "",
        updatedAt: data.updatedAt || null,
        ...data,
    };
}

export function isCloudAvailable() {
    return hasFirebaseConfig();
}

export function onCloudAuthChange(callback) {
    const clients = ensureFirebaseClients();
    if (!clients) {
        callback(null);
        return () => {};
    }
    return onAuthStateChanged(clients.auth, callback);
}

export async function cloudSignIn(email, password) {
    const clients = ensureFirebaseClients();
    if (!clients) throw new Error("Firebase config is missing.");
    const result = await signInWithEmailAndPassword(clients.auth, email, password);
    return result.user;
}

export async function cloudSignInWithGoogleIdToken(idToken, accessToken = "") {
    const clients = ensureFirebaseClients();
    if (!clients) throw new Error("Firebase config is missing.");
    if (!idToken && !accessToken) throw new Error("Missing Google auth token.");

    const credential = GoogleAuthProvider.credential(idToken || null, accessToken || null);
    const result = await signInWithCredential(clients.auth, credential);
    return result.user;
}

export async function cloudSignOut() {
    const clients = ensureFirebaseClients();
    if (!clients) return;
    await signOut(clients.auth);
}

export async function fetchCoursesForUser(user) {
    const clients = ensureFirebaseClients();
    if (!clients) return [];

    const snapshot = await getDocs(collection(clients.db, "courses"));
    const all = snapshot.docs.map(normalizeCourse);
    if (!user?.uid) return all;

    return all.filter((course) => {
        const owner = course.ownerId || course.userId || course.uid || "";
        if (!owner) return true;
        return owner === user.uid;
    });
}

export async function fetchMaterialsForCourse(courseId) {
    const clients = ensureFirebaseClients();
    if (!clients || !courseId) return [];

    const topQuery = query(
        collection(clients.db, "materials"),
        where("courseId", "==", courseId),
        limit(300)
    );

    const topSnapshot = await getDocs(topQuery);
    const nestedSnapshot = await getDocs(collection(clients.db, "courses", courseId, "materials"));
    const all = [
        ...topSnapshot.docs.map((docSnap) => normalizeMaterial(docSnap, courseId)),
        ...nestedSnapshot.docs.map((docSnap) => normalizeMaterial(docSnap, courseId)),
    ];

    const unique = new Map();
    all.forEach((material) => {
        const key = `${material.courseId || courseId}:${material.id}`;
        if (!unique.has(key)) unique.set(key, material);
    });
    return [...unique.values()];
}

export async function resolveMaterialDownloadUrl(material) {
    const clients = ensureFirebaseClients();
    if (!clients) throw new Error("Firebase config is missing.");

    if (material.storageUrl) return material.storageUrl;
    if (!material.storagePath) throw new Error("Missing storagePath/storageUrl for material.");

    return getDownloadURL(ref(clients.storage, material.storagePath));
}

export async function downloadMaterialBlob(material) {
    const clients = ensureFirebaseClients();
    if (!clients) throw new Error("Firebase config is missing.");

    if (material.storagePath) {
        try {
            return await getBlob(ref(clients.storage, material.storagePath));
        } catch (e) {
            if (!material.storageUrl) throw e;
        }
    }

    const url = await resolveMaterialDownloadUrl(material);
    const res = await fetch(url);
    if (!res.ok) throw new Error(`Material download failed: ${res.status}`);
    return res.blob();
}
