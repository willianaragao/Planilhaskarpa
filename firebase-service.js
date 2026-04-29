// ─────────────────────────────────────────────────────────────────
// FIREBASE SERVICE — Skarpa Logística
// Centraliza todas as operações Firebase:
//   - Auth (login/logout com email)
//   - Firestore (categorias, formatação, condomínios)
//   - Storage (arquivo XLSX da planilha)
// ─────────────────────────────────────────────────────────────────

import { initializeApp }        from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth,
         signInWithEmailAndPassword,
         signOut,
         onAuthStateChanged }   from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { getFirestore,
         doc, getDoc, setDoc }  from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import { getStorage,
         ref, uploadBytes,
         getDownloadURL }       from "https://www.gstatic.com/firebasejs/10.12.2/firebase-storage.js";

// ── Inicialização ──────────────────────────────────────────────
const app       = initializeApp(window.FIREBASE_CONFIG);
const auth      = getAuth(app);
const db        = getFirestore(app);
const storage   = getStorage(app);

// ── AUTH ──────────────────────────────────────────────────────

/** Login com email + senha. Retorna Promise<UserCredential>. */
export function fbLogin(email, password) {
    return signInWithEmailAndPassword(auth, email, password);
}

/** Logout. */
export function fbLogout() {
    return signOut(auth);
}

/** Observa mudanças de auth. Chama callback(user | null). */
export function fbOnAuthChanged(callback) {
    return onAuthStateChanged(auth, callback);
}

/** Retorna o usuário autenticado atual ou null. */
export function fbCurrentUser() {
    return auth.currentUser;
}

// ── STORAGE (planilha XLSX) ────────────────────────────────────

/** Caminho do arquivo XLSX do usuário no Storage. */
function workbookPath(uid) {
    return `users/${uid}/planilha.xlsx`;
}

/** Salva o workbook (ArrayBuffer) no Firebase Storage. */
export async function fbSaveWorkbook(arrayBuffer) {
    const user = auth.currentUser;
    if (!user) throw new Error('Usuário não autenticado');

    const bytes = new Uint8Array(arrayBuffer);
    const storageRef = ref(storage, workbookPath(user.uid));
    await uploadBytes(storageRef, bytes, { contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/** Carrega o workbook do Firebase Storage. Retorna ArrayBuffer ou null. */
export async function fbLoadWorkbook() {
    const user = auth.currentUser;
    if (!user) return null;

    try {
        const storageRef = ref(storage, workbookPath(user.uid));
        const url = await getDownloadURL(storageRef);
        const res = await fetch(url);
        if (!res.ok) return null;
        return await res.arrayBuffer();
    } catch (e) {
        // Arquivo não existe ainda para este usuário
        if (e.code === 'storage/object-not-found') return null;
        throw e;
    }
}

// ── FIRESTORE (metadata do usuário) ───────────────────────────

function userDocRef(uid, collection) {
    return doc(db, 'users', uid, collection, 'data');
}

/** Salva categorias no Firestore. */
export async function fbSaveCategories(categories) {
    const user = auth.currentUser;
    if (!user) return;
    await setDoc(userDocRef(user.uid, 'categories'), { value: JSON.stringify(categories) });
}

/** Carrega categorias do Firestore. */
export async function fbLoadCategories() {
    const user = auth.currentUser;
    if (!user) return {};
    try {
        const snap = await getDoc(userDocRef(user.uid, 'categories'));
        if (!snap.exists()) return {};
        return JSON.parse(snap.data().value || '{}');
    } catch { return {}; }
}

/** Salva formatação no Firestore. */
export async function fbSaveFormatting(formatting) {
    const user = auth.currentUser;
    if (!user) return;
    await setDoc(userDocRef(user.uid, 'formatting'), { value: JSON.stringify(formatting) });
}

/** Carrega formatação do Firestore. */
export async function fbLoadFormatting() {
    const user = auth.currentUser;
    if (!user) return {};
    try {
        const snap = await getDoc(userDocRef(user.uid, 'formatting'));
        if (!snap.exists()) return {};
        return JSON.parse(snap.data().value || '{}');
    } catch { return {}; }
}

/** Salva condomínios no Firestore. */
export async function fbSaveCondominios(condominios) {
    const user = auth.currentUser;
    if (!user) return;
    await setDoc(userDocRef(user.uid, 'condominios'), { value: JSON.stringify(condominios) });
}

/** Carrega condomínios do Firestore. */
export async function fbLoadCondominios() {
    const user = auth.currentUser;
    if (!user) return [];
    try {
        const snap = await getDoc(userDocRef(user.uid, 'condominios'));
        if (!snap.exists()) return [];
        return JSON.parse(snap.data().value || '[]');
    } catch { return []; }
}
