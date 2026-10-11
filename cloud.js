// 初晴之森共用雲端層：Google 登入 + 成員名單權限 + Firestore 資料
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.14.1/firebase-app.js";
import {
  getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect,
  onAuthStateChanged, signOut
} from "https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js";
import {
  getFirestore, collection, doc, getDoc, addDoc, setDoc, deleteDoc, onSnapshot
} from "https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore.js";

const app = initializeApp({
  apiKey: "AIzaSyCar-w8B5h769FfIczVbQeQTosws9kyR8A",
  authDomain: "sunny-forest.firebaseapp.com",
  projectId: "sunny-forest",
  storageBucket: "sunny-forest.firebasestorage.app",
  messagingSenderId: "981586744223",
  appId: "1:981586744223:web:97e0ddd869f1e36e1123aa"
});
const auth = getAuth(app);
const db = getFirestore(app);
const provider = new GoogleAuthProvider();
provider.setCustomParameters({ prompt: "select_account" });

let started = false;

function panel() {
  let el = document.getElementById("sf-gate");
  if (!el) {
    el = document.createElement("div");
    el.id = "sf-gate";
    el.style.cssText = "position:fixed;inset:0;z-index:99999;background:#fbf6ef;color:#4d4a47;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px;text-align:center;font:16px/1.7 'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif";
    document.body.appendChild(el);
  }
  el.replaceChildren();
  el.style.display = "flex";
  return el;
}
function hidePanel() {
  const el = document.getElementById("sf-gate");
  if (el) el.style.display = "none";
}
function line(text, strong) {
  const p = document.createElement("p");
  p.style.margin = "6px 0";
  if (strong) p.style.fontSize = "20px", p.style.fontWeight = "700";
  p.textContent = text;
  return p;
}
function button(text, fn) {
  const b = document.createElement("button");
  b.textContent = text;
  b.style.cssText = "margin-top:18px;border:0;border-radius:999px;padding:12px 26px;font:inherit;font-size:16px;background:#86ab98;color:#fff;cursor:pointer";
  b.onclick = fn;
  return b;
}
async function login() {
  try {
    await signInWithPopup(auth, provider);
  } catch (e) {
    if (e.code === "auth/popup-closed-by-user" || e.code === "auth/cancelled-popup-request") return;
    if (e.code === "auth/popup-blocked" || e.code === "auth/operation-not-supported-in-this-environment") {
      return signInWithRedirect(auth, provider);
    }
    alert("登入失敗：" + (e.message || e.code));
  }
}
function showLogin() {
  const el = panel();
  el.append(
    line((document.title.split("｜")[0] || "初晴之森"), true),
    line("請使用 Google 帳號登入。"),
    line("只有管理者加入名單的成員能使用。"),
    button("使用 Google 登入", login)
  );
}
function showDenied(email) {
  const el = panel();
  el.append(
    line("尚未開通使用權限", true),
    line("帳號 " + (email || "（無）") + " 不在成員名單內。"),
    line("請聯絡管理者加入名單，或換一個帳號登入。"),
    button("換一個帳號", () => signOut(auth))
  );
}

// 只有名單內（members 集合裡有該信箱）的人才會進入頁面
function gate(onAllowed) {
  onAuthStateChanged(auth, async (user) => {
    if (!user) return showLogin();
    const email = (user.email || "").toLowerCase();
    let role = null;
    try {
      const m = await getDoc(doc(db, "members", email));
      if (m.exists()) role = m.data().role || "staff";
    } catch (e) {
      role = null;
    }
    if (!role) return showDenied(email);
    hidePanel();
    if (!started) {
      started = true;
      onAllowed({ email, role, admin: role === "admin" });
    }
  });
}

const safe = (p) => p.catch((e) => {
  console.error(e);
  alert("儲存失敗，請確認網路連線，或聯絡管理者確認您的權限");
});
const onErr = (e) => console.error(e);

export const SF = {
  gate,
  watch(name, cb) {
    return onSnapshot(collection(db, name),
      (snap) => cb(snap.docs.map((d) => ({ id: d.id, ...d.data() })), snap.metadata.hasPendingWrites),
      onErr);
  },
  watchDoc(name, id, cb) {
    return onSnapshot(doc(db, name, id),
      (snap) => cb(snap.exists() ? snap.data() : null, snap.metadata.hasPendingWrites),
      onErr);
  },
  add: (name, data) => safe(addDoc(collection(db, name), data)),
  addStrict: (name, data) => addDoc(collection(db, name), data),
  set: (name, id, data) => safe(setDoc(doc(db, name, id), data)),
  remove: (name, id) => safe(deleteDoc(doc(db, name, id))),
  logout: () => signOut(auth)
};
