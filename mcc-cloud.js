import { initializeApp } from 'https://www.gstatic.com/firebasejs/12.19.0/firebase-app.js';
import { getAuth, GoogleAuthProvider, onAuthStateChanged, signInWithPopup, signOut } from 'https://www.gstatic.com/firebasejs/12.19.0/firebase-auth.js';
import { Bytes, collection, deleteDoc, doc, getDoc, getDocs, getFirestore, limit, orderBy, query, serverTimestamp, setDoc } from 'https://www.gstatic.com/firebasejs/12.19.0/firebase-firestore.js';

// Firebase Web configuration is public. Access is controlled by Firestore rules.
const app = initializeApp({
  apiKey: 'AIzaSyAqkG-HvR_C7cG-v8eT8_ePI14M7_T2zM0',
  authDomain: 'trimble-model-control-center.firebaseapp.com',
  projectId: 'trimble-model-control-center',
  appId: '1:796142431592:web:cca91f7dcef28b3b20badd'
});
const auth = getAuth(app);
const db = getFirestore(app);
const OWNER = 'heodaigia1983@gmail.com';
const CHUNK = 180000;
let projectId = null;
let pendingDraft = null;
const pendingImports = [];
let draftTimer = null;
let savingDraft = false;
let activeUser = null;

function status(message, bad = false) {
  const el = document.getElementById('cloudStatus');
  if (el) { el.textContent = message; el.style.color = bad ? '#d93025' : '#5f6368'; }
}
function safeId(value) { return encodeURIComponent(String(value || '')); }
function projectRef() { return doc(db, 'mccProjects', safeId(projectId)); }
function recordRef(kind, id) { return doc(projectRef(), kind, safeId(id)); }
function requireReady() {
  if (!activeUser || activeUser.email !== OWNER) throw new Error('Đăng nhập đúng tài khoản chủ Firebase để đồng bộ.');
  if (!projectId) throw new Error('Trimble chưa trả Project ID; chưa thể lưu dữ liệu lên Firebase.');
}
async function resolveProject() {
  const api = await window.getAPI();
  const project = await api.project.getProject();
  if (!project || !project.id) throw new Error('Không đọc được Project ID từ Trimble.');
  projectId = String(project.id);
  document.getElementById('cloudProject').textContent = project.name || projectId;
  return projectId;
}
async function encode(bytes) {
  if (!('CompressionStream' in window)) return { bytes, encoding: 'raw' };
  const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream('gzip'));
  return { bytes: new Uint8Array(await new Response(stream).arrayBuffer()), encoding: 'gzip' };
}
async function decode(bytes, encoding) {
  if (encoding !== 'gzip') return bytes;
  if (!('DecompressionStream' in window)) throw new Error('Trình duyệt không hỗ trợ giải nén dữ liệu Firebase.');
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
function revisionId() { return Date.now().toString(36) + '-' + crypto.randomUUID(); }
async function attachExternalIds(state) {
  if (!Array.isArray(state.colorLedger) || !state.colorLedger.length) return;
  const api = await window.getAPI(), byModel = new Map();
  for (const row of state.colorLedger) {
    if (row.value.externalId) continue;
    const at = row.key.lastIndexOf(':'), modelId = row.key.slice(0, at), runtimeId = Number(row.key.slice(at + 1));
    if (at < 1 || !Number.isFinite(runtimeId)) continue;
    if (!byModel.has(modelId)) byModel.set(modelId, []);
    byModel.get(modelId).push({ key: row.key, runtimeId, row });
  }
  for (const [modelId, rows] of byModel) {
    try {
      const result = await window.convertObjectIdsSafely(api, modelId, rows);
      rows.forEach(item => { const externalId = result.ids.get(item.key); if (externalId) item.row.value.externalId = externalId; });
    } catch (_) { /* The summary remains usable even if some external IDs are unavailable. */ }
  }
}
async function remapExternalIds(state) {
  if (!Array.isArray(state.colorLedger) || !state.colorLedger.length || state.colorLedger.some(row => !row.value.externalId)) return false;
  const api = await window.getAPI(), byModel = new Map(), replacement = new Map();
  for (const row of state.colorLedger) {
    const at = row.key.lastIndexOf(':'), modelId = row.key.slice(0, at);
    if (at < 1) return false;
    if (!byModel.has(modelId)) byModel.set(modelId, []);
    byModel.get(modelId).push(row);
  }
  try {
    for (const [modelId, rows] of byModel) {
      for (let i = 0; i < rows.length; i += 200) {
        const piece = rows.slice(i, i + 200);
        const ids = await api.viewer.convertToObjectRuntimeIds(modelId, piece.map(row => row.value.externalId));
        if (!Array.isArray(ids) || ids.length !== piece.length || ids.some(id => id == null || !Number.isFinite(Number(id)))) return false;
        piece.forEach((row, index) => replacement.set(row, modelId + ':' + Number(ids[index])));
      }
    }
  } catch (_) { return false; }
  if (new Set(replacement.values()).size !== state.colorLedger.length) return false;
  state.colorLedger = state.colorLedger.map(row => ({ key: replacement.get(row), value: row.value }));
  return true;
}
async function writePayload(ref, bytes, metadata) {
  const packed = await encode(bytes);
  const rev = revisionId();
  const count = Math.ceil(packed.bytes.length / CHUNK) || 1;
  if (count > 500) throw new Error('File vượt 90 MB sau nén; cần Cloud Storage cho file lớn hơn.');
  for (let i = 0; i < count; i++) {
    const part = packed.bytes.slice(i * CHUNK, (i + 1) * CHUNK);
    await setDoc(doc(ref, 'payloads', rev, 'chunks', String(i)), { data: Bytes.fromUint8Array(part) });
  }
  const previous = await getDoc(ref);
  const old = previous.exists() ? previous.data() : null;
  await setDoc(ref, { ...metadata, revision: rev, chunks: count, encoding: packed.encoding, updatedAt: serverTimestamp() });
  // Keep one previous revision as a recoverable copy; remove the revision before it.
  if (old && old.previousRevision && Number.isInteger(old.previousChunks)) {
    for (let i = 0; i < old.previousChunks; i++) {
      await deleteDoc(doc(ref, 'payloads', old.previousRevision, 'chunks', String(i))).catch(() => {});
    }
  }
  if (old && old.revision) await setDoc(ref, { previousRevision: old.revision, previousChunks: old.chunks }, { merge: true });
  return rev;
}
async function readPayload(ref, meta) {
  if (!meta || !meta.revision || !Number.isInteger(meta.chunks) || meta.chunks < 1 || meta.chunks > 500) throw new Error('Bản lưu Firebase thiếu danh sách phần dữ liệu.');
  const parts = [];
  for (let i = 0; i < meta.chunks; i++) {
    const part = await getDoc(doc(ref, 'payloads', meta.revision, 'chunks', String(i)));
    if (!part.exists() || !part.data().data) throw new Error('Bản lưu Firebase thiếu phần ' + (i + 1) + '/' + meta.chunks + '.');
    parts.push(part.data().data.toUint8Array());
  }
  const merged = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let offset = 0;
  for (const part of parts) { merged.set(part, offset); offset += part.length; }
  return decode(merged, meta.encoding);
}
async function login() {
  try {
    await signInWithPopup(auth, new GoogleAuthProvider());
  } catch (error) { status('Không đăng nhập được Firebase: ' + error.message, true); }
}
async function logout() { await signOut(auth); }
function scheduleDraft() {
  pendingDraft = true;
  if (!activeUser) { status('Chưa đăng nhập: dữ liệu mới đang ở trên máy anh.', true); return; }
  clearTimeout(draftTimer);
  draftTimer = setTimeout(saveDraft, 2500);
}
async function saveDraft() {
  if (savingDraft) { pendingDraft = true; return; }
  if (!pendingDraft) return;
  savingDraft = true;
  pendingDraft = false;
  try {
    await resolveProject(); requireReady();
    const payload = window.mccCloudCapture('Bản đang tô màu');
    await attachExternalIds(payload.state);
    const ref = recordRef('drafts', activeUser.uid);
    await writePayload(ref, new TextEncoder().encode(JSON.stringify(payload.state)), { name: payload.state.name, groups: payload.groups, viewId: window._currentViewId || null, ownerUid: activeUser.uid, kind: 'draft' });
    status('Đã đồng bộ bản đang thao tác lên Firebase.');
    await refreshDashboard();
  } catch (error) { pendingDraft = true; status('Chưa đồng bộ: ' + error.message, true); }
  finally { savingDraft = false; if (pendingDraft && activeUser) { clearTimeout(draftTimer); draftTimer = setTimeout(saveDraft, 10000); } }
}
async function saveView(viewId, payload) {
  await resolveProject(); requireReady();
  await attachExternalIds(payload.state);
  await writePayload(recordRef('views', viewId), new TextEncoder().encode(JSON.stringify(payload.state)), { name: payload.state.name, groups: payload.groups, ownerUid: activeUser.uid, kind: 'view', viewId: String(viewId) });
  status('View ' + payload.state.name + ' đã lưu trên Firebase.');
  await refreshDashboard();
}
async function loadView(viewId) {
  await resolveProject(); requireReady();
  const ref = recordRef('views', viewId), snapshot = await getDoc(ref);
  if (!snapshot.exists()) return null;
  const meta = snapshot.data();
  const state = JSON.parse(new TextDecoder().decode(await readPayload(ref, meta)));
  const remapped = await remapExternalIds(state);
  return { state, groups: meta.groups || [], remapped };
}
async function saveImport(file, slot, guidCount) {
  await resolveProject(); requireReady();
  const id = crypto.randomUUID(), ref = recordRef('imports', id);
  await writePayload(ref, new Uint8Array(await file.arrayBuffer()), { name: file.name, size: file.size, lastModified: file.lastModified, slot, guidCount, ownerUid: activeUser.uid, kind: 'import' });
  status('Đã lưu file ' + file.name + ' lên Firebase.');
  await refreshDashboard();
  return id;
}
function queueImport(file, slot, guidCount) {
  if (activeUser) return saveImport(file, slot, guidCount);
  status('File ' + file.name + ' đang chờ đăng nhập để lưu Firebase.', true);
  return new Promise((resolve, reject) => pendingImports.push({ file, slot, guidCount, resolve, reject }));
}
async function downloadImport(id) {
  try {
    await resolveProject(); requireReady();
    const ref = recordRef('imports', id), snapshot = await getDoc(ref);
    if (!snapshot.exists()) throw new Error('Không tìm thấy file.');
    const meta = snapshot.data(), bytes = await readPayload(ref, meta);
    const anchor = document.createElement('a');
    anchor.href = URL.createObjectURL(new Blob([bytes], { type: 'application/octet-stream' }));
    anchor.download = meta.name || 'import.xsr';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(anchor.href), 1000);
  } catch (error) { status('Không tải được file: ' + error.message, true); }
}
async function restoreDraft() {
  try {
    await resolveProject(); requireReady();
    const ref = recordRef('drafts', activeUser.uid), snapshot = await getDoc(ref);
    if (!snapshot.exists()) throw new Error('Không tìm thấy bản đang thao tác.');
    const meta = snapshot.data();
    const state = JSON.parse(new TextDecoder().decode(await readPayload(ref, meta)));
    const remapped = await remapExternalIds(state);
    await window.mccRestoreCloudDraft(state, meta.groups || [], remapped);
    status(remapped || !state.colorLedger.length ? 'Đã khôi phục bản đang thao tác.' : 'Đã khôi phục số liệu; model chưa xác nhận đủ ID để tô lại màu.', !remapped && !!state.colorLedger.length);
  } catch (error) { status('Không khôi phục được bản đang thao tác: ' + error.message, true); }
}
function esc(value) { return String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch]); }
function fmt(value) { return Number(value).toLocaleString('vi-VN', { maximumFractionDigits: 3 }); }
async function refreshDashboard() {
  const el = document.getElementById('cloudDashboard');
  if (!el || !activeUser) return;
  try {
    await resolveProject(); requireReady();
    const views = await getDocs(query(collection(projectRef(), 'views'), orderBy('updatedAt', 'desc'), limit(40)));
    const imports = await getDocs(query(collection(projectRef(), 'imports'), orderBy('updatedAt', 'desc'), limit(10)));
    const draft = await getDoc(recordRef('drafts', activeUser.uid));
    let html = draft.exists() ? '<button type="button" id="cloudRestoreDraft" class="text-blue-600 underline mb-2">Khôi phục bản đang tô (thay trạng thái hiện tại)</button>' : '';
    html += '<div class="font-semibold mb-1">View đã lưu (' + views.size + ' gần nhất)</div>';
    if (views.empty) html += '<div>Chưa có View trên Firebase.</div>';
    views.forEach(snapshot => {
      const item = snapshot.data(), groups = Array.isArray(item.groups) ? item.groups : [];
      const volume = groups.length && groups.every(g => g.volumeComplete) ? fmt(groups.reduce((n, g) => n + Number(g.volumeM3 || 0), 0)) + ' m³' : 'chưa đủ V';
      const weight = groups.length && groups.every(g => g.weightComplete) ? fmt(groups.reduce((n, g) => n + Number(g.weightKg || 0), 0) / 1000) + ' tấn' : 'chưa đủ KL';
      html += '<div class="border-t py-1"><button type="button" class="text-blue-600 underline" data-cloud-view="' + esc(item.viewId) + '">' + esc(item.name) + '</button><div>' + groups.map(g => '<span style="display:inline-block;width:9px;height:9px;background:' + (/^#[0-9A-Fa-f]{6}$/.test(g.color) ? g.color : '#999') + '"></span> ' + esc(g.color) + ' · ' + fmt(g.count) + ' cấu kiện').join(' · ') + '</div><div>' + volume + ' · ' + weight + '</div></div>';
    });
    html += '<div class="font-semibold mt-2 mb-1">File đã nạp (' + imports.size + ' gần nhất)</div>';
    if (imports.empty) html += '<div>Chưa có file trên Firebase.</div>';
    imports.forEach(snapshot => { const item = snapshot.data(); html += '<div class="border-t py-1"><button type="button" class="text-blue-600 underline" data-cloud-import="' + esc(snapshot.id) + '">' + esc(item.name) + '</button>' + (item.slot ? ' · ' + fmt(item.guidCount) + ' GUID' : ' · file lưu riêng') + '</div>'; });
    el.innerHTML = html;
    el.querySelector('#cloudRestoreDraft')?.addEventListener('click', restoreDraft);
    el.querySelectorAll('[data-cloud-view]').forEach(button => button.addEventListener('click', () => window.loadView(button.dataset.cloudView)));
    el.querySelectorAll('[data-cloud-import]').forEach(button => button.addEventListener('click', () => downloadImport(button.dataset.cloudImport)));
  } catch (error) { el.textContent = 'Chưa tải được dashboard: ' + error.message; }
}
onAuthStateChanged(auth, async user => {
  activeUser = user && user.email === OWNER ? user : null;
  const loginButton = document.getElementById('cloudLogin');
  const logoutButton = document.getElementById('cloudLogout');
  if (loginButton) loginButton.hidden = !!activeUser;
  if (logoutButton) logoutButton.hidden = !activeUser;
  status(activeUser ? 'Đã đăng nhập Firebase: ' + activeUser.email : 'Chưa đăng nhập Firebase. Dữ liệu chỉ đang lưu trên máy.');
  if (activeUser) {
    try {
      await resolveProject(); await refreshDashboard();
      while (pendingImports.length) {
        const item = pendingImports.shift();
        try { item.resolve(await saveImport(item.file, item.slot, item.guidCount)); }
        catch (error) { item.reject(error); }
      }
      if (pendingDraft) saveDraft();
      if (window._currentViewId && window.loadLedger) await window.loadLedger(window._currentViewId);
    } catch (error) { status(error.message, true); }
  }
});
document.getElementById('cloudLogin')?.addEventListener('click', login);
document.getElementById('cloudLogout')?.addEventListener('click', logout);
document.getElementById('cloudRefresh')?.addEventListener('click', refreshDashboard);
document.getElementById('cloudChooseFile')?.addEventListener('click', () => document.getElementById('cloudFile').click());
document.getElementById('cloudFile')?.addEventListener('change', event => {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  queueImport(file, 0, 0).catch(error => status('Chưa lưu được file: ' + error.message, true));
  event.target.value = '';
});
window.MccCloud = { scheduleDraft, saveView, loadView, queueImport, refreshDashboard };
