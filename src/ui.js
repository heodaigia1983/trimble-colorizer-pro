// =============================================================================
// ui.js  —  MODULE: DOM controls (toolbar, selection panel, status bar)
// -----------------------------------------------------------------------------
// Pure view layer: no engine knowledge. renderer.js wires these callbacks to
// the viewer module. Keeps Mảnh 1 UI isolated from the 3D logic.
// =============================================================================

const el = {
  openBtn: document.getElementById('openFileBtn'),
  openXktBtn: document.getElementById('openXktBtn'),
  focusBtn: document.getElementById('focusBtn'),
  fadeBtn: document.getElementById('fadeBtn'),
  themeBtn: document.getElementById('themeBtn'),
  infoName: document.getElementById('infoName'),
  infoGuid: document.getElementById('infoGuid'),
  status: document.getElementById('statusText'),
  overlay: document.getElementById('viewerOverlay'),
};

/** Wire toolbar buttons to handlers. */
export function bindControls({ onOpen, onOpenXkt, onFocus, onFade, onTheme }) {
  el.openBtn.addEventListener('click', onOpen);
  el.openXktBtn.addEventListener('click', onOpenXkt);
  el.focusBtn.addEventListener('click', onFocus);
  el.fadeBtn.addEventListener('click', onFade);
  el.themeBtn.addEventListener('click', onTheme);
}

/** Status bar text, e.g. "Ready — 248 assemblies". */
export function setStatus(text) {
  el.status.textContent = text;
}

/**
 * Warn that an .ifc is large enough to risk freezing the renderer, and let the
 * user decide. Returns true if they choose to parse it live anyway.
 * @param {number} sizeMB  file size in MB (already rounded by the caller)
 */
export function confirmLargeIfc(sizeMB) {
  return window.confirm(
    `Model lớn (~${sizeMB} MB).\n\n` +
    `Parse .ifc trực tiếp trên UI có thể làm treo app. Nên convert sang XKT ` +
    `trước rồi mở .xkt:\n\n    npm run convert -- "đường-dẫn-file.ifc"\n\n` +
    `→ tạo file .xkt trong thư mục /models, sau đó bấm "Mở file XKT".\n\n` +
    `Bấm OK để vẫn thử parse .ifc trực tiếp (có thể đơ), hoặc Cancel để hủy.`
  );
}

/** Center overlay for transient messages; pass '' to hide. */
export function setOverlay(text) {
  el.overlay.textContent = text || '';
  el.overlay.style.display = text ? 'grid' : 'none';
}

/** Fill the selection panel with a picked object's Name + GUID. */
export function showSelection({ name, guid }) {
  el.infoName.textContent = name;
  el.infoGuid.textContent = guid;
}

/** Reset the selection panel to its empty state. */
export function clearSelection() {
  el.infoName.textContent = 'Chưa chọn object.';
  el.infoGuid.textContent = '—';
}

/** Enable/disable the buttons that act on a selection (Focus, Fade). */
export function setSelectionEnabled(on) {
  el.focusBtn.disabled = !on;
  el.fadeBtn.disabled = !on;
}

export function setFadeLabel(on) {
  el.fadeBtn.textContent = on ? 'Fade: ON' : 'Fade';
}

export function setThemeLabel(isDark) {
  // Button shows the mode it will switch TO.
  el.themeBtn.textContent = isDark ? 'Light' : 'Dark';
}
