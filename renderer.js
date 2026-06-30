// =============================================================================
// renderer.js  —  ENTRY POINT (glue)
// -----------------------------------------------------------------------------
// Wires the three modules together. Holds only the small bit of app state that
// links a user gesture to a viewer operation (current selection, fade flag,
// theme). All heavy lifting lives in src/viewer.js, src/file-source.js, src/ui.js.
// =============================================================================

import { createViewerCore } from './src/viewer.js';
import * as files from './src/file-source.js';
import * as ui from './src/ui.js';

let core = null;          // viewer facade (src/viewer.js)
let selectedId = null;    // currently picked object id (== IFC GUID)
let fadeOn = false;       // Fade/X-ray toggle state
let isDark = true;        // viewer background theme

// --- IFC loading (live parse — small models) ---------------------------------

async function loadBuffer(buffer, path) {
  resetSelection();
  ui.setOverlay(`Đang load IFC: ${files.fileName(path)}…`);
  try {
    const { objectCount } = await core.loadIFC(buffer);
    ui.setOverlay('');
    ui.setStatus(`Ready — ${objectCount} assemblies`);
    console.log(`[viewer] loaded ${files.fileName(path)} — ${objectCount} objects`);
  } catch (err) {
    console.error('[viewer] IFC load failed:', err && err.stack ? err.stack : err);
    ui.setOverlay('Không load được IFC — xem console.');
  }
}

async function onOpen() {
  const f = await files.openIfcFromDisk();
  if (!f) return;
  // Big .ifc would freeze the renderer if parsed live — warn and let the user
  // decide instead of auto-parsing. The .xkt path (convert.js) is the fix.
  if (f.data.byteLength > files.IFC_WARN_BYTES) {
    const sizeMB = Math.round(f.data.byteLength / (1024 * 1024));
    if (!ui.confirmLargeIfc(sizeMB)) {
      ui.setStatus('Đã hủy mở IFC lớn — hãy convert sang XKT trước.');
      return;
    }
  }
  await loadBuffer(f.data, f.path);
}

// --- XKT loading (fast path — big models, offline-converted) ------------------

async function loadXktBuffer(buffer, path) {
  resetSelection();
  ui.setOverlay(`Đang load XKT: ${files.fileName(path)}…`);
  try {
    const { objectCount } = await core.loadXKT(buffer);
    ui.setOverlay('');
    ui.setStatus(`Ready — ${objectCount} assemblies`);
    console.log(`[viewer] loaded ${files.fileName(path)} — ${objectCount} objects (XKT)`);
  } catch (err) {
    console.error('[viewer] XKT load failed:', err && err.stack ? err.stack : err);
    ui.setOverlay('Không load được XKT — xem console.');
  }
}

async function onOpenXkt() {
  const f = await files.openXktFromDisk();
  if (f) await loadXktBuffer(f.data, f.path);
}

// --- selection ---------------------------------------------------------------

function resetSelection() {
  if (selectedId) core.setHighlighted(selectedId, false);
  if (fadeOn) {
    core.setXRayOthers(null, false);
    fadeOn = false;
    ui.setFadeLabel(false);
  }
  selectedId = null;
  ui.clearSelection();
  ui.setSelectionEnabled(false);
}

function onPick(canvasCoords) {
  const hit = core.pick(canvasCoords);

  // Drop previous highlight first.
  if (selectedId) core.setHighlighted(selectedId, false);

  if (!hit) {
    // Clicked empty space — clear selection but keep Fade state if it was on.
    selectedId = null;
    ui.clearSelection();
    ui.setSelectionEnabled(false);
    if (fadeOn) core.setXRayOthers(null, true);
    return;
  }

  selectedId = hit.id;
  core.setHighlighted(selectedId, true);
  ui.showSelection(hit);
  ui.setSelectionEnabled(true);

  // Keep the newly selected object sharp if Fade is active.
  if (fadeOn) core.setXRayOthers(selectedId, true);
}

// --- toolbar actions ---------------------------------------------------------

function onFocus() {
  if (selectedId) core.focusObject(selectedId);
}

function onFade() {
  if (!selectedId) return;
  fadeOn = !fadeOn;
  core.setXRayOthers(selectedId, fadeOn);
  ui.setFadeLabel(fadeOn);
}

function onTheme() {
  isDark = !isDark;
  core.setBackground(isDark ? core.BG_DARK : core.BG_LIGHT);
  document.body.classList.toggle('light-theme', !isDark);
  ui.setThemeLabel(isDark);
}

// --- bootstrap ---------------------------------------------------------------

async function main() {
  ui.setOverlay('Đang khởi tạo viewer…');
  core = await createViewerCore('viewerCanvas');

  // mouseclicked fires the canvas coords as [x, y].
  core.viewer.scene.input.on('mouseclicked', onPick);

  ui.bindControls({ onOpen, onOpenXkt, onFocus, onFade, onTheme });
  ui.setSelectionEnabled(false);
  ui.setThemeLabel(isDark);
  ui.setStatus('Ready — no model loaded');

  // Auto-load the bundled sample so the pipeline is testable on first launch.
  const sample = await files.loadSampleIfc();
  if (sample) {
    await loadBuffer(sample.data, sample.path);
  } else {
    ui.setOverlay('');
  }
}

main().catch((err) => {
  console.error('[viewer] init failed:', err);
  ui.setOverlay('Khởi tạo thất bại — xem console.');
});
