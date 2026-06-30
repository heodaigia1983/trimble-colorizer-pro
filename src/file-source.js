// =============================================================================
// file-source.js  —  MODULE: file input (disk <-> renderer)
// -----------------------------------------------------------------------------
// Thin bridge over the Electron IPC that preload.js exposes as `window.api`.
// All file reading happens in the main process (Node fs); the renderer only
// receives ArrayBuffers — everything stays local, no network.
//
// [EXT] Mảnh 2 (Excel) will add an `openExcelFromDisk()` here in the same shape.
// =============================================================================

// Above this size, parsing a raw .ifc on the renderer thread risks freezing the
// app — we warn and steer the user to convert.js + the .xkt path instead.
export const IFC_WARN_BYTES = 50 * 1024 * 1024; // ~50 MB

function toArrayBuffer(data) {
  if (data instanceof ArrayBuffer) return data;
  if (ArrayBuffer.isView(data)) return data.buffer;
  return data;
}

/** Opens the OS file dialog and returns the chosen .ifc as an ArrayBuffer. */
export async function openIfcFromDisk() {
  const res = await window.api.openIfcFile();
  if (!res) return null;
  return { path: res.path, data: toArrayBuffer(res.data) };
}

/**
 * Opens the OS file dialog and returns the chosen .xkt as an ArrayBuffer.
 * This is the fast path for big models (offline-converted by convert.js).
 */
export async function openXktFromDisk() {
  const res = await window.api.openXktFile();
  if (!res) return null;
  return { path: res.path, data: toArrayBuffer(res.data) };
}

/** Loads the bundled sample model (sample/Duplex.ifc), if present. */
export async function loadSampleIfc() {
  const res = await window.api.readSampleIfc();
  if (!res) return null;
  return { path: res.path, data: toArrayBuffer(res.data) };
}

/** Strips a path down to its file name (Windows or POSIX separators). */
export function fileName(p) {
  return p ? p.split(/[\\/]/).pop() : '';
}
