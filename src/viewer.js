// =============================================================================
// viewer.js  —  MODULE: 3D engine (xeokit) + IFC parsing (web-ifc)
// -----------------------------------------------------------------------------
// Verified against installed versions:
//   - @xeokit/xeokit-sdk 2.6.111  (WebIFCLoaderPlugin, CameraControl, cameraFlight)
//   - web-ifc 0.0.77              (IfcAPI, loaded as IIFE -> window.WebIFC)
//
// This module owns the Viewer and exposes a small, stable API so that later
// pieces can plug in WITHOUT touching the engine wiring:
//   * Mảnh 2 (tô màu theo Excel) -> dùng `core.viewer.scene.setObjectsColorized([id], rgb)`
//   * Mảnh 3 (timeline 4D)        -> dùng `core.viewer.scene.setObjectsVisible([id], bool)`
// Extension hooks are marked with  // [EXT] .
// =============================================================================

import { Viewer, WebIFCLoaderPlugin, XKTLoaderPlugin } from '../node_modules/@xeokit/xeokit-sdk/dist/xeokit-sdk.es.js';
// web-ifc is pinned to 0.0.51 — the version xeokit 2.6.111's WebIFCLoaderPlugin
// is built/tested against (newer web-ifc returns geometry that overflows
// xeokit's VBO buffer compile -> "RangeError: offset is out of bounds").
// Imported as an ES module (0.0.51 ships no IIFE build); served over app://
// so its WASM loads via fetch.
import * as WebIFC from '../node_modules/web-ifc/web-ifc-api.js';

// Viewer background presets for the Dark/Light toggle (RGB, 0..1).
export const BG_DARK = [0.027, 0.039, 0.078];
export const BG_LIGHT = [0.901, 0.929, 0.964];

/**
 * Boots web-ifc, creates the xeokit Viewer, and returns a thin facade.
 * @param {string} canvasId  id of the <canvas> element.
 */
export async function createViewerCore(canvasId) {
  // --- web-ifc engine (parses .ifc -> geometry + metadata) -------------------
  // Point web-ifc's WASM loader at an explicit app:// URL. The app:// scheme
  // (see main.js) supports fetch/instantiateStreaming, which file:// does not.
  // crossOriginIsolated is false here, so web-ifc loads the single-thread WASM.
  const ifcAPI = new WebIFC.IfcAPI();
  await ifcAPI.Init((file) => `app://bundle/node_modules/web-ifc/${file}`);

  // --- xeokit Viewer ---------------------------------------------------------
  const viewer = new Viewer({
    canvasId,
    transparent: false,
  });

  // Solid background (so Dark/Light toggle is visible).
  viewer.scene.canvas.backgroundColorFromAmbientLight = false;
  viewer.scene.canvas.backgroundColor = BG_DARK;

  // CameraControl is auto-created by Viewer. Orbit + smooth follow-pointer zoom.
  viewer.cameraControl.navMode = 'orbit';
  viewer.cameraControl.followPointer = true;

  // High-tech cyan highlight for the selected object.
  const hi = viewer.scene.highlightMaterial;
  hi.fill = true;
  hi.fillColor = [0.36, 0.83, 1.0];
  hi.fillAlpha = 0.45;
  hi.edges = true;
  hi.edgeColor = [0.42, 0.90, 1.0];

  // X-ray (Fade) look for the non-selected objects.
  const xr = viewer.scene.xrayMaterial;
  xr.fillAlpha = 0.08;
  xr.edgeAlpha = 0.25;

  // --- IFC loader plugin (bridges web-ifc -> xeokit) -------------------------
  // Live path for SMALL models: parses .ifc on the fly via web-ifc 0.0.51.
  const ifcLoader = new WebIFCLoaderPlugin(viewer, { WebIFC, IfcAPI: ifcAPI });

  // --- XKT loader plugin (fast path for BIG models) --------------------------
  // Loads .xkt files pre-built offline by convert.js. No web-ifc involved here:
  // geometry + metadata are already in xeokit's native format, so even ~47k-object
  // models stream in fast without freezing the UI. This is a fully separate path
  // from ifcLoader — the two never interact.
  const xktLoader = new XKTLoaderPlugin(viewer);

  let currentModel = null;

  function destroyModel() {
    if (currentModel) {
      currentModel.destroy();
      currentModel = null;
    }
  }

  /** Wait for a freshly-issued model to finish loading; resolve on 'loaded'. */
  function awaitLoaded(model) {
    return new Promise((resolve, reject) => {
      model.on('loaded', resolve);
      model.on('error', reject);
    });
  }

  /**
   * Loads an IFC file (already read into an ArrayBuffer) and frames it.
   * @returns {Promise<{objectCount:number}>}
   */
  async function loadIFC(arrayBuffer) {
    destroyModel();
    const model = ifcLoader.load({
      id: 'ifcModel',
      ifc: arrayBuffer,      // ArrayBuffer — parsed locally, no network
      edges: true,
      loadMetadata: true,    // required for the Name + GUID panel
    });

    await awaitLoaded(model);

    currentModel = model;
    viewer.cameraFlight.flyTo({ aabb: viewer.scene.aabb }); // fit whole model
    return { objectCount: viewer.scene.objectIds.length };
  }

  /**
   * Loads a pre-converted .xkt (ArrayBuffer) — the fast path for big models.
   * Metadata is embedded in the XKT, so the Name + GUID panel and Mảnh 2/3
   * (colorize / visibility by object id) work exactly as on the IFC path.
   * @returns {Promise<{objectCount:number}>}
   */
  async function loadXKT(arrayBuffer) {
    destroyModel();
    const model = xktLoader.load({
      id: 'xktModel',
      xkt: arrayBuffer,      // ArrayBuffer — already in xeokit's native format
      edges: true,
    });

    await awaitLoaded(model);

    currentModel = model;
    viewer.cameraFlight.flyTo({ aabb: viewer.scene.aabb }); // fit whole model
    return { objectCount: viewer.scene.objectIds.length };
  }

  /**
   * Picks the object under canvas coordinates [x, y].
   * @returns {null | {id:string, name:string, guid:string, type:string}}
   */
  function pick(canvasCoords) {
    const hit = viewer.scene.pick({ canvasPos: canvasCoords });
    if (!hit || !hit.entity) return null;
    const id = hit.entity.id;
    // For WebIFCLoaderPlugin (globalizeObjectIds=false) the entity id IS the
    // IfcGloballyUniqueId, and metaObject.id matches it.
    const meta = viewer.metaScene.metaObjects[id];
    return {
      id,
      name: meta?.name || `Object ${id}`,
      guid: id,
      type: meta?.type || '',
    };
  }

  function setHighlighted(id, on) {
    if (id) viewer.scene.setObjectsHighlighted([id], on);
  }

  /** Fade: X-ray every object except `keepId` (pass on=false to clear). */
  function setXRayOthers(keepId, on) {
    const ids = viewer.scene.objectIds;
    viewer.scene.setObjectsXRayed(ids, on);
    if (on && keepId) viewer.scene.setObjectsXRayed([keepId], false);
  }

  /** Focus: fly the camera to fit a single object. */
  function focusObject(id) {
    if (!id) return;
    const aabb = viewer.scene.getAABB([id]);
    viewer.cameraFlight.flyTo({ aabb });
  }

  function setBackground(rgb) {
    viewer.scene.canvas.backgroundColor = rgb;
  }

  return {
    // raw handles — extension points for later pieces  // [EXT]
    viewer,
    ifcLoader,
    xktLoader,
    ifcAPI,
    // background presets
    BG_DARK,
    BG_LIGHT,
    // operations used by the UI
    loadIFC,
    loadXKT,
    pick,
    setHighlighted,
    setXRayOthers,
    focusObject,
    setBackground,
    getObjectCount: () => viewer.scene.objectIds.length,
  };
}
