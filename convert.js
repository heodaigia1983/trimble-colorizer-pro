#!/usr/bin/env node
// =============================================================================
// convert.js  —  OFFLINE IFC -> XKT converter (standalone Node process)
// -----------------------------------------------------------------------------
// The heavy IFC parsing happens HERE, in its own Node process — NEVER in the
// Electron renderer — so the viewer UI can't freeze. It produces a pre-optimized
// .xkt that XKTLoaderPlugin streams into the scene almost instantly, which is the
// fast path for big models (~47k objects). The small-model live .ifc path
// (WebIFCLoaderPlugin in src/viewer.js) is left untouched.
//
// USAGE
//   node convert.js <input.ifc> [output.xkt]
//   npm run convert -- <input.ifc> [output.xkt]
//
// If no output is given, writes models/<same-name>.xkt (dir auto-created).
//
// WEB-IFC VERSION NOTE (important — see project memory):
//   The app pins web-ifc 0.0.51 because xeokit 2.6.111's WebIFCLoaderPlugin only
//   tolerates that build's geometry. convert2xkt ships its OWN web-ifc (0.0.70)
//   and we deliberately DO NOT pass a `WebIFC` here, so convert2xkt self-resolves
//   its bundled copy. The two web-ifc versions never mix: 0.0.51 stays on the
//   live viewer path, 0.0.70 stays inside this offline converter.
// =============================================================================

'use strict';

const path = require('node:path');
const fs = require('node:fs');

function mb(bytes) {
  return (bytes / (1024 * 1024)).toFixed(2);
}

async function main() {
  const args = process.argv.slice(2);
  if (!args[0] || args[0] === '-h' || args[0] === '--help') {
    console.log('Usage: node convert.js <input.ifc> [output.xkt]');
    console.log('       npm run convert -- <input.ifc> [output.xkt]');
    process.exit(args[0] ? 0 : 1);
  }

  const inputPath = path.resolve(args[0]);
  if (!fs.existsSync(inputPath)) {
    console.error(`[convert] ERROR: input file not found: ${inputPath}`);
    process.exit(1);
  }

  // Output: explicit 2nd arg, else models/<basename>.xkt next to this script.
  const baseName = path.parse(inputPath).name;
  const outputPath = args[1]
    ? path.resolve(args[1])
    : path.resolve(__dirname, 'models', `${baseName}.xkt`);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });

  // --- polyfills convert2xkt's loaders.gl + web-ifc stack expects in Node -----
  // (mirrors the setup in @xeokit/xeokit-convert's own CLI bin)
  const { TextEncoder, TextDecoder } = require('node:util');
  if (!global.TextEncoder) global.TextEncoder = TextEncoder;
  if (!global.TextDecoder) global.TextDecoder = TextDecoder;
  await import('@loaders.gl/polyfills');
  const { installFilePolyfills } = await import('@loaders.gl/polyfills');
  installFilePolyfills();

  // convert2xkt is ESM-only. Import the FUNCTION module (src/convert2xkt.js),
  // not the package's convert2xkt.js bin — the bin parses argv on import.
  const { convert2xkt } = await import('@xeokit/xeokit-convert/src/convert2xkt.js');

  const t0 = Date.now();
  const elapsed = () => ((Date.now() - t0) / 1000).toFixed(1).padStart(6);
  const log = (msg) => console.log(`[convert +${elapsed()}s] ${msg}`);

  const inMB = mb(fs.statSync(inputPath).size);
  console.log('============================================================');
  console.log('[convert] IFC -> XKT (offline, standalone Node process)');
  console.log(`[convert] input : ${inputPath}`);
  console.log(`[convert] output: ${outputPath}`);
  console.log(`[convert] source size: ${inMB} MB`);
  console.log('============================================================');

  const stats = {};
  try {
    await convert2xkt({
      source: inputPath,     // convert2xkt infers format from the .ifc extension
      output: outputPath,    // writes the .xkt to disk for us
      stats,                 // filled in below with object/triangle counts
      log,                   // streamed progress with elapsed timestamps
      // WebIFC intentionally omitted — self-resolves convert2xkt's bundled 0.0.70.
    });
  } catch (err) {
    console.error('============================================================');
    console.error(`[convert] FAILED after ${((Date.now() - t0) / 1000).toFixed(1)}s`);
    console.error(`[convert] ${err && err.stack ? err.stack : err}`);
    console.error('============================================================');
    process.exit(1);
  }

  const secs = ((Date.now() - t0) / 1000).toFixed(1);
  const outMB = mb(fs.statSync(outputPath).size);
  console.log('============================================================');
  console.log(`[convert] DONE in ${secs}s`);
  console.log(`[convert]   objects     : ${stats.numObjects ?? '?'}`);
  console.log(`[convert]   meta objects: ${stats.numMetaObjects ?? '?'}`);
  console.log(`[convert]   geometries  : ${stats.numGeometries ?? '?'}`);
  console.log(`[convert]   triangles   : ${stats.numTriangles ?? '?'}`);
  console.log(`[convert]   xkt size    : ${outMB} MB  (source ${inMB} MB, ratio ${stats.compressionRatio ?? '?'})`);
  console.log(`[convert]   written     : ${outputPath}`);
  console.log('============================================================');
  process.exit(0);
}

main().catch((err) => {
  console.error('[convert] Fatal:', err && err.stack ? err.stack : err);
  process.exit(1);
});
