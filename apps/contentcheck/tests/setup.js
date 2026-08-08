/**
 * setup.js
 * Loads the app's plain browser scripts (utils.js, normalizer.js,
 * comparer.js) into a fake `window` so they can be exercised with Node's
 * built-in test runner. No bundler, no DOM, no npm dependency — mirrors
 * how the app itself has no build step, just <script> tags in order.
 *
 * parser.js is intentionally NOT loaded here: it needs a real DOMParser/
 * JSZip/pdf.js, which only exist in a browser. Its logic is exercised via
 * manual in-browser checks instead (see apps/contentcheck/README.md).
 */
'use strict';
const path = require('path');

function loadApp() {
  const win = {};
  global.window = win;

  const files = ['../js/utils.js', '../js/normalizer.js', '../js/comparer.js'];
  files.forEach((rel) => {
    const abs = require.resolve(rel);
    delete require.cache[abs];
    require(abs);
  });

  return win;
}

module.exports = { loadApp };
