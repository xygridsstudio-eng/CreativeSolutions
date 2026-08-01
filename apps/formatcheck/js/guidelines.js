/**
 * guidelines.js
 * Single responsibility: the brand-guideline data model — defaults,
 * localStorage persistence, and JSON import/export. No DOM/UI logic here.
 */
(function (global) {
  'use strict';

  const STORAGE_KEY = 'fc_guidelines_v1';

  // Sensible starting point matching the Creative Solutions suite's own
  // default theme — teams should adjust these to their actual brand kit.
  const DEFAULTS = {
    name: 'Creative Solutions — Default Brand Kit',
    fonts: {
      allowed: ['Space Grotesk', 'Inter'],
    },
    fontSizes: {
      allowed: [11, 14, 18, 24, 28, 32, 40],
      minBodyPt: 11,
    },
    colors: {
      allowed: ['#4FD1C5', '#F2A93B', '#EF6461', '#0A0E13', '#121821', '#E7ECF2', '#8592A6'],
    },
    lineSpacing: {
      targetPct: 100,
      tolerancePct: 10,
    },
    paragraphSpacing: {
      // Points, applies to spcBef/spcAft when expressed in points
      spaceBeforePt: 0,
      spaceAfterPt: 6,
      tolerancePt: 2,
    },
    shapeMargins: {
      leftIn: 0.1,
      topIn: 0.05,
      rightIn: 0.1,
      bottomIn: 0.05,
      toleranceIn: 0.02,
    },
    checks: {
      font: true,
      fontSize: true,
      textColor: true,
      shapeFillColor: true,
      lineSpacing: true,
      paragraphSpacing: true,
      shapeMargins: true,
    },
  };

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return JSON.parse(JSON.stringify(DEFAULTS));
      const parsed = JSON.parse(raw);
      return mergeDefaults(parsed);
    } catch (e) {
      return JSON.parse(JSON.stringify(DEFAULTS));
    }
  }

  function mergeDefaults(partial) {
    const base = JSON.parse(JSON.stringify(DEFAULTS));
    return deepMerge(base, partial || {});
  }

  function deepMerge(target, source) {
    Object.keys(source).forEach((key) => {
      if (
        source[key] && typeof source[key] === 'object' && !Array.isArray(source[key]) &&
        target[key] && typeof target[key] === 'object' && !Array.isArray(target[key])
      ) {
        deepMerge(target[key], source[key]);
      } else {
        target[key] = source[key];
      }
    });
    return target;
  }

  function save(guidelines) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(guidelines));
    } catch (e) { /* not essential to function */ }
  }

  function clear() {
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) { /* ignore */ }
  }

  function toJsonBlob(guidelines) {
    return new Blob([JSON.stringify(guidelines, null, 2)], { type: 'application/json' });
  }

  function fromJsonText(text) {
    const parsed = JSON.parse(text);
    return mergeDefaults(parsed);
  }

  global.FCGuidelines = { DEFAULTS, load, save, clear, mergeDefaults, toJsonBlob, fromJsonText };
})(window);
