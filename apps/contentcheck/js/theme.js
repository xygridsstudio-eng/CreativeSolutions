/**
 * theme.js
 * Single responsibility: theme state — defaults, applying a theme to CSS
 * custom properties, deriving light/dark tint shades for a base color,
 * persisting to localStorage, and loading Google Fonts on demand.
 * No DOM wiring (that's ui.js's job) beyond writing CSS variables and
 * injecting <link> tags for fonts.
 */
(function (global) {
  'use strict';

  const STORAGE_KEY = 'contentcheck_theme_v1';

  const FONT_STACKS = {
    system: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
    georgia: "Georgia, 'Times New Roman', serif",
    playfair: "'Playfair Display', Georgia, serif",
    times: "'Times New Roman', Times, serif",
    trebuchet: "'Trebuchet MS', sans-serif",
    verdana: "Verdana, Geneva, sans-serif",
    courier: "'Courier New', Courier, monospace",
    arial: "Arial, Helvetica, sans-serif",
    roboto: "'Roboto', sans-serif",
    opensans: "'Open Sans', sans-serif",
  };

  const DEFAULTS = {
    headerFont: 'system',
    bodyFont: 'system',
    fontColor: '#1f2937',
    bgColor: '#f4f6f9',
    bannerColor: '#1e293b',
    accentColor: '#2563eb',
    matchColor: '#16a34a',
    modifiedColor: '#b45309',
    missingColor: '#dc2626',
    addedColor: '#1d4ed8',
  };

  /* ---- color math (hex <-> rgb, tint/shade) ---- */
  function hexToRgb(hex) {
    const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex || '');
    if (!m) return { r: 0, g: 0, b: 0 };
    return { r: parseInt(m[1], 16), g: parseInt(m[2], 16), b: parseInt(m[3], 16) };
  }
  function toHex(n) {
    return Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
  }
  /** Mix a color toward white by `ratio` (0-1) — used for light status backgrounds. */
  function tint(hex, ratio) {
    const { r, g, b } = hexToRgb(hex);
    const mix = (c) => c + (255 - c) * ratio;
    return `#${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
  }
  /** Mix a color toward black by `ratio` (0-1) — used for hover/dark shades. */
  function shade(hex, ratio) {
    const { r, g, b } = hexToRgb(hex);
    const mix = (c) => c * (1 - ratio);
    return `#${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
  }

  function fontStack(key) {
    return FONT_STACKS[key] || FONT_STACKS.system;
  }

  function loadGoogleFont(spec) {
    if (!spec || typeof document === 'undefined') return;
    const id = `gf-${spec.replace(/[^a-zA-Z0-9]/g, '')}`;
    if (document.getElementById(id)) return;
    const link = document.createElement('link');
    link.id = id;
    link.rel = 'stylesheet';
    link.href = `https://fonts.googleapis.com/css2?family=${spec}&display=swap`;
    document.head.appendChild(link);
  }

  function applyTheme(theme) {
    if (typeof document === 'undefined') return;
    const t = Object.assign({}, DEFAULTS, theme || {});
    const root = document.documentElement.style;

    root.setProperty('--theme-header-font', fontStack(t.headerFont));
    root.setProperty('--theme-body-font', fontStack(t.bodyFont));

    root.setProperty('--text', t.fontColor);
    root.setProperty('--bg', t.bgColor);
    root.setProperty('--banner-bg', t.bannerColor);
    root.setProperty('--banner-bg-2', shade(t.bannerColor, 0.35));

    root.setProperty('--primary', t.accentColor);
    root.setProperty('--primary-dark', shade(t.accentColor, 0.15));

    root.setProperty('--green', t.matchColor);
    root.setProperty('--green-bg', tint(t.matchColor, 0.88));
    root.setProperty('--yellow', t.modifiedColor);
    root.setProperty('--yellow-bg', tint(t.modifiedColor, 0.88));
    root.setProperty('--red', t.missingColor);
    root.setProperty('--red-bg', tint(t.missingColor, 0.9));
    root.setProperty('--blue', t.addedColor);
    root.setProperty('--blue-bg', tint(t.addedColor, 0.88));
  }

  function save(theme) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(theme));
    } catch (e) {
      // localStorage may be unavailable (private browsing, file:// in some
      // browsers) — theming still works for the current page load.
    }
  }

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? Object.assign({}, DEFAULTS, JSON.parse(raw)) : Object.assign({}, DEFAULTS);
    } catch (e) {
      return Object.assign({}, DEFAULTS);
    }
  }

  function clear() {
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) { /* ignore */ }
  }

  global.CCTheme = { DEFAULTS, applyTheme, save, load, clear, loadGoogleFont, tint, shade };
})(window);
