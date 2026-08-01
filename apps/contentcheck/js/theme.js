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
    spacegrotesk: "'Space Grotesk', sans-serif",
    inter: "'Inter', sans-serif",
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

  // Matches the Creative Solutions Dashboard's dark theme, so Content Check
  // looks consistent with the rest of the suite by default.
  const DEFAULTS = {
    headerFont: 'spacegrotesk',
    bodyFont: 'inter',
    fontColor: '#E7ECF2',
    bgColor: '#0A0E13',
    bannerColor: '#171F2A',
    accentColor: '#4FD1C5',
    matchColor: '#4FD1C5',
    modifiedColor: '#F2A93B',
    missingColor: '#EF6461',
    addedColor: '#7C93F2',
  };

  /* ---- color math (hex <-> rgb, tint/shade/overlay) ---- */
  function hexToRgb(hex) {
    const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex || '');
    if (!m) return { r: 0, g: 0, b: 0 };
    return { r: parseInt(m[1], 16), g: parseInt(m[2], 16), b: parseInt(m[3], 16) };
  }
  function toHex(n) {
    return Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
  }
  /** Mix a color toward white by `ratio` (0-1). */
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
  /**
   * Translucent overlay of a color at `alpha` opacity — used for status
   * backgrounds (stat cards, badges, row highlights). Unlike tint(), this
   * looks correct regardless of whether the page background is light or
   * dark, since it blends with whatever's underneath rather than assuming
   * a white base.
   */
  function overlay(hex, alpha) {
    const { r, g, b } = hexToRgb(hex);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
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
    root.setProperty('--primary-soft', overlay(t.accentColor, 0.1));

    root.setProperty('--green', t.matchColor);
    root.setProperty('--green-bg', overlay(t.matchColor, 0.16));
    root.setProperty('--yellow', t.modifiedColor);
    root.setProperty('--yellow-bg', overlay(t.modifiedColor, 0.16));
    root.setProperty('--red', t.missingColor);
    root.setProperty('--red-bg', overlay(t.missingColor, 0.16));
    root.setProperty('--blue', t.addedColor);
    root.setProperty('--blue-bg', overlay(t.addedColor, 0.16));
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

  global.CCTheme = { DEFAULTS, applyTheme, save, load, clear, loadGoogleFont, tint, shade, overlay };
})(window);
