/**
 * shell.js
 * Single responsibility: sidebar navigation for the Creative Solutions
 * suite — switches which tool is loaded into the content iframe, tracks
 * the active nav item, and remembers the last-selected tool + sidebar
 * collapse state across visits (localStorage).
 */
(function () {
  'use strict';

  const TOOLS = {
    dashboard: { src: 'apps/dashboard/index.html', title: 'Team Dashboard' },
    newsletter: { src: 'apps/newsletter/index.html', title: 'Newsletter Builder' },
    contentcheck: { src: 'apps/contentcheck/index.html', title: 'Content Check' },
    about: { src: 'apps/about/index.html', title: 'About & How to Use' },
  };
  const DEFAULT_TOOL = 'dashboard';
  const DEFAULT_LOGO_SRC = 'assets/logo.svg';
  const STORAGE_KEY = 'cs_suite_state_v1';
  const LOGO_STORAGE_KEY = 'cs_suite_logo_v1';

  const el = (id) => document.getElementById(id);
  const frame = el('toolFrame');
  const loading = el('contentLoading');
  const sidebar = el('sidebar');
  const navItems = Array.from(document.querySelectorAll('.nav-item'));

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (e) {
      return {};
    }
  }
  function saveState(patch) {
    try {
      const current = loadState();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(Object.assign(current, patch)));
    } catch (e) { /* ignore — not essential to function */ }
  }

  function initLogoUpload() {
    const slot = el('logoSlot');
    const input = el('logoInput');
    const img = el('brandMark');
    const removeBtn = el('logoRemoveBtn');

    const applyLogo = (dataUrl) => {
      img.src = dataUrl || DEFAULT_LOGO_SRC;
      removeBtn.classList.toggle('hidden', !dataUrl);
    };

    // Restore a previously uploaded logo, if any.
    try {
      const saved = localStorage.getItem(LOGO_STORAGE_KEY);
      if (saved) applyLogo(saved);
    } catch (e) { /* ignore */ }

    slot.addEventListener('click', () => input.click());

    input.addEventListener('change', () => {
      const file = input.files && input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result;
        try { localStorage.setItem(LOGO_STORAGE_KEY, dataUrl); } catch (e) { /* ignore */ }
        applyLogo(dataUrl);
      };
      reader.readAsDataURL(file);
    });

    removeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      try { localStorage.removeItem(LOGO_STORAGE_KEY); } catch (e2) { /* ignore */ }
      applyLogo(null);
      input.value = '';
    });
  }

  function setActiveTool(key) {
    const tool = TOOLS[key] || TOOLS[DEFAULT_TOOL];
    const resolvedKey = TOOLS[key] ? key : DEFAULT_TOOL;

    navItems.forEach((btn) => btn.classList.toggle('active', btn.dataset.tool === resolvedKey));
    loading.classList.remove('hidden');
    frame.src = tool.src;
    document.title = `${tool.title} — Creative Solutions`;
    saveState({ tool: resolvedKey });
  }

  frame.addEventListener('load', () => loading.classList.add('hidden'));

  navItems.forEach((btn) => {
    btn.addEventListener('click', () => setActiveTool(btn.dataset.tool));
  });

  el('collapseBtn').addEventListener('click', () => {
    const collapsed = sidebar.classList.toggle('collapsed');
    saveState({ collapsed });
  });

  // Restore last session's tool + sidebar state.
  const saved = loadState();
  if (saved.collapsed) sidebar.classList.add('collapsed');
  setActiveTool(saved.tool || DEFAULT_TOOL);
  initLogoUpload();
})();
