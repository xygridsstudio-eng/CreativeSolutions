/**
 * ui.js
 * Single responsibility: all DOM reads/writes. No parsing/comparison logic
 * lives here — app.js calls into these render functions and reads state
 * back out via the getters below.
 */
(function (global) {
  'use strict';

  const el = (id) => document.getElementById(id);

  const state = {
    sourceFile: null,
    outputFile: null,
  };

  function formatBytes(bytes) {
    return global.CCUtils.formatBytes(bytes);
  }

  function renderFileSlot(slotEl, file) {
    const nameEl = slotEl.querySelector('.file-name');
    const metaEl = slotEl.querySelector('.file-meta');
    if (file) {
      slotEl.classList.add('has-file');
      nameEl.textContent = file.name;
      metaEl.textContent = `${formatBytes(file.size)} • ${(file.name.split('.').pop() || '').toUpperCase()}`;
    } else {
      slotEl.classList.remove('has-file');
      nameEl.textContent = 'Drag & drop a file, or click to browse';
      metaEl.textContent = 'DOCX, PPTX, PDF, or TXT';
    }
  }

  function wireUploadSlot(slotId, inputId, onFileSelected) {
    const slot = el(slotId);
    const input = el(inputId);

    slot.addEventListener('click', () => input.click());
    input.addEventListener('change', () => {
      if (input.files && input.files[0]) onFileSelected(input.files[0]);
    });

    ['dragenter', 'dragover'].forEach((evt) => {
      slot.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        slot.classList.add('dragover');
      });
    });
    ['dragleave', 'drop'].forEach((evt) => {
      slot.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        slot.classList.remove('dragover');
      });
    });
    slot.addEventListener('drop', (e) => {
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) onFileSelected(file);
    });
  }

  function wireCollapsible(headerId, contentId, startOpen) {
    const header = el(headerId);
    const content = el(contentId);
    if (!header || !content) return;
    const setOpen = (open) => {
      header.classList.toggle('open', open);
      content.classList.toggle('hidden', !open);
    };
    setOpen(!!startOpen);
    header.addEventListener('click', () => setOpen(content.classList.contains('hidden')));
  }

  const THEME_FIELD_IDS = {
    headerFont: 'themeHeaderFont',
    bodyFont: 'themeBodyFont',
    fontColor: 'themeFontColor',
    bgColor: 'themeBgColor',
    bannerColor: 'themeBannerColor',
    accentColor: 'themeAccentColor',
    matchColor: 'themeMatchColor',
    modifiedColor: 'themeModifiedColor',
    missingColor: 'themeMissingColor',
    addedColor: 'themeAddedColor',
  };

  function readThemeFromControls() {
    const theme = {};
    Object.keys(THEME_FIELD_IDS).forEach((key) => {
      theme[key] = el(THEME_FIELD_IDS[key]).value;
    });
    return theme;
  }

  function writeThemeToControls(theme) {
    Object.keys(THEME_FIELD_IDS).forEach((key) => {
      el(THEME_FIELD_IDS[key]).value = theme[key];
    });
  }

  function loadFontsForTheme(theme) {
    ['themeHeaderFont', 'themeBodyFont'].forEach((id) => {
      const select = el(id);
      const opt = select.options[select.selectedIndex];
      if (opt && opt.dataset.google) global.CCTheme.loadGoogleFont(opt.dataset.google);
    });
  }

  function initThemePanel() {
    wireCollapsible('themeToggle', 'themeContent', false);

    const saved = global.CCTheme.load();
    writeThemeToControls(saved);
    global.CCTheme.applyTheme(saved);
    loadFontsForTheme(saved);

    const onChange = () => {
      const theme = readThemeFromControls();
      loadFontsForTheme(theme);
      global.CCTheme.applyTheme(theme);
      global.CCTheme.save(theme);
    };
    Object.values(THEME_FIELD_IDS).forEach((id) => {
      el(id).addEventListener('input', onChange);
      el(id).addEventListener('change', onChange);
    });

    el('themeResetBtn').addEventListener('click', () => {
      writeThemeToControls(global.CCTheme.DEFAULTS);
      global.CCTheme.applyTheme(global.CCTheme.DEFAULTS);
      global.CCTheme.clear();
    });
  }

  const SUPPORTED_EXTENSIONS = ['docx', 'pptx', 'pdf', 'txt'];
  const URL_STORAGE_KEY = 'contentcheck_last_urls_v1';

  function loadLastUrls() {
    try {
      const raw = localStorage.getItem(URL_STORAGE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (e) {
      return {};
    }
  }
  function saveLastUrl(prefix, url) {
    try {
      const all = loadLastUrls();
      all[prefix] = url;
      localStorage.setItem(URL_STORAGE_KEY, JSON.stringify(all));
    } catch (e) { /* ignore — remembering the path is a convenience, not required */ }
  }

  function filenameFromContentDisposition(header) {
    if (!header) return null;
    const starMatch = /filename\*=(?:UTF-8'')?([^;]+)/i.exec(header);
    if (starMatch) {
      try { return decodeURIComponent(starMatch[1].replace(/["']/g, '').trim()); } catch (e) { /* fall through */ }
    }
    const plainMatch = /filename="?([^";]+)"?/i.exec(header);
    return plainMatch ? plainMatch[1].trim() : null;
  }

  function filenameFromUrl(url) {
    try {
      const u = new URL(url, window.location.href);
      const last = decodeURIComponent(u.pathname.split('/').filter(Boolean).pop() || '');
      return last || null;
    } catch (e) {
      return null;
    }
  }

  function setUrlStatus(statusId, text, kind) {
    const el2 = el(statusId);
    el2.textContent = text;
    el2.className = `url-status${kind ? ` ${kind}` : ''}`;
  }

  function wireUrlFetch(prefix, onFileSelected) {
    const input = el(`${prefix}UrlInput`);
    const btn = el(`${prefix}UrlFetchBtn`);
    const statusId = `${prefix}UrlStatus`;

    const saved = loadLastUrls();
    if (saved[prefix]) input.value = saved[prefix];

    const doFetch = async () => {
      const url = input.value.trim();
      if (!url) return;
      saveLastUrl(prefix, url);

      btn.disabled = true;
      setUrlStatus(statusId, 'Fetching…', 'pending');
      try {
        const response = await fetch(url, { mode: 'cors' });
        if (!response.ok) throw new Error(`Server responded with ${response.status} ${response.statusText}`);

        const blob = await response.blob();
        const cd = response.headers.get('content-disposition');
        const filename = filenameFromContentDisposition(cd) || filenameFromUrl(url) || 'downloaded-file';
        const ext = global.CCUtils.fileExtension(filename);

        if (!SUPPORTED_EXTENSIONS.includes(ext)) {
          const e = new Error(
            `That link resolved to "${filename}", which isn't a supported file type. ` +
            `Supported: DOCX, PPTX, PDF, TXT.`
          );
          e.isValidationError = true;
          throw e;
        }

        const file = new File([blob], filename, { type: blob.type });
        onFileSelected(file);
        setUrlStatus(statusId, `Loaded "${filename}" (${formatBytes(file.size)})`, 'ok');
      } catch (err) {
        if (err && err.isValidationError) {
          setUrlStatus(statusId, err.message, 'err');
        } else {
          setUrlStatus(
            statusId,
            `Couldn't fetch that link (${err && err.message ? err.message : err}). This is often a CORS or ` +
            `authentication restriction on corporate SharePoint/network links — download the file locally and ` +
            `use the box above instead.`,
            'err'
          );
        }
      } finally {
        btn.disabled = false;
      }
    };

    btn.addEventListener('click', doFetch);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); doFetch(); }
    });
  }

  function init({ onCompare, onReset, onDownload }) {
    const handleSourceFile = (file) => {
      state.sourceFile = file;
      renderFileSlot(el('sourceSlot'), file);
      updateCompareButton();
    };
    const handleOutputFile = (file) => {
      state.outputFile = file;
      renderFileSlot(el('outputSlot'), file);
      updateCompareButton();
    };

    wireUploadSlot('sourceSlot', 'sourceInput', handleSourceFile);
    wireUploadSlot('outputSlot', 'outputInput', handleOutputFile);
    wireUrlFetch('source', handleSourceFile);
    wireUrlFetch('output', handleOutputFile);
    initThemePanel();

    el('compareBtn').addEventListener('click', () => onCompare(getOptions()));
    el('resetBtn').addEventListener('click', () => {
      reset();
      onReset();
    });
    el('downloadBtn').addEventListener('click', onDownload);

    renderFileSlot(el('sourceSlot'), null);
    renderFileSlot(el('outputSlot'), null);
    updateCompareButton();
  }

  function updateCompareButton() {
    el('compareBtn').disabled = !(state.sourceFile && state.outputFile);
  }

  function getOptions() {
    return {
      ignoreCase: el('optIgnoreCase').checked,
      ignoreExtraSpaces: el('optIgnoreSpaces').checked,
      ignoreBlankLines: el('optIgnoreBlankLines').checked,
      ignorePunctuation: el('optIgnorePunctuation').checked,
      ignoreBulletSymbols: el('optIgnoreBullets').checked,
      ignoreHeaders: el('optIgnoreHeaders').checked,
      ignoreFooters: el('optIgnoreFooters').checked,
      compareNumbers: el('optCompareNumbers').checked,
      compareDates: el('optCompareDates').checked,
      compareUrls: el('optCompareUrls').checked,
      compareEmails: el('optCompareEmails').checked,
    };
  }

  function getFiles() {
    return { sourceFile: state.sourceFile, outputFile: state.outputFile };
  }

  function reset() {
    state.sourceFile = null;
    state.outputFile = null;
    el('sourceInput').value = '';
    el('outputInput').value = '';
    renderFileSlot(el('sourceSlot'), null);
    renderFileSlot(el('outputSlot'), null);
    setUrlStatus('sourceUrlStatus', '', '');
    setUrlStatus('outputUrlStatus', '', '');
    updateCompareButton();
    hideProgress();
    hideSummary();
    hidePreview();
    hideError();
    el('downloadBtn').disabled = true;
  }

  function showProgress(label) {
    el('progressWrap').classList.remove('hidden');
    setProgress(0, label || 'Starting…');
  }
  function setProgress(pct, label) {
    el('progressBarFill').style.width = `${Math.min(100, Math.max(0, pct))}%`;
    if (label) el('progressLabel').textContent = label;
  }
  function hideProgress() {
    el('progressWrap').classList.add('hidden');
  }

  function showError(message) {
    const box = el('errorBox');
    box.textContent = message;
    box.classList.remove('hidden');
  }
  function hideError() {
    el('errorBox').classList.add('hidden');
  }

  function hideSummary() {
    el('summaryWrap').classList.add('hidden');
  }

  function renderSummary(report) {
    const s = report.stats;
    const cards = [
      ['Total Blocks', s.totalParagraphs, 'neutral'],
      ['Total Sentences', s.totalSentences, 'neutral'],
      ['Matched', s.matched, 'match'],
      ['Modified', s.modified, 'modified'],
      ['Missing', s.missing, 'missing'],
      ['Added', s.added, 'added'],
      ['Match %', `${s.matchPercentage}%`, 'neutral'],
      ['Processing Time', report.meta.processingTime, 'neutral'],
    ];
    el('summaryGrid').innerHTML = cards.map(([label, value, cls]) => `
      <div class="stat-card stat-${cls}">
        <div class="stat-value">${global.CCUtils.escapeHtml(String(value))}</div>
        <div class="stat-label">${global.CCUtils.escapeHtml(label)}</div>
      </div>
    `).join('');
    el('summaryWrap').classList.remove('hidden');
    el('downloadBtn').disabled = false;
  }

  const STATUS_LABELS = { match: 'Match', modified: 'Modified', missing: 'Missing', added: 'Added' };

  function hidePreview() {
    el('previewWrap').classList.add('hidden');
  }

  function renderPreview(report) {
    const interesting = report.detailRows.filter((r) => r.status !== 'match').slice(0, 200);
    const rows = interesting.length ? interesting : report.detailRows.slice(0, 50);

    el('previewCount').textContent = interesting.length
      ? `Showing ${rows.length} of ${interesting.length} changed sentences (full detail in the Excel report)`
      : `No differences found — showing first ${rows.length} matched sentences`;

    el('previewBody').innerHTML = rows.map((r) => `
      <tr class="row-${r.status}">
        <td>${r.blockId}</td>
        <td>${global.CCUtils.escapeHtml(r.section || '')}</td>
        <td><span class="badge badge-${r.status}">${STATUS_LABELS[r.status] || r.status}</span></td>
        <td>${global.CCUtils.escapeHtml((r.sourceText || '').slice(0, 160))}</td>
        <td>${global.CCUtils.escapeHtml((r.outputText || '').slice(0, 160))}</td>
        <td>${global.CCUtils.escapeHtml(r.comments || '')}</td>
      </tr>
    `).join('');
    el('previewWrap').classList.remove('hidden');
  }

  global.CCUi = {
    init,
    getOptions,
    getFiles,
    reset,
    showProgress,
    setProgress,
    hideProgress,
    showError,
    hideError,
    renderSummary,
    hideSummary,
    renderPreview,
    hidePreview,
  };
})(window);
