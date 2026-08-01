/**
 * ui.js
 * Single responsibility: all DOM reads/writes. No parsing/checking logic
 * lives here — app.js calls into these functions and reads state back via
 * the getters below.
 */
(function (global) {
  'use strict';

  const el = (id) => document.getElementById(id);
  const state = { deckFile: null, guidelines: null };

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  /* ---------------------------------------------------------------- *
   * Upload slot
   * ---------------------------------------------------------------- */
  function wireUpload(onFileSelected) {
    const slot = el('deckSlot');
    const input = el('deckInput');
    slot.addEventListener('click', () => input.click());
    input.addEventListener('change', () => { if (input.files[0]) onFileSelected(input.files[0]); });
    ['dragenter', 'dragover'].forEach((evt) => slot.addEventListener(evt, (e) => { e.preventDefault(); slot.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach((evt) => slot.addEventListener(evt, (e) => { e.preventDefault(); slot.classList.remove('dragover'); }));
    slot.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) onFileSelected(f); });
  }

  function setDeckFile(file) {
    state.deckFile = file;
    const slot = el('deckSlot');
    if (file) {
      slot.classList.add('has-file');
      el('deckFileName').textContent = file.name;
      el('deckFileMeta').textContent = formatBytes(file.size);
    } else {
      slot.classList.remove('has-file');
      el('deckFileName').textContent = 'Drag & drop a .pptx file, or click to browse';
      el('deckFileMeta').textContent = 'PowerPoint (.pptx) only';
    }
    el('runBtn').disabled = !file;
  }

  /* ---------------------------------------------------------------- *
   * Collapsible panel
   * ---------------------------------------------------------------- */
  function wireCollapsible(headerId, contentId, startOpen) {
    const header = el(headerId);
    const content = el(contentId);
    const setOpen = (open) => { header.classList.toggle('open', open); content.classList.toggle('hidden', !open); };
    setOpen(!!startOpen);
    header.addEventListener('click', () => setOpen(content.classList.contains('hidden')));
  }

  /* ---------------------------------------------------------------- *
   * Guideline form <-> model binding
   * ---------------------------------------------------------------- */
  function renderChipList(containerId, items, onRemove) {
    const wrap = el(containerId);
    wrap.innerHTML = items.map((item, i) => `
      <span class="chip">${global.FCUtils.escapeHtml(String(item))}<button type="button" data-i="${i}" aria-label="Remove">×</button></span>
    `).join('');
    wrap.querySelectorAll('button').forEach((btn) => {
      btn.addEventListener('click', () => onRemove(Number(btn.dataset.i)));
    });
  }

  function renderColorList(containerId, items, onRemove) {
    const wrap = el(containerId);
    wrap.innerHTML = items.map((hex, i) => `
      <span class="swatch-chip">
        <span class="swatch-dot" style="background:${global.FCUtils.escapeHtml(hex)}"></span>
        <span>${global.FCUtils.escapeHtml(hex)}</span>
        <button type="button" data-i="${i}" aria-label="Remove">×</button>
      </span>
    `).join('');
    wrap.querySelectorAll('button').forEach((btn) => {
      btn.addEventListener('click', () => onRemove(Number(btn.dataset.i)));
    });
  }

  function writeGuidelinesToForm(g) {
    el('gName').value = g.name || '';
    renderChipList('gFontList', g.fonts.allowed, (i) => { g.fonts.allowed.splice(i, 1); writeGuidelinesToForm(g); });
    renderChipList('gSizeList', g.fontSizes.allowed, (i) => { g.fontSizes.allowed.splice(i, 1); writeGuidelinesToForm(g); });
    el('gMinBody').value = g.fontSizes.minBodyPt;
    renderColorList('gColorList', g.colors.allowed, (i) => { g.colors.allowed.splice(i, 1); writeGuidelinesToForm(g); });
    el('gLineTarget').value = g.lineSpacing.targetPct;
    el('gLineTol').value = g.lineSpacing.tolerancePct;
    el('gSpaceBefore').value = g.paragraphSpacing.spaceBeforePt;
    el('gSpaceAfter').value = g.paragraphSpacing.spaceAfterPt;
    el('gSpaceTol').value = g.paragraphSpacing.tolerancePt;
    el('gMarginL').value = g.shapeMargins.leftIn;
    el('gMarginT').value = g.shapeMargins.topIn;
    el('gMarginR').value = g.shapeMargins.rightIn;
    el('gMarginB').value = g.shapeMargins.bottomIn;
    el('gMarginTol').value = g.shapeMargins.toleranceIn;
    el('cFont').checked = g.checks.font;
    el('cFontSize').checked = g.checks.fontSize;
    el('cTextColor').checked = g.checks.textColor;
    el('cShapeFill').checked = g.checks.shapeFillColor;
    el('cLineSpacing').checked = g.checks.lineSpacing;
    el('cParaSpacing').checked = g.checks.paragraphSpacing;
    el('cMargins').checked = g.checks.shapeMargins;
  }

  function readGuidelinesFromForm() {
    const g = state.guidelines;
    g.name = el('gName').value.trim() || 'Untitled Brand Kit';
    g.fontSizes.minBodyPt = Number(el('gMinBody').value) || 0;
    g.lineSpacing.targetPct = Number(el('gLineTarget').value) || 0;
    g.lineSpacing.tolerancePct = Number(el('gLineTol').value) || 0;
    g.paragraphSpacing.spaceBeforePt = Number(el('gSpaceBefore').value) || 0;
    g.paragraphSpacing.spaceAfterPt = Number(el('gSpaceAfter').value) || 0;
    g.paragraphSpacing.tolerancePt = Number(el('gSpaceTol').value) || 0;
    g.shapeMargins.leftIn = Number(el('gMarginL').value) || 0;
    g.shapeMargins.topIn = Number(el('gMarginT').value) || 0;
    g.shapeMargins.rightIn = Number(el('gMarginR').value) || 0;
    g.shapeMargins.bottomIn = Number(el('gMarginB').value) || 0;
    g.shapeMargins.toleranceIn = Number(el('gMarginTol').value) || 0;
    g.checks.font = el('cFont').checked;
    g.checks.fontSize = el('cFontSize').checked;
    g.checks.textColor = el('cTextColor').checked;
    g.checks.shapeFillColor = el('cShapeFill').checked;
    g.checks.lineSpacing = el('cLineSpacing').checked;
    g.checks.paragraphSpacing = el('cParaSpacing').checked;
    g.checks.shapeMargins = el('cMargins').checked;
    return g;
  }

  function wireGuidelineForm({ onSave, onExport, onImportFile, onReset }) {
    state.guidelines = global.FCGuidelines.load();
    writeGuidelinesToForm(state.guidelines);
    wireCollapsible('guidelinesToggle', 'guidelinesContent', false);

    const addFont = () => {
      const val = el('gFontInput').value.trim();
      if (!val) return;
      if (!state.guidelines.fonts.allowed.some((f) => f.toLowerCase() === val.toLowerCase())) {
        state.guidelines.fonts.allowed.push(val);
        writeGuidelinesToForm(state.guidelines);
      }
      el('gFontInput').value = '';
      el('gFontInput').focus();
    };
    el('gFontAddBtn').addEventListener('click', addFont);
    el('gFontInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addFont(); } });

    const addSize = () => {
      const val = Number(el('gSizeInput').value);
      if (!val) return;
      if (!state.guidelines.fontSizes.allowed.includes(val)) {
        state.guidelines.fontSizes.allowed.push(val);
        state.guidelines.fontSizes.allowed.sort((a, b) => a - b);
        writeGuidelinesToForm(state.guidelines);
      }
      el('gSizeInput').value = '';
      el('gSizeInput').focus();
    };
    el('gSizeAddBtn').addEventListener('click', addSize);
    el('gSizeInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addSize(); } });

    el('gColorInput').addEventListener('input', () => { el('gColorHexInput').value = el('gColorInput').value.toUpperCase(); });
    const addColor = () => {
      let val = (el('gColorHexInput').value || el('gColorInput').value).trim();
      if (!val) return;
      if (!val.startsWith('#')) val = '#' + val;
      if (!/^#[0-9A-Fa-f]{6}$/.test(val)) { alert('Enter a valid 6-digit hex color, e.g. #4FD1C5'); return; }
      val = val.toUpperCase();
      if (!state.guidelines.colors.allowed.includes(val)) {
        state.guidelines.colors.allowed.push(val);
        writeGuidelinesToForm(state.guidelines);
      }
      el('gColorHexInput').value = '';
    };
    el('gColorAddBtn').addEventListener('click', addColor);
    el('gColorHexInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addColor(); } });

    el('gSaveBtn').addEventListener('click', () => {
      readGuidelinesFromForm();
      onSave(state.guidelines);
      el('gSavedNote').textContent = `Saved "${state.guidelines.name}" — will be remembered next time you open Format Check.`;
      setTimeout(() => { el('gSavedNote').textContent = ''; }, 4000);
    });

    el('gExportBtn').addEventListener('click', () => { readGuidelinesFromForm(); onExport(state.guidelines); });

    el('gImportBtn').addEventListener('click', () => el('gImportInput').click());
    el('gImportInput').addEventListener('change', () => {
      const file = el('gImportInput').files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const merged = onImportFile(reader.result);
          state.guidelines = merged;
          writeGuidelinesToForm(state.guidelines);
          el('gSavedNote').textContent = `Loaded guideline file "${file.name}".`;
          setTimeout(() => { el('gSavedNote').textContent = ''; }, 4000);
        } catch (e) {
          alert(`Could not read that guideline file: ${e.message}`);
        }
      };
      reader.readAsText(file);
      el('gImportInput').value = '';
    });

    el('gResetBtn').addEventListener('click', () => {
      if (!confirm('Reset all brand guideline fields to the suite defaults?')) return;
      state.guidelines = onReset();
      writeGuidelinesToForm(state.guidelines);
    });
  }

  function getGuidelines() {
    return readGuidelinesFromForm();
  }

  /* ---------------------------------------------------------------- *
   * Progress / errors
   * ---------------------------------------------------------------- */
  function showProgress(label) { el('progressWrap').classList.remove('hidden'); setProgress(0, label); }
  function setProgress(pct, label) { el('progressFill').style.width = `${Math.min(100, Math.max(0, pct))}%`; if (label) el('progressLabel').textContent = label; }
  function hideProgress() { el('progressWrap').classList.add('hidden'); }
  function showError(msg) { const b = el('errorBox'); b.textContent = msg; b.classList.remove('hidden'); }
  function hideError() { el('errorBox').classList.add('hidden'); }

  /* ---------------------------------------------------------------- *
   * Results
   * ---------------------------------------------------------------- */
  const CATEGORY_SEVERITY = {
    'Font': 'critical', 'Text Color': 'critical', 'Shape Fill Color': 'critical',
    'Font Size': 'warn', 'Line Spacing': 'warn', 'Bullet/Paragraph Spacing': 'warn', 'Shape Margin': 'warn',
  };

  function hideResults() { el('resultsPanel').classList.add('hidden'); }

  function renderResults(result) {
    const cards = [
      [result.totalSlides, 'Total Slides', 'c-cyan'],
      [result.slidesWithIssues, 'Slides With Issues', 'c-coral'],
      [result.totalViolations, 'Total Occurrences', 'c-amber'],
      [result.uniqueViolationRows, 'Unique Issues', 'c-amber'],
    ];
    el('scoreRow').innerHTML = cards.map(([n, l, cls]) => `
      <div class="score-card ${cls}"><div class="num">${n}</div><div class="lbl">${global.FCUtils.escapeHtml(l)}</div></div>
    `).join('');

    el('categoryRow').innerHTML = Object.entries(result.byCategory).map(([cat, count]) => `
      <span class="category-pill">${global.FCUtils.escapeHtml(cat)}: <b>${count}</b></span>
    `).join('') || '<span class="category-pill">No issues found 🎉</span>';

    const bySlide = new Map();
    result.violations.forEach((v) => {
      if (!bySlide.has(v.slide)) bySlide.set(v.slide, []);
      bySlide.get(v.slide).push(v);
    });

    if (!bySlide.size) {
      el('sectionsWrap').innerHTML = '<div class="empty-state">No issues found — every checked slide matches the brand guidelines ✓</div>';
    } else {
      const slideNums = Array.from(bySlide.keys()).sort((a, b) => a - b);
      el('sectionsWrap').innerHTML = slideNums.map((n) => {
        const items = bySlide.get(n);
        const totalOccurrences = items.reduce((s, v) => s + v.count, 0);
        return `
          <div class="section-card">
            <button class="section-head open">
              <span class="section-num">Slide ${n}</span>
              <span class="section-stat dirty">${items.length} issue${items.length > 1 ? 's' : ''} (${totalOccurrences}×)</span>
              <svg class="section-chevron" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
            </button>
            <div class="section-blocks open">
              ${items.map((v) => {
                const sev = CATEGORY_SEVERITY[v.category] || 'warn';
                return `
                <div class="violation-row">
                  <div class="violation-meta">
                    <span class="cat-badge ${sev}">${global.FCUtils.escapeHtml(v.category)}</span>
                    <span class="violation-shape">${global.FCUtils.escapeHtml(v.shapeName)}</span>
                    ${v.count > 1 ? `<span class="violation-shape">× ${v.count}</span>` : ''}
                  </div>
                  <div class="violation-detail"><b>Found:</b> ${global.FCUtils.escapeHtml(v.found)} &nbsp;·&nbsp; <b>Expected:</b> ${global.FCUtils.escapeHtml(v.expected)}</div>
                  <div class="violation-detail">${global.FCUtils.escapeHtml(v.examples[0])}</div>
                </div>`;
              }).join('')}
            </div>
          </div>`;
      }).join('');
    }

    el('resultsPanel').classList.remove('hidden');
    el('downloadBtn').disabled = false;
    setTimeout(() => {
      if (typeof el('resultsPanel').scrollIntoView === 'function') {
        el('resultsPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }, 80);
  }

  document.addEventListener('click', (e) => {
    const head = e.target.closest('.section-head');
    if (!head) return;
    head.classList.toggle('open');
    const blocks = head.nextElementSibling;
    if (blocks) blocks.classList.toggle('open');
  });

  function reset() {
    setDeckFile(null);
    el('deckInput').value = '';
    hideProgress();
    hideError();
    hideResults();
    el('downloadBtn').disabled = true;
  }

  function getDeckFile() { return state.deckFile; }

  global.FCUi = {
    wireUpload, setDeckFile, getDeckFile,
    wireGuidelineForm, getGuidelines,
    showProgress, setProgress, hideProgress,
    showError, hideError,
    renderResults, hideResults,
    reset,
  };
})(window);
