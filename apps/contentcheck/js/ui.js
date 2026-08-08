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

    el('compareBtn').addEventListener('click', () => onCompare({}));
    el('resetBtn').addEventListener('click', () => {
      reset();
      onReset();
    });
    el('downloadBtn').addEventListener('click', onDownload);

    // Delegated on the container (not per-header) since preview groups are
    // re-rendered wholesale on every comparison run.
    el('previewGroups').addEventListener('click', (e) => {
      const header = e.target.closest('.preview-group-header');
      if (header) toggleGroup(header);
    });
    el('previewGroups').addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      const header = e.target.closest('.preview-group-header');
      if (!header) return;
      e.preventDefault();
      toggleGroup(header);
    });

    renderFileSlot(el('sourceSlot'), null);
    renderFileSlot(el('outputSlot'), null);
    updateCompareButton();
  }

  function updateCompareButton() {
    el('compareBtn').disabled = !(state.sourceFile && state.outputFile);
  }

  function getFiles() {
    return { sourceFile: state.sourceFile, outputFile: state.outputFile };
  }

  function useBackendParsing() {
    return el('useBackendParsing').checked;
  }

  function reset() {
    state.sourceFile = null;
    state.outputFile = null;
    el('sourceInput').value = '';
    el('outputInput').value = '';
    renderFileSlot(el('sourceSlot'), null);
    renderFileSlot(el('outputSlot'), null);
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
      ['Reformatted', s.reformatted, 'reformatted'],
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

  const STATUS_LABELS = { match: 'Match', modified: 'Modified', missing: 'Missing', added: 'Added', reformatted: 'Reformatted' };

  function hidePreview() {
    el('previewWrap').classList.add('hidden');
  }

  /** Group rows by section, preserving order of first appearance. */
  function groupRowsBySection(rows) {
    const order = [];
    const map = new Map();
    rows.forEach((r) => {
      const key = r.section || '(No section)';
      if (!map.has(key)) {
        // page/slide is a property of the section, not the individual row —
        // grab it from whichever row we see first for this section.
        map.set(key, { section: key, page: r.page, slide: r.slide, rows: [] });
        order.push(key);
      }
      map.get(key).rows.push(r);
    });
    return order.map((key) => map.get(key));
  }

  function sectionLabel(g) {
    const location = g.slide != null ? `Slide ${g.slide}` : (g.page != null ? `Page ${g.page}` : null);
    // Slides/pages with no real title fall back to "Slide N"/"Page N" as the
    // section name itself (see parser.js) — don't double it up as "Slide 1 — Slide 1".
    if (!location || g.section === location) return g.section;
    return `${location} — ${g.section}`;
  }

  function countsByStatus(rows) {
    const counts = { match: 0, modified: 0, missing: 0, added: 0, reformatted: 0 };
    rows.forEach((r) => { if (counts[r.status] != null) counts[r.status] += 1; });
    return counts;
  }

  function renderDiffTokens(tokens) {
    return tokens.map((t) => {
      const text = global.CCUtils.escapeHtml(t.text);
      if (t.type === 'equal') return `<span class="diff-equal">${text}</span>`;
      return `<mark class="diff-${t.type}">${text}</mark>`;
    }).join(' ');
  }

  /**
   * One Word "Track Changes"-style redline for a preview row: unchanged
   * words plain, deleted words struck through, inserted words underlined,
   * all read left-to-right as a single sentence — instead of two separate
   * before/after columns a reviewer has to compare by eye. A wholly missing
   * paragraph reads as entirely struck through (nothing survived); a wholly
   * added one reads as entirely underlined (nothing to compare against);
   * "modified"/"reformatted" interleave both inline, exactly like Word's
   * redline for an edited sentence.
   */
  function renderRowContent(r) {
    if (r.status === 'missing') {
      return `<mark class="diff-removed">${global.CCUtils.escapeHtml(r.sourceText || '')}</mark>`;
    }
    if (r.status === 'added') {
      return `<mark class="diff-added">${global.CCUtils.escapeHtml(r.outputText || '')}</mark>`;
    }
    if (r.status !== 'modified' && r.status !== 'reformatted') {
      return `<span class="diff-equal">${global.CCUtils.escapeHtml(r.sourceText || r.outputText || '')}</span>`;
    }
    const tokens = global.CCUtils.diffWordsMerged(r.sourceText || '', r.outputText || '');
    return renderDiffTokens(tokens);
  }

  function renderPreview(report) {
    const interesting = report.previewRows.filter((r) => r.status !== 'match').slice(0, 200);
    const rows = interesting.length ? interesting : report.previewRows.slice(0, 50);

    el('previewCount').textContent = interesting.length
      ? `Showing ${rows.length} of ${interesting.length} changed items, grouped by section — click a section to expand (full detail in the Excel report)`
      : `No differences found — showing first ${rows.length} matched items`;

    const groups = groupRowsBySection(rows);

    el('previewGroups').innerHTML = groups.map((g) => {
      const counts = countsByStatus(g.rows);
      const countBadges = ['modified', 'missing', 'added', 'reformatted', 'match']
        .filter((key) => counts[key])
        .map((key) => `<span class="badge badge-${key}">${counts[key]} ${STATUS_LABELS[key]}</span>`)
        .join('');
      const rowsHtml = g.rows.map((r) => `
        <tr class="row-${r.status}">
          <td style="width:56px">${r.blockId}</td>
          <td style="width:90px"><span class="badge badge-${r.status}">${STATUS_LABELS[r.status] || r.status}</span></td>
          <td>${renderRowContent(r)}</td>
          <td style="width:160px">${global.CCUtils.escapeHtml(r.comments || '')}</td>
        </tr>
      `).join('');
      return `
        <div class="preview-group">
          <div class="preview-group-header" role="button" tabindex="0" aria-expanded="false">
            <span class="preview-group-chevron">▸</span>
            <span class="preview-group-title">${global.CCUtils.escapeHtml(sectionLabel(g))}</span>
            <span class="preview-group-counts">${countBadges}</span>
          </div>
          <div class="preview-group-body">
            <div class="table-wrap">
              <table class="preview-table">
                <thead>
                  <tr>
                    <th>ID</th>
                    <th>Status</th>
                    <th>Changes</th>
                    <th>Comments</th>
                  </tr>
                </thead>
                <tbody>${rowsHtml}</tbody>
              </table>
            </div>
          </div>
        </div>
      `;
    }).join('');

    el('previewWrap').classList.remove('hidden');
  }

  function toggleGroup(header) {
    const body = header.nextElementSibling;
    const open = header.classList.toggle('open');
    body.classList.toggle('open', open);
    header.setAttribute('aria-expanded', String(open));
  }

  global.CCUi = {
    init,
    getFiles,
    useBackendParsing,
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
