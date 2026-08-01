/**
 * app.js
 * Single responsibility: orchestrate the pipeline (Upload -> Parse ->
 * Check -> Render -> Export) and glue ui.js to the other modules.
 */
(function () {
  'use strict';

  let lastResult = null;
  let lastGuidelines = null;

  function nextFrame() {
    return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
  }

  async function runCheck() {
    const file = window.FCUi.getDeckFile();
    if (!file) return;

    window.FCUi.hideError();
    window.FCUi.hideResults();
    window.FCUi.showProgress('Reading file…');

    try {
      window.FCUi.setProgress(15, `Unzipping ${file.name}…`);
      await nextFrame();
      const model = await window.FCParser.parsePptx(file);

      window.FCUi.setProgress(55, `Checking ${model.slides.length} slide(s) against brand guidelines…`);
      await nextFrame();
      const guidelines = window.FCUi.getGuidelines();
      const result = window.FCChecker.checkDocument(model, guidelines);

      window.FCUi.setProgress(100, 'Done');
      setTimeout(() => window.FCUi.hideProgress(), 350);

      lastResult = result;
      lastGuidelines = guidelines;
      window.FCUi.renderResults(result);
    } catch (err) {
      console.error(err);
      window.FCUi.hideProgress();
      window.FCUi.showError(
        `Format check failed: ${err && err.message ? err.message : err}. ` +
        `Make sure this is a valid, non-password-protected .pptx file.`
      );
    }
  }

  async function downloadReport() {
    if (!lastResult) return;
    const btn = document.getElementById('downloadBtn');
    const original = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Generating…';
    try {
      const blob = await window.FCExcel.generateWorkbook(lastResult, lastGuidelines);
      const base = (lastResult.fileName || 'deck').replace(/\.[^.]+$/, '');
      window.FCUtils.downloadBlob(blob, `FormatCheck_Report_${base}.xlsx`);
    } catch (err) {
      console.error(err);
      window.FCUi.showError(`Could not generate the Excel report: ${err && err.message ? err.message : err}`);
    } finally {
      btn.disabled = false;
      btn.textContent = original;
    }
  }

  function exportGuidelines(guidelines) {
    const blob = window.FCGuidelines.toJsonBlob(guidelines);
    const safeName = (guidelines.name || 'brand-guidelines').replace(/[^a-z0-9]+/gi, '-').toLowerCase();
    window.FCUtils.downloadBlob(blob, `${safeName}.json`);
  }

  document.addEventListener('DOMContentLoaded', () => {
    window.FCUi.wireUpload((file) => window.FCUi.setDeckFile(file));

    window.FCUi.wireGuidelineForm({
      onSave: (g) => window.FCGuidelines.save(g),
      onExport: (g) => exportGuidelines(g),
      onImportFile: (text) => window.FCGuidelines.fromJsonText(text),
      onReset: () => {
        window.FCGuidelines.clear();
        return window.FCGuidelines.load();
      },
    });

    document.getElementById('runBtn').addEventListener('click', runCheck);
    document.getElementById('resetRunBtn').addEventListener('click', () => {
      lastResult = null;
      lastGuidelines = null;
      window.FCUi.reset();
    });
    document.getElementById('downloadBtn').addEventListener('click', downloadReport);
  });
})();
