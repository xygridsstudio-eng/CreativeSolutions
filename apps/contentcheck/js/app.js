/**
 * app.js
 * Single responsibility: orchestrate the pipeline (Upload -> Extract ->
 * Normalize -> Compare -> Summarize -> Report) and glue ui.js to the
 * other modules. Contains no parsing/comparison logic itself.
 */
(function () {
  'use strict';

  let lastReport = null;

  async function runComparison(options) {
    const { sourceFile, outputFile } = window.CCUi.getFiles();
    if (!sourceFile || !outputFile) return;

    window.CCUtils.resetIdCounter();
    window.CCUi.hideError();
    window.CCUi.hideSummary();
    window.CCUi.hidePreview();
    window.CCUi.showProgress('Reading files…');

    const t0 = performance.now();
    try {
      window.CCUi.setProgress(10, `Extracting content from ${sourceFile.name}…`);
      const srcRaw = await window.CCParser.parseFile(sourceFile);

      window.CCUi.setProgress(35, `Extracting content from ${outputFile.name}…`);
      const outRaw = await window.CCParser.parseFile(outputFile);

      window.CCUi.setProgress(55, 'Normalizing content…');
      const srcDoc = window.CCNormalizer.buildHierarchy(srcRaw, options);
      const outDoc = window.CCNormalizer.buildHierarchy(outRaw, options);

      window.CCUi.setProgress(70, 'Comparing sentence by sentence…');
      // Yield to the browser so the progress bar actually paints before the
      // (synchronous, CPU-bound) comparison work runs.
      await nextFrame();
      const comparison = window.CCComparer.compareDocuments(srcDoc, outDoc, options);

      window.CCUi.setProgress(92, 'Building summary…');
      const t1 = performance.now();
      const report = window.CCReport.buildReport(comparison, {
        sourceName: sourceFile.name,
        outputName: outputFile.name,
        processingTime: window.CCUtils.formatDuration(t1 - t0),
      });

      lastReport = report;
      window.CCUi.setProgress(100, 'Done');
      setTimeout(() => window.CCUi.hideProgress(), 400);
      window.CCUi.renderSummary(report);
      window.CCUi.renderPreview(report);
    } catch (err) {
      console.error(err);
      window.CCUi.hideProgress();
      window.CCUi.showError(
        `Comparison failed: ${err && err.message ? err.message : err}. ` +
        `Check that both files are valid and, for PDFs, that they contain selectable text (not scanned images).`
      );
    }
  }

  function nextFrame() {
    return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
  }

  async function downloadReport() {
    if (!lastReport) return;
    const btn = document.getElementById('downloadBtn');
    const originalLabel = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Generating…';
    try {
      const blob = await window.CCExcel.generateWorkbook(lastReport);
      const base = (lastReport.meta.sourceName || 'ContentCheck').replace(/\.[^.]+$/, '');
      window.CCUtils.downloadBlob(blob, `ContentCheck_Report_${base}.xlsx`);
    } catch (err) {
      console.error(err);
      window.CCUi.showError(`Could not generate the Excel report: ${err && err.message ? err.message : err}`);
    } finally {
      btn.disabled = false;
      btn.textContent = originalLabel;
    }
  }

  function resetAll() {
    lastReport = null;
  }

  document.addEventListener('DOMContentLoaded', () => {
    window.CCUi.init({
      onCompare: runComparison,
      onReset: resetAll,
      onDownload: downloadReport,
    });
  });
})();
