/**
 * report.js
 * Single responsibility: turn the raw comparison result (from comparer.js)
 * into a clean "report model" that both ui.js (on-screen summary) and
 * excel.js (workbook generation) can consume without re-deriving anything.
 */
(function (global) {
  'use strict';

  function buildReport(comparison, meta) {
    const { stats, detailRows, tableRows, duplicates } = comparison;

    const numberChanges = [];
    const dateChanges = [];
    const missingContent = [];
    const addedContent = [];
    const modifiedContent = [];
    const reformattedContent = [];

    detailRows.forEach((row) => {
      if (row.status === 'missing') missingContent.push(row);
      if (row.status === 'added') addedContent.push(row);
      if (row.status === 'reformatted') reformattedContent.push(row);
      if (row.status === 'modified') {
        modifiedContent.push(row);
        if (row.special) {
          if (row.special.number) numberChanges.push({ ...row, changeType: 'Number', ...row.special.number });
          if (row.special.percentage) numberChanges.push({ ...row, changeType: 'Percentage', ...row.special.percentage });
          if (row.special.currency) numberChanges.push({ ...row, changeType: 'Currency', ...row.special.currency });
          if (row.special.fiscalYear) numberChanges.push({ ...row, changeType: 'Fiscal Year', ...row.special.fiscalYear });
          if (row.special.date) dateChanges.push({ ...row, ...row.special.date });
        }
      }
    });

    const tableRowsWithComments = tableRows.map((row) => ({
      ...row,
      sentenceNumber: null,
      // Reconciliation (comparer.js) already sets a specific comment when it
      // relabels a row 'reformatted' — keep that instead of overwriting it.
      comments: row.comments ? row.comments
        : row.status === 'missing' ? 'Table row removed'
        : row.status === 'added' ? 'Table row added'
        : row.status === 'modified' ? 'Table row modified'
        : '',
    }));

    tableRowsWithComments.forEach((row) => {
      if (row.status === 'missing') missingContent.push(row);
      if (row.status === 'added') addedContent.push(row);
      if (row.status === 'reformatted') reformattedContent.push(row);
      if (row.status === 'modified') modifiedContent.push(row);
    });

    // Sentence-level and table-level changes are computed separately (they
    // go through different alignment logic), but the on-screen preview
    // needs one combined, chronological-by-blockId list so table changes
    // are actually visible there instead of only in the Excel report.
    const previewRows = detailRows.concat(tableRowsWithComments).sort((a, b) => a.blockId - b.blockId);

    return {
      meta,
      stats,
      detailRows,
      tableRows,
      previewRows,
      missingContent,
      addedContent,
      modifiedContent,
      reformattedContent,
      numberChanges,
      dateChanges,
      duplicates,
      generatedAt: new Date(),
    };
  }

  global.CCReport = { buildReport };
})(window);
