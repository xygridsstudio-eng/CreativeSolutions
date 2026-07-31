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

    detailRows.forEach((row) => {
      if (row.status === 'missing') missingContent.push(row);
      if (row.status === 'added') addedContent.push(row);
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

    tableRows.forEach((row) => {
      if (row.status === 'missing') missingContent.push({ ...row, sentenceNumber: null, comments: 'Table row removed' });
      if (row.status === 'added') addedContent.push({ ...row, sentenceNumber: null, comments: 'Table row added' });
      if (row.status === 'modified') modifiedContent.push({ ...row, sentenceNumber: null, comments: 'Table row modified' });
    });

    return {
      meta,
      stats,
      detailRows,
      tableRows,
      missingContent,
      addedContent,
      modifiedContent,
      numberChanges,
      dateChanges,
      duplicates,
      generatedAt: new Date(),
    };
  }

  global.CCReport = { buildReport };
})(window);
