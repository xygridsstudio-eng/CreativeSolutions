/**
 * excel.js
 * Single responsibility: turn a report model (report.js) into a styled,
 * multi-worksheet .xlsx workbook and trigger a browser download.
 *
 * Uses ExcelJS (lib/exceljs.min.js -> window.ExcelJS) because, unlike the
 * community build of SheetJS, it reliably writes cell fills/fonts/borders,
 * autofilters and frozen panes from the browser with no server round-trip.
 */
(function (global) {
  'use strict';

  const U = global.CCUtils;

  // PepsiCo Digital Minikit palette — the "50" step of each ramp as a light
  // Excel cell fill (spreadsheets are a light-background context, unlike
  // the app's dark theme) paired with that ramp's darkest step as font
  // color, straight from the brand's own core/expanded color values.
  const COLORS = {
    headerFill: 'FF02355A',   // Fresh 500
    headerFont: 'FFFFFFFF',
    match: 'FFDDEEB9',        // Leaf 50
    matchFont: 'FF0F440E',    // Leaf 500
    modified: 'FFFFF4D6',     // Grain 50
    modifiedFont: 'FF613305', // Grain 600
    missing: 'FFFBDFE2',      // Grapefruit 50
    missingFont: 'FF450D16',  // Grapefruit 600
    added: 'FFF3E6F9',        // Plum 50
    addedFont: 'FF270D2B',    // Plum 600
  };

  function statusFill(status) {
    switch (status) {
      case 'match': return { fill: COLORS.match, font: COLORS.matchFont };
      case 'modified': return { fill: COLORS.modified, font: COLORS.modifiedFont };
      case 'missing': return { fill: COLORS.missing, font: COLORS.missingFont };
      case 'added': return { fill: COLORS.added, font: COLORS.addedFont };
      default: return null;
    }
  }

  function styleHeaderRow(sheet) {
    const row = sheet.getRow(1);
    row.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: COLORS.headerFont } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.headerFill } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } } };
    });
    row.height = 20;
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
  }

  function autoFitColumns(sheet, columns) {
    columns.forEach((col, idx) => {
      const key = col.key;
      let max = (col.header || '').length + 2;
      sheet.eachRow({ includeEmpty: false }, (row) => {
        const cell = row.getCell(idx + 1);
        const v = cell.value == null ? '' : String(cell.value);
        // Cap very long text blocks so the sheet stays readable; wrapText handles overflow.
        max = Math.max(max, Math.min(v.length, 60));
      });
      sheet.getColumn(idx + 1).width = Math.min(Math.max(max + 2, 10), 62);
    });
  }

  function addAutoFilter(sheet, lastColLetter, rowCount) {
    if (rowCount < 1) return;
    sheet.autoFilter = { from: 'A1', to: `${lastColLetter}1` };
  }

  function colLetter(n) {
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function addSheetFromRows(workbook, name, columns, rows, options = {}) {
    const sheet = workbook.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    sheet.columns = columns;
    rows.forEach((r) => {
      const row = sheet.addRow(r);
      if (options.colorByStatusKey) {
        const status = r[options.colorByStatusKey];
        const style = statusFill(status);
        if (style) {
          row.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: style.fill } };
            cell.font = Object.assign({ color: { argb: style.font } }, cell.font || {});
          });
        }
      }
      row.eachCell((cell) => { cell.alignment = { vertical: 'top', wrapText: true }; });
    });
    styleHeaderRow(sheet);
    autoFitColumns(sheet, columns);
    addAutoFilter(sheet, colLetter(columns.length), rows.length);
    return sheet;
  }

  function buildSummarySheet(workbook, report) {
    const sheet = workbook.addWorksheet('Summary');
    sheet.columns = [{ width: 30 }, { width: 46 }];

    const title = sheet.addRow(['Content Check — Comparison Summary']);
    sheet.mergeCells('A1:B1');
    title.getCell(1).font = { bold: true, size: 16, color: { argb: 'FF1F2937' } };
    title.height = 26;

    sheet.addRow([]);
    const meta = report.meta || {};
    const metaRows = [
      ['Source File', meta.sourceName || ''],
      ['Output File', meta.outputName || ''],
      ['Generated On', report.generatedAt.toLocaleString()],
      ['Processing Time', meta.processingTime || ''],
    ];
    metaRows.forEach((r) => {
      const row = sheet.addRow(r);
      row.getCell(1).font = { bold: true };
    });

    sheet.addRow([]);
    const statsHeader = sheet.addRow(['Metric', 'Value']);
    statsHeader.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: COLORS.headerFont } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.headerFill } };
    });

    const s = report.stats;
    const statRows = [
      ['Total Sections', s.totalSections],
      ['Total Paragraphs', s.totalParagraphs],
      ['Total Sentences', s.totalSentences],
      ['Matched', s.matched],
      ['Modified', s.modified],
      ['Missing', s.missing],
      ['Added', s.added],
      ['Overall Match %', `${s.matchPercentage}%`],
    ];
    const statusColorKey = { Matched: 'match', Modified: 'modified', Missing: 'missing', Added: 'added' };
    statRows.forEach(([label, value]) => {
      const row = sheet.addRow([label, value]);
      const style = statusFill(statusColorKey[label]);
      if (style) {
        row.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: style.fill } };
          cell.font = { color: { argb: style.font }, bold: true };
        });
      }
    });

    sheet.addRow([]);
    const legendHeader = sheet.addRow(['Color Legend', '']);
    legendHeader.getCell(1).font = { bold: true };
    [['Match', 'match'], ['Modified', 'modified'], ['Missing', 'missing'], ['Added', 'added']].forEach(([label, key]) => {
      const row = sheet.addRow([label, '']);
      const style = statusFill(key);
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: style.fill } };
      row.getCell(1).font = { color: { argb: style.font } };
    });

    return sheet;
  }

  function detailToRow(r) {
    let sourceText = r.sourceText || '';
    let outputText = r.outputText || '';
    let comments = r.comments || '';

    if ((r.status || '').toLowerCase() === 'modified') {
      const trimmed = U.trimAroundChange(sourceText, outputText, { contextWords: 4, lengthThreshold: 20 });
      sourceText = trimmed.source;
      outputText = trimmed.output;
      if (trimmed.trimmed) {
        comments = comments
          ? `${comments} (long sentence trimmed to change context)`
          : 'Long sentence trimmed to change context';
      }
    }

    return {
      blockId: r.blockId,
      section: r.section,
      page: r.page || '',
      slide: r.slide || '',
      paragraph: r.paragraphNumber || '',
      sentence: r.sentenceNumber || '',
      status: (r.status || '').charAt(0).toUpperCase() + (r.status || '').slice(1),
      sourceText,
      outputText,
      comments,
    };
  }

  const DETAIL_COLUMNS = [
    { header: 'Block ID', key: 'blockId', width: 10 },
    { header: 'Section', key: 'section', width: 24 },
    { header: 'Page', key: 'page', width: 8 },
    { header: 'Slide', key: 'slide', width: 8 },
    { header: 'Paragraph #', key: 'paragraph', width: 12 },
    { header: 'Sentence #', key: 'sentence', width: 12 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Source Text', key: 'sourceText', width: 45 },
    { header: 'Output Text', key: 'outputText', width: 45 },
    { header: 'Comments', key: 'comments', width: 30 },
  ];

  async function generateWorkbook(report) {
    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = 'Content Check Tool';
    workbook.created = new Date();

    buildSummarySheet(workbook, report);

    addSheetFromRows(
      workbook, 'Modified Content', DETAIL_COLUMNS,
      report.modifiedContent.map(detailToRow)
    ).eachRow((row, i) => { if (i > 1) row.eachCell((c) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.modified } }; c.font = { color: { argb: COLORS.modifiedFont } }; }); });

    const detailSheet = addSheetFromRows(
      workbook, 'Detailed Comparison', DETAIL_COLUMNS,
      report.detailRows.map(detailToRow)
    );
    detailSheet.eachRow((row, i) => {
      if (i === 1) return;
      const status = (row.getCell(7).value || '').toString().toLowerCase();
      const style = statusFill(status);
      if (style) {
        row.eachCell((c) => {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: style.fill } };
          c.font = { color: { argb: style.font } };
        });
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  global.CCExcel = { generateWorkbook };
})(window);
