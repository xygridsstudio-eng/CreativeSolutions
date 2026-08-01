/**
 * excel.js
 * Single responsibility: turn a checkDocument() result into a styled
 * .xlsx workbook (Summary + Violations sheets) and return it as a Blob.
 */
(function (global) {
  'use strict';

  const COLORS = {
    headerFill: 'FF1B2431',
    headerFont: 'FFFFFFFF',
    critical: 'FFF7D6D4',
    criticalFont: 'FF8A2E2A',
    warn: 'FFFCEBD0',
    warnFont: 'FF8A5A0A',
  };

  const CATEGORY_SEVERITY = {
    'Font': 'critical',
    'Text Color': 'critical',
    'Shape Fill Color': 'critical',
    'Font Size': 'warn',
    'Line Spacing': 'warn',
    'Bullet/Paragraph Spacing': 'warn',
    'Shape Margin': 'warn',
  };

  function styleHeaderRow(sheet) {
    const row = sheet.getRow(1);
    row.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: COLORS.headerFont } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.headerFill } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    });
    row.height = 20;
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
  }

  function autoFitColumns(sheet, columns) {
    columns.forEach((col, idx) => {
      let max = (col.header || '').length + 2;
      sheet.eachRow({ includeEmpty: false }, (row) => {
        const v = row.getCell(idx + 1).value;
        const s = v == null ? '' : String(v);
        max = Math.max(max, Math.min(s.length, 60));
      });
      sheet.getColumn(idx + 1).width = Math.min(Math.max(max + 2, 10), 55);
    });
  }

  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function buildSummarySheet(workbook, result, guidelines) {
    const sheet = workbook.addWorksheet('Summary');
    sheet.columns = [{ width: 32 }, { width: 50 }];

    const title = sheet.addRow(['Format Check — Slide Consistency Report']);
    sheet.mergeCells('A1:B1');
    title.getCell(1).font = { bold: true, size: 16, color: { argb: 'FF1B2431' } };
    title.height = 26;

    sheet.addRow([]);
    [
      ['File Checked', result.fileName || ''],
      ['Brand Guideline', guidelines.name || ''],
      ['Generated On', new Date().toLocaleString()],
    ].forEach((r) => { const row = sheet.addRow(r); row.getCell(1).font = { bold: true }; });

    sheet.addRow([]);
    const statsHeader = sheet.addRow(['Metric', 'Value']);
    statsHeader.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: COLORS.headerFont } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.headerFill } };
    });
    [
      ['Total Slides', result.totalSlides],
      ['Slides With Issues', result.slidesWithIssues],
      ['Slides Clean', result.totalSlides - result.slidesWithIssues],
      ['Total Occurrences', result.totalViolations],
      ['Unique Issues Found', result.uniqueViolationRows],
    ].forEach((r) => sheet.addRow(r));

    sheet.addRow([]);
    const catHeader = sheet.addRow(['Issues by Category', 'Count']);
    catHeader.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: COLORS.headerFont } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.headerFill } };
    });
    Object.entries(result.byCategory).forEach(([cat, count]) => {
      const row = sheet.addRow([cat, count]);
      const sev = CATEGORY_SEVERITY[cat];
      if (sev) {
        const fill = sev === 'critical' ? COLORS.critical : COLORS.warn;
        const font = sev === 'critical' ? COLORS.criticalFont : COLORS.warnFont;
        row.eachCell((c) => {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
          c.font = { color: { argb: font } };
        });
      }
    });

    sheet.addRow([]);
    const gHeader = sheet.addRow(['Guideline Rules Applied', '']);
    gHeader.getCell(1).font = { bold: true };
    sheet.addRow(['Allowed Fonts', (guidelines.fonts.allowed || []).join(', ')]);
    sheet.addRow(['Allowed Font Sizes', (guidelines.fontSizes.allowed || []).map((s) => s + 'pt').join(', ')]);
    sheet.addRow(['Minimum Body Size', `${guidelines.fontSizes.minBodyPt}pt`]);
    sheet.addRow(['Allowed Colors', (guidelines.colors.allowed || []).join(', ')]);
    sheet.addRow(['Target Line Spacing', `${guidelines.lineSpacing.targetPct}% (± ${guidelines.lineSpacing.tolerancePct}%)`]);
    sheet.addRow(['Paragraph Spacing (Before/After)', `${guidelines.paragraphSpacing.spaceBeforePt}pt / ${guidelines.paragraphSpacing.spaceAfterPt}pt (± ${guidelines.paragraphSpacing.tolerancePt}pt)`]);
    sheet.addRow(['Shape Margins (L/T/R/B)', `${guidelines.shapeMargins.leftIn}" / ${guidelines.shapeMargins.topIn}" / ${guidelines.shapeMargins.rightIn}" / ${guidelines.shapeMargins.bottomIn}" (± ${guidelines.shapeMargins.toleranceIn}")`]);

    return sheet;
  }

  function buildViolationsSheet(workbook, result) {
    const columns = [
      { header: 'Slide #', key: 'slide', width: 9 },
      { header: 'Shape', key: 'shape', width: 22 },
      { header: 'Category', key: 'category', width: 20 },
      { header: 'Found', key: 'found', width: 28 },
      { header: 'Expected', key: 'expected', width: 32 },
      { header: 'Occurrences', key: 'count', width: 12 },
      { header: 'Example Text', key: 'example', width: 45 },
    ];
    const sheet = workbook.addWorksheet('Violations', { views: [{ state: 'frozen', ySplit: 1 }] });
    sheet.columns = columns;

    result.violations.forEach((v) => {
      const row = sheet.addRow({
        slide: v.slide,
        shape: v.shapeName,
        category: v.category,
        found: v.found,
        expected: v.expected,
        count: v.count,
        example: v.examples.join(' | '),
      });
      const sev = CATEGORY_SEVERITY[v.category];
      if (sev) {
        const fill = sev === 'critical' ? COLORS.critical : COLORS.warn;
        const font = sev === 'critical' ? COLORS.criticalFont : COLORS.warnFont;
        row.eachCell((c) => {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
          c.font = Object.assign({ color: { argb: font } }, c.font || {});
          c.alignment = { vertical: 'top', wrapText: true };
        });
      } else {
        row.eachCell((c) => { c.alignment = { vertical: 'top', wrapText: true }; });
      }
    });

    styleHeaderRow(sheet);
    autoFitColumns(sheet, columns);
    if (result.violations.length) sheet.autoFilter = { from: 'A1', to: `${colLetter(columns.length)}1` };
    return sheet;
  }

  async function generateWorkbook(result, guidelines) {
    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = 'Format Check';
    workbook.created = new Date();

    buildSummarySheet(workbook, result, guidelines);
    buildViolationsSheet(workbook, result);

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  global.FCExcel = { generateWorkbook };
})(window);
