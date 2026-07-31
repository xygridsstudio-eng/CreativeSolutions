/**
 * parser.js
 * Single responsibility: turn an uploaded File into a raw, structural
 * representation of its content — BEFORE normalization/sentence-splitting.
 *
 * Output shape (consumed by normalizer.js):
 * {
 *   docType: 'docx'|'pptx'|'pdf'|'txt',
 *   fileName: string,
 *   sections: [
 *     {
 *       title: string,          // slide title / heading text / page label
 *       page: number|null,      // page number (pdf)
 *       slide: number|null,     // slide number (pptx)
 *       blocks: [
 *         { type: 'heading'|'paragraph'|'list'|'notes'|'tableRow', text: string,
 *           cells: string[]|null, tableId: string|null, rowIndex: number|null }
 *       ]
 *     }
 *   ]
 * }
 *
 * Adding a new format later = adding one more `parseXxx(file)` function and
 * one branch in `parseFile`.
 */
(function (global) {
  'use strict';

  const U = global.CCUtils;

  // ---------------------------------------------------------------------
  // Entry point
  // ---------------------------------------------------------------------
  async function parseFile(file) {
    const ext = U.fileExtension(file.name);
    switch (ext) {
      case 'docx':
        return parseDocx(file);
      case 'pptx':
        return parsePptx(file);
      case 'pdf':
        return parsePdf(file);
      case 'txt':
        return parseTxt(file);
      default:
        throw new Error(`Unsupported file type: .${ext}. Supported types: docx, pptx, pdf, txt.`);
    }
  }

  // ---------------------------------------------------------------------
  // DOCX (via mammoth.js -> HTML -> DOM walk)
  // ---------------------------------------------------------------------
  async function parseDocx(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await window.mammoth.convertToHtml({ arrayBuffer });
    const html = result.value || '';
    const doc = new DOMParser().parseFromString(html, 'text/html');

    const sections = [];
    let current = { title: 'Document', page: null, slide: null, blocks: [] };
    sections.push(current);
    let headingCount = 0;

    let tableCounter = 0;

    Array.from(doc.body.children).forEach((el) => {
      const tag = el.tagName.toLowerCase();
      if (/^h[1-6]$/.test(tag)) {
        const text = el.textContent.trim();
        if (!text) return;
        headingCount += 1;
        // Start a new section at each heading so the hierarchy mirrors the
        // document's own structure (Document > Section > Paragraph > Sentence).
        current = { title: text, page: null, slide: null, blocks: [] };
        sections.push(current);
        current.blocks.push({ type: 'heading', text });
      } else if (tag === 'p') {
        const text = el.textContent.trim();
        if (text) current.blocks.push({ type: 'paragraph', text });
      } else if (tag === 'ul' || tag === 'ol') {
        Array.from(el.querySelectorAll('li')).forEach((li) => {
          const text = li.textContent.trim();
          if (text) current.blocks.push({ type: 'list', text });
        });
      } else if (tag === 'table') {
        tableCounter += 1;
        const tableId = `docx_table_${tableCounter}`;
        Array.from(el.querySelectorAll('tr')).forEach((tr, rowIndex) => {
          const cells = Array.from(tr.querySelectorAll('td,th')).map((c) => c.textContent.trim());
          if (cells.some((c) => c)) {
            current.blocks.push({
              type: 'tableRow',
              text: cells.join(' | '),
              cells,
              tableId,
              rowIndex,
            });
          }
        });
      } else {
        const text = el.textContent.trim();
        if (text) current.blocks.push({ type: 'paragraph', text });
      }
    });

    // Drop the initial empty placeholder section if headings created real ones.
    const nonEmpty = sections.filter((s) => s.blocks.length > 0);
    return { docType: 'docx', fileName: file.name, sections: nonEmpty.length ? nonEmpty : sections };
  }

  // ---------------------------------------------------------------------
  // PPTX (via JSZip + manual OOXML parsing — no external pptx library needed)
  // ---------------------------------------------------------------------
  async function parsePptx(file) {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await window.JSZip.loadAsync(arrayBuffer);

    const slideFiles = Object.keys(zip.files)
      .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
      .sort((a, b) => slideIndex(a) - slideIndex(b));

    const sections = [];
    const parser = new DOMParser();

    for (const slidePath of slideFiles) {
      const slideNum = slideIndex(slidePath);
      const xmlText = await zip.files[slidePath].async('text');
      const xml = parser.parseFromString(xmlText, 'application/xml');

      const section = { title: `Slide ${slideNum}`, page: null, slide: slideNum, blocks: [] };
      let titleFound = false;
      let tableCounter = 0;

      const shapeNodes = Array.from(xml.getElementsByTagNameNS('*', 'sp'));
      shapeNodes.forEach((sp) => {
        const isTitlePlaceholder = !!sp.querySelector &&
          Array.from(sp.getElementsByTagNameNS('*', 'ph')).some((ph) => {
            const t = ph.getAttribute('type') || '';
            return /title|ctrTitle/i.test(t);
          });

        const textRuns = Array.from(sp.getElementsByTagNameNS('*', 't')).map((n) => n.textContent);
        const text = collapseRuns(textRuns);
        if (!text) return;

        if (isTitlePlaceholder && !titleFound) {
          section.title = text;
          titleFound = true;
          section.blocks.push({ type: 'heading', text });
        } else {
          section.blocks.push({ type: 'paragraph', text });
        }
      });

      // Tables live in graphicFrame > tbl, not in <p:sp>.
      const tables = Array.from(xml.getElementsByTagNameNS('*', 'tbl'));
      tables.forEach((tbl) => {
        tableCounter += 1;
        const tableId = `slide${slideNum}_table_${tableCounter}`;
        const rows = Array.from(tbl.getElementsByTagNameNS('*', 'tr'));
        rows.forEach((tr, rowIndex) => {
          const tcs = Array.from(tr.getElementsByTagNameNS('*', 'tc'));
          const cells = tcs.map((tc) => collapseRuns(Array.from(tc.getElementsByTagNameNS('*', 't')).map((n) => n.textContent)));
          if (cells.some((c) => c)) {
            section.blocks.push({ type: 'tableRow', text: cells.join(' | '), cells, tableId, rowIndex });
          }
        });
      });

      // Speaker notes live in a sibling notesSlideN.xml.
      const notesPath = `ppt/notesSlides/notesSlide${slideNum}.xml`;
      if (zip.files[notesPath]) {
        const notesXmlText = await zip.files[notesPath].async('text');
        const notesXml = parser.parseFromString(notesXmlText, 'application/xml');
        // Skip the slide-number/placeholder text box; keep body text boxes only.
        const notesShapes = Array.from(notesXml.getElementsByTagNameNS('*', 'sp'));
        notesShapes.forEach((sp) => {
          const isBody = Array.from(sp.getElementsByTagNameNS('*', 'ph')).some((ph) => {
            const t = ph.getAttribute('type') || '';
            return /body|undefined|^$/i.test(t) || !t;
          });
          const textRuns = Array.from(sp.getElementsByTagNameNS('*', 't')).map((n) => n.textContent);
          const text = collapseRuns(textRuns);
          if (text && isBody) {
            section.blocks.push({ type: 'notes', text });
          }
        });
      }

      sections.push(section);
    }

    return { docType: 'pptx', fileName: file.name, sections };
  }

  function slideIndex(path) {
    const m = /slide(\d+)\.xml$/.exec(path);
    return m ? parseInt(m[1], 10) : 0;
  }

  function collapseRuns(runs) {
    return runs.join('').replace(/\s+/g, ' ').trim();
  }

  // ---------------------------------------------------------------------
  // PDF (text-based only, via pdf.js). Basic table detection using
  // x-position clustering of text items on the same line.
  // ---------------------------------------------------------------------
  let workerConfigured = false;
  function getPdfjs() {
    if (!window.pdfjsLib) {
      throw new Error('pdf.min.js failed to load (lib/pdf.min.js). PDF comparison is unavailable.');
    }
    if (!workerConfigured) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
      workerConfigured = true;
    }
    return window.pdfjsLib;
  }

  async function parsePdf(file) {
    const pdfjsLib = getPdfjs();
    const arrayBuffer = await file.arrayBuffer();
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    const pdf = await loadingTask.promise;

    const sections = [];
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      const section = { title: `Page ${pageNum}`, page: pageNum, slide: null, blocks: [] };

      // Group items into visual lines by rounded y-coordinate.
      const lines = new Map();
      textContent.items.forEach((item) => {
        const y = Math.round(item.transform[5]);
        if (!lines.has(y)) lines.set(y, []);
        lines.get(y).push(item);
      });
      // pdf.js coordinates increase upward; sort lines top-to-bottom.
      const sortedYs = [...lines.keys()].sort((a, b) => b - a);

      let paragraphBuffer = [];
      let tableCounter = 0;
      let lastY = null;

      const flushParagraph = () => {
        const text = paragraphBuffer.join(' ').replace(/\s+/g, ' ').trim();
        if (text) section.blocks.push({ type: 'paragraph', text });
        paragraphBuffer = [];
      };

      sortedYs.forEach((y) => {
        const items = lines.get(y).sort((a, b) => a.transform[4] - b.transform[4]);
        // Detect a "table-like" line: 3+ items separated by large x gaps.
        const gaps = [];
        for (let i = 1; i < items.length; i++) {
          gaps.push(items[i].transform[4] - (items[i - 1].transform[4] + (items[i - 1].width || 0)));
        }
        const bigGaps = gaps.filter((g) => g > 18).length;
        const looksLikeTable = items.length >= 3 && bigGaps >= 2;

        if (looksLikeTable) {
          flushParagraph();
          tableCounter += 1;
          const cells = [];
          let cellBuf = [items[0].str];
          for (let i = 1; i < items.length; i++) {
            if (gaps[i - 1] > 18) {
              cells.push(cellBuf.join('').trim());
              cellBuf = [items[i].str];
            } else {
              cellBuf.push(items[i].str);
            }
          }
          cells.push(cellBuf.join('').trim());
          if (cells.some((c) => c)) {
            section.blocks.push({
              type: 'tableRow',
              text: cells.join(' | '),
              cells,
              tableId: `page${pageNum}_table_${Math.ceil(tableCounter / 4)}`,
              rowIndex: tableCounter,
            });
          }
        } else {
          const lineText = items.map((it) => it.str).join(' ').replace(/\s+/g, ' ').trim();
          if (!lineText) return;
          // Blank-gap heuristic: a big vertical jump starts a new paragraph.
          if (lastY !== null && lastY - y > 22) flushParagraph();
          paragraphBuffer.push(lineText);
          lastY = y;
        }
      });
      flushParagraph();

      sections.push(section);
    }

    return { docType: 'pdf', fileName: file.name, sections };
  }

  // ---------------------------------------------------------------------
  // TXT
  // ---------------------------------------------------------------------
  async function parseTxt(file) {
    const text = await file.text();
    const rawParagraphs = text.split(/\r?\n\s*\r?\n/); // blank-line separated
    const section = { title: 'Document', page: null, slide: null, blocks: [] };
    rawParagraphs.forEach((p) => {
      const cleaned = p.replace(/\r?\n/g, ' ').trim();
      if (cleaned) section.blocks.push({ type: 'paragraph', text: cleaned });
    });
    return { docType: 'txt', fileName: file.name, sections: [section] };
  }

  global.CCParser = { parseFile };
})(window);
