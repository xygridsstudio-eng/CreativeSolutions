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

  // Opt-in only — see the "Use enhanced parsing" toggle in index.html.
  // Content Check works fully offline without this; it's only used when a
  // user explicitly asks for it, since it means uploading files to a server.
  const BACKEND_URL = 'https://content-check-backend-h62t.onrender.com';

  async function parseFileViaBackend(file) {
    const formData = new FormData();
    formData.append('file', file, file.name);

    let response;
    try {
      response = await fetch(`${BACKEND_URL}/parse`, { method: 'POST', body: formData });
    } catch (e) {
      throw new Error(
        `Could not reach the enhanced parsing server. It may be waking up from sleep (free hosting spins down ` +
        `when idle) — wait a moment and try again, or turn off "Use enhanced parsing" to parse locally instead.`
      );
    }
    if (!response.ok) {
      let detail = '';
      try { detail = (await response.json()).detail || ''; } catch (e) { /* response wasn't JSON */ }
      throw new Error(`Enhanced parsing failed for ${file.name}: ${detail || response.statusText}`);
    }
    return response.json();
  }

  // ---------------------------------------------------------------------
  // Entry point
  // ---------------------------------------------------------------------
  async function parseFile(file, options) {
    if (options && options.useBackend) {
      return parseFileViaBackend(file);
    }
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

  /**
   * Table-cell text, with <br> line breaks turned into spaces first.
   * Element.textContent concatenates descendant text nodes with no
   * separator, so a cell with a manual line break (mammoth renders it as
   * <br>) would otherwise glue the two lines together, e.g. "Make &" and
   * "Move" becoming "Make &Move".
   */
  function cellText(cell) {
    const clone = cell.cloneNode(true);
    clone.querySelectorAll('br').forEach((br) => br.replaceWith(' '));
    return clone.textContent.replace(/\s+/g, ' ').trim();
  }

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
          const cells = Array.from(tr.querySelectorAll('td,th')).map((c) => cellText(c));
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

        // Each <a:p> is one paragraph/bullet line. Extracting text per-paragraph
        // (rather than joining every <a:t> run in the whole shape into one
        // string) keeps separate bullets as separate blocks — joining them
        // used to glue unrelated lines together, e.g. "VF" + "XYZ" + "XYZ"
        // becoming "VFXYZXYZ", which could then never match the corresponding
        // lines in the other document.
        const paraTexts = Array.from(sp.getElementsByTagNameNS('*', 'p'))
          .map((p) => collapseRuns(Array.from(p.getElementsByTagNameNS('*', 't')).map((n) => n.textContent)))
          .filter(Boolean);
        if (!paraTexts.length) return;

        if (isTitlePlaceholder && !titleFound) {
          // A slide title is one semantic unit even if it wraps across
          // multiple paragraphs, so join them instead of splitting them.
          const text = paraTexts.join(' ');
          section.title = text;
          titleFound = true;
          section.blocks.push({ type: 'heading', text });
        } else {
          const type = paraTexts.length > 1 ? 'list' : 'paragraph';
          paraTexts.forEach((text) => section.blocks.push({ type, text }));
        }
      });

      // SmartArt diagrams live in a separate part referenced by graphicFrame,
      // not inline as shape text — see extractDiagramTexts for why this matters.
      const diagramTexts = await extractDiagramTexts(zip, slidePath, xml);
      diagramTexts.forEach((text) => section.blocks.push({ type: 'list', text }));

      // Charts are likewise a separate part, not inline shape text — see
      // extractChartTexts for why this matters.
      const chartTexts = await extractChartTexts(zip, slidePath, xml);
      chartTexts.forEach((text) => section.blocks.push({ type: 'list', text }));

      // Tables live in graphicFrame > tbl, not in <p:sp>.
      const tables = Array.from(xml.getElementsByTagNameNS('*', 'tbl'));
      tables.forEach((tbl) => {
        tableCounter += 1;
        const tableId = `slide${slideNum}_table_${tableCounter}`;
        const rows = Array.from(tbl.getElementsByTagNameNS('*', 'tr'));
        rows.forEach((tr, rowIndex) => {
          const tcs = Array.from(tr.getElementsByTagNameNS('*', 'tc'));
          // Join per-<a:p> paragraph text with a space rather than flattening
          // every <a:t> run in the cell together — a cell with a manual line
          // break (a soft-wrapped label like "Make &" / "Move") is stored as
          // two paragraphs, and collapseRuns() on the raw run list has no
          // separator between them, gluing the lines into "Make &Move".
          const cells = tcs.map((tc) => Array.from(tc.getElementsByTagNameNS('*', 'p'))
            .map((p) => collapseRuns(Array.from(p.getElementsByTagNameNS('*', 't')).map((n) => n.textContent)))
            .filter(Boolean)
            .join(' '));
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
          if (!isBody) return;
          const paraTexts = Array.from(sp.getElementsByTagNameNS('*', 'p'))
            .map((p) => collapseRuns(Array.from(p.getElementsByTagNameNS('*', 't')).map((n) => n.textContent)))
            .filter(Boolean);
          paraTexts.forEach((text) => section.blocks.push({ type: 'notes', text }));
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

  const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /** 'ppt/slides/slide1.xml' -> 'ppt/slides/_rels/slide1.xml.rels' */
  function relsPathFor(partPath) {
    const slash = partPath.lastIndexOf('/');
    const dir = partPath.slice(0, slash);
    const file = partPath.slice(slash + 1);
    return `${dir}/_rels/${file}.rels`;
  }

  /** Resolve a relationship Target (e.g. '../diagrams/data1.xml') against the part that references it. */
  function resolveRelPath(basePartPath, relTarget) {
    if (/^[a-z]+:\/\//i.test(relTarget)) return null; // external target, not in the zip
    const baseParts = basePartPath.split('/');
    baseParts.pop(); // drop the file itself, keep its directory
    relTarget.split('/').forEach((part) => {
      if (part === '..') baseParts.pop();
      else if (part && part !== '.') baseParts.push(part);
    });
    return baseParts.join('/');
  }

  /**
   * SmartArt diagrams are NOT stored inline in the slide XML as <p:sp>
   * shapes — the slide only holds a <p:graphicFrame> pointing (via a
   * relationship id) to a separate ppt/diagrams/dataN.xml part that holds
   * the actual text. Without this, any content redesigned into a SmartArt
   * diagram silently disappears from parsing (extracted as nothing, not
   * even reported as removed) rather than showing up as a real change.
   */
  async function extractDiagramTexts(zip, slidePath, xml) {
    const relIds = Array.from(xml.getElementsByTagNameNS('*', 'graphicFrame'))
      .map((gf) => gf.getElementsByTagNameNS('*', 'graphicData')[0])
      .filter((gd) => gd && /diagram/i.test(gd.getAttribute('uri') || ''))
      .map((gd) => gd.getElementsByTagNameNS('*', 'relIds')[0])
      .filter(Boolean)
      .map((el) => el.getAttributeNS(RELATIONSHIPS_NS, 'dm'))
      .filter(Boolean);
    if (!relIds.length) return [];

    const relsPath = relsPathFor(slidePath);
    if (!zip.files[relsPath]) return [];
    const relsXml = new DOMParser().parseFromString(await zip.files[relsPath].async('text'), 'application/xml');
    const targetById = new Map(
      Array.from(relsXml.getElementsByTagNameNS('*', 'Relationship')).map((r) => [r.getAttribute('Id'), r.getAttribute('Target')])
    );

    const texts = [];
    for (const relId of relIds) {
      const target = targetById.get(relId);
      const dataPath = target && resolveRelPath(slidePath, target);
      if (!dataPath || !zip.files[dataPath]) continue;
      const dataXml = new DOMParser().parseFromString(await zip.files[dataPath].async('text'), 'application/xml');
      Array.from(dataXml.getElementsByTagNameNS('*', 'pt')).forEach((pt) => {
        // Non-"node" points (parTrans/sibTrans/doc) are structural connectors
        // with no real content, not actual diagram text.
        const ptType = pt.getAttribute('type');
        if (ptType && ptType !== 'node') return;
        Array.from(pt.getElementsByTagNameNS('*', 'p'))
          .map((p) => collapseRuns(Array.from(p.getElementsByTagNameNS('*', 't')).map((n) => n.textContent)))
          .filter(Boolean)
          .forEach((text) => texts.push(text));
      });
    }
    return texts;
  }

  /** idx -> trimmed text, from a <c:cat>/<c:val>'s <c:pt idx="N"><c:v>...</c:v></c:pt> points. */
  function chartPointsByIdx(container) {
    const map = new Map();
    if (!container) return map;
    Array.from(container.getElementsByTagNameNS('*', 'pt')).forEach((pt) => {
      const idx = pt.getAttribute('idx');
      const v = pt.getElementsByTagNameNS('*', 'v')[0];
      if (idx != null && v) map.set(idx, v.textContent.trim());
    });
    return map;
  }

  function chartTitleText(chartXml) {
    const titleEl = chartXml.getElementsByTagNameNS('*', 'title')[0];
    if (!titleEl) return '';
    return collapseRuns(Array.from(titleEl.getElementsByTagNameNS('*', 't')).map((n) => n.textContent));
  }

  /** One line per series: "<series name> — <category>: <value>, <category>: <value>, ...". */
  function chartSeriesLines(chartXml) {
    return Array.from(chartXml.getElementsByTagNameNS('*', 'ser')).map((ser) => {
      const txEl = ser.getElementsByTagNameNS('*', 'tx')[0];
      const nameV = txEl ? txEl.getElementsByTagNameNS('*', 'v')[0] : null;
      const seriesName = nameV ? nameV.textContent.trim() : '';

      const categories = chartPointsByIdx(ser.getElementsByTagNameNS('*', 'cat')[0]);
      const values = chartPointsByIdx(ser.getElementsByTagNameNS('*', 'val')[0]);
      const idxs = Array.from(new Set([...categories.keys(), ...values.keys()])).sort((a, b) => Number(a) - Number(b));
      const pairs = idxs.map((idx) => `${categories.get(idx) || `#${idx}`}: ${values.get(idx) != null ? values.get(idx) : ''}`);

      return ((seriesName ? `${seriesName} — ` : '') + pairs.join(', ')).trim();
    }).filter(Boolean);
  }

  /**
   * Chart data is likewise not inline in the slide XML — a chart's
   * graphicFrame only holds <c:chart r:id="rIdX"/>, pointing to a separate
   * ppt/charts/chartN.xml part with the actual title/series/category/value
   * data. Without this, edits to a chart's underlying numbers (the whole
   * point of a chart) are completely invisible to the comparison.
   * Covers standard chart types (bar/line/pie/etc. via <c:cat>/<c:val>);
   * modern "chartEx" types (waterfall, funnel, ...) use a different, not
   * currently supported, schema and a distinct graphicData uri, so they're
   * safely skipped rather than mis-parsed.
   */
  async function extractChartTexts(zip, slidePath, xml) {
    const relIds = Array.from(xml.getElementsByTagNameNS('*', 'graphicFrame'))
      .map((gf) => gf.getElementsByTagNameNS('*', 'graphicData')[0])
      .filter((gd) => gd && /\/chart$/i.test(gd.getAttribute('uri') || ''))
      .map((gd) => gd.getElementsByTagNameNS('*', 'chart')[0])
      .filter(Boolean)
      .map((el) => el.getAttributeNS(RELATIONSHIPS_NS, 'id'))
      .filter(Boolean);
    if (!relIds.length) return [];

    const relsPath = relsPathFor(slidePath);
    if (!zip.files[relsPath]) return [];
    const relsXml = new DOMParser().parseFromString(await zip.files[relsPath].async('text'), 'application/xml');
    const targetById = new Map(
      Array.from(relsXml.getElementsByTagNameNS('*', 'Relationship')).map((r) => [r.getAttribute('Id'), r.getAttribute('Target')])
    );

    const texts = [];
    for (const relId of relIds) {
      const target = targetById.get(relId);
      const chartPath = target && resolveRelPath(slidePath, target);
      if (!chartPath || !zip.files[chartPath]) continue;
      const chartXml = new DOMParser().parseFromString(await zip.files[chartPath].async('text'), 'application/xml');
      const title = chartTitleText(chartXml);
      if (title) texts.push(`Chart title: ${title}`);
      texts.push(...chartSeriesLines(chartXml));
    }
    return texts;
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

  /**
   * Join PDF text items into one string, inserting a space only where the
   * original PDF actually had a visible horizontal gap. Text is often split
   * into several items at hyphen/kerning-pair boundaries with zero or even
   * negative gap between them (still part of the same word) — blindly
   * joining every item with a space turned "build-up" into "build - up".
   */
  function joinTextItems(items) {
    let out = '';
    for (let i = 0; i < items.length; i++) {
      if (i > 0) {
        const prev = items[i - 1];
        const gap = items[i].transform[4] - (prev.transform[4] + (prev.width || 0));
        if (gap > 1) out += ' ';
      }
      out += items[i].str;
    }
    return out.replace(/\s+/g, ' ').trim();
  }

  const PDF_BULLET_RE = /^[•▪◦‣∙·–—\-\*]\s+/;

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
          const cellGroups = [[items[0]]];
          for (let i = 1; i < items.length; i++) {
            if (gaps[i - 1] > 18) {
              cellGroups.push([items[i]]);
            } else {
              cellGroups[cellGroups.length - 1].push(items[i]);
            }
          }
          const cells = cellGroups.map((group) => joinTextItems(group));
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
          const lineText = joinTextItems(items);
          if (!lineText) return;
          // Blank-gap heuristic: a big vertical jump starts a new paragraph.
          // A bulleted line is always its own paragraph too, regardless of
          // gap size — bullet lists commonly use the same tight spacing
          // between separate bullets as between wrapped lines within one
          // bullet, so the vertical gap alone can't tell those apart.
          const isBullet = PDF_BULLET_RE.test(lineText);
          if (lastY !== null && (lastY - y > 22 || isBullet)) flushParagraph();
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
