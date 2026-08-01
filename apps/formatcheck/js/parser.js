/**
 * parser.js
 * Single responsibility: read a .pptx (a zip of OOXML) and produce a plain
 * JS model of every slide's shapes, paragraphs, and text runs, with their
 * relevant formatting properties (font, size, color, spacing, margins)
 * already resolved to concrete values — including basic theme-color
 * resolution (a:schemeClr -> actual hex via the slide's theme).
 *
 * Scope note: this reads *explicit* formatting on each element and the
 * slide's theme color scheme. It does not fully resolve slide-master /
 * layout placeholder inheritance chains (a real PPTX renderer's full
 * cascade), so a run that inherits its font/size purely from the master
 * with no direct or layout-level override may not be flagged even if it
 * differs from the guideline. This keeps the checker fast and dependency-free
 * while still catching the overwhelming majority of real-world drift,
 * which happens via direct formatting on runs/shapes.
 */
(function (global) {
  'use strict';

  const EMU_PER_INCH = 914400;
  const NS = {
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  };

  function q(el, tag) {
    return el ? Array.from(el.getElementsByTagNameNS('*', tag)) : [];
  }
  function qDirect(el, tag) {
    // Direct children only (avoid descending into nested txBody, e.g. table cells)
    if (!el) return [];
    return Array.from(el.childNodes).filter(
      (n) => n.nodeType === 1 && n.localName === tag
    );
  }
  function firstDirect(el, tag) {
    const list = qDirect(el, tag);
    return list.length ? list[0] : null;
  }
  function emuToInches(v) {
    return v == null ? null : Math.round((Number(v) / EMU_PER_INCH) * 1000) / 1000;
  }
  function centiToPt(v) {
    // a:rPr sz is in hundredths of a point
    return v == null ? null : Math.round(Number(v)) / 100;
  }

  /* ---------------------------------------------------------------- *
   * Theme color resolution
   * ---------------------------------------------------------------- */
  function parseTheme(themeXmlDoc) {
    const scheme = {};
    if (!themeXmlDoc) return scheme;
    const clrScheme = themeXmlDoc.getElementsByTagNameNS('*', 'clrScheme')[0];
    if (!clrScheme) return scheme;
    Array.from(clrScheme.childNodes)
      .filter((n) => n.nodeType === 1)
      .forEach((node) => {
        const name = node.localName; // dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink
        const srgb = node.getElementsByTagNameNS('*', 'srgbClr')[0];
        const sysClr = node.getElementsByTagNameNS('*', 'sysClr')[0];
        if (srgb) scheme[name] = '#' + srgb.getAttribute('val').toUpperCase();
        else if (sysClr) scheme[name] = '#' + (sysClr.getAttribute('lastClr') || '000000').toUpperCase();
      });
    // Standard PowerPoint scheme-name aliasing used inside slide XML
    const alias = {
      bg1: scheme.lt1, tx1: scheme.dk1, bg2: scheme.lt2, tx2: scheme.dk2,
      accent1: scheme.accent1, accent2: scheme.accent2, accent3: scheme.accent3,
      accent4: scheme.accent4, accent5: scheme.accent5, accent6: scheme.accent6,
      hlink: scheme.hlink, folHlink: scheme.folHlink,
    };
    return Object.assign({}, scheme, alias);
  }

  function resolveColor(solidFillEl, theme) {
    if (!solidFillEl) return null;
    const srgb = firstDirect(solidFillEl, 'srgbClr') || solidFillEl.getElementsByTagNameNS('*', 'srgbClr')[0];
    if (srgb) return '#' + srgb.getAttribute('val').toUpperCase();
    const scheme = firstDirect(solidFillEl, 'schemeClr') || solidFillEl.getElementsByTagNameNS('*', 'schemeClr')[0];
    if (scheme) {
      const name = scheme.getAttribute('val');
      return theme[name] || null;
    }
    return null;
  }

  /* ---------------------------------------------------------------- *
   * Per-shape / paragraph / run extraction
   * ---------------------------------------------------------------- */
  function extractBodyPr(spEl) {
    const txBody = spEl.getElementsByTagNameNS('*', 'txBody')[0] || spEl.getElementsByTagNameNS('*', 'bodyPr')[0];
    const bodyPr = spEl.getElementsByTagNameNS('*', 'bodyPr')[0];
    if (!bodyPr) return null;
    const get = (attr) => (bodyPr.hasAttribute(attr) ? Number(bodyPr.getAttribute(attr)) : null);
    return {
      lIns: emuToInches(get('lIns')),
      tIns: emuToInches(get('tIns')),
      rIns: emuToInches(get('rIns')),
      bIns: emuToInches(get('bIns')),
      // OOXML defaults when the attribute is simply absent
      lInsExplicit: bodyPr.hasAttribute('lIns'),
      tInsExplicit: bodyPr.hasAttribute('tIns'),
      rInsExplicit: bodyPr.hasAttribute('rIns'),
      bInsExplicit: bodyPr.hasAttribute('bIns'),
    };
  }

  function extractLineSpacing(pPrEl) {
    if (!pPrEl) return null;
    const lnSpc = firstDirect(pPrEl, 'lnSpc');
    if (!lnSpc) return null;
    const pct = firstDirect(lnSpc, 'spcPct');
    if (pct) return { type: 'pct', value: Number(pct.getAttribute('val')) / 1000 };
    const pts = firstDirect(lnSpc, 'spcPts');
    if (pts) return { type: 'pts', value: centiToPt(pts.getAttribute('val')) };
    return null;
  }

  function extractParaSpacing(pPrEl, tag) {
    if (!pPrEl) return null;
    const node = firstDirect(pPrEl, tag); // spcBef | spcAft
    if (!node) return null;
    const pts = firstDirect(node, 'spcPts');
    if (pts) return { type: 'pts', value: centiToPt(pts.getAttribute('val')) };
    const pct = firstDirect(node, 'spcPct');
    if (pct) return { type: 'pct', value: Number(pct.getAttribute('val')) / 1000 };
    return null;
  }

  function textOf(runEl) {
    const t = runEl.getElementsByTagNameNS('*', 't')[0];
    return t ? t.textContent : '';
  }

  function parseSlideXml(xmlDoc, theme, slideNumber) {
    const shapes = [];
    const spNodes = Array.from(xmlDoc.getElementsByTagNameNS('*', 'sp'));

    spNodes.forEach((sp, shapeIdx) => {
      const nvPr = sp.getElementsByTagNameNS('*', 'cNvPr')[0];
      const shapeName = nvPr ? nvPr.getAttribute('name') : `Shape ${shapeIdx + 1}`;

      // Shape fill color (spPr > solidFill), skipping line/outline fills
      const spPr = sp.getElementsByTagNameNS('*', 'spPr')[0];
      let shapeFill = null;
      if (spPr) {
        const directFill = firstDirect(spPr, 'solidFill');
        if (directFill) shapeFill = resolveColor(directFill, theme);
      }

      const bodyPr = extractBodyPr(sp);
      const txBody = sp.getElementsByTagNameNS('*', 'txBody')[0];
      const paragraphs = [];

      if (txBody) {
        const pNodes = qDirect(txBody, 'p');
        pNodes.forEach((pEl, paraIdx) => {
          const pPr = firstDirect(pEl, 'pPr');
          const marL = pPr && pPr.hasAttribute('marL') ? emuToInches(Number(pPr.getAttribute('marL'))) : null;
          const indent = pPr && pPr.hasAttribute('indent') ? emuToInches(Number(pPr.getAttribute('indent'))) : null;
          const lnSpc = extractLineSpacing(pPr);
          const spcBef = extractParaSpacing(pPr, 'spcBef');
          const spcAft = extractParaSpacing(pPr, 'spcAft');
          const hasBullet = !!(pPr && (firstDirect(pPr, 'buChar') || firstDirect(pPr, 'buAutoNum')));

          const runs = qDirect(pEl, 'r').map((rEl) => {
            const rPr = firstDirect(rEl, 'rPr');
            const latin = rPr ? rPr.getElementsByTagNameNS('*', 'latin')[0] : null;
            const sz = rPr && rPr.hasAttribute('sz') ? centiToPt(rPr.getAttribute('sz')) : null;
            const solidFill = rPr ? firstDirect(rPr, 'solidFill') : null;
            return {
              text: textOf(rEl),
              font: latin ? latin.getAttribute('typeface') : null,
              sizePt: sz,
              color: solidFill ? resolveColor(solidFill, theme) : null,
              bold: rPr ? rPr.getAttribute('b') === '1' : false,
            };
          });

          paragraphs.push({
            index: paraIdx,
            hasBullet,
            marLIn: marL,
            indentIn: indent,
            lineSpacing: lnSpc,
            spaceBefore: spcBef,
            spaceAfter: spcAft,
            runs,
            text: runs.map((r) => r.text).join(''),
          });
        });
      }

      shapes.push({
        shapeName,
        shapeIndex: shapeIdx,
        fill: shapeFill,
        bodyPr,
        paragraphs,
        hasText: paragraphs.some((p) => p.text.trim().length > 0),
      });
    });

    return { slideNumber, shapes };
  }

  async function parsePptx(file) {
    const zip = await window.JSZip.loadAsync(file);
    const parser = new DOMParser();

    // Theme: map each slide -> its theme via slide layout -> master -> theme.
    // Simplification: PPTX decks overwhelmingly use one theme for all slides,
    // so we resolve the *first available* theme file and apply it globally.
    // (Multi-theme decks are rare; if present, only the first theme's colors
    // will be used for scheme-color resolution on later slides.)
    let theme = {};
    const themeFile = Object.keys(zip.files).find((f) => /^ppt\/theme\/theme\d+\.xml$/.test(f));
    if (themeFile) {
      const themeXml = await zip.files[themeFile].async('text');
      theme = parseTheme(parser.parseFromString(themeXml, 'application/xml'));
    }

    const slideFiles = Object.keys(zip.files)
      .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
      .sort((a, b) => {
        const na = Number(a.match(/slide(\d+)\.xml/)[1]);
        const nb = Number(b.match(/slide(\d+)\.xml/)[1]);
        return na - nb;
      });

    if (!slideFiles.length) {
      throw new Error('No slides found — is this a valid .pptx file?');
    }

    const slides = [];
    for (let i = 0; i < slideFiles.length; i++) {
      const xmlText = await zip.files[slideFiles[i]].async('text');
      const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
      slides.push(parseSlideXml(xmlDoc, theme, i + 1));
    }

    return { fileName: file.name, theme, slides };
  }

  global.FCParser = { parsePptx };
})(window);
