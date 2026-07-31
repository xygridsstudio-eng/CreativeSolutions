/**
 * normalizer.js
 * Single responsibility: turn the raw parser output into the normalized
 * Document > Section > Paragraph > Sentence hierarchy that the comparison
 * engine operates on. Never changes the *words* of the original text —
 * only whitespace/case/punctuation/bullet handling, as controlled by options.
 */
(function (global) {
  'use strict';

  const U = global.CCUtils;

  const DEFAULT_OPTIONS = {
    ignoreCase: true,
    ignoreExtraSpaces: true,
    ignoreBlankLines: true,
    ignorePunctuation: false,
    ignoreBulletSymbols: true,
    ignoreHeaders: false,
    ignoreFooters: false,
  };

  /** Normalize a single line of text according to the active options. */
  function normalizeText(text, options) {
    if (!text) return '';
    let t = text;

    // Standardize quotes and dashes so purely typographic differences never
    // register as content changes, regardless of option toggles.
    t = t
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201C\u201D]/g, '"')
      .replace(/[\u2013\u2014]/g, '-')
      .replace(/\u00A0/g, ' ');

    if (options.ignoreBulletSymbols) {
      t = t.replace(U.PATTERNS.bulletPrefix, '');
    }
    if (options.ignoreExtraSpaces) {
      t = t.replace(U.PATTERNS.multiSpace, ' ').trim();
    }
    if (options.ignorePunctuation) {
      t = t.replace(U.PATTERNS.punctuation, '');
      t = t.replace(U.PATTERNS.multiSpace, ' ').trim();
    }
    if (options.ignoreCase) {
      t = t.toLowerCase();
    }
    return t.trim();
  }

  /** Split a paragraph's original text into sentence strings (on original text, not normalized). */
  function splitSentences(text) {
    if (!text) return [];
    const parts = text.split(U.PATTERNS.sentenceSplit).map((s) => s.trim()).filter(Boolean);
    return parts.length ? parts : [text.trim()];
  }

  function looksLikeHeaderFooter(block, blockIndexInSection, totalInSection) {
    // Heuristic: very short first/last block of a section, no sentence punctuation,
    // e.g. running headers like "Confidential" or footers like "Page 3 of 40".
    const shortText = block.text.length <= 60;
    const isEdge = blockIndexInSection === 0 || blockIndexInSection === totalInSection - 1;
    const noPunctuation = !/[.!?]$/.test(block.text.trim());
    return shortText && isEdge && noPunctuation;
  }

  /**
   * Build the full hierarchy for one parsed document.
   * @param {object} raw - output of CCParser.parseFile
   * @param {object} options - comparison options (ignoreCase, etc.)
   * @returns {object} normalized Document object
   */
  function buildHierarchy(raw, options) {
    const opts = Object.assign({}, DEFAULT_OPTIONS, options || {});
    const document = {
      docType: raw.docType,
      fileName: raw.fileName,
      sections: [],
    };

    raw.sections.forEach((rawSection, sIdx) => {
      const section = {
        id: U.nextId('sec'),
        index: sIdx,
        title: rawSection.title,
        page: rawSection.page,
        slide: rawSection.slide,
        paragraphs: [],
      };

      const totalBlocks = rawSection.blocks.length;
      let paragraphCounter = 0;

      rawSection.blocks.forEach((block, bIdx) => {
        if (opts.ignoreBlankLines && !block.text.trim()) return;

        if (opts.ignoreHeaders && block.type === 'heading') return;
        if ((opts.ignoreHeaders || opts.ignoreFooters) && block.type !== 'tableRow') {
          if (looksLikeHeaderFooter(block, bIdx, totalBlocks)) {
            if (opts.ignoreHeaders && bIdx === 0) return;
            if (opts.ignoreFooters && bIdx === totalBlocks - 1) return;
          }
        }

        paragraphCounter += 1;
        const normalizedText = normalizeText(block.text, opts);
        if (opts.ignoreBlankLines && !normalizedText) return;

        const paragraph = {
          id: U.nextId('para'),
          paragraphNumber: paragraphCounter,
          type: block.type,
          originalText: block.text,
          normalizedText,
          cells: block.cells || null,
          tableId: block.tableId || null,
          rowIndex: block.rowIndex != null ? block.rowIndex : null,
          sentences: [],
        };

        if (block.type === 'tableRow') {
          // Table rows are compared cell-by-cell in comparer.js; still provide
          // a single "sentence" so generic sentence-level comparison also works.
          paragraph.sentences.push({
            id: U.nextId('sent'),
            sentenceNumber: 1,
            originalText: block.text,
            normalizedText,
          });
        } else {
          const rawSentences = splitSentences(block.text);
          rawSentences.forEach((sentText, i) => {
            paragraph.sentences.push({
              id: U.nextId('sent'),
              sentenceNumber: i + 1,
              originalText: sentText,
              normalizedText: normalizeText(sentText, opts),
            });
          });
        }

        section.paragraphs.push(paragraph);
      });

      if (section.paragraphs.length) document.sections.push(section);
    });

    return document;
  }

  global.CCNormalizer = {
    DEFAULT_OPTIONS,
    normalizeText,
    splitSentences,
    buildHierarchy,
  };
})(window);
