/**
 * comparer.js
 * Single responsibility: deterministic, rule-based comparison of two
 * normalized Document hierarchies (see normalizer.js). No AI, no semantic
 * matching — only exact string comparison, Levenshtein distance, and LCS.
 *
 * Comparison happens at the SENTENCE level (most accurate), then results
 * are aggregated back up to paragraph, section, and document level.
 */
(function (global) {
  'use strict';

  const U = global.CCUtils;

  const MODIFIED_THRESHOLD = 0.45; // below this similarity, treat as unrelated (missing+added) not "modified"
  const SECTION_MATCH_THRESHOLD = 0.3;

  // ---------------------------------------------------------------------
  // Generic, position-independent sequence alignment.
  // Produces an ordered list of ops: match | modify | delete | insert
  //
  // Deliberately NOT order-preserving: the spec requires ignoring
  // formatting/LAYOUT differences, and moving a block of content to a
  // different position (e.g. reordering boxes/shapes on a slide) is a
  // layout change, not a content change. A naive order-preserving LCS
  // would report a moved-but-unchanged block as delete+insert
  // (Missing+Added) simply because it now appears out of sequence
  // relative to its neighbors. Instead:
  //   Phase 1 — pair up items with an IDENTICAL normalized key, regardless
  //             of position, so reordered-but-unchanged content is always
  //             a Match.
  //   Phase 2 — among what's left, greedily pair the highest-similarity
  //             src/out items ANYWHERE in the sequence (not just a local
  //             "gap"), so a block that was both edited and moved is still
  //             recognized as Modified rather than Missing+Added.
  //   Phase 3 — anything still unpaired is genuinely Missing or Added.
  // ---------------------------------------------------------------------
  function alignSequences(srcItems, outItems, keyFn, simFn, threshold) {
    const n = srcItems.length;
    const m = outItems.length;
    const srcKeys = srcItems.map(keyFn);
    const outKeys = outItems.map(keyFn);

    const usedSrc = new Set();
    const usedOut = new Set();
    const ops = [];

    // Phase 1: exact-key matches, position-independent.
    const outByKey = new Map();
    outKeys.forEach((k, j) => {
      if (!outByKey.has(k)) outByKey.set(k, []);
      outByKey.get(k).push(j);
    });
    srcKeys.forEach((k, i) => {
      const queue = outByKey.get(k);
      if (queue && queue.length) {
        const j = queue.shift();
        usedSrc.add(i);
        usedOut.add(j);
        ops.push({ type: 'match', srcIndex: i, outIndex: j, score: 1 });
      }
    });

    // Phase 2: near-match (modified) pairing among remaining items, scored
    // globally rather than within a positional window, so moved+edited
    // content still gets matched up.
    const remainingSrc = [];
    for (let i = 0; i < n; i++) if (!usedSrc.has(i)) remainingSrc.push(i);
    const remainingOut = [];
    for (let j = 0; j < m; j++) if (!usedOut.has(j)) remainingOut.push(j);

    const scored = [];
    // Guard against pathological O(n*m) blowups on very large documents:
    // beyond this many remaining items, restrict candidate pairs to a
    // proportional window around each item's relative position instead of
    // scoring every possible pair. Typical sections/paragraphs/sentences
    // are far smaller than this threshold, so normal documents are
    // unaffected and still get full, unrestricted (order-independent)
    // matching.
    const FULL_SCAN_LIMIT = 200000; // remainingSrc.length * remainingOut.length
    if (remainingSrc.length * remainingOut.length <= FULL_SCAN_LIMIT || remainingOut.length === 0) {
      remainingSrc.forEach((i) => {
        remainingOut.forEach((j) => {
          const score = simFn(srcItems[i], outItems[j]);
          if (score >= threshold) scored.push({ i, j, score });
        });
      });
    } else {
      const window = Math.max(25, Math.ceil(FULL_SCAN_LIMIT / Math.max(remainingSrc.length, remainingOut.length, 1)));
      remainingSrc.forEach((i, pos) => {
        const center = Math.round((pos / remainingSrc.length) * remainingOut.length);
        const lo = Math.max(0, center - window);
        const hi = Math.min(remainingOut.length - 1, center + window);
        for (let p = lo; p <= hi; p++) {
          const j = remainingOut[p];
          const score = simFn(srcItems[i], outItems[j]);
          if (score >= threshold) scored.push({ i, j, score });
        }
      });
    }

    scored.sort((a, b) => b.score - a.score);
    scored.forEach(({ i, j, score }) => {
      if (usedSrc.has(i) || usedOut.has(j)) return;
      usedSrc.add(i);
      usedOut.add(j);
      ops.push({ type: 'modify', srcIndex: i, outIndex: j, score });
    });

    // Phase 3: whatever's left is genuinely missing or added.
    for (let i = 0; i < n; i++) if (!usedSrc.has(i)) ops.push({ type: 'delete', srcIndex: i, outIndex: null });
    for (let j = 0; j < m; j++) if (!usedOut.has(j)) ops.push({ type: 'insert', srcIndex: null, outIndex: j });

    // Ordering is for readable report output only — it does not affect
    // match/modify/missing/added classification above.
    ops.sort((a, b) => {
      const av = a.srcIndex != null ? a.srcIndex : n + a.outIndex / (m + 1);
      const bv = b.srcIndex != null ? b.srcIndex : n + b.outIndex / (m + 1);
      return av - bv;
    });
    return ops;
  }

  // ---------------------------------------------------------------------
  // Special content detectors: returns { added: [...], removed: [...] }
  // ---------------------------------------------------------------------
  function diffPatternMatches(srcText, outText, regex) {
    const srcMatches = U.extractMatches(srcText, regex);
    const outMatches = U.extractMatches(outText, regex);
    const srcCount = countBy(srcMatches);
    const outCount = countBy(outMatches);
    const removed = [];
    const added = [];
    srcCount.forEach((count, val) => {
      const remaining = count - (outCount.get(val) || 0);
      for (let k = 0; k < remaining; k++) removed.push(val);
    });
    outCount.forEach((count, val) => {
      const remaining = count - (srcCount.get(val) || 0);
      for (let k = 0; k < remaining; k++) added.push(val);
    });
    return { added, removed };
  }

  function countBy(arr) {
    const m = new Map();
    arr.forEach((v) => m.set(v, (m.get(v) || 0) + 1));
    return m;
  }

  function detectSpecialChanges(srcText, outText, options) {
    const opts = Object.assign(
      { compareNumbers: true, compareDates: true, compareUrls: true, compareEmails: true },
      options || {}
    );
    const P = U.PATTERNS;
    const result = {};

    const percentage = opts.compareNumbers ? diffPatternMatches(srcText, outText, P.percentage) : null;
    const currency = opts.compareNumbers ? diffPatternMatches(srcText, outText, P.currency) : null;
    const fiscalYear = opts.compareNumbers ? diffPatternMatches(srcText, outText, P.fiscalYear) : null;
    const date = opts.compareDates ? diffPatternMatches(srcText, outText, P.date) : null;
    const email = opts.compareEmails ? diffPatternMatches(srcText, outText, P.email) : null;
    const url = opts.compareUrls ? diffPatternMatches(srcText, outText, P.url) : null;

    if (percentage && (percentage.added.length || percentage.removed.length)) result.percentage = percentage;
    if (currency && (currency.added.length || currency.removed.length)) result.currency = currency;
    if (fiscalYear && (fiscalYear.added.length || fiscalYear.removed.length)) result.fiscalYear = fiscalYear;
    if (date && (date.added.length || date.removed.length)) result.date = date;
    if (email && (email.added.length || email.removed.length)) result.email = email;
    if (url && (url.added.length || url.removed.length)) result.url = url;

    // Plain numbers: strip out substrings already claimed by %, currency,
    // fiscal-year, and date patterns *before* scanning for numbers, so a
    // date's day/month/year digits are never also reported as a separate
    // "number changed" entry.
    if (opts.compareNumbers) {
      const strip = (text) => [P.percentage, P.currency, P.fiscalYear, P.date].reduce(
        (t, re) => t.replace(new RegExp(re.source, re.flags.includes('g') ? re.flags : re.flags + 'g'), ' '),
        text
      );
      const number = diffPatternMatches(strip(srcText), strip(outText), P.number);
      if (number.added.length || number.removed.length) result.number = number;
    }
    return result;
  }

  function hasAnySpecialChange(special) {
    return Object.keys(special).length > 0;
  }

  // ---------------------------------------------------------------------
  // Sentence-level comparison within a matched/modified paragraph pair
  // ---------------------------------------------------------------------
  function compareSentences(srcPara, outPara, options) {
    const ops = alignSequences(
      srcPara.sentences,
      outPara.sentences,
      (s) => s.normalizedText,
      (a, b) => U.combinedSimilarity(a.normalizedText, b.normalizedText),
      MODIFIED_THRESHOLD
    );

    return ops.map((op) => {
      const srcSent = op.srcIndex != null ? srcPara.sentences[op.srcIndex] : null;
      const outSent = op.outIndex != null ? outPara.sentences[op.outIndex] : null;
      let status;
      let special = {};
      if (op.type === 'match') status = 'match';
      else if (op.type === 'delete') status = 'missing';
      else if (op.type === 'insert') status = 'added';
      else {
        special = detectSpecialChanges(srcSent.originalText, outSent.originalText, options);
        status = 'modified';
      }
      return {
        srcSentence: srcSent,
        outSentence: outSent,
        status,
        similarity: op.score != null ? op.score : (status === 'match' ? 1 : 0),
        special,
      };
    });
  }

  // ---------------------------------------------------------------------
  // Table comparison: align rows by tableId groups, then diff cell-by-cell
  // ---------------------------------------------------------------------
  function compareTables(srcParas, outParas) {
    const groupByTable = (paras) => {
      const groups = new Map();
      paras.forEach((p) => {
        if (p.type !== 'tableRow') return;
        const key = p.tableId || 'table';
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(p);
      });
      return groups;
    };

    const srcTables = groupByTable(srcParas);
    const outTables = groupByTable(outParas);
    const allTableIds = new Set([...srcTables.keys(), ...outTables.keys()]);
    const results = [];

    allTableIds.forEach((tableId) => {
      const srcRows = srcTables.get(tableId) || [];
      const outRows = outTables.get(tableId) || [];
      const ops = alignSequences(
        srcRows,
        outRows,
        (r) => r.normalizedText,
        (a, b) => U.combinedSimilarity(a.normalizedText, b.normalizedText),
        0.4
      );

      ops.forEach((op) => {
        const srcRow = op.srcIndex != null ? srcRows[op.srcIndex] : null;
        const outRow = op.outIndex != null ? outRows[op.outIndex] : null;
        if (op.type === 'match') {
          results.push({ tableId, status: 'match', srcRow, outRow, cellDiffs: [] });
        } else if (op.type === 'delete') {
          results.push({ tableId, status: 'missing', srcRow, outRow: null, cellDiffs: [] });
        } else if (op.type === 'insert') {
          results.push({ tableId, status: 'added', srcRow: null, outRow, cellDiffs: [] });
        } else {
          const cellDiffs = diffCells(srcRow.cells || [], outRow.cells || []);
          const modified = cellDiffs.some((c) => c.status !== 'match');
          results.push({ tableId, status: modified ? 'modified' : 'match', srcRow, outRow, cellDiffs });
        }
      });
    });

    return results;
  }

  function diffCells(srcCells, outCells) {
    const len = Math.max(srcCells.length, outCells.length);
    const diffs = [];
    for (let i = 0; i < len; i++) {
      const s = srcCells[i] != null ? srcCells[i] : '';
      const o = outCells[i] != null ? outCells[i] : '';
      let status;
      if (s === o) status = 'match';
      else if (!s) status = 'added';
      else if (!o) status = 'missing';
      else status = 'modified';
      diffs.push({ cellIndex: i, srcText: s, outText: o, status });
    }
    return diffs;
  }

  // ---------------------------------------------------------------------
  // Duplicate content detection (within a single document)
  // ---------------------------------------------------------------------
  function findDuplicates(docHierarchy) {
    const seen = new Map();
    const duplicates = [];
    docHierarchy.sections.forEach((section) => {
      section.paragraphs.forEach((p) => {
        if (p.type === 'tableRow' || !p.normalizedText || p.normalizedText.length < 8) return;
        if (seen.has(p.normalizedText)) {
          duplicates.push({ text: p.originalText, firstLocation: seen.get(p.normalizedText), location: describeLocation(section, p) });
        } else {
          seen.set(p.normalizedText, describeLocation(section, p));
        }
      });
    });
    return duplicates;
  }

  function describeLocation(section, paragraph) {
    return {
      section: section.title,
      page: section.page,
      slide: section.slide,
      paragraphNumber: paragraph.paragraphNumber,
    };
  }

  // ---------------------------------------------------------------------
  // Top-level document comparison
  // ---------------------------------------------------------------------
  function compareDocuments(srcDoc, outDoc, options) {
    const sectionOps = alignSequences(
      srcDoc.sections,
      outDoc.sections,
      (s) => (s.title || '').trim().toLowerCase(),
      (a, b) => U.combinedSimilarity((a.title || '').toLowerCase(), (b.title || '').toLowerCase()),
      SECTION_MATCH_THRESHOLD
    );

    const detailRows = [];
    const tableRows = [];
    let blockId = 0;

    const stats = {
      totalSections: 0,
      totalParagraphs: 0,
      totalSentences: 0,
      matched: 0,
      modified: 0,
      missing: 0,
      added: 0,
    };

    function pushDetail(status, section, srcPara, outPara, srcSent, outSent, special, comments) {
      blockId += 1;
      detailRows.push({
        blockId,
        section: section ? section.title : '',
        page: section ? section.page : null,
        slide: section ? section.slide : null,
        paragraphNumber: srcPara ? srcPara.paragraphNumber : (outPara ? outPara.paragraphNumber : null),
        sentenceNumber: srcSent ? srcSent.sentenceNumber : (outSent ? outSent.sentenceNumber : null),
        status,
        sourceText: srcSent ? srcSent.originalText : (srcPara ? srcPara.originalText : ''),
        outputText: outSent ? outSent.originalText : (outPara ? outPara.originalText : ''),
        special: special || {},
        comments: comments || '',
      });
      stats.totalSentences += 1;
      if (status === 'match') stats.matched += 1;
      else if (status === 'modified') stats.modified += 1;
      else if (status === 'missing') stats.missing += 1;
      else if (status === 'added') stats.added += 1;
    }

    sectionOps.forEach((sOp) => {
      const srcSection = sOp.srcIndex != null ? srcDoc.sections[sOp.srcIndex] : null;
      const outSection = sOp.outIndex != null ? outDoc.sections[sOp.outIndex] : null;
      stats.totalSections += 1;

      const activeSection = srcSection || outSection;

      if (sOp.type === 'delete') {
        srcSection.paragraphs.forEach((p) => {
          stats.totalParagraphs += 1;
          if (p.type === 'tableRow') return; // handled in table pass below
          p.sentences.forEach((s) => pushDetail('missing', srcSection, p, null, s, null, {}, 'Entire section removed'));
        });
        return;
      }
      if (sOp.type === 'insert') {
        outSection.paragraphs.forEach((p) => {
          stats.totalParagraphs += 1;
          if (p.type === 'tableRow') return;
          p.sentences.forEach((s) => pushDetail('added', outSection, null, p, null, s, {}, 'Entire section added'));
        });
        return;
      }

      // match or modify at section level: align paragraphs within.
      const srcParas = srcSection.paragraphs.filter((p) => p.type !== 'tableRow');
      const outParas = outSection.paragraphs.filter((p) => p.type !== 'tableRow');

      const paraOps = alignSequences(
        srcParas,
        outParas,
        (p) => p.normalizedText,
        (a, b) => U.combinedSimilarity(a.normalizedText, b.normalizedText),
        MODIFIED_THRESHOLD
      );

      paraOps.forEach((pOp) => {
        const srcPara = pOp.srcIndex != null ? srcParas[pOp.srcIndex] : null;
        const outPara = pOp.outIndex != null ? outParas[pOp.outIndex] : null;
        stats.totalParagraphs += 1;

        if (pOp.type === 'match') {
          srcPara.sentences.forEach((s, idx) => {
            pushDetail('match', activeSection, srcPara, outPara, s, outPara.sentences[idx] || s, {}, '');
          });
        } else if (pOp.type === 'delete') {
          srcPara.sentences.forEach((s) => pushDetail('missing', activeSection, srcPara, null, s, null, {}, ''));
        } else if (pOp.type === 'insert') {
          outPara.sentences.forEach((s) => pushDetail('added', activeSection, null, outPara, null, s, {}, ''));
        } else {
          const sentenceDiffs = compareSentences(srcPara, outPara, options);
          sentenceDiffs.forEach((sd) => {
            const comment = hasAnySpecialChange(sd.special) ? summarizeSpecial(sd.special) : '';
            pushDetail(sd.status, activeSection, srcPara, outPara, sd.srcSentence, sd.outSentence, sd.special, comment);
          });
        }
      });

      // Tables within this section.
      const srcTableParas = srcSection.paragraphs.filter((p) => p.type === 'tableRow');
      const outTableParas = outSection.paragraphs.filter((p) => p.type === 'tableRow');
      if (srcTableParas.length || outTableParas.length) {
        const tDiffs = compareTables(srcTableParas, outTableParas);
        tDiffs.forEach((t) => {
          blockId += 1;
          tableRows.push({
            blockId,
            section: activeSection.title,
            page: activeSection.page,
            slide: activeSection.slide,
            tableId: t.tableId,
            status: t.status,
            sourceText: t.srcRow ? t.srcRow.text : '',
            outputText: t.outRow ? t.outRow.text : '',
            cellDiffs: t.cellDiffs,
          });
          if (t.status === 'match') stats.matched += 1;
          else if (t.status === 'modified') stats.modified += 1;
          else if (t.status === 'missing') stats.missing += 1;
          else if (t.status === 'added') stats.added += 1;
        });
      }
    });

    const duplicatesSrc = findDuplicates(srcDoc);
    const duplicatesOut = findDuplicates(outDoc);

    const matchPercentage = stats.totalSentences + tableRows.length > 0
      ? Math.round((stats.matched / (stats.totalSentences + (tableRows.length ? 0 : 0) || 1)) * 1000) / 10
      : 100;

    return {
      stats: Object.assign(stats, {
        matchPercentage: computeMatchPercentage(stats),
      }),
      detailRows,
      tableRows,
      duplicates: { source: duplicatesSrc, output: duplicatesOut },
    };
  }

  function computeMatchPercentage(stats) {
    const total = stats.matched + stats.modified + stats.missing + stats.added;
    if (total === 0) return 100;
    return Math.round((stats.matched / total) * 1000) / 10;
  }

  function summarizeSpecial(special) {
    const parts = [];
    Object.keys(special).forEach((key) => {
      const { added, removed } = special[key];
      const label = key.charAt(0).toUpperCase() + key.slice(1);
      if (removed.length) parts.push(`${label} removed: ${removed.join(', ')}`);
      if (added.length) parts.push(`${label} added: ${added.join(', ')}`);
    });
    return parts.join(' | ');
  }

  global.CCComparer = {
    alignSequences,
    compareDocuments,
    detectSpecialChanges,
    findDuplicates,
  };
})(window);
