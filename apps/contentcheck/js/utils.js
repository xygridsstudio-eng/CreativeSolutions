/**
 * utils.js
 * Generic, dependency-free helper functions shared across the Content Check app:
 * ID generation, regex pattern libraries (numbers/dates/currency/%/email/url),
 * string similarity algorithms (Levenshtein, LCS), and small DOM/formatting helpers.
 *
 * No global variables are created; everything is exposed on window.CCUtils
 * so other modules (loaded as plain classic scripts) can consume it.
 */
(function (global) {
  'use strict';

  // ---------------------------------------------------------------------
  // ID generation
  // ---------------------------------------------------------------------
  let idCounter = 0;
  function nextId(prefix) {
    idCounter += 1;
    return `${prefix || 'id'}_${idCounter}_${Math.random().toString(36).slice(2, 7)}`;
  }
  function resetIdCounter() {
    idCounter = 0;
  }

  // ---------------------------------------------------------------------
  // Regex pattern library
  // ---------------------------------------------------------------------
  const PATTERNS = {
    // 5000, 5,000, 12.5, 12345.67
    number: /-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\d+(?:\.\d+)?/g,
    // 12%, 12.5 %
    percentage: /-?\d+(?:\.\d+)?\s?%/g,
    // $5,000  ₹1,000  Rs. 1,000  INR 500  USD 5,000.00  €100
    currency: /(?:[$₹€£¥]|Rs\.?|INR|USD|EUR|GBP)\s?-?\d{1,3}(?:,\d{2,3})*(?:\.\d+)?/gi,
    // FY25, FY2026, Q1FY26
    fiscalYear: /\bFY\s?-?\d{2,4}\b|\bQ[1-4]\s?FY\s?-?\d{2,4}\b/gi,
    // emails
    email: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g,
    // urls (with or without protocol/www)
    url: /(https?:\/\/[^\s,)"'<>]+)|(www\.[^\s,)"'<>]+\.[a-zA-Z]{2,}[^\s,)"'<>]*)/gi,
    // dates: 01-Jan-2026 | 1 January 2026 | 2026-01-01 | 15/01/2026 | Jan 1, 2026
    date: new RegExp(
      [
        '\\b\\d{1,2}[-/ ](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-/ ]\\d{2,4}\\b',
        '\\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[ .,]+\\d{1,2}[a-z]{0,2},?[ ]+\\d{2,4}\\b',
        '\\b\\d{4}-\\d{2}-\\d{2}\\b',
        '\\b\\d{1,2}/\\d{1,2}/\\d{2,4}\\b',
        '\\b\\d{1,2}-\\d{1,2}-\\d{2,4}\\b',
      ].join('|'),
      'gi'
    ),
    bulletPrefix: /^[\s]*[•▪◦‣∙·–—\-\*]\s+/,
    multiSpace: /[ \t]+/g,
    multiBlankLines: /\n{3,}/g,
    punctuation: /[.,;:!?"'“”‘’()\[\]{}]/g,
    sentenceSplit: /(?<=[.!?])\s+(?=[A-Z0-9"“'(])|\n+/g,
  };

  function extractMatches(text, regex) {
    if (!text) return [];
    const re = new RegExp(regex.source, regex.flags.includes('g') ? regex.flags : regex.flags + 'g');
    return [...text.matchAll(re)].map((m) => m[0].trim());
  }

  // ---------------------------------------------------------------------
  // String similarity
  // ---------------------------------------------------------------------

  /** Classic Levenshtein edit distance (iterative, O(n*m) time, O(min(n,m)) space). */
  function levenshtein(a, b) {
    if (a === b) return 0;
    const al = a.length;
    const bl = b.length;
    if (al === 0) return bl;
    if (bl === 0) return al;

    // Ensure b is the shorter string for reduced memory use.
    if (al < bl) return levenshtein(b, a);

    let prevRow = new Array(bl + 1);
    for (let j = 0; j <= bl; j++) prevRow[j] = j;

    for (let i = 1; i <= al; i++) {
      const curRow = new Array(bl + 1);
      curRow[0] = i;
      const ca = a.charCodeAt(i - 1);
      for (let j = 1; j <= bl; j++) {
        const cost = ca === b.charCodeAt(j - 1) ? 0 : 1;
        curRow[j] = Math.min(
          prevRow[j] + 1, // deletion
          curRow[j - 1] + 1, // insertion
          prevRow[j - 1] + cost // substitution
        );
      }
      prevRow = curRow;
    }
    return prevRow[bl];
  }

  /** Normalized similarity in [0,1] derived from Levenshtein distance. */
  function levenshteinSimilarity(a, b) {
    if (a === b) return 1;
    const maxLen = Math.max(a.length, b.length);
    if (maxLen === 0) return 1;
    return 1 - levenshtein(a, b) / maxLen;
  }

  /** Length of the Longest Common Subsequence between two token arrays. */
  function lcsLength(tokensA, tokensB) {
    const n = tokensA.length;
    const m = tokensB.length;
    if (n === 0 || m === 0) return 0;
    let prev = new Array(m + 1).fill(0);
    for (let i = 1; i <= n; i++) {
      const cur = new Array(m + 1).fill(0);
      for (let j = 1; j <= m; j++) {
        cur[j] = tokensA[i - 1] === tokensB[j - 1] ? prev[j - 1] + 1 : Math.max(prev[j], cur[j - 1]);
      }
      prev = cur;
    }
    return prev[m];
  }

  /** Word-level LCS-based similarity in [0,1]. Good for reordering / partial overlap. */
  function lcsSimilarity(a, b) {
    const ta = tokenize(a);
    const tb = tokenize(b);
    if (ta.length === 0 && tb.length === 0) return 1;
    if (ta.length === 0 || tb.length === 0) return 0;
    const lcs = lcsLength(ta, tb);
    return (2 * lcs) / (ta.length + tb.length);
  }

  function tokenize(text) {
    if (!text) return [];
    return text.split(/\s+/).filter(Boolean);
  }

  /** Text with every number replaced by a placeholder, to compare surrounding context only. */
  function numberSkeleton(text) {
    return text.replace(/-?\d+(?:,\d{3})*(?:\.\d+)?/g, '#');
  }

  /**
   * Combined similarity score blending character-level (Levenshtein) and
   * word-level (LCS) signals. Robust to both small edits and word reordering.
   */
  function combinedSimilarity(a, b) {
    if (a === b) return 1;
    const lev = levenshteinSimilarity(a, b);
    const lcs = lcsSimilarity(a, b);
    let score = lev * 0.6 + lcs * 0.4;

    // Two short strings that are otherwise identical except for their
    // numbers (e.g. "2.3%" vs "5.7%") are a value update, not unrelated
    // content — always a Modified pair, never Missing+Added. Plain
    // edit-distance scores these as unrelated because so few characters
    // are shared, so pairing is boosted whenever stripping the digits
    // makes both sides equal.
    const skelA = numberSkeleton(a);
    const skelB = numberSkeleton(b);
    if (skelA === skelB && skelA !== a) score = Math.max(score, 0.9);

    return score;
  }

  /** Word-level LCS DP table for two token arrays (full table, not space-optimized, so callers can backtrack an alignment). */
  function wordDiffTable(ta, tb) {
    const n = ta.length;
    const m = tb.length;
    const dp = new Array(n + 1);
    for (let i = 0; i <= n; i++) dp[i] = new Array(m + 1).fill(0);
    for (let i = 1; i <= n; i++) {
      for (let j = 1; j <= m; j++) {
        dp[i][j] = ta[i - 1] === tb[j - 1] ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
    return dp;
  }

  /**
   * Word-level diff between two strings, so a "modified" row can show just
   * the words that actually changed instead of two full, mostly-identical
   * sentences.
   * Returns { srcTokens, outTokens }, each an array of
   * { type: 'equal'|'removed'|'added', text } covering every word on that
   * side, in original order — for rendering the source and output as two
   * separate highlighted strings.
   */
  function diffWords(a, b) {
    const ta = tokenize(a);
    const tb = tokenize(b);
    const dp = wordDiffTable(ta, tb);

    const srcTokens = [];
    const outTokens = [];
    let i = ta.length;
    let j = tb.length;
    while (i > 0 && j > 0) {
      if (ta[i - 1] === tb[j - 1]) {
        srcTokens.push({ type: 'equal', text: ta[i - 1] });
        outTokens.push({ type: 'equal', text: tb[j - 1] });
        i -= 1;
        j -= 1;
      } else if (dp[i - 1][j] >= dp[i][j - 1]) {
        srcTokens.push({ type: 'removed', text: ta[i - 1] });
        i -= 1;
      } else {
        outTokens.push({ type: 'added', text: tb[j - 1] });
        j -= 1;
      }
    }
    while (i > 0) { srcTokens.push({ type: 'removed', text: ta[i - 1] }); i -= 1; }
    while (j > 0) { outTokens.push({ type: 'added', text: tb[j - 1] }); j -= 1; }

    srcTokens.reverse();
    outTokens.reverse();
    return { srcTokens, outTokens };
  }

  /**
   * Word-level diff as ONE inline sequence — Word "Track Changes"/redline
   * style: unchanged words plain, deleted words marked for strikethrough,
   * inserted words marked for underline, all in reading order, instead of
   * two separate before/after strings. Same alignment as diffWords, just
   * emitted as a single interleaved stream.
   * Returns an array of { type: 'equal'|'removed'|'added', text }.
   */
  function diffWordsMerged(a, b) {
    const ta = tokenize(a);
    const tb = tokenize(b);
    const dp = wordDiffTable(ta, tb);

    const tokens = [];
    let i = ta.length;
    let j = tb.length;
    while (i > 0 && j > 0) {
      if (ta[i - 1] === tb[j - 1]) {
        tokens.push({ type: 'equal', text: ta[i - 1] });
        i -= 1;
        j -= 1;
      } else if (dp[i - 1][j] >= dp[i][j - 1]) {
        tokens.push({ type: 'removed', text: ta[i - 1] });
        i -= 1;
      } else {
        tokens.push({ type: 'added', text: tb[j - 1] });
        j -= 1;
      }
    }
    while (i > 0) { tokens.push({ type: 'removed', text: ta[i - 1] }); i -= 1; }
    while (j > 0) { tokens.push({ type: 'added', text: tb[j - 1] }); j -= 1; }

    tokens.reverse();
    return tokens;
  }

  // ---------------------------------------------------------------------
  // Misc helpers
  // ---------------------------------------------------------------------
  function escapeHtml(str) {
    if (str == null) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatBytes(bytes) {
    if (!bytes) return '0 B';
    const units = ['B', 'KB', 'MB', 'GB'];
    const i = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
    return `${(bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
  }

  function formatDuration(ms) {
    if (ms < 1000) return `${Math.round(ms)} ms`;
    return `${(ms / 1000).toFixed(2)} s`;
  }

  function debounce(fn, wait) {
    let t;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), wait);
    };
  }

  function fileExtension(name) {
    const m = /\.([a-z0-9]+)$/i.exec(name || '');
    return m ? m[1].toLowerCase() : '';
  }

  function clamp(n, min, max) {
    return Math.max(min, Math.min(max, n));
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  function splitWords(text) {
    return (text || '').split(/\s+/).filter(Boolean);
  }

  /**
   * For a modified sentence pair, if either side is long, trim both the
   * source and output down to just the words around the actual change
   * (a few words of context before/after) instead of showing the full
   * sentence. Matches word-for-word from both ends to find the common
   * prefix/suffix, then keeps only `contextWords` words of that shared
   * prefix/suffix around the differing middle, with an ellipsis marking
   * anything cut. Short sentences are returned unchanged.
   */
  function trimAroundChange(srcText, outText, options) {
    const opts = Object.assign({ contextWords: 4, lengthThreshold: 20 }, options || {});
    const srcWords = splitWords(srcText);
    const outWords = splitWords(outText);

    if (srcWords.length <= opts.lengthThreshold && outWords.length <= opts.lengthThreshold) {
      return { source: srcText || '', output: outText || '', trimmed: false };
    }

    const maxPrefix = Math.min(srcWords.length, outWords.length);
    let prefix = 0;
    while (prefix < maxPrefix && srcWords[prefix] === outWords[prefix]) prefix++;

    const maxSuffix = Math.min(srcWords.length, outWords.length) - prefix;
    let suffix = 0;
    while (
      suffix < maxSuffix &&
      srcWords[srcWords.length - 1 - suffix] === outWords[outWords.length - 1 - suffix]
    ) suffix++;

    const srcStart = Math.max(0, prefix - opts.contextWords);
    const srcEnd = Math.min(srcWords.length, srcWords.length - suffix + opts.contextWords);
    const outStart = Math.max(0, prefix - opts.contextWords);
    const outEnd = Math.min(outWords.length, outWords.length - suffix + opts.contextWords);

    // Safety net: if there's no real differing middle to anchor on (e.g. a
    // change that's invisible at the word-split level), fall back to the
    // untrimmed text rather than risk an empty/nonsensical snippet.
    if (srcEnd <= srcStart || outEnd <= outStart) {
      return { source: srcText || '', output: outText || '', trimmed: false };
    }

    const buildSnippet = (words, start, end) => {
      const middle = words.slice(start, end).join(' ');
      const before = start > 0 ? '… ' : '';
      const after = end < words.length ? ' …' : '';
      return before + middle + after;
    };

    return {
      source: buildSnippet(srcWords, srcStart, srcEnd),
      output: buildSnippet(outWords, outStart, outEnd),
      trimmed: true,
    };
  }

  global.CCUtils = {
    nextId,
    resetIdCounter,
    PATTERNS,
    extractMatches,
    levenshtein,
    levenshteinSimilarity,
    lcsLength,
    lcsSimilarity,
    combinedSimilarity,
    diffWords,
    diffWordsMerged,
    tokenize,
    escapeHtml,
    formatBytes,
    formatDuration,
    debounce,
    fileExtension,
    clamp,
    downloadBlob,
    trimAroundChange,
  };
})(window);
