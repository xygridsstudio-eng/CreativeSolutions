'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadApp } = require('./setup');

const { CCComparer: C, CCNormalizer: N, CCUtils: U } = loadApp();

function rawDoc(blocks) {
  return {
    docType: 'txt',
    fileName: 'test.txt',
    sections: [{ title: 'Section', page: null, slide: null, blocks }],
  };
}

function compare(srcBlocks, outBlocks, options) {
  const srcDoc = N.buildHierarchy(rawDoc(srcBlocks), N.DEFAULT_OPTIONS);
  const outDoc = N.buildHierarchy(rawDoc(outBlocks), N.DEFAULT_OPTIONS);
  return C.compareDocuments(srcDoc, outDoc, options || {});
}

function nonMatchRows(comparison) {
  return comparison.detailRows
    .filter((r) => r.status !== 'match')
    .map((r) => ({ status: r.status, src: r.sourceText, out: r.outputText }));
}

// ---------------------------------------------------------------------
// alignSequences: the core position-independent matching primitive
// ---------------------------------------------------------------------

test('alignSequences: phase 1 pairs identical items regardless of order', () => {
  const ops = C.alignSequences(
    ['a', 'b', 'c'],
    ['c', 'a', 'b'],
    (s) => s,
    (a, b) => U.combinedSimilarity(a, b),
    0.45
  );
  assert.ok(ops.every((op) => op.type === 'match'));
});

test('alignSequences: phase 2 pairs similar-but-edited items as modify', () => {
  const ops = C.alignSequences(
    ['The quarterly report is ready for review'],
    ['The quarterly report is ready for approval'],
    (s) => s,
    (a, b) => U.combinedSimilarity(a, b),
    0.45
  );
  assert.equal(ops.length, 1);
  assert.equal(ops[0].type, 'modify');
});

test('alignSequences: below-threshold, unrelated items stay Missing + Added by default', () => {
  const ops = C.alignSequences(
    ['Raw material price optimization (All Extrudes)'],
    ['Bag fill improvement unlocks (for Soft Extrudes)'],
    (s) => s,
    (a, b) => U.combinedSimilarity(a, b),
    0.45
  );
  const types = ops.map((op) => op.type).sort();
  assert.deepEqual(types, ['delete', 'insert']);
});

test('alignSequences: positionalFallback pairs wholesale replacements as modify', () => {
  // Same content as the test above, but with positionalFallback on — this
  // is the exact "copy-pasted a different banner's text over a label"
  // scenario: too little shared text for similarity scoring, but still the
  // same slot being edited.
  const ops = C.alignSequences(
    ['Context', 'Raw material price optimization (All Extrudes)', 'Leftover A'],
    ['Acquisition strategy', 'Bag fill improvement unlocks (for Soft Extrudes)'],
    (s) => s,
    (a, b) => U.combinedSimilarity(a, b),
    0.45,
    { positionalFallback: true }
  );
  const modified = ops.filter((op) => op.type === 'modify');
  assert.equal(modified.length, 2);
  assert.equal(ops.find((op) => op.type === 'delete').srcIndex, 2); // "Leftover A" has no counterpart left
});

// ---------------------------------------------------------------------
// detectSpecialChanges
// ---------------------------------------------------------------------

test('detectSpecialChanges: flags a changed percentage', () => {
  const special = C.detectSpecialChanges('Growth was 2.3% last year', 'Growth was 5.7% last year');
  assert.deepEqual(special.percentage.removed, ['2.3%']);
  assert.deepEqual(special.percentage.added, ['5.7%']);
});

test('detectSpecialChanges: date digits are not double-reported as plain numbers', () => {
  const special = C.detectSpecialChanges('Due 15/01/2026', 'Due 20/02/2026');
  assert.ok(special.date);
  assert.equal(special.number, undefined);
});

// ---------------------------------------------------------------------
// compareDocuments: end-to-end regressions for bugs fixed this session
// ---------------------------------------------------------------------

test('regression: a number-only change inside a sentence is one Modified row, not Missing+Added', () => {
  const comparison = compare(
    [{ type: 'paragraph', text: 'South Africa 5.7%, Pakistan 2.3%, India 3.1%.' }],
    [{ type: 'paragraph', text: 'South Africa 5.7%, Pakistan 5.7%, India 5.7%.' }]
  );
  const rows = nonMatchRows(comparison);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].status, 'modified');
});

test('regression: a table cell numeric change is Modified with both values shown', () => {
  const comparison = compare(
    [{ type: 'tableRow', text: 'Pakistan | 2.3%', cells: ['Pakistan', '2.3%'], tableId: 't1', rowIndex: 0 }],
    [{ type: 'tableRow', text: 'Pakistan | 5.7%', cells: ['Pakistan', '5.7%'], tableId: 't1', rowIndex: 0 }]
  );
  assert.equal(comparison.tableRows.length, 1);
  const row = comparison.tableRows[0];
  assert.equal(row.status, 'modified');
  const cell = row.cellDiffs.find((c) => c.cellIndex === 1);
  assert.equal(cell.status, 'modified');
  assert.equal(cell.srcText, '2.3%');
  assert.equal(cell.outText, '5.7%');
});

test('regression: a heading wholesale-replaced with unrelated text is Modified, not Missing+Added', () => {
  const comparison = compare(
    [
      { type: 'heading', text: 'Context' },
      { type: 'heading', text: 'Raw material price optimization (All Extrudes)' },
    ],
    [
      { type: 'heading', text: 'Acquisition strategy' },
      { type: 'heading', text: 'Bag fill improvement unlocks (for Soft Extrudes)' },
    ]
  );
  const rows = nonMatchRows(comparison);
  assert.equal(rows.length, 2);
  assert.ok(rows.every((r) => r.status === 'modified'));
});

test('a paragraph genuinely deleted with no replacement is still Missing', () => {
  const comparison = compare(
    [
      { type: 'paragraph', text: 'This stays the same content here.' },
      { type: 'paragraph', text: 'This paragraph gets removed entirely with nothing replacing it.' },
    ],
    [{ type: 'paragraph', text: 'This stays the same content here.' }]
  );
  const rows = nonMatchRows(comparison);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].status, 'missing');
});

test('a paragraph genuinely added with no prior counterpart is still Added', () => {
  const comparison = compare(
    [{ type: 'paragraph', text: 'This stays the same content here.' }],
    [
      { type: 'paragraph', text: 'This stays the same content here.' },
      { type: 'paragraph', text: 'This is brand new paragraph content that did not exist before.' },
    ]
  );
  const rows = nonMatchRows(comparison);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].status, 'added');
});

test('identical documents produce zero differences', () => {
  const blocks = [
    { type: 'heading', text: 'Title' },
    { type: 'paragraph', text: 'Body text that does not change.' },
  ];
  const comparison = compare(blocks, blocks.map((b) => Object.assign({}, b)));
  assert.equal(nonMatchRows(comparison).length, 0);
  assert.equal(comparison.stats.matchPercentage, 100);
});
