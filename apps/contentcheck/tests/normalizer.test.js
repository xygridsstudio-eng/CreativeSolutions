'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadApp } = require('./setup');

const { CCNormalizer: N } = loadApp();

function rawDoc(blocks) {
  return {
    docType: 'txt',
    fileName: 'test.txt',
    sections: [{ title: 'Section', page: null, slide: null, blocks }],
  };
}

test('normalizeText: collapses extra whitespace and lowercases', () => {
  const out = N.normalizeText('  Hello   World  ', N.DEFAULT_OPTIONS);
  assert.equal(out, 'hello world');
});

test('normalizeText: strips a leading bullet symbol', () => {
  const out = N.normalizeText('• Some bullet point', N.DEFAULT_OPTIONS);
  assert.equal(out, 'some bullet point');
});

test('normalizeText: smart quotes and dashes normalize regardless of options', () => {
  const out = N.normalizeText('“Hello” — test', N.DEFAULT_OPTIONS);
  assert.equal(out, '"hello" - test');
});

test('splitSentences: splits on sentence-ending punctuation', () => {
  const parts = N.splitSentences('First sentence. Second sentence! Third one?');
  assert.equal(parts.length, 3);
});

test('splitSentences: a single sentence with no terminator stays one part', () => {
  const parts = N.splitSentences('Just one line of text');
  assert.deepEqual(parts, ['Just one line of text']);
});

test('buildHierarchy: assigns sequential paragraph and sentence numbers', () => {
  const doc = N.buildHierarchy(
    rawDoc([
      { type: 'paragraph', text: 'First para. Second sentence.' },
      { type: 'paragraph', text: 'Another paragraph.' },
    ]),
    N.DEFAULT_OPTIONS
  );
  const [p1, p2] = doc.sections[0].paragraphs;
  assert.equal(p1.paragraphNumber, 1);
  assert.equal(p1.sentences.length, 2);
  assert.equal(p2.paragraphNumber, 2);
  assert.equal(p2.sentences.length, 1);
});

test('buildHierarchy: drops blank blocks when ignoreBlankLines is set', () => {
  const doc = N.buildHierarchy(
    rawDoc([
      { type: 'paragraph', text: '   ' },
      { type: 'paragraph', text: 'Real content.' },
    ]),
    N.DEFAULT_OPTIONS
  );
  assert.equal(doc.sections[0].paragraphs.length, 1);
  assert.equal(doc.sections[0].paragraphs[0].originalText, 'Real content.');
});

test('buildHierarchy: a tableRow block keeps its cells and gets one whole-row sentence', () => {
  const doc = N.buildHierarchy(
    rawDoc([
      { type: 'tableRow', text: 'A | B | C', cells: ['A', 'B', 'C'], tableId: 't1', rowIndex: 0 },
    ]),
    N.DEFAULT_OPTIONS
  );
  const row = doc.sections[0].paragraphs[0];
  assert.deepEqual(row.cells, ['A', 'B', 'C']);
  assert.equal(row.sentences.length, 1);
  assert.equal(row.sentences[0].originalText, 'A | B | C');
});
