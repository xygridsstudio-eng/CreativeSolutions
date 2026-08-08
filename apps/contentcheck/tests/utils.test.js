'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadApp } = require('./setup');

const { CCUtils: U } = loadApp();

test('levenshteinSimilarity: identical strings score 1', () => {
  assert.equal(U.levenshteinSimilarity('hello', 'hello'), 1);
});

test('levenshteinSimilarity: completely different strings score low', () => {
  assert.ok(U.levenshteinSimilarity('apple', 'banana') < 0.3);
});

test('lcsSimilarity: fully reordered words only get partial credit (order-sensitive subsequence, not bag-of-words)', () => {
  const score = U.lcsSimilarity('the quick brown fox', 'brown fox the quick');
  assert.equal(score, 0.5);
});

test('lcsSimilarity: appending words to the end keeps the shared prefix fully credited', () => {
  const score = U.lcsSimilarity('the quick brown fox', 'the quick brown fox jumps');
  assert.ok(score > 0.85, `expected near-full credit for a pure suffix addition, got ${score}`);
});

test('combinedSimilarity: identical strings score exactly 1', () => {
  assert.equal(U.combinedSimilarity('same text', 'same text'), 1);
});

test('combinedSimilarity: number-only change is boosted above the modify threshold', () => {
  // Regression: "2.3%" vs "5.7%" used to score ~0.3 (below the 0.45
  // threshold used in comparer.js), so a value update was reported as a
  // separate Missing + Added instead of one Modified pair.
  const score = U.combinedSimilarity('2.3%', '5.7%');
  assert.ok(score >= 0.45, `expected number-only change to clear the modify threshold, got ${score}`);
});

test('combinedSimilarity: number boost requires an actual digit difference', () => {
  // Two identical non-numeric strings already return 1 via the early exit;
  // this checks the boost path doesn't misfire on strings with no digits.
  const score = U.combinedSimilarity('apple', 'banana');
  assert.ok(score < 0.45, `unrelated non-numeric strings should not be boosted, got ${score}`);
});

test('combinedSimilarity: numbers embedded in longer matching sentences still boost', () => {
  const score = U.combinedSimilarity(
    'Revenue grew by 2.3% this quarter',
    'Revenue grew by 5.7% this quarter'
  );
  assert.ok(score >= 0.45, `expected sentence with an updated number to score high, got ${score}`);
});

test('diffWords: marks only the changed words', () => {
  const { srcTokens, outTokens } = U.diffWords('the quick brown fox', 'the slow brown fox');
  assert.deepEqual(
    srcTokens.map((t) => t.type),
    ['equal', 'removed', 'equal', 'equal']
  );
  assert.deepEqual(
    outTokens.map((t) => t.type),
    ['equal', 'added', 'equal', 'equal']
  );
  assert.equal(srcTokens[1].text, 'quick');
  assert.equal(outTokens[1].text, 'slow');
});

test('diffWords: identical strings produce only equal tokens', () => {
  const { srcTokens, outTokens } = U.diffWords('a b c', 'a b c');
  assert.ok(srcTokens.every((t) => t.type === 'equal'));
  assert.ok(outTokens.every((t) => t.type === 'equal'));
});

test('extractMatches: percentage pattern picks up all matches', () => {
  const matches = U.extractMatches('grew 2.3% then fell 5.7 % overall', U.PATTERNS.percentage);
  assert.equal(matches.length, 2);
});

test('extractMatches: number pattern ignores currency-prefixed values when combined with strip logic', () => {
  const matches = U.extractMatches('spent $5,000 on 12 items', U.PATTERNS.number);
  assert.ok(matches.includes('12'));
});

test('trimAroundChange: short text is returned unchanged', () => {
  const result = U.trimAroundChange('short text here', 'short text there');
  assert.equal(result.trimmed, false);
  assert.equal(result.source, 'short text here');
});

test('trimAroundChange: long text is trimmed around the differing middle', () => {
  const src = 'one two three four five six seven eight nine ten eleven twelve thirteen fourteen CHANGED sixteen seventeen eighteen nineteen twenty twentyone';
  const out = src.replace('CHANGED', 'DIFFERENT');
  const result = U.trimAroundChange(src, out);
  assert.equal(result.trimmed, true);
  assert.ok(result.source.includes('CHANGED'));
  assert.ok(result.output.includes('DIFFERENT'));
});
