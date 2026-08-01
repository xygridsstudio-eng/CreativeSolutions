/**
 * checker.js
 * Single responsibility: compare a parsed PPTX model (parser.js) against a
 * guideline model (guidelines.js) and produce a flat, deduplicated list of
 * violations. No parsing or reporting logic lives here.
 *
 * Violations are deduplicated per (slide, shape, category, found-value) —
 * a text box with the same wrong font on 12 runs produces ONE row with an
 * occurrence count, not 12 near-identical rows.
 */
(function (global) {
  'use strict';

  // OOXML's own defaults when a bodyPr inset attribute is simply absent.
  const DEFAULT_INSETS_IN = { l: 0.1, t: 0.05, r: 0.1, b: 0.05 };

  function snippet(text, len) {
    const t = (text || '').trim().replace(/\s+/g, ' ');
    if (!t) return '(empty)';
    return t.length > len ? t.slice(0, len) + '…' : t;
  }

  function within(actual, target, tolerance) {
    return Math.abs(actual - target) <= tolerance + 1e-9;
  }

  function normHex(hex) {
    return (hex || '').toUpperCase();
  }

  function makeCollector() {
    const map = new Map();
    return {
      add(v) {
        const key = [v.slide, v.shapeName, v.category, v.found].join('|||');
        if (map.has(key)) {
          const existing = map.get(key);
          existing.count += 1;
          if (existing.examples.length < 3 && !existing.examples.includes(v.example)) {
            existing.examples.push(v.example);
          }
        } else {
          map.set(key, Object.assign({}, v, { count: 1, examples: [v.example] }));
        }
      },
      list() {
        return Array.from(map.values());
      },
    };
  }

  function checkDocument(model, guidelines) {
    const c = makeCollector();
    const checks = guidelines.checks || {};
    const allowedFonts = new Set((guidelines.fonts.allowed || []).map((f) => f.toLowerCase()));
    const allowedSizes = guidelines.fontSizes.allowed || [];
    const minBodyPt = guidelines.fontSizes.minBodyPt;
    const allowedColors = new Set((guidelines.colors.allowed || []).map(normHex));

    model.slides.forEach((slide) => {
      slide.shapes.forEach((shape) => {
        const shapeName = shape.shapeName;

        // ---- Shape internal margins ----
        if (checks.shapeMargins && shape.bodyPr && shape.hasText) {
          const bp = shape.bodyPr;
          const actual = {
            l: bp.lInsExplicit ? bp.lIns : DEFAULT_INSETS_IN.l,
            t: bp.tInsExplicit ? bp.tIns : DEFAULT_INSETS_IN.t,
            r: bp.rInsExplicit ? bp.rIns : DEFAULT_INSETS_IN.r,
            b: bp.bInsExplicit ? bp.bIns : DEFAULT_INSETS_IN.b,
          };
          const expected = guidelines.shapeMargins;
          const tol = expected.toleranceIn;
          const pairs = [
            ['left', actual.l, expected.leftIn], ['top', actual.t, expected.topIn],
            ['right', actual.r, expected.rightIn], ['bottom', actual.b, expected.bottomIn],
          ];
          pairs.forEach(([side, val, target]) => {
            if (val != null && !within(val, target, tol)) {
              c.add({
                slide: slide.slideNumber, shapeName, category: 'Shape Margin',
                found: `${side}: ${val}"`, expected: `${side}: ${target}" (±${tol}")`,
                example: shapeName,
              });
            }
          });
        }

        shape.paragraphs.forEach((para) => {
          if (!para.text.trim()) return;

          // ---- Line spacing ----
          if (checks.lineSpacing && para.lineSpacing) {
            const ls = para.lineSpacing;
            if (ls.type === 'pct') {
              if (!within(ls.value, guidelines.lineSpacing.targetPct, guidelines.lineSpacing.tolerancePct)) {
                c.add({
                  slide: slide.slideNumber, shapeName, category: 'Line Spacing',
                  found: `${ls.value}%`, expected: `${guidelines.lineSpacing.targetPct}% (±${guidelines.lineSpacing.tolerancePct}%)`,
                  example: snippet(para.text, 60),
                });
              }
            } else if (ls.type === 'pts') {
              c.add({
                slide: slide.slideNumber, shapeName, category: 'Line Spacing',
                found: `${ls.value}pt (fixed)`, expected: `${guidelines.lineSpacing.targetPct}% (relative)`,
                example: snippet(para.text, 60),
              });
            }
          }

          // ---- Paragraph (bullet) spacing ----
          if (checks.paragraphSpacing) {
            [['spaceBefore', para.spaceBefore, guidelines.paragraphSpacing.spaceBeforePt],
             ['spaceAfter', para.spaceAfter, guidelines.paragraphSpacing.spaceAfterPt]]
              .forEach(([label, sp, targetPt]) => {
                if (!sp) return;
                if (sp.type === 'pts') {
                  if (!within(sp.value, targetPt, guidelines.paragraphSpacing.tolerancePt)) {
                    c.add({
                      slide: slide.slideNumber, shapeName, category: 'Bullet/Paragraph Spacing',
                      found: `${label}: ${sp.value}pt`, expected: `${label}: ${targetPt}pt (±${guidelines.paragraphSpacing.tolerancePt}pt)`,
                      example: snippet(para.text, 60),
                    });
                  }
                } else if (sp.type === 'pct') {
                  c.add({
                    slide: slide.slideNumber, shapeName, category: 'Bullet/Paragraph Spacing',
                    found: `${label}: ${sp.value}% (relative)`, expected: `${label}: ${targetPt}pt (fixed)`,
                    example: snippet(para.text, 60),
                  });
                }
              });
          }

          // ---- Per-run checks: font, size, color ----
          para.runs.forEach((run) => {
            if (!run.text.trim()) return;

            if (checks.font && run.font) {
              if (!allowedFonts.has(run.font.toLowerCase())) {
                c.add({
                  slide: slide.slideNumber, shapeName, category: 'Font',
                  found: run.font, expected: (guidelines.fonts.allowed || []).join(' / '),
                  example: snippet(run.text, 50),
                });
              }
            }

            if (checks.fontSize && run.sizePt != null) {
              const matches = allowedSizes.some((s) => Math.abs(s - run.sizePt) < 0.01);
              if (!matches) {
                c.add({
                  slide: slide.slideNumber, shapeName, category: 'Font Size',
                  found: `${run.sizePt}pt`, expected: allowedSizes.map((s) => s + 'pt').join(' / '),
                  example: snippet(run.text, 50),
                });
              } else if (minBodyPt != null && run.sizePt < minBodyPt) {
                c.add({
                  slide: slide.slideNumber, shapeName, category: 'Font Size',
                  found: `${run.sizePt}pt (below minimum)`, expected: `≥ ${minBodyPt}pt`,
                  example: snippet(run.text, 50),
                });
              }
            }

            if (checks.textColor && run.color) {
              if (!allowedColors.has(normHex(run.color))) {
                c.add({
                  slide: slide.slideNumber, shapeName, category: 'Text Color',
                  found: run.color, expected: (guidelines.colors.allowed || []).join(' / '),
                  example: snippet(run.text, 50),
                });
              }
            }
          });
        });

        // ---- Shape fill color ----
        if (checks.shapeFillColor && shape.fill) {
          if (!allowedColors.has(normHex(shape.fill))) {
            c.add({
              slide: slide.slideNumber, shapeName, category: 'Shape Fill Color',
              found: shape.fill, expected: (guidelines.colors.allowed || []).join(' / '),
              example: shapeName,
            });
          }
        }
      });
    });

    const violations = c.list().sort((a, b) => a.slide - b.slide || a.category.localeCompare(b.category));

    const bySlide = new Map();
    violations.forEach((v) => bySlide.set(v.slide, (bySlide.get(v.slide) || 0) + 1));

    const byCategory = {};
    violations.forEach((v) => { byCategory[v.category] = (byCategory[v.category] || 0) + 1; });

    return {
      fileName: model.fileName,
      totalSlides: model.slides.length,
      slidesWithIssues: bySlide.size,
      totalViolations: violations.reduce((sum, v) => sum + v.count, 0),
      uniqueViolationRows: violations.length,
      byCategory,
      violations,
    };
  }

  global.FCChecker = { checkDocument };
})(window);
