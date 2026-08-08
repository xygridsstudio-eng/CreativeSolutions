# Content Check Tool

A lightweight, standalone, browser-only application that compares a **Source**
document against an **Output** document, ignores formatting/layout
differences, and produces a professional multi-sheet Excel report of the
content differences.

No installation, no server, no build step, no AI. Pure HTML + CSS +
vanilla ES6 JavaScript, running 100% in your browser.

---

## Quick Start

1. Extract the ZIP anywhere on your computer.
2. Open `index.html` directly in a modern browser (Chrome, Edge, or Firefox).
3. Upload a **Source** file and an **Output** file (DOCX, PPTX, PDF, or TXT) —
   either by browsing/dragging a file, or by pasting a direct URL into the
   "or paste a file URL / SharePoint link…" box under each upload area and
   clicking **Fetch**.
4. Adjust the comparison options if needed, then click **Compare**.
5. Review the on-screen summary and preview, then click
   **Download Excel Report** to save the full `.xlsx` report.

### Loading a file from a URL, shared drive, or SharePoint

Each upload box has a "Fetch from URL" field underneath it. This is a
genuine convenience for files that are directly, publicly reachable —
e.g. a file hosted on an internal web server, or a SharePoint "anyone with
the link can download" sharing URL.

**Important limitation:** this is a static, no-backend app that makes no
server-side requests — everything happens in your browser. That means it
is bound by the same security rules every website is:

- There is **no way for a browser page to read a file by an arbitrary
  local/network path string** (e.g. `\\server\share\file.docx` or a mapped
  drive letter) without you picking it via the file dialog — this is an
  intentional browser security restriction, not a limitation of this tool,
  and no web app can bypass it.
- A **URL** (`https://...`) can be fetched directly, but most authenticated
  corporate SharePoint/Teams links block this via CORS — the browser will
  refuse to hand the response back to this page even though the request
  reached the server. If a Fetch fails, the error message will say so;
  simply download the file from that link normally and use the upload box
  instead.
- The last URL you fetch for each side is remembered (via `localStorage`)
  so you don't have to retype it next time.

### If the browser blocks local file access

Some browsers restrict `fetch`/module loading from `file://` URLs. If the
libraries in `lib/` don't load (blank page, console errors about CORS),
serve the folder instead of opening it directly:

```bash
# From inside the ContentCheck folder:
python3 -m http.server 8000
# then open http://localhost:8000 in your browser
```

Any static file server works — Node's `npx serve`, VS Code's "Live Server"
extension, etc. This app makes **no network requests**; a local server is
only needed to satisfy the browser's local-file security policy.

### Try it with the included samples

The `samples/` folder has a ready-made Source/Output pair for every
supported format (`source.docx`/`output.docx`, `source.pptx`/`output.pptx`,
`source.pdf`/`output.pdf`, `source.txt`/`output.txt`). Each pair has a
handful of deliberate changes — a percentage change, a date change, a
currency change, an email domain change, a missing sentence, an added
sentence, and a modified table cell — so you can see every difference type
represented in one comparison.

---

## Design Settings

Click **🎨 Design Settings** (collapsed by default) to customize:

- **Header Font** / **Body Font** — a curated list of system-safe fonts
  plus a few nice Google Fonts (Playfair Display, Roboto, Open Sans) loaded
  on demand only if you pick one.
- **Font Color**, **Page Background**, **Section Banner Background**,
  and **Primary Accent** (buttons/links).
- **KPI & Chart Colors** — separate colors for Matched / Modified /
  Missing / Added, used in the summary stat cards and the preview table's
  row highlighting. Each color's lighter background tint (e.g. a status
  card's pale fill) is derived automatically, so you only ever pick one
  color per status.

Changes apply instantly (no page reload) and are remembered in this
browser via `localStorage` for next time. **Reset to Defaults** restores
the original look and clears the saved settings. If a page is opened with
`localStorage` unavailable (e.g. some strict private-browsing modes),
theming still works for that session — it just won't persist.

---

## Project Structure

```
ContentCheck/
  index.html            Application shell — loads libraries & modules in order
  css/
    styles.css           All styling
  js/
    app.js                Orchestrates the pipeline; wires ui.js to everything else
    ui.js                 All DOM reads/writes (uploads, options, progress, summary, preview)
    parser.js             File -> raw structural content (docx/pptx/pdf/txt)
    normalizer.js         Raw content -> normalized Document>Section>Paragraph>Sentence hierarchy
    comparer.js            Deterministic sentence-level comparison engine (LCS + Levenshtein)
    report.js             Comparison result -> report model consumed by the UI and Excel export
    excel.js               Report model -> styled 3-sheet .xlsx workbook (via ExcelJS)
    theme.js                Design Settings state: defaults, CSS variable application, color math, persistence
    utils.js               Shared helpers: IDs, regex pattern library, string similarity, misc.
  lib/
    mammoth.min.js         DOCX -> HTML conversion
    JSZip.min.js            .pptx (zip/OOXML) reading
    pdf.min.js / pdf.worker.min.js   PDF text extraction (pdf.js)
    exceljs.min.js          Styled .xlsx generation
  samples/                Ready-to-use Source/Output file pairs for every format
  assets/                 (reserved for icons/images if you add any)
  tests/                  Node-based regression tests for utils.js/normalizer.js/comparer.js
```

Each JS file has a single responsibility, communicates only through its
small public API (`window.CCUtils`, `window.CCParser`, etc.), and can be
extended independently. To support a new input format, add one
`parseXxx(file)` function and one branch to `parser.js` — nothing else
needs to change.

> **Note on the Excel library:** the spec text mentions SheetJS; this build
> uses **ExcelJS** instead (`lib/exceljs.min.js`, still 100% client-side,
> still a static file, no server). SheetJS's free/community build does not
> reliably write cell colors, fonts, or borders from the browser, and the
> spec explicitly requires color-coded rows (green/yellow/red/blue),
> header styling, and a frozen header row — all of which ExcelJS supports
> natively. Swapping the library back is a small, isolated change confined
> to `excel.js` if you'd prefer SheetJS for another reason.

---

## Running the Tests

`tests/` covers the comparison engine's pure logic — string similarity,
sentence/table alignment, and the normalizer — using Node's built-in test
runner. No npm install, no dependencies:

```
node --test apps/contentcheck/tests/*.test.js
```

`parser.js` isn't covered here since it needs a real browser DOMParser/
JSZip/pdf.js; changes to it should still be spot-checked in the browser
(open `index.html`, or see the samples in `samples/`).

---

## How Comparison Works

1. **Upload** — Source and Output files are read via the File API (nothing
   ever leaves your browser).
2. **Extract** (`parser.js`) — Each format is parsed into a common raw
   structure of headings/paragraphs/lists/table rows (+ slide/page numbers,
   speaker notes for PPTX).
3. **Normalize** (`normalizer.js`) — Whitespace, quote/dash style, bullet
   symbols, punctuation, and case are normalized per your chosen options.
   **Words are never altered.** The result is a
   `Document > Section > Paragraph > Sentence` hierarchy, matching the
   architecture in the spec, with a unique ID, page/slide number, and
   both original + normalized text on every node.
4. **Compare** (`comparer.js`) — Sections are aligned, then paragraphs
   within matched sections, then sentences within matched/modified
   paragraphs — each level using an LCS-anchored alignment plus a
   Levenshtein/word-overlap similarity score to decide match vs. modify vs.
   missing vs. added. This means a single changed sentence inside an
   otherwise-unchanged paragraph is reported as **one** modified sentence,
   not an entire paragraph rewrite. All comparison is deterministic
   (exact string comparison, Levenshtein distance, LCS) — no AI or
   semantic matching, per the spec.
5. **Detect special changes** — Within any modified sentence, numbers,
   percentages, currency values, fiscal years, dates, emails, and URLs are
   extracted and diffed separately (each togglable in the Options panel),
   so e.g. "12% → 14%" is called out explicitly rather than just marked
   "modified."
6. **Tables** — Rows are grouped by table and aligned the same way;
   changed rows are further diffed cell-by-cell.
7. **Summarize & report** (`report.js`, `excel.js`) — An on-screen summary
   and preview render immediately; the full detail is available in the
   downloadable Excel workbook: Summary, Modified Content, and Detailed
   Comparison — each with header formatting, autofit columns, autofilter,
   a frozen header row, and green/yellow/red/blue status color coding.

---

## Supported Formats & Known Limitations

| Format | Extracts | Notes |
|---|---|---|
| DOCX | Headings, paragraphs, lists, tables | Via mammoth.js |
| PPTX | Slide title, text boxes, tables, speaker notes | Manual OOXML parsing via JSZip — no extra pptx library needed |
| PDF | Text, basic tables | **Text-based PDFs only.** Scanned/image-only PDFs have no extractable text (OCR is out of scope, but the architecture leaves room for it — see below) |
| TXT | All text, split on blank lines | — |

PDF table detection is heuristic (column gaps between text items on the
same line) and works well for simple grid tables; complex multi-line or
merged-cell tables may be read as plain paragraphs instead.

---

## Performance

- Comparison runs at the sentence level with LCS-anchored alignment, which
  keeps each section's comparison close to linear in the common case
  (mostly-matching documents) rather than the full quadratic worst case.
- Large inputs (500-slide decks, 1,000-page Word docs) are processed
  section-by-section so memory stays bounded; the progress bar yields to
  the browser between phases so the UI never appears frozen.
- For very large files, comparison time is dominated by parsing (mammoth/
  pdf.js), which is proportional to document size.

---

## Extending the App

- **New file format:** add `parseXxx()` to `parser.js` (must return the
  documented raw shape) and one line in `parseFile`'s switch statement.
- **New change-detector** (e.g. phone numbers): add a regex to
  `utils.js`'s `PATTERNS`, and a branch in `comparer.js`'s
  `detectSpecialChanges`.
- **New report sheet:** add a builder function in `excel.js` and call it
  from `generateWorkbook`.
- **OCR / semantic matching:** the architecture deliberately keeps
  extraction (`parser.js`) and comparison (`comparer.js`) decoupled from
  everything else, so an OCR-based parser or an optional AI-assisted
  comparison mode could be added later as an alternate implementation
  behind the same interfaces, without touching normalization, reporting,
  or the UI.

---

## Privacy

Everything — file parsing, comparison, and Excel generation — happens
locally in your browser. No file content, comparison results, or telemetry
is ever sent to any server.
