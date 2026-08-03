# Creative Solutions — Internal Tools Suite

A single-entry-point suite that brings together three standalone,
browser-only tools behind one shared left-hand navigation:

- **KPI Dashboard** — utilization/KPI dashboard with Excel-driven charts and PPTX export
- **Newsletter Builder** ("The Signal") — converts a Word doc into a styled HTML newsletter
- **Content Check** ("PEP SmartQC") — sentence-level document comparison with an Excel diff report

No installation, no build step, no backend required. Extract the ZIP and
open `index.html`. Content Check has one optional, off-by-default
"enhanced parsing" toggle backed by a small Python service — see
[Optional: enhanced parsing backend](#optional-enhanced-parsing-backend).

---

## Quick Start

```bash
# From inside the CreativeSolutionsSuite folder:
python3 -m http.server 8000
# then open http://localhost:8000
```

Serving it (rather than double-clicking `index.html`) is recommended here
specifically because the shell loads each tool into an `<iframe>`, and
some browsers restrict iframe/script loading from `file://` pages more
strictly than from `http://localhost`. If your browser handles `file://`
fine, opening `index.html` directly also works.

---

## How it's put together

```
CreativeSolutionsSuite/
  index.html              The shell: sidebar nav + iframe content area
  css/shell.css
  js/shell.js
  assets/logo.svg
  apps/
    dashboard/             KPI Dashboard — untouched functionality, CDN libs localized
      index.html
      lib/  (xlsx.full.min.js, chart.umd.js, pptxgen.bundle.js)
    newsletter/            Newsletter Builder ("The Signal") — untouched functionality
      index.html
      lib/  (mammoth.browser.min.js)
    contentcheck/           Content Check ("PEP SmartQC") — the app from previous requests
      index.html, css/, js/, lib/, samples/, README.md
```

**Why iframes, not one merged page?** Each tool was already a complete,
independently-built application with its own CSS reset, its own global
JavaScript state, and — critically — its own same-named things (e.g. all
three define their own `.panel`, `.btn`, color variables, etc., and the
Dashboard and Content Check both happen to use an `app.js`/`ui.js`
naming pattern). Merging all three into a single DOM/script context would
risk silent CSS collisions and JavaScript variable clashes between apps
that were never written to coexist. Loading each into its own `<iframe>`
gives each tool a fully isolated document — guaranteeing nothing in one
tool can break another — while the shell still gives you one URL, one nav,
one set of browser tabs to manage.

The only change made *inside* each tool was pointing their `<script src>`
tags at a locally-bundled copy of the exact same library version they
already used from a CDN (SheetJS 0.18.5, Chart.js 4.4.0, PptxGenJS 3.12.0
for the Dashboard; Mammoth.js 1.6.0 for the Newsletter Builder) — so the
whole suite works fully offline aside from optional Google Fonts, and
every tool's actual functionality, layout, and logic is unchanged from
what you uploaded.

---

## What's unified vs. what isn't

**Fully unified (by construction, since it's shared chrome):**
- One left sidebar with the Creative Solutions logo + team name, always visible
- One navigation model — click a tool, it loads in place, browser back/forward isn't needed
- The active tool and sidebar collapsed/expanded state are remembered
  (`localStorage`) so reopening the suite returns you to where you left off
- A collapsible sidebar (icon-only rail) for smaller screens, with a manual
  collapse toggle on desktop too

**Intentionally left as-is (each tool's own internal design):**
Each tool keeps the internal visual identity it was already built with —
the Dashboard's dark KPI-console theme, the Newsletter Builder's cream/
editorial "Signal" branding, and Content Check's light blue tool styling.
These were each deliberately designed for their own purpose (a dark
data-dense dashboard vs. a newsletter template preview vs. a QC utility),
and forcing them into one identical internal color/font scheme would mean
substantially rewriting each tool's CSS rather than just wrapping them —
a bigger, separate piece of work. If you'd like a deeper pass to make the
*internal* look of all three visually match (same palette/fonts inside
each tool, not just the shared sidebar), that's a reasonable follow-up —
just say the word.

---

## Optional: enhanced parsing backend

`backend/` is a small FastAPI service (`python-pptx` / `python-docx` /
`pdfplumber`) that Content Check can optionally send files to for more
robust parsing — better PDF table detection and PowerPoint SmartArt
support than the in-browser JS parser can manage. It is entirely
opt-in: Content Check works fully offline by default, and this service
is only called when a user explicitly checks "Use enhanced parsing" in
the Upload panel, which clearly discloses that doing so uploads the
files to a server. No uploaded content is stored — files are parsed in
memory and discarded once the response is sent.

Deployed as a second free-tier Render service via the root `render.yaml`
(`content-check-backend`, pointing `apps/contentcheck/js/parser.js`'s
`BACKEND_URL` at its `.onrender.com` URL). Free-tier web services spin
down after 15 minutes idle, so the first request after a quiet period
takes ~30-50s to wake up — the toggle's help text mentions this.

To run it locally:
```bash
cd backend
pip install -r requirements.txt
uvicorn app:app --reload --port 8900
```

## Adding another tool later

1. Drop the new tool's folder under `apps/<name>/` (localize any CDN
   `<script src>`/`<link href>` the same way the existing tools were).
2. Add one `<button class="nav-item" data-tool="<name>">…</button>` to the
   nav list in `index.html`.
3. Add one entry to the `TOOLS` map at the top of `js/shell.js`.

That's the entire integration surface — no other file needs to change.
