# CLAUDE.md: 2026 World Cup Brand Tracker

Global rules in `~/.claude/CLAUDE.md` (Windows/PowerShell, git etiquette, scope,
verification, UK English) apply here and are not repeated. This file covers only what
you cannot infer from the code.

## What this is

A static, hand-written site publishing the **2026 World Cup Brand Tracker** for
**Voice of the Trade** (a Potentia Sports product). Two survey waves, two panels each:

| Wave | Fielded | Trade n | Public n | Asks about |
| --- | --- | --- | --- | --- |
| 1 | 5-10 June 2026 | 116 | 244 | **Intent** (will you watch, what would sway you) |
| 2 | 14-16 July 2026 | 136 | 200 | **Behaviour** (what you watched, how late, what you did) |

No framework, no build step, no package.json. Plain HTML + CSS + vanilla JS, with
survey aggregates embedded inline so every page works offline from `file://`.

**This repo is public** and the whole root is served as the site. Nothing that isn't
publishable goes in a tracked file. Raw respondent-level data, the questionnaire and
the internal reference folders are gitignored on purpose (`.gitignore` explains each).

## Layout

Root = the deployed site. Numbered folders are the working project and mostly
gitignored.

```
index.html        tab shell: loads the three tabs as auto-sized iframes
insights.html     narrative findings. Figures are HARD-CODED in the copy.
explore.html      1.4 MB. Every question, break-by-demographic. Data inline as JS consts.
compare.html      "compare yourself" quiz. Data inline in <script type="application/json">.
brand.css         shared design system, used by insights.html and compare.html ONLY
embed-child.js    height reporting + in-app nav, loaded by all three tab pages
pptx-export.js    PowerPoint export (PREVIEW, ?export=1 only), loaded by explore.html ONLY
libs/             vendored pptxgen.bundle.js (MIT), fetched on demand by the export
fonts/            licensed Rework Text woff/woff2 (the .otf trial files are gitignored)
05 Web Report/_build/   python data pipeline (gitignored, internal)
```

## Architecture

**Shell + tabs.** `index.html` is a tab bar over three lazy-loaded iframes. Each child
loads `embed-child.js`, which posts `{type:"wc-embed-size", height}` up; the shell sets
`iframe.style.height` to match and forwards its own total height to the host site as
`{type:"resize"}`. In-app links use `data-nav="explore"` and are routed through the
shell via `{type:"wc-nav"}` so they don't break out of the host iframe. Opened standalone,
the child does nothing and the `href` fallback (`index.html#explore`) takes over.

**Wave switching.** `insights.html` and `explore.html` each hold two `.wave-panel` divs
behind a `.wavebar` switcher. In `explore.html`, Wave 2 chart keys are **prefixed**
(`data-chart="W2Q1"` -> `#chart-W2Q1`) so DOM ids stay unique; `DK()` maps a chart key
back to its data key and `useWave()` repoints the `DATA` global per card before render.
Wave 2 routed multis get their base from the `bases` map via `multiBase()`.
`compare.html` uses a `WAVES` registry (`{data, questions}`) and swaps `DATA` +
`QUESTIONS` in `switchWave()`.

**Styling is duplicated in four places.** `brand.css` is linked only by `insights.html`
and `compare.html`. `index.html` and `explore.html` carry their own inline `<style>`
copies of the same tokens, pitch backdrop and pitch markings. A design change must be
applied to each place that uses it. Check all four before declaring a restyle done.

## Data pipeline

- `05 Web Report/_build/compute_figures.py` -> `data.json` (Wave 1)
- `05 Web Report/_build/compute_figures_wave2.py` -> `data_w2.json` (Wave 2)

Both read the xlsx files in `01 Raw Data/` and print a reconciliation block against the
published LinkedIn carousel figures. **Check that block before trusting any number.**
Run them from `_build/`. They only write JSON, so they are safe.

Aggregates reach the site by being pasted inline into the page. `_build/` is gitignored,
so nothing is fetched at runtime.

**`assemble.py` is superseded. Do not run it.** It stitches `template.html` + `data.json`
and overwrites the repo-root `index.html`, which is now the tab shell. `template.html` is
~640 lines behind the live files. Running it clobbers the shell and the redesign. Edit
the root HTML directly. Same for `05 Web Report/worldcup_tracker_vott_wave1.html`: it is
tracked and deployed, but it is the superseded single-page version.

### Wave 2 traps that have already cost time

1. **Never assume `Q<n>` means the same thing across waves.** Both waves have Q1..Q14 and
   only the S1-S6 demographics carry across.
2. **Q14 (nation supported) bases on all respondents**, not just answerers. Wave 2 records
   a non-answer explicitly (`noanswerQ14_n1`); "not supporting anyone" is a real answer.
   Basing on answerers overstates England as 84%/75% instead of the published 74%/62%.
3. **Short rows.** `dict(zip(headers, row))` silently drops trailing columns (Q15, Q16).
   Pre-seed every header to `None` first.
4. Brand recall (Q8/Q13) is coded from the TEXT LABELS workbook by alias matching and
   lands within ~2pp of published. **Use the published figures in copy** so the site and
   the LinkedIn carousel agree.

Anything hard-coded in `insights.html` copy must be re-checked against the JSON when the
data changes. Nothing recomputes it.

## PowerPoint export (preview only)

**Off on the live site.** `pptx-export.js` returns immediately unless `?export=1` is in its
own query string, so the deployed Pages site shows no tick boxes and no Export button.
`index.html` reads the flag off the shell URL and forwards it to each tab iframe, so
`https://brighty36.github.io/world_cup_survey/?export=1` turns it on, and
`explore.html?export=1` works standalone. The tab links use a bare `#name`, which preserves
the query string, so the flag survives switching tabs. This is hidden, not
access-controlled: anyone who knows the flag can use it.

When enabled it adds a tick box to every Explore card and an Export button that writes the
selected charts to a `.pptx`, one native editable chart per slide, honouring each card's
own Audience and Break-by selection.

**The Export button is repeated inline on every card and in both wave toolbars, and that is
deliberate.** The shell sizes this page's iframe to its full content height and never
scrolls it, so the iframe viewport IS the whole document: `position:fixed` pins to the
bottom of a 14,000px box, far below anything visible. The page's own sticky `.jumpbox` is
inert for the same reason. Do not "fix" this with a floating bar.

**It reads the numbers back out of the rendered DOM, not out of `DATA`.** Every renderer
funnels into `pairBar()` and `singleBar()`, and each bar carries its true value in
`data-w`, so scraping inherits sort order, top-N slicing, break-by grouping and routed
bases for free. The cost is a hard coupling to that markup: **if you change what
`pairBar`/`singleBar` emit, or the class names `.pair` / `.pair-lab` / `.pair-row.t|p|a` /
`.bar-row` / `.bar-lab` / `.bar-fill` / `.stat` / `.div-side`, fix `readRows()` to match.**
Read `data-w`, never `style.width` — bar widths animate from 0 and are wrong on frame one.

`libs/pptxgen.bundle.js` is vendored unmodified (PptxGenJS 4.0.1, MIT, JSZip included) and
injected on first export, so the page weight is unchanged until someone actually exports.
Slides use **Arial** deliberately: Rework Text is licensed for the web and must not be
shipped into a deck that leaves the building.

## Verification

Serve over HTTP rather than `file://` when testing the shell, because the iframes and
postMessage handshake need an origin:

```powershell
python -m http.server 8731 --bind 127.0.0.1
```

Then open `http://127.0.0.1:8731/`. Confirm the Chrome MCP is connected with one
screenshot before starting UI work.

**Three things look like bugs and are not.** Check these before diagnosing:

1. **rAF is throttled to zero in a background tab.** `embed-child.js` posts its height
   inside `requestAnimationFrame`, so in an automated or non-foregrounded tab the iframe
   sits at its 600px CSS default and looks broken. Check `document.visibilityState`
   first. Taking a screenshot forces a render and the resize completes. `fillBars()` in
   `explore.html` uses rAF too, so bars stay at width 0.
2. **Bars animate over 0.9s** (`transition:width`). A screenshot straight after a scroll
   catches them mid-flight. Read `style.width` from the DOM instead of trusting the image.
3. **`html{scroll-behavior:smooth}`** means a scripted `scrollTo` loop interrupts its own
   animation, so intermediate positions never render and IntersectionObservers never fire.
   Use real scrolling or `behavior:'instant'`. Do not "fix" the page for this.

The reveal IntersectionObserver's implicit root is the **top-level** viewport, not the
iframe's, so cards reveal only when scrolled onto screen in the host page. Do not force it
by re-observing on resize: `unobserve` + `observe` restarts async delivery, starves the
callback and breaks every card.

Because the pages are huge single files, syntax-check extracted script blocks with
`node --check` before opening a browser.

## Shipping

Branch, commit, push, PR against `main`. Pushing to `main` triggers
`.github/workflows/deploy-pages.yml`, which publishes the repo root to GitHub Pages at
`https://brighty36.github.io/world_cup_survey/`. There is no staging, so anything merged
is live. Verify locally first.
