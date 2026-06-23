# CLAUDE.md

Guidance for AI assistants (Claude Code and others) working in this repository.

## What this is

A **single-file, client-side web dashboard** that monitors battery / deep-tech
industry news in real time. The UI and all copy are in **Korean** (`lang="ko"`).

- 🔋 **배터리·딥테크 모니터링** — "Battery & Deep-Tech Monitoring"
- Tracks news across categories: 수주/계약 (orders/contracts), 생산 차질
  (production disruptions), 운송 차질 (logistics disruptions), M&A·딜
  (M&A/deals), 기술 동향 (tech trends), 방산·딥테크 (defense/deep-tech).
- Collects **Google News RSS** for both Korean (🇰🇷) and English (🇺🇸)
  search queries, classifies each article, and renders charts, summary cards,
  a company-frequency list, source tabs, and a filterable news feed.

## Repository layout

This repo contains **exactly one file**:

```
index.html   ← the entire application (HTML + CSS + JS, all inline)
```

There is **no build step, no package manager, no dependencies to install, and
no test suite.** The only external runtime dependency is Chart.js, loaded from
a CDN (`cdn.jsdelivr.net/npm/chart.js@4.4.0`).

> **Note on the file header:** The comment at the top of `index.html` describes
> this as the "standalone" version and refers to an "official" Flask backend
> (`app.py`, port 5001, with Naver/TheBell support). **That `app.py` does not
> exist in this repository** — only the standalone `index.html` is present. Do
> not assume a backend exists; if a task requires one, confirm with the user
> before creating it.

## How to run

Because everything is client-side, just open the file:

- Open `index.html` directly in a browser (works from `file://`), **or**
- Serve it locally, e.g. `python3 -m http.server 8000` then visit
  `http://localhost:8000/index.html`.

On load it immediately fetches news and then auto-refreshes every **10 minutes**
(`AUTO_REFRESH = 10 * 60 * 1000`).

## Architecture (all inside `index.html`)

The file has three sections: a `<style>` block, the static HTML markup, and two
`<script>` blocks (config, then the engine).

### 1. Configuration script (first `<script>`)

Plain JS arrays/objects — **this is the main place to edit content/behavior**:

- `BATTERY_COMPANIES` — company names matched against article text (Korean +
  English aliases, e.g. `"LG에너지솔루션"`, `"LGES"`).
- `CATEGORY_KEYWORDS` — keyword lists per category; an article gets a category
  if any keyword appears in its text.
- `SEVERITY_KEYWORDS` — `critical` / `warning` / `info` keyword lists; **order
  matters** — the first matching tier wins (critical is checked first).

### 2. Engine script (second `<script>`)

- `CORS_PROXIES` — ordered list of public CORS proxy builders (codetabs,
  allorigins, corsproxy.io). `fetchWithProxy()` tries them in order until one
  returns RSS-looking text. Open proxies are first so `file://` works.
- `SEARCH_QUERIES` (Korean) and `SEARCH_QUERIES_EN` (English) — the Google News
  search terms. Add/remove queries here to change coverage.
- `fetchGoogleNewsRSS(query, locale)` builds a Google News RSS URL with the
  right locale params (`hl/gl/ceid`) and parses the result.
- `parseRSS()` parses RSS XML with `DOMParser`; `hashStr()` makes a stable id
  from `title + source` for de-duplication.
- `fetchAllNews()` runs all queries in parallel, de-dupes by id, classifies,
  and renders progressively as results arrive.
- Classification: `classifyCategory()`, `classifySeverity()`,
  `findCompanies()`.
- Filtering: `applyFilters()` plus `filterBy*` helpers (category, severity,
  company, source) driven by the dropdowns, search box, summary cards, and
  source tabs. Event handling uses **event delegation** (set up in the
  `DOMContentLoaded` listener) — dynamically rendered elements carry `data-*`
  attributes rather than inline handlers.
- Rendering: `renderDashboard()` orchestrates `renderSummaryCards`,
  `renderCharts` (Chart.js doughnut + bar), `renderCompanyList`,
  `renderSourceTabs`, `renderAlertBanner`, `renderNewsFeed`.
- Translation: English articles get a 🌐 번역 button that calls the free
  **MyMemory** API (`translateArticle` / `_mymemory`, en→ko, no key), with an
  in-memory `_translationCache` and original/translation toggle.
- Utilities: `esc()` (HTML-escape — **always use it when injecting text into
  `innerHTML`**), `truncate()`, `fmtDate()` (relative Korean time),
  `showToast()`, `showLoading()`.

## Conventions to follow

- **Keep it single-file.** All HTML, CSS, and JS stay inline in `index.html`
  unless the user explicitly asks to split the project.
- **No frameworks/bundlers.** Vanilla JS only; the sole CDN dependency is
  Chart.js. Don't introduce npm/build tooling without being asked.
- **Korean-first UI.** User-facing strings are Korean. Match the existing tone
  and the emoji-prefixed labels (📝 수주, 🏭 생산 차질, etc.).
- **Security:** any user/feed-derived text inserted via `innerHTML` must go
  through `esc()` (and `CSS.escape()` for selectors, as in `translateArticle`).
- **Styling:** use the existing CSS custom properties in `:root` (e.g.
  `--bg-card`, `--accent-blue`) and the established dark-theme palette; the
  layout is responsive with breakpoints at 1200/768/480px.
- **Comments** in the source are in Korean and use box-drawing separators
  (`═══`, `──`); keep that style if adding comments nearby.
- **Categories are a fixed set.** If you add a category, update *all* of:
  `CATEGORY_KEYWORDS`, the summary card markup, the filter `<select>`, the
  chart labels/colors in `renderCharts`, the `computeStats` stats object, the
  `catBadges` map, and `updateActiveCards`.

## Git workflow

- Active development branch for AI-assisted work: **`claude/claude-md-docs-4032wb`**
  (do not push to `main` without explicit permission).
- Push with `git push -u origin <branch-name>`.
- Do **not** open a pull request unless the user explicitly asks.

## Gotchas

- The app depends on **public CORS proxies** and **Google News RSS** being
  reachable; if no news loads, it's usually a proxy/network issue, not a code
  bug (an empty state with a 📡 message is shown).
- `encodeURIComponent` is used deliberately for queries; the original code
  notes a double-encoding pitfall with `+` — prefer spaces in query strings.
- There is no persistence — all state is in memory and reset on refresh.
</content>
</invoke>
