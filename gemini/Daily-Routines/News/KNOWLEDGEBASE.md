# Knowledgebase — Market Intelligence Screener

## Data Schema

### `dips[]` — Equities Trading Down

| Field | Type | Description |
|-------|------|-------------|
| `ticker` | string | Exchange ticker symbol (e.g. `"NVDA"`) |
| `company` | string | Full company name |
| `title` | string | Why it's interesting despite the dip (≤15 words) |
| `summary` | string | 2–3 sentences: what happened, why it may be a buying opportunity |
| `change_pct` | string | Today's % change, e.g. `"-3.2"` (negative = down) |
| `sector` | string | Sector label (e.g. `"Semiconductors"`) |
| `source` | string | Publication name |
| `url` | string | Article URL |
| `conviction` | int 1–5 | How compelling the opportunity is |

**Selection criteria Claude is prompted to apply:**
- Company has strong fundamentals (earnings, balance sheet, competitive moat)
- Price drop is due to temporary/overblown factors: sector rotation, market fear, one-time miss
- NOT in structural decline (e.g. avoid legacy tech being displaced)
- Ideally showing insider buying, analyst upgrades, or upcoming catalysts

---

### `sectors[]` — High-Growth Sectors

| Field | Type | Description |
|-------|------|-------------|
| `sector` | string | Sector/industry name |
| `title` | string | Growth thesis headline |
| `summary` | string | 2–3 sentences on concrete growth drivers |
| `catalyst` | string | Specific news event or data point driving the thesis |
| `timeframe` | string | `"near-term"` / `"medium-term"` / `"long-term"` |
| `source` | string | Publication name |
| `url` | string | Article URL |
| `conviction` | int 1–5 | Strength of the growth signal |

**Selection criteria:**
- Tied to a specific event (contract wins, earnings beats, policy tailwinds, tech adoption inflection)
- Not vague background trends — must be news-driven from past 24–48 hours
- Timeframe reflects how soon the growth is expected to materialise

---

### `news[]` — Market News

| Field | Type | Description |
|-------|------|-------------|
| `title` | string | Concise headline |
| `summary` | string | 2–3 sentences |
| `category` | string | See category taxonomy below |
| `source` | string | Publication name |
| `url` | string | Article URL |
| `significance` | int 1–5 | Market-wide importance |

**Category taxonomy:**

| Category | Examples |
|----------|---------|
| `macro` | Fed decisions, CPI, jobs report, GDP |
| `earnings` | Quarterly results, guidance updates |
| `ipo` | New listings, SPAC activity |
| `crypto` | BTC/ETH price moves, regulation, protocols |
| `commodities` | Oil, gold, copper, agricultural |
| `policy` | Tariffs, regulation, central bank statements |
| `m&a` | Mergers, acquisitions, divestitures |

---

## Prompt Engineering Notes

### System Prompt Design

The system prompt is minimal and schema-first. Key choices:
- **"Return ONLY raw JSON"** — no markdown, no explanation. Claude tends to wrap JSON in backticks if not instructed otherwise.
- Schema is written inline as a comment block, not as JSON Schema — this fits within the natural language context window more efficiently.
- Arrays named `dips`, `sectors`, `news` (not nested or deeply namespaced) to reduce hallucination on field names.
- `conviction` and `significance` use the same 1–5 scale but different names to signal different evaluation dimensions.

### User Prompt Design

Includes today's date explicitly — Claude uses this to anchor recency filtering. Without the date, Claude may surface older stories.

The prompt lists the three deliverables as a numbered list to match the system schema, reinforcing structure consistency.

### JSON Extraction

Claude sometimes prefixes/suffixes the JSON with text even when instructed not to. The parser:
1. Tries `json.loads(raw.strip())` — catches clean responses
2. Falls back to `re.search(r'\{[\s\S]*\}', raw)` — extracts embedded JSON objects
3. Raises `ValueError` with a preview if both fail

---

## HTML Rendering Pipeline

```
fetch_market_data()
    └─ returns dict: { dips, sectors, news }

build_html(data, generated_at)
    ├─ _dip_card(item)     → HTML string per dip
    ├─ _sector_card(item)  → HTML string per sector
    ├─ _news_card(item)    → HTML string per news item
    ├─ _section(...)       → wraps cards in a labeled section
    └─ returns full HTML page string
```

Helper functions:
- `_safe(text)` — HTML-escapes all user-facing strings to prevent XSS in card content
- `_dots(value, max, color)` — renders conviction dots (● filled, ● empty)
- `_link(url, label, color)` — renders a `Read →` anchor; skips if URL is empty or not HTTP

### CSS Architecture

All styles are in a single `CSS` constant embedded in a `<style>` block — no external deps.
The layout uses CSS Grid with `auto-fill, minmax(320px, 1fr)` for responsive card columns.
Dark theme uses a fixed palette:
- Background: `#070b14`
- Card: `#111827`
- Border: `#1f2937`
- Text primary: `#e5e7eb` / secondary: `#9ca3af` / muted: `#6b7280`

Section accent colors:
- Dips: `#ef4444` (red)
- Sectors: `#10b981` (emerald)
- News: category-specific (blue, purple, yellow, orange, etc.)

---

## API Usage

### Claude
- **Model:** `claude-sonnet-4-6`
- **Tool:** `web_search_20250305` (Anthropic server-side)
- **Max tokens:** `8000`

### Gemini
- **Model:** `gemini-2.0-flash`
- **Tool:** `google_search` (Google Search Grounding)
- **Response Format:** `application/json` (ensures structural integrity)

### Approximate token cost per run

| Stage | Tokens |
|-------|--------|
| System prompt | ~350 |
| User prompt | ~80 |
| Web search context injected | ~2000–4000 (estimated) |
| Output JSON | ~1500–2500 |

---

## Extending the Screener

**Add a new data category:**
1. Add the field to `SYSTEM_PROMPT` schema comment
2. Write a `_yourtype_card(item)` renderer function
3. Add a `_section(...)` call in `build_html()`
4. Update `KNOWLEDGEBASE.md` schema tables

**Schedule daily runs:**
```bash
# crontab: run at 7am every weekday
0 7 * * 1-5 cd /path/to/project && python generate_screener.py
```

**Add filtering in the HTML:**
The rendered HTML currently has no JS. To add topic/category filtering, wrap cards in
`data-category` attributes and add a vanilla JS filter function before `</body>`.
