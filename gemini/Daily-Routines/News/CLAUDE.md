# Market Intelligence News Screener

## What This Is

A Python script that calls Claude or Gemini (with live web search/grounding) and generates a self-contained
`news-screener.html` file — a daily finance dashboard covering:

1. **Dips to Watch** — quality equities/ETFs trading down but potentially worth buying
2. **High-Growth Sectors** — sectors with concrete near/medium/long-term growth catalysts
3. **Market News** — trending macro, earnings, IPO, crypto, commodities, policy, and M&A stories

## How to Run

### Option A: Gemini (Recommended if you have a Google API Key)
```bash
export GOOGLE_API_KEY=your_key_here
pip install google-genai
python generate_screener_gemini.py
# → opens/writes news-screener.html
```

### Option B: Claude
```bash
export ANTHROPIC_API_KEY=sk-...
pip install anthropic
python generate_screener.py
# → opens/writes news-screener.html
```

Open `news-screener.html` in any browser. No server needed — it's fully static.

## Files

| File | Purpose |
|------|---------|
| `generate_screener.py` | Claude version — fetches data, renders HTML |
| `generate_screener_gemini.py` | Gemini version — fetches data, renders HTML |
| `news-screener.html` | Output — regenerated each run |

## Key Behaviors

- Models use web search (Claude: `web_search_20250305`, Gemini: `google_search`) to pull real-time market data
- All content is from the past 24–48 hours
- Each item is scored by **conviction** (dips/sectors) or **significance** (news) on a 1–5 scale
- The HTML file is fully self-contained — no CDN, no JS framework, no external dependencies

## Models

- **Claude:** Uses `claude-sonnet-4-6` with `max_tokens: 8000`.
- **Gemini:** Uses `gemini-2.0-flash`.

## Output Format

Three sections of cards, color-coded:
- 🔴 **Dips** — red accent, shows ticker + % change + conviction dots
- 🟢 **Growth Sectors** — green accent, shows timeframe + catalyst callout box
- 🔵 **Market News** — blue/color-coded by category (macro, earnings, IPO, etc.)
