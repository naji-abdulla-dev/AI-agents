#!/usr/bin/env python3
"""
Generate a self-contained HTML finance news screener using Gemini.

Uses Gemini (gemini-2.0-flash) with Google Search grounding to fetch today's market intelligence:
  - Equities trading down that look like buying opportunities
  - High-growth sectors with concrete catalysts
  - Trending market / finance news

Usage:
    export GOOGLE_API_KEY=your-api-key
    pip install google-genai
    python generate_screener_gemini.py
    # → writes news-screener.html
"""

import json
import os
import re
import sys
from datetime import datetime

try:
    from google import genai
    from google.genai import types
except ImportError:
    print("Error: google-genai package not found. Please run: pip install google-genai", file=sys.stderr)
    sys.exit(1)


# ── Prompts ────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are a financial market intelligence analyst with real-time Google Search access.
Return ONLY a raw JSON object — no markdown, no backticks, no preamble.

{
  "dips": [...],     // 6-8 quality stocks/ETFs trading DOWN but worth watching
  "sectors": [...],  // 4-6 high-growth sectors with concrete catalysts
  "news": [...]      // 6-10 trending market/finance stories
}

dips item:
  ticker, company, title (why interesting despite dip), summary (2-3 sentences),
  change_pct (e.g. "-3.2"), sector, source, url, conviction (1-5)

sectors item:
  sector, title (growth thesis), summary (2-3 sentences), catalyst (specific news/event),
  timeframe ("near-term" | "medium-term" | "long-term"), source, url, conviction (1-5)

news item:
  title, summary (2-3 sentences),
  category ("macro" | "earnings" | "ipo" | "crypto" | "commodities" | "policy" | "m&a"),
  source, url, significance (1-5)

Rules:
- All items from today or past 24-48 hours only
- Dips: quality companies down on temporary/overblown factors, not structural decline
- Sectors: concrete catalysts only, not vague background trends
- Sort each array by conviction/significance descending
- Include real article URLs when found"""


def _user_prompt() -> str:
    today = datetime.now().strftime("%A, %B %d, %Y")
    return (
        f"Today is {today}. Use Google Search to find current market data and news. Find:\n"
        "1. Quality equities/ETFs that are down today but look like buying opportunities\n"
        "2. Sectors with strong future growth signals from recent news\n"
        "3. Other important market/finance stories from the past 24-48 hours\n\n"
        "Return ONLY the raw JSON object according to the system instructions."
    )


# ── API ────────────────────────────────────────────────────────────────────────

def fetch_market_data(api_key: str) -> dict:
    client = genai.Client(api_key=api_key)

    print("Searching for market intelligence using Gemini...", flush=True)
    
    # Define the grounding tool
    grounding_tool = types.Tool(
        google_search=types.GoogleSearch()
    )

    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=_user_prompt(),
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            tools=[grounding_tool],
            response_mime_type="application/json",
        )
    )

    raw = response.text

    try:
        return json.loads(raw.strip())
    except json.JSONDecodeError:
        # Fallback in case Gemini includes markdown or other text
        match = re.search(r"\{[\s\S]*\}", raw)
        if match:
            return json.loads(match.group(0))
        raise ValueError(f"No valid JSON in response:\n{raw[:600]}")


# ── HTML helpers ───────────────────────────────────────────────────────────────

def _esc(text) -> str:
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
    )


def _dots(value: int, max_val: int = 5, color: str = "#10b981") -> str:
    filled = f'<span style="color:{color}">&#9679;</span>'
    empty = '<span style="color:#374151">&#9679;</span>'
    return "".join(filled if i < value else empty for i in range(max_val))


def _link(url: str, color: str = "#60a5fa") -> str:
    if not url or not url.startswith("http"):
        return ""
    return (
        f'<a href="{_esc(url)}" target="_blank" rel="noopener" '
        f'style="color:{color};font-size:12px;font-weight:600;text-decoration:none;">Read &#8594;</a>'
    )


# ── Card renderers ─────────────────────────────────────────────────────────────

def _dip_card(item: dict) -> str:
    ticker = _esc(item.get("ticker", ""))
    company = _esc(item.get("company", ""))
    title = _esc(item.get("title", ""))
    summary = _esc(item.get("summary", ""))
    change = _esc(str(item.get("change_pct", "")))
    sector = _esc(item.get("sector", ""))
    source = _esc(item.get("source", ""))
    conviction = max(1, min(5, int(item.get("conviction", 3))))
    url = item.get("url", "")

    change_color = "#ef4444" if "-" in change else "#10b981"
    change_html = (
        f'<span style="font-size:13px;font-weight:700;color:{change_color};">{change}%</span>'
        if change else ""
    )

    return f"""
<div class="card">
  <div class="card-bar" style="background:#ef4444;"></div>
  <div class="card-header">
    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <span class="badge" style="background:#ef444422;color:#ef4444;">{ticker}</span>
      {change_html}
      <span class="badge" style="background:#1f293780;color:#9ca3af;">{sector}</span>
    </div>
    <div class="conviction">{_dots(conviction, color="#f59e0b")} {conviction}/5</div>
  </div>
  <div style="font-size:12px;color:#6b7280;margin-bottom:5px;">{company}</div>
  <h3 class="card-title">{title}</h3>
  <p class="card-summary">{summary}</p>
  <div class="card-footer">
    <span class="source">{source}</span>
    {_link(url, color="#ef4444")}
  </div>
</div>"""


def _sector_card(item: dict) -> str:
    sector = _esc(item.get("sector", ""))
    title = _esc(item.get("title", ""))
    summary = _esc(item.get("summary", ""))
    catalyst = _esc(item.get("catalyst", ""))
    timeframe = _esc(item.get("timeframe", ""))
    source = _esc(item.get("source", ""))
    conviction = max(1, min(5, int(item.get("conviction", 3))))
    url = item.get("url", "")

    tf_colors = {
        "near-term": "#10b981",
        "medium-term": "#f59e0b",
        "long-term": "#8b5cf6",
    }
    tf_color = tf_colors.get(timeframe.lower(), "#60a5fa")
    tf_badge = (
        f'<span class="badge" style="background:{tf_color}22;color:{tf_color};">{timeframe}</span>'
        if timeframe else ""
    )
    catalyst_box = (
        f'<div class="catalyst-box">&#9889; {catalyst}</div>'
        if catalyst else ""
    )

    return f"""
<div class="card">
  <div class="card-bar" style="background:#10b981;"></div>
  <div class="card-header">
    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <span class="badge" style="background:#10b98122;color:#10b981;">{sector}</span>
      {tf_badge}
    </div>
    <div class="conviction">{_dots(conviction, color="#10b981")} {conviction}/5</div>
  </div>
  <h3 class="card-title">{title}</h3>
  <p class="card-summary">{summary}</p>
  {catalyst_box}
  <div class="card-footer">
    <span class="source">{source}</span>
    {_link(url, color="#10b981")}
  </div>
</div>"""


_CAT_COLORS = {
    "macro": "#60a5fa",
    "earnings": "#a78bfa",
    "ipo": "#34d399",
    "crypto": "#fbbf24",
    "commodities": "#fb923c",
    "policy": "#f87171",
    "m&a": "#c084fc",
}


def _news_card(item: dict) -> str:
    title = _esc(item.get("title", ""))
    summary = _esc(item.get("summary", ""))
    category = _esc(item.get("category", "news"))
    source = _esc(item.get("source", ""))
    significance = max(1, min(5, int(item.get("significance", 3))))
    url = item.get("url", "")

    color = _CAT_COLORS.get(category.lower(), "#60a5fa")

    return f"""
<div class="card">
  <div class="card-bar" style="background:{color};"></div>
  <div class="card-header">
    <span class="badge" style="background:{color}22;color:{color};">{category.upper()}</span>
    <div class="conviction">{_dots(significance, color=color)} {significance}/5</div>
  </div>
  <h3 class="card-title">{title}</h3>
  <p class="card-summary">{summary}</p>
  <div class="card-footer">
    <span class="source">{source}</span>
    {_link(url, color=color)}
  </div>
</div>"""


def _section(title: str, subtitle: str, accent: str, cards_html: str, count: int) -> str:
    return f"""
<section class="section">
  <div class="section-header">
    <div>
      <h2 class="section-title" style="color:{accent};">{title}</h2>
      <p class="section-subtitle">{subtitle}</p>
    </div>
    <span class="count-badge" style="background:{accent}22;color:{accent};">{count} items</span>
  </div>
  <div class="grid">{cards_html}</div>
</section>"""


# ── CSS ────────────────────────────────────────────────────────────────────────

CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: 'IBM Plex Sans', 'SF Pro Display', system-ui, -apple-system, sans-serif;
  background: #070b14;
  color: #e5e7eb;
  min-height: 100vh;
  padding: 32px 20px 60px;
}
.container { max-width: 1100px; margin: 0 auto; }

.header { margin-bottom: 36px; }
.header-eyebrow {
  display: flex; align-items: center; gap: 10px; margin-bottom: 6px;
}
.header-title {
  font-size: 26px; font-weight: 800; letter-spacing: -0.03em;
  background: linear-gradient(135deg, #60a5fa, #a78bfa);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
}
.header-meta { font-size: 13px; color: #6b7280; margin-top: 4px; }

.stats-bar {
  display: flex; gap: 28px; flex-wrap: wrap; align-items: center;
  padding: 14px 20px;
  background: #111827; border: 1px solid #1f2937; border-radius: 12px;
  margin-bottom: 40px;
}
.stat { display: flex; flex-direction: column; gap: 3px; }
.stat-value { font-size: 22px; font-weight: 700; line-height: 1; }
.stat-label { font-size: 11px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.05em; }

.section { margin-bottom: 44px; }
.section-header {
  display: flex; justify-content: space-between; align-items: flex-start;
  margin-bottom: 16px;
}
.section-title { font-size: 17px; font-weight: 800; letter-spacing: -0.02em; margin-bottom: 3px; }
.section-subtitle { font-size: 12px; color: #6b7280; }
.count-badge {
  font-size: 12px; font-weight: 700; padding: 4px 12px;
  border-radius: 100px; white-space: nowrap; flex-shrink: 0;
}

.grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
  gap: 12px;
}

.card {
  background: #111827; border: 1px solid #1f2937; border-radius: 12px;
  padding: 18px 18px 14px; position: relative; overflow: hidden;
  transition: border-color 0.2s ease, transform 0.2s ease, box-shadow 0.2s ease;
}
.card:hover {
  border-color: #374151;
  transform: translateY(-2px);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
}
.card-bar { position: absolute; top: 0; left: 0; right: 0; height: 3px; }
.card-header {
  display: flex; justify-content: space-between; align-items: flex-start;
  margin-bottom: 10px; gap: 8px;
}
.badge {
  font-size: 11px; font-weight: 700; padding: 3px 10px;
  border-radius: 100px; white-space: nowrap; letter-spacing: 0.04em;
}
.conviction { font-size: 11px; color: #6b7280; white-space: nowrap; flex-shrink: 0; }
.card-title {
  font-size: 14px; font-weight: 700; line-height: 1.4;
  color: #f3f4f6; margin-bottom: 8px;
}
.card-summary {
  font-size: 13px; line-height: 1.6; color: #9ca3af; margin-bottom: 12px;
}
.catalyst-box {
  font-size: 12px; color: #6ee7b7;
  background: #10b98118; border: 1px solid #10b98130;
  border-radius: 8px; padding: 8px 12px; margin-bottom: 12px; line-height: 1.4;
}
.card-footer {
  display: flex; justify-content: space-between; align-items: center;
  border-top: 1px solid #1f2937; padding-top: 10px;
}
.source { font-size: 11px; color: #6b7280; font-style: italic; }

.disclaimer {
  text-align: center; margin-top: 48px;
  font-size: 12px; color: #374151; line-height: 1.7;
}
"""


# ── HTML builder ───────────────────────────────────────────────────────────────

def build_html(data: dict, generated_at: datetime) -> str:
    dips = data.get("dips", [])
    sectors = data.get("sectors", [])
    news = data.get("news", [])

    dips_html = "".join(_dip_card(i) for i in dips)
    sectors_html = "".join(_sector_card(i) for i in sectors)
    news_html = "".join(_news_card(i) for i in news)

    total = len(dips) + len(sectors) + len(news)
    timestamp = generated_at.strftime("%B %d, %Y at %I:%M %p")
    date_str = generated_at.strftime("%Y-%m-%d")

    avg_conviction = (
        f"{sum(int(d.get('conviction', 3)) for d in dips) / len(dips):.1f}"
        if dips else "—"
    )
    avg_stat = (
        f'<div class="stat" style="margin-left:auto;">'
        f'<span class="stat-value" style="color:#f59e0b;">{avg_conviction}</span>'
        f'<span class="stat-label">Avg Dip Conviction</span></div>'
        if dips else ""
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Market Intelligence Screener (Gemini) &mdash; {date_str}</title>
  <style>{CSS}</style>
</head>
<body>
<div class="container">

  <div class="header">
    <div class="header-eyebrow">
      <span style="font-size:24px;">&#128225;</span>
      <h1 class="header-title">Market Intelligence Screener</h1>
    </div>
    <p class="header-meta">Generated {timestamp} &middot; Powered by Gemini + Google Search</p>
  </div>

  <div class="stats-bar">
    <div class="stat">
      <span class="stat-value" style="color:#ef4444;">{len(dips)}</span>
      <span class="stat-label">Dips to Watch</span>
    </div>
    <div class="stat">
      <span class="stat-value" style="color:#10b981;">{len(sectors)}</span>
      <span class="stat-label">Growth Sectors</span>
    </div>
    <div class="stat">
      <span class="stat-value" style="color:#60a5fa;">{len(news)}</span>
      <span class="stat-label">News Items</span>
    </div>
    <div class="stat">
      <span class="stat-value" style="color:#e5e7eb;">{total}</span>
      <span class="stat-label">Total Stories</span>
    </div>
    {avg_stat}
  </div>

  {_section(
      "Dips to Watch",
      "Quality equities trading down &mdash; potential buying opportunities",
      "#ef4444", dips_html, len(dips)
  )}
  {_section(
      "High-Growth Sectors",
      "Sectors with strong forward catalysts and growth signals",
      "#10b981", sectors_html, len(sectors)
  )}
  {_section(
      "Market News",
      "Trending finance &amp; market stories from the past 24&ndash;48 hours",
      "#60a5fa", news_html, len(news)
  )}

  <div class="disclaimer">
    This screener is AI-generated for informational purposes only.<br>
    Not financial advice. Always do your own research before making investment decisions.
  </div>

</div>
</body>
</html>"""


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        print("Error: GOOGLE_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    try:
        data = fetch_market_data(api_key)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    generated_at = datetime.now()
    html = build_html(data, generated_at)

    output = "news-screener.html"
    with open(output, "w", encoding="utf-8") as fh:
        fh.write(html)

    dips = len(data.get("dips", []))
    sectors = len(data.get("sectors", []))
    news = len(data.get("news", []))
    print(f"Done. {output} — {dips} dips, {sectors} sectors, {news} news items")


if __name__ == "__main__":
    main()
