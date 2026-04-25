# KNOWLEDGEBASE.md

Detailed implementation knowledge for generating financial analysis Excel reports. This file allows future Claude instances to skip the planning phase and go straight to coding.

---

## Data Collection Phase

This section documents exactly what data to collect, where to find it, and how to organise it before writing a single line of the script. Complete this phase fully before starting the Excel script — every number in the script must come from verified research, not guesses.

### Data Sources — Where to Find Each Data Point

#### SEC EDGAR (sec.gov/cgi-bin/browse-edgar)

The authoritative source for all US-listed company filings. For Canadian companies cross-listed in the US, use EDGAR for their SEC filings; use SEDAR+ (sedarplus.ca) for Canadian filings.

**Finding a company's filings:**
1. Go to `https://efts.sec.gov/LATEST/search-index?q=%22TICKER%22&dateRange=custom&startdt=2023-01-01&forms=10-K` — or simpler: search `site:sec.gov TICKER 10-K` in a web search.
2. Or navigate directly: `https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=TICKER&type=10-K&dateb=&owner=include&count=10`
3. Click the company name → "Filings" → filter by form type.

**What to pull from each form:**

| Form | What lives here | Where in the document |
|---|---|---|
| **10-K** (annual) | 5-yr income statement, balance sheet, cash flow; segment breakdown; risk factors; business description | Item 1 (Business), Item 1A (Risk Factors), Item 7 (MD&A), Item 8 (Financial Statements) |
| **10-Q** (quarterly) | Most recent quarter's P&L, YTD cash flow, guidance updates | Same structure as 10-K but quarterly; MD&A section has guidance |
| **DEF 14A** (proxy) | Executive pay structure, LTIP design, CEO share ownership, say-on-pay vote | "Executive Compensation" section; "Security Ownership" table |
| **Form 4** (insider trades) | Individual insider buy/sell transactions, date, price, shares, transaction type | The form itself — "Table I" for open-market; "Table II" for derivative; check "Transaction Code" (P=purchase, S=sale, A=grant, D=disposition) |
| **8-K** (current report) | Earnings releases, M&A announcements, guidance updates | Exhibit 99.1 is usually the press release with the financial tables |
| **SC 13G / 13F** | Institutional ownership ≥5%; quarterly fund holdings | 13G = major holder; 13F = fund's full portfolio (quarterly) |

**Searching Form 4s specifically:**
- Go to EDGAR full-text search: `https://efts.sec.gov/LATEST/search-index?q=%22TICKER%22&forms=4&dateRange=custom&startdt=2024-01-01`
- Or search: `SEC EDGAR Form 4 TICKER insider transactions`
- Look for transactions by the CEO, CFO, and board — ignore RSU vesting (code A/D); focus on open-market purchases (P) and sales (S).

---

#### Yahoo Finance

URL pattern: `https://finance.yahoo.com/quote/TICKER/`

Sub-pages and what lives on each:

| Sub-page | URL suffix | Data available |
|---|---|---|
| **Summary** | `/` | Current price, market cap, 52-wk range, P/E, EPS, dividend yield, avg volume |
| **Financials** | `/financials/` | Annual and quarterly income statement (toggle "Annual") — Revenue, Gross Profit, EBITDA, Net Income, EPS |
| **Balance Sheet** | `/balance-sheet/` | Total assets, cash, total debt, equity — annual and quarterly |
| **Cash Flow** | `/cash-flow/` | CFO, Capex, FCF, dividends paid, buybacks — annual and quarterly |
| **Statistics** | `/key-statistics/` | Enterprise value, EV/EBITDA, EV/Revenue, P/FCF, short interest, shares outstanding, float, insider %, institutional % |
| **Analysis** | `/analysis/` | Analyst EPS and revenue estimates, number of analysts, consensus |
| **Profile** | `/profile/` | Business description, sector, industry, key executives list |
| **Holders** | `/holders/` | Top institutional holders with % owned; insider transactions summary |

**How to read the financials pages:** Toggle between "Annual" and "Quarterly" using the button at the top. Numbers are in thousands by default — divide by 1,000 for millions or 1,000,000 for billions. The most recent column is the trailing twelve months (TTM) or the latest fiscal year.

**Getting historical data:** Yahoo Finance only shows 4 years on screen. For 5-year history, use Macrotrends (see below) or download the CSV using the download button on the chart page.

---

#### Macrotrends (macrotrends.net)

Best for clean, pre-formatted 10-year historical financial data. Use when you need more than 4 years of history or want pre-calculated margins and ratios.

URL pattern: `https://www.macrotrends.net/stocks/charts/TICKER/company-name/METRIC`

Common metric slugs:

| Metric | URL slug |
|---|---|
| Revenue | `/revenue` |
| Gross profit margin | `/gross-profit` |
| EBITDA | `/ebitda` |
| Net income | `/net-income` |
| EPS (diluted) | `/eps-diluted` |
| Free cash flow | `/free-cash-flow` |
| Return on equity | `/return-on-equity` |
| Return on assets | `/return-on-assets` |
| Debt / equity ratio | `/debt-equity-ratio` |
| P/E ratio (historical) | `/pe-ratio` |
| EV/EBITDA (historical) | `/ev-ebitda` |

**How to use:** Navigate to the page, set the "Annual" toggle, and the table below the chart gives you year-by-year data. Copy the relevant years. The company name in the URL is usually hyphenated lowercase — easiest to search `macrotrends TICKER revenue` to get the correct URL.

---

#### TipRanks (tipranks.com)

Best for aggregated analyst consensus, individual analyst ratings with price targets, and hedge fund activity.

**How to get analyst data:**
1. Search `TipRanks TICKER analyst ratings` — the result page shows consensus (Buy/Hold/Sell count), average price target, high PT, low PT, and each analyst's name, firm, rating, and PT.
2. Alternatively search `TICKER analyst consensus site:tipranks.com`.
3. The page also shows the "Smart Score" (not needed) and analyst track record (useful for assessing reliability).

**What to capture:** Firm name, analyst name (optional), current rating, current price target, date of most recent rating action. Aim for 5–8 analysts covering the stock.

---

#### Seeking Alpha (seekingalpha.com)

Best for earnings call transcripts and news sentiment.

**Earnings call transcripts:**
- Search: `TICKER earnings call transcript Q4 2025 site:seekingalpha.com`
- Or: `TICKER Q4 2025 earnings transcript`
- The transcript includes management prepared remarks (guidance language, strategic commentary) and Q&A (analyst pushback, risk discussion).

**What to extract from the transcript:**
- Management's exact guidance language for the next quarter/year
- Any mention of macro tailwinds or headwinds
- Capital allocation commentary (M&A pipeline, buyback intentions)
- Any new or changed KPI targets
- Analyst questions that surface risks management didn't volunteer

---

#### Company Investor Relations (IR) Page

Every public company has an IR page, typically at `ir.companyname.com` or `investors.companyname.com`. Search `TICKER investor relations` to find it.

**What lives here:**
- **Earnings releases** (8-K Exhibit 99.1) — press release with financial tables, segment breakdown, and guidance. These are usually cleaner than the 10-Q for current-quarter data.
- **Investor day presentations** (slide decks, PDFs) — the best source for long-term guidance targets and management's own KPI framework.
- **Annual reports** — glossy PDF with business narrative; the 10-K is the legal filing equivalent.
- **SEC filings link** — most IR pages link directly to EDGAR.

**Navigation tip:** Look for a "Financial Results", "Presentations", or "Events & Presentations" tab on the IR page for investor day slide decks.

---

#### Web Search Queries (for filling gaps)

When a specific number isn't on a standard financial data site, use targeted web searches. Effective query patterns:

| Data needed | Search query |
|---|---|
| Segment revenue breakdown | `TICKER segment revenue breakdown FY2025 annual report` |
| Geographic revenue split | `TICKER revenue by geography FY2025 10-K` |
| Analyst price targets | `TICKER analyst price target consensus 2025` |
| CEO ownership stake | `TICKER CEO shares owned proxy 2025` |
| WACC estimate | `TICKER WACC cost of capital estimate` |
| Historical EV/EBITDA | `TICKER EV/EBITDA historical average macrotrends` |
| Peer EV/EBITDA multiples | `[Peer1] [Peer2] [Peer3] EV/EBITDA comparison 2025` |
| Moat / competitive analysis | `TICKER competitive advantages moat analysis` |
| Market themes / tailwinds | `[sector] industry trends tailwinds headwinds 2025 2026` |
| Short interest | `TICKER short interest % float` |
| Institutional ownership | `TICKER institutional ownership % 13F` |
| Order backlog | `TICKER backlog Q4 2025 earnings` |
| Rule of 50 (SaaS) | `TICKER revenue growth FCF margin 2025` |

---

#### SEDI (for Canadian companies, sedarplus.ca)

Equivalent to SEC EDGAR for companies primarily listed on the TSX. Use for Brookfield entities (BN.TO, BAM.TO, BEP.UN, etc.).

- Insider transactions: search the company on SEDAR+, click "Insider Filings" — equivalent to Form 4.
- Annual information form (AIF) = Canadian equivalent of 10-K.
- Management Information Circular = Canadian equivalent of proxy (DEF 14A).

---

### Template Section → Source Mapping

Quick reference: which source answers each data template section.

| Template Section | Primary Source | Secondary / Fallback |
|---|---|---|
| A. Company Identity | Yahoo Finance `/profile/` or company IR page | 10-K Item 1 (Business) |
| B. Market Snapshot | Yahoo Finance `/` (summary page) + `/key-statistics/` | Google Finance |
| C. Income Statement | Yahoo Finance `/financials/` (annual) | Macrotrends; 10-K Item 8 |
| D. Balance Sheet | Yahoo Finance `/balance-sheet/` (annual) | 10-K Item 8 |
| E. Cash Flow | Yahoo Finance `/cash-flow/` (annual) | 10-K Item 8 |
| F. Return on Capital | Macrotrends `/return-on-equity`, `/return-on-assets` | Calculate from IS + BS data |
| G. Segment Detail | 10-K Item 1 + Item 8 segment footnotes; earnings release | Investor day slide deck |
| H. Geographic Revenue | 10-K Item 1 or geographic segment footnote | Earnings release |
| I. Management + Proxy | DEF 14A on EDGAR (proxy statement) | Company IR → Governance |
| I. Insider Transactions | SEC EDGAR Form 4 filings | Yahoo Finance `/holders/` |
| J. Moat / Competitors | 10-K Item 1 (Competition section) + web search | Seeking Alpha analysis pieces |
| K. Risks | 10-K Item 1A (Risk Factors) + earnings call transcript | Web search for sector risks |
| L. Valuation — Multiples | Yahoo Finance `/key-statistics/` | Macrotrends historical ratios |
| L. Valuation — Peers | Yahoo Finance for each peer ticker | Macrotrends peer comparison |
| L. Valuation — NAV/DCF | Analyst reports + company IR (asset values) | Calculate from segment data |
| M. Analyst Ratings | TipRanks | Seeking Alpha analyst page |
| M. Market Themes | Earnings call transcript + sector news search | Seeking Alpha sector articles |
| M. Ownership / Short | Yahoo Finance `/holders/` + `/key-statistics/` | EDGAR 13F/13G filings |
| N. Guidance & Targets | Earnings release + earnings call transcript | Investor day presentation |

---

### Data Collection Template (fill before scripting)

This is the master data structure to populate during research. Every section maps directly to a tab in the Excel report.

#### A. Company Identity
```
Company full name:
Ticker (exchange):
Sector / sub-sector:
Headquarters:
Founded / incorporated:
CEO name + tenure:
CFO name:
Fiscal year end (calendar Dec vs. Sep vs. other):
Report date (today's date):
```

#### B. Market Snapshot (as of report date)
```
Current stock price:             $___  (date: ___)
Market capitalisation:           $___B
Enterprise value:                $___B  (Mkt Cap + Net Debt)
52-week high:                    $___
52-week low:                     $___
YTD performance:                 ___%
Shares outstanding (diluted):    ___M / ___B
Dividend per share (annual):     $___
Dividend yield:                  ___%
Analyst consensus:               Buy / Hold / Sell  (__ Buy, __ Hold, __ Sell)
Avg analyst price target:        $___
High analyst price target:       $___
Short interest (% of float):     ___%
```

#### C. Income Statement — 5 Years (USD $M or $B, be consistent)

Columns: FY2021, FY2022, FY2023, FY2024, FY2025 (or FY2025E), FY2026E guidance

```
Revenue (total):
  Segment 1 revenue:
  Segment 2 revenue:
  (add segments as needed)
Gross profit:
Gross margin %:
EBITDA:
EBITDA margin %:
EBIT / Operating income:
EBIT margin %:
Net income (GAAP):
EPS (GAAP, diluted):
Adj. EPS (non-GAAP, if company reports):
YoY Revenue growth %:

For alt-asset managers, additionally:
  Fee-Related Earnings (FRE):
  Distributable Earnings (DE) by segment:
  Fee-Bearing Capital (FBC):
  Carried interest realised:
```

#### D. Balance Sheet — 3 Years (most recent + 2 prior)

Columns: FY2022 (or FY2023), FY2023 (or FY2024), FY2024 (or FY2025)

```
Cash & equivalents:
Accounts receivable:
Inventories:
Total current assets:
Goodwill:
Intangible assets:
PP&E (net):
Total assets:

Accounts payable:
Short-term debt:
Total current liabilities:
Long-term debt:
Total liabilities:

Total shareholders' equity:
Book value per share:

Derived:
  Net debt (Total Debt - Cash):
  Net Debt / EBITDA:
  Debt / Equity ratio:
  Current ratio:
```

#### E. Cash Flow Statement — 5 Years

```
Cash from operations (CFO):
Capital expenditures (Capex):
Free cash flow (FCF = CFO - Capex):
  FCF margin %:
Asset disposals / divestitures:
Acquisitions paid:
Dividends paid:
Share buybacks / repurchases:
Net change in debt:
```

#### F. Return on Capital — 5 Years

```
ROE (Net Income / Avg Equity):
ROA (Net Income / Avg Assets):
ROIC (NOPAT / Invested Capital):
WACC (estimate):
Economic spread (ROIC - WACC):

For asset-light businesses (alt managers, SaaS):
  Asset management FRE ROE:
  Incremental capital deployed ($):
  Incremental DE/earnings on new capital:
  Incremental ROIC:

Third-party LP returns (if available):
  Fund strategy 1 net IRR:
  Fund strategy 2 net IRR:
  30-year or long-term track record CAGR:
```

#### G. Segment Detail

For each reporting segment:
```
Segment name:
FY2024 revenue + % of total:
FY2025 revenue + % of total + YoY growth:
Adj EBIT margin:
Key products / services:
Key clients / customer types:
Primary geographies:
Value proposition (why customers choose this):
Buying process (how customers procure):
Seasonality pattern:
Revenue model (product / recurring / project / spread):
Moat elements:
```

#### H. Geographic Revenue Split

```
Geography 1 (e.g. North America): $___  ___%
Geography 2 (e.g. Europe):        $___  ___%
Geography 3 (e.g. APAC):          $___  ___%
(etc.)
Notes on FX exposure or natural hedges:
```

#### I. Management

```
For each key executive (CEO, CFO, key division heads):
  Name:
  Role:
  Tenure (joined + in role since):
  Key contributions or track record:

Proxy statement (DEF 14A):
  LTIP structure (vesting period, performance metrics used):
  Performance Share Units tied to (TSR vs peers? EPS? Cash flow?):
  CEO base salary: $___
  CEO total comp (last proxy): $___M
  CEO share ownership: ___% of company (~$___M at current price)
  Key executives' deferred comp: held in stock? cash?

Insider transactions (Form 4, last 24 months):
  [Date]: [Name] [bought/sold] [shares] at $[price] — [open market / 10b5-1 plan / RSU vest]
  (list all material transactions)
  Net insider verdict: BULLISH / NEUTRAL / BEARISH

Capital allocation decisions (last 3–5 years):
  [Year]: [Decision] — [Outcome] — [Assessment: EXCELLENT / POSITIVE / NEUTRAL / WARNING]
```

#### J. Competitive Moat

```
Moat rating (overall): Wide / Narrow / None

For each moat element:
  Element name:
  Width (Wide / Moderate / Narrow):
  Evidence / quantitative support:

Primary competitors:
  Competitor 1: [Name, Ticker, relative size, key diff]
  Competitor 2: ...

Competitive comparison (vs each main peer):
  [Peer] vs [Company]: [Where company leads / lags]
```

#### K. Risks

For each risk category (Financial, Operational, Macro, Regulatory):
```
Risk name:
Probability (HIGH / MEDIUM / LOW):
Impact (HIGH / MEDIUM / LOW):
Mitigation in place:
Net risk after mitigation (HIGH / MEDIUM / LOW / LOW-MED):
```

#### L. Valuation Inputs

```
Current trading multiples:
  P/E (GAAP):
  P/Adj EPS or P/DE:
  EV/EBITDA:
  EV/Revenue:
  Price/FCF:
  Price/Book:

Historical average multiples (5yr avg if available):
  Historical avg P/E:
  Historical avg EV/EBITDA:

Peer multiples comparison:
  [Peer 1]: Ticker, P/E, EV/EBITDA, AUM or Revenue, comments
  [Peer 2]: ...
  (3–5 peers)
  Peer median P/E:         ___x
  Peer median EV/EBITDA:   ___x
  Premium / discount to peers: ___%

NAV / Sum-of-Parts (for conglomerates, alt managers, holding cos):
  Asset 1 name:  method used:  estimated value $___B  per share $___
  Asset 2 name:  ...
  Less: Corporate debt: -$___B
  Total NAV estimate:  $___B  ($___/share)
  Current price vs NAV:  ___% discount or premium

DCF / Scenario analysis:
  Bear case: revenue/DE CAGR ___%, terminal multiple ___x, discount rate ___% → implied $___/share
  Base case: revenue/DE CAGR ___%, terminal multiple ___x, discount rate ___% → implied $___/share
  Bull case: revenue/DE CAGR ___%, terminal multiple ___x, discount rate ___% → implied $___/share
  Current price: $___
```

#### M. Market Sentiment

```
Analyst ratings (list individually):
  [Firm]: [Rating] [Price Target] [Commentary / recent action]
  (aim for 5–8 firms)

Market themes:
  [Theme]: TAILWIND / HEADWIND / CATALYST — [1-sentence explanation]
  (aim for 6–10 themes)

Ownership structure:
  Institutional:   ___%
  Insider:         ___%
  Retail/Other:    ___%

Short interest:
  Short % of float:  ___%
  Days to cover:     ___
```

#### N. Key Indicators & Management Targets

```
KPI scorecard (3 most recent years):
  [KPI name]: [FY2023] [FY2024] [FY2025] [Trend: POSITIVE / NEUTRAL / WATCH]
  (list 12–18 KPIs relevant to the business model)

Management guidance and targets:
  Near-term (FY2026): Revenue $___  EPS $___  (any other guided metrics)
  Medium-term (2–3yr): [metric] → [target]
  Long-term (5yr): [metric] → [target]
```

#### SaaS / High-Growth Tech Only

```
ARR / MRR:
Net Revenue Retention (NRR):
Customer count + growth:
Gross margin:
FCF margin:
Rule of 50 score (Revenue growth % + FCF margin %):
  Verdict: PASS (≥50) / BORDERLINE (40–49) / FAIL (<40)
CAC payback period:
Magic number (sales efficiency):
```

---

### Data Collection Workflow

Follow this order — each step builds on the last.

1. **10-K on SEC EDGAR** — Fills templates C, D, E, G, H, J, K. The ground truth for all financial figures and business description. Navigate: `sec.gov → search TICKER → filter by 10-K → most recent filing → Item 1, 1A, 7, 8`.

2. **Most recent 10-Q on SEC EDGAR** — Updates template C and N with the latest quarter and any guidance revision since the 10-K. Same navigation, filter by 10-Q.

3. **DEF 14A (Proxy) on SEC EDGAR** — Fills template I (management incentives, CEO ownership). Filter filings by "DEF 14A". Go to "Executive Compensation" and "Security Ownership" sections.

4. **Form 4 filings on SEC EDGAR** — Fills template I (insider transactions). Filter by Form 4, date range last 24 months. Focus on Transaction Code P (open-market buy) and S (open-market sale); ignore A/D (grants/vesting).

5. **Yahoo Finance summary + key-statistics** — Fills template B (market snapshot). Visit `finance.yahoo.com/quote/TICKER/` for price/cap/52-wk, then `/key-statistics/` for EV, short interest, shares outstanding, institutional %.

6. **Yahoo Finance financials** — Cross-check template C, D, E figures against what was pulled from the 10-K. Use `/financials/`, `/balance-sheet/`, `/cash-flow/` with the "Annual" toggle. Use Macrotrends as a fallback for 5-year history if Yahoo only shows 4.

7. **TipRanks** — Fills template M (analyst ratings). Search `TICKER analyst ratings tipranks` — capture firm, rating, and price target for each covering analyst (5–8 firms).

8. **Earnings call transcript on Seeking Alpha** — Fills templates K and N (risks, guidance, market themes). Search `TICKER Q4 2025 earnings call transcript seekingalpha`. Read the prepared remarks for guidance language and the Q&A for risks.

9. **Company IR page** — Fills template N gaps (long-term targets). Find the investor day slide deck if one exists in the last 18 months — this is the richest source for multi-year KPI targets.

10. **Peer research** — Fills template L (peer multiples). For each of 3–5 competitors, visit their Yahoo Finance `/key-statistics/` page and record P/E, EV/EBITDA. Use Macrotrends for historical average multiples of the subject company.

---

### Data Quality Rules

- All prices and market data must be as of a specific date — record it and use it consistently throughout the report.
- If a number is estimated or derived (not directly disclosed), mark it with `~` or add a NOTE row explaining the source.
- If a segment metric is not separately disclosed, note `"Not separately reported"` rather than guessing.
- GAAP vs non-GAAP: always show both where material. Explain the difference in a NOTE row.
- 5-year history is standard for income statement and cash flow; 3-year is sufficient for balance sheet.
- Fiscal years that end off-calendar (Sep, Jun, etc.) — label clearly as FY2025 not just "2025" to avoid confusion.
- For non-US companies with dual-listed tickers, note both exchanges and the reporting currency.

---

## Script Skeleton

```python
"""
<COMPANY NAME> (<TICKER>) - Comprehensive Financial Analysis
Generated: <Month Year>
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

OUTPUT_PATH = "/Users/naji/WORK/github.com/AI/claude/Agent/MarketResearch/output/<TICKER>_Financial_Analysis.xlsx"

# ── Colour palette ────────────────────────────────────────────────────────────
# Pick a thematically appropriate palette per company (see "Colour Palettes" below)
DARK_PRIMARY  = "1F3864"   # header backgrounds (darkest)
MID_PRIMARY   = "2E5FAC"   # section sub-headers
LIGHT_PRIMARY = "BDD7EE"   # totals / highlighted rows
ACCENT        = "C9A84C"   # accent / gold highlight
GREEN         = "375623"
LIGHT_GREEN   = "E2EFDA"
RED           = "C00000"
LIGHT_RED     = "FFCCCC"
GREY          = "F2F2F2"   # alternating row (even)
WHITE         = "FFFFFF"
DARK_GREY     = "595959"

FONT_SIZE = 14

# ── Style helpers ─────────────────────────────────────────────────────────────
def font(bold=False, size=FONT_SIZE, color=None, italic=False):
    kw = {"name": "Calibri", "bold": bold, "size": size, "italic": italic}
    if color: kw["color"] = color
    return Font(**kw)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def center(wrap=True):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def left(wrap=True):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

def set_col_width(ws, col, width):   # col is integer
    ws.column_dimensions[get_column_letter(col)].width = width

def merge_write(ws, cell_range, value, bold=False, size=FONT_SIZE,
                bg=None, fg=WHITE, align=None, italic=False):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = value
    cell.font = font(bold=bold, size=size, color=fg, italic=italic)
    if bg: cell.fill = fill(bg)
    cell.alignment = align or center()
    return cell

def write_cell(ws, row, col, value, bold=False, size=FONT_SIZE,
               bg=None, fg=None, align=None, num_fmt=None, italic=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font(bold=bold, size=size,
                     color=fg or ("000000" if not bg else WHITE),
                     italic=italic)
    if bg: cell.fill = fill(bg)
    cell.alignment = align or left()
    if num_fmt: cell.number_format = num_fmt
    return cell

def header_row(ws, row, col_values, bg=DARK_PRIMARY, fg=WHITE, height=30):
    ws.row_dimensions[row].height = height
    for col, val in enumerate(col_values, 1):
        write_cell(ws, row, col, val, bold=True, bg=bg, fg=fg,
                   align=center(), size=FONT_SIZE)

def section_header(ws, row, col_from, col_to, title, bg=MID_PRIMARY, fg=WHITE):
    col_range = f"{get_column_letter(col_from)}{row}:{get_column_letter(col_to)}{row}"
    merge_write(ws, col_range, title, bold=True, bg=bg, fg=fg,
                size=FONT_SIZE + 1, align=center())
    ws.row_dimensions[row].height = 28
```

---

## Tab-by-Tab Implementation Blueprint

### TAB 1 — Cover

**Purpose:** One-page snapshot investors read first.

**Structure:**
- Rows 1–7: Dark banner with company name + "Comprehensive Financial Analysis", merged A:E, font size 26
- Row 8: Accent-colour bar with `"Prepared: <Month Year>  |  Analyst: Research Mode"`
- Rows 10–27: Two-column key-facts table — label (col A, MID_PRIMARY bg) | value (cols B:E merged, LIGHT_PRIMARY bg). Typical fields:
  - Company, Ticker, Sector, Headquarters, Founded, CEO, Report Date
  - Current Price, Market Cap, 52-Week Range, Analyst Consensus
  - Key metric #1, Key metric #2, Dividend Yield
- Row 30: `section_header(..., "INVESTMENT THESIS", bg=GREEN)`
- Rows 31–37: Large merged text block with the thesis paragraph (wrap=True), LIGHT_PRIMARY bg, row height ~130
- Row 39: Final rating bar — `"RATING:  BUY/HOLD/SELL  |  12-Month Target: $XX  |  Upside: ~XX%"`, GREEN bg, font size+2

**Config:** `ws.sheet_view.showGridLines = False`. 5 columns at width 28.

---

### TAB 2 — Business Overview

**Sections (in order):**
1. **Three Core Business Segments** — table: Segment | Revenue/DE ($B) | YoY Growth | Key Metric | % of Total | Description. Total row in LIGHT_PRIMARY.
2. **Products & Services** — table: Category | Products/Vehicles | Strategy | Target Client | Geography | Fee Type
3. **Revenue by Geography** — table: Geography | % share | Key Markets (cols C:F merged)
4. **Value Proposition & Key Clients** — two-col: bold label | merged description (cols B:F)
5. **Seasonality & Margin Structure** — table: Metric | Range | Comment (cols C:F merged)

**Config:** 6 columns, widths approx `[32, 22, 22, 22, 22, 28]`. Alternating GREY/WHITE rows, height 22–30.

---

### TAB 3 — Moat

**Sections:** Group moat elements into 4–5 themed blocks (e.g. Scale, Brand, Capital Structure, Track Record, Competitive Comparison). Each block:
- `section_header` in MID_PRIMARY
- `header_row`: Moat Element | Moat Width | Evidence / Commentary | (filler)
- Data rows: col 1 bold, col 2 colour-coded badge (GREEN=Wide, ACCENT_GOLD=Moderate, RED=Narrow), cols 3:4 merged description

End with `section_header(..., "MOAT VERDICT", bg=GREEN)` followed by a large merged text verdict cell in LIGHT_GREEN.

**Config:** 4 columns, widths `[30, 14, 14, 32]`. Row height 26.

---

### TAB 4 — Income Statement

**Sections:**
1. **GAAP Revenue** (5yr table) — Revenue ($B), YoY Growth %. Include a "NOTE:" row in yellow (`"FFF2CC"` bg, italic, merged) when GAAP metrics are misleading.
2. **Management Earnings View** (e.g. Distributable Earnings, or EBITDA by segment) — 5yr table, subtotal rows in LIGHT_PRIMARY bold.
3. **Key Profitability Metrics** — P/E, P/DE or P/EBITDA, EPS, key segment KPIs

**Config:** 6 columns, widths `[36, 16, 16, 16, 16, 16]`. Year columns header: `["Metric", "2021", "2022", "2023", "2024", "2025E"]`.

---

### TAB 5 — Balance Sheet

**Sections:**
1. Consolidated balance sheet table — ASSETS section header, line items, TOTAL ASSETS in LIGHT_PRIMARY; LIABILITIES section, TOTAL LIABILITIES; EQUITY section
2. LEVERAGE METRICS section — Debt/Assets, Net Debt/EBITDA, avg maturity
3. **Key Observations** — 3–5 rows: bold label with colour-coded bg (GREEN=positive, ACCENT_GOLD=caution) | merged description

**Config:** 5 columns, widths `[36, 18, 18, 18, 18]`. Years: 3 historical + Change column.

---

### TAB 6 — Cash Flow Analysis

**Sections:**
1. **Operating Cash Flows** — CFO, YoY growth
2. **Capex & Investment Activity** — Capex, Disposals/Recycling, Net Capex, GAAP FCF (note row explaining negative FCF if applicable)
3. **Management/Distributable Free Cash Flow** — segment DE breakdown, totals in LIGHT_PRIMARY. Negative values → LIGHT_RED cell bg.
4. **Capital Allocation** — Dividends, Buybacks, Investments, Retained
5. **Cash Flow Quality Assessment** — colour-coded verdict rows (GREEN=POSITIVE, ACCENT_GOLD=WARNING)

**Config:** 6 columns. NOTE rows: merged A:F, italic, `"FFF2CC"` bg.

---

### TAB 7 — Return on Capital

**Sections:**
1. **GAAP Returns** — ROE, ROA, ROIC (include NOTE row explaining suppression by depreciation)
2. **Management/Segment Returns** — asset-light business ROE, segment ROIC
3. **Incremental ROIC** — New capital deployed, DE on new capital, ROIC, WACC, Spread (ROIC - WACC)
4. **Third-Party/LP Fund Returns** — by strategy: IRR % across years; 30-yr track record
5. **Return on Capital Verdict** — large merged GREEN/LIGHT_GREEN cell, numbered conclusions

**Config:** 6 columns, widths `[34, 16, 16, 16, 16, 20]`.

---

### TAB 8 — Management

**Sections:**
1. **Key Executives** — Name | Role | Tenure | Key Contribution (row height 28)
2. **Executive Incentives (Proxy Statement)** — Incentive Element | Structure | Alignment (colour-coded badge: HIGH=GREEN, MODERATE=ACCENT_GOLD) | Commentary
3. **Capital Allocation Track Record** — Decision | Action Taken | Outcome | Assessment (EXCELLENT=GREEN, POSITIVE=MID_PRIMARY, NEUTRAL=ACCENT_GOLD, WARNING=RED)
4. **Insider Activity (SEC Form 4)** — Period | Activity | Volume | Signal (BULLISH=GREEN, NEUTRAL=ACCENT_GOLD)
5. **Management Verdict** — merged large cell `"DOES MANAGEMENT ACT LIKE AN OWNER?"`, GREEN header, LIGHT_GREEN content, row height ~140

**Config:** 4 columns, widths `[30, 18, 16, 32]`.

---

### TAB 9 — Risks

**Structure:** Group risks into 4 categories (Financial, Operational, Macro/Market, Regulatory). For each category:
- `section_header` in DARK_GREY
- `header_row`: Risk Factor | Probability | Impact | Mitigation | Net Risk
- Data rows: col 1 bold; cols 2, 3, 5 colour-coded badge using:
  ```python
  color_map = {"HIGH": RED, "MEDIUM": ACCENT_GOLD, "LOW": GREEN, "LOW-MED": "92D050", "N/A": GREY}
  ```
  (white text on badge); col 4 plain description

End with **Overall Risk Verdict** — merged cell in LIGHT_BLUE.

**Config:** 5 columns, widths `[30, 14, 14, 30, 16]`. Row height 28.

---

### TAB 10 — Valuation

**Sections:**
1. **Current Market Snapshot** — 3-col table: Metric (bold) | Value | Note (cols C:E merged). Metrics: Price, Mkt Cap, Shares, P/E, P/DE or EV/EBITDA, Price/Book, Dividend Yield.
2. **Peer Valuation Comparison** — Company | Ticker | P/DE (or EV/EBITDA) | AUM/Revenue | Comments. Highlight subject company in LIGHT_BLUE; implied fair value row in LIGHT_GREEN.
3. **NAV / Sum-of-Parts Analysis** (if applicable) — Asset Component | Value ($B) | Method | Per Share. Total row in LIGHT_BLUE; "Less: Debt" in LIGHT_RED; Discount row in LIGHT_RED.
4. **DCF / Scenario Analysis** — Bear/Base/Bull: key assumption | terminal multiple | discount rate | implied value. Row colours: LIGHT_RED / LIGHT_BLUE / LIGHT_GREEN / GREY.
5. **Margin of Safety Assessment** — large merged GREEN-header + LIGHT_GREEN content cell. Numbered conclusions + rating, row height ~130.

**Config:** 5 columns, widths `[32, 18, 18, 18, 20]`.

---

### TAB 11 — Market Sentiment

**Sections:**
1. **Analyst Consensus** — Rating (Buy/Hold/Sell) | Count | % of Coverage | Avg Price Target. Consensus row in LIGHT_GREEN bold.
2. **Recent Analyst Actions** — Firm | Rating (colour-coded badge: GREEN=Buy/Outperform/Overweight, ACCENT_GOLD=Hold) | Price Target | Commentary
3. **Market Themes & Tailwinds/Headwinds** — Theme | Type (TAILWIND=GREEN, HEADWIND=RED, CATALYST=MID_PRIMARY) | Detail (cols C:D merged)
4. **Ownership & Short Interest** — Category | % | Notes (cols C:D merged)

**Config:** 4 columns, widths `[32, 18, 18, 28]`.

---

### TAB 12 — Key Indicators

**Sections:**
1. **Financial Scorecard** — KPI | 3 years | Trend/Signal. Signal cell colour-coded: GREEN=POSITIVE, RED=WATCH/NEGATIVE, ACCENT_GOLD=NEUTRAL.
2. **Operating Metrics & Growth Targets** — Metric | Current | Near-term Target | Long-term Target | Management Guidance (cols 1 and 5 left-aligned, rest centered)
3. **Investment Summary & Final Recommendation** — Large merged cell (row height ~210), GREEN header "INVESTMENT SUMMARY & FINAL RECOMMENDATION", LIGHT_GREEN body. Structure: FINAL RECOMMENDATION, thesis bullet points, price/upside, risks to watch, catalysts.

**SaaS companies only:** Add a "Rule of 50" section between sections 1 and 2:
```
Revenue Growth (%) + FCF Margin (%) = Rule of 50 Score
Score ≥ 50 → PASS (GREEN), < 40 → FAIL (RED), 40-49 → BORDERLINE (ACCENT_GOLD)
```

**Config:** 5 columns, widths `[32, 18, 18, 18, 22]`.

---

## `main()` Function Pattern

```python
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # remove default sheet

    build_cover(wb)
    build_business(wb)
    build_moat(wb)
    build_income(wb)
    build_balance_sheet(wb)
    build_cashflow(wb)
    build_return_on_capital(wb)
    build_management(wb)
    build_risks(wb)
    build_valuation(wb)
    build_sentiment(wb)
    build_key_indicators(wb)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"Saved to: {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
```

---

## Styling Conventions

| Pattern | Rule |
|---|---|
| All cells | `FONT_SIZE = 14`, Calibri |
| `showGridLines` | `False` on every sheet |
| Section headers | `section_header()` — MID_PRIMARY bg, height 28, font size 15 |
| Tab top banner | `section_header()` with DARK_PRIMARY bg |
| Column headers | `header_row()` — DARK_PRIMARY bg, height 30, bold white |
| Alternating rows | `GREY if i % 2 == 0 else WHITE`, height 22 |
| Totals / subtotals | LIGHT_PRIMARY bg, bold |
| NOTE / warning rows | `"FFF2CC"` (yellow) bg, italic, merged full width |
| Negative numbers | LIGHT_RED cell bg |
| Verdict blocks | GREEN section header + LIGHT_GREEN merged content, height 90–210 |
| Badge cells | Colour-coded bg + WHITE bold text (used for rating, moat width, risk level, signal) |

---

## Colour Palettes — By Company Type

Use a thematically matched palette rather than copy-pasting BN's exact colours:

| Company Type | Dark Primary | Mid Primary | Light Primary | Accent |
|---|---|---|---|---|
| Financials / Alt-Asset Mgmt | `1F3864` Navy | `2E5FAC` Blue | `BDD7EE` Light Blue | `C9A84C` Gold |
| Industrials / Engineering | `1B2A4A` Dark Navy | `2B5FA6` Blue | `D6E4F0` Pale Blue | `E87722` Orange |
| Environmental / Waste | `1B4332` Dark Green | `2D6A4F` Mid Green | `D8F3DC` Light Green | `B7950B` Gold |
| Technology / SaaS | `1A1A2E` Dark Navy | `16213E` Midnight | `E8F4FD` Pale Blue | `4CC9F0` Cyan |
| Consumer / Retail | `4A1942` Dark Purple | `7B2D8B` Purple | `F3E5F5` Lavender | `FF8F00` Amber |

Always keep RED=`C00000`, LIGHT_RED=`FFCCCC`, GREY=`F2F2F2`, WHITE=`FFFFFF`.

---

## Data Research Checklist (per ticker)

Before writing any data into the script, pull and verify:

- [ ] Current stock price, market cap, 52-week range
- [ ] Most recent 5 years of: Revenue, EBITDA/DE/FCF, EPS
- [ ] Balance sheet: total assets, total debt, cash, equity (latest 3 years)
- [ ] Operating cash flow and capex (latest 5 years)
- [ ] Analyst consensus: Buy/Hold/Sell counts, average PT, high PT
- [ ] 3–5 recent analyst firm ratings with price targets
- [ ] Key executives: CEO, CFO, relevant division heads + tenure
- [ ] Proxy statement: LTIP structure, CEO ownership/compensation
- [ ] SEC Form 4 / SEDI: recent insider buys/sells (last 12–24 months)
- [ ] Segment revenue/earnings breakdown + YoY growth
- [ ] 3–5 primary competitors + peer valuation multiples
- [ ] Company-stated growth targets / guidance (investor day, earnings calls)
- [ ] Current macro tailwinds / headwinds relevant to the sector

---

## Common Pitfalls to Avoid

- **`merge_cells` then write**: always write to the top-left cell of the merged range, not the merged range itself.
- **Column width by integer**: use `get_column_letter(col)` — `ws.column_dimensions["A"].width`, not `ws.column_dimensions[1].width`.
- **Row height on merged cells**: set height on the first row of the merge; other rows in the merge can be set to 0 or left default.
- **`fg` defaults**: `write_cell` sets fg to WHITE when `bg` is provided — pass `fg="000000"` explicitly for dark-on-light cells.
- **GAAP opacity note**: for companies where GAAP metrics are misleading (infrastructure, alt-asset managers), always include a NOTE row explaining the correct lens (DE, Distributable FCF, etc.).
