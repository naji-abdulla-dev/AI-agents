# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

This is a financial research agent. Given a stock ticker in `work.md`, perform a deep-dive analysis following the methodology in `instructions.md` and produce a styled Excel workbook.

## Workflow

1. Read the ticker from `work.md`.
2. Pull live data and online sources (current trends, market sentiment, competitive analysis) — research mode.
3. Write a self-contained Python script to `tmp/<ticker_lower>_analysis.py`.
4. Run the script; it saves the Excel file to `output/<TICKER>_Financial_Analysis.xlsx`.
5. Open and verify the output is readable (fonts render, no truncated text, all tabs populated).

## Output Requirements

The Excel workbook must contain these tabs in order:

1. Cover
2. Business Overview
3. Moat
4. Income Statements
5. Balance Sheet
6. Cash Flow Analysis
7. Return on Capital
8. Management
9. Risks
10. Valuation
11. Market Sentiment
12. Key Indicators

For SaaS companies, apply Rule of 50 (Revenue Growth % + FCF Margin %) in the Key Indicators tab.

## Detailed Implementation Reference

See `KNOWLEDGEBASE.md` for the complete tab-by-tab blueprint, full script skeleton, helper functions, styling conventions, colour palettes, data research checklist, and common pitfalls. Use it to write the script directly without a planning phase.

## Script Conventions

- Use `openpyxl` for all Excel generation (already the established pattern).
- Global `FONT_SIZE = 14` — apply to every cell.
- Define a colour palette as module-level constants at the top of each script.
- Helper functions (`_font`, `_fill`, `_border`, `_align`, `_hdr_cell`, `_data_cell`) keep cell styling DRY — follow the pattern from existing scripts in `tmp/`.
- Hard-code the output path relative to the repo root: `output/<TICKER>_Financial_Analysis.xlsx`.
- Each tab is built in its own function; `main()` assembles and saves the workbook.

## Analysis Depth

Think like a company owner, not a trader:
- **Business**: products/services, revenue by segment/geography/channel, key customers, buying process, seasonality, moat.
- **Financials**: income statement trends, balance sheet health, free cash flow quality and conversion.
- **Management**: proxy statement incentives, capital allocation track record, insider buys/sells (SEC filings), whether the CEO acts like an owner.
- **Valuation**: intrinsic value mechanics (DCF or appropriate method), margin of safety at current price.
- **Risks**: company-specific and macro risks that could impair the thesis.
