# Phase 2: Execute Analysis Scripts

Research and script writing are already complete. Your only job is to run the scripts and verify their output.

## Steps (per ticker listed in "Scripts to Execute")

1. Run `python tmp/<ticker_lower>_analysis.py` from this directory
2. Confirm `output/<TICKER>_Financial_Analysis.xlsx` was created and is non-empty
3. If the script raises a runtime error (import, path, syntax), fix only that error and re-run
4. Report pass / fail for each ticker

## Scope constraints

- Do NOT fetch any web pages or perform additional research
- Do NOT change financial figures, tab structure, or script logic beyond fixing runtime errors
- If a script has a fundamental data error that requires re-research, log it and move on — do not restart Phase 1 yourself

## Completion

When all scripts have been executed (succeeded or explicitly skipped), stop. The agent resets the task queue automatically.
