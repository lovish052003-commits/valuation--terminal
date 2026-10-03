# Universal Valuation Automation

## Setup (one time, in your Antigravity terminal)

    pip install requests openpyxl
    export ANTHROPIC_API_KEY=sk-ant-...          # your key

## Use, for any company

    python generate_valuation.py "RBL Bank"
    python generate_valuation.py "Sun Pharmaceutical Industries"
    python generate_valuation.py "HDFC Bank" --out ./models

Each run:
1. Sends VALUATION_PROMPT.md + a strict JSON schema to Claude with web search
   enabled, so it researches the company itself — no manual Screener.in
   copy/paste.
2. Runs the audit checks that came out of the RBL Bank and Sun Pharma model
   reviews (self-inclusion in peer set, Tax==PBT data corruption, stale share
   count, reinvestment-rate clamp) BEFORE computing anything — printed to the
   terminal as an audit log, and saved as a sheet in the output workbook.
3. Branches to an Excess Return Model (banks/NBFCs/insurers) or a WACC/FCFF
   DCF (everyone else) — decided by Claude's classification of the company,
   which you can see in the audit log's first line.
4. Writes `<Company>_Valuation_Model.xlsx` with Data / Peers & WACC /
   Valuation / Comps / Audit Log sheets.

## Files

- `VALUATION_PROMPT.md` — the methodology instructions. Edit this to change
  how any future company gets valued (e.g. tweak the reinvestment-rate
  formula, add a new audit check, change the peer-selection rule) — every
  company you run after that picks up the change automatically.
- `generate_valuation.py` — the script. The audit_checks() and compute_*()
  functions are where the actual math lives; the prompt only controls what
  data Claude fetches, not how it's turned into a valuation.

## Extending it

- New audit check: add a `check_*(data, log)` function and call it from
  `run_audit()`.
- New company type (e.g. Insurance gets its own model instead of sharing the
  bank path): add a branch in `main()` and a `compute_*()` function.
- Batch mode: wrap `main()`'s body in a loop over a list of company names if
  you want to regenerate your whole coverage list in one run.
