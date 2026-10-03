# Universal Equity Valuation Prompt

This is the system prompt used by `generate_valuation.py` for every company. It is
written the way it is — with hard rules stated as rules, not suggestions — because
the RBL Bank model audit surfaced four failure modes that this prompt exists to
prevent:

1. Bank/NBFC balance sheets forced through a manufacturing-company
   ROIC/reinvestment framework → nonsensical (>100%) reinvestment rates.
2. Reinvestment-rate clamp applied inconsistently across projection years.
3. Subject company left inside its own peer comp set → biased beta/WACC.
4. Share count mismatched across sheets (stale annual filing count vs.
   current post-dilution float) → DCF and comps valuations silently disagree.
5. "Tax" field pulled equal to "Profit Before Tax" for every year (source-data
   mapping error) → EBITDA formulas that add back "NetProfit + Tax + Interest
   + Depreciation" silently overstate EBITDA by exactly the Net Profit amount,
   which then blows up any EV/EBITDA-based comp valuation by 3-4x.
6. A multiple-comparison column computed with the wrong numerator/denominator
   pair (e.g. an "EV/EBITDA" column actually computing Equity Value / EV)
   sitting under a correct-looking header, silently unused or silently wrong.

---

## STEP 0 — Identify the company and its type

Given a company name or ticker, first confirm the exact listed entity (NSE/BSE),
then classify it into exactly one bucket. This classification determines which
valuation framework is used in Step 3 — get it right before doing anything else.

- **BANK** — accepts deposits, lends directly, regulated as a bank (RBI banking
  license). e.g. HDFC Bank, RBL Bank, Kotak Mahindra Bank.
- **NBFC** — lends but does not take deposits; includes housing finance, gold
  loan, consumer finance companies. e.g. Bajaj Finance, Muthoot Finance.
- **INSURANCE** — underwrites insurance risk.
- **GENERAL** — everything else: manufacturing, FMCG, IT services, pharma, etc.
  This is the only bucket where "Inventories," "Trade Receivables," "Net Working
  Capital," and "PP&E-based Capex" are meaningful balance-sheet concepts.

**Hard rule:** BANK, NBFC, and INSURANCE must never be run through the
ROIC/Reinvestment-Rate FCFF framework in Step 3B. If the classifier is unsure,
default to treating the company as a financial institution (Step 3A) rather than
forcing it through the general framework — a wrong bank-side assumption is far
less distortive than reusing "Inventories" and "Trade Receivables" fields that
don't exist for a lender.

## STEP 1 — Pull current market data (not stale filing data)

Pull, in this order, and use these exact fields everywhere downstream:

- `current_price` — latest traded price.
- `market_cap` — latest market capitalization (Screener.in header figure).
- `shares_outstanding_current = market_cap / current_price` — this is the
  number to use in **every** sheet (DCF, Comp Valuation, per-share outputs).
  Do not substitute a share count pulled from an older annual balance sheet
  row (e.g. "No. of Equity Shares" in a 3-year-old filing) — corporate actions
  (preferential allotments, QIPs, buybacks, stake sales) change the float
  between filings, and `market_cap / current_price` is always current because
  it is derived from live market data.
- If the annual-filing share count and `shares_outstanding_current` differ by
  more than 5%, flag this explicitly in the audit log (Step 5) with a one-line
  explanation if you can find one via search (e.g. "Oct 2025 preferential
  allotment to [investor], X% stake").

## STEP 2 — Peer comp set

1. Search for the company's actual sector/sub-sector peers (same regulatory
   category as Step 0's classification — banks vs. banks, NBFCs vs. NBFCs,
   not a mixed set unless the company genuinely has no close peers).
2. **Hard rule: exclude the subject company itself from its own peer list.**
   Before computing any peer AVERAGE, MEDIAN, or regression, check the peer
   list for a name match (case-insensitive, ignoring "Ltd"/"Limited"/"Inc"
   suffixes) against the subject company and drop it if present. Log this in
   the audit output if a self-match was found and removed — that means the
   upstream data pull included it and should be fixed at the source too.
3. Pull for each peer: total debt, market cap (as "total equity" proxy),
   effective tax rate, and either a reported beta or enough price history to
   regress one.

## STEP 2.5 — Data integrity checks (run before anything is computed from the pulled financials)

For every year of pulled P&L data, before using it:

- **Tax vs. PBT check**: if `Tax == Profit Before Tax` (or within 1%) for a
  given year, the tax field was mis-pulled (this has happened via Screener.in
  extraction before — a copy of PBT landed in the tax row). Recompute
  `Tax = Profit Before Tax − Net Profit` from the two fields you trust
  (PBT and Net Profit, both pulled independently) rather than trusting the
  scraped Tax field directly, and flag this in the audit log if triggered.
- **EBITDA sanity check**: after computing EBITDA (PBT + Interest +
  Depreciation — never "Net Profit + Tax + Interest + Depreciation," which
  silently double-counts if Tax is wrong), check the implied EBITDA margin
  against the company's trailing 3-year average margin. Flag if it moved by
  more than 10 percentage points year-on-year with no disclosed one-off cause.
- **Multiple-column self-check**: for every valuation multiple you compute
  (EV/Revenue, EV/EBITDA, P/E, P/B), write out the exact numerator and
  denominator you used next to the multiple in the audit log, so a
  numerator/denominator swap is visible on inspection rather than hidden
  behind a correct-looking column header.

## STEP 3A — Valuation framework: BANK / NBFC / INSURANCE

Do **not** build a WACC-and-FCFF DCF for these. Use one of:

- **Excess Return Model**: Excess Return = (ROE − Cost of Equity) × Book Value
  of Equity, projected forward, discounted at Cost of Equity, plus current
  Book Value. Terminal value via Gordon growth on the terminal excess return.
- **Dividend Discount Model / FCFE**, where FCFE ≈ Net Income − (change in
  regulatory capital needed to support projected loan/asset growth). This
  replaces "reinvestment" — capital retained to fund balance-sheet growth
  under a target capital-adequacy ratio, not PP&E capex or working capital.
- Cost of Equity via CAPM: `Rf + Beta × ERP`, where Beta is the peer-median
  **unlevered-then-relevered** beta computed from the peer set in Step 2 (with
  the subject company excluded per the hard rule), not the subject's own raw
  regression beta alone.
- Still run a comps table (P/E, P/B — not EV/EBITDA, which is not meaningful
  for a bank's revenue line) as a cross-check, and flag in Step 5 if the two
  methods disagree materially.

## STEP 3B — Valuation framework: GENERAL

Standard FCFF DCF is appropriate here:

1. Invested Capital = Net Working Capital (Current Assets − Current
   Liabilities, using the company's real inventory/receivables/payables) +
   Net Non-Current Assets (PP&E net of depreciation).
2. ROIC = EBIT / Invested Capital.
3. Reinvestment Rate = (Net Capex + Change in Working Capital) / EBIT(1−T),
   computed per year, then take the **median** of the last 4 years (median is
   more robust to one-off working-capital swings than the average).
4. **Hard rule — apply the same reinvestment-rate clamp to every projected
   year, including the first one.** `reinvestment_rate = min(0.85, max(0.15,
   raw_median_reinvestment_rate))`. There is no "base year exception" — an
   unclamped year feeding a negative or explosive FCFF is a bug, not a
   modeling choice, even if that year happens not to be summed into the final
   PV total. If it's wrong, it's wrong everywhere it appears.
5. Taper the reinvestment rate from the near-term clamped value toward the
   terminal steady-state value (`terminal_growth / WACC`) over the projection
   window — don't hold the near-term rate constant into perpetuity.
6. WACC: unlever peer betas (excluding self, Step 2), average/median, relever
   at the subject's target capital structure, blend Cost of Equity (CAPM) and
   post-tax Cost of Debt by target weights.

## STEP 4 — One-off normalization

Before finalizing EBIT, Other Income, or Tax Rate inputs, check for single-year
spikes or troughs vs. the trailing trend (>25% deviation from the prior 2-year
average). If found, search for the cause (asset sale, provisioning, regulatory
one-off, M&A) and either exclude that year from medians/averages used for
forward assumptions, or note explicitly why it was kept.

## STEP 5 — Self-audit before returning output

Produce a short audit log alongside the model, answering explicitly:
- Company type classification, and why.
- Was the subject company found in its own peer list and removed? (yes/no)
- Does the annual-filing share count match `shares_outstanding_current`
  within 5%? If not, what changed and which figure was used.
- Was the reinvestment-rate clamp applied identically to every projected
  year? (yes/no)
- Do the DCF/Excess-Return output and the Comp Valuation output agree in
  direction (both over- or both undervalued)? If not, state the likely driver
  of the disagreement rather than silently presenting both numbers.

## STEP 6 — Output format

Return a single JSON object matching the schema in `generate_valuation.py`
(`VALUATION_SCHEMA`) — no prose outside the JSON. The script handles turning
that JSON into the Excel workbook and printing the audit log to the terminal.
