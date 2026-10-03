---
name: financial-valuation
description: >-
  Use this skill when the user asks to run a financial valuation for a specific company or stock.
---

# Financial Valuation Workflow

When asked to value a company, follow these steps exactly:

1. **Source Data & Sector Peers**:
   - Search data from Screener.in for the latest financial statements, growth metrics, and ratios.
   - **Critical Peer Selection**: ALWAYS pick true **Industry / Sector Group** peers (from the Screener market classification hierarchy, e.g., `Metals & Mining` for Adani Enterprises or Tata Steel, `Fast Moving Consumer Goods` for FMCG companies, `Information Technology` for IT firms) sorted by Market Cap descending. **NEVER** pick micro-cap companies from narrow sub-categories (e.g. Starlineps or Rajdarshan) for large-cap comp valuation; peers must be viable and of comparable scale.
2. **Company Background & Recent Updates**:
   - **About the Company**: Sourced dynamically from Wikipedia (or company profile) for the target company and populated in cell `B8` of both `Dupont Analysis` and `Altman's Z Score` sheets.
   - **Recent Updates**: Sourced from Economic Times (or financial news) and structured in the exact 5-bullet institutional format (capacity/capex expansion, segment demand/volume drivers, tech/strategic initiatives, latest quarterly revenue & profit performance, and management roadmap) in rows 37-45 of `Dupont Analysis` and rows 36-44 of `Altman's Z Score`.
   - **52-Week Range**: Live 52-week high and low must be updated in cell `B5` in both sheets.
3. **Reference the Model**: Open and study `Nestle India Model.xlsx` and `ITC Model.xlsx` in the workspace to understand the preferred valuation methodology, structure, and formatting.
4. **Peer-Comparables Data Block Standing Rules (MANDATORY)**:
   When generating or updating the peer-comparables data block for ANY company's valuation model (the block holding Market Cap / Net Debt / EV / Sales / EBITDA / Net Profit for the subject company + its peers — e.g. `Raw FS!AR:BA` and `Comp_Valuation`), follow these 5 rules without exception:
   - **Rule 1: KEY BY IDENTITY, NOT BY ROW POSITION.**
     Never assume "row N in the name column corresponds to row N in the financial-data columns." Write each company's financial data using a lookup keyed to that company's name/ticker (INDEX/MATCH or equivalent), not a fixed row offset. If hardcoding values instead of formulas, paste name and financials as a single atomic row-write — never populate the name column and the financial columns in two separate passes.
   - **Rule 2: THE SUBJECT COMPANY'S OWN ROW IS NOT A PEER ROW.**
     The subject company (the one being valued) must have its own financials pulled fresh from the primary data source for THIS company (`Data Sheet` / `Raw Data` — whichever sheet holds the live Screener.in export for the company currently being valued), never copied, shifted, or interpolated from an adjacent peer row. Before finalizing the sheet, explicitly check: does the subject company's row in the peer-comps block use a formula or reference that traces back to that company's OWN Data Sheet import — not a peer's? Row 56 in `Raw FS` belongs strictly to the target company.
   - **Rule 3: SELF-CHECK BEFORE MOVING ON.**
     For every row in the peer-comps block (including the subject company's own row), verify `Market Cap ≈ Share Price × Shares Outstanding` for THAT SAME named row, within ~2% tolerance. If any row fails this check, the financial-data columns and the name column are misaligned somewhere in the block — stop and find the offset before using the block in `Comp_Valuation` or `WACC`.
   - **Rule 4: NEVER LEAVE A ROW BLANK MID-BLOCK.**
     If the peer list has N companies, the financial-data columns must also have exactly N populated rows, aligned 1:1 with the name rows. A block that's short by one row (data starting one row before/after the names) is the single most common failure mode here — check row counts match exactly before using the block downstream.
   - **Rule 5: AFTER POPULATING, PRINT A VALIDATION TABLE.**
     Print the validation table: `company name | mkt cap (computed from price×shares) | mkt cap (as entered in the comps block) | %diff`. Flag any row with >2% diff before proceeding to `Comp_Valuation`, `DCF`, or `WACC`.
5. **Perform Valuation**: Using the data you found, perform a valuation matching the logic found in the models (DCF, WACC, DuPont, Altman Z, and Comparable Multiples).
6. **Calculate and Report**: Calculate the intrinsic value. Output a clean Markdown artifact summarizing the findings and the calculated intrinsic value.

