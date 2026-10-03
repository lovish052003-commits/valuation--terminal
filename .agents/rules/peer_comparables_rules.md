# Standing Rules for Peer-Comparables Data Block

When generating or updating the peer-comparables data block for ANY company's
valuation model (the block holding Market Cap / Net Debt / EV / Sales / EBITDA /
Net Profit for the subject company + its peers — e.g. `Raw FS!AR:BA` and `Comp_Valuation`),
follow these rules without exception:

## 1. KEY BY IDENTITY, NOT BY ROW POSITION
Never assume "row N in the name column corresponds to row N in the financial-data columns."
Write each company's financial data using a lookup keyed to that company's name/ticker (INDEX/MATCH or equivalent), not a fixed row offset.
If you must hardcode values instead of formulas, paste name and financials as a single atomic row-write — never populate the name column and the financial columns in two separate passes, since that's how blocks drift out of sync with each other.

## 2. THE SUBJECT COMPANY'S OWN ROW IS NOT A PEER ROW
The subject company (the one being valued) must have its own financials pulled fresh from the primary data source for THIS company (`Data Sheet` / `Raw Data` — whichever sheet holds the live Screener.in export for the company currently being valued), never copied, shifted, or interpolated from an adjacent peer row.
Before finalizing the sheet, explicitly check: does the subject company's row in the peer-comps block use a formula or reference that traces back to that company's OWN Data Sheet import — not a peer's?
Row 56 in `Raw FS` belongs strictly to the subject company. Peer rows begin at Row 57 (Rows 57 to 66).

## 3. SELF-CHECK BEFORE MOVING ON
For every row in the peer-comps block (including the subject company's own row), verify:
$$\text{Market Cap} \approx \text{Share Price} \times \text{Shares Outstanding}$$
for THAT SAME named row, within ~2% tolerance.
If any row fails this check, the financial-data columns and the name column are misaligned somewhere in the block — stop and find the offset before using the block in `Comp_Valuation` or `WACC`.

## 4. NEVER LEAVE A ROW BLANK MID-BLOCK
If the peer list has N companies, the financial-data columns must also have exactly N populated rows, aligned 1:1 with the name rows.
A block that's short by one row (data starting one row before/after the names) is the single most common failure mode here — check row counts match exactly before using the block downstream.

## 5. AFTER POPULATING, PRINT A VALIDATION TABLE
After populating, print a validation table:
`company name | mkt cap (computed from price×shares) | mkt cap (as entered in the comps block) | %diff`
Flag any row with >2% diff before proceeding to `Comp_Valuation`, `DCF`, or `WACC`.

Apply this to every company run through the terminal or platform, not just the one currently being valued — this is a standing rule for how the peer-comps block gets built or refreshed.
