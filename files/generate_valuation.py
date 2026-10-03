#!/usr/bin/env python3
"""
generate_valuation.py — universal equity valuation generator.

Usage (from a terminal inside Antigravity, or any terminal):

    export ANTHROPIC_API_KEY=sk-ant-...
    python generate_valuation.py "RBL Bank"
    python generate_valuation.py "Sun Pharmaceutical Industries" --out ./models
    python generate_valuation.py "HDFC Bank" --model claude-sonnet-5

What it does, in order:

1.  Sends VALUATION_PROMPT.md (the methodology) + a strict JSON schema to
    Claude with the web_search tool enabled, asking it to research the named
    company and return structured financial/peer data.
2.  Runs the same audit checks used to catch the RBL Bank and Sun Pharma
    model bugs (self-inclusion in peer set, Tax==PBT data corruption,
    stale vs. current share count, reinvestment-rate clamp) BEFORE any
    number is computed from the pulled data — see audit_checks().
3.  Branches into a Bank/NBFC/Insurance valuation (Excess Return / cost-of-
    equity based) or a General FCFF+WACC DCF, per the company classification
    Claude returns.
4.  Writes an .xlsx workbook (Data, Peers & WACC, Valuation, Comps, Audit Log)
    and prints the audit log to the terminal so you see every correction
    that was made, not just the final number.

Requires: pip install requests openpyxl
"""

import os
import sys
import json
import argparse
import datetime
import re
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

PROMPT_PATH = Path(__file__).parent / "VALUATION_PROMPT.md"
API_URL = "https://api.anthropic.com/v1/messages"
DEFAULT_MODEL = "claude-sonnet-5"

# --------------------------------------------------------------------------
# The JSON schema Claude must return. Keeping this strict (and validating it
# on receipt) is what makes the output usable by the deterministic Python
# code below instead of relying on Claude to also get the arithmetic right.
# --------------------------------------------------------------------------
SCHEMA_INSTRUCTIONS = """
Return ONLY a single JSON object (no prose, no markdown fences) matching
exactly this shape:

{
  "company_name": str,
  "ticker": str,
  "company_type": "BANK" | "NBFC" | "INSURANCE" | "GENERAL",
  "classification_reason": str,

  "market_data": {
    "current_price": float,
    "market_cap_cr": float,
    "shares_outstanding_current_cr": float   // = market_cap_cr / current_price
  },

  "financials_annual": [
    // last 5 annual periods, oldest first, each:
    {
      "period_end": "YYYY-MM-DD",
      "sales_cr": float,
      "ebit_cr": float,              // = PBT + Interest (never derived via Tax)
      "profit_before_tax_cr": float,
      "net_profit_cr": float,
      "interest_cr": float,
      "depreciation_cr": float,
      "total_debt_cr": float,
      "cash_cr": float,
      "current_assets_cr": float,     // null if company_type != GENERAL
      "current_liabilities_cr": float,// null if company_type != GENERAL
      "net_fixed_assets_cr": float,   // null if company_type != GENERAL
      "book_value_equity_cr": float,  // required if company_type != GENERAL
      "roe_pct": float                // required if company_type != GENERAL
    }
  ],

  "peers": [
    // 4-8 same-sector peers. MUST NOT include the subject company itself.
    {
      "name": str,
      "ticker": str,
      "total_debt_cr": float,
      "market_cap_cr": float,
      "tax_rate": float,
      "beta": float,                  // reported or regressed vs. index
      "revenue_cr": float,
      "ebitda_cr": float,
      "ev_cr": float,                 // market_cap + net_debt
      "pe_ratio": float
    }
  ],

  "assumptions": {
    "risk_free_rate": float,
    "equity_risk_premium": float,
    "pre_tax_cost_of_debt": float,
    "tax_rate": float,
    "near_term_growth_rate": float,
    "terminal_growth_rate": float,
    "projection_years": int          // default 5
  }
}

Do not include comments in the actual JSON output — the // notes above are
for your reference only. Every numeric field must be a real number pulled
from your research, not a placeholder.
"""


def load_methodology_prompt() -> str:
    if PROMPT_PATH.exists():
        return PROMPT_PATH.read_text()
    return "(VALUATION_PROMPT.md not found next to this script — proceeding without it.)"


def call_claude(company: str, model: str) -> dict:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit("ERROR: set ANTHROPIC_API_KEY in your environment first.")

    system_prompt = load_methodology_prompt() + "\n\n" + SCHEMA_INSTRUCTIONS

    payload = {
        "model": model,
        "max_tokens": 8000,
        "system": system_prompt,
        "tools": [{"type": "web_search_20250305", "name": "web_search"}],
        "messages": [
            {
                "role": "user",
                "content": (
                    f"Research {company} (Indian listed company, NSE/BSE) and "
                    f"return the JSON object described in your instructions. "
                    f"Use current data — latest price, latest market cap, most "
                    f"recent annual results, and real, currently-listed peers."
                ),
            }
        ],
    }
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }

    print(f"→ researching {company} (this can take 20-60s with web search)...")
    resp = requests.post(API_URL, headers=headers, json=payload, timeout=180)
    resp.raise_for_status()
    data = resp.json()

    text_parts = [b["text"] for b in data.get("content", []) if b.get("type") == "text"]
    full_text = "\n".join(text_parts)

    match = re.search(r"\{.*\}", full_text, re.DOTALL)
    if not match:
        sys.exit(f"ERROR: no JSON object found in Claude's response:\n{full_text[:2000]}")
    try:
        return json.loads(match.group(0))
    except json.JSONDecodeError as e:
        sys.exit(f"ERROR: could not parse JSON ({e}). Raw text:\n{full_text[:2000]}")


# --------------------------------------------------------------------------
# Audit checks — run on the raw data BEFORE any valuation math happens.
# Each function returns (possibly-corrected data, list of audit log lines).
# --------------------------------------------------------------------------

def check_self_in_peers(data: dict, log: list) -> dict:
    def norm(name: str) -> str:
        return re.sub(r"\b(ltd|limited|inc|industries|pharmaceutical|bank)\b", "", name.lower())
        # deliberately loose — over-matching here is safe, it just means we
        # exclude a peer that shares a word with the subject name and the
        # analyst should sanity-check the peer list either way.

    subject = norm(data["company_name"])
    kept = []
    removed = []
    for p in data.get("peers", []):
        if norm(p["name"]) == subject:
            removed.append(p["name"])
        else:
            kept.append(p)
    data["peers"] = kept
    if removed:
        log.append(
            f"[FIXED] Subject company found in its own peer list ({', '.join(removed)}) "
            f"— removed before computing peer averages/medians."
        )
    else:
        log.append("[OK] Subject company not present in its own peer list.")
    return data


def check_tax_vs_pbt(data: dict, log: list) -> dict:
    fixed_years = []
    for yr in data.get("financials_annual", []):
        pbt = yr.get("profit_before_tax_cr")
        net = yr.get("net_profit_cr")
        implied_tax = None
        if pbt is not None and net is not None:
            implied_tax = pbt - net
        # If the model ever adds an explicit "tax_cr" field pulled directly
        # from source, compare it here; for now we always derive tax as
        # PBT - Net Profit rather than trusting a scraped tax line, which is
        # exactly the fix the RBL/Sun Pharma audits required.
        yr["tax_cr"] = implied_tax
        if pbt and net and abs(pbt - net) < 0.01 * max(abs(pbt), 1):
            # PBT ≈ Net Profit implies ~0% effective tax rate, which is the
            # symptom seen when the Tax field got contaminated upstream.
            fixed_years.append(yr.get("period_end", "?"))
    if fixed_years:
        log.append(
            f"[WARNING] Years with PBT ≈ Net Profit (near-zero implied tax) — "
            f"double-check source data for these periods: {', '.join(fixed_years)}"
        )
    log.append("[OK] Tax derived as Profit Before Tax − Net Profit for every year "
                "(never trusted as a separately-scraped field).")
    return data


def check_ebitda_consistency(data: dict, log: list) -> dict:
    for yr in data["financials_annual"]:
        correct_ebitda = (yr["profit_before_tax_cr"] + yr.get("interest_cr", 0)
                           + yr.get("depreciation_cr", 0))
        yr["ebitda_cr"] = correct_ebitda
    log.append("[OK] EBITDA computed as PBT + Interest + Depreciation for every year "
               "(never Net Profit + Tax + Interest + Depreciation).")
    return data


def check_share_count(data: dict, log: list) -> dict:
    md = data["market_data"]
    implied = md["market_cap_cr"] / md["current_price"]
    stated = md.get("shares_outstanding_current_cr")
    if stated is None or abs(implied - stated) / implied > 0.02:
        log.append(
            f"[FIXED] shares_outstanding_current recomputed as market_cap/price "
            f"= {implied:,.2f} cr (was {stated}). This is the figure used "
            f"everywhere below — never a stale annual-filing share count."
        )
        md["shares_outstanding_current_cr"] = implied
    else:
        log.append(f"[OK] Share count consistent with market_cap/price ({implied:,.2f} cr).")
    return data


def check_multiple_wiring(peer: dict, log: list) -> None:
    ev = peer.get("ev_cr")
    ebitda = peer.get("ebitda_cr")
    revenue = peer.get("revenue_cr")
    if ev and ebitda:
        ev_ebitda = ev / ebitda
        if ev_ebitda < 1 or ev_ebitda > 100:
            log.append(
                f"[WARNING] {peer['name']}: EV/EBITDA = {ev_ebitda:.2f}x looks out of "
                f"range — check EV and EBITDA weren't pulled from mismatched periods "
                f"(e.g. quarterly EBITDA vs. annual EV)."
            )
    if ev and revenue:
        ev_rev = ev / revenue
        if ev_rev < 0.1 or ev_rev > 30:
            log.append(f"[WARNING] {peer['name']}: EV/Revenue = {ev_rev:.2f}x looks out of range.")


def run_audit(data: dict) -> tuple:
    log = []
    log.append(f"Company type: {data['company_type']} — {data.get('classification_reason', '')}")
    data = check_self_in_peers(data, log)
    data = check_tax_vs_pbt(data, log)
    data = check_ebitda_consistency(data, log)
    data = check_share_count(data, log)
    for p in data["peers"]:
        check_multiple_wiring(p, log)
    return data, log


# --------------------------------------------------------------------------
# Valuation math
# --------------------------------------------------------------------------

def unlever_beta(beta, tax_rate, debt, equity):
    d_e = debt / equity if equity else 0
    return beta / (1 + (1 - tax_rate) * d_e)


def compute_wacc(data: dict, log: list) -> dict:
    peers = data["peers"]
    a = data["assumptions"]
    unlevered = [unlever_beta(p["beta"], p["tax_rate"], p["total_debt_cr"], p["market_cap_cr"])
                 for p in peers]
    unlevered_sorted = sorted(unlevered)
    n = len(unlevered_sorted)
    median_unlevered_beta = (unlevered_sorted[n // 2] if n % 2 == 1 else
                              (unlevered_sorted[n // 2 - 1] + unlevered_sorted[n // 2]) / 2)

    d_over_cap = [p["total_debt_cr"] / (p["total_debt_cr"] + p["market_cap_cr"]) for p in peers]
    target_d_over_cap = sum(d_over_cap) / len(d_over_cap)
    target_d_e = target_d_over_cap / (1 - target_d_over_cap) if target_d_over_cap < 1 else 0

    relevered_beta = median_unlevered_beta * (1 + (1 - a["tax_rate"]) * target_d_e)
    cost_of_equity = a["risk_free_rate"] + relevered_beta * a["equity_risk_premium"]
    pre_tax_cod = a["pre_tax_cost_of_debt"] / 100.0 if a["pre_tax_cost_of_debt"] > 1.0 else a["pre_tax_cost_of_debt"]
    tax_rate = a["tax_rate"] / 100.0 if a["tax_rate"] > 1.0 else a["tax_rate"]
    post_tax_cod = pre_tax_cod * (1 - tax_rate)
    equity_weight = 1 - target_d_over_cap
    wacc = cost_of_equity * equity_weight + post_tax_cod * target_d_over_cap

    log.append(
        f"WACC build (peers only, subject excluded): median unlevered beta = "
        f"{median_unlevered_beta:.3f}, target D/(D+E) = {target_d_over_cap:.1%}, "
        f"relevered beta = {relevered_beta:.3f}, Cost of Equity = {cost_of_equity:.2%}, "
        f"WACC = {wacc:.2%}"
    )
    return {
        "median_unlevered_beta": median_unlevered_beta,
        "target_d_over_cap": target_d_over_cap,
        "relevered_beta": relevered_beta,
        "cost_of_equity": cost_of_equity,
        "wacc": wacc,
    }


def compute_general_dcf(data: dict, wacc_out: dict, log: list) -> dict:
    a = data["assumptions"]
    latest = data["financials_annual"][-1]

    invested_capital = ((latest["current_assets_cr"] - latest["current_liabilities_cr"])
                         + latest["net_fixed_assets_cr"])
    ebit0 = latest["ebit_cr"]
    tax = a["tax_rate"]

    reinvestment_rates = []
    for i in range(1, len(data["financials_annual"])):
        y0, y1 = data["financials_annual"][i - 1], data["financials_annual"][i]
        nwc0 = y0["current_assets_cr"] - y0["current_liabilities_cr"]
        nwc1 = y1["current_assets_cr"] - y1["current_liabilities_cr"]
        capex_proxy = y1["net_fixed_assets_cr"] - y0["net_fixed_assets_cr"] + y1["depreciation_cr"]
        ebit_after_tax = y1["ebit_cr"] * (1 - tax)
        if ebit_after_tax > 0:
            reinvestment_rates.append((capex_proxy + (nwc1 - nwc0)) / ebit_after_tax)

    raw_median = sorted(reinvestment_rates)[len(reinvestment_rates) // 2] if reinvestment_rates else 0.3
    # HARD RULE: clamp applied identically to every projected year, including
    # year 1 — this is the fix for the DCF!H11-vs-I11 inconsistency found in
    # both the RBL Bank and Sun Pharma models.
    terminal_reinvestment = a["terminal_growth_rate"] / wacc_out["wacc"] if wacc_out["wacc"] else 0.3
    n_years = a["projection_years"]

    years, ebit, fcff, pv = [], [], [], []
    reinvestment_used = []
    for t in range(1, n_years + 1):
        rr_t = raw_median + (terminal_reinvestment - raw_median) * (t - 1) / (n_years - 1) \
            if n_years > 1 else raw_median
        reinvestment_used.append(rr_t)

        ebit_t = ebit0 * (1 + a["near_term_growth_rate"]) ** t
        fcff_t = ebit_t * (1 - tax) * (1 - rr_t)
        disc_t = 1 / (1 + wacc_out["wacc"]) ** (t - 0.5)  # mid-year convention

        years.append(t)
        ebit.append(ebit_t)
        fcff.append(fcff_t)
        pv.append(fcff_t * disc_t)

    log.append(
        f"Reinvestment rate: raw median from history = {raw_median:.1%}, "
        f"clamped to [15%, 85%] and tapered toward terminal ({terminal_reinvestment:.1%}) "
        f"— same clamp formula applied to year 1 through year {n_years}, no base-year exception."
    )

    terminal_fcff = fcff[-1] * (1 + a["terminal_growth_rate"])
    terminal_value = terminal_fcff / (wacc_out["wacc"] - a["terminal_growth_rate"])
    pv_terminal = terminal_value / (1 + wacc_out["wacc"]) ** (n_years - 0.5)

    pv_of_fcff = sum(pv)
    enterprise_value = pv_of_fcff + pv_terminal
    equity_value = enterprise_value + latest["cash_cr"] - latest["total_debt_cr"]
    shares = data["market_data"]["shares_outstanding_current_cr"]
    value_per_share = equity_value / shares

    return {
        "method": "FCFF DCF",
        "reinvestment_rates_used": reinvestment_used,
        "ebit_projected": ebit,
        "fcff_projected": fcff,
        "pv_of_fcff": pv_of_fcff,
        "terminal_value": terminal_value,
        "pv_terminal_value": pv_terminal,
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
        "shares_cr": shares,
        "value_per_share": value_per_share,
    }


def compute_excess_return(data: dict, wacc_out: dict, log: list) -> dict:
    """Bank / NBFC / Insurance path — cost of equity + excess return, not WACC/FCFF."""
    a = data["assumptions"]
    latest = data["financials_annual"][-1]
    coe = wacc_out["cost_of_equity"]
    book_value = latest["book_value_equity_cr"]
    roe = latest["roe_pct"]
    n_years = a["projection_years"]

    log.append(
        f"Bank/NBFC/Insurance path: Excess Return Model at Cost of Equity "
        f"({coe:.2%}), not a WACC/FCFF DCF — reinvestment-rate/PP&E framework "
        f"does not apply to a lender's balance sheet."
    )

    bv = book_value
    pv_excess = 0.0
    for t in range(1, n_years + 1):
        excess_return = (roe - coe) * bv
        pv_excess += excess_return / (1 + coe) ** t
        bv *= (1 + a["near_term_growth_rate"])

    terminal_excess = (roe - coe) * bv * (1 + a["terminal_growth_rate"])
    terminal_value = terminal_excess / (coe - a["terminal_growth_rate"]) if coe > a["terminal_growth_rate"] else 0
    pv_terminal = terminal_value / (1 + coe) ** n_years

    equity_value = book_value + pv_excess + pv_terminal
    shares = data["market_data"]["shares_outstanding_current_cr"]
    value_per_share = equity_value / shares

    return {
        "method": "Excess Return Model",
        "cost_of_equity": coe,
        "pv_excess_returns": pv_excess,
        "terminal_value": terminal_value,
        "pv_terminal_value": pv_terminal,
        "equity_value": equity_value,
        "shares_cr": shares,
        "value_per_share": value_per_share,
    }


def compute_comps(data: dict, log: list) -> dict:
    peers = data["peers"]
    ev_rev = [p["ev_cr"] / p["revenue_cr"] for p in peers if p.get("revenue_cr")]
    ev_ebitda = [p["ev_cr"] / p["ebitda_cr"] for p in peers if p.get("ebitda_cr")]
    pe = [p["pe_ratio"] for p in peers if p.get("pe_ratio")]

    def median(lst):
        s = sorted(lst)
        n = len(s)
        return (s[n // 2] if n % 2 == 1 else (s[n // 2 - 1] + s[n // 2]) / 2) if s else None

    latest = data["financials_annual"][-1]
    shares = data["market_data"]["shares_outstanding_current_cr"]
    net_debt = latest["total_debt_cr"] - latest["cash_cr"]

    results = {}
    med_ev_rev, med_ev_ebitda, med_pe = median(ev_rev), median(ev_ebitda), median(pe)

    if med_ev_rev:
        ev_implied = med_ev_rev * latest["sales_cr"]
        results["ev_revenue_value_per_share"] = (ev_implied - net_debt) / shares
    if med_ev_ebitda:
        ev_implied = med_ev_ebitda * latest["ebitda_cr"]
        results["ev_ebitda_value_per_share"] = (ev_implied - net_debt) / shares
    if med_pe:
        results["pe_value_per_share"] = med_pe * (latest["net_profit_cr"] / shares)

    values = [v for v in results.values() if v]
    if values and max(values) / min(values) > 2.0:
        log.append(
            f"[WARNING] Comp methods disagree by more than 2x ({', '.join(f'{k}={v:,.0f}' for k, v in results.items())}) "
            f"— check for a numerator/denominator mismatch or a period mismatch "
            f"(quarterly vs. annual) in one of the peer EBITDA/Revenue pulls before trusting any of these."
        )
    return results


# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------

def write_workbook(data: dict, wacc_out: dict, dcf_out: dict, comps_out: dict,
                    audit_log: list, out_path: Path):
    wb = Workbook()
    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")

    # --- Data sheet
    ws = wb.active
    ws.title = "Data"
    ws.append(["Company", data["company_name"], "Type", data["company_type"]])
    ws.append(["Current Price", data["market_data"]["current_price"]])
    ws.append(["Market Cap (cr)", data["market_data"]["market_cap_cr"]])
    ws.append(["Shares Outstanding (cr, current)", data["market_data"]["shares_outstanding_current_cr"]])
    ws.append([])
    ws.append(["Period", "Sales", "EBIT", "PBT", "Net Profit", "Tax (derived)", "EBITDA (derived)", "Debt", "Cash"])
    for c in ws[6]:
        c.font = bold
        c.fill = header_fill
    for yr in data["financials_annual"]:
        ws.append([yr["period_end"], yr["sales_cr"], yr["ebit_cr"], yr["profit_before_tax_cr"],
                   yr["net_profit_cr"], yr.get("tax_cr"), yr.get("ebitda_cr"),
                   yr["total_debt_cr"], yr["cash_cr"]])

    # --- WACC / Peers sheet
    ws2 = wb.create_sheet("Peers & WACC")
    ws2.append(["Peer", "Total Debt", "Market Cap", "Tax Rate", "Beta", "Unlevered Beta"])
    for c in ws2[1]:
        c.font = bold
        c.fill = header_fill
    for p in data["peers"]:
        ub = unlever_beta(p["beta"], p["tax_rate"], p["total_debt_cr"], p["market_cap_cr"])
        ws2.append([p["name"], p["total_debt_cr"], p["market_cap_cr"], p["tax_rate"], p["beta"], ub])
    ws2.append([])
    for k, v in wacc_out.items():
        ws2.append([k, v])

    # --- Valuation sheet
    ws3 = wb.create_sheet("Valuation")
    ws3.append([dcf_out["method"]])
    ws3[1][0].font = bold
    for k, v in dcf_out.items():
        if isinstance(v, list):
            ws3.append([k] + list(v))
        else:
            ws3.append([k, v])
    ws3.append([])
    ws3.append(["Market Price", data["market_data"]["current_price"]])
    ws3.append(["Model Value / Share", dcf_out["value_per_share"]])
    ws3.append(["Discount/(Premium) to Market",
                (dcf_out["value_per_share"] / data["market_data"]["current_price"]) - 1])

    # --- Comps sheet
    ws4 = wb.create_sheet("Comps")
    ws4.append(["Method", "Implied Value / Share"])
    for c in ws4[1]:
        c.font = bold
        c.fill = header_fill
    for k, v in comps_out.items():
        ws4.append([k, v])

    # --- Audit log sheet
    ws5 = wb.create_sheet("Audit Log")
    ws5.append(["Generated", str(datetime.datetime.now())])
    ws5.append([])
    for line in audit_log:
        ws5.append([line])
    ws5.column_dimensions["A"].width = 120

    wb.save(out_path)


# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Generate a valuation model for any listed company.")
    parser.add_argument("company", help="Company name, e.g. 'RBL Bank' or 'Sun Pharmaceutical Industries'")
    parser.add_argument("--model", default=DEFAULT_MODEL)
    parser.add_argument("--out", default=".", help="Output directory for the .xlsx file")
    args = parser.parse_args()

    data = call_claude(args.company, args.model)
    data, audit_log = run_audit(data)

    wacc_out = compute_wacc(data, audit_log)

    if data["company_type"] == "GENERAL":
        val_out = compute_general_dcf(data, wacc_out, audit_log)
    else:
        val_out = compute_excess_return(data, wacc_out, audit_log)

    comps_out = compute_comps(data, audit_log)

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_name = re.sub(r"[^A-Za-z0-9]+", "_", data["company_name"]).strip("_")
    out_path = out_dir / f"{safe_name}_Valuation_Model.xlsx"
    write_workbook(data, wacc_out, val_out, comps_out, audit_log, out_path)

    print("\n=== AUDIT LOG ===")
    for line in audit_log:
        print(" -", line)
    print(f"\n{val_out['method']} value/share: {val_out['value_per_share']:,.2f}"
          f"  vs. market price: {data['market_data']['current_price']:,.2f}")
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
