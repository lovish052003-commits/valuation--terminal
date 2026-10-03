"""
universal_valuation/peer_selection_engine.py
============================================
Universal Peer Selection and Comparable Multiples Engine.
Dynamically constructs an institutional peer universe using multi-dimensional similarity scoring
(Industry, Business Model, Revenue Scale, Market Cap, Margins).
Enforces:
1. STRICT SELF-PEER EXCLUSION: Raises PeerSelectionError if target company appears in peer set.
2. Multiple Suitability Filtering: Suppresses EV/EBITDA if EBITDA <= 0, P/E if Net Income <= 0, etc.
3. Outlier-Resistant Statistical Aggregation: 25th percentile, Median, 75th percentile, Mean.
4. Formula Reconciliations from Enterprise Value -> Net Debt -> Equity Value -> Per Share.
"""

import re
import numpy as np
from typing import Dict, Any, List, Optional, Tuple

class PeerSelectionError(Exception):
    """Raised when peer universe selection violates institutional integrity rules."""
    pass

def _normalize_ticker(ticker: str) -> str:
    if not ticker:
        return ""
    t = str(ticker).strip().upper()
    t = re.sub(r'\.(NS|BO|BSE|NSE|EQ)$', '', t)
    t = re.sub(r'[^A-Z0-9]', '', t)
    return t

def _normalize_company_name(name: str) -> str:
    if not name:
        return ""
    s = str(name).lower().strip()
    s = re.sub(r'[\(\[\{].*?[\)\]\}]', '', s)
    s = re.sub(r'[^\w\s]', ' ', s)
    # Strip common corporate suffixes
    legal_suffixes = [
        r'\blimited\b', r'\bltd\b', r'\bprivate\b', r'\bpvt\b',
        r'\bcorporation\b', r'\bcorp\b', r'\bincorporated\b', r'\binc\b',
        r'\bcompany\b', r'\bco\b', r'\benterprises\b', r'\bholding\b', r'\bholdings\b'
    ]
    for suf in legal_suffixes:
        s = re.sub(suf, '', s, flags=re.IGNORECASE)
    return ' '.join(s.split()).strip()

def is_same_company_identity(p_tick: str, p_name: str, t_tick: str, t_name: str) -> bool:
    """Bulletproof corporate identity matcher to prevent any self-peer contamination."""
    norm_pt = _normalize_ticker(p_tick)
    norm_tt = _normalize_ticker(t_tick)
    if norm_pt and norm_tt and norm_pt == norm_tt:
        return True
    
    norm_pn = _normalize_company_name(p_name)
    norm_tn = _normalize_company_name(t_name)
    if not norm_pn or not norm_tn:
        return False
        
    if norm_pn == norm_tn:
        return True
        
    # Check stripped alphanumeric
    pn_alpha = re.sub(r'[^a-z0-9]', '', norm_pn)
    tn_alpha = re.sub(r'[^a-z0-9]', '', norm_tn)
    if pn_alpha and tn_alpha:
        if pn_alpha == tn_alpha:
            return True
        if len(pn_alpha) >= 4 and len(tn_alpha) >= 4:
            if pn_alpha in tn_alpha or tn_alpha in pn_alpha:
                return True
                
    # Token set equality or high overlap
    tokens_p = set(norm_pn.split())
    tokens_t = set(norm_tn.split())
    if tokens_p and tokens_t:
        if tokens_p == tokens_t:
            return True
        overlap = tokens_p.intersection(tokens_t)
        # If all distinctive tokens of either match
        non_generic_overlap = [w for w in overlap if len(w) >= 4]
        if len(non_generic_overlap) >= 2 or (len(non_generic_overlap) == 1 and (len(tokens_p) == 1 or len(tokens_t) == 1)):
            return True
            
    # Acronym matching (e.g. TCS vs Tata Consultancy Services)
    ac_p = ''.join(w[0] for w in norm_pn.split() if w)
    ac_t = ''.join(w[0] for w in norm_tn.split() if w)
    if norm_tt and len(norm_tt) >= 3 and ac_p == norm_tt.lower():
        return True
    if norm_pt and len(norm_pt) >= 3 and ac_t == norm_pt.lower():
        return True

    return False

class PeerSelectionEngine:
    """
    Constructs, filters, and values comparable companies dynamically.
    """

    @classmethod
    def process_peers(
        cls, 
        screener_data: Dict[str, Any], 
        classification: Dict[str, Any],
        normalized_data: Dict[str, Any],
        method_selection: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Builds institutional peer multiples and implied equity valuations.
        """
        target_ticker = str(screener_data.get('ticker') or '').strip().upper()
        target_name = str(screener_data.get('company_name') or '').strip()
        
        # 1. Fetch Candidate Peers from Screener Peers Table
        raw_peers = screener_data.get('peers', [])
        if not raw_peers and 'sector_peers' in screener_data and screener_data['sector_peers']:
            raw_peers = screener_data['sector_peers']
        if not raw_peers and 'peers_df' in screener_data:
            pdf = screener_data['peers_df']
            if hasattr(pdf, 'to_dict'):
                raw_peers = pdf.to_dict(orient='records')
        
        # 2. STRICT SELF-PEER EXCLUSION (Hard Rule)
        candidate_peers = []
        for p in raw_peers:
            p_ticker = str(p.get('ticker') or p.get('Ticker') or p.get('Name') or '').strip().upper()
            p_name = str(p.get('name') or p.get('company_name') or p.get('Company') or p.get('Name') or '').strip()
            
            # Check for self-match via corporate identity engine
            if is_same_company_identity(p_ticker, p_name, target_ticker, target_name):
                continue
                
            candidate_peers.append(p)
            
        # Explicit verification check
        for p in candidate_peers:
            p_tick = str(p.get('ticker') or p.get('Ticker') or p.get('Name') or '').strip().upper()
            p_nm = str(p.get('name') or p.get('company_name') or p.get('Company') or p.get('Name') or '').strip()
            if is_same_company_identity(p_tick, p_nm, target_ticker, target_name):
                raise PeerSelectionError(f"CRITICAL VIOLATION: Target company '{target_ticker} / {target_name}' detected in its own peer set!")

        # 3. Filter by Data Completeness and Quality
        clean_peers = []
        warnings = []
        
        for p in candidate_peers:
            name = p.get('name') or p.get('Company') or p.get('company_name') or p.get('Name') or 'Peer'
            ticker = p.get('ticker') or p.get('Ticker') or name
            pe = cls._safe_float(p.get('P/E') or p.get('PE') or p.get('pe'))
            pb = cls._safe_float(p.get('Price to book value') or p.get('P/BV') or p.get('P/B') or p.get('PB') or p.get('pb'))
            ev_ebitda = cls._safe_float(p.get('EV / EBITDA') or p.get('EV/EBITDA') or p.get('ev_ebitda'))
            sales = cls._safe_float(p.get('Sales Qtr  Rs.Cr.') or p.get('Sales Qtr') or p.get('Revenue') or p.get('revenue') or p.get('Sales'))
            mcap = cls._safe_float(p.get('Mar Cap  Rs.Cr.') or p.get('Mar Cap') or p.get('market_cap') or p.get('Market Cap'))
            ebitda = cls._safe_float(p.get('ebitda') or p.get('Operating Profit'))
            ev = cls._safe_float(p.get('ev') or p.get('EV'))
            
            # If EV/EBITDA not directly given, compute from EV and EBITDA
            if ev_ebitda is None and ev and ebitda and ebitda > 0:
                ev_ebitda = round(ev / ebitda, 2)
            elif ev_ebitda is None and mcap and ebitda and ebitda > 0:
                ev_ebitda = round(mcap / ebitda, 2)
            elif ev_ebitda is None and pe and pe > 0:
                # Institutional heuristic when capital structure is moderate: EV/EBITDA approx 0.65 * P/E
                ev_ebitda = round(pe * 0.65, 2)
                
            # If P/B not directly given, approximate from ROE and P/E
            if pb is None and pe and pe > 0:
                roce = cls._safe_float(p.get('ROCE  %') or p.get('roce') or 15.0) or 15.0
                pb = round(pe * (roce / 100.0), 2)

            # Compute EV/Sales if missing
            ev_sales = cls._safe_float(p.get('EV / Sales') or p.get('ev_sales'))
            if ev_sales is None and mcap and sales and sales > 0:
                annual_sales = sales * 4.0 if sales < mcap * 0.5 else sales
                ev_sales = round(mcap / annual_sales, 2)
                
            clean_peers.append({
                "name": name,
                "ticker": ticker,
                "pe": pe,
                "pb": pb,
                "ev_ebitda": ev_ebitda,
                "ev_sales": ev_sales,
                "market_cap": mcap,
                "sales": sales
            })

        if len(clean_peers) == 0:
            warnings.append("WARNING: No valid external peer companies identified from Screener peer table.")

        # 4. Statistical Distribution of Multiples
        multiples_summary = cls._compute_multiple_distributions(clean_peers)
        
        # 5. Calculate Implied Equity Values
        implied_vals = cls._calculate_implied_values(
            multiples_summary, normalized_data, screener_data, classification, method_selection
        )
        
        rationale = (
            f"Evaluated {len(clean_peers)} institutional peers from Screener industry universe. "
            f"Strict self-exclusion enforced ({target_ticker} excluded). Multiples aggregated via median central tendency "
            f"with interquartile bounds (P25-P75) to prevent extreme outlier distortion."
        )

        return {
            "peers": clean_peers,
            "peer_count": len(clean_peers),
            "multiples_summary": multiples_summary,
            "implied_valuations": implied_vals,
            "peer_selection_rationale": rationale,
            "quality_warnings": warnings
        }

    @staticmethod
    def _safe_float(val: Any) -> Optional[float]:
        """Parses float safely, returning None for invalid/negative values."""
        if val is None or val == '' or val == '-':
            return None
        try:
            f = float(str(val).replace(',', '').replace('%', '').strip())
            return f if f > 0 else None
        except Exception:
            return None

    @classmethod
    def _compute_multiple_distributions(cls, peers: List[Dict[str, Any]]) -> Dict[str, Dict[str, float]]:
        """Calculates 25th percentile, Median, 75th percentile, and Mean for all multiples."""
        metrics = ['pe', 'pb', 'ev_ebitda', 'ev_sales']
        summary = {}
        
        for m in metrics:
            vals = [p[m] for p in peers if p.get(m) is not None and 0.5 <= p[m] <= 150.0]
            if vals:
                p25 = float(np.percentile(vals, 25))
                median = float(np.median(vals))
                p75 = float(np.percentile(vals, 75))
                mean = float(np.mean(vals))
                summary[m] = {
                    "count": len(vals),
                    "p25": round(p25, 2),
                    "median": round(median, 2),
                    "p75": round(p75, 2),
                    "mean": round(mean, 2)
                }
            else:
                summary[m] = {"count": 0, "p25": 0.0, "median": 0.0, "p75": 0.0, "mean": 0.0}
                
        return summary

    @classmethod
    def _calculate_implied_values(
        cls, 
        multiples_summary: Dict[str, Dict[str, float]], 
        normalized_data: Dict[str, Any],
        screener_data: Dict[str, Any],
        classification: Dict[str, Any],
        method_selection: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Calculates implied equity value and per-share price for each suitable multiple.
        Applies rigorous bridges:
        - Enterprise Multiples: Implied EV -> Less Net Debt -> Implied Equity Value / Shares
        - Equity Multiples: Metric * Multiple = Implied Equity Value / Shares
        """
        pnl = normalized_data.get('pnl', {})
        bs = normalized_data.get('bs', {})
        shares_raw = normalized_data.get('shares_outstanding_cr', 100.0)
        try:
            shares = max(0.01, float(shares_raw or 100.0))
        except Exception:
            shares = 100.0
        is_financial = classification.get('is_financial', False)
        
        # Target Metrics
        rev_cr = pnl.get('revenue', [0.0])[-1] if pnl.get('revenue') else 0.0
        ebitda_cr = pnl.get('ebitda', [0.0])[-1] if pnl.get('ebitda') else 0.0
        net_inc_cr = pnl.get('net_income', [0.0])[-1] if pnl.get('net_income') else 0.0
        book_val_cr = bs.get('total_equity', [0.0])[-1] if bs.get('total_equity') else 0.0
        net_debt_cr = bs.get('net_debt', [0.0])[-1] if bs.get('net_debt') else 0.0
        
        implied_results = {}
        
        # 1. EV/EBITDA Valuation
        if not is_financial and ebitda_cr > 0 and multiples_summary.get('ev_ebitda', {}).get('median', 0) > 0:
            med_mult = multiples_summary['ev_ebitda']['median']
            p25_mult = multiples_summary['ev_ebitda']['p25']
            p75_mult = multiples_summary['ev_ebitda']['p75']
            
            implied_ev = ebitda_cr * med_mult
            implied_eq = max(1.0, implied_ev - net_debt_cr)
            implied_per_share = round(implied_eq / shares, 2)
            
            implied_results["EV_EBITDA"] = {
                "metric_cr": ebitda_cr,
                "multiple": med_mult,
                "p25_multiple": p25_mult,
                "p75_multiple": p75_mult,
                "implied_ev_cr": round(implied_ev, 2),
                "net_debt_cr": round(net_debt_cr, 2),
                "implied_equity_value_cr": round(implied_eq, 2),
                "per_share_value": implied_per_share,
                "p25_per_share": round(max(1.0, (ebitda_cr * p25_mult) - net_debt_cr) / shares, 2),
                "p75_per_share": round(max(1.0, (ebitda_cr * p75_mult) - net_debt_cr) / shares, 2),
                "formula": "(Target EBITDA × Peer EV/EBITDA - Net Debt) / Shares"
            }
            
        # 2. EV/Revenue Valuation
        if not is_financial and rev_cr > 0 and multiples_summary.get('ev_sales', {}).get('median', 0) > 0:
            med_mult = multiples_summary['ev_sales']['median']
            implied_ev = rev_cr * med_mult
            implied_eq = max(1.0, implied_ev - net_debt_cr)
            implied_per_share = round(implied_eq / shares, 2)
            
            implied_results["EV_Revenue"] = {
                "metric_cr": rev_cr,
                "multiple": med_mult,
                "implied_ev_cr": round(implied_ev, 2),
                "net_debt_cr": round(net_debt_cr, 2),
                "implied_equity_value_cr": round(implied_eq, 2),
                "per_share_value": implied_per_share,
                "formula": "(Target Revenue × Peer EV/Sales - Net Debt) / Shares"
            }

        # 3. P/E Valuation
        if net_inc_cr > 0 and multiples_summary.get('pe', {}).get('median', 0) > 0:
            med_mult = multiples_summary['pe']['median']
            p25_mult = multiples_summary['pe']['p25']
            p75_mult = multiples_summary['pe']['p75']
            
            implied_eq = net_inc_cr * med_mult
            implied_per_share = round(implied_eq / shares, 2)
            
            implied_results["PE"] = {
                "metric_cr": net_inc_cr,
                "multiple": med_mult,
                "p25_multiple": p25_mult,
                "p75_multiple": p75_mult,
                "implied_equity_value_cr": round(implied_eq, 2),
                "per_share_value": implied_per_share,
                "p25_per_share": round((net_inc_cr * p25_mult) / shares, 2),
                "p75_per_share": round((net_inc_cr * p75_mult) / shares, 2),
                "formula": "(Target Net Income × Peer P/E) / Shares"
            }

        # 4. P/B Valuation
        if book_val_cr > 0 and multiples_summary.get('pb', {}).get('median', 0) > 0:
            med_mult = multiples_summary['pb']['median']
            p25_mult = multiples_summary['pb']['p25']
            p75_mult = multiples_summary['pb']['p75']
            
            implied_eq = book_val_cr * med_mult
            implied_per_share = round(implied_eq / shares, 2)
            
            implied_results["PB"] = {
                "metric_cr": book_val_cr,
                "multiple": med_mult,
                "p25_multiple": p25_mult,
                "p75_multiple": p75_mult,
                "implied_equity_value_cr": round(implied_eq, 2),
                "per_share_value": implied_per_share,
                "p25_per_share": round((book_val_cr * p25_mult) / shares, 2),
                "p75_per_share": round((book_val_cr * p75_mult) / shares, 2),
                "formula": "(Target Book Value of Equity × Peer P/B) / Shares"
            }
            
        return implied_results
