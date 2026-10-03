"""
universal_valuation/confidence_engine.py
========================================
Institutional Valuation Confidence Scoring Engine.
Evaluates data completeness, statement stability, forecast uncertainty, peer depth,
and terminal value sensitivity to assign a calibrated institutional confidence rating.
Never used as a disguised buy/sell recommendation.
"""

from typing import Dict, Any, List

class ValuationConfidenceEngine:
    """
    Computes objective multi-factor confidence rating for valuation models.
    """

    @classmethod
    def evaluate(
        cls, 
        normalized_data: Dict[str, Any],
        peer_data: Dict[str, Any],
        dcf_data: Dict[str, Any],
        cost_of_capital: Dict[str, Any],
        classification: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Assesses valuation confidence.
        
        Returns:
            Dict containing:
            - confidence_level: 'HIGH', 'MEDIUM', or 'LOW'
            - score_pct: float (0 - 100)
            - factor_scores: Dict of component scores
            - positive_factors: List[str]
            - risk_factors: List[str]
            - audit_rationale: str
        """
        score = 100.0
        positives = []
        risks = []
        factor_scores = {}
        
        # 1. Historical Data Completeness (Weight: 20 pts)
        num_years = normalized_data.get('num_years', 5)
        if num_years >= 7:
            factor_scores['data_completeness'] = 20.0
            positives.append(f"Robust historical financial statement depth ({num_years} reporting years available).")
        elif num_years >= 5:
            factor_scores['data_completeness'] = 16.0
            positives.append(f"Standard historical depth ({num_years} reporting years available).")
        else:
            factor_scores['data_completeness'] = 8.0
            score -= 12.0
            risks.append(f"Limited historical disclosure ({num_years} years); elevates forecast estimation error.")

        # 2. Earnings and Margin Stability (Weight: 20 pts)
        data_flags = normalized_data.get('data_quality_flags', [])
        pnl = normalized_data.get('pnl', {})
        neg_ebitda = any(v < 0 for v in pnl.get('ebitda', []))
        neg_net_inc = any(v < 0 for v in pnl.get('net_income', []))
        
        if not neg_ebitda and not neg_net_inc:
            factor_scores['earnings_stability'] = 20.0
            positives.append("Consistently positive historical operating profits and bottom-line earnings.")
        elif neg_ebitda:
            factor_scores['earnings_stability'] = 6.0
            score -= 14.0
            risks.append("Historical periods of negative operating profit (EBITDA) increase cash flow volatility.")
        else:
            factor_scores['earnings_stability'] = 12.0
            score -= 8.0
            risks.append("Intermittent historical net losses detected; normalized metrics required.")

        # 3. Peer Universe Quality & Breadth (Weight: 20 pts)
        peer_count = peer_data.get('peer_count', 0)
        if peer_count >= 5:
            factor_scores['peer_universe'] = 20.0
            positives.append(f"Deep peer universe ({peer_count} comparable industry companies available).")
        elif peer_count >= 2:
            factor_scores['peer_universe'] = 14.0
            score -= 6.0
            positives.append(f"Moderate peer universe ({peer_count} peers available).")
        else:
            factor_scores['peer_universe'] = 5.0
            score -= 15.0
            risks.append("Thin peer universe (<2 comparable peers); limits multiple-based cross-validation.")

        # 4. Terminal Value Sensitivity & Capital Structure (Weight: 20 pts)
        tv_pct = dcf_data.get('tv_pct_of_ev', 65.0)
        if tv_pct <= 75.0:
            factor_scores['tv_sensitivity'] = 20.0
            positives.append(f"Healthy terminal value proportion ({tv_pct:.1f}% of Enterprise Value).")
        elif tv_pct <= 85.0:
            factor_scores['tv_sensitivity'] = 15.0
            score -= 5.0
        else:
            factor_scores['tv_sensitivity'] = 8.0
            score -= 12.0
            risks.append(f"High terminal value dependence ({tv_pct:.1f}% of EV); model sensitivity is high.")

        # 5. Valuation Framework Appropriateness (Weight: 20 pts)
        is_financial = classification.get('is_financial', False)
        is_cyclical = classification.get('is_cyclical', False)
        
        if is_financial and dcf_data.get('valuation_model') == 'EXCESS_RETURN_MODEL':
            factor_scores['method_alignment'] = 20.0
            positives.append("Institutional Excess Return equity framework accurately deployed for financial institution.")
        elif is_cyclical and normalized_data.get('normalized_metrics', {}).get('normalization_applied'):
            factor_scores['method_alignment'] = 20.0
            positives.append("Mid-cycle margin normalization successfully dampens peak/trough cyclical distortions.")
        else:
            factor_scores['method_alignment'] = 18.0
            
        final_score = max(10.0, min(100.0, score))
        if final_score >= 80.0:
            conf_level = "HIGH"
        elif final_score >= 55.0:
            conf_level = "MEDIUM"
        else:
            conf_level = "LOW"
            
        rationale = (
            f"Overall model confidence rated {conf_level} ({final_score:.0f}/100). "
            f"Evaluated across {len(positives)} supporting pillars and {len(risks)} uncertainty factors."
        )

        return {
            "confidence_level": conf_level,
            "score_pct": round(final_score, 1),
            "factor_scores": factor_scores,
            "positive_factors": positives,
            "risk_factors": risks,
            "audit_rationale": rationale
        }
