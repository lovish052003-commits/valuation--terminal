"""
universal_valuation/company_classifier.py
=========================================
Universal Company Classification Engine (Phase 7).
Automatically classifies any listed company into one of 28 institutional categories
and valuation families using industry taxonomy, business description, and financial
statement balance-sheet/P&L characteristics WITHOUT hardcoding company names or tickers.
"""

import re
from typing import Dict, Any

class CompanyClassificationEngine:
    """
    Classifies any publicly listed company into 28 Institutional Categories (Phase 7):
     1. Industrial
     2. Consumer
     3. Technology
     4. Software
     5. Telecom
     6. Banking
     7. NBFC
     8. Insurance
     9. Financial services
    10. Real estate
    11. REIT
    12. InvIT
    13. Infrastructure
    14. Utilities
    15. Energy
    16. Oil & Gas
    17. Metals & Mining
    18. Chemicals
    19. Pharmaceuticals
    20. Healthcare
    21. Auto
    22. Auto ancillary
    23. FMCG
    24. Retail
    25. Conglomerate
    26. Holding company
    27. Cyclical commodity
    28. Other

    Also groups into Valuation Families for DCF vs Equity Models vs SOTP.
    """

    FINANCIAL_FAMILIES = {'BANK', 'NBFC', 'INSURANCE', 'REIT_INVIT', 'FINANCIAL_SERVICES'}
    CYCLICAL_FAMILIES = {'COMMODITY_CYCLICAL', 'ASSET_HEAVY_INDUSTRIAL', 'AUTO_ANCILLARY', 'REAL_ESTATE', 'INFRASTRUCTURE', 'METALS_MINING', 'ENERGY', 'OIL_AND_GAS'}

    @classmethod
    def classify(cls, screener_data: dict) -> dict:
        res = cls._classify_internal(screener_data)
        
        # Loss-Making Evaluation (Phase 8 & Phase 25)
        tables = screener_data.get('tables', {})
        pl_df = tables.get('profit-loss')
        is_loss_making = False
        valid_date_re = re.compile(r'(?:mar|jun|sep|dec|fy|\b20\d\d\b|ttm|12m)', re.I)
        exclude_col_re = re.compile(r'(?:scenario|case|growth|cagr|pct|%|ratio|turnover|note|desc|00:00:00|variance|diff|col_)', re.I)
        if pl_df is not None and hasattr(pl_df, 'empty') and not pl_df.empty:
            for _, row in pl_df.iterrows():
                m_str = str(row.get('Metric', '')).lower()
                if 'net profit' in m_str or 'pat' in m_str:
                    cand_cols = [c for c in pl_df.columns if c not in ['Metric', 'metric'] and valid_date_re.search(str(c)) and not exclude_col_re.search(str(c))]
                    for col in reversed(cand_cols):
                        v = str(row[col]).replace(',', '').strip()
                        try:
                            val = float(v)
                            is_loss_making = (val <= 0)
                            break
                        except (ValueError, TypeError):
                            pass
                    break
        res['is_loss_making'] = is_loss_making

        # Map to 9 Universal Canonical Company Types & Valuation Frameworks (Section 7)
        v_fam = str(res.get('valuation_family', '')).upper()
        c_type = str(res.get('company_type', '')).upper()
        sec = str(screener_data.get('sector') or res.get('sector') or '').upper()
        ind = str(screener_data.get('industry') or res.get('industry') or '').upper()
        name_t = f"{screener_data.get('company_name', '')} {screener_data.get('ticker', '')}".upper()

        if v_fam == 'BANK' or 'BANK' in c_type or 'BANK' in ind or 'BANK' in name_t:
            canonical_type = "Bank"
            canonical_framework = "Bank Equity Valuation"
            res['is_bank'] = True
            res['is_financial'] = True
        elif v_fam in ('NBFC', 'FINANCIAL_SERVICES') or 'NBFC' in c_type or 'FINANCE' in ind:
            canonical_type = "NBFC / Financial Services"
            canonical_framework = "Financial Services Equity Valuation"
            res['is_bank'] = False
            res['is_financial'] = True
        elif v_fam == 'INSURANCE' or 'INSURANCE' in c_type or 'INSURANCE' in ind:
            canonical_type = "Insurance"
            canonical_framework = "Insurance Equity / Embedded Value"
            res['is_bank'] = False
            res['is_financial'] = True
        elif v_fam == 'REIT_INVIT' or 'REIT' in c_type or 'INVIT' in c_type or 'REIT' in ind or 'INVIT' in ind:
            canonical_type = "REIT / InvIT"
            canonical_framework = "Yield / Distribution Valuation"
            res['is_bank'] = False
            res['is_financial'] = False
        elif v_fam in ('CONSUMER_STABLE', 'CONSUMER_FMCG', 'CONSUMER_DISCRETIONARY', 'RETAIL') or any(k in sec for k in ['CONSUMER', 'FMCG', 'RETAIL', 'FOOD', 'BEVERAGE']):
            canonical_type = "Consumer"
            canonical_framework = "FCFF / Enterprise Valuation"
            res['is_bank'] = False
            res['is_financial'] = False
        elif v_fam in ('TECH_SAAS', 'IT_SERVICES', 'TELECOM') or any(k in sec for k in ['TECH', 'SOFTWARE', 'TELECOM', 'INFORMATION TECH']) or bool(re.search(r'\bIT\b', sec)):
            canonical_type = "IT / Services"
            canonical_framework = "FCFF / Enterprise Valuation"
            res['is_bank'] = False
            res['is_financial'] = False
        elif v_fam in ('UTILITIES_REGULATED', 'ENERGY', 'OIL_AND_GAS') or any(k in sec for k in ['ENERGY', 'OIL', 'GAS', 'POWER', 'UTILITIES']):
            canonical_type = "Energy / Utilities"
            canonical_framework = "FCFF / Enterprise Valuation"
            res['is_bank'] = False
            res['is_financial'] = False
        elif v_fam in ('ASSET_HEAVY_INDUSTRIAL', 'METALS_MINING', 'AUTO_OEM', 'AUTO_ANCILLARY', 'CHEMICALS', 'PHARMA_HEALTHCARE', 'REAL_ESTATE', 'INFRASTRUCTURE', 'COMMODITY_CYCLICAL', 'CONGLOMERATE') or any(k in sec for k in ['INDUSTRIAL', 'METAL', 'MINING', 'AUTO', 'CHEMICAL', 'PHARMA', 'HEALTH', 'CONSTRUCTION', 'STEEL']):
            canonical_type = "Industrial / Manufacturing"
            canonical_framework = "FCFF / Enterprise Valuation"
            res['is_bank'] = False
            res['is_financial'] = False
        else:
            canonical_type = "Other / Unknown"
            canonical_framework = "FCFF / Enterprise Valuation"
            res['is_bank'] = False
            res['is_financial'] = False

        res['canonical_company_type'] = canonical_type
        res['canonical_valuation_framework'] = canonical_framework
        res['company_type'] = canonical_type
        res['valuation_framework'] = canonical_framework
        return res

    @classmethod
    def _classify_internal(cls, screener_data: dict) -> dict:
        sector = str(screener_data.get('sector') or '').strip()
        industry = str(screener_data.get('industry') or '').strip()
        about = str(screener_data.get('about') or '').strip()
        company_name = str(screener_data.get('company_name') or '').strip()
        ticker = str(screener_data.get('ticker') or '').strip()
        
        combined_text = f"{company_name} {ticker} {sector} {industry} {about}".lower()

        # 1. Structural Financial Statement Signature Analysis
        tables = screener_data.get('tables', {})
        pl_df = tables.get('profit-loss')
        bs_df = tables.get('balance-sheet')
        
        has_financing_profit = False
        has_interest_earned = False
        has_deposits = False
        has_loans_advances = False
        has_net_premiums = False
        has_gross_block = False
        has_inventory = False
        is_financial_structural = False

        if pl_df is not None and hasattr(pl_df, 'empty') and not pl_df.empty:
            metrics_pl = " ".join([str(m) for m in pl_df.get('Metric', [])]).lower()
            has_financing_profit = 'financing profit' in metrics_pl
            has_interest_earned = 'interest earned' in metrics_pl or 'interest income' in metrics_pl
            has_net_premiums = 'premiums earned' in metrics_pl or 'premium' in metrics_pl

        if bs_df is not None and hasattr(bs_df, 'empty') and not bs_df.empty:
            bs_rows = [str(m).strip().lower() for m in bs_df.get('Metric', [])]
            # Strip out ratio rows appended below BS
            asset_rows = []
            for m in bs_rows:
                if any(r_word in m for r_word in ['turnover', 'debtor days', 'return on', 'roce', 'roe', 'working capital']):
                    break
                asset_rows.append(m)
            metrics_bs = " ".join(asset_rows)
            has_deposits = 'deposits' in metrics_bs
            has_loans_advances = 'advances' in metrics_bs or 'loans' in metrics_bs
            has_gross_block = 'fixed assets' in metrics_bs or 'net block' in metrics_bs or 'cwip' in metrics_bs
            has_inventory = any(m in ['inventories', 'inventory', 'raw materials', 'stock in trade'] for m in asset_rows)

            # Quantitative Capital Structure Ratio Check
            try:
                borrowings_val = 0.0
                total_assets_val = 0.0
                net_block_val = 0.0
                for _, b_row in bs_df.iterrows():
                    m_label = str(b_row.get('Metric', '')).strip().lower()
                    for col in reversed([c for c in bs_df.columns if c not in ['Metric', 'metric']]):
                        v_str = str(b_row[col]).replace(',', '').strip()
                        try:
                            v_num = float(v_str)
                            if 'borrowing' in m_label:
                                borrowings_val = v_num
                            elif m_label in ('total', 'total assets', 'total liabilities'):
                                total_assets_val = max(total_assets_val, v_num)
                            elif 'fixed assets' in m_label or 'net block' in m_label:
                                net_block_val = v_num
                            break
                        except Exception:
                            continue
                if total_assets_val > 0:
                    d_ratio = borrowings_val / total_assets_val
                    fa_ratio = net_block_val / total_assets_val
                    if d_ratio > 0.40 and fa_ratio < 0.15:
                        is_financial_structural = True
            except Exception:
                pass

        # 2. Rule-Based Classification Hierarchy (Structural + Semantic)
        
        # A. Conglomerate & Multi-Business Incubator (Phase 16 SOTP Candidate)
        conglomerate_keywords = [
            'conglomerate', 'multi-business', 'incubator', 'diversified operations',
            'airports, roads', 'multi business', 'diversified industrial'
        ]
        is_conglomerate_explicit = (
            any(k in combined_text for k in conglomerate_keywords) 
            or ('enterprises' in company_name.lower() and not any(k in combined_text for k in ['software', 'technology', 'bank', 'nbfc', 'insurance', 'pharma', 'hospital']))
            or ('diversified' in sector.lower() and ('infrastructure' in combined_text or 'incubator' in combined_text or 'enterprises' in combined_text))
            or (sector.lower() in ['diversified', 'conglomerates', 'trading'] and any(k in combined_text for k in ['energy', 'airports', 'roads', 'mining', 'ports', 'metals', 'enterprises']))
        )

        if is_conglomerate_explicit:
            return {
                "sector": sector or "Diversified",
                "industry": industry or "Conglomerate",
                "business_model": "Multi-Segment Conglomerate & Incubator Platform",
                "company_type": "Conglomerate",
                "valuation_family": "CONGLOMERATE",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": True,
                "is_holding_co": False,
                "requires_sotp": True,
                "confidence": "HIGH",
                "classification_rationale": "Identified as multi-segment conglomerate operating distinct business divisions (Airports/Roads/Mining/Energy), requiring SOTP alongside consolidated DCF."
            }

        # Disambiguate Bank vs NBFC vs Insurance (Must evaluate before Holding Company to avoid misclassifying commercial banks)
        nbfc_explicit = any(k in combined_text for k in [
            'nbfc', 'non banking', 'non-banking', 'housing finance', 'microfinance',
            'consumer finance', 'gold loan', 'vehicle finance', 'asset finance', 'commercial finance'
        ])

        # Banking (Commercial Banking / Deposit-Taking Lending)
        bank_keywords = ['commercial bank', 'private sector bank', 'public sector bank', 'small finance bank', 'banking', ' bank ']
        is_bank_semantic = (any(k in combined_text for k in bank_keywords) or combined_text.strip().endswith(' bank') or 'bank' in company_name.lower().split() or 'bank' in ticker.lower()) and not any(k in combined_text for k in ['blood bank', 'food bank', 'data bank', 'non banking', 'non-banking', 'investment bank'])
        
        if (has_deposits or is_bank_semantic) and not nbfc_explicit:
            return {
                "sector": sector or "Banking",
                "industry": industry or "Banks",
                "business_model": "Commercial Banking / Deposit-Taking Lending",
                "company_type": "Bank",
                "valuation_family": "BANK",
                "is_financial": True,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via banking financial statement architecture (Deposits / Financing Profit / Interest Earned) and regulatory banking taxonomy."
            }

        # B. Holding Company / Investment Trust (Checked after genuine banks)
        holding_keywords = ['holding company', 'investment trust', 'core investment company']
        has_holding_kw = any(k in combined_text for k in holding_keywords) or bool(re.search(r'\bcic\b', combined_text))
        if has_holding_kw and not has_deposits and not has_gross_block:
            return {
                "sector": sector or "Financial Services",
                "industry": industry or "Holding Company",
                "business_model": "Equity Investment Holding & Capital Allocation",
                "company_type": "Holding company",
                "valuation_family": "HOLDING_COMPANY",
                "is_financial": True,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": True,
                "requires_sotp": True,
                "confidence": "HIGH",
                "classification_rationale": "Identified via equity holding/treasury asset structure, requiring NAV and holding company discount framework."
            }

        # Insurance
        ins_keywords = ['insurance', 'life insurance', 'general insurance', 'reinsurance', 'assurance', 'policyholder']
        if has_net_premiums or any(k in combined_text for k in ins_keywords):
            return {
                "sector": sector or "Financial Services",
                "industry": industry or "Insurance",
                "business_model": "Underwriting / Premium Investment Float",
                "company_type": "Insurance",
                "valuation_family": "INSURANCE",
                "is_financial": True,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via insurance underwriting operations and regulatory insurance taxonomy."
            }

        # NBFC & Financial Services (Spread Lending / No Deposits)
        nbfc_keywords = [
            'nbfc', 'non banking', 'non-banking', 'housing finance', 'microfinance', 
            'consumer finance', 'gold loan', 'vehicle finance', 'asset finance', 
            'lending', 'financial services', 'finance'
        ]
        nbfc_ticker_marker = any(ticker.upper().endswith(suf) for suf in ['HFL', 'FIN', 'CAP', 'CREDIT', 'LEND']) or any(k in ticker.upper() for k in ['FINANCE', 'CAPITAL', 'HOLDING'])
        if (has_financing_profit and not has_deposits) or (has_interest_earned and not has_deposits) or nbfc_explicit or nbfc_ticker_marker or is_financial_structural or any(k in combined_text for k in nbfc_keywords):
            return {
                "sector": sector or "Financial Services",
                "industry": industry or "NBFC / Lending",
                "business_model": "Non-Banking Financial Intermediation",
                "company_type": "NBFC",
                "valuation_family": "NBFC",
                "is_financial": True,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via wholesale/retail lending operations, non-deposit financing liabilities, and NBFC taxonomy."
            }

        # REIT / InvIT
        reit_keywords = ['reit', 'invit', 'real estate investment trust', 'infrastructure investment trust', 'distributable cash flow']
        if any(k in combined_text for k in reit_keywords):
            c_type = "REIT" if "reit" in combined_text else "InvIT"
            return {
                "sector": sector or "Real Estate / Infrastructure",
                "industry": industry or "REIT / InvIT",
                "business_model": "Yield-Generating Asset Trust",
                "company_type": c_type,
                "valuation_family": "REIT_INVIT",
                "is_financial": False,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via REIT/InvIT trust structure and asset-backed distributable yield architecture."
            }

        # Metals & Mining / Cyclical Commodity
        metals_keywords = ['steel', 'iron', 'metal', 'mining', 'mineral', 'aluminum', 'copper', 'zinc', 'coal mining']
        if any(k in combined_text for k in metals_keywords):
            return {
                "sector": sector or "Metals & Mining",
                "industry": industry or "Metals & Mining",
                "business_model": "Commodity / Metals Mining Capital-Intensive Producer",
                "company_type": "Metals & Mining",
                "valuation_family": "COMMODITY_CYCLICAL",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via metals extraction, mining concession footprint, and cyclical commodity pricing."
            }

        # Oil & Gas / Energy / Refinery
        oil_keywords = ['oil & gas', 'refinery', 'refining', 'petrochemical', 'upstream', 'crude oil', 'natural gas', 'petroleum']
        if any(k in combined_text for k in oil_keywords):
            return {
                "sector": sector or "Energy",
                "industry": industry or "Oil & Gas",
                "business_model": "Hydrocarbon Exploration, Refining & Petrochemicals",
                "company_type": "Oil & Gas",
                "valuation_family": "COMMODITY_CYCLICAL",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via oil & gas refining, upstream extraction, and energy commodity exposure."
            }

        # Cement & Cyclical Commodity
        if 'cement' in combined_text or 'clinker' in combined_text:
            return {
                "sector": sector or "Materials",
                "industry": industry or "Cement",
                "business_model": "Capital-Intensive Building Materials Producer",
                "company_type": "Cyclical commodity",
                "valuation_family": "COMMODITY_CYCLICAL",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via cement manufacturing, heavy kilns, and construction material cycles."
            }

        # Chemicals & Agrochemicals
        chemicals_keywords = ['chemical', 'chemicals', 'speciality chemical', 'agrochemical', 'fertilizer', 'pesticide']
        if any(k in combined_text for k in chemicals_keywords):
            return {
                "sector": sector or "Chemicals",
                "industry": industry or "Specialty Chemicals",
                "business_model": "Chemical Synthesis & Formulation",
                "company_type": "Chemicals",
                "valuation_family": "OPERATING_COMPANY",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via specialty chemical synthesis, batch manufacturing, and chemical processing."
            }

        # Technology & Software (Tech SaaS / Services)
        software_keywords = ['software', 'saas', 'cloud', 'digital platform', 'application software']
        tech_keywords = ['it services', 'information technology', 'consultancy services', 'data analytics', 'cybersecurity']
        if any(k in combined_text for k in software_keywords) or any(k in combined_text for k in tech_keywords):
            c_type = "Software" if any(k in combined_text for k in software_keywords) else "Technology"
            return {
                "sector": sector or "Technology",
                "industry": industry or "IT Services / Software",
                "business_model": "Human Capital & Intellectual Property Services",
                "company_type": c_type,
                "valuation_family": "TECH_SAAS",
                "is_financial": False,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via low tangible capital intensity, intellectual property/software services, and technology taxonomy."
            }

        # Pharmaceuticals & Healthcare
        pharma_keywords = ['pharma', 'pharmaceutical', 'drug', 'medicine', 'active pharmaceutical', 'formulation', 'biotechnology']
        health_keywords = ['healthcare', 'hospital', 'diagnostics', 'clinic', 'medical devices']
        if any(k in combined_text for k in pharma_keywords) or any(k in combined_text for k in health_keywords):
            c_type = "Pharmaceuticals" if any(k in combined_text for k in pharma_keywords) else "Healthcare"
            return {
                "sector": sector or "Healthcare",
                "industry": industry or "Pharmaceuticals & Healthcare",
                "business_model": "R&D, Formulations & Healthcare Delivery",
                "company_type": c_type,
                "valuation_family": "PHARMA_HEALTHCARE",
                "is_financial": False,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via drug development/manufacturing, healthcare facilities, and life-sciences taxonomy."
            }

        # FMCG & Consumer Staples
        fmcg_keywords = ['fmcg', 'consumer goods', 'beverage', 'food', 'dairy', 'personal care', 'household', 'tobacco', 'cigarettes', 'packaged food']
        if any(k in combined_text for k in fmcg_keywords):
            return {
                "sector": sector or "Consumer Staples",
                "industry": industry or "FMCG / Consumer",
                "business_model": "Brand-Driven Consumer Goods & Distribution",
                "company_type": "FMCG",
                "valuation_family": "CONSUMER_FMCG",
                "is_financial": False,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via consumer retail distribution, brand franchise equity, and FMCG business model."
            }

        # Retail
        retail_keywords = ['retail', 'supermarket', 'hypermarket', 'department store', 'e-commerce', 'apparel retail', 'fashion retail']
        if any(k in combined_text for k in retail_keywords):
            return {
                "sector": sector or "Consumer Discretionary",
                "industry": industry or "Retail",
                "business_model": "Store Network & Omnichannel Retail Distribution",
                "company_type": "Retail",
                "valuation_family": "OPERATING_COMPANY",
                "is_financial": False,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via store front retail merchandising, working capital turn, and retail footfall."
            }

        # Utilities & Power
        utility_keywords = ['utility', 'utilities', 'power generation', 'power transmission', 'electricity', 'water utility', 'city gas', 'renewable energy', 'solar power', 'wind power']
        if any(k in combined_text for k in utility_keywords):
            c_type = "Energy" if any(k in combined_text for k in ['renewable energy', 'solar power', 'wind power']) else "Utilities"
            return {
                "sector": sector or "Utilities",
                "industry": industry or "Power / Utilities",
                "business_model": "Regulated Utility Asset Operator",
                "company_type": c_type,
                "valuation_family": "UTILITIES",
                "is_financial": False,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via regulated rate-base return model and utility service infrastructure."
            }

        # Telecom
        telecom_keywords = ['telecom', 'telecommunications', 'wireless', 'cellular', 'mobile network', 'telecom towers']
        if any(k in combined_text for k in telecom_keywords):
            return {
                "sector": sector or "Telecommunications",
                "industry": industry or "Telecom",
                "business_model": "Network Infrastructure & Subscriber Subscription",
                "company_type": "Telecom",
                "valuation_family": "TELECOM",
                "is_financial": False,
                "is_cyclical": False,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via high spectrum/network capex and subscriber subscription revenue."
            }

        # Infrastructure & Construction
        infra_keywords = ['infrastructure', 'construction', 'engineering', 'roads', 'highways', 'ports', 'airports', 'bridges', 'epc']
        if any(k in combined_text for k in infra_keywords):
            return {
                "sector": sector or "Infrastructure",
                "industry": industry or "EPC & Infrastructure",
                "business_model": "Contracting, Concession & Long-Duration Asset Construction",
                "company_type": "Infrastructure",
                "valuation_family": "INFRASTRUCTURE",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via long-cycle EPC contracts, project concession assets, and civil construction."
            }

        # Automobile & Ancillaries
        auto_keywords = [
            'automobile', 'auto ancillary', 'vehicles', 'vehicle', 'automotive', 
            'passenger cars', 'commercial vehicles', 'passenger vehicles', 'utility vehicles',
            'two wheeler', 'four wheeler', 'auto components', 'tyres', 'motors', 'motor', 
            'motocorp', 'trucks', 'buses', 'tractors'
        ]
        is_auto_word = bool(re.search(r'\bauto\b', combined_text))
        if any(k in combined_text for k in auto_keywords) or is_auto_word:
            c_type = "Auto ancillary" if any(k in combined_text for k in ['ancillary', 'components', 'tyres']) else "Auto"
            return {
                "sector": sector or "Automotive",
                "industry": industry or "Automobiles & Components",
                "business_model": "OEM Vehicle Manufacturing & Ancillary Engineering",
                "company_type": c_type,
                "valuation_family": "AUTO_ANCILLARY",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via automotive OEM assembly, component engineering, and discretionary auto cycles."
            }

        # Real Estate Development
        re_keywords = ['real estate', 'property development', 'residential developer', 'commercial developer', 'housing projects']
        if any(k in combined_text for k in re_keywords):
            return {
                "sector": sector or "Real Estate",
                "industry": industry or "Real Estate Development",
                "business_model": "Land Bank Acquisition & Property Development",
                "company_type": "Real estate",
                "valuation_family": "REAL_ESTATE",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "HIGH",
                "classification_rationale": "Identified via land-bank inventory, pre-sales booking cycles, and real estate construction."
            }

        # Asset-Heavy Industrial Manufacturing (Fallback for capital goods, machinery, electricals)
        if not is_financial_structural and has_gross_block and has_inventory:
            return {
                "sector": sector or "Industrials",
                "industry": industry or "Industrial Manufacturing",
                "business_model": "Asset-Heavy Manufacturing & Equipment",
                "company_type": "Industrial",
                "valuation_family": "ASSET_HEAVY_INDUSTRIAL",
                "is_financial": False,
                "is_cyclical": True,
                "is_conglomerate": False,
                "is_holding_co": False,
                "requires_sotp": False,
                "confidence": "MEDIUM",
                "classification_rationale": "Identified via significant fixed assets (gross block) and working capital inventory footprint."
            }

        # General Operating Company (Catch-all)
        return {
            "sector": sector or "Diversified Operations",
            "industry": industry or "Operating Company",
            "business_model": "Commercial Operating Company",
            "company_type": "Other",
            "valuation_family": "OPERATING_COMPANY",
            "is_financial": False,
            "is_cyclical": False,
            "is_conglomerate": False,
            "is_holding_co": False,
            "requires_sotp": False,
            "confidence": "MEDIUM",
            "classification_rationale": "Defaulted to standard operating corporate valuation architecture based on non-financial operational statements."
        }

class ClassificationResult(dict):
    """Dictionary subclass providing dot-notation attribute access."""
    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError:
            return None

    def __setattr__(self, name, value):
        self[name] = value

def classify_company(screener_data: dict) -> ClassificationResult:
    """Universal helper function that returns dot-accessible classification info."""
    res = CompanyClassificationEngine.classify(screener_data)
    return ClassificationResult(res)

