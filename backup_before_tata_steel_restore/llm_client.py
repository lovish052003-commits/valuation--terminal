"""
llm_client.py
Universal Multi-Provider AI Equity Research & Valuation Report Generator.
Supports:
- NVIDIA NIM (https://integrate.api.nvidia.com/v1)
- OpenRouter (https://openrouter.ai/api/v1)
- OpenAI (https://api.openai.com/v1)
- Google Gemini (google-genai SDK)
- Custom OpenAI-Compatible Endpoints (Groq, Together, Ollama, LM Studio, etc.)
"""

import os
from openai import OpenAI

# Provider configuration defaults
PROVIDER_CONFIGS = {
    'nvidia': {
        'name': 'NVIDIA NIM',
        'base_url': 'https://integrate.api.nvidia.com/v1',
        'default_model': 'meta/llama-3.3-70b-instruct',
        'fallback_models': [
            'meta/llama-3.3-70b-instruct',
            'deepseek-ai/deepseek-r1',
            'nvidia/llama-3.1-nemotron-70b-instruct',
            'meta/llama-3.1-70b-instruct',
            'meta/llama-3.1-8b-instruct'
        ]
    },
    'openrouter': {
        'name': 'OpenRouter',
        'base_url': 'https://openrouter.ai/api/v1',
        'default_model': 'deepseek/deepseek-r1',
        'fallback_models': [
            'deepseek/deepseek-r1',
            'meta-llama/llama-3.3-70b-instruct',
            'anthropic/claude-3.5-sonnet',
            'google/gemini-2.0-flash-001',
            'openai/gpt-4o-mini'
        ]
    },
    'openai': {
        'name': 'OpenAI',
        'base_url': 'https://api.openai.com/v1',
        'default_model': 'gpt-4o',
        'fallback_models': ['gpt-4o', 'gpt-4o-mini', 'o3-mini']
    },
    'gemini': {
        'name': 'Google Gemini',
        'default_model': 'gemini-2.5-flash',
        'fallback_models': ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash']
    },
    'custom': {
        'name': 'Custom Endpoint',
        'base_url': '',
        'default_model': '',
        'fallback_models': []
    }
}

SYSTEM_PROMPT = """You are an elite Senior Equity Research Analyst and Valuation Expert at a premier global investment bank.
You are provided with verified financial statements from Screener.in and exact mathematical models computed by the institutional Universal Valuation Engine.

Your goal is to synthesize these quantitative results into an authoritative, institutional-grade Equity Research Memorandum.
CRITICAL MANDATES:
1. DO NOT generate automatic BUY, SELL, or HOLD ratings. The model remains strictly analytical.
2. Present the mathematical VALUATION GAP (Discount or Premium to Current Market Price) and explain valuation uncertainty.
3. NEVER modify or invent underlying numerical metrics. Only interpret the verified outputs provided.

The title of your memorandum must strictly begin with:
# {companyName} ({ticker}) - Institutional Valuation Memorandum

## 1. Executive Summary & Valuation Gap Analysis
- Intrinsic Base Value vs Current Market Price (CMP)
- Valuation Gap: Percentage Discount (Margin of Safety) or Premium
- Analytical interpretation of valuation uncertainty across market scenarios

---

## 2. Company Profile & Business Architecture
- Business model, industry categorization, and revenue drivers
- Classification rationale and reporting basis (Consolidated vs Standalone)

---

## 3. Valuation Methodology & Method Selection
- Primary and secondary valuation architectures deployed (FCFF DCF vs Excess Return Model vs Comps)
- Justification of selected methods and explicit reasons for excluded methodologies

---

## 4. Key Valuation Drivers & Assumptions
- Top-line revenue trajectory, normalized margins, and mid-cycle adjustments
- Fundamental Reinvestment Rate (Expected Growth / Sustainable ROIC) and explicit fade trajectory
- Dynamic Cost of Capital (WACC / Cost of Equity Ke): Risk-Free Rate, ERP, Beta, Capital Structure

---

## 5. Discounted Cash Flow / Intrinsic Valuation
- Summary of explicit cash flow forecast schedule and present values
- Terminal Value determination, Terminal ROIC convergence, and % share of Enterprise Value

---

## 6. Comparable Company Multiple Analysis
- Peer group selection criteria and strict self-exclusion validation
- Peer distribution (P25, Median, P75) across EV/EBITDA, EV/Sales, P/E, and P/B
- Implied per-share values derived from relative market pricing

---

## 7. Valuation Reconciliation & Dispersion Diagnosis
- Comparison across intrinsic DCF and relative multiples
- Quantified dispersion % and diagnosis of divergence between cash flow intrinsic value and multiple pricing
- Central valuation range and multi-scenario bounds (Bear, Base, Bull)

---

## 8. 2D Sensitivity Analysis
- WACC × Terminal Growth sensitivity matrix interpretation
- Growth × Margin sensitivity matrix interpretation

---

## 9. Key Catalysts & Downside Risks
- Top 3 catalysts for fundamental value expansion
- Critical operational, regulatory, commodity/cyclical, or leverage risks

---

## 10. Data Quality, Integrity & Model Confidence
- Audit of financial statement depth, margin stability, and accounting flags
- Institutional confidence rating (High / Medium / Low) and quality scorecard
"""

def build_valuation_context(screener_data, valuation_result):
    """Formats the company's financial data and valuation results into a clean markdown prompt context."""
    company_name = screener_data['company_name']
    ticker = screener_data['ticker']
    current_price = screener_data['current_price']
    market_cap = screener_data['market_cap_cr']
    meta = screener_data['meta_raw']
    about = screener_data.get('about', '')
    sector = screener_data.get('sector', '')
    industry = screener_data.get('industry', '')

    v = valuation_result
    
    # Financial tables summary
    tables = screener_data.get('tables', {})
    pl_md = tables['profit-loss'].to_markdown(index=False) if 'profit-loss' in tables else "N/A"
    bs_md = tables['balance-sheet'].to_markdown(index=False) if 'balance-sheet' in tables else "N/A"
    cf_md = tables['cash-flow'].to_markdown(index=False) if 'cash-flow' in tables else "N/A"
    peers_md = screener_data['peers_df'].to_markdown(index=False) if not screener_data['peers_df'].empty else "N/A"

    # DCF Table markdown
    dcf_rows = "\n".join([
        f"| {row['year']} | ₹{row['ebit']} | ₹{row['nopat']} | {row['reinvestment_rate']}% | ₹{row['fcff']} | {row['discount_factor']} | ₹{row['pv_fcff']} |"
        for row in v['dcf_table']
    ])
    dcf_table_md = (
        "| Year | EBIT (Cr) | NOPAT (Cr) | Reinvest Rate | FCFF (Cr) | Discount Factor | PV of FCFF (Cr) |\n"
        "|---|---|---|---|---|---|---|\n"
        + dcf_rows
    )

    # Sensitivity matrix markdown
    sens_headers = " | ".join(v['sensitivity']['tg_headers'])
    sens_div = " | ".join(["---"] * len(v['sensitivity']['tg_headers']))
    sens_rows = "\n".join([
        f"| WACC {r['wacc']} | " + " | ".join([f"₹{val}" if val is not None else "N/A" for val in r['values']]) + " |"
        for r in v['sensitivity']['rows']
    ])
    sens_table_md = f"| WACC \\ Terminal Growth | {sens_headers} |\n|---|{sens_div}|\n{sens_rows}"

    fp = v.get('four_pillars', {})
    p1 = fp.get('pillar1_fcf', {})
    p2 = fp.get('pillar2_growth', {})
    p3 = fp.get('pillar3_wacc', {})
    p4 = fp.get('pillar4_multiples', {})

    prompt = f"""
### COMPANY PROFILE:
- **Company Name**: {company_name} ({ticker})
- **Sector / Industry**: {sector} / {industry}
- **Current Market Price (CMP)**: ₹{current_price}
- **Market Capitalization**: ₹{market_cap} Cr
- **Shares Outstanding**: {v['shares_cr']} Cr shares
- **About**: {about}

### 4 VALUATION PILLARS (QUANTITATIVE ENGINE OUTPUTS):

#### PILLAR 1: FREE CASH FLOW GENERATION (THE ENGINE)
- **Operating Cash Flow (CFO)**: ₹{p1.get('cfo', 0)} Cr vs **Net Profit (PAT)**: ₹{p1.get('net_profit', 0)} Cr
- **Earnings Quality Ratio (CFO/PAT)**: {p1.get('earnings_quality_ratio', 1.0)}x ({'High Cash Backing' if p1.get('earnings_quality_ratio', 1) >= 1.0 else 'Accrual / Working Capital Drag'})
- **Operating Margins**: EBITDA Margin = {p1.get('ebitda_margin', 0)}%, EBIT Margin = {p1.get('ebit_margin', 0)}%
- **Capital Expenditures**: Total CapEx = ₹{p1.get('total_capex', 0)} Cr (Maintenance CapEx ~ ₹{p1.get('maintenance_capex', 0)} Cr, Growth CapEx ~ ₹{p1.get('growth_capex', 0)} Cr)
- **Actual Free Cash Flow (CFO - CapEx)**: ₹{p1.get('actual_fcf', 0)} Cr
- **Working Capital Efficiency**:
  - Debtor Days (Receivables): {p1.get('debtor_days', 'N/A')} days
  - Inventory Days: {p1.get('inventory_days', 'N/A')} days
  - Payable Days: {p1.get('payable_days', 'N/A')} days
  - Cash Conversion Cycle (CCC): {p1.get('cash_conversion_cycle', 'N/A')} days

#### PILLAR 2: GROWTH ASSUMPTIONS (THE TRAJECTORY)
- **Historical Sales CAGR**: 3-Year CAGR = {p2.get('sales_cagr_3yr', 'N/A')}%, 5-Year CAGR = {p2.get('sales_cagr_5yr', 'N/A')}%
- **Historical Net Profit CAGR (3Y)**: {p2.get('pat_cagr_3yr', 'N/A')}%
- **Explicit 5-Year Forecast EBIT Growth Rate**: {p2.get('explicit_growth_rate', v['growth_rate'])}%
- **Terminal Value Perpetual Growth Rate (g)**: {p2.get('terminal_growth_rate', v['terminal_growth'])}% (Pegged to long-term GDP / inflation)
- **Terminal Value**: ₹{p2.get('terminal_value', v['terminal_value'])} Cr (PV of TV: ₹{p2.get('pv_terminal_value', v['pv_terminal_value'])} Cr)
- **Terminal Value Share of Enterprise Value**: {p2.get('tv_share_of_ev', 'N/A')}% of total EV

#### PILLAR 3: COST OF CAPITAL & RISK (THE DISCOUNT RATE - WACC)
- **Calculated WACC**: {p3.get('wacc', v['wacc'])}%
- **Cost of Equity (Ke)**: {p3.get('cost_of_equity', v['cost_of_equity'])}% (Risk-Free Rate Rf = 7.0%, Equity Risk Premium ERP = 6.5%, Beta = {p3.get('beta', v['beta'])})
- **Cost of Debt (Kd post-tax)**: {p3.get('post_tax_cost_of_debt', v['cost_of_debt'])}% (Pre-tax Kd = {p3.get('pre_tax_cost_of_debt', '8.05')}%, Effective Tax Rate = {p3.get('tax_rate', v['tax_rate'])}%)
- **Capital Structure Weights**: Equity Weight = {p3.get('weight_equity', v['weight_equity'])}%, Debt Weight = {p3.get('weight_debt', v['weight_debt'])}%
- **Debt-to-Equity Ratio (D/E)**: {p3.get('debt_to_equity', 0.0)}x

#### PILLAR 4: RELATIVE MARKET MULTIPLES (THE SANITY CHECK)
- **Target Company Multiples**:
  - EV/EBITDA: {p4.get('target_ev_ebitda', 'N/A')}x
  - P/E Ratio: {p4.get('target_pe', 'N/A')}x
  - Price-to-Book (P/B): {p4.get('target_pb', 'N/A')}x
  - EV/Sales: {p4.get('target_ev_sales', 'N/A')}x
- **Peer Benchmark Multiples**:
  - Peer Median P/E: {p4.get('peer_median_pe', 'N/A')}x
  - Peer Median EV/EBITDA: {p4.get('peer_median_ev_ebitda', 'N/A')}x
- **Cross-Check Implied Share Prices**:
  - DCF Intrinsic Target: ₹{v['intrinsic_value_per_share']} (Margin of Safety: {v['margin_of_safety_pct']}%)
  - Peer Median P/E Implied Price: ₹{p4.get('implied_price_pe', 'N/A')}
  - Peer Median EV/EBITDA Implied Price: ₹{p4.get('implied_price_ev_ebitda', 'N/A')}
  - 52-Week Price Range: ₹{p4.get('low_52wk', 'N/A')} - ₹{p4.get('high_52wk', 'N/A')} (CMP: ₹{current_price})

### 5-YEAR DCF CASH FLOW PROJECTION SCHEDULE:
{dcf_table_md}

### VALUATION SENSITIVITY MATRIX (WACC vs Terminal Growth):
{sens_table_md}

### DUPONT & FINANCIAL HEALTH RESILIENCE:
- **DuPont 3-Stage ROE**: {v['dupont']['roe_3stage']}% (Net Margin: {v['dupont']['net_profit_margin']}%, Asset Turnover: {v['dupont']['asset_turnover']}x, Equity Multiplier: {v['dupont']['equity_multiplier']}x)
- **DuPont 5-Stage ROE**: {v['dupont']['roe_5stage']}% (Operating Margin: {v['dupont']['operating_margin']}%, Interest Burden: {v['dupont']['interest_burden']}x, Tax Burden: {v['dupont']['tax_burden']}x)
- **Return on Assets (ROA)**: {v['dupont'].get('roa', 'N/A')}%
- **Altman Z-Score**: {v['altman_z']['display_text'] if 'display_text' in v['altman_z'] else (f"{v['altman_z']['score']} ({v['altman_z']['zone']})" if v['altman_z'].get('score') is not None else "Altman's Z-Score: Not applicable / insufficient data")}

### DATA SOURCE & PROVENANCE:
- **Data Source**: {screener_data.get('data_source', 'Screener.in')}
- **Data Timestamp**: {screener_data.get('data_as_of', 'Live')}
- **Run ID**: {v.get('run_id', 'N/A')}

### HISTORICAL STATEMENTS (SCREENER.IN):
#### 1. Profit & Loss:
{pl_md}
#### 2. Balance Sheet:
{bs_md}
#### 3. Cash Flows:
{cf_md}
#### 4. Peer Comparison Table:
{peers_md}

CRITICAL ANTI-HALLUCINATION CONSTRAINTS:
1. You must use ONLY the supplied structured financial data for {company_name} ({ticker}).
2. If any metric or line item is missing or marked N/A, explicitly state that it is unavailable. Do NOT fabricate or invent any financial values.
3. The report must pertain strictly to {company_name} ({ticker}). Never reference any previous company.
4. Title of the report must strictly be: `# {company_name} ({ticker}) - Institutional Valuation`. Never invent prices, financial results, peer multiples, WACC, growth rates, news, or company facts.

Please generate the comprehensive Institutional Equity Research Memorandum strictly structured around the 4 Core Valuation Pillars now.
"""
    return prompt

def generate_report(provider, api_key, screener_data, valuation_result, model=None, custom_base_url=None):
    """
    Calls the specified AI provider and returns the generated markdown valuation report.
    """
    provider_key = provider.lower().strip()
    if not api_key:
        raise ValueError(f"API key is required for provider '{provider}'.")

    prompt_context = build_valuation_context(screener_data, valuation_result)

    # 1. Google Gemini Provider
    if provider_key == 'gemini':
        try:
            from google import genai
            from google.genai import types
            from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeout
            import time
            client = genai.Client(api_key=api_key)
            
            gemini_models = []
            if model:
                gemini_models.append(model)
            for m in PROVIDER_CONFIGS['gemini'].get('fallback_models', []):
                if m not in gemini_models:
                    gemini_models.append(m)

            gemini_config = types.GenerateContentConfig(
                temperature=0.3,
                top_p=0.85,
                max_output_tokens=4096,
                automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True)
            )

            deadline = time.time() + 45.0  # Max total Gemini budget: 45 seconds
            last_gemini_err = None

            for chosen_model in gemini_models:
                remaining_time = deadline - time.time()
                if remaining_time <= 5.0:
                    break
                model_timeout = min(30.0, remaining_time)
                try:
                    with ThreadPoolExecutor(max_workers=1) as executor:
                        future = executor.submit(
                            client.models.generate_content,
                            model=chosen_model,
                            contents=[SYSTEM_PROMPT, prompt_context],
                            config=gemini_config
                        )
                        try:
                            response = future.result(timeout=model_timeout)
                        except FutureTimeout:
                            raise RuntimeError(f"Gemini model '{chosen_model}' timed out after {model_timeout:.0f} seconds.")

                    return {
                        'report_markdown': response.text,
                        'provider': 'Google Gemini',
                        'model': chosen_model
                    }
                except Exception as e_m:
                    last_gemini_err = str(e_m)
                    if 'API_KEY_INVALID' in last_gemini_err or '401' in last_gemini_err or 'invalid api key' in last_gemini_err.lower():
                        raise ValueError("Authentication Error: Invalid API Key for Google Gemini.")
                    continue

            raise RuntimeError(f"Gemini synthesis reached timeout limit (45s). Tried models: {gemini_models}. Last notice: {last_gemini_err}")
        except ValueError:
            raise
        except Exception as e:
            raise RuntimeError(f"Gemini API Error: {str(e)}")

    # 2. OpenAI / OpenRouter / NVIDIA NIM / Custom OpenAI-Compatible
    cfg = PROVIDER_CONFIGS.get(provider_key, PROVIDER_CONFIGS['custom'])
    base_url = custom_base_url or cfg.get('base_url')
    
    if provider_key == 'custom' and not base_url:
        raise ValueError("Please provide a custom Base URL for the custom provider.")

    models_to_try = []
    if model:
        models_to_try.append(model)
    elif cfg.get('fallback_models'):
        models_to_try.extend(cfg['fallback_models'])
    else:
        models_to_try.append(cfg.get('default_model', 'meta/llama-3.3-70b-instruct'))

    client = OpenAI(
        base_url=base_url if base_url else None,
        api_key=api_key,
        timeout=85.0
    )

    last_error = None
    for m in models_to_try:
        try:
            completion = client.chat.completions.create(
                model=m,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": prompt_context}
                ],
                temperature=0.3,
                top_p=0.85,
                max_tokens=4096
            )
            content = completion.choices[0].message.content
            return {
                'report_markdown': content,
                'provider': cfg['name'],
                'model': m
            }
        except Exception as e:
            last_error = str(e)
            # If 401 Unauthorized, immediately return error
            if '401' in last_error or 'invalid api key' in last_error.lower() or 'unauthorized' in last_error.lower():
                raise ValueError(f"Authentication Error: Invalid API Key for {cfg['name']}.")
            continue

    raise RuntimeError(f"Model generation failed. Tried models: {models_to_try}. Last error: {last_error}")
