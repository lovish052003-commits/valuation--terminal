"""
app.py
Flask Web Application for the ITC Model Financial Valuation Platform.
Integrates live Screener.in data, mathematical valuation engine,
multi-provider AI research reports, and Excel export.
"""

import os
import sys
import threading
import time
import re
import json
import math
import dataclasses
import markdown
from flask import Flask, render_template, request, jsonify, send_file

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import screener_client
import valuation_engine
import llm_client
import excel_exporter
import company_logo_manager

app = Flask(__name__)

# In-memory session cache for latest analysis to enable instant parameter slider recalculations
ANALYSIS_CACHE = {}
EXPORT_EVENTS = {}
EXPORT_STATUS = {}
RUNNING_ANALYSES = set()
ANALYSIS_LOCK = threading.Lock()

def sanitize_for_json(obj):
    """
    Recursively converts and sanitizes Python data structures for strict RFC 8259 JSON compliance.
    Handles domain objects (CompanyData, CanonicalFinancialStatementLayer, etc.),
    dataclasses, pandas DataFrames/Series, numpy types, and replaces NaN/Inf with 0.0.
    Strips internal private Python pointers like '_company_data_obj'.
    """
    if obj is None or isinstance(obj, (str, bool)):
        return obj

    if isinstance(obj, dict):
        return {
            str(k): sanitize_for_json(v)
            for k, v in obj.items()
            if k != '_company_data_obj'
        }

    if isinstance(obj, (list, tuple, set, frozenset)):
        return [sanitize_for_json(v) for v in obj]

    if isinstance(obj, (float, int)):
        try:
            if math.isnan(obj) or math.isinf(obj):
                return 0.0
        except Exception:
            return 0.0
        return obj

    try:
        import numpy as np
        if isinstance(obj, (np.floating, np.integer)):
            if np.isnan(obj) or np.isinf(obj):
                return 0.0
            return float(obj)
        if isinstance(obj, np.ndarray):
            return [sanitize_for_json(v) for v in obj.tolist()]
    except Exception:
        pass

    try:
        import pandas as pd
        if isinstance(obj, pd.DataFrame):
            return sanitize_for_json(obj.fillna('-').to_dict(orient='records'))
        if isinstance(obj, pd.Series):
            return sanitize_for_json(obj.fillna('-').to_dict())
    except Exception:
        pass

    if hasattr(obj, 'to_dict') and callable(getattr(obj, 'to_dict')):
        try:
            return sanitize_for_json(obj.to_dict())
        except Exception:
            pass

    if hasattr(obj, 'to_summary_dict') and callable(getattr(obj, 'to_summary_dict')):
        try:
            return sanitize_for_json(obj.to_summary_dict())
        except Exception:
            pass

    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        try:
            return sanitize_for_json(dataclasses.asdict(obj))
        except Exception:
            pass

    if hasattr(obj, '__dict__'):
        try:
            return sanitize_for_json({
                k: v for k, v in obj.__dict__.items()
                if not str(k).startswith('_')
            })
        except Exception:
            pass

    try:
        json.dumps(obj)
        return obj
    except (TypeError, OverflowError):
        return str(obj)

@app.after_request
def add_cache_headers(response):
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/search', methods=['GET'])
def search():
    query = request.args.get('q', '').strip()
    if not query:
        return jsonify([])
    try:
        results = screener_client.search_companies(query)
        return jsonify(results)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/valuation', methods=['POST'])
@app.route('/api/analyze', methods=['POST'])
def analyze():
    data = request.get_json(force=True) or {}
    company_input = data.get('company', '').strip()
    if not company_input:
        return jsonify({'error': 'Please enter a company name or ticker symbol.'}), 400

    provider = data.get('provider', 'nvidia').lower().strip()
    api_key = data.get('apiKey', '').strip()
    model = data.get('model', '').strip() or None
    custom_base_url = data.get('customBaseUrl', '').strip() or None
    skip_ai = data.get('skipAi', False)
    print(f"[API] Valuation requested: '{company_input}' (provider: {provider}, skip_ai: {skip_ai})")

    custom_params = {}
    if data.get('wacc') is not None:
        try:
            custom_params['wacc'] = float(data['wacc']) / 100.0
        except (ValueError, TypeError):
            pass
    if data.get('growthRate') is not None:
        try:
            custom_params['growth_rate'] = float(data['growthRate']) / 100.0
        except (ValueError, TypeError):
            pass
    if data.get('terminalGrowth') is not None:
        try:
            custom_params['terminal_growth'] = float(data['terminalGrowth']) / 100.0
        except (ValueError, TypeError):
            pass
    if data.get('taxRate') is not None:
        try:
            custom_params['tax_rate'] = float(data['taxRate']) / 100.0
        except (ValueError, TypeError):
            pass

    clean_company_key = re.sub(r'[^A-Za-z0-9]', '', company_input).upper()
    with ANALYSIS_LOCK:
        if clean_company_key in RUNNING_ANALYSES:
            return jsonify({
                'error': f"Valuation analysis for '{company_input}' is already in progress. Please allow it a moment to complete."
            }), 429
        RUNNING_ANALYSES.add(clean_company_key)

    # Unique Valuation Run ID and state isolation
    run_id = f"{clean_company_key}_{int(time.time() * 1000)}"

    try:
        # 1. Fetch Screener data (with fast local cache & offline institutional archive fallback)
        try:
            screener_data = screener_client.fetch_company_data(company_input)
        except screener_client.AmbiguousCompanyError as e_amb:
            return jsonify({
                'success': False,
                'ambiguous': True,
                'error': str(e_amb),
                'message': str(e_amb),
                'candidates': e_amb.candidates
            }), 400
        except Exception as e_fetch:
            print(f"[API ERROR] /api/analyze failed for '{company_input}': {e_fetch}")
            return jsonify({
                'error': f"{str(e_fetch)}"
            }), 400

        ticker = screener_data.get('ticker', 'COMPANY').upper()
        
        # 2. Run Valuation Engine
        val_result = valuation_engine.calculate_valuation(screener_data, custom_params)
        val_result['run_id'] = run_id
        val_result['data_source'] = screener_data.get('data_source', 'Screener.in')
        val_result['data_as_of'] = screener_data.get('data_as_of')
        val_result['last_updated'] = screener_data.get('last_updated')

        # Cache isolated by canonical ticker
        cached_payload = {
            'screener_data': screener_data,
            'valuation_result': val_result,
            'run_id': run_id
        }
        ANALYSIS_CACHE[ticker] = cached_payload
        ANALYSIS_CACHE['latest'] = cached_payload

        # 3. AI Report Generation
        report_markdown = ""
        report_html = ""
        provider_name = ""
        model_name = ""

        if not skip_ai:
            if not api_key:
                return jsonify({
                    'error': f"Please enter your {provider.upper()} API key to generate the AI report, or check 'Skip AI' for quantitative DCF only."
                }), 400
            
            try:
                ai_res = llm_client.generate_report(
                    provider=provider,
                    api_key=api_key,
                    screener_data=screener_data,
                    valuation_result=val_result,
                    model=model,
                    custom_base_url=custom_base_url
                )
                report_markdown = ai_res['report_markdown']
                provider_name = ai_res['provider']
                model_name = ai_res['model']
                report_html = markdown.markdown(
                    report_markdown,
                    extensions=['tables', 'fenced_code', 'nl2br']
                )
            except Exception as e:
                # If AI fails, still return valuation math but with error note
                report_markdown = f"> [!WARNING]\n> AI Report Generation Warning: {str(e)}"
                report_html = f"<div class='alert alert-warning'><strong>AI Report Notice:</strong> {str(e)}</div>"

        # 4. Asynchronous Excel Model Generation
        excel_filename = f"{ticker}_Valuation_Model.xlsx"
        old_file_path = os.path.join(excel_exporter.EXPORT_DIR, excel_filename)
        if os.path.exists(old_file_path):
            try:
                os.remove(old_file_path)
            except Exception:
                pass

        EXPORT_EVENTS[ticker] = threading.Event()
        EXPORT_STATUS[ticker] = {'status': 'generating', 'run_id': run_id}
        
        def run_bg_excel():
            try:
                p = excel_exporter.export_valuation_model(screener_data, val_result, report_markdown)
                wb_val = excel_exporter.LAST_WORKBOOK_VALUATION.get(ticker) or excel_exporter.extract_workbook_valuation(p, val_result)
                EXPORT_STATUS[ticker] = {'status': 'done', 'path': p, 'run_id': run_id, 'workbook_valuation': wb_val}
            except Exception as e_bg:
                print(f"[Background Excel Export] Notice: {e_bg}")
                EXPORT_STATUS[ticker] = {'status': 'error', 'error': str(e_bg), 'run_id': run_id}
            finally:
                if ticker in EXPORT_EVENTS:
                    EXPORT_EVENTS[ticker].set()
                
        threading.Thread(target=run_bg_excel, daemon=True).start()

        # Convert raw pandas tables to serializable records for the UI safely
        def _safe_df_to_records(df):
            if df is None or not hasattr(df, 'empty') or df.empty:
                return {'columns': [], 'rows': []}
            cols = []
            seen = {}
            for c in df.columns:
                c_str = str(c)
                if c_str in seen:
                    seen[c_str] += 1
                    cols.append(f"{c_str}_{seen[c_str]}")
                else:
                    seen[c_str] = 0
                    cols.append(c_str)
            df_unique = df.copy(deep=False)
            df_unique.columns = cols
            return {
                'columns': cols,
                'rows': df_unique.fillna('-').to_dict(orient='records')
            }

        raw_tables = {}
        for k, df in screener_data.get('tables', {}).items():
            if df is not None and hasattr(df, 'empty') and not df.empty:
                raw_tables[k] = _safe_df_to_records(df)

        peers_df = screener_data.get('peers_df')
        peers_table = None
        if peers_df is not None and hasattr(peers_df, 'empty') and not peers_df.empty:
            peers_table = _safe_df_to_records(peers_df)

        canonical = screener_data.get('canonical_company', {})
        c_name = screener_data.get('company_name', ticker)
        return jsonify(sanitize_for_json({
            'success': True,
            'run_id': run_id,
            'runId': run_id,
            'data_source': screener_data.get('data_source', 'Screener.in'),
            'data_as_of': screener_data.get('data_as_of'),
            'dataAsOf': screener_data.get('data_as_of'),
            'last_updated': screener_data.get('last_updated'),
            'company': {
                'id': canonical.get('id', screener_data.get('company_id')),
                'name': c_name,
                'companyName': c_name,
                'shortName': canonical.get('shortName', c_name.split()[0]),
                'ticker': screener_data.get('ticker', ticker),
                'nseSymbol': screener_data.get('nse_ticker', screener_data.get('ticker', ticker)),
                'bseCode': screener_data.get('bse_code', ''),
                'company_type': screener_data.get('company_type', 'OTHER_NON_FINANCIAL'),
                'companyType': screener_data.get('company_type', 'OTHER_NON_FINANCIAL'),
                'nse_ticker': screener_data.get('nse_ticker', screener_data.get('ticker', ticker)),
                'bse_code': screener_data.get('bse_code', ''),
                'sector': screener_data.get('sector', ''),
                'industry': screener_data.get('industry', ''),
                'about': screener_data.get('about', '') or screener_data.get('about_company', ''),
                'url': screener_data.get('url', f"/company/{screener_data.get('ticker', ticker)}/consolidated/"),
                'current_price': screener_data.get('current_price', 0.0),
                'market_cap_cr': screener_data.get('market_cap_cr', 0.0),
                'shares_cr': screener_data.get('shares_in_cr', 0.0),
                'meta': screener_data.get('meta_raw', {})
            },
            'market': {
                'currentPrice': screener_data.get('current_price', 0.0),
                'marketCap': screener_data.get('market_cap_cr', 0.0),
                'sharesOutstanding': screener_data.get('shares_in_cr', 0.0),
                'peRatio': screener_data.get('pe_ratio', 0.0),
                'bookValue': screener_data.get('book_value', 0.0),
                'dividendYield': screener_data.get('dividend_yield', 0.0),
                'roce': screener_data.get('roce', 0.0),
                'roe': screener_data.get('roe', 0.0)
            },
            'financials': screener_data.get('normalized_financials', {}).get('historical', {}),
            'classification': {
                'sector': screener_data.get('sector', ''),
                'industry': screener_data.get('industry', ''),
                'company_type': screener_data.get('company_type', 'OTHER_NON_FINANCIAL'),
                'companyType': screener_data.get('company_type', 'OTHER_NON_FINANCIAL')
            },
            'peers': screener_data.get('sector_peers', []),
            'wacc': val_result.get('four_pillars', {}).get('pillar3_wacc', {}),
            'dcf': {
                'table': val_result.get('dcf_table', []),
                'intrinsicValue': val_result.get('intrinsic_value_per_share'),
                'enterpriseValue': val_result.get('enterprise_value'),
                'equityValue': val_result.get('equity_value'),
                'terminalValue': val_result.get('terminal_value')
            },
            'comparables': val_result.get('comps', {}),
            'dupont': val_result.get('dupont', {}),
            'altman': val_result.get('altman_z', {}),
            'aiReport': {
                'markdown': report_markdown,
                'html': report_html,
                'provider': provider_name,
                'model': model_name
            },
            'valuation': val_result,
            'report': {
                'markdown': report_markdown,
                'html': report_html,
                'provider': provider_name,
                'model': model_name
            },
            'excel_filename': excel_filename,
            'raw_tables': raw_tables,
            'peers_table': peers_table
        }))

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        with ANALYSIS_LOCK:
            RUNNING_ANALYSES.discard(clean_company_key)

@app.route('/api/recalculate', methods=['POST'])
def recalculate():
    """Recalculates DCF valuation on the fly when user tweaks parameter sliders."""
    data = request.get_json(force=True) or {}
    ticker = data.get('ticker', '').strip().upper()
    cached = ANALYSIS_CACHE.get(ticker) if ticker else ANALYSIS_CACHE.get('latest')
    if not cached:
        return jsonify({'error': 'No active company valuation in session. Please run an analysis first.'}), 400

    custom_params = {}
    if data.get('wacc') is not None:
        custom_params['wacc'] = float(data['wacc']) / 100.0
    if data.get('growthRate') is not None:
        custom_params['growth_rate'] = float(data['growthRate']) / 100.0
    if data.get('terminalGrowth') is not None:
        custom_params['terminal_growth'] = float(data['terminalGrowth']) / 100.0
    if data.get('taxRate') is not None:
        custom_params['tax_rate'] = float(data['taxRate']) / 100.0

    val_result = valuation_engine.calculate_valuation(cached['screener_data'], custom_params)
    val_result['run_id'] = cached.get('run_id')
    val_result['data_source'] = cached['screener_data'].get('data_source', 'Screener.in')
    val_result['data_as_of'] = cached['screener_data'].get('data_as_of')
    val_result['last_updated'] = cached['screener_data'].get('last_updated')

    cached['valuation_result'] = val_result

    t = cached['screener_data'].get('ticker', 'COMPANY').upper()
    excel_filename = f"{t}_Valuation_Model.xlsx"

    return jsonify(sanitize_for_json({
        'success': True,
        'valuation': val_result,
        'excel_filename': excel_filename
    }))

@app.route('/api/export-status/<ticker>')
def export_status(ticker):
    t = ticker.upper()
    status_info = dict(EXPORT_STATUS.get(t, {'status': 'not_started'}))
    # If not actively generating, check if a completed model exists on disk
    if status_info.get('status') != 'generating':
        primary_path = os.path.join(excel_exporter.EXPORT_DIR, f"{t}_Valuation_Model.xlsx")
        if os.path.exists(primary_path):
            wb_val = status_info.get('workbook_valuation') or excel_exporter.LAST_WORKBOOK_VALUATION.get(t) or excel_exporter.extract_workbook_valuation(primary_path)
            status_info = {
                'status': 'done',
                'path': primary_path,
                'filename': f"{t}_Valuation_Model.xlsx",
                'workbook_valuation': wb_val
            }
    elif status_info.get('status') == 'done' and not status_info.get('workbook_valuation'):
        wb_val = excel_exporter.LAST_WORKBOOK_VALUATION.get(t) or excel_exporter.extract_workbook_valuation(status_info.get('path'))
        if wb_val:
            status_info['workbook_valuation'] = wb_val
    return jsonify(sanitize_for_json(status_info))

@app.route('/api/download-excel/<filename>')
def download_excel(filename):
    # Sanitize filename
    clean_filename = os.path.basename(filename)
    ticker_prefix = clean_filename.replace('.xlsx', '').split('_')[0].upper()
    
    # If a background export was initiated or is generating, wait for it to complete
    status_info = EXPORT_STATUS.get(ticker_prefix, {})
    if status_info.get('status') == 'generating' or ticker_prefix in EXPORT_EVENTS:
        if ticker_prefix in EXPORT_EVENTS:
            EXPORT_EVENTS[ticker_prefix].wait(timeout=120)
        status_info = EXPORT_STATUS.get(ticker_prefix, {})

    if status_info.get('status') == 'done' and status_info.get('path'):
        fresh_path = status_info['path']
        if os.path.exists(fresh_path) and os.path.getsize(fresh_path) > 10000:
            return send_file(fresh_path, as_attachment=True, download_name=clean_filename)

    # Check if we have cached analysis for this ticker to generate on demand if file is missing, stale, or failed
    cached = ANALYSIS_CACHE.get(ticker_prefix) or (ANALYSIS_CACHE.get('latest') if ANALYSIS_CACHE.get('latest', {}).get('screener_data', {}).get('ticker') == ticker_prefix else None)
    file_path = os.path.join(excel_exporter.EXPORT_DIR, clean_filename)

    # If cached analysis is available and the file is missing or status was error, synchronously generate fresh!
    if cached and cached.get('screener_data') and (status_info.get('status') in ('error', 'not_started', None) or not os.path.exists(file_path)):
        try:
            print(f"[Download Excel] Generating fresh model for {ticker_prefix} on-demand...")
            fresh_p = excel_exporter.export_valuation_model(cached['screener_data'], cached.get('val_result', {}), cached.get('report_markdown', ''))
            if fresh_p and os.path.exists(fresh_p):
                return send_file(fresh_p, as_attachment=True, download_name=clean_filename)
        except Exception as e_sync:
            print(f"[Download Excel] On-demand generation error: {e_sync}")

    # 1. Immediate disk check: If the completed, verified file is already on disk and not generating, serve it!
    if os.path.exists(file_path) and not clean_filename.startswith('.') and os.path.getsize(file_path) > 10000:
        if status_info.get('status') != 'generating':
            return send_file(file_path, as_attachment=True, download_name=clean_filename)
        if 'text/html' in request.headers.get('Accept', ''):
            return f"""<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="refresh" content="3">
    <title>Generating Valuation Model...</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #0f172a; color: #f8fafc; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; margin: 0; }}
        .spinner {{ border: 4px solid rgba(255,255,255,0.1); width: 44px; height: 44px; border-radius: 50%; border-left-color: #38bdf8; animation: spin 1s linear infinite; margin-bottom: 20px; }}
        @keyframes spin {{ 0% {{ transform: rotate(0deg); }} 100% {{ transform: rotate(360deg); }} }}
        h2 {{ font-weight: 600; margin-bottom: 8px; }}
        p {{ color: #94a3b8; font-size: 14px; text-align: center; max-width: 480px; line-height: 1.5; }}
    </style>
</head>
<body>
    <div class="spinner"></div>
    <h2>Generating {clean_filename}...</h2>
    <p>Calculating all 22 sheets and validating peer comparables. This page will automatically refresh and download the workbook once ready.</p>
</body>
</html>""", 202
        return "Excel model is still calculating all 22 sheets natively. Please wait a few seconds and try again.", 202

    # Fallback to checking EXPORT_DIR for completed, verified files
    for _ in range(10):
        if os.path.exists(excel_exporter.EXPORT_DIR):
            candidates = [
                f for f in os.listdir(excel_exporter.EXPORT_DIR)
                if not f.startswith('.')
                and not f.endswith('.tmp')
                and not f.endswith('.logotmp')
                and f.endswith('.xlsx')
                and (f == clean_filename or f.startswith(f"{ticker_prefix}_Valuation_Model") or f == f"TEST_{ticker_prefix}_COM.xlsx")
            ]
            if candidates:
                candidates.sort(key=lambda x: os.path.getmtime(os.path.join(excel_exporter.EXPORT_DIR, x)), reverse=True)
                latest_path = os.path.join(excel_exporter.EXPORT_DIR, candidates[0])
                if os.path.exists(latest_path) and os.path.getsize(latest_path) > 10000:
                    return send_file(latest_path, as_attachment=True, download_name=clean_filename)
        time.sleep(0.5)

    file_path = os.path.join(excel_exporter.EXPORT_DIR, clean_filename)
    if os.path.exists(file_path) and not clean_filename.startswith('.') and os.path.getsize(file_path) > 10000:
        return send_file(file_path, as_attachment=True, download_name=clean_filename)

    if 'text/html' in request.headers.get('Accept', ''):
        return f"""<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="refresh" content="3">
    <title>Preparing Valuation Model...</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #0f172a; color: #f8fafc; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; margin: 0; }}
        .spinner {{ border: 4px solid rgba(255,255,255,0.1); width: 44px; height: 44px; border-radius: 50%; border-left-color: #38bdf8; animation: spin 1s linear infinite; margin-bottom: 20px; }}
        @keyframes spin {{ 0% {{ transform: rotate(0deg); }} 100% {{ transform: rotate(360deg); }} }}
        h2 {{ font-weight: 600; margin-bottom: 8px; }}
        p {{ color: #94a3b8; font-size: 14px; text-align: center; max-width: 480px; line-height: 1.5; }}
    </style>
</head>
<body>
    <div class="spinner"></div>
    <h2>Preparing {clean_filename}...</h2>
    <p>Excel model is being finalized. This page will automatically refresh and download the file.</p>
</body>
</html>""", 202

    return "Excel model is still generating. Please try again in a few seconds.", 404

if __name__ == '__main__':
    # Ensure export directory exists
    excel_exporter.ensure_export_dir()
    print("Starting Institutional Valuation Platform on http://127.0.0.1:5000 ...")
    app.run(debug=False, port=5000, threaded=True)

