"""Adds the Alt Valuation button + panel to templates/index.html and static/js/app.js.
Usage: python patch_alt_frontend.py <path/index.html> <path/app.js>   (makes .bak_alt backups; aborts on mismatch)"""
import sys, shutil
html_p, js_p = sys.argv[1], sys.argv[2]
h = open(html_p, encoding='utf-8').read()
j = open(js_p, encoding='utf-8').read()
if 'btnDownloadAlt' in h or 'applyAltModel' in j:
    print('already patched'); sys.exit(0)

def rep(s, old, new):
    assert s.count(old) == 1, f'expected 1 match, found {s.count(old)}: {old[:70]!r}'
    return s.replace(old, new)

# ---------- HTML ----------
h = rep(h, """                </div>
            </div>

            <!-- Hero Valuation KPI Grid -->""",
"""                    <a href="#" id="btnDownloadAlt" class="btn-action btn-excel" style="display:none;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM17 13l-5 5-5-5h3V9h4v4h3z"/></svg>
                        <span>Download Alternate Valuation Model (.xlsx)</span>
                    </a>
                </div>
            </div>

            <!-- Alternate Valuation Panel (shown only when the FCFF DCF is not applicable) -->
            <div id="altModelPanel" class="glass-panel" style="display:none; margin-bottom:1.5rem; padding:1.25rem 1.5rem; border:1px solid rgba(245,158,11,0.35); background:rgba(245,158,11,0.06);">
                <div style="display:flex; justify-content:space-between; gap:1rem; flex-wrap:wrap;">
                    <div>
                        <div class="kpi-label">ALTERNATE VALUATION &bull; DCF NOT APPLICABLE</div>
                        <div id="altReason" class="text-muted" style="font-size:0.85rem; margin:0.25rem 0 0.75rem;"></div>
                        <div>Method: <strong id="altMethod">&mdash;</strong> &bull; <span id="altType" class="text-muted"></span></div>
                    </div>
                    <div style="display:flex; gap:2rem; flex-wrap:wrap; align-items:flex-start;">
                        <div><div class="pm-label">Alt. Intrinsic Value</div><div class="pm-val text-cyan" id="altValue">&mdash;</div></div>
                        <div><div class="pm-label">Upside / (Downside)</div><div class="pm-val" id="altUpside">&mdash;</div></div>
                        <div><div class="pm-label">Verdict</div><div class="pm-val" id="altVerdict">&mdash;</div></div>
                    </div>
                </div>
                <div style="display:flex; gap:1.5rem; flex-wrap:wrap; margin-top:0.75rem; font-size:0.82rem;">
                    <span>P/B: <strong id="altPb" class="font-mono">&mdash;</strong></span>
                    <span>ROE: <strong id="altRoe" class="font-mono">&mdash;</strong></span>
                    <span>Market-implied ROE: <strong id="altImpRoe" class="font-mono">&mdash;</strong></span>
                    <span>Cost of equity: <strong id="altKe" class="font-mono">&mdash;</strong></span>
                    <span>Peer P/E value: <strong id="altPeVal" class="font-mono">&mdash;</strong></span>
                    <span>EV/EBITDA value: <strong id="altEvVal" class="font-mono">&mdash;</strong></span>
                </div>
                <div id="altConfidence" class="text-muted" style="font-size:0.75rem; margin-top:0.6rem;"></div>
                <div class="text-muted" style="font-size:0.75rem;">All figures above are read from the downloadable Alternate Valuation workbook.</div>
            </div>

            <!-- Hero Valuation KPI Grid -->""")

# ---------- JS ----------
j = rep(j, """        // Excel Download Button
        const tkr = (data.company && data.company.ticker) ? data.company.ticker : 'COMPANY';
        monitorExcelExport(tkr, data.excel_filename);
    }
""", """        // Excel Download Button
        const tkr = (data.company && data.company.ticker) ? data.company.ticker : 'COMPANY';
        applyAltModel(data.alt_model);
        monitorExcelExport(tkr, data.excel_filename);
    }

    // --- Alternate valuation (banks / NBFCs / conglomerates where FCFF DCF is not meaningful) ---
    window.__altActive = false;
    function applyAltModel(alt) {
        const btn = document.getElementById('btnDownloadAlt');
        const panel = document.getElementById('altModelPanel');
        const params = document.querySelector('.params-toolbar');
        if (!alt || !alt.value) {
            window.__altActive = false;
            if (btn) btn.style.display = 'none';
            if (panel) panel.style.display = 'none';
            if (params) params.style.display = '';
            return;
        }
        window.__altActive = true;
        if (btn) { btn.href = alt.download_url; btn.style.display = ''; }
        if (panel) panel.style.display = 'block';
        if (params) params.style.display = 'none';
        const fx = (x, d, suf) => (typeof x === 'number' && isFinite(x)) ? x.toFixed(d) + (suf || '') : 'n/a';
        const pct = x => (typeof x === 'number' && isFinite(x)) ? (x * 100).toFixed(1) + '%' : 'n/a';
        const val = Number(alt.value), up = Number(alt.upside_pct);
        safeSetText('altReason', alt.reason || '');
        safeSetText('altMethod', alt.method || '');
        safeSetText('altType', alt.company_type || '');
        safeSetText('altValue', `₹${val.toFixed(2)}`);
        safeSetText('altUpside', `${up > 0 ? '+' : ''}${up.toFixed(1)}%`);
        safeSetText('altVerdict', alt.verdict || '');
        safeSetText('altPb', fx(alt.pb, 2, 'x'));
        safeSetText('altRoe', pct(alt.roe));
        safeSetText('altImpRoe', pct(alt.implied_roe));
        safeSetText('altKe', pct(alt.cost_of_equity));
        safeSetText('altPeVal', typeof alt.pe_value === 'number' ? `₹${alt.pe_value.toFixed(0)}` : 'n/a');
        safeSetText('altEvVal', typeof alt.ev_ebitda_value === 'number' ? `₹${alt.ev_ebitda_value.toFixed(0)}` : 'n/a');
        safeSetText('altConfidence', alt.confidence || '');
        // headline KPIs follow the alternate workbook
        safeSetText('kpiIntrinsic', `₹${val.toFixed(2)}`);
        const upsideEl = document.getElementById('kpiUpside');
        if (upsideEl) {
            upsideEl.textContent = `${up > 0 ? '+' : ''}${up.toFixed(1)}%`;
            upsideEl.className = `kpi-val ${up >= 0 ? 'val-positive' : 'val-negative'}`;
        }
        const verdictEl = document.getElementById('verdictBadge');
        if (verdictEl && alt.verdict) {
            verdictEl.textContent = alt.verdict;
            const vc = /UNDER/.test(alt.verdict) ? 'buy' : (/OVER/.test(alt.verdict) ? 'sell' : 'hold');
            verdictEl.className = `verdict-badge verdict-${vc}`;
        }
    }
""")
j = rep(j, "    function syncWorkbookValuationToTerminal(wbVal) {\n        if (!wbVal || typeof wbVal !== 'object') return;\n",
           "    function syncWorkbookValuationToTerminal(wbVal) {\n        if (window.__altActive) return;\n        if (!wbVal || typeof wbVal !== 'object') return;\n")
shutil.copy(html_p, html_p + '.bak_alt'); shutil.copy(js_p, js_p + '.bak_alt')
open(html_p, 'w', encoding='utf-8').write(h); open(js_p, 'w', encoding='utf-8').write(j)
print('OK')
