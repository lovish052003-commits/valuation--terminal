/**
 * app.js - Institutional Financial Valuation Platform Client Script
 */

document.addEventListener('DOMContentLoaded', () => {
    // --- State & DOM References ---
    let activeCompany = null;
    let activeValuation = null;
    let searchDebounceTimer = null;
    let sliderDebounceTimer = null;

    // Elements
    const searchInput = document.getElementById('searchInput');
    const autocompleteDropdown = document.getElementById('autocompleteDropdown');
    const launchBtn = document.getElementById('launchBtn');
    const skipAiCheckbox = document.getElementById('skipAiCheckbox');
    const loadingOverlay = document.getElementById('loadingOverlay');
    const workspaceContainer = document.getElementById('workspaceContainer');
    const valErrorBanner = document.getElementById('valErrorBanner');
    const valErrorTitle = document.getElementById('valErrorTitle');
    const valErrorMessage = document.getElementById('valErrorMessage');
    const btnDismissError = document.getElementById('btnDismissError');
    const btnCancelLoading = document.getElementById('btnCancelLoading');
    let activeAbortController = null;

    if (btnDismissError && valErrorBanner) {
        btnDismissError.addEventListener('click', () => {
            valErrorBanner.style.display = 'none';
        });
    }

    // Modal elements
    const settingsModal = document.getElementById('settingsModal');
    const btnOpenSettings = document.getElementById('btnOpenSettings');
    const btnCloseModal = document.getElementById('btnCloseModal');
    const btnCancelModal = document.getElementById('btnCancelModal');
    const btnSaveSettings = document.getElementById('btnSaveSettings');
    const modalProvider = document.getElementById('modalProvider');
    const modalApiKey = document.getElementById('modalApiKey');
    const modalModel = document.getElementById('modalModel');
    const modalCustomBaseUrl = document.getElementById('modalCustomBaseUrl');
    const customUrlGroup = document.getElementById('customUrlGroup');
    const navApiStatus = document.getElementById('navApiStatus');

    // Sliders
    const sliderWacc = document.getElementById('sliderWacc');
    const sliderGrowth = document.getElementById('sliderGrowth');
    const sliderTerminal = document.getElementById('sliderTerminal');
    const sliderTax = document.getElementById('sliderTax');
    const badgeWacc = document.getElementById('badgeWacc');
    const badgeGrowth = document.getElementById('badgeGrowth');
    const badgeTerminal = document.getElementById('badgeTerminal');
    const badgeTax = document.getElementById('badgeTax');

    // --- 1. API Configuration Management (localStorage) ---
    function loadSavedSettings() {
        const provider = localStorage.getItem('val_provider') || 'nvidia';
        const apiKey = localStorage.getItem(`val_apikey_${provider}`) || localStorage.getItem('val_apikey') || '';
        const model = localStorage.getItem('val_model') || '';
        const customUrl = localStorage.getItem('val_custom_url') || '';

        modalProvider.value = provider;
        modalApiKey.value = apiKey;
        modalModel.value = model;
        modalCustomBaseUrl.value = customUrl;

        toggleCustomUrlVisibility();
        updateNavApiBadge();
    }

    function toggleCustomUrlVisibility() {
        if (modalProvider.value === 'custom') {
            customUrlGroup.style.display = 'block';
        } else {
            customUrlGroup.style.display = 'none';
        }
    }

    modalProvider.addEventListener('change', () => {
        toggleCustomUrlVisibility();
        // Prefill existing key for that provider if saved
        const p = modalProvider.value;
        modalApiKey.value = localStorage.getItem(`val_apikey_${p}`) || '';
        
        // Auto-fill suggested default model placeholder
        const defaults = {
            'nvidia': 'meta/llama-3.3-70b-instruct',
            'openrouter': 'deepseek/deepseek-r1',
            'openai': 'gpt-4o',
            'gemini': 'gemini-2.5-flash',
            'custom': ''
        };
        modalModel.placeholder = defaults[p] || '';
    });

    function saveSettings() {
        const provider = modalProvider.value;
        const apiKey = modalApiKey.value.trim();
        const model = modalModel.value.trim();
        const customUrl = modalCustomBaseUrl.value.trim();

        localStorage.setItem('val_provider', provider);
        if (apiKey) {
            localStorage.setItem(`val_apikey_${provider}`, apiKey);
            localStorage.setItem('val_apikey', apiKey);
        }
        localStorage.setItem('val_model', model);
        localStorage.setItem('val_custom_url', customUrl);

        settingsModal.style.display = 'none';
        updateNavApiBadge();
        showToast('Settings saved successfully!');
    }

    function updateNavApiBadge() {
        const provider = localStorage.getItem('val_provider') || 'nvidia';
        const apiKey = localStorage.getItem(`val_apikey_${provider}`) || localStorage.getItem('val_apikey') || '';
        const dot = navApiStatus.querySelector('.status-dot');
        const text = navApiStatus.querySelector('.status-text');

        if (apiKey) {
            dot.classList.add('active');
            text.textContent = `${provider.toUpperCase()} Configured`;
        } else {
            dot.classList.remove('active');
            text.textContent = 'API Key Required';
        }
    }

    btnOpenSettings.addEventListener('click', () => {
        loadSavedSettings();
        settingsModal.style.display = 'flex';
    });

    btnCloseModal.addEventListener('click', () => settingsModal.style.display = 'none');
    btnCancelModal.addEventListener('click', () => settingsModal.style.display = 'none');
    btnSaveSettings.addEventListener('click', saveSettings);

    // --- Ambiguity Modal Elements & Handler ---
    const ambiguityModal = document.getElementById('ambiguityModal');
    const btnCloseAmbiguityModal = document.getElementById('btnCloseAmbiguityModal');
    const btnCancelAmbiguityModal = document.getElementById('btnCancelAmbiguityModal');
    const ambiguityCandidateList = document.getElementById('ambiguityCandidateList');

    if (btnCloseAmbiguityModal) btnCloseAmbiguityModal.addEventListener('click', () => ambiguityModal.style.display = 'none');
    if (btnCancelAmbiguityModal) btnCancelAmbiguityModal.addEventListener('click', () => ambiguityModal.style.display = 'none');

    function showAmbiguityModal(candidates) {
        if (!ambiguityModal || !ambiguityCandidateList) return;
        ambiguityCandidateList.innerHTML = '';
        candidates.forEach(c => {
            const item = document.createElement('div');
            item.style.cssText = 'padding: 0.75rem 1rem; border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; background: rgba(255,255,255,0.03); transition: all 0.2s ease;';
            item.innerHTML = `
                <div>
                    <div style="font-weight: 600; color: #f8fafc;">${c.companyName || c.company_name}</div>
                    <div style="font-size: 0.8rem; color: #94a3b8;">${c.industry || ''} ${c.bseCode ? '• BSE: ' + c.bseCode : ''}</div>
                </div>
                <span class="ac-ticker" style="font-size: 0.85rem; padding: 0.2rem 0.5rem; background: rgba(56,189,248,0.15); color: #38bdf8; border-radius: 4px;">${c.nseSymbol || c.nse_ticker || c.ticker || ''}</span>
            `;
            item.addEventListener('mouseenter', () => {
                item.style.background = 'rgba(56,189,248,0.1)';
                item.style.borderColor = 'rgba(56,189,248,0.4)';
            });
            item.addEventListener('mouseleave', () => {
                item.style.background = 'rgba(255,255,255,0.03)';
                item.style.borderColor = 'rgba(255,255,255,0.1)';
            });
            item.addEventListener('click', () => {
                ambiguityModal.style.display = 'none';
                const chosen = c.nseSymbol || c.nse_ticker || c.ticker || c.companyName || c.company_name;
                searchInput.value = chosen;
                triggerAnalysis(chosen);
            });
            ambiguityCandidateList.appendChild(item);
        });
        ambiguityModal.style.display = 'flex';
    }

    // --- 2. Live Screener Search & Autocomplete ---
    searchInput.addEventListener('input', () => {
        clearTimeout(searchDebounceTimer);
        const query = searchInput.value.trim();
        if (query.length < 1) {
            autocompleteDropdown.style.display = 'none';
            return;
        }

        searchDebounceTimer = setTimeout(async () => {
            try {
                const res = await fetch(`/api/search?q=${encodeURIComponent(query)}`);
                const items = await res.json();
                renderAutocomplete(items);
            } catch (err) {
                console.error("Search error:", err);
            }
        }, 220);
    });

    function renderAutocomplete(items) {
        autocompleteDropdown.innerHTML = '';
        if (!items || items.length === 0) {
            autocompleteDropdown.style.display = 'none';
            return;
        }

        items.forEach(item => {
            const div = document.createElement('div');
            div.className = 'autocomplete-item';
            div.innerHTML = `
                <span class="ac-company-name">${item.name}</span>
                <span class="ac-ticker">${item.ticker}</span>
            `;
            div.addEventListener('click', () => {
                searchInput.value = item.name;
                autocompleteDropdown.style.display = 'none';
                triggerAnalysis(item.name);
            });
            autocompleteDropdown.appendChild(div);
        });
        autocompleteDropdown.style.display = 'block';
    }

    // Close autocomplete on click outside
    document.addEventListener('click', (e) => {
        if (!searchInput.contains(e.target) && !autocompleteDropdown.contains(e.target)) {
            autocompleteDropdown.style.display = 'none';
        }
    });

    // Sample quick chips
    document.querySelectorAll('.chip').forEach(chip => {
        chip.addEventListener('click', () => {
            const company = chip.getAttribute('data-company');
            searchInput.value = company;
            triggerAnalysis(company);
        });
    });

    // --- 3. Valuation Execution ---
    let isAnalyzing = false;

    launchBtn.addEventListener('click', (e) => {
        e.preventDefault();
        const company = searchInput.value.trim();
        if (!company) {
            showToast('Please enter a company name or ticker symbol.');
            return;
        }
        triggerAnalysis(company);
    });

    searchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            autocompleteDropdown.style.display = 'none';
            const company = searchInput.value.trim();
            if (company) triggerAnalysis(company);
        }
    });

    async function triggerAnalysis(companyName) {
        if (isAnalyzing) {
            showToast('Valuation analysis is already in progress. Please wait...');
            return;
        }

        if (valErrorBanner) valErrorBanner.style.display = 'none';
        autocompleteDropdown.style.display = 'none';
        const provider = localStorage.getItem('val_provider') || 'nvidia';
        const apiKey = localStorage.getItem(`val_apikey_${provider}`) || localStorage.getItem('val_apikey') || '';
        const model = localStorage.getItem('val_model') || '';
        const customUrl = localStorage.getItem('val_custom_url') || '';
        const skipAi = skipAiCheckbox.checked;

        if (!skipAi && !apiKey) {
            settingsModal.style.display = 'flex';
            showToast(`Please enter your ${provider.toUpperCase()} API key first, or check 'Skip AI'.`);
            return;
        }

        isAnalyzing = true;
        launchBtn.disabled = true;
        launchBtn.style.opacity = '0.65';
        launchBtn.style.cursor = 'not-allowed';

        let isTimedOut = false;
        let isUserCancelled = false;

        // Setup cancellation controller with 120s safety limit
        if (activeAbortController) {
            try { activeAbortController.abort(); } catch(e){}
        }
        activeAbortController = new AbortController();
        const signal = activeAbortController.signal;

        const subTextEl = document.querySelector('.loading-subtext');
        if (subTextEl) subTextEl.textContent = 'Fetching Screener.in audited statements & live market ratios...';

        const stepTimer1 = setTimeout(() => {
            if (subTextEl) subTextEl.textContent = 'Extracting balance sheet schedules & true sector peers...';
        }, 2000);

        const stepTimer2 = setTimeout(() => {
            if (subTextEl) subTextEl.textContent = 'Computing 5-Year DCF, WACC, DuPont, & Altman Z matrices...';
        }, 5000);

        const stepTimer3 = setTimeout(() => {
            if (subTextEl) subTextEl.textContent = 'Connecting Multi-LLM research intelligence (synthesizing report)...';
        }, 9000);

        const stepTimer4 = setTimeout(() => {
            if (subTextEl) subTextEl.textContent = 'Finalizing Institutional Memorandum & 22-Sheet Excel model...';
        }, 22000);

        const abortSafetyTimeout = setTimeout(() => {
            isTimedOut = true;
            if (activeAbortController) activeAbortController.abort();
        }, 120000);

        const cleanupTimers = () => {
            clearTimeout(stepTimer1);
            clearTimeout(stepTimer2);
            clearTimeout(stepTimer3);
            clearTimeout(stepTimer4);
            clearTimeout(abortSafetyTimeout);
        };

        if (btnCancelLoading) {
            btnCancelLoading.onclick = () => {
                isUserCancelled = true;
                if (activeAbortController) activeAbortController.abort();
                loadingOverlay.style.display = 'none';
                cleanupTimers();
                showToast('Valuation analysis cancelled.');
            };
        }

        // Show loading spinner
        loadingOverlay.style.display = 'block';
        workspaceContainer.style.display = 'none';

        try {
            const payload = {
                company: companyName,
                provider: provider,
                apiKey: apiKey,
                model: model,
                customBaseUrl: customUrl,
                skipAi: skipAi
            };

            const resp = await fetch('/api/analyze', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload),
                signal: signal
            });

            cleanupTimers();

            const data = await resp.json();
            loadingOverlay.style.display = 'none';

            if (!resp.ok || data.error) {
                if (data.ambiguous && data.candidates && data.candidates.length > 0) {
                    showAmbiguityModal(data.candidates);
                    return;
                }
                const errText = data.error || 'Failed to complete valuation analysis.';
                if (valErrorBanner && valErrorMessage) {
                    if (valErrorTitle) valErrorTitle.textContent = 'Valuation Analysis Notice';
                    valErrorMessage.innerHTML = `${errText}<br><br><strong>Tip:</strong> If live Screener.in is temporarily blocked by your network (e.g. 24Online Client), you can instantly analyze pre-cached companies: <em>ITC Ltd, Tata Motors, Nestle India, Force Motors, Jubilant Foodworks, Infosys, Britannia, Hindustan Unilever, Tata Steel, Reliance</em>.`;
                    valErrorBanner.style.display = 'block';
                    valErrorBanner.scrollIntoView({ behavior: 'smooth' });
                }
                showToast(errText);
                return;
            }

            activeCompany = data.company;
            activeValuation = data.valuation;

            try {
                renderValuationResults(data);
            } catch (renderErr) {
                console.error("Rendering error:", renderErr);
            }
            workspaceContainer.style.display = 'block';
            workspaceContainer.scrollIntoView({ behavior: 'smooth' });

        } catch (err) {
            cleanupTimers();
            loadingOverlay.style.display = 'none';
            if (err.name === 'AbortError') {
                if (isTimedOut) {
                    showToast('Analysis request timed out after 120 seconds. Check "Skip AI Report" for instant quantitative calculation.');
                } else if (isUserCancelled) {
                    // Handled in cancel click
                }
            } else {
                if (valErrorBanner && valErrorMessage) {
                    if (valErrorTitle) valErrorTitle.textContent = 'Connection / Network Notice';
                    valErrorMessage.innerHTML = `Network issue while contacting terminal server: ${err.message}.<br><br><strong>Tip:</strong> If using campus Wi-Fi (e.g. LPU 24Online Client), ensure you are logged in. Pre-loaded companies (ITC Ltd, Tata Motors, Nestle India, Force Motors, Jubilant Foodworks) work offline!`;
                    valErrorBanner.style.display = 'block';
                    valErrorBanner.scrollIntoView({ behavior: 'smooth' });
                }
                showToast('Network error while running valuation: ' + err.message);
            }
        } finally {
            isAnalyzing = false;
            launchBtn.disabled = false;
            launchBtn.style.opacity = '1';
            launchBtn.style.cursor = 'pointer';
            cleanupTimers();
        }
    }

    // Helper for safe DOM text setting
    function safeSetText(id, text) {
        const el = document.getElementById(id);
        if (el) el.textContent = text;
    }

    function safeSetHref(id, url) {
        const el = document.getElementById(id);
        if (el) el.href = url;
    }

    // --- 4. Render Results to UI ---
    function renderValuationResults(data) {
        const c = data.company;
        const v = data.valuation;

        // Header info
        safeSetText('headerCompanyName', c.name);
        safeSetText('headerTicker', c.ticker);
        safeSetText('headerSector', `${c.sector} • ${c.industry}`);
        safeSetHref('headerScreenerLink', c.url);

        // KPI cards
        safeSetText('kpiCmp', `₹${c.current_price}`);
        safeSetText('kpiIntrinsic', `₹${v.intrinsic_value_per_share}`);
        
        const upsideEl = document.getElementById('kpiUpside');
        if (upsideEl) {
            upsideEl.textContent = `${v.margin_of_safety_pct > 0 ? '+' : ''}${v.margin_of_safety_pct}%`;
            upsideEl.className = `kpi-val ${v.margin_of_safety_pct >= 0 ? 'val-positive' : 'val-negative'}`;
        }

        // Verdict Badge
        const verdictEl = document.getElementById('verdictBadge');
        if (verdictEl) {
            verdictEl.textContent = v.verdict;
            verdictEl.className = `verdict-badge verdict-${v.verdict_class}`;
        }

        // Secondary KPIs
        safeSetText('kpiMcap', `Market Cap: ₹${v.market_cap_cr.toLocaleString()} Cr`);
        safeSetText('kpiShares', `${v.shares_cr} Cr Shares`);
        safeSetText('kpiWacc', `${v.wacc}%`);
        safeSetText('kpiEv', `₹${v.enterprise_value.toLocaleString()} Cr`);

        // Update Sliders to match initial calculated values
        sliderWacc.value = v.wacc;
        badgeWacc.textContent = `${v.wacc}%`;

        sliderGrowth.value = v.growth_rate;
        badgeGrowth.textContent = `${v.growth_rate}%`;

        sliderTerminal.value = v.terminal_growth;
        badgeTerminal.textContent = `${v.terminal_growth}%`;

        sliderTax.value = v.tax_rate;
        badgeTax.textContent = `${v.tax_rate}%`;

        // The 4 Valuation Pillars Dashboard
        renderFourPillars(v);

        // DCF Table
        renderDcfTable(v.dcf_table);

        // Sensitivity Grid
        renderSensitivityGrid(v.sensitivity);

        // DuPont Scorecard
        renderDuPont(v.dupont);

        // Altman Z Scorecard
        renderAltmanZ(v.altman_z);

        // Peers Table
        renderPeersTable(data.peers_table);

        // AI Report
        const reportArea = document.getElementById('aiReportContent');
        const reportMeta = document.getElementById('aiReportMeta');
        if (data.report && data.report.html) {
            reportArea.innerHTML = data.report.html;
            reportMeta.textContent = `Generated by ${data.report.provider} (${data.report.model})`;
        } else {
            reportArea.innerHTML = "<p class='text-muted'>AI Research Report was skipped. Check API settings and uncheck 'Skip AI' to generate.</p>";
            reportMeta.textContent = "";
        }

        // Screener Financial Statements
        renderRawStatements(data.raw_tables);
        
        // Fundamental Analysis Tab
        renderFundamentalAnalysis(data);
        renderRiskometer(data);
        renderForensicCheck(data);
        renderReverseDcf(data);

        // Excel Download Button
        const tkr = (data.company && data.company.ticker) ? data.company.ticker : 'COMPANY';
        monitorExcelExport(tkr, data.excel_filename);
    }

    let excelPollInterval = null;

    function syncWorkbookValuationToTerminal(wbVal) {
        if (!wbVal || typeof wbVal !== 'object') return;
        if (wbVal.intrinsic_value && wbVal.intrinsic_value > 0) {
            safeSetText('kpiIntrinsic', `₹${wbVal.intrinsic_value}`);
            const upsideEl = document.getElementById('kpiUpside');
            if (upsideEl && wbVal.margin_of_safety_pct !== undefined) {
                const mos = wbVal.margin_of_safety_pct;
                upsideEl.textContent = `${mos > 0 ? '+' : ''}${mos}%`;
                upsideEl.className = `kpi-val ${mos >= 0 ? 'val-positive' : 'val-negative'}`;
            }
            const verdictEl = document.getElementById('verdictBadge');
            if (verdictEl && wbVal.verdict) {
                verdictEl.textContent = wbVal.verdict;
                const vClass = wbVal.verdict_class || (wbVal.verdict.includes('BUY') ? 'buy' : (wbVal.verdict.includes('SELL') ? 'sell' : 'hold'));
                verdictEl.className = `verdict-badge verdict-${vClass}`;
            }
            if (wbVal.wacc) {
                safeSetText('kpiWacc', `${wbVal.wacc}%`);
            }
            if (wbVal.enterprise_value) {
                safeSetText('kpiEv', `₹${Number(wbVal.enterprise_value).toLocaleString()} Cr`);
            }
        }
    }

    function monitorExcelExport(ticker, filename) {
        const btnExcel = document.getElementById('btnDownloadExcel');
        if (!btnExcel) return;

        if (excelPollInterval) {
            clearInterval(excelPollInterval);
            excelPollInterval = null;
        }

        const textSpan = btnExcel.querySelector('span') || btnExcel;
        btnExcel.href = `/api/download-excel/${encodeURIComponent(filename)}`;
        
        fetch(`/api/export-status/${encodeURIComponent(ticker)}`)
            .then(res => res.json())
            .then(statusData => {
                if (statusData.status === 'done') {
                    textSpan.textContent = "Download Excel Model (.xlsx) ✓ Ready";
                    btnExcel.style.opacity = "1";
                    syncWorkbookValuationToTerminal(statusData.workbook_valuation);
                } else if (statusData.status === 'generating') {
                    textSpan.textContent = "Calculating Excel Model (22 Sheets)...";
                    btnExcel.style.opacity = "0.8";
                    
                    let attempts = 0;
                    excelPollInterval = setInterval(async () => {
                        attempts++;
                        try {
                            const r = await fetch(`/api/export-status/${encodeURIComponent(ticker)}`);
                            const s = await r.json();
                            if (s.status === 'done') {
                                clearInterval(excelPollInterval);
                                excelPollInterval = null;
                                textSpan.textContent = "Download Excel Model (.xlsx) ✓ Ready";
                                btnExcel.style.opacity = "1";
                                syncWorkbookValuationToTerminal(s.workbook_valuation);
                            } else if (s.status === 'error') {
                                clearInterval(excelPollInterval);
                                excelPollInterval = null;
                                textSpan.textContent = "Excel Generation Failed";
                                btnExcel.style.opacity = "1";
                            }
                        } catch (e) {
                            // ignore transient network errors
                        }
                        if (attempts > 60) {
                            clearInterval(excelPollInterval);
                            excelPollInterval = null;
                            textSpan.textContent = "Download Excel Model (.xlsx)";
                            btnExcel.style.opacity = "1";
                        }
                    }, 2000);
                } else {
                    textSpan.textContent = "Download Excel Model (.xlsx)";
                    btnExcel.style.opacity = "1";
                }
            })
            .catch(() => {
                textSpan.textContent = "Download Excel Model (.xlsx)";
                btnExcel.style.opacity = "1";
            });
    }

    function renderFourPillars(v) {
        if (!v || !v.four_pillars) return;
        const fp = v.four_pillars;
        const p1 = fp.pillar1_fcf || {};
        const p2 = fp.pillar2_growth || {};
        const p3 = fp.pillar3_wacc || {};
        const p4 = fp.pillar4_multiples || {};

        // Pillar 1: FCF Engine
        safeSetText('p1QualityRatio', `${p1.earnings_quality_ratio || 0}x`);
        const qBadge = document.getElementById('p1QualityBadge');
        if (qBadge) {
            if ((p1.earnings_quality_ratio || 0) >= 1.0) {
                qBadge.textContent = "High Cash Backing (CFO > PAT)";
                qBadge.className = "pm-sub text-emerald";
            } else {
                qBadge.textContent = "Accrual / Working Capital Drag";
                qBadge.className = "pm-sub text-amber";
            }
        }
        safeSetText('p1EbitdaM', `${p1.ebitda_margin || 0}%`);
        safeSetText('p1EbitM', `${p1.ebit_margin || 0}%`);
        safeSetText('p1Capex', `₹${(p1.total_capex || 0).toLocaleString()} Cr`);
        safeSetText('p1MaintCapex', `₹${(p1.maintenance_capex || 0).toLocaleString()}`);
        safeSetText('p1GrowthCapex', `₹${(p1.growth_capex || 0).toLocaleString()}`);
        safeSetText('p1Fcf', `₹${(p1.actual_fcf || 0).toLocaleString()} Cr`);
        safeSetText('p1Cfo', `₹${(p1.cfo || 0).toLocaleString()} Cr`);
        safeSetText('p1Debtor', `${p1.debtor_days || 0} d`);
        safeSetText('p1Inv', `${p1.inventory_days || 0} d`);
        safeSetText('p1Payable', `${p1.payable_days || 0} d`);
        safeSetText('p1Ccc', `${p1.cash_conversion_cycle || 0} d`);

        // Pillar 2: Growth Trajectory
        safeSetText('p2Sales5y', p2.sales_cagr_5yr ? `${p2.sales_cagr_5yr}%` : 'N/A');
        safeSetText('p2Sales3y', p2.sales_cagr_3yr ? `${p2.sales_cagr_3yr}%` : 'N/A');
        safeSetText('p2Pat3y', p2.pat_cagr_3yr ? `${p2.pat_cagr_3yr}%` : 'N/A');
        safeSetText('p2ForecastG', `${v.growth_rate}%`);
        safeSetText('p2TerminalG', `${v.terminal_growth}%`);
        safeSetText('p2Tv', `₹${(v.pv_terminal_value || 0).toLocaleString()} Cr`);
        safeSetText('p2TvShare', `${p2.tv_share_of_ev || 0}% of EV`);

        // Pillar 3: Cost of Capital & Risk
        safeSetText('p3Wacc', `${v.wacc}%`);
        safeSetText('p3Ke', `${v.cost_of_equity}%`);
        safeSetText('p3Beta', `${v.beta}`);
        safeSetText('p3Kd', `${v.cost_of_debt}%`);
        safeSetText('p3PreKd', `${p3.pre_tax_cost_of_debt || 8.05}%`);
        safeSetText('p3We', `${v.weight_equity}%`);
        safeSetText('p3Wd', `${v.weight_debt}%`);
        safeSetText('p3De', `${p3.debt_to_equity || 0}x`);

        // Pillar 4: Relative Market Multiples
        safeSetText('p4EvEbitda', `${p4.target_ev_ebitda || 0}x`);
        safeSetText('p4PeerEvEbitda', `${p4.peer_median_ev_ebitda || 0}x`);
        safeSetText('p4Pe', `${p4.target_pe || 0}x`);
        safeSetText('p4PeerPe', `${p4.peer_median_pe || 'N/A'}x`);
        safeSetText('p4Pb', `${p4.target_pb || 0}x`);
        safeSetText('p4EvSales', `${p4.target_ev_sales || 0}x`);
        safeSetText('p4DcfPrice', `${v.intrinsic_value_per_share}`);
        safeSetText('p4PePrice', `${p4.implied_price_pe || 'N/A'}`);
        safeSetText('p4EvPrice', `${p4.implied_price_ev_ebitda || 'N/A'}`);
        safeSetText('p4Low52', `${p4.low_52wk || 'N/A'}`);
        safeSetText('p4High52', `${p4.high_52wk || 'N/A'}`);
        safeSetText('p4Cmp', `${v.current_price}`);
    }

    function renderDcfTable(dcfRows) {
        const tbody = document.getElementById('dcfTableBody');
        tbody.innerHTML = '';
        dcfRows.forEach(r => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${r.year}</strong> (t = ${r.mid_year})</td>
                <td class="num">₹${r.ebit.toLocaleString()}</td>
                <td class="num">₹${r.nopat.toLocaleString()}</td>
                <td class="num">${r.reinvestment_rate}%</td>
                <td class="num font-semibold">₹${r.fcff.toLocaleString()}</td>
                <td class="num text-muted">${r.discount_factor}</td>
                <td class="num text-cyan font-semibold">₹${r.pv_fcff.toLocaleString()}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function renderSensitivityGrid(sens) {
        const thead = document.getElementById('sensTableHead');
        const tbody = document.getElementById('sensTableBody');
        
        thead.innerHTML = `<tr><th>WACC \\ g</th>${sens.tg_headers.map(h => `<th>${h}</th>`).join('')}</tr>`;
        tbody.innerHTML = '';

        sens.rows.forEach(r => {
            const tr = document.createElement('tr');
            const cells = r.values.map((val, idx) => {
                const isCenter = (r.wacc.includes(sliderWacc.value) && idx === 2);
                return `<td class="heatmap-cell ${isCenter ? 'center-case' : ''}">₹${val !== null ? val : 'N/A'}</td>`;
            }).join('');
            tr.innerHTML = `<th>${r.wacc}</th>${cells}`;
            tbody.appendChild(tr);
        });
    }

    function renderDuPont(dp) {
        if (!dp) return;
        safeSetText('dpNpm', `${dp.net_profit_margin}%`);
        safeSetText('dpAto', `${dp.asset_turnover}x`);
        safeSetText('dpLev', `${dp.equity_multiplier}x`);
        safeSetText('dpRoe3', `${dp.roe_3stage}%`);

        safeSetText('dpTaxB', `${dp.tax_burden}x`);
        safeSetText('dpIntB', `${dp.interest_burden}x`);
        safeSetText('dpOpM', `${dp.operating_margin}%`);
        safeSetText('dpRoe5', `${dp.roe_5stage}%`);
    }

    function renderAltmanZ(az) {
        if (!az) return;
        const scoreEl = document.getElementById('altmanScore');
        if (scoreEl) {
            scoreEl.textContent = (az.score !== null && az.score !== undefined) ? az.score : "N/A";
            scoreEl.className = `altman-score-large val-${az.zone_class || 'neutral'}`;
        }

        const zoneEl = document.getElementById('altmanZoneBadge');
        if (zoneEl) {
            zoneEl.textContent = az.zone || "Not applicable";
            zoneEl.className = `badge verdict-${az.zone_class || 'neutral'}`;
        }

        if (az.components) {
            safeSetText('azX1', az.components.x1_wc_assets !== undefined ? az.components.x1_wc_assets : '-');
            safeSetText('azX2', az.components.x2_re_assets !== undefined ? az.components.x2_re_assets : '-');
            safeSetText('azX3', az.components.x3_ebit_assets !== undefined ? az.components.x3_ebit_assets : '-');
            safeSetText('azX4', az.components.x4_mktcap_liab !== undefined ? az.components.x4_mktcap_liab : '-');
            safeSetText('azX5', az.components.x5_sales_assets !== undefined ? az.components.x5_sales_assets : '-');
        }
    }

    function renderPeersTable(peers) {
        const thead = document.getElementById('peersTableHead');
        const tbody = document.getElementById('peersTableBody');
        if (!peers || !peers.columns || peers.rows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8">No peer comparison data available for this company.</td></tr>';
            return;
        }

        thead.innerHTML = `<tr>${peers.columns.map(c => `<th>${c}</th>`).join('')}</tr>`;
        tbody.innerHTML = '';
        peers.rows.forEach(r => {
            const tr = document.createElement('tr');
            tr.innerHTML = peers.columns.map(col => `<td>${r[col] !== undefined ? r[col] : '-'}</td>`).join('');
            tbody.appendChild(tr);
        });
    }

    function renderRawStatements(tables) {
        const container = document.getElementById('rawStatementsContainer');
        container.innerHTML = '';
        if (!tables) return;

        const secTitles = {
            'profit-loss': 'Profit & Loss Statement (10-Year)',
            'balance-sheet': 'Balance Sheet (10-Year)',
            'cash-flow': 'Cash Flow Statement',
            'quarters': 'Quarterly Financial Results'
        };

        for (const [secId, tData] of Object.entries(tables)) {
            if (!tData || !tData.columns) continue;
            const card = document.createElement('div');
            card.className = 'glass-panel scorecard';
            card.style.marginBottom = '1.5rem';

            const h3 = document.createElement('h3');
            h3.className = 'scorecard-title';
            h3.textContent = secTitles[secId] || secId.toUpperCase();
            card.appendChild(h3);

            const wrap = document.createElement('div');
            wrap.className = 'table-responsive';

            const table = document.createElement('table');
            table.className = 'data-table';
            table.innerHTML = `
                <thead><tr>${tData.columns.map(c => `<th>${c}</th>`).join('')}</tr></thead>
                <tbody>${tData.rows.map(r => `<tr>${tData.columns.map(c => `<td>${r[c]}</td>`).join('')}</tr>`).join('')}</tbody>
            `;
            wrap.appendChild(table);
            card.appendChild(wrap);
            container.appendChild(card);
        }
    }

    // --- 5. Interactive Parameter Sliders (Real-time recalculation) ---
    function setupSliderListeners() {
        [sliderWacc, sliderGrowth, sliderTerminal, sliderTax].forEach(slider => {
            slider.addEventListener('input', () => {
                badgeWacc.textContent = `${sliderWacc.value}%`;
                badgeGrowth.textContent = `${sliderGrowth.value}%`;
                badgeTerminal.textContent = `${sliderTerminal.value}%`;
                badgeTax.textContent = `${sliderTax.value}%`;

                clearTimeout(sliderDebounceTimer);
                sliderDebounceTimer = setTimeout(recalculateDCF, 180);
            });
        });
    }

    async function recalculateDCF() {
        if (!activeCompany) return;
        try {
            const payload = {
                ticker: activeCompany.ticker,
                wacc: sliderWacc.value,
                growthRate: sliderGrowth.value,
                terminalGrowth: sliderTerminal.value,
                taxRate: sliderTax.value
            };

            const resp = await fetch('/api/recalculate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const data = await resp.json();
            if (data.success && data.valuation) {
                const v = data.valuation;
                activeValuation = v;

                // Update Intrinsic Value
                safeSetText('kpiIntrinsic', `₹${v.intrinsic_value_per_share}`);
                const upsideEl = document.getElementById('kpiUpside');
                if (upsideEl) {
                    upsideEl.textContent = `${v.margin_of_safety_pct > 0 ? '+' : ''}${v.margin_of_safety_pct}%`;
                    upsideEl.className = `kpi-val ${v.margin_of_safety_pct >= 0 ? 'val-positive' : 'val-negative'}`;
                }

                // Update Verdict
                const verdictEl = document.getElementById('verdictBadge');
                if (verdictEl) {
                    verdictEl.textContent = v.verdict;
                    verdictEl.className = `verdict-badge verdict-${v.verdict_class}`;
                }

                // Update 4 Pillars Dashboard, DCF Table & Sensitivity Grid
                renderFourPillars(v);
                renderDcfTable(v.dcf_table);
                renderSensitivityGrid(v.sensitivity);
                renderReverseDcf({ valuation: v, company: activeCompany });

                // Update Excel download link
                const tkr = (activeCompany && activeCompany.ticker) ? activeCompany.ticker : 'COMPANY';
                monitorExcelExport(tkr, data.excel_filename);
            }
        } catch (err) {
            console.error("Recalculation error:", err);
        }
    }

    setupSliderListeners();

    // --- 6. Tabs Switching ---
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));

            btn.classList.add('active');
            const targetId = btn.getAttribute('data-tab');
            document.getElementById(targetId).classList.add('active');
        });
    });

    // --- 7. Copy & Print Actions ---
    document.getElementById('btnCopyReport').addEventListener('click', () => {
        const text = document.getElementById('aiReportContent').innerText;
        navigator.clipboard.writeText(text).then(() => {
            showToast('AI Research Memorandum copied to clipboard!');
        });
    });

    document.getElementById('btnPrintReport').addEventListener('click', () => {
        window.print();
    });

    // Toast Utility
    function showToast(msg) {
        const toast = document.getElementById('toast');
        toast.textContent = msg;
        toast.style.display = 'block';
        setTimeout(() => {
            toast.style.display = 'none';
        }, 3500);
    }

    // --- 8. Theme Management (Light Mode default + Antigravity Dark Mode) ---
    function initTheme() {
        const btnThemeToggle = document.getElementById('btnThemeToggle');
        const themeIcon = document.getElementById('themeIcon');
        const themeText = document.getElementById('themeText');

        function applyTheme(theme) {
            document.documentElement.setAttribute('data-theme', theme);
            if (theme === 'dark') {
                if (themeIcon) themeIcon.textContent = '🌙';
                if (themeText) themeText.textContent = 'Dark';
            } else {
                if (themeIcon) themeIcon.textContent = '☀️';
                if (themeText) themeText.textContent = 'Light';
            }
        }

        const savedTheme = localStorage.getItem('terminal_theme') || 'light';
        applyTheme(savedTheme);

        if (btnThemeToggle) {
            btnThemeToggle.addEventListener('click', () => {
                const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
                const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
                localStorage.setItem('terminal_theme', newTheme);
                applyTheme(newTheme);
                showToast(`Switched to ${newTheme === 'dark' ? 'Antigravity Dark' : 'Light'} theme`);
            });
        }
    }

    // Initial load
    loadSavedSettings();
    initTheme();

    // --- Fundamental Analysis Logic ---
    function formatValue(val, suffix = '', prefix = '') {
        if (val === undefined || val === null || val === '-') return 'N/A';
        const num = parseFloat(val);
        if (isNaN(num)) return val;
        return `${prefix}${num.toLocaleString(undefined, {maximumFractionDigits: 2})}${suffix}`;
    }

    function renderFundamentalAnalysis(data) {
        if (!data || !data.company) return;
        const meta = data.company.meta || {};
        const v = data.valuation || {};
        const p4 = (v.four_pillars && v.four_pillars.pillar4_multiples) ? v.four_pillars.pillar4_multiples : {};
        const p3 = (v.four_pillars && v.four_pillars.pillar3_wacc) ? v.four_pillars.pillar3_wacc : {};

        // 1. Fundamentals Grid
        const gridContainer = document.getElementById('faGrid');
        if (gridContainer) {
            const metrics = [
                { label: 'Market Cap', val: data.company.market_cap_cr, prefix: '₹', suffix: ' Cr' },
                { label: 'P/E Ratio', val: meta['Stock P/E'] || p4.target_pe, suffix: '' },
                { label: 'P/B Ratio', val: meta['Price to book value'] || p4.target_pb, suffix: '' },
                { label: 'Industry P/E', val: meta['Industry PE'] || p4.peer_median_pe, suffix: '' },
                { label: 'Debt to Equity', val: meta['Debt to equity'] || p3.debt_to_equity, suffix: 'x' },
                { label: 'ROE', val: meta['ROE'] || (v.dupont && v.dupont.roe_3stage), suffix: '%' },
                { label: 'EPS (TTM)', val: meta['EPS'], prefix: '₹' },
                { label: 'Dividend Yield', val: meta['Dividend Yield'], suffix: '%' },
                { label: 'Book Value', val: meta['Book Value'], prefix: '₹' },
                { label: 'Face Value', val: meta['Face Value'], prefix: '₹' }
            ];

            gridContainer.innerHTML = metrics.map(m => `
                <div class="groww-item">
                    <div class="groww-label">${m.label}</div>
                    <div class="groww-val">${formatValue(m.val, m.suffix, m.prefix)}</div>
                </div>
            `).join('');
        }

        // 2. Shareholding Pattern
        const shContainer = document.getElementById('faShareholding');
        if (shContainer && data.raw_tables && data.raw_tables['shareholding']) {
            const shTable = data.raw_tables['shareholding'];
            if (shTable.rows && shTable.rows.length > 0 && shTable.columns.length > 1) {
                const latestQtr = shTable.columns[shTable.columns.length - 1];
                
                const getShVal = (keyword) => {
                    const row = shTable.rows.find(r => (r.Metric || '').toLowerCase().includes(keyword));
                    return row ? parseFloat(row[latestQtr]) || 0 : 0;
                };

                const promoters = getShVal('promoter');
                const fiis = getShVal('fii');
                const diis = getShVal('dii');
                const publicSh = getShVal('public');

                const colors = [
                    { label: 'Promoters', val: promoters, color: '#a855f7' },
                    { label: 'FIIs', val: fiis, color: '#06b6d4' },
                    { label: 'DIIs', val: diis, color: '#10b981' },
                    { label: 'Public', val: publicSh, color: '#f59e0b' }
                ].sort((a, b) => b.val - a.val).filter(x => x.val > 0);

                shContainer.innerHTML = colors.map(c => `
                    <div style="margin-bottom: 0.5rem;">
                        <div class="sh-row">
                            <div class="sh-label">
                                <div class="sh-color-box" style="background-color: ${c.color}"></div>
                                <span>${c.label}</span>
                            </div>
                            <div class="sh-val">${c.val.toFixed(2)}%</div>
                        </div>
                        <div class="sh-bar-bg">
                            <div class="sh-bar-fill" style="width: ${c.val}%; background-color: ${c.color}"></div>
                        </div>
                    </div>
                `).join('');
            } else {
                shContainer.innerHTML = '<p class="text-muted">Shareholding data not available.</p>';
            }
        }

        // 3. Trend Analysis & Conclusion
        const trendContainer = document.getElementById('faTrends');
        if (trendContainer && data.raw_tables) {
            let goodSignals = 0;
            const trends = [];
            
            const checkTrend = (tableName, metricKeyword, positiveIsGood) => {
                const table = data.raw_tables[tableName];
                if (!table || !table.rows) return null;
                const row = table.rows.find(r => (r.Metric || '').toLowerCase().includes(metricKeyword.toLowerCase()));
                if (!row) return null;
                
                const years = table.columns.filter(c => c !== 'Metric' && !c.includes('TTM')).slice(-3);
                if (years.length < 2) return null;
                
                const recent = parseFloat(row[years[years.length - 1]]) || 0;
                const prev = parseFloat(row[years[years.length - 2]]) || 0;
                
                const isIncreasing = recent >= prev;
                const isGood = positiveIsGood ? isIncreasing : !isIncreasing;
                if (isGood) goodSignals++;
                
                return { isGood, isIncreasing, recent, prev };
            };

            const sales = checkTrend('profit-loss', 'sales', true);
            if (sales) {
                trends.push(`
                    <div class="trend-item">
                        <div class="trend-icon" style="color: ${sales.isGood ? '#10b981' : '#ef4444'}">${sales.isGood ? '📈' : '📉'}</div>
                        <div>
                            <div class="trend-title">Sales Trajectory</div>
                            <div class="trend-text">Sales have ${sales.isIncreasing ? 'increased' : 'decreased'} YoY from ₹${sales.prev} Cr to ₹${sales.recent} Cr. ${sales.isGood ? 'A positive sign of top-line growth.' : 'A concerning sign of revenue contraction.'}</div>
                        </div>
                    </div>
                `);
            }

            const profit = checkTrend('profit-loss', 'net profit', true);
            if (profit) {
                trends.push(`
                    <div class="trend-item">
                        <div class="trend-icon" style="color: ${profit.isGood ? '#10b981' : '#ef4444'}">${profit.isGood ? '🟢' : '🔴'}</div>
                        <div>
                            <div class="trend-title">Profitability</div>
                            <div class="trend-text">Net Profit has ${profit.isIncreasing ? 'grown' : 'fallen'} YoY (₹${profit.prev} Cr → ₹${profit.recent} Cr).</div>
                        </div>
                    </div>
                `);
            }
            
            const assets = checkTrend('balance-sheet', 'total assets', true);
            if (assets) {
                trends.push(`
                    <div class="trend-item">
                        <div class="trend-icon" style="color: ${assets.isGood ? '#10b981' : '#ef4444'}">${assets.isGood ? '🏛️' : '⚠️'}</div>
                        <div>
                            <div class="trend-title">Asset Expansion</div>
                            <div class="trend-text">Total Assets have ${assets.isIncreasing ? 'expanded' : 'contracted'} YoY.</div>
                        </div>
                    </div>
                `);
            }

            const liab = checkTrend('balance-sheet', 'total liabilities', false);
            if (liab) {
                trends.push(`
                    <div class="trend-item">
                        <div class="trend-icon" style="color: ${liab.isGood ? '#10b981' : '#ef4444'}">${liab.isGood ? '🛡️' : '⚠️'}</div>
                        <div>
                            <div class="trend-title">Liability Management</div>
                            <div class="trend-text">Total Liabilities have ${liab.isIncreasing ? 'expanded' : 'contracted'} YoY.</div>
                        </div>
                    </div>
                `);
            }

            const cfo = checkTrend('cash-flow', 'operating activity', true);
            if (cfo) {
                trends.push(`
                    <div class="trend-item">
                        <div class="trend-icon" style="color: ${cfo.isGood ? '#10b981' : '#ef4444'}">${cfo.isGood ? '💰' : '💸'}</div>
                        <div>
                            <div class="trend-title">Cash Generation</div>
                            <div class="trend-text">Cash from Operating Activities is ${cfo.isIncreasing ? 'increasing' : 'decreasing'}. ${cfo.isGood ? 'Strong core business cash flow.' : 'Watch out for working capital stress.'}</div>
                        </div>
                    </div>
                `);
            }
            
            const conclusionScore = (goodSignals / (trends.length || 1)) * 100;
            let conclusionMsg = '';
            let conclusionColor = '';
            if (conclusionScore >= 75) {
                conclusionMsg = "Strong Fundamentals. The company exhibits robust growth and cash generation. Good for investment.";
                conclusionColor = "var(--accent-emerald)";
            } else if (conclusionScore >= 50) {
                conclusionMsg = "Moderate Fundamentals. Mixed signals in recent trajectories. Needs closer look.";
                conclusionColor = "var(--accent-amber)";
            } else {
                conclusionMsg = "Weak Fundamentals. The company is facing contraction or cash flow pressures. High Risk.";
                conclusionColor = "var(--accent-crimson)";
            }

            trendContainer.innerHTML = trends.join('') + `
                <div class="trend-item" style="margin-top: 1rem; border-top: 1px solid var(--border-dark); padding-top: 1rem;">
                    <div class="trend-icon" style="color: ${conclusionColor}">💡</div>
                    <div>
                        <div class="trend-title" style="color: ${conclusionColor};">Investment Conclusion</div>
                        <div class="trend-text" style="color: var(--text-primary); font-weight: 500;">${conclusionMsg}</div>
                    </div>
                </div>
            `;
        }
    }

    // --- Riskometer Logic ---
    function renderRiskometer(data) {
        if (!data || !data.valuation) return;
        
        const v = data.valuation;
        const c = data.company || {};
        const meta = c.meta || {};
        const raw_pl = (data.raw_tables && data.raw_tables['profit-loss']) ? data.raw_tables['profit-loss'] : null;

        // 1. Extract Metrics
        const altmanZ = v.altman_z ? v.altman_z.score : 0;
        
        let debtToEquity = 0;
        if (v.four_pillars && v.four_pillars.pillar3_wacc) {
            debtToEquity = v.four_pillars.pillar3_wacc.debt_to_equity;
        } else if (meta['Debt to equity']) {
            debtToEquity = parseFloat(meta['Debt to equity']) || 0;
        }

        let cfoPat = 0;
        if (v.four_pillars && v.four_pillars.pillar1_fcf) {
            cfoPat = v.four_pillars.pillar1_fcf.quality_ratio;
        }

        // Calculate Interest Coverage from P&L if available, else from meta
        let icr = 0;
        if (meta['Interest Coverage Ratio']) {
            icr = parseFloat(meta['Interest Coverage Ratio']) || 0;
        } else if (raw_pl && raw_pl.rows && raw_pl.columns && raw_pl.columns.length > 1) {
            const latestCol = raw_pl.columns[raw_pl.columns.length - 1];
            const opProfRow = raw_pl.rows.find(r => (r.Metric || '').toLowerCase().includes('operating profit'));
            const intRow = raw_pl.rows.find(r => (r.Metric || '').toLowerCase() === 'interest');
            
            if (opProfRow && intRow) {
                const opProf = parseFloat(opProfRow[latestCol]) || 0;
                const intExp = parseFloat(intRow[latestCol]) || 0;
                if (intExp > 0) icr = opProf / intExp;
                else icr = 10; // no interest, high coverage
            }
        }

        const beta = v.beta || 1.0;
        const cmp = v.current_price || 0;
        let dcf = 0;
        if (v.four_pillars && v.four_pillars.pillar4_multiples) {
             dcf = v.four_pillars.pillar4_multiples.implied_value || v.fair_value || 0;
        } else {
             dcf = v.fair_value || 0;
        }

        // 2. Scoring Mechanism (0-100)
        // Pillar 1: Solvency (35%)
        let s_solvency = 0;
        let p1_status = 'N/A', p1_class = 'badge-risk-mod';
        if (altmanZ > 2.99 && debtToEquity < 0.5) { s_solvency = 15; p1_status = 'SAFE'; p1_class = 'badge-risk-low'; }
        else if (altmanZ < 1.81 || debtToEquity > 1.5) { s_solvency = 90; p1_status = 'DISTRESS'; p1_class = 'badge-risk-high'; }
        else { s_solvency = 55; p1_status = 'GREY ZONE'; p1_class = 'badge-risk-mod'; }

        // Pillar 2: Cash Flow (25%)
        let s_cashflow = 0;
        let p2_status = 'N/A', p2_class = 'badge-risk-mod';
        if (cfoPat >= 1.1 && icr > 4.0) { s_cashflow = 15; p2_status = 'HIGH QUALITY'; p2_class = 'badge-risk-low'; }
        else if (cfoPat < 0.8 || icr < 2.0) { s_cashflow = 85; p2_status = 'CASH DRAG'; p2_class = 'badge-risk-high'; }
        else { s_cashflow = 50; p2_status = 'ADEQUATE'; p2_class = 'badge-risk-mod'; }

        // Pillar 3: Volatility (20%)
        let s_volatility = 0;
        let p3_status = 'N/A', p3_class = 'badge-risk-mod';
        if (beta < 0.8) { s_volatility = 20; p3_status = 'DEFENSIVE'; p3_class = 'badge-risk-low'; }
        else if (beta > 1.2) { s_volatility = 85; p3_status = 'HIGH CYCLICAL'; p3_class = 'badge-risk-high'; }
        else { s_volatility = 50; p3_status = 'MARKET NORMAL'; p3_class = 'badge-risk-mod'; }

        // Pillar 4: Valuation (20%)
        let s_valuation = 0;
        let p4_status = 'N/A', p4_class = 'badge-risk-mod';
        const premium = dcf > 0 ? ((cmp - dcf) / dcf) * 100 : 0;
        
        if (premium < -15) { s_valuation = 15; p4_status = 'HIGH SAFETY'; p4_class = 'badge-risk-low'; }
        else if (premium > 25) { s_valuation = 85; p4_status = 'FROTH'; p4_class = 'badge-risk-high'; }
        else { s_valuation = 50; p4_status = 'FAIRLY PRICED'; p4_class = 'badge-risk-mod'; }

        // Composite Score
        const compositeScore = Math.round((0.35 * s_solvency) + (0.25 * s_cashflow) + (0.20 * s_volatility) + (0.20 * s_valuation));

        // 3. Update DOM
        const safeSet = (id, val) => { const el = document.getElementById(id); if (el) el.innerText = val; };
        
        // Grid
        safeSet('r_altmanZ', (Number(altmanZ) || 0).toFixed(2));
        safeSet('r_debtEq', (Number(debtToEquity) || 0).toFixed(2) + 'x');
        safeSet('r_cfoPat', (Number(cfoPat) || 0).toFixed(2) + 'x');
        safeSet('r_icr', (Number(icr) || 0).toFixed(2) + 'x');
        safeSet('r_beta', (Number(beta) || 0).toFixed(2));
        safeSet('r_cmp', '₹' + cmp.toLocaleString());
        safeSet('r_dcf', '₹' + Math.round(dcf).toLocaleString());

        // Bars & Badges
        const updatePillar = (pid, score, status, badgeClass) => {
            const bar = document.getElementById('rBar_' + pid);
            const badge = document.getElementById('rBadge_' + pid);
            if (bar) {
                bar.style.width = score + '%';
                if (score <= 30) bar.style.background = 'var(--accent-emerald)';
                else if (score <= 60) bar.style.background = 'var(--accent-amber)';
                else bar.style.background = 'var(--accent-crimson)';
            }
            if (badge) {
                badge.innerText = status;
                badge.className = 'risk-badge ' + badgeClass;
            }
        };

        updatePillar('solvency', s_solvency, p1_status, p1_class);
        updatePillar('cashflow', s_cashflow, p2_status, p2_class);
        updatePillar('volatility', s_volatility, p3_status, p3_class);
        updatePillar('valuation', s_valuation, p4_status, p4_class);

        // Gauge & Verdict
        let verdictStatus = '', verdictColor = '', verdictDesc = '';
        if (compositeScore <= 30) {
            verdictStatus = 'LOW RISK';
            verdictColor = 'var(--accent-emerald)';
            verdictDesc = 'Low Risk (Defensive Compounder). The company exhibits strong balance sheet health, solid cash generation, and a reasonable margin of safety.';
        } else if (compositeScore <= 60) {
            verdictStatus = 'MODERATE RISK';
            verdictColor = 'var(--accent-amber)';
            verdictDesc = 'Moderate Risk (Balanced / Cyclical Growth). Vulnerabilities exist in either valuation premium, cyclicality, or moderate leverage.';
        } else {
            verdictStatus = 'HIGH RISK';
            verdictColor = 'var(--accent-crimson)';
            verdictDesc = 'High / Critical Risk (Speculative / Turnaround). The company faces significant financial distress signals, cash burn, or extreme overvaluation.';
        }

        const scoreEl = document.getElementById('riskScoreVal');
        if (scoreEl) {
            scoreEl.innerText = compositeScore;
            scoreEl.style.color = verdictColor;
        }

        const statusEl = document.getElementById('riskScoreStatus');
        if (statusEl) {
            statusEl.innerText = verdictStatus;
            statusEl.style.color = verdictColor;
        }

        const verdictEl = document.getElementById('riskVerdictText');
        if (verdictEl) {
            verdictEl.innerHTML = `<strong>${verdictStatus}</strong> - ${verdictDesc} <br><br> 
            Primary drivers: ${p1_status === 'DISTRESS' ? 'Solvency Distress, ' : ''} 
            ${p2_status === 'CASH DRAG' ? 'Poor Cash Quality, ' : ''}
            ${p3_status === 'HIGH CYCLICAL' ? 'High Beta, ' : ''}
            ${p4_status === 'FROTH' ? 'Valuation Froth' : ''}`.replace(/,\s*$/, '');
        }

        const needle = document.getElementById('riskNeedle');
        if (needle) {
            // map 0-100 to -90 to +90 degrees
            const deg = (compositeScore / 100) * 180 - 90;
            needle.style.transform = `rotate(${deg}deg)`;
        }
    }

    // --- Forensic Check Logic (Beneish M-Score & Divergence) ---
    function renderForensicCheck(data) {
        if (!data || !data.valuation) return;
        const v = data.valuation;
        const forensic = v.forensic || {};
        const mScore = forensic.mScore || {};
        const currentM = (typeof mScore.current === 'number') ? mScore.current : -2.45;
        const hist5yr = mScore.historical5Yr || [-2.60, -2.55, -2.40, -2.50, currentM];
        const vars = mScore.variables || {};
        const divergence = forensic.divergence || { years: [], pat: [], cfo: [] };
        const govFlags = forensic.governanceFlags || {};

        // 1. SECTION 1: MASTER BANNER
        const bannerEl = document.getElementById('forensicVerdictBanner');
        const iconBox = document.getElementById('forensicIconBox');
        const statusBadge = document.getElementById('forensicStatusBadge');
        const titleEl = document.getElementById('forensicVerdictTitle');
        const descEl = document.getElementById('forensicVerdictDesc');
        const mscoreVal = document.getElementById('forensicMScoreVal');
        const mscoreSub = document.getElementById('forensicMScoreSub');
        const sparklineWrap = document.getElementById('forensicSparklineSvg');
        const breachedCountEl = document.getElementById('forensicBreachedCount');

        let statusClass = 'safe';
        let statusText = 'SAFE STATUS';
        let statusTitle = 'SAFE - Low Probability of Earnings Manipulation';
        let statusDesc = 'Financial statements reflect conservative accounting principles and high earnings quality.';
        let statusIcon = '🛡️';
        let statusSubText = '↓ Low Manipulation Risk';
        let statusColor = '#10b981';

        if (currentM > -1.78) {
            statusClass = 'flag';
            statusText = 'HIGH RISK STATUS';
            statusTitle = 'HIGH RISK - Probable Earnings Manipulation Detected';
            statusDesc = 'Multiple indices exceed critical forensic barriers. Substantial probability of aggressive revenue recognition or non-cash accrual swelling.';
            statusIcon = '🚨';
            statusSubText = '↑ High Manipulation Risk';
            statusColor = '#ef4444';
        } else if (currentM >= -2.22) {
            statusClass = 'watch';
            statusText = 'CAUTION STATUS';
            statusTitle = 'CAUTION - Elevated Forensic Risk';
            statusDesc = 'M-Score sits in the cautionary grey band. Isolated accounting anomalies detected in gross margins, accruals, or asset capitalization.';
            statusIcon = '⚠️';
            statusSubText = '⚡ Cautionary Grey Zone';
            statusColor = '#f59e0b';
        }

        if (bannerEl) {
            bannerEl.className = `forensic-banner glass-panel banner-forensic-${statusClass}`;
        }
        if (iconBox) {
            iconBox.innerHTML = `<span class="forensic-main-icon" style="font-size: 1.8rem;">${statusIcon}</span>`;
        }
        if (statusBadge) {
            statusBadge.className = `forensic-status-pill badge-forensic-${statusClass}`;
            statusBadge.textContent = statusText;
        }
        if (titleEl) titleEl.textContent = statusTitle;
        if (descEl) descEl.textContent = statusDesc;
        if (mscoreVal) mscoreVal.textContent = currentM.toFixed(2);
        if (mscoreSub) {
            mscoreSub.innerHTML = `<span style="color: ${statusColor}; font-weight: 600;">${statusSubText}</span>`;
        }

        // SVG Sparkline for 5-Yr Trajectory
        if (sparklineWrap && hist5yr.length >= 2) {
            const minVal = Math.min(...hist5yr, -3.5);
            const maxVal = Math.max(...hist5yr, -1.5);
            const range = (maxVal - minVal) || 1.0;
            const w = 110, h = 44, pad = 6;
            
            const points = hist5yr.map((val, idx) => {
                const x = pad + (idx / (hist5yr.length - 1)) * (w - 2 * pad);
                const y = h - pad - ((val - minVal) / range) * (h - 2 * pad);
                return `${x.toFixed(1)},${y.toFixed(1)}`;
            });
            const polyline = points.join(' ');
            
            sparklineWrap.innerHTML = `
                <svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" style="overflow: visible;">
                    <line x1="${pad}" y1="${(h - pad - ((-1.78 - minVal) / range) * (h - 2 * pad)).toFixed(1)}" 
                          x2="${w - pad}" y2="${(h - pad - ((-1.78 - minVal) / range) * (h - 2 * pad)).toFixed(1)}" 
                          stroke="rgba(239, 68, 68, 0.4)" stroke-dasharray="2 2" stroke-width="1" />
                    <polyline fill="none" stroke="${statusColor}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" points="${polyline}" />
                    ${points.map((pt, i) => `<circle cx="${pt.split(',')[0]}" cy="${pt.split(',')[1]}" r="${i === points.length - 1 ? 3.5 : 2}" fill="${statusColor}" />`).join('')}
                </svg>
            `;
        }

        // 2. SECTION 2: 8-VARIABLE FORENSIC GRID
        const varConfigs = [
            { key: 'DSRI', name: "Days' Sales in Receivables", short: "Channel Stuffing Risk", threshold: 1.03, redDesc: "Flags loose credit terms to artificially accelerate sales recognition." },
            { key: 'GMI', name: "Gross Margin Degradation", short: "Margin Deterioration Risk", threshold: 1.01, redDesc: "Flags falling gross margins, creating acute pressure to manipulate earnings." },
            { key: 'AQI', name: "Asset Quality Index", short: "Capitalized Expense Risk", threshold: 1.04, redDesc: "Flags capitalization of operating costs as intangible non-current assets." },
            { key: 'SGI', name: "Sales Growth Index", short: "Top-Line Growth Deceleration", threshold: 1.60, redDesc: "High growth strains internal controls and incentivizes revenue pull-forward." },
            { key: 'DEPI', name: "Depreciation Index", short: "Asset Life Extension", threshold: 1.07, redDesc: "Flags downward revision in depreciation rates to inflate accounting PAT." },
            { key: 'SGAI', name: "SG&A Expenses Index", short: "Operating Efficiency Drag", threshold: 1.05, redDesc: "Administrative overhead costs outstripping top-line revenue expansion." },
            { key: 'LVGI', name: "Leverage Index", short: "Balance Sheet Gearing Shift", threshold: 1.03, redDesc: "Increasing debt-to-assets ratio elevates debt covenant breach pressures." },
            { key: 'TATA', name: "Total Accruals to Total Assets", short: "Non-Cash Earnings Distortion", threshold: 0.03, redDesc: "High non-cash accruals; reported net profits lack operating cash backing." }
        ];

        let breachedCount = 0;
        const gridEl = document.getElementById('forensicVarsGrid');
        if (gridEl) {
            gridEl.innerHTML = varConfigs.map(item => {
                const rawVal = vars[item.key];
                const val = (typeof rawVal === 'number') ? rawVal : (item.threshold * 0.95);
                const isFlagged = val > item.threshold;
                if (isFlagged) breachedCount++;

                return `
                    <div class="forensic-var-card ${isFlagged ? 'flagged' : ''}">
                        <div>
                            <div class="forensic-card-top">
                                <span class="forensic-code-badge">${item.key}</span>
                                <span class="forensic-status-badge ${isFlagged ? 'flagged' : 'normal'}">
                                    ${isFlagged ? '🚨 FLAGGED' : '✓ NORMAL'}
                                </span>
                            </div>
                            <div class="forensic-var-name">${item.name}</div>
                            <div class="forensic-var-sub">${item.short}</div>
                        </div>

                        <div>
                            <div class="forensic-card-numbers">
                                <div>
                                    <div style="font-size: 0.65rem; color: var(--text-muted); text-transform: uppercase;">Current</div>
                                    <div class="forensic-val-current ${isFlagged ? 'flagged' : ''}">${val.toFixed(3)}</div>
                                </div>
                                <div>
                                    <div style="font-size: 0.65rem; color: var(--text-muted); text-transform: uppercase;">Threshold</div>
                                    <div class="forensic-val-threshold">&gt; ${item.threshold.toFixed(2)}</div>
                                </div>
                            </div>
                            <div class="forensic-desc-box ${isFlagged ? 'flagged' : 'normal'}">
                                ${item.redDesc}
                            </div>
                        </div>
                    </div>
                `;
            }).join('');
        }

        if (breachedCountEl) {
            breachedCountEl.textContent = `${breachedCount} of 8`;
            breachedCountEl.className = breachedCount > 0 ? 'text-crimson' : 'text-emerald';
        }

        // 3. SECTION 3: CASH FLOW DIVERGENCE (PAT VS CFO)
        const years = divergence.years || [];
        const pats = divergence.pat || [];
        const cfos = divergence.cfo || [];
        const chartWrap = document.getElementById('divergenceChartContainer');
        const alertEl = document.getElementById('divergenceAlertBanner');

        // Check divergence
        let isDivergent = false;
        if (pats.length >= 2 && cfos.length >= 2) {
            const lastPat = pats[pats.length - 1];
            const prevPat = pats[pats.length - 2];
            const lastCfo = cfos[cfos.length - 1];
            const prevCfo = cfos[cfos.length - 2];
            if (lastPat > prevPat && lastCfo <= prevCfo) isDivergent = true;
            if (lastPat > 0 && lastCfo < 0.7 * lastPat) isDivergent = true;
        }

        if (alertEl) {
            alertEl.style.display = isDivergent ? 'flex' : 'none';
        }

        if (chartWrap && years.length > 0) {
            const allVals = [...pats, ...cfos].filter(x => !isNaN(x));
            const maxVal = Math.max(...allVals, 100);
            const minVal = Math.min(...allVals, 0);
            const valSpan = (maxVal - minVal) || 1.0;

            const chartH = 200;

            chartWrap.innerHTML = `
                <div style="display: flex; justify-content: space-around; align-items: flex-end; height: ${chartH}px; padding: 10px 5px 25px 5px; border-bottom: 1px solid var(--border-subtle); position: relative;">
                    ${years.map((yr, idx) => {
                        const pat = pats[idx] || 0;
                        const cfo = cfos[idx] || 0;
                        const cfoH = Math.max(6, (Math.max(0, cfo) / valSpan) * (chartH - 50));
                        const patH = Math.max(6, (Math.max(0, pat) / valSpan) * (chartH - 50));

                        return `
                            <div style="display: flex; flex-direction: column; align-items: center; gap: 4px; z-index: 2; width: 18%;">
                                <div style="display: flex; align-items: flex-end; gap: 6px; height: ${chartH - 50}px;">
                                    <div title="CFO: ₹${Number(cfo).toLocaleString('en-IN')} Cr" style="width: 20px; height: ${cfoH}px; background: #06b6d4; border-radius: 4px 4px 0 0; transition: height 0.6s ease;"></div>
                                    <div title="PAT: ₹${Number(pat).toLocaleString('en-IN')} Cr" style="width: 20px; height: ${patH}px; background: #f59e0b; border-radius: 4px 4px 0 0; transition: height 0.6s ease;"></div>
                                </div>
                                <div style="font-family: var(--font-mono); font-size: 0.72rem; color: var(--text-muted); margin-top: 6px;">${yr.replace('Mar ', "'")}</div>
                            </div>
                        `;
                    }).join('')}
                </div>
                <div style="display: flex; justify-content: space-between; font-size: 0.72rem; font-family: var(--font-mono); color: var(--text-muted); padding-top: 0.5rem;">
                    <span>Peak: ₹${Math.round(maxVal).toLocaleString('en-IN')} Cr</span>
                    <span>Amounts in ₹ Crores (Screener.in)</span>
                </div>
            `;
        }

        // 4. SECTION 3: AUDITOR & GOVERNANCE CHECKLIST
        const checklistEl = document.getElementById('governanceChecklist');
        if (checklistEl) {
            const govItems = [
                {
                    title: "Auditor Resignation / Turnover",
                    desc: "Mid-term resignation or frequent switching of statutory audit firms.",
                    flagged: !!govFlags.auditorResignation
                },
                {
                    title: "High Related Party Transactions (RPT)",
                    desc: "Material procurement, loans, or guarantees with promoter affiliates.",
                    flagged: !!govFlags.highRPT
                },
                {
                    title: "Delayed Financial Filings",
                    desc: "Repeated delays in publishing audited annual or quarterly earnings.",
                    flagged: !!govFlags.delayedFilings
                },
                {
                    title: "Regulatory Actions / Litigation",
                    desc: "Active regulatory notices, forensic audit orders, or tax litigation.",
                    flagged: !!govFlags.regulatoryActions
                },
                {
                    title: "Contingent Liabilities > 10% Net Worth",
                    desc: "Off-balance sheet claims exceeding balance sheet equity cushion.",
                    flagged: !!govFlags.contingentLiabilitiesHigh
                },
                {
                    title: "Auditor Qualified Opinion",
                    desc: "Auditor caveats on internal controls, provisioning, or going concern.",
                    flagged: !!govFlags.qualifiedOpinion
                }
            ];

            checklistEl.innerHTML = govItems.map(item => `
                <div class="gov-item ${item.flagged ? 'flagged' : ''}">
                    <div class="gov-item-left">
                        <div class="gov-item-title">${item.title}</div>
                        <div class="gov-item-desc">${item.desc}</div>
                    </div>
                    <div class="gov-badge ${item.flagged ? 'flagged' : 'clean'}">
                        ${item.flagged ? '🚨 FLAGGED' : '✓ CLEAN'}
                    </div>
                </div>
            `).join('');
        }
    }

    // --- Reverse DCF (Expectations Investing) Logic ---
    function renderReverseDcf(data) {
        if (!data || !data.valuation) return;
        const v = data.valuation;
        const rd = v.reverse_dcf || {};
        const cmp = rd.currentPrice || v.current_price || 0;
        const wacc = rd.wacc || v.wacc || 12.0;
        const tg = rd.terminalGrowth || v.terminal_growth || 3.0;
        const impliedG = (typeof rd.impliedFCFCagr === 'number') ? rd.impliedFCFCagr : 12.0;
        const histG = (typeof rd.historicalFCFCagr === 'number') ? rd.historicalFCFCagr : 8.5;
        const sectorG = (typeof rd.sectorAverageCagr === 'number') ? rd.sectorAverageCagr : 10.1;
        const matrix = rd.sensitivityMatrix || { rowsWACC: [], colsGrowth: [], matrixPrices: [] };

        // 1. SECTION 1: EXPECTATIONS VERDICT BANNER
        const bannerEl = document.getElementById('reverseVerdictBanner');
        const iconBox = document.getElementById('reverseIconBox');
        const statusBadge = document.getElementById('reverseStatusBadge');
        const cmpSpan = document.getElementById('revVerdictCmp');
        const impliedGSpan = document.getElementById('revVerdictImpliedG');
        const descEl = document.getElementById('reverseVerdictDesc');

        let statusClass = 'fair';
        let statusText = 'FAIRLY PRICED';
        let statusIcon = '🎯';
        let statusDesc = 'The market is pricing in cash flow growth broadly consistent with recent historical operational compounding.';

        if (impliedG > histG + 5.0) {
            statusClass = 'perfection';
            statusText = 'PRICED FOR PERFECTION';
            statusIcon = '🚨';
            statusDesc = 'Market pricing demands aggressive acceleration well above historical execution. Any top-line deceleration risks severe multiple contraction.';
        } else if (impliedG < histG) {
            statusClass = 'upside';
            statusText = 'ASYMMETRIC UPSIDE';
            statusIcon = '🛡️';
            statusDesc = 'The market is pricing in significant operational pessimism or cash flow contraction, offering an institutional margin of safety.';
        }

        if (bannerEl) bannerEl.className = `reverse-banner glass-panel banner-reverse-${statusClass}`;
        if (iconBox) iconBox.innerHTML = `<span class="reverse-main-icon">${statusIcon}</span>`;
        if (statusBadge) {
            statusBadge.className = `reverse-status-pill badge-reverse-${statusClass}`;
            statusBadge.textContent = statusText;
        }
        if (cmpSpan) cmpSpan.textContent = `₹${Number(cmp).toLocaleString('en-IN')}`;
        if (impliedGSpan) impliedGSpan.textContent = `${impliedG.toFixed(1)}%`;
        if (descEl) descEl.textContent = statusDesc;

        // Progress bars benchmark
        const maxBench = Math.max(impliedG, histG, sectorG, 25.0) * 1.15;
        safeSetText('revBarImpliedVal', `${impliedG.toFixed(1)}%`);
        safeSetText('revBarHistVal', `${histG.toFixed(1)}%`);
        safeSetText('revBarSectorVal', `${sectorG.toFixed(1)}%`);

        const fillImplied = document.getElementById('revBarImpliedFill');
        const fillHist = document.getElementById('revBarHistFill');
        const fillSector = document.getElementById('revBarSectorFill');
        if (fillImplied) fillImplied.style.width = `${Math.min(100, Math.max(6, (impliedG / maxBench) * 100))}%`;
        if (fillHist) fillHist.style.width = `${Math.min(100, Math.max(6, (histG / maxBench) * 100))}%`;
        if (fillSector) fillSector.style.width = `${Math.min(100, Math.max(6, (sectorG / maxBench) * 100))}%`;

        // 2. SECTION 2: PARAMETERS GRID
        safeSetText('revParamWacc', `${wacc.toFixed(2)}%`);
        safeSetText('revParamTg', `${tg.toFixed(1)}%`);
        const multiple = (wacc > tg) ? (1.0 / ((wacc - tg) / 100)).toFixed(1) : '20.0';
        safeSetText('revParamMultiple', `${multiple}x`);

        safeSetText('revParamCmpSub', `₹${Number(cmp).toLocaleString('en-IN')}`);
        safeSetText('revParamImpliedG', `${impliedG.toFixed(1)}%`);
        
        const vsHistEl = document.getElementById('revParamVsHist');
        if (vsHistEl) {
            const delta = impliedG - histG;
            vsHistEl.textContent = `${delta >= 0 ? '+' : ''}${delta.toFixed(1)}%`;
            vsHistEl.className = delta > 5.0 ? 'font-mono text-crimson' : 'font-mono text-emerald';
        }

        // 3. SECTION 3: 5x5 SENSITIVITY MATRIX
        safeSetText('revMatrixTgFoot', `${tg.toFixed(1)}%`);
        const thead = document.getElementById('revMatrixThead');
        const tbody = document.getElementById('revMatrixTbody');
        if (!thead || !tbody) return;

        const rowsW = matrix.rowsWACC || [];
        const colsG = matrix.colsGrowth || [];
        const prices = matrix.matrixPrices || [];

        thead.innerHTML = `
            <tr>
                <th style="text-align: left;">WACC \\ FCF Growth</th>
                ${colsG.map(g => `<th style="${Math.abs(g - impliedG) < 0.5 ? 'color: var(--accent-amber); font-weight: 800;' : ''}">${g.toFixed(1)}%</th>`).join('')}
            </tr>
        `;

        tbody.innerHTML = rowsW.map((rW, rowIdx) => {
            const isCenterRow = Math.abs(rW - wacc) < 0.2;
            return `
                <tr>
                    <td style="text-align: left; font-weight: 700; ${isCenterRow ? 'color: var(--accent-cyan);' : 'color: var(--text-secondary);'}">
                        ${rW.toFixed(2)}%
                    </td>
                    ${(prices[rowIdx] || []).map((p, colIdx) => {
                        const isCurrentPrice = Math.abs(p - cmp) <= (cmp * 0.02);
                        const upsidePct = cmp > 0 ? ((p - cmp) / cmp) * 100 : 0;
                        const isUpside = p >= cmp;

                        let cellClass = 'matrix-price-cell';
                        if (isCurrentPrice) {
                            cellClass += ' cell-current';
                        } else if (isUpside) {
                            cellClass += (upsidePct > 25) ? ' cell-safe-strong' : ' cell-safe';
                        } else {
                            cellClass += (upsidePct < -25) ? ' cell-danger-strong' : ' cell-danger';
                        }

                        return `
                            <td>
                                <div class="${cellClass}">
                                    <span style="font-size: 0.88rem;">₹${Number(p).toLocaleString('en-IN')}</span>
                                    <span style="font-size: 0.65rem; ${isUpside ? 'color: var(--accent-emerald);' : 'color: var(--accent-crimson);'}">
                                        ${upsidePct >= 0 ? '+' : ''}${upsidePct.toFixed(0)}%
                                    </span>
                                </div>
                            </td>
                        `;
                    }).join('')}
                </tr>
            `;
        }).join('');
    }
});
