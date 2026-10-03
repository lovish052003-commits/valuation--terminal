import React, { useMemo } from 'react';
import { 
  AlertOctagon, ShieldCheck, Search, Scale, FileWarning, 
  Activity, CheckCircle, XCircle, TrendingUp, TrendingDown,
  AlertTriangle, ArrowUpRight, ArrowDownRight, HelpCircle
} from 'lucide-react';
import { 
  ResponsiveContainer, ComposedChart, Bar, Line, 
  XAxis, YAxis, Tooltip, CartesianGrid, Legend, LineChart
} from 'recharts';

/**
 * ForensicAccountingTab Component
 * 
 * An institutional forensic financial analysis terminal tab.
 * Computes and visualizes:
 *  1. Master Beneish M-Score Verdict Banner with 5-Year Historical Sparkline
 *  2. 8-Variable Forensic Indices Grid with Threshold Flags
 *  3. Cash Flow Divergence (PAT vs CFO ComposedChart) & Auditor/Governance Checks
 */
export default function ForensicAccountingTab({ data }) {
  // --- FALLBACK / DEFAULT DATA STRUCTURE ---
  const forensicData = useMemo(() => {
    if (!data) return null;
    return {
      ticker: data.ticker || 'COMPANY',
      mScore: {
        current: data.mScore?.current ?? -2.45,
        historical5Yr: data.mScore?.historical5Yr ?? [-2.80, -2.65, -1.95, -2.10, -2.45],
        variables: {
          DSRI: data.mScore?.variables?.DSRI ?? 0.98,
          GMI: data.mScore?.variables?.GMI ?? 1.02,
          AQI: data.mScore?.variables?.AQI ?? 1.01,
          SGI: data.mScore?.variables?.SGI ?? 1.15,
          DEPI: data.mScore?.variables?.DEPI ?? 0.95,
          SGAI: data.mScore?.variables?.SGAI ?? 1.01,
          LVGI: data.mScore?.variables?.LVGI ?? 0.99,
          TATA: data.mScore?.variables?.TATA ?? 0.015,
        }
      },
      divergence: {
        years: data.divergence?.years ?? [2021, 2022, 2023, 2024, 2025],
        pat: data.divergence?.pat ?? [7490, 41749, 8075, -4910, 6800],
        cfo: data.divergence?.cfo ?? [44327, 44381, 29215, 20268, 24500]
      },
      governanceFlags: {
        auditorResignation: data.governanceFlags?.auditorResignation ?? false,
        highRPT: data.governanceFlags?.highRPT ?? false,
        delayedFilings: data.governanceFlags?.delayedFilings ?? false,
        regulatoryActions: data.governanceFlags?.regulatoryActions ?? false,
        contingentLiabilitiesHigh: data.governanceFlags?.contingentLiabilitiesHigh ?? false,
        qualifiedOpinion: data.governanceFlags?.qualifiedOpinion ?? false
      }
    };
  }, [data]);

  if (!forensicData) {
    return (
      <div className="flex items-center justify-center h-64 bg-[#0a0a0a] text-gray-400 font-mono text-sm border border-gray-800 rounded-xl">
        <Activity className="w-5 h-5 mr-2 animate-spin text-cyan-400" />
        Awaiting financial statements for forensic parsing...
      </div>
    );
  }

  const { mScore, divergence, governanceFlags } = forensicData;
  const currentM = mScore.current;

  // --- SECTION 1: MASTER VERDICT THRESHOLDS ---
  // M < -2.22: SAFE (Low Risk)
  // -2.22 <= M <= -1.78: CAUTION (Moderate Risk)
  // M > -1.78: HIGH RISK (Probable Manipulation)
  let verdictStatus = 'SAFE';
  let verdictTitle = 'SAFE - Low Probability of Earnings Manipulation';
  let verdictDesc = 'Financial statements reflect conservative accounting principles and high earnings quality.';
  let verdictColor = '#10b981'; // Emerald
  let verdictBadgeBg = 'bg-emerald-500/10 border-emerald-500/30 text-emerald-400';
  let verdictIcon = <ShieldCheck className="w-9 h-9 text-emerald-400" />;

  if (currentM > -1.78) {
    verdictStatus = 'FLAG';
    verdictTitle = 'HIGH RISK - Probable Earnings Manipulation Detected';
    verdictDesc = 'Multiple indices exceed critical forensic barriers. Substantial probability of aggressive revenue recognition or non-cash accrual swelling.';
    verdictColor = '#ef4444'; // Crimson
    verdictBadgeBg = 'bg-red-500/10 border-red-500/40 text-red-400 shadow-[0_0_20px_rgba(239,68,68,0.2)]';
    verdictIcon = <AlertOctagon className="w-9 h-9 text-red-500 animate-pulse" />;
  } else if (currentM >= -2.22) {
    verdictStatus = 'WATCH';
    verdictTitle = 'CAUTION - Elevated Forensic Risk';
    verdictDesc = 'M-Score sits in the cautionary grey band. Isolated accounting anomalies detected in gross margins, accruals, or asset capitalization.';
    verdictColor = '#f59e0b'; // Amber
    verdictBadgeBg = 'bg-amber-500/10 border-amber-500/30 text-amber-400';
    verdictIcon = <AlertTriangle className="w-9 h-9 text-amber-400" />;
  }

  // Format 5-Yr Sparkline Data
  const sparklineData = mScore.historical5Yr.map((val, idx) => ({
    period: `Y-${mScore.historical5Yr.length - 1 - idx === 0 ? 'Current' : mScore.historical5Yr.length - 1 - idx}`,
    mScore: val,
    threshold: -1.78
  }));

  // --- SECTION 2: 8 BENEISH VARIABLES DEFINITION & THRESHOLD RULES ---
  const variablesConfig = [
    {
      key: 'DSRI',
      name: "Days' Sales in Receivables Index",
      short: 'Days Sales in Receivables',
      threshold: 1.03,
      thresholdDisplay: '> 1.03',
      operator: 'gt',
      val: mScore.variables.DSRI,
      redFlagDesc: 'Flags channel stuffing or loose credit terms to artificially inflate sales.'
    },
    {
      key: 'GMI',
      name: 'Gross Margin Index',
      short: 'Gross Margin Degradation',
      threshold: 1.01,
      thresholdDisplay: '> 1.01',
      operator: 'gt',
      val: mScore.variables.GMI,
      redFlagDesc: 'Flags deteriorating margins, creating acute management pressure to manipulate profits.'
    },
    {
      key: 'AQI',
      name: 'Asset Quality Index',
      short: 'Capitalized Expense Pressure',
      threshold: 1.04,
      thresholdDisplay: '> 1.04',
      operator: 'gt',
      val: mScore.variables.AQI,
      redFlagDesc: 'Flags capitalization of routine operating expenses as non-current intangible assets.'
    },
    {
      key: 'SGI',
      name: 'Sales Growth Index',
      short: 'Extreme Top-Line Growth Pressure',
      threshold: 1.60,
      thresholdDisplay: '> 1.60',
      operator: 'gt',
      val: mScore.variables.SGI,
      redFlagDesc: 'Flags unsustainable top-line deceleration or aggressive revenue pull-forward.'
    },
    {
      key: 'DEPI',
      name: 'Depreciation Index',
      short: 'Useful Asset Life Extension',
      threshold: 1.07,
      thresholdDisplay: '> 1.07',
      operator: 'gt',
      val: mScore.variables.DEPI,
      redFlagDesc: 'Flags sudden reduction in depreciation rates to boost reported net income.'
    },
    {
      key: 'SGAI',
      name: 'SG&A Expenses Index',
      short: 'Operating Efficiency Drag',
      threshold: 1.05,
      thresholdDisplay: '> 1.05',
      operator: 'gt',
      val: mScore.variables.SGAI,
      redFlagDesc: 'Flags administrative costs compounding faster than revenues (operational drag).'
    },
    {
      key: 'LVGI',
      name: 'Leverage Index',
      short: 'Balance Sheet Gearing Shift',
      threshold: 1.03,
      thresholdDisplay: '> 1.03',
      operator: 'gt',
      val: mScore.variables.LVGI,
      redFlagDesc: 'Flags rising debt-to-assets ratio and increasing debt covenants pressure.'
    },
    {
      key: 'TATA',
      name: 'Total Accruals to Total Assets',
      short: 'Non-Cash Earnings Distortion',
      threshold: 0.03,
      thresholdDisplay: '> 0.03',
      operator: 'gt',
      val: mScore.variables.TATA,
      redFlagDesc: 'Flags high non-cash accounting accruals with low cash backing.'
    }
  ];

  // Count flagged variables
  const flaggedCount = variablesConfig.filter(v => v.val > v.threshold).length;

  // --- SECTION 3: CASH FLOW DIVERGENCE DETECTION ---
  const divergenceChartData = divergence.years.map((year, i) => ({
    year: year.toString(),
    pat: divergence.pat[i] ?? 0,
    cfo: divergence.cfo[i] ?? 0,
  }));

  // Check divergence: Latest PAT rising while CFO falling/flat or CFO < PAT significantly
  const n = divergence.pat.length;
  let divergenceDetected = false;
  if (n >= 2) {
    const patRecent = divergence.pat[n - 1];
    const patPrev = divergence.pat[n - 2];
    const cfoRecent = divergence.cfo[n - 1];
    const cfoPrev = divergence.cfo[n - 2];
    
    // Condition 1: PAT rising YoY while CFO falling YoY
    if (patRecent > patPrev && cfoRecent <= cfoPrev) {
      divergenceDetected = true;
    }
    // Condition 2: High PAT with significantly lower or negative CFO
    if (patRecent > 0 && cfoRecent < 0.7 * patRecent) {
      divergenceDetected = true;
    }
  }

  // --- SECTION 3: GOVERNANCE & AUDITOR CHECKS LIST ---
  const governanceList = [
    {
      label: 'Frequent Auditor Resignations / Changes',
      description: 'Mid-term resignation or frequent switching of statutory audit firms.',
      flagged: !!governanceFlags.auditorResignation
    },
    {
      label: 'High Related Party Transactions (RPT)',
      description: 'Material procurement or loans routed through promoter-controlled affiliates.',
      flagged: !!governanceFlags.highRPT
    },
    {
      label: 'Delayed Financial Statement Filings',
      description: 'Repeated delays in publishing quarterly or annual audited results.',
      flagged: !!governanceFlags.delayedFilings
    },
    {
      label: 'Material Regulatory Actions / SEBI Inquiries',
      description: 'Active regulatory notices, forensic audits, or enforcement proceedings.',
      flagged: !!governanceFlags.regulatoryActions
    },
    {
      label: 'Contingent Liabilities > 10% of Net Worth',
      description: 'Off-balance sheet tax or contractual claims exceeding equity cushion.',
      flagged: !!governanceFlags.contingentLiabilitiesHigh
    },
    {
      label: 'Auditor Qualified Opinion / Matter of Emphasis',
      description: 'Auditor caveats on internal controls or going concern assumptions.',
      flagged: !!governanceFlags.qualifiedOpinion
    }
  ];

  return (
    <div className="w-full bg-[#0a0a0a] text-gray-100 p-4 sm:p-6 space-y-6 font-sans antialiased">
      
      {/* ------------------------------------------------------------- */}
      {/* SECTION 1: MASTER FORENSIC VERDICT (TOP BANNER) */}
      {/* ------------------------------------------------------------- */}
      <div className={`w-full rounded-2xl p-6 border transition-all duration-300 ${verdictBadgeBg} bg-opacity-40 backdrop-blur-md`}>
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-6">
          
          {/* Left Column: Verdict Badge & Description */}
          <div className="flex items-start gap-4 max-w-2xl">
            <div className="p-3 bg-[#141414] border border-gray-800 rounded-xl shadow-inner flex-shrink-0">
              {verdictIcon}
            </div>
            <div>
              <div className="flex items-center gap-2 mb-1">
                <span className="text-xs font-mono uppercase tracking-wider px-2 py-0.5 rounded bg-gray-900/80 border border-gray-700 text-gray-300">
                  Forensic Health Audit
                </span>
                <span className={`text-xs font-mono font-semibold px-2 py-0.5 rounded border ${
                  verdictStatus === 'SAFE' 
                    ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/40' 
                    : verdictStatus === 'WATCH' 
                    ? 'bg-amber-500/20 text-amber-400 border-amber-500/40' 
                    : 'bg-red-500/20 text-red-400 border-red-500/40'
                }`}>
                  {verdictStatus} STATUS
                </span>
              </div>
              <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-white mb-1">
                {verdictTitle}
              </h2>
              <p className="text-sm text-gray-300 leading-relaxed">
                {verdictDesc}
              </p>
              <div className="flex items-center gap-4 mt-3 text-xs text-gray-400 font-mono">
                <span>Critical Threshold: <strong className="text-red-400">&gt; -1.78 (Flag)</strong></span>
                <span>•</span>
                <span>Safe Threshold: <strong className="text-emerald-400">&lt; -2.22 (Clean)</strong></span>
                <span>•</span>
                <span>Breached Variables: <strong className={flaggedCount > 0 ? 'text-red-400' : 'text-emerald-400'}>{flaggedCount} of 8</strong></span>
              </div>
            </div>
          </div>

          {/* Right Column: Large M-Score Display & 5-Year Mini Sparkline */}
          <div className="flex items-center gap-6 bg-[#141414]/90 border border-gray-800/80 rounded-xl p-4 self-stretch lg:self-auto min-w-[280px] justify-between">
            <div>
              <div className="text-xs font-mono text-gray-400 uppercase">Beneish M-Score</div>
              <div className="text-3xl sm:text-4xl font-mono font-black tracking-tight text-white mt-0.5">
                {currentM.toFixed(2)}
              </div>
              <div className="text-[11px] font-mono mt-1 text-gray-400">
                {currentM > -1.78 ? (
                  <span className="text-red-400 flex items-center gap-1 font-semibold">
                    <ArrowUpRight className="w-3.5 h-3.5" /> High Manipulation Risk
                  </span>
                ) : currentM < -2.22 ? (
                  <span className="text-emerald-400 flex items-center gap-1 font-semibold">
                    <ArrowDownRight className="w-3.5 h-3.5" /> Low Manipulation Risk
                  </span>
                ) : (
                  <span className="text-amber-400 flex items-center gap-1 font-semibold">
                    <Activity className="w-3.5 h-3.5" /> Cautionary Zone
                  </span>
                )}
              </div>
            </div>

            {/* Historical 5-Year M-Score Sparkline */}
            <div className="w-32 h-16 flex flex-col justify-end">
              <div className="text-[10px] font-mono text-gray-500 text-right mb-1">5-Yr Trajectory</div>
              <div className="w-full h-12">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={sparklineData}>
                    <Line 
                      type="monotone" 
                      dataKey="mScore" 
                      stroke={verdictColor} 
                      strokeWidth={2.5} 
                      dot={{ r: 2.5, fill: verdictColor }} 
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>

          </div>
        </div>
      </div>

      {/* ------------------------------------------------------------- */}
      {/* SECTION 2: THE 8-VARIABLE FORENSIC GRID (CENTER) */}
      {/* ------------------------------------------------------------- */}
      <div className="space-y-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Scale className="w-5 h-5 text-cyan-400" />
            <h3 className="text-base sm:text-lg font-bold text-white tracking-wide">
              The 8-Variable Beneish Forensic Grid
            </h3>
          </div>
          <div className="text-xs font-mono text-gray-400 hidden sm:block">
            Formula: <span className="text-cyan-400">M = -4.84 + 0.920·DSRI + 0.528·GMI + 0.404·AQI + 0.892·SGI + 0.115·DEPI - 0.172·SGAI + 4.679·TATA - 0.327·LVGI</span>
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {variablesConfig.map((item) => {
            const isFlagged = item.val > item.threshold;
            return (
              <div 
                key={item.key}
                className={`bg-[#141414] rounded-xl p-4 border transition-all duration-200 flex flex-col justify-between ${
                  isFlagged 
                    ? 'border-red-500/50 bg-red-950/10 hover:border-red-500 hover:shadow-[0_0_15px_rgba(239,68,68,0.25)]' 
                    : 'border-gray-800/80 hover:border-gray-700'
                }`}
              >
                <div>
                  {/* Card Header: Variable Key + Status Badge */}
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-mono text-sm font-bold text-cyan-400 bg-cyan-950/40 px-2 py-0.5 rounded border border-cyan-800/50">
                      {item.key}
                    </span>
                    <span className={`text-[11px] font-mono font-semibold px-2 py-0.5 rounded-full flex items-center gap-1 border ${
                      isFlagged 
                        ? 'bg-red-500/15 text-red-400 border-red-500/30' 
                        : 'bg-emerald-500/15 text-emerald-400 border-emerald-500/30'
                    }`}>
                      {isFlagged ? (
                        <>
                          <AlertOctagon className="w-3 h-3" />
                          FLAGGED
                        </>
                      ) : (
                        <>
                          <CheckCircle className="w-3 h-3" />
                          NORMAL
                        </>
                      )}
                    </span>
                  </div>

                  {/* Variable Full Name */}
                  <h4 className="text-xs font-medium text-gray-200 line-clamp-1 mb-1" title={item.name}>
                    {item.name}
                  </h4>
                  <div className="text-[11px] text-gray-400 mb-3">
                    {item.short}
                  </div>
                </div>

                <div>
                  {/* Metrics Row: Current Value vs Threshold */}
                  <div className="flex items-baseline justify-between border-t border-gray-800/60 pt-2.5 mt-1 font-mono">
                    <div>
                      <div className="text-[10px] text-gray-500 uppercase">Current</div>
                      <div className={`text-lg font-bold ${isFlagged ? 'text-red-400' : 'text-gray-100'}`}>
                        {item.val.toFixed(3)}
                      </div>
                    </div>
                    <div className="text-right">
                      <div className="text-[10px] text-gray-500 uppercase">Threshold</div>
                      <div className="text-xs font-semibold text-gray-300">
                        {item.thresholdDisplay}
                      </div>
                    </div>
                  </div>

                  {/* Forensic Description */}
                  <div className={`text-[11px] mt-2.5 p-2 rounded-lg leading-snug border ${
                    isFlagged 
                      ? 'bg-red-900/20 border-red-800/40 text-red-300' 
                      : 'bg-gray-900/50 border-gray-800/40 text-gray-400'
                  }`}>
                    {item.redFlagDesc}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* ------------------------------------------------------------- */}
      {/* SECTION 3: CASH FLOW DIVERGENCE & AUDITOR FLAGS (BOTTOM) */}
      {/* ------------------------------------------------------------- */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        
        {/* Left Column: Cash Flow Divergence (7 Cols) */}
        <div className="lg:col-span-7 bg-[#141414] border border-gray-800/80 rounded-2xl p-5 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-gray-800/80 pb-3">
            <div>
              <div className="flex items-center gap-2">
                <Activity className="w-5 h-5 text-cyan-400" />
                <h3 className="text-base font-bold text-white">
                  Cash Flow Divergence (PAT vs CFO)
                </h3>
              </div>
              <p className="text-xs text-gray-400 mt-0.5">
                Compares reported Net Profit (accrual accounting) vs Operating Cash Flow (real cash).
              </p>
            </div>
            <div className="flex items-center gap-4 text-xs font-mono">
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded-sm bg-cyan-500"></span>
                <span className="text-gray-300">CFO (Bars)</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-0.5 bg-amber-400"></span>
                <span className="text-gray-300">PAT (Line)</span>
              </div>
            </div>
          </div>

          {/* Warning Banner if Divergence Detected */}
          {divergenceDetected && (
            <div className="flex items-center gap-3 p-3.5 bg-red-950/40 border border-red-500/40 rounded-xl text-red-300 animate-pulse">
              <FileWarning className="w-5 h-5 text-red-400 flex-shrink-0" />
              <div className="text-xs">
                <strong className="font-bold text-red-200 uppercase tracking-wide mr-1.5">
                  DIVERGENCE DETECTED:
                </strong>
                Reported net profits lack operating cash backing. Earnings may be propped up by non-cash receivables or delayed vendor disbursements.
              </div>
            </div>
          )}

          {/* Recharts ComposedChart: PAT as Line, CFO as Bars */}
          <div className="w-full h-64 sm:h-72 mt-2">
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={divergenceChartData} margin={{ top: 15, right: 15, left: -10, bottom: 5 }}>
                <CartesianGrid stroke="#222" strokeDasharray="3 3" vertical={false} />
                <XAxis 
                  dataKey="year" 
                  stroke="#666" 
                  fontSize={11} 
                  tickLine={false} 
                />
                <YAxis 
                  stroke="#666" 
                  fontSize={11} 
                  tickLine={false} 
                  tickFormatter={(val) => `₹${(val / 1000).toFixed(0)}k`}
                />
                <Tooltip 
                  contentStyle={{ 
                    backgroundColor: '#18181b', 
                    borderColor: '#3f3f46', 
                    borderRadius: '8px', 
                    fontSize: '12px',
                    color: '#fff' 
                  }}
                  formatter={(value, name) => [
                    `₹${value.toLocaleString('en-IN')} Cr`, 
                    name === 'pat' ? 'Net Profit (PAT)' : 'Operating Cash Flow (CFO)'
                  ]}
                />
                <Legend 
                  verticalAlign="top" 
                  height={30} 
                  formatter={(val) => val === 'pat' ? 'Net Profit (PAT)' : 'Operating Cash Flow (CFO)'}
                />
                <Bar 
                  dataKey="cfo" 
                  fill="#06b6d4" 
                  radius={[4, 4, 0, 0]} 
                  name="cfo"
                  maxBarSize={45} 
                />
                <Line 
                  type="monotone" 
                  dataKey="pat" 
                  stroke="#f59e0b" 
                  strokeWidth={2.5} 
                  dot={{ r: 4, fill: '#f59e0b' }} 
                  activeDot={{ r: 6 }} 
                  name="pat"
                />
              </ComposedChart>
            </ResponsiveContainer>
          </div>

        </div>

        {/* Right Column: Auditor & Governance Checks (5 Cols) */}
        <div className="lg:col-span-5 bg-[#141414] border border-gray-800/80 rounded-2xl p-5 space-y-4">
          <div className="border-b border-gray-800/80 pb-3">
            <div className="flex items-center gap-2">
              <Search className="w-5 h-5 text-purple-400" />
              <h3 className="text-base font-bold text-white">
                Auditor & Governance Checklist
              </h3>
            </div>
            <p className="text-xs text-gray-400 mt-0.5">
              Qualitative corporate governance integrity filters and auditor sanity checks.
            </p>
          </div>

          <div className="space-y-2.5">
            {governanceList.map((item, idx) => (
              <div 
                key={idx}
                className={`p-3 rounded-xl border flex items-start justify-between gap-3 transition-colors ${
                  item.flagged 
                    ? 'bg-red-950/20 border-red-500/40 text-red-200' 
                    : 'bg-gray-900/30 border-gray-800/60 text-gray-300'
                }`}
              >
                <div className="space-y-0.5">
                  <div className="text-xs font-semibold text-gray-200">
                    {item.label}
                  </div>
                  <div className="text-[11px] text-gray-400 leading-snug">
                    {item.description}
                  </div>
                </div>

                <div className="flex-shrink-0 mt-0.5">
                  {item.flagged ? (
                    <span className="flex items-center gap-1 font-mono text-[11px] text-red-400 bg-red-950/80 px-2 py-0.5 rounded border border-red-800">
                      <XCircle className="w-3.5 h-3.5 text-red-400" />
                      FLAGGED
                    </span>
                  ) : (
                    <span className="flex items-center gap-1 font-mono text-[11px] text-emerald-400 bg-emerald-950/80 px-2 py-0.5 rounded border border-emerald-800">
                      <CheckCircle className="w-3.5 h-3.5 text-emerald-400" />
                      CLEAN
                    </span>
                  )}
                </div>
              </div>
            ))}
          </div>

          <div className="p-3 bg-gray-900/60 border border-gray-800 rounded-xl text-[11px] text-gray-400 leading-relaxed font-mono">
            <span className="text-cyan-400 font-bold">Forensic Note:</span> Any simultaneous occurrence of high RPT, auditor turnover, and M-Score &gt; -1.78 warrants an immediate discount to intrinsic valuation multiples.
          </div>
        </div>

      </div>

    </div>
  );
}
