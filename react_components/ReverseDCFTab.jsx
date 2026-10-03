import React, { useMemo } from 'react';
import { 
  Target, Activity, TrendingUp, AlertTriangle, 
  ArrowRightLeft, ShieldCheck, Scale, CheckCircle2, TrendingDown 
} from 'lucide-react';

/**
 * ReverseDCFTab Component
 * 
 * An institutional Expectations Investing (Reverse DCF) terminal tab.
 * Inverts DCF valuation to reveal the cash flow growth expectations priced into the current stock price.
 * 
 * Features:
 *  1. Section 1: Expectations Verdict Banner with Implied vs Historical Benchmark Bars
 *  2. Section 2: Reverse Engine Parameters (WACC, Terminal Rate, Implied FCF Target)
 *  3. Section 3: 5x5 Implied Price & Margin of Safety Sensitivity Heatmap Matrix
 */
export default function ReverseDCFTab({ data }) {
  // --- FALLBACK / DEFAULT DATA CONTRACT ---
  const reverseData = useMemo(() => {
    if (!data) return null;
    return {
      ticker: data.ticker || 'COMPANY',
      currentPrice: data.currentPrice ?? 184,
      wacc: data.wacc ?? 11.73,
      terminalGrowth: data.terminalGrowth ?? 3.0,
      historicalFCFCagr: data.historicalFCFCagr ?? 8.5,
      impliedFCFCagr: data.impliedFCFCagr ?? 14.2,
      sectorAverageCagr: data.sectorAverageCagr ?? 10.1,
      sensitivityMatrix: data.sensitivityMatrix ?? {
        rowsWACC: [9.73, 10.73, 11.73, 12.73, 13.73],
        colsGrowth: [6.2, 10.2, 14.2, 18.2, 22.2],
        matrixPrices: [
          [220, 245, 275, 310, 350],
          [190, 215, 242, 270, 305],
          [165, 188, 184, 235, 260],
          [145, 162, 181, 205, 225],
          [125, 140, 158, 178, 195]
        ]
      }
    };
  }, [data]);

  if (!reverseData) {
    return (
      <div className="flex items-center justify-center h-64 bg-[#0a0a0a] text-gray-400 font-mono text-sm border border-gray-800 rounded-xl">
        <Activity className="w-5 h-5 mr-2 animate-spin text-cyan-400" />
        Computing reverse cash flow expectations...
      </div>
    );
  }

  const {
    ticker,
    currentPrice,
    wacc,
    terminalGrowth,
    historicalFCFCagr,
    impliedFCFCagr,
    sectorAverageCagr,
    sensitivityMatrix
  } = reverseData;

  // --- SECTION 1: EXPECTATIONS STATUS BADGE & THEMES ---
  // If Implied > Historical + 5%: "Priced for Perfection" (Crimson)
  // If Implied < Historical: "Asymmetric Upside" (Emerald)
  // Else: "Fairly Priced" (Amber)
  let statusBadge = 'Fairly Priced';
  let badgeColorClass = 'bg-amber-500/15 text-amber-400 border-amber-500/30';
  let bannerBorderClass = 'border-amber-500/30 bg-amber-500/5';
  let iconComponent = <AlertTriangle className="w-8 h-8 text-amber-400" />;
  let verdictSummary = 'The market is pricing in growth broadly consistent with recent historical operational compounding.';

  if (impliedFCFCagr > historicalFCFCagr + 5.0) {
    statusBadge = 'Priced for Perfection';
    badgeColorClass = 'bg-red-500/15 text-red-400 border-red-500/40';
    bannerBorderClass = 'border-red-500/40 bg-red-950/20 shadow-[0_0_25px_rgba(239,68,68,0.15)]';
    iconComponent = <TrendingUp className="w-8 h-8 text-red-500 animate-pulse" />;
    verdictSummary = 'Market pricing demands heroic acceleration beyond historical execution. Any deceleration risks severe multiple contraction.';
  } else if (impliedFCFCagr < historicalFCFCagr) {
    statusBadge = 'Asymmetric Upside';
    badgeColorClass = 'bg-emerald-500/15 text-emerald-400 border-emerald-500/40';
    bannerBorderClass = 'border-emerald-500/40 bg-emerald-950/20';
    iconComponent = <ShieldCheck className="w-8 h-8 text-emerald-400" />;
    verdictSummary = 'The market is pricing in significant pessimism or growth contraction, offering an institutional margin of safety.';
  }

  // Max value for progress bars
  const maxBenchmark = Math.max(impliedFCFCagr, historicalFCFCagr, sectorAverageCagr, 25.0) * 1.15;

  // Implied Terminal Multiple: 1 / (WACC - g)
  const impliedTerminalMultiple = wacc > terminalGrowth 
    ? (1.0 / ((wacc - terminalGrowth) / 100)).toFixed(1) 
    : '20.0';

  // Format currency
  const formatINR = (num) => `₹${Number(num).toLocaleString('en-IN', { maximumFractionDigits: 0 })}`;

  return (
    <div className="w-full bg-[#0a0a0a] text-gray-100 p-4 sm:p-6 space-y-6 font-sans antialiased">
      
      {/* ------------------------------------------------------------- */}
      {/* SECTION 1: EXPECTATIONS VERDICT (TOP BANNER) */}
      {/* ------------------------------------------------------------- */}
      <div className={`w-full rounded-2xl p-6 border transition-all duration-300 ${bannerBorderClass} backdrop-blur-md`}>
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-6">
          
          {/* Banner Left: Core Verdict Sentence */}
          <div className="flex items-start gap-4 max-w-3xl">
            <div className="p-3.5 bg-[#141414] border border-gray-800 rounded-xl shadow-inner flex-shrink-0">
              {iconComponent}
            </div>
            <div className="space-y-2">
              <div className="flex items-center gap-2.5">
                <span className="text-xs font-mono uppercase tracking-wider px-2 py-0.5 rounded bg-gray-900 border border-gray-700 text-gray-300">
                  Expectations Investing
                </span>
                <span className={`text-xs font-mono font-bold px-2.5 py-0.5 rounded-full border ${badgeColorClass}`}>
                  {statusBadge}
                </span>
              </div>

              {/* Core Reverse DCF Verdict Sentence */}
              <h2 className="text-lg sm:text-2xl font-bold tracking-tight text-white leading-snug">
                To justify the current share price of <span className="text-cyan-400 font-mono font-black">{formatINR(currentPrice)}</span>, the market is pricing in a <span className="text-amber-400 font-mono font-black">{impliedFCFCagr.toFixed(1)}%</span> annual Free Cash Flow CAGR over the next 5 years.
              </h2>
              
              <p className="text-xs sm:text-sm text-gray-300 leading-relaxed">
                {verdictSummary}
              </p>
            </div>
          </div>

          {/* Banner Right: Horizontal Comparison Bars */}
          <div className="w-full lg:w-80 bg-[#141414]/90 border border-gray-800/90 rounded-xl p-4 flex-shrink-0 space-y-3">
            <div className="text-xs font-mono text-gray-400 uppercase tracking-wide border-b border-gray-800 pb-1.5 flex justify-between items-center">
              <span>FCF Growth Benchmark</span>
              <span className="text-cyan-400 text-[10px]">5-Yr Compounding</span>
            </div>

            {/* Implied CAGR */}
            <div className="space-y-1">
              <div className="flex justify-between text-xs font-mono">
                <span className="text-gray-300 font-medium flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-cyan-400"></span>
                  Priced-in (Implied)
                </span>
                <span className="font-bold text-cyan-400">{impliedFCFCagr.toFixed(1)}%</span>
              </div>
              <div className="w-full h-2 bg-gray-900 rounded-full overflow-hidden">
                <div 
                  className="h-full bg-cyan-500 rounded-full transition-all duration-700" 
                  style={{ width: `${Math.min(100, Math.max(6, (impliedFCFCagr / maxBenchmark) * 100))}%` }}
                />
              </div>
            </div>

            {/* Historical CAGR */}
            <div className="space-y-1">
              <div className="flex justify-between text-xs font-mono">
                <span className="text-gray-400 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-emerald-400"></span>
                  Company Historical
                </span>
                <span className="font-semibold text-emerald-400">{historicalFCFCagr.toFixed(1)}%</span>
              </div>
              <div className="w-full h-2 bg-gray-900 rounded-full overflow-hidden">
                <div 
                  className="h-full bg-emerald-500 rounded-full transition-all duration-700" 
                  style={{ width: `${Math.min(100, Math.max(6, (historicalFCFCagr / maxBenchmark) * 100))}%` }}
                />
              </div>
            </div>

            {/* Sector Average CAGR */}
            <div className="space-y-1">
              <div className="flex justify-between text-xs font-mono">
                <span className="text-gray-400 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-purple-400"></span>
                  Sector Peer Median
                </span>
                <span className="font-semibold text-purple-400">{sectorAverageCagr.toFixed(1)}%</span>
              </div>
              <div className="w-full h-2 bg-gray-900 rounded-full overflow-hidden">
                <div 
                  className="h-full bg-purple-500 rounded-full transition-all duration-700" 
                  style={{ width: `${Math.min(100, Math.max(6, (sectorAverageCagr / maxBenchmark) * 100))}%` }}
                />
              </div>
            </div>

          </div>

        </div>
      </div>

      {/* ------------------------------------------------------------- */}
      {/* SECTION 2: THE REVERSE ENGINE PARAMETERS (CENTER) */}
      {/* ------------------------------------------------------------- */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        
        {/* Card 1: Cost of Capital (WACC) */}
        <div className="bg-[#141414] border border-gray-800/80 hover:border-cyan-500/30 rounded-xl p-5 transition-all duration-200 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between text-xs font-mono text-gray-400 mb-2">
              <span className="uppercase tracking-wider">Hurdle Rate</span>
              <Scale className="w-4 h-4 text-cyan-400" />
            </div>
            <h3 className="text-base font-bold text-white mb-1">
              Cost of Capital (WACC)
            </h3>
            <p className="text-xs text-gray-400 leading-relaxed">
              Discount rate reflecting market risk, cost of debt, and equity hurdle rate applied.
            </p>
          </div>
          <div className="mt-4 pt-3 border-t border-gray-800/80 flex items-baseline justify-between font-mono">
            <span className="text-2xl sm:text-3xl font-black text-cyan-400">
              {wacc.toFixed(2)}%
            </span>
            <span className="text-xs text-gray-400">
              Held Constant
            </span>
          </div>
        </div>

        {/* Card 2: Terminal Value Assumption */}
        <div className="bg-[#141414] border border-gray-800/80 hover:border-cyan-500/30 rounded-xl p-5 transition-all duration-200 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between text-xs font-mono text-gray-400 mb-2">
              <span className="uppercase tracking-wider">Long-Term Horizon</span>
              <Target className="w-4 h-4 text-purple-400" />
            </div>
            <h3 className="text-base font-bold text-white mb-1">
              Terminal Value Horizon
            </h3>
            <p className="text-xs text-gray-400 leading-relaxed">
              Perpetual growth (pegged to GDP) and implied exit cash multiple beyond Year 5.
            </p>
          </div>
          <div className="mt-4 pt-3 border-t border-gray-800/80 flex items-baseline justify-between font-mono">
            <div>
              <span className="text-2xl sm:text-3xl font-black text-purple-400">
                {terminalGrowth.toFixed(1)}%
              </span>
              <span className="text-xs text-gray-500 ml-1">Terminal g</span>
            </div>
            <div className="text-right">
              <span className="text-sm font-bold text-gray-300">
                {impliedTerminalMultiple}x
              </span>
              <div className="text-[10px] text-gray-500">Implied Multiple</div>
            </div>
          </div>
        </div>

        {/* Card 3: The Implied Growth Target (Solved Variable) */}
        <div className="bg-[#141414] border-2 border-cyan-500/50 hover:border-cyan-400 rounded-xl p-5 transition-all duration-200 shadow-[0_0_20px_rgba(6,182,212,0.12)] flex flex-col justify-between relative overflow-hidden">
          <div className="absolute -right-8 -top-8 w-24 h-24 bg-cyan-500/10 rounded-full blur-xl pointer-events-none" />
          <div>
            <div className="flex items-center justify-between text-xs font-mono text-cyan-400 mb-2">
              <span className="uppercase tracking-wider font-bold">Solved Expectation</span>
              <ArrowRightLeft className="w-4 h-4 text-cyan-400" />
            </div>
            <h3 className="text-base font-bold text-white mb-1">
              Implied 5-Yr FCF CAGR
            </h3>
            <p className="text-xs text-gray-300 leading-relaxed">
              Required annual compounding rate for DCF intrinsic value to match CMP of {formatINR(currentPrice)}.
            </p>
          </div>
          <div className="mt-4 pt-3 border-t border-cyan-900/40 flex items-baseline justify-between font-mono">
            <span className="text-3xl sm:text-4xl font-black text-amber-400">
              {impliedFCFCagr.toFixed(1)}%
            </span>
            <div className="text-right">
              <span className={`text-xs font-bold ${
                impliedFCFCagr > historicalFCFCagr ? 'text-red-400' : 'text-emerald-400'
              }`}>
                {impliedFCFCagr > historicalFCFCagr ? '+' : ''}
                {(impliedFCFCagr - historicalFCFCagr).toFixed(1)}%
              </span>
              <div className="text-[10px] text-gray-400">vs. Historical</div>
            </div>
          </div>
        </div>

      </div>

      {/* ------------------------------------------------------------- */}
      {/* SECTION 3: SENSITIVITY & MARGIN OF SAFETY MATRIX (BOTTOM) */}
      {/* ------------------------------------------------------------- */}
      <div className="bg-[#141414] border border-gray-800/80 rounded-2xl p-5 space-y-4">
        
        {/* Matrix Header */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-gray-800 pb-3">
          <div>
            <div className="flex items-center gap-2">
              <Activity className="w-5 h-5 text-cyan-400" />
              <h3 className="text-base font-bold text-white">
                Reverse Valuation Heatmap Matrix
              </h3>
            </div>
            <p className="text-xs text-gray-400 mt-0.5">
              Simulates implied share prices (₹) across varying 5-year FCF Growth (X-Axis) and WACC Discount Rates (Y-Axis).
            </p>
          </div>

          {/* Matrix Legend */}
          <div className="flex items-center gap-4 text-xs font-mono">
            <div className="flex items-center gap-1.5">
              <span className="w-3 h-3 rounded bg-emerald-500/40 border border-emerald-500/60" />
              <span className="text-gray-300">Price &gt; CMP (Undervalued)</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-3 h-3 rounded bg-cyan-500/40 border border-cyan-400" />
              <span className="text-cyan-300">Current Price (₹{currentPrice})</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-3 h-3 rounded bg-red-500/40 border border-red-500/60" />
              <span className="text-gray-300">Price &lt; CMP (Overvalued)</span>
            </div>
          </div>
        </div>

        {/* Matrix Heatmap Table */}
        <div className="overflow-x-auto">
          <table className="w-full text-center text-xs font-mono border-collapse min-w-[550px]">
            <thead>
              <tr>
                <th className="p-2.5 text-gray-500 font-medium text-left uppercase text-[11px] border-b border-gray-800">
                  WACC \ Growth
                </th>
                {sensitivityMatrix.colsGrowth.map((g, idx) => (
                  <th 
                    key={idx} 
                    className={`p-2.5 font-bold border-b border-gray-800 ${
                      Math.abs(g - impliedFCFCagr) < 0.5 ? 'text-amber-400 underline underline-offset-4 decoration-amber-400' : 'text-gray-300'
                    }`}
                  >
                    {g.toFixed(1)}%
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sensitivityMatrix.rowsWACC.map((rWacc, rowIdx) => {
                const isCurrentWacc = Math.abs(rWacc - wacc) < 0.2;
                return (
                  <tr key={rowIdx} className="hover:bg-white/[0.02] transition-colors">
                    {/* Y-Axis Header: WACC */}
                    <td className={`p-2.5 text-left font-bold border-r border-gray-800/80 ${
                      isCurrentWacc ? 'text-cyan-400 bg-cyan-950/20' : 'text-gray-400'
                    }`}>
                      {rWacc.toFixed(2)}%
                    </td>

                    {/* Matrix Cells */}
                    {sensitivityMatrix.matrixPrices[rowIdx]?.map((price, colIdx) => {
                      const isCurrentPriceCell = Math.abs(price - currentPrice) <= 3;
                      const upsidePct = ((price - currentPrice) / currentPrice) * 100;
                      const isUpside = price >= currentPrice;

                      // Color gradient intensity based on upside/downside
                      let cellStyle = 'bg-gray-900/40 text-gray-300 border-gray-800';
                      if (isCurrentPriceCell) {
                        cellStyle = 'bg-cyan-500/20 border-cyan-400 text-cyan-300 font-black ring-1 ring-cyan-400/60 shadow-[0_0_12px_rgba(6,182,212,0.3)]';
                      } else if (isUpside) {
                        if (upsidePct > 30) {
                          cellStyle = 'bg-emerald-500/25 border-emerald-500/40 text-emerald-300 font-bold';
                        } else if (upsidePct > 15) {
                          cellStyle = 'bg-emerald-500/15 border-emerald-500/30 text-emerald-400 font-medium';
                        } else {
                          cellStyle = 'bg-emerald-500/10 border-emerald-500/20 text-emerald-400';
                        }
                      } else {
                        if (upsidePct < -30) {
                          cellStyle = 'bg-red-500/25 border-red-500/40 text-red-300 font-bold';
                        } else if (upsidePct < -15) {
                          cellStyle = 'bg-red-500/15 border-red-500/30 text-red-400 font-medium';
                        } else {
                          cellStyle = 'bg-red-500/10 border-red-500/20 text-red-400';
                        }
                      }

                      return (
                        <td key={colIdx} className="p-1.5">
                          <div className={`p-2.5 rounded-lg border transition-all duration-150 flex flex-col items-center justify-center ${cellStyle}`}>
                            <span className="text-xs sm:text-sm">
                              {formatINR(price)}
                            </span>
                            <span className={`text-[10px] mt-0.5 ${
                              isUpside ? 'text-emerald-400' : 'text-red-400'
                            }`}>
                              {upsidePct >= 0 ? '+' : ''}{upsidePct.toFixed(0)}%
                            </span>
                          </div>
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Matrix Subtitle / Footnote */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between text-[11px] font-mono text-gray-500 pt-2 border-t border-gray-800">
          <div>
            Base FCF terminal valuation model: <span className="text-gray-400">Terminal g = {terminalGrowth.toFixed(1)}%</span>
          </div>
          <div>
            Center cell highlights implied pricing benchmark at CMP ({formatINR(currentPrice)})
          </div>
        </div>

      </div>

    </div>
  );
}
