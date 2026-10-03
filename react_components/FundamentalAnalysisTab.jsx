import React, { useMemo } from 'react';
import { 
  TrendingUp, TrendingDown, ShieldCheck, AlertTriangle, 
  PieChart, Landmark, Users, Wallet, ArrowUpRight, ArrowDownRight 
} from 'lucide-react';
import { 
  LineChart, Line, ResponsiveContainer, 
  PieChart as RechartsPieChart, Pie, Cell 
} from 'recharts';

export default function FundamentalAnalysisTab({ data }) {
  // --- UTILS ---
  const formatCurrency = (val) => new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 }).format(val);
  const formatPercent = (val) => `${(val).toFixed(2)}%`;
  const formatRatio = (val) => `${(val).toFixed(2)}x`;
  
  const computeCAGR = (start, end, periods) => {
    if (start <= 0 || end <= 0 || periods <= 0) return 0;
    return (Math.pow(end / start, 1 / periods) - 1) * 100;
  };

  const getMedian = (arr) => {
    const sorted = [...arr].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  };

  // --- SECTION 1 COMPUTATIONS ---
  const hist = data.historical;
  const curr = data.currentValuation;
  
  const salesCAGR5 = computeCAGR(hist.sales[0], hist.sales[hist.sales.length - 1], hist.sales.length - 1);
  const salesCAGR3 = computeCAGR(hist.sales[hist.sales.length - 3], hist.sales[hist.sales.length - 1], 2);
  const salesConsistent = salesCAGR5 > 5 && salesCAGR3 > 0;
  
  const patCAGR5 = computeCAGR(hist.netProfit[0], hist.netProfit[hist.netProfit.length - 1], hist.netProfit.length - 1);
  const patConsistent = patCAGR5 > 5;
  
  const assetsGrowth = computeCAGR(hist.totalAssets[0], hist.totalAssets[hist.totalAssets.length - 1], hist.totalAssets.length - 1);
  const liabGrowth = computeCAGR(hist.totalLiabilities[0], hist.totalLiabilities[hist.totalLiabilities.length - 1], hist.totalLiabilities.length - 1);
  const deleveraging = liabGrowth < assetsGrowth;
  const netDebtDelta = hist.totalLiabilities[hist.totalLiabilities.length - 1] - hist.totalLiabilities[0];

  const recentCFO = hist.cfo[hist.cfo.length - 1];
  const recentPAT = hist.netProfit[hist.netProfit.length - 1];
  const cfoPatRatio = recentPAT > 0 ? recentCFO / recentPAT : 0;
  const cashQuality = cfoPatRatio > 1.0;

  // Chart data formatting
  const sparkData = (key) => hist.years.map((y, i) => ({ year: y, value: hist[key][i] }));

  // --- SHAREHOLDING COMPUTATIONS ---
  const sh = data.shareholding;
  const pieData = [
    { name: 'Promoters', value: sh.promoters, color: '#a855f7' },
    { name: 'FIIs', value: sh.fii, color: '#06b6d4' },
    { name: 'DIIs', value: sh.dii, color: '#10b981' },
    { name: 'Retail', value: sh.retail, color: '#f59e0b' },
  ];
  const totalInst = sh.fii + sh.dii;

  // --- RENDERING HELPERS ---
  const Badge = ({ condition, goodText, badText, goodIcon: GoodIcon, badIcon: BadIcon }) => (
    <div className={`flex items-center gap-1.5 text-xs font-medium px-2.5 py-1 rounded-full w-fit ${condition ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'bg-amber-500/10 text-amber-400 border border-amber-500/20'}`}>
      {condition ? <GoodIcon className="w-3.5 h-3.5" /> : <BadIcon className="w-3.5 h-3.5" />}
      {condition ? goodText : badText}
    </div>
  );

  const Sparkline = ({ dataKey }) => (
    <div className="h-12 w-full mt-4 opacity-70 pointer-events-none">
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={sparkData(dataKey)}>
          <Line type="monotone" dataKey="value" stroke="#06b6d4" strokeWidth={2} dot={false} isAnimationActive={false} />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );

  return (
    <div className="p-6 bg-[#0a0a0a] min-h-screen text-gray-200 font-sans space-y-8">
      
      {/* SECTION 1: EXECUTIVE TRAJECTORY */}
      <section>
        <h2 className="text-xl font-semibold mb-4 text-gray-100 flex items-center gap-2">
          <ShieldCheck className="w-5 h-5 text-cyan-500" /> Executive Health Verdicts
        </h2>
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
          
          {/* Sales Card */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors">
            <h3 className="text-sm text-gray-400 font-medium mb-3">Sales Trajectory</h3>
            <Badge condition={salesConsistent} goodText="Consistent Growth" badText="Volatile / Contracting" goodIcon={TrendingUp} badIcon={TrendingDown} />
            <div className="mt-4 flex gap-4 text-sm">
              <div>
                <span className="text-gray-500 block text-xs">3Y CAGR</span>
                <span className="font-mono text-gray-200">{formatPercent(salesCAGR3)}</span>
              </div>
              <div>
                <span className="text-gray-500 block text-xs">5Y CAGR</span>
                <span className="font-mono text-gray-200">{formatPercent(salesCAGR5)}</span>
              </div>
            </div>
            <p className="text-xs text-gray-500 mt-4 leading-relaxed">
              Revenue has {salesConsistent ? 'demonstrated reliable expansion' : 'shown historical volatility'}, growing by {formatPercent(salesCAGR5)} annualized over 5 years.
            </p>
          </div>

          {/* Profitability Card */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors">
            <h3 className="text-sm text-gray-400 font-medium mb-3">Profitability Trajectory</h3>
            <Badge condition={patConsistent} goodText="Compounding Earnings" badText="Margin Compression" goodIcon={TrendingUp} badIcon={AlertTriangle} />
            <div className="mt-4 flex gap-4 text-sm">
              <div>
                <span className="text-gray-500 block text-xs">5Y PAT CAGR</span>
                <span className="font-mono text-gray-200">{formatPercent(patCAGR5)}</span>
              </div>
            </div>
            <p className="text-xs text-gray-500 mt-4 leading-relaxed">
              Net profitability is {patConsistent ? 'expanding efficiently' : 'under pressure'}, reflecting current operating leverage dynamics.
            </p>
          </div>

          {/* Balance Sheet Card */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors">
            <h3 className="text-sm text-gray-400 font-medium mb-3">Balance Sheet Quality</h3>
            <Badge condition={deleveraging} goodText="Strong Deleveraging" badText="Debt Expansion" goodIcon={ShieldCheck} badIcon={TrendingDown} />
            <div className="mt-4 flex gap-4 text-sm">
              <div>
                <span className="text-gray-500 block text-xs">Liabilities Delta (5Y)</span>
                <span className={`font-mono ${netDebtDelta <= 0 ? 'text-emerald-400' : 'text-amber-400'}`}>
                  {netDebtDelta > 0 ? '+' : ''}{formatCurrency(netDebtDelta)}
                </span>
              </div>
            </div>
            <p className="text-xs text-gray-500 mt-4 leading-relaxed">
              Total Assets expanded by {formatPercent(assetsGrowth)} while Liabilities {deleveraging ? 'reduced/grew slower' : 'expanded rapidly'} by {formatPercent(liabGrowth)}.
            </p>
          </div>

          {/* Cash Flow Card */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors">
            <h3 className="text-sm text-gray-400 font-medium mb-3">Cash Flow Quality</h3>
            <Badge condition={cashQuality} goodText="High Cash Quality" badText="Working Capital Drag" goodIcon={Wallet} badIcon={AlertTriangle} />
            <div className="mt-4 flex gap-4 text-sm">
              <div>
                <span className="text-gray-500 block text-xs">CFO / PAT</span>
                <span className="font-mono text-gray-200">{formatRatio(cfoPatRatio)}</span>
              </div>
            </div>
            <p className="text-xs text-gray-500 mt-4 leading-relaxed">
              Operating cash flow is {cashQuality ? 'robust and fully backing earnings' : 'lagging behind reported net income'}, representing {formatPercent(cfoPatRatio * 100)} of PAT.
            </p>
          </div>
          
        </div>
      </section>

      {/* SECTION 2: METRICS GRID */}
      <section>
        <h2 className="text-xl font-semibold mb-4 text-gray-100 flex items-center gap-2">
          <Landmark className="w-5 h-5 text-cyan-500" /> Core Valuation & Metrics
        </h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
          
          {/* P/E */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors flex flex-col justify-between">
            <div>
              <div className="flex justify-between items-start">
                <h3 className="text-sm text-gray-400 font-medium">P/E Ratio</h3>
                <span className="text-[10px] font-semibold tracking-wider text-gray-500 uppercase bg-gray-800/50 px-2 py-0.5 rounded">Valuation</span>
              </div>
              <div className="mt-2 flex items-end gap-3">
                <span className="text-3xl font-bold font-mono text-gray-100">{formatRatio(curr.pe)}</span>
                <span className={`text-sm mb-1 font-mono ${curr.pe < curr.sectorPE ? 'text-emerald-400' : 'text-red-400'}`}>
                  vs {formatRatio(curr.sectorPE)} Ind
                </span>
              </div>
            </div>
            <Sparkline dataKey="pe" />
          </div>

          {/* P/B */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors flex flex-col justify-between">
            <div>
              <div className="flex justify-between items-start">
                <h3 className="text-sm text-gray-400 font-medium">P/B Ratio</h3>
                <span className="text-[10px] font-semibold tracking-wider text-gray-500 uppercase bg-gray-800/50 px-2 py-0.5 rounded">Valuation</span>
              </div>
              <div className="mt-2 flex items-end gap-3">
                <span className="text-3xl font-bold font-mono text-gray-100">{formatRatio(curr.pb)}</span>
                <span className="text-sm mb-1 text-gray-500 font-mono">Med {formatRatio(getMedian(hist.pb))}</span>
              </div>
            </div>
            <Sparkline dataKey="pb" />
          </div>

          {/* D/E */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors flex flex-col justify-between">
            <div>
              <div className="flex justify-between items-start">
                <h3 className="text-sm text-gray-400 font-medium">Debt to Equity</h3>
                <span className="text-[10px] font-semibold tracking-wider text-gray-500 uppercase bg-gray-800/50 px-2 py-0.5 rounded">Solvency</span>
              </div>
              <div className="mt-2 flex items-center gap-3">
                <span className="text-3xl font-bold font-mono text-gray-100">{formatRatio(curr.debtToEquity)}</span>
                {curr.debtToEquity < 0.5 && <span className="text-[10px] bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 px-1.5 py-0.5 rounded-full">Low Leverage</span>}
                {curr.debtToEquity > 1.0 && <span className="text-[10px] bg-red-500/10 text-red-400 border border-red-500/20 px-1.5 py-0.5 rounded-full">High Leverage</span>}
              </div>
            </div>
            <Sparkline dataKey="debtToEquity" />
          </div>

          {/* Div Yield */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors flex flex-col justify-between">
            <div>
              <div className="flex justify-between items-start">
                <h3 className="text-sm text-gray-400 font-medium">Div Yield & Payout</h3>
                <span className="text-[10px] font-semibold tracking-wider text-gray-500 uppercase bg-gray-800/50 px-2 py-0.5 rounded">Returns</span>
              </div>
              <div className="mt-2 flex items-end gap-3">
                <span className="text-3xl font-bold font-mono text-gray-100">{formatPercent(curr.dividendYield)}</span>
                <span className="text-sm mb-1 text-gray-500 font-mono">PO: {formatPercent(curr.dividendPayout)}</span>
              </div>
            </div>
            <Sparkline dataKey="dividendYield" />
          </div>

          {/* ROE & ROCE */}
          <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-5 hover:border-cyan-500/40 transition-colors flex flex-col justify-between sm:col-span-2 lg:col-span-1 xl:col-span-2">
            <div>
              <div className="flex justify-between items-start">
                <h3 className="text-sm text-gray-400 font-medium">Return on Capital (ROE & ROCE)</h3>
                <span className="text-[10px] font-semibold tracking-wider text-gray-500 uppercase bg-gray-800/50 px-2 py-0.5 rounded">Profitability</span>
              </div>
              <div className="mt-3 flex gap-8">
                <div>
                  <span className="block text-xs text-gray-500 mb-1">ROE</span>
                  <span className="text-2xl font-bold font-mono text-gray-100">{formatPercent(curr.roe)}</span>
                </div>
                <div>
                  <span className="block text-xs text-gray-500 mb-1">ROCE</span>
                  <span className="text-2xl font-bold font-mono text-gray-100">{formatPercent(curr.roce)}</span>
                </div>
              </div>
            </div>
          </div>

        </div>
      </section>

      {/* SECTION 3: SHAREHOLDING PATTERN */}
      <section>
        <h2 className="text-xl font-semibold mb-4 text-gray-100 flex items-center gap-2">
          <PieChart className="w-5 h-5 text-cyan-500" /> Shareholding Pattern
        </h2>
        <div className="bg-[#141414] border border-gray-800/80 rounded-xl p-6 flex flex-col md:flex-row gap-8 items-center hover:border-cyan-500/40 transition-colors">
          
          <div className="w-full md:w-1/3 h-64 relative flex items-center justify-center">
            <ResponsiveContainer width="100%" height="100%">
              <RechartsPieChart>
                <Pie data={pieData} cx="50%" cy="50%" innerRadius={70} outerRadius={100} paddingAngle={2} dataKey="value" stroke="none">
                  {pieData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Pie>
              </RechartsPieChart>
            </ResponsiveContainer>
            <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
              <span className="text-xs text-gray-500">Inst. Backing</span>
              <span className="text-2xl font-bold font-mono text-gray-200">{formatPercent(totalInst)}</span>
            </div>
          </div>

          <div className="w-full md:w-2/3 flex flex-col gap-5">
            {pieData.map((item, idx) => {
              const qoqKey = item.name.toLowerCase().replace('s', ''); 
              const qoq = sh.qoqChange[qoqKey] ?? sh.qoqChange[item.name.toLowerCase()] ?? 0;
              const isPositive = qoq > 0;
              const isNegative = qoq < 0;

              return (
                <div key={idx} className="w-full">
                  <div className="flex justify-between items-end mb-2">
                    <div className="flex items-center gap-2">
                      <div className="w-3 h-3 rounded-sm" style={{ backgroundColor: item.color }} />
                      <span className="text-sm font-medium text-gray-300">{item.name}</span>
                      {item.name === 'Promoters' && sh.pledgedPromoterShares > 0 && (
                        <span className="text-[10px] text-red-400 bg-red-400/10 border border-red-500/20 px-1.5 py-0.5 rounded ml-2">Pledged: {formatPercent(sh.pledgedPromoterShares)}</span>
                      )}
                    </div>
                    <div className="flex items-center gap-4">
                      <span className="text-sm font-bold font-mono text-gray-200">{formatPercent(item.value)}</span>
                      <span className={`text-xs font-mono w-20 text-right flex items-center justify-end ${isPositive ? 'text-emerald-400' : isNegative ? 'text-red-400' : 'text-gray-600'}`}>
                        {isPositive ? <ArrowUpRight className="w-3.5 h-3.5 mr-0.5"/> : (isNegative ? <ArrowDownRight className="w-3.5 h-3.5 mr-0.5"/> : null)}
                        {qoq === 0 ? 'Unchanged' : `${qoq > 0 ? '+' : ''}${qoq}%`}
                      </span>
                    </div>
                  </div>
                  <div className="w-full bg-gray-800/50 rounded-full h-1.5 overflow-hidden">
                    <div className="h-full rounded-full transition-all duration-1000 ease-out" style={{ width: `${item.value}%`, backgroundColor: item.color }} />
                  </div>
                </div>
              );
            })}
          </div>

        </div>
      </section>

    </div>
  );
}
