import React from 'react';
import { FileSpreadsheet, TrendingUp, Check, AlertCircle } from 'lucide-react';

export interface SpreadsheetProps {
  scenario: 'formula' | 'highlight' | 'chart' | 'rollback';
  step: number; // 0: initial, 1: step1, 2: step2, 3: step3, 4: done
}

export const OfficeSpreadsheet: React.FC<SpreadsheetProps> = ({ scenario, step }) => {
  // Scenario 1: Formula and data completion
  // Step 0: Raw initial table (8月 and 同比 are empty)
  // Step 1: Scanning range (A4:C8 highlighted with scan border)
  // Step 2: D column (8月) filled in with green flash
  // Step 3: E column (同比) formulas calculated and filled in
  // Step 4: Finished

  // Scenario 2: Highlight anomaly
  // Step 0: Normal data table
  // Step 1: Scanning/Search
  // Step 2: Row 6 (CAC) highlighted in light yellow
  // Step 3: Row 8 (Refund rate) highlighted in light yellow

  // Scenario 3: Add chart
  // Step 0: No chart
  // Step 1: Range selected
  // Step 2: Chart appears
  // Step 3: Peak point highlighted

  // Scenario 4: Rollback
  // Step 0: Table with changes
  // Step 1: Reading snapshot
  // Step 2: Reverted to clean original

  const isScanning = (scenario === 'formula' || scenario === 'highlight' || scenario === 'chart') && step === 1;

  // Formula scenario flags
  const isDataFilled = (scenario === 'formula' && step >= 2) || (scenario !== 'formula' && scenario !== 'rollback') || (scenario === 'rollback' && step === 0);
  const isFormulaCalculated = (scenario === 'formula' && step >= 3) || (scenario !== 'formula' && scenario !== 'rollback') || (scenario === 'rollback' && step === 0);

  // Highlight scenario flags
  const isCacYellow = (scenario === 'highlight' && step >= 2);
  const isRefundYellow = (scenario === 'highlight' && step >= 3);

  // Chart scenario flags
  const isChartVisible = (scenario === 'chart' && step >= 2);
  const isChartPeak = (scenario === 'chart' && step >= 3);

  // Rollback scenario flags
  const isRolledBack = (scenario === 'rollback' && step >= 2);

  return (
    <div className="flex flex-col h-full bg-[#181820] rounded-xl overflow-hidden border border-white/[0.1] shadow-2xl font-sans text-xs select-none">
      {/* Excel Title Bar */}
      <div className="bg-[#107c41] px-3 py-1.5 flex items-center justify-between text-white border-b border-black/20">
        <div className="flex items-center gap-2">
          <FileSpreadsheet className="w-4 h-4 text-emerald-100" />
          <span className="font-semibold text-xs tracking-wide">2026年经营分析表.xlsx - WPS / Excel 本地打开中</span>
        </div>
        <div className="flex items-center gap-2 text-[11px] text-emerald-100/90 font-mono">
          <span className="w-2 h-2 rounded-full bg-emerald-300 animate-pulse" />
          <span>实时同步中</span>
        </div>
      </div>

      {/* Ribbon Bar Mockup */}
      <div className="bg-[#24242e] px-3 py-1 flex items-center gap-4 text-zinc-300 border-b border-white/[0.06] text-[11px]">
        <span className="text-white font-medium border-b-2 border-[#107c41] pb-0.5">开始</span>
        <span className="hover:text-white cursor-pointer">插入</span>
        <span className="hover:text-white cursor-pointer">公式</span>
        <span className="hover:text-white cursor-pointer">数据</span>
        <span className="hover:text-white cursor-pointer">审阅</span>
      </div>

      {/* Formula Bar */}
      <div className="bg-[#1f1f28] px-3 py-1.5 flex items-center gap-2 border-b border-white/[0.06] text-zinc-300">
        <div className="bg-[#14141c] px-2 py-0.5 rounded border border-white/[0.08] font-mono text-[10px] text-zinc-400 min-w-[44px] text-center">
          {scenario === 'formula' && step >= 3 ? 'E4' : isScanning ? 'A4:C8' : 'D4'}
        </div>
        <div className="text-zinc-500 font-serif italic text-xs">fx</div>
        <div className="flex-1 bg-[#14141c] px-2 py-0.5 rounded border border-white/[0.08] font-mono text-[11px] text-zinc-200 truncate">
          {scenario === 'formula' && step >= 3
            ? '=(D4-C4)/C4'
            : isScanning
              ? '选区读取: =Sheet1!$A$4:$C$8'
              : '=SUM(D4:D8)'}
        </div>
      </div>

      {/* Main Table Grid */}
      <div className="flex-1 p-3 overflow-auto flex flex-col justify-between">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-[11px] text-zinc-300">
            <thead>
              <tr className="bg-[#23232c] text-zinc-400 font-medium">
                <th className="border border-white/[0.08] px-2 py-1 w-8 text-center bg-[#282834] text-[10px]">#</th>
                <th className="border border-white/[0.08] px-3 py-1.5 text-left">A (指标)</th>
                <th className="border border-white/[0.08] px-3 py-1.5 text-right">B (6月)</th>
                <th className="border border-white/[0.08] px-3 py-1.5 text-right">C (7月)</th>
                <th className={`border border-white/[0.08] px-3 py-1.5 text-right transition-colors ${
                  scenario === 'formula' && step === 2 ? 'bg-blue-500/20 text-blue-300 font-bold' : ''
                }`}>
                  D (8月实际)
                </th>
                <th className={`border border-white/[0.08] px-3 py-1.5 text-right transition-colors ${
                  scenario === 'formula' && step >= 3 ? 'bg-emerald-500/20 text-emerald-300 font-bold' : ''
                }`}>
                  E (同比增长)
                </th>
              </tr>
            </thead>
            <tbody>
              {/* Row 4: 销售总额 */}
              <tr className={`transition-all duration-300 ${isScanning ? 'ring-1 ring-blue-400/50 bg-blue-500/[0.04]' : 'hover:bg-white/[0.02]'}`}>
                <td className="border border-white/[0.08] px-2 py-1.5 text-center text-zinc-500 bg-[#202028] text-[10px]">4</td>
                <td className="border border-white/[0.08] px-3 py-1.5 font-medium text-zinc-200">销售总额 (万元)</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">540</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">560</td>
                <td className={`border border-white/[0.08] px-3 py-1.5 text-right font-mono font-bold transition-all duration-500 ${
                  isRolledBack ? 'text-zinc-600' : isDataFilled ? 'bg-emerald-500/15 text-emerald-300' : 'text-zinc-600'
                }`}>
                  {isRolledBack ? '-' : isDataFilled ? '612' : '-'}
                </td>
                <td className={`border border-white/[0.08] px-3 py-1.5 text-right font-mono transition-all duration-500 ${
                  isRolledBack ? 'text-zinc-600' : isFormulaCalculated ? 'bg-blue-500/15 text-blue-300 font-bold' : 'text-zinc-600'
                }`}>
                  {isRolledBack ? '-' : isFormulaCalculated ? '+9.3%' : '-'}
                </td>
              </tr>

              {/* Row 5: 转化率 */}
              <tr className={`transition-all duration-300 ${isScanning ? 'ring-1 ring-blue-400/50 bg-blue-500/[0.04]' : 'hover:bg-white/[0.02]'}`}>
                <td className="border border-white/[0.08] px-2 py-1.5 text-center text-zinc-500 bg-[#202028] text-[10px]">5</td>
                <td className="border border-white/[0.08] px-3 py-1.5 font-medium text-zinc-200">订单转化率</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">3.4%</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">3.5%</td>
                <td className={`border border-white/[0.08] px-3 py-1.5 text-right font-mono font-bold transition-all duration-500 ${
                  isRolledBack ? 'text-zinc-600' : isDataFilled ? 'bg-emerald-500/15 text-emerald-300' : 'text-zinc-600'
                }`}>
                  {isRolledBack ? '-' : isDataFilled ? '3.8%' : '-'}
                </td>
                <td className={`border border-white/[0.08] px-3 py-1.5 text-right font-mono transition-all duration-500 ${
                  isRolledBack ? 'text-zinc-600' : isFormulaCalculated ? 'bg-blue-500/15 text-blue-300 font-bold' : 'text-zinc-600'
                }`}>
                  {isRolledBack ? '-' : isFormulaCalculated ? '+0.3%' : '-'}
                </td>
              </tr>

              {/* Row 6: 获客成本 (CAC) - Can be highlighted */}
              <tr className={`transition-all duration-500 ${
                isCacYellow && !isRolledBack
                  ? 'bg-amber-400/25 border-amber-400 text-amber-200' 
                  : isScanning ? 'ring-1 ring-blue-400/50 bg-blue-500/[0.04]' : 'hover:bg-white/[0.02]'
              }`}>
                <td className="border border-white/[0.08] px-2 py-1.5 text-center text-zinc-500 bg-[#202028] text-[10px]">6</td>
                <td className="border border-white/[0.08] px-3 py-1.5 font-medium flex items-center justify-between">
                  <span>获客成本 CAC (元)</span>
                  {isCacYellow && !isRolledBack && (
                    <span className="text-[9px] px-1.5 py-0.2 rounded bg-amber-400/30 text-amber-300 font-bold">
                      ⚠️ 成本异常
                    </span>
                  )}
                </td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">45</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">48</td>
                <td className={`border border-white/[0.08] px-3 py-1.5 text-right font-mono font-bold ${
                  isCacYellow && !isRolledBack ? 'text-amber-300' : isDataFilled ? 'text-zinc-200' : 'text-zinc-600'
                }`}>
                  {isRolledBack ? '-' : isDataFilled ? '68' : '-'}
                </td>
                <td className={`border border-white/[0.08] px-3 py-1.5 text-right font-mono ${
                  isCacYellow && !isRolledBack ? 'text-amber-300 font-bold' : isFormulaCalculated ? 'text-zinc-300' : 'text-zinc-600'
                }`}>
                  {isRolledBack ? '-' : isFormulaCalculated ? '+41.7%' : '-'}
                </td>
              </tr>

              {/* Row 7: 净利润率 */}
              <tr className={`transition-all duration-300 ${isScanning ? 'ring-1 ring-blue-400/50 bg-blue-500/[0.04]' : 'hover:bg-white/[0.02]'}`}>
                <td className="border border-white/[0.08] px-2 py-1.5 text-center text-zinc-500 bg-[#202028] text-[10px]">7</td>
                <td className="border border-white/[0.08] px-3 py-1.5 font-medium text-zinc-200">净利润率</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">19.1%</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">19.4%</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono font-bold text-zinc-300">
                  {isRolledBack ? '-' : isDataFilled ? '14.8%' : '-'}
                </td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono text-zinc-400">
                  {isRolledBack ? '-' : isFormulaCalculated ? '-4.6%' : '-'}
                </td>
              </tr>

              {/* Row 8: 退款客诉率 - Can be highlighted */}
              <tr className={`transition-all duration-500 ${
                isRefundYellow && !isRolledBack
                  ? 'bg-amber-400/25 border-amber-400 text-amber-200' 
                  : isScanning ? 'ring-1 ring-blue-400/50 bg-blue-500/[0.04]' : 'hover:bg-white/[0.02]'
              }`}>
                <td className="border border-white/[0.08] px-2 py-1.5 text-center text-zinc-500 bg-[#202028] text-[10px]">8</td>
                <td className="border border-white/[0.08] px-3 py-1.5 font-medium flex items-center justify-between">
                  <span>退款客诉率</span>
                  {isRefundYellow && !isRolledBack && (
                    <span className="text-[9px] px-1.5 py-0.2 rounded bg-amber-400/30 text-amber-300 font-bold">
                      ⚠️ 客诉偏高
                    </span>
                  )}
                </td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">1.2%</td>
                <td className="border border-white/[0.08] px-3 py-1.5 text-right font-mono">1.3%</td>
                <td className={`border border-white/[0.08] px-3 py-1.5 text-right font-mono font-bold ${
                  isRefundYellow && !isRolledBack ? 'text-amber-300' : isDataFilled ? 'text-zinc-200' : 'text-zinc-600'
                }`}>
                  {isRolledBack ? '-' : isDataFilled ? '2.5%' : '-'}
                </td>
                <td className={`border border-white/[0.08] px-3 py-1.5 text-right font-mono ${
                  isRefundYellow && !isRolledBack ? 'text-amber-300 font-bold' : isFormulaCalculated ? 'text-zinc-300' : 'text-zinc-600'
                }`}>
                  {isRolledBack ? '-' : isFormulaCalculated ? '+1.2%' : '-'}
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* Dynamic Chart Area (for Chart Scenario) */}
        {scenario === 'chart' && (
          <div className="mt-3 p-3 bg-[#1e1e28] rounded-lg border border-white/[0.08] transition-all">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2">
                <TrendingUp className="w-3.5 h-3.5 text-blue-400" />
                <span className="font-semibold text-zinc-200 text-[11px]">月度营收走势 (就地插入原生图表)</span>
              </div>
              <span className="text-[10px] text-zinc-400 font-mono">
                {isChartVisible ? '已绑定数据源: $B$4:$D$4' : '等待图表指令执行...'}
              </span>
            </div>

            {isChartVisible ? (
              <div className="h-20 w-full relative flex items-end">
                <svg className="w-full h-full overflow-visible" viewBox="0 0 360 60">
                  <line x1="0" y1="15" x2="360" y2="15" stroke="rgba(255,255,255,0.06)" strokeDasharray="3 3" />
                  <line x1="0" y1="40" x2="360" y2="40" stroke="rgba(255,255,255,0.06)" strokeDasharray="3 3" />

                  <path
                    d="M 40 45 L 130 38 L 220 30 L 310 10"
                    fill="none"
                    stroke="#3b82f6"
                    strokeWidth="2.5"
                    strokeLinecap="round"
                  />
                  <circle cx="40" cy="45" r="3.5" fill="#3b82f6" />
                  <circle cx="130" cy="38" r="3.5" fill="#3b82f6" />
                  <circle cx="220" cy="30" r="3.5" fill="#3b82f6" />

                  {isChartPeak ? (
                    <g>
                      <circle cx="310" cy="10" r="5" fill="#10b981" className="animate-ping" />
                      <circle cx="310" cy="10" r="4" fill="#10b981" />
                      <text x="310" y="25" textAnchor="middle" fill="#34d399" fontWeight="bold" fontSize="10">
                        8月 612万 (最高点 🚀)
                      </text>
                    </g>
                  ) : (
                    <circle cx="310" cy="10" r="3.5" fill="#3b82f6" />
                  )}

                  <text x="40" y="58" textAnchor="middle" fill="#71717a" fontSize="9">6月</text>
                  <text x="130" y="58" textAnchor="middle" fill="#71717a" fontSize="9">7月</text>
                  <text x="220" y="58" textAnchor="middle" fill="#71717a" fontSize="9">8月</text>
                  <text x="310" y="58" textAnchor="middle" fill="#34d399" fontSize="9" fontWeight="bold">9月预测</text>
                </svg>
              </div>
            ) : (
              <div className="h-16 flex items-center justify-center text-zinc-500 text-[11px] border border-dashed border-white/[0.06] rounded">
                尚未插入图表（点击执行后生成）
              </div>
            )}
          </div>
        )}

        {/* Snapshot Notification for Rollback Scenario */}
        {scenario === 'rollback' && isRolledBack && (
          <div className="mt-3 p-2.5 rounded bg-amber-500/10 border border-amber-500/30 flex items-center gap-2 text-amber-300 text-[11px]">
            <AlertCircle className="w-4 h-4 flex-shrink-0" />
            <span>已通过快照 #snap-8021 回退至初始状态，未保存任何变更，随时由人工接管。</span>
          </div>
        )}
      </div>

      {/* Sheet Tabs Bar */}
      <div className="bg-[#181822] px-3 py-1 flex items-center justify-between border-t border-white/[0.06] text-[10px] text-zinc-400">
        <div className="flex items-center gap-2">
          <span className="px-2 py-0.5 bg-[#22222e] text-white rounded font-medium border-t border-emerald-500">
            Sheet1 (经营月报)
          </span>
          <span className="hover:text-white cursor-pointer px-1">Sheet2 (明细)</span>
        </div>
        <span className="text-zinc-500">缩放: 100% ｜ 就绪</span>
      </div>
    </div>
  );
};
