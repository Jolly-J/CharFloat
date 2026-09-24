import React from 'react';
import { FileSpreadsheet, TrendingUp, Check, AlertCircle } from 'lucide-react';

export interface SpreadsheetProps {
  scenario: 'formula' | 'highlight' | 'chart' | 'rollback';
  step: number; // 0: initial, 1: step1, 2: step2, 3: step3, 4: done
}

export const OfficeSpreadsheet: React.FC<SpreadsheetProps> = ({ scenario, step }) => {
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
    <div className="flex flex-col h-full bg-white rounded-2xl overflow-hidden border border-slate-200 shadow-xl font-sans text-xs select-none">
      {/* Excel / WPS Green Title Bar */}
      <div className="bg-[#107c41] px-4 py-2 flex items-center justify-between text-white border-b border-black/10">
        <div className="flex items-center gap-2">
          <FileSpreadsheet className="w-4 h-4 text-emerald-100" />
          <span className="font-semibold text-xs tracking-wide">2026年经营分析表.xlsx - WPS / Excel 本地打开中</span>
        </div>
        <div className="flex items-center gap-2 text-[11px] text-emerald-100/90 font-mono">
          <span className="w-2 h-2 rounded-full bg-emerald-300 animate-pulse" />
          <span>本地实时直连</span>
        </div>
      </div>

      {/* Clean Ribbon Bar Mockup */}
      <div className="bg-[#f8fafc] px-4 py-1.5 flex items-center gap-5 text-slate-600 border-b border-slate-200 text-[11px] font-medium">
        <span className="text-emerald-800 font-bold border-b-2 border-[#107c41] pb-1">开始</span>
        <span className="hover:text-slate-900 cursor-pointer">插入</span>
        <span className="hover:text-slate-900 cursor-pointer">页面布局</span>
        <span className="hover:text-slate-900 cursor-pointer">公式</span>
        <span className="hover:text-slate-900 cursor-pointer">数据</span>
        <span className="hover:text-slate-900 cursor-pointer">审阅</span>
      </div>

      {/* Formula Bar */}
      <div className="bg-white px-3 py-1.5 flex items-center gap-2 border-b border-slate-200 text-slate-700">
        <div className="bg-slate-100 px-2 py-0.5 rounded border border-slate-200 font-mono text-[10px] text-slate-600 min-w-[44px] text-center font-semibold">
          {scenario === 'formula' && step >= 3 ? 'E4' : isScanning ? 'A4:C8' : 'D4'}
        </div>
        <div className="text-slate-400 font-serif italic text-xs font-bold">fx</div>
        <div className="flex-1 bg-slate-50 px-2 py-0.5 rounded border border-slate-200 font-mono text-[11px] text-slate-800 truncate">
          {scenario === 'formula' && step >= 3
            ? '=(D4-C4)/C4'
            : isScanning
              ? '选区读取: =Sheet1!$A$4:$C$8'
              : '=SUM(D4:D8)'}
        </div>
      </div>

      {/* Main Clean Table Grid */}
      <div className="flex-1 p-3.5 overflow-auto flex flex-col justify-between bg-white">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-[11px] text-slate-700">
            <thead>
              <tr className="bg-[#f1f5f9] text-slate-600 font-semibold">
                <th className="border border-slate-200 px-2 py-1.5 w-8 text-center bg-slate-200/70 text-[10px]">#</th>
                <th className="border border-slate-200 px-3 py-2 text-left">A (核心指标)</th>
                <th className="border border-slate-200 px-3 py-2 text-right">B (6月)</th>
                <th className="border border-slate-200 px-3 py-2 text-right">C (7月)</th>
                <th className={`border border-slate-200 px-3 py-2 text-right transition-colors ${
                  scenario === 'formula' && step === 2 ? 'bg-blue-100 text-blue-800 font-bold' : ''
                }`}>
                  D (8月实际)
                </th>
                <th className={`border border-slate-200 px-3 py-2 text-right transition-colors ${
                  scenario === 'formula' && step >= 3 ? 'bg-emerald-100 text-emerald-800 font-bold' : ''
                }`}>
                  E (同比增长)
                </th>
              </tr>
            </thead>
            <tbody>
              {/* Row 4: 销售总额 */}
              <tr className={`transition-all duration-300 ${isScanning ? 'bg-blue-50/80 ring-1 ring-blue-400' : 'hover:bg-slate-50'}`}>
                <td className="border border-slate-200 px-2 py-2 text-center text-slate-400 bg-slate-100 text-[10px]">4</td>
                <td className="border border-slate-200 px-3 py-2 font-medium text-slate-900">销售总额 (万元)</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">540</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">560</td>
                <td className={`border border-slate-200 px-3 py-2 text-right font-mono font-bold transition-all duration-500 ${
                  isRolledBack ? 'text-slate-300' : isDataFilled ? 'bg-emerald-50 text-emerald-700' : 'text-slate-300'
                }`}>
                  {isRolledBack ? '-' : isDataFilled ? '612' : '-'}
                </td>
                <td className={`border border-slate-200 px-3 py-2 text-right font-mono transition-all duration-500 ${
                  isRolledBack ? 'text-slate-300' : isFormulaCalculated ? 'bg-blue-50 text-blue-700 font-bold' : 'text-slate-300'
                }`}>
                  {isRolledBack ? '-' : isFormulaCalculated ? '+9.3%' : '-'}
                </td>
              </tr>

              {/* Row 5: 订单转化率 */}
              <tr className={`transition-all duration-300 ${isScanning ? 'bg-blue-50/80 ring-1 ring-blue-400' : 'hover:bg-slate-50'}`}>
                <td className="border border-slate-200 px-2 py-2 text-center text-slate-400 bg-slate-100 text-[10px]">5</td>
                <td className="border border-slate-200 px-3 py-2 font-medium text-slate-900">订单转化率</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">3.4%</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">3.5%</td>
                <td className={`border border-slate-200 px-3 py-2 text-right font-mono font-bold transition-all duration-500 ${
                  isRolledBack ? 'text-slate-300' : isDataFilled ? 'bg-emerald-50 text-emerald-700' : 'text-slate-300'
                }`}>
                  {isRolledBack ? '-' : isDataFilled ? '3.8%' : '-'}
                </td>
                <td className={`border border-slate-200 px-3 py-2 text-right font-mono transition-all duration-500 ${
                  isRolledBack ? 'text-slate-300' : isFormulaCalculated ? 'bg-blue-50 text-blue-700 font-bold' : 'text-slate-300'
                }`}>
                  {isRolledBack ? '-' : isFormulaCalculated ? '+0.3%' : '-'}
                </td>
              </tr>

              {/* Row 6: 获客成本 (CAC) - Soft Yellow Highlight */}
              <tr className={`transition-all duration-500 ${
                isCacYellow && !isRolledBack
                  ? 'bg-amber-100/90 text-amber-950 font-medium' 
                  : isScanning ? 'bg-blue-50/80 ring-1 ring-blue-400' : 'hover:bg-slate-50'
              }`}>
                <td className="border border-slate-200 px-2 py-2 text-center text-slate-400 bg-slate-100 text-[10px]">6</td>
                <td className="border border-slate-200 px-3 py-2 font-medium flex items-center justify-between">
                  <span className="text-slate-900">获客成本 CAC (元)</span>
                  {isCacYellow && !isRolledBack && (
                    <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-200 text-amber-900 font-bold">
                      ⚠️ 成本超标
                    </span>
                  )}
                </td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">45</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">48</td>
                <td className={`border border-slate-200 px-3 py-2 text-right font-mono font-bold ${
                  isCacYellow && !isRolledBack ? 'text-amber-900 font-black' : isDataFilled ? 'text-slate-900' : 'text-slate-300'
                }`}>
                  {isRolledBack ? '-' : isDataFilled ? '68' : '-'}
                </td>
                <td className={`border border-slate-200 px-3 py-2 text-right font-mono ${
                  isCacYellow && !isRolledBack ? 'text-amber-900 font-bold' : isFormulaCalculated ? 'text-slate-700' : 'text-slate-300'
                }`}>
                  {isRolledBack ? '-' : isFormulaCalculated ? '+41.7%' : '-'}
                </td>
              </tr>

              {/* Row 7: 净利润率 */}
              <tr className={`transition-all duration-300 ${isScanning ? 'bg-blue-50/80 ring-1 ring-blue-400' : 'hover:bg-slate-50'}`}>
                <td className="border border-slate-200 px-2 py-2 text-center text-slate-400 bg-slate-100 text-[10px]">7</td>
                <td className="border border-slate-200 px-3 py-2 font-medium text-slate-900">净利润率</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">19.1%</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">19.4%</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono font-bold text-slate-800">
                  {isRolledBack ? '-' : isDataFilled ? '14.8%' : '-'}
                </td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">
                  {isRolledBack ? '-' : isFormulaCalculated ? '-4.6%' : '-'}
                </td>
              </tr>

              {/* Row 8: 退款客诉率 - Soft Yellow Highlight */}
              <tr className={`transition-all duration-500 ${
                isRefundYellow && !isRolledBack
                  ? 'bg-amber-100/90 text-amber-950 font-medium' 
                  : isScanning ? 'bg-blue-50/80 ring-1 ring-blue-400' : 'hover:bg-slate-50'
              }`}>
                <td className="border border-slate-200 px-2 py-2 text-center text-slate-400 bg-slate-100 text-[10px]">8</td>
                <td className="border border-slate-200 px-3 py-2 font-medium flex items-center justify-between">
                  <span className="text-slate-900">退款客诉率</span>
                  {isRefundYellow && !isRolledBack && (
                    <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-200 text-amber-900 font-bold">
                      ⚠️ 客诉偏高
                    </span>
                  )}
                </td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">1.2%</td>
                <td className="border border-slate-200 px-3 py-2 text-right font-mono text-slate-600">1.3%</td>
                <td className={`border border-slate-200 px-3 py-2 text-right font-mono font-bold ${
                  isRefundYellow && !isRolledBack ? 'text-amber-900 font-black' : isDataFilled ? 'text-slate-900' : 'text-slate-300'
                }`}>
                  {isRolledBack ? '-' : isDataFilled ? '2.5%' : '-'}
                </td>
                <td className={`border border-slate-200 px-3 py-2 text-right font-mono ${
                  isRefundYellow && !isRolledBack ? 'text-amber-900 font-bold' : isFormulaCalculated ? 'text-slate-700' : 'text-slate-300'
                }`}>
                  {isRolledBack ? '-' : isFormulaCalculated ? '+1.2%' : '-'}
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* Dynamic Native Chart Area (Clean White Modern Styling) */}
        {scenario === 'chart' && (
          <div className="mt-3.5 p-3.5 bg-slate-50 rounded-xl border border-slate-200 shadow-sm transition-all animate-fadeIn">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2">
                <TrendingUp className="w-4 h-4 text-blue-600" />
                <span className="font-bold text-slate-800 text-xs">月度营收趋势 (WPS/Excel 原生嵌入图表)</span>
              </div>
              <span className="text-[10px] text-slate-500 font-mono">
                {isChartVisible ? '动态绑定序列: $B$4:$D$4' : '图表生成中...'}
              </span>
            </div>

            {isChartVisible ? (
              <div className="h-24 w-full relative flex items-end pt-2">
                <svg className="w-full h-full overflow-visible" viewBox="0 0 360 65">
                  <line x1="0" y1="20" x2="360" y2="20" stroke="#e2e8f0" strokeDasharray="3 3" />
                  <line x1="0" y1="45" x2="360" y2="45" stroke="#e2e8f0" strokeDasharray="3 3" />

                  {/* Gradient fill under trend line */}
                  <defs>
                    <linearGradient id="chartGradient" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor="#2563eb" stopOpacity="0.2" />
                      <stop offset="100%" stopColor="#2563eb" stopOpacity="0.0" />
                    </linearGradient>
                  </defs>

                  <polygon
                    points="40,50 130,42 220,32 310,12 310,65 40,65"
                    fill="url(#chartGradient)"
                  />

                  <path
                    d="M 40 50 L 130 42 L 220 32 L 310 12"
                    fill="none"
                    stroke="#2563eb"
                    strokeWidth="3"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                  <circle cx="40" cy="50" r="4" fill="#2563eb" />
                  <circle cx="130" cy="42" r="4" fill="#2563eb" />
                  <circle cx="220" cy="32" r="4" fill="#2563eb" />
                  <circle cx="310" cy="12" r="5" fill="#1d4ed8" className="animate-ping" opacity="0.4" />
                  <circle cx="310" cy="12" r="5" fill="#1d4ed8" />

                  <text x="35" y="63" fontSize="9" fill="#64748b" fontFamily="sans-serif">6月 (540万)</text>
                  <text x="125" y="58" fontSize="9" fill="#64748b" fontFamily="sans-serif">7月 (560万)</text>
                  <text x="215" y="48" fontSize="9" fill="#64748b" fontFamily="sans-serif">8月预测</text>
                  <text x="285" y="10" fontSize="10" fill="#1d4ed8" fontWeight="bold" fontFamily="sans-serif">8月实际 612万 👑</text>
                </svg>
              </div>
            ) : (
              <div className="h-16 flex items-center justify-center text-slate-400 text-xs">
                正在向 WPS 图表引擎申请画板...
              </div>
            )}
          </div>
        )}
      </div>

      {/* Spreadsheet Bottom Status Bar */}
      <div className="bg-[#f8fafc] px-3.5 py-1.5 border-t border-slate-200 flex items-center justify-between text-[11px] text-slate-500">
        <div className="flex items-center gap-3">
          <span className="font-semibold text-emerald-700">就绪</span>
          <span className="hidden sm:inline">工作表 1 / 1</span>
        </div>
        <div className="flex items-center gap-2 font-mono text-[10px]">
          <span>求和: 1712</span>
          <span>平均值: 570.6</span>
        </div>
      </div>
    </div>
  );
};
