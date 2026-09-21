import React, { useState, useEffect, useRef } from 'react';
import { Sparkles, Terminal, CheckCircle2, ChevronDown, ArrowRight, CornerDownLeft, Zap, FileSpreadsheet } from 'lucide-react';
import { OfficeSpreadsheet } from './OfficeSpreadsheet';

type ScenarioType = 'formula' | 'highlight' | 'chart' | 'rollback';

interface ScenarioItem {
  id: ScenarioType;
  tabTitle: string;
  userPrompt: string;
  toolCall: string;
  toolActionDesc: string;
  agentReply: string;
}

const SCENARIOS: ScenarioItem[] = [
  {
    id: 'formula',
    tabTitle: '需求 ①：补数据并算公式',
    userPrompt: '“把 8 月数据补齐，并按去年的计算口径把同比增长率算出来。”',
    toolCall: 'excel_patch_cells',
    toolActionDesc: '写入 8 月各指标数据并继承同比增长公式',
    agentReply: '已识别历史口径。8 月营收 (612万) 及同比增长 (+9.3%) 已就地计算完成。',
  },
  {
    id: 'highlight',
    tabTitle: '需求 ②：找异常就地标黄',
    userPrompt: '“把成本超标的和未达标指标标出来，不要标红，改成浅黄色。”',
    toolCall: 'excel_format_cells',
    toolActionDesc: '批量将第 6 行与第 8 行背景设为浅黄警示',
    agentReply: '已定位获客成本 (68元) 及客诉率 (2.5%)，已将背景调整为浅黄色。',
  },
  {
    id: 'chart',
    tabTitle: '需求 ③：就地生成走势图',
    userPrompt: '“根据 6 到 8 月的营收走势画一张折线图，放在表格右下方。”',
    toolCall: 'excel_add_chart',
    toolActionDesc: '提取连续月度序列并绘制原生嵌入式折线图',
    agentReply: '折线走势图已在当前表格右侧生成，并自动标注了 8 月最高峰值点。',
  },
  {
    id: 'rollback',
    tabTitle: '需求 ④：一键撤回与人机交接',
    userPrompt: '“刚才那步撤回，这个我自己微调。”',
    toolCall: 'wps_rollback',
    toolActionDesc: '调取快照 #snap-8021，毫秒级还原单元格初态',
    agentReply: '已撤回最近一次改动并恢复初始状态，随时由您直接接管编辑。',
  },
];

export const LiveWorkspace: React.FC = () => {
  const containerRef = useRef<HTMLDivElement>(null);
  const [activeIdx, setActiveIdx] = useState<number>(0);
  const [animKey, setAnimKey] = useState<number>(0);

  // Scroll listener: detects scroll progress within container track and switches scenario
  useEffect(() => {
    const handleScroll = () => {
      if (!containerRef.current) return;
      const rect = containerRef.current.getBoundingClientRect();
      const viewportHeight = window.innerHeight;
      const totalScrollable = rect.height - viewportHeight;

      if (totalScrollable <= 0) return;

      // Current distance scrolled into the container
      const scrolled = -rect.top;
      const progress = Math.max(0, Math.min(1, scrolled / totalScrollable));

      // Map progress 0..1 to scenario index 0..3
      const newIdx = Math.min(
        SCENARIOS.length - 1,
        Math.floor(progress * SCENARIOS.length)
      );

      setActiveIdx((prev) => {
        if (prev !== newIdx) {
          setAnimKey((k) => k + 1);
          return newIdx;
        }
        return prev;
      });
    };

    window.addEventListener('scroll', handleScroll, { passive: true });
    return () => window.removeEventListener('scroll', handleScroll);
  }, []);

  // Clicking tab smoothly scrolls the page to that specific section
  const handleTabClick = (idx: number) => {
    if (!containerRef.current) return;
    const rect = containerRef.current.getBoundingClientRect();
    const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
    const totalScrollable = rect.height - window.innerHeight;
    const targetScroll = scrollTop + rect.top + (idx / SCENARIOS.length) * totalScrollable + 20;

    window.scrollTo({ top: targetScroll, behavior: 'smooth' });
    setActiveIdx(idx);
    setAnimKey((k) => k + 1);
  };

  const currentScenario = SCENARIOS[activeIdx];

  return (
    // Outer scroll track: 320vh ensures comfortable, natural scrolling distance
    <section ref={containerRef} className="relative h-[320vh] w-full">
      {/* Sticky Inner Frame: Stays fixed in viewport while user scrolls */}
      <div className="sticky top-6 z-20 max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Top Demands Switcher Bar (User's Exact Screenshot UI + Scroll Hijacking) */}
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4 bg-[#121216]/90 p-2.5 rounded-2xl border border-white/[0.08] backdrop-blur-md shadow-xl">
          <div className="flex items-center gap-2">
            <span className="text-xs font-semibold text-zinc-400 pl-2">
              各种常见修改需求：
            </span>
            <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
              {SCENARIOS.map((scenario, idx) => {
                const isSelected = idx === activeIdx;
                return (
                  <button
                    key={scenario.id}
                    onClick={() => handleTabClick(idx)}
                    className={`px-3 sm:px-4 py-2 rounded-xl text-xs font-semibold transition-all duration-300 ${
                      isSelected
                        ? 'bg-blue-600 text-white shadow-lg shadow-blue-500/30 scale-[1.03] ring-1 ring-blue-400/50'
                        : 'bg-white/[0.03] text-zinc-400 hover:text-white hover:bg-white/[0.06]'
                    }`}
                  >
                    <span>{scenario.tabTitle}</span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Subtle scroll progress hint */}
          <div className="hidden lg:flex items-center gap-2 pr-3 text-[11px] text-zinc-500 font-medium">
            <span>向下滚动页面自动切换</span>
            <ChevronDown className="w-3.5 h-3.5 animate-bounce text-blue-400" />
          </div>
        </div>

        {/* MacBook Window Container */}
        <div className="rounded-2xl overflow-hidden bg-[#0c0c10] border border-white/[0.12] shadow-[0_25px_70px_-15px_rgba(0,0,0,0.85)]">
          {/* Top Window Header */}
          <div className="h-9 bg-[#16161e] px-4 flex items-center justify-between border-b border-white/[0.08] select-none">
            <div className="flex items-center gap-2">
              <span className="w-3 h-3 rounded-full bg-[#ff5f56]" />
              <span className="w-3 h-3 rounded-full bg-[#ffbd2e]" />
              <span className="w-3 h-3 rounded-full bg-[#27c93f]" />
              <span className="ml-3 text-xs text-zinc-400 font-mono hidden sm:inline">
                Office Agent Bridge · 实时就地改动
              </span>
            </div>
            <div className="flex items-center gap-2 text-xs text-zinc-400 font-mono">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
              <span>当前需求 {activeIdx + 1} / 4</span>
            </div>
          </div>

          {/* Split-Screen: Left Large Text & AI Response + Right Spreadsheet */}
          <div className="grid grid-cols-1 lg:grid-cols-12 min-h-[480px]">
            {/* Left 5 Cols: Prominent User Prompt Sending & Snappy AI Response */}
            <div className="lg:col-span-5 bg-[#111116] p-5 sm:p-6 flex flex-col justify-between border-b lg:border-b-0 lg:border-r border-white/[0.08]">
              {/* Dynamic Animated Dialogue Area */}
              <div key={animKey} className="space-y-5 animate-fadeIn">
                {/* 1. Large User Prompt Sending Effect */}
                <div>
                  <div className="flex items-center gap-1.5 text-xs font-semibold text-blue-400 mb-2">
                    <CornerDownLeft className="w-4 h-4" />
                    <span>你在 AI 对话框中发送指令：</span>
                  </div>

                  <div className="p-4 sm:p-5 rounded-2xl bg-gradient-to-br from-[#1a1b26] to-[#14141c] border border-blue-500/30 shadow-xl shadow-blue-500/10 transition-all">
                    <p className="text-base sm:text-lg md:text-xl font-extrabold text-white leading-relaxed tracking-tight">
                      {currentScenario.userPrompt}
                    </p>
                  </div>
                </div>

                {/* 2. Snappy AI Tool Call & Natural Reply */}
                <div className="space-y-3 pt-2">
                  <div className="flex items-center gap-2 text-xs font-medium text-zinc-400">
                    <Zap className="w-3.5 h-3.5 text-emerald-400" />
                    <span>AI 实时处理并调用工具：</span>
                  </div>

                  {/* Clean Tool Call Badge */}
                  <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg bg-[#0a0a0f] border border-white/[0.08] font-mono text-xs text-blue-300">
                    <span className="w-2 h-2 rounded-full bg-blue-400 animate-pulse" />
                    <span className="text-zinc-500">tool:</span>
                    <span className="font-bold text-blue-400">{currentScenario.toolCall}</span>
                    <span className="text-zinc-500 text-[10px]">({currentScenario.toolActionDesc})</span>
                  </div>

                  {/* Concise AI Reply Card */}
                  <div className="p-3.5 rounded-xl bg-[#161622] border border-white/[0.06] text-xs sm:text-sm text-zinc-200 leading-relaxed shadow-sm">
                    <p>{currentScenario.agentReply}</p>
                    <div className="mt-2.5 pt-2 border-t border-white/[0.04] flex items-center justify-between text-[11px] text-zinc-500">
                      <span className="text-emerald-400 flex items-center gap-1 font-medium">
                        <CheckCircle2 className="w-3.5 h-3.5" /> 状态：当场在右侧完成
                      </span>
                      <span className="font-mono text-[10px]">耗时 24ms</span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Bottom Scroll Hint */}
              <div className="mt-6 pt-3 border-t border-white/[0.06] flex items-center justify-between text-xs text-zinc-500">
                <span>滚动页面继续看下一个操作</span>
                <span className="font-mono text-[11px] text-blue-400">
                  {activeIdx + 1} / {SCENARIOS.length}
                </span>
              </div>
            </div>

            {/* Right 7 Cols: Real Office Simulator In Action */}
            <div className="lg:col-span-7 bg-[#14141c] p-3 sm:p-4 flex flex-col justify-center">
              <div className="flex-1 min-h-[440px]">
                <OfficeSpreadsheet
                  scenario={currentScenario.id}
                  step={3} // Shows the completed action state for the current scenario
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
};
