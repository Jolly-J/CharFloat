import React from 'react';
import { RefreshCw, RotateCcw, ShieldCheck, Lock, CheckCircle2, Sliders, PlayCircle, Eye } from 'lucide-react';

export const ValueProps: React.FC = () => {
  return (
    <section id="philosophy" className="py-20 bg-[#0d0d12] relative overflow-hidden">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 items-center">
          {/* Left Column: 已经做好的，不需要重来 (80% vs 20%) */}
          <div className="space-y-6">
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-blue-500/10 text-blue-400 border border-blue-500/20 text-xs font-semibold">
              <RefreshCw className="w-3.5 h-3.5" />
              <span>保留进度 · 拒绝重来</span>
            </div>

            <h2 className="text-3xl sm:text-4xl font-extrabold text-white tracking-tight leading-tight">
              已经做好的，<br />
              <span className="text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-indigo-300">
                根本不需要重来
              </span>
            </h2>

            <p className="text-zinc-300 text-sm sm:text-base leading-relaxed">
              真实办公中，绝大多数工作并不是从 0 开始。你可能已经完成了精心排版的 80%。
            </p>

            <p className="text-zinc-400 text-sm sm:text-base leading-relaxed">
              真正耗费大量精力的是那最后的 <strong className="text-white">20%</strong>：<br />
              <span className="text-blue-300 font-medium">补数据、调公式、改图表、统一格式、移动元素、调整页面、反复微调。</span>
            </p>

            <div className="p-4 rounded-xl bg-white/[0.03] border border-white/[0.08] text-sm text-zinc-300 space-y-2">
              <div className="flex items-center gap-2 text-zinc-500">
                <span className="w-2 h-2 rounded-full bg-zinc-600" />
                <span>传统生成式 AI：擅长 “推倒重来重新生成一份”。</span>
              </div>
              <div className="flex items-center gap-2 text-emerald-300 font-semibold">
                <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                <span>Office Agent Bridge：擅长 “接着你现在的进度继续做”。</span>
              </div>
            </div>

            <div className="grid grid-cols-3 gap-3 pt-2 text-center text-xs">
              <div className="p-3 rounded-lg bg-[#14141c] border border-white/[0.06]">
                <span className="text-white font-bold block mb-1">保留已有内容</span>
                <span className="text-zinc-400 text-[11px]">排版格式完全不变</span>
              </div>
              <div className="p-3 rounded-lg bg-[#14141c] border border-white/[0.06]">
                <span className="text-white font-bold block mb-1">理解当前结构</span>
                <span className="text-zinc-400 text-[11px]">按原有统计口径继续</span>
              </div>
              <div className="p-3 rounded-lg bg-[#14141c] border border-white/[0.06]">
                <span className="text-white font-bold block mb-1">只改指定位置</span>
                <span className="text-zinc-400 text-[11px]">绝不波及其它页面</span>
              </div>
            </div>
          </div>

          {/* Right Column: 每一步，都在你的掌控之中 (随时接管与撤销) */}
          <div className="p-6 sm:p-8 rounded-2xl bg-gradient-to-br from-[#161622] to-[#121218] border border-white/[0.1] shadow-2xl relative">
            <div className="flex items-center gap-2 text-amber-400 text-xs font-semibold uppercase tracking-wider mb-4">
              <ShieldCheck className="w-4 h-4" />
              <span>全链路可控与快照审计</span>
            </div>

            <h3 className="text-2xl sm:text-3xl font-bold text-white mb-4">
              每一步，都在你的掌控之中
            </h3>

            <p className="text-zinc-400 text-sm leading-relaxed mb-6">
              AI 的动作直接发生在当前 Office 软件中，修改过程完全透明，没有任何黑盒。你随时可以打断、微调或接管。
            </p>

            {/* Quote Pill Cards */}
            <div className="space-y-3 mb-6">
              <div className="flex items-center gap-3 p-3 rounded-xl bg-[#1c1c28] border border-white/[0.06] text-xs sm:text-sm text-zinc-300">
                <span className="px-2 py-0.5 rounded bg-rose-500/20 text-rose-300 font-mono text-[11px]">你随时可以说</span>
                <span className="font-semibold text-white">“停。”</span>
                <span className="text-zinc-500 text-xs ml-auto">毫秒级暂停执行</span>
              </div>

              <div className="flex items-center gap-3 p-3 rounded-xl bg-[#1c1c28] border border-white/[0.06] text-xs sm:text-sm text-zinc-300">
                <span className="px-2 py-0.5 rounded bg-amber-500/20 text-amber-300 font-mono text-[11px]">你随时可以说</span>
                <span className="font-semibold text-white">“刚才那步撤回。”</span>
                <span className="text-zinc-500 text-xs ml-auto">内存快照一键还原</span>
              </div>

              <div className="flex items-center gap-3 p-3 rounded-xl bg-[#1c1c28] border border-white/[0.06] text-xs sm:text-sm text-zinc-300">
                <span className="px-2 py-0.5 rounded bg-blue-500/20 text-blue-300 font-mono text-[11px]">你随时可以说</span>
                <span className="font-semibold text-white">“这里只改数字，不改格式。”</span>
                <span className="text-zinc-500 text-xs ml-auto">精准颗粒度约束</span>
              </div>

              <div className="flex items-center gap-3 p-3 rounded-xl bg-[#1c1c28] border border-white/[0.06] text-xs sm:text-sm text-zinc-300">
                <span className="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 font-mono text-[11px]">你随时可以说</span>
                <span className="font-semibold text-white">“这个我自己来。”</span>
                <span className="text-zinc-500 text-xs ml-auto">人类随时直接接手</span>
              </div>
            </div>

            {/* Takeaway */}
            <div className="p-3.5 rounded-xl bg-blue-500/10 border border-blue-500/20 text-center">
              <span className="text-xs sm:text-sm font-bold text-blue-200">
                AI 提供充沛执行力，人始终保留最终控制权。
              </span>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
};
