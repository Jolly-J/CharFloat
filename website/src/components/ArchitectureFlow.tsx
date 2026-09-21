import React from 'react';
import { User, Cpu, FileText, ArrowRightLeft, Sparkles, CheckCircle2, ShieldCheck } from 'lucide-react';

export const ArchitectureFlow: React.FC = () => {
  return (
    <section id="architecture" className="py-20 relative overflow-hidden bg-[#0a0a0e]">
      {/* Background glow */}
      <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[500px] h-[300px] bg-blue-600/10 blur-[140px] rounded-full pointer-events-none" />

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center max-w-3xl mx-auto mb-14">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 text-xs font-semibold mb-3">
            <ArrowRightLeft className="w-3.5 h-3.5" />
            <span>全新协同范式</span>
          </div>
          <h2 className="text-3xl sm:text-4xl font-extrabold text-white tracking-tight mb-4">
            和 AI 在同一个 Office 工作区
          </h2>
          <p className="text-zinc-400 text-sm sm:text-base leading-relaxed">
            不再是人和 AI 轮流上传、下载、处理文件，而是在同一个工作区共同完成一件事。
          </p>
        </div>

        {/* Visual Architecture Diagram Container */}
        <div className="max-w-4xl mx-auto p-6 sm:p-8 rounded-2xl bg-gradient-to-b from-[#14141c] to-[#0f0f14] border border-white/[0.08] shadow-2xl relative">
          {/* Top Layer: 3 Nodes Connected */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6 relative z-10">
            {/* Node 1: You */}
            <div className="p-5 rounded-xl bg-[#1c1c24] border border-white/[0.08] flex flex-col items-center text-center shadow-lg hover:border-blue-500/40 transition-colors">
              <div className="w-12 h-12 rounded-full bg-blue-500/20 border border-blue-500/40 flex items-center justify-center text-blue-400 mb-3 shadow">
                <User className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-white mb-1">你 (人类操作者)</h3>
              <p className="text-xs text-zinc-400">口语化提需求 · 边看边说 · 保持最终掌控</p>
            </div>

            {/* Node 2: The Bridge / Same Workspace (Center Hub) */}
            <div className="p-5 rounded-xl bg-gradient-to-br from-blue-900/30 via-[#1e1e2c] to-[#161622] border-2 border-blue-500/50 flex flex-col items-center text-center shadow-xl ring-1 ring-blue-500/20 relative">
              <span className="absolute -top-3 px-3 py-0.5 rounded-full bg-gradient-to-r from-blue-600 to-indigo-600 text-white text-[10px] font-bold uppercase tracking-wider shadow">
                核心枢纽
              </span>
              <div className="w-12 h-12 rounded-full bg-gradient-to-tr from-blue-600 to-indigo-600 flex items-center justify-center text-white mb-3 shadow-lg shadow-blue-500/30 animate-pulse">
                <Sparkles className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-white mb-1">同一个 Office 工作区</h3>
              <p className="text-xs text-blue-200/90 font-medium">Office Agent Bridge</p>
              <span className="mt-2 text-[10px] text-zinc-400 font-mono">本地直连 · 审计快照 · 零泄密</span>
            </div>

            {/* Node 3: Your Favorite AI */}
            <div className="p-5 rounded-xl bg-[#1c1c24] border border-white/[0.08] flex flex-col items-center text-center shadow-lg hover:border-indigo-500/40 transition-colors">
              <div className="w-12 h-12 rounded-full bg-indigo-500/20 border border-indigo-500/40 flex items-center justify-center text-indigo-400 mb-3 shadow">
                <Cpu className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-white mb-1">你习惯使用的 AI</h3>
              <p className="text-xs text-zinc-400">Claude / Cursor / ChatGPT / 豆包 / 千问</p>
            </div>
          </div>

          {/* Bottom Layer: Office Documents Target */}
          <div className="mt-8 pt-8 border-t border-white/[0.08] text-center">
            <div className="inline-flex flex-wrap items-center justify-center gap-3 px-6 py-3 rounded-xl bg-white/[0.03] border border-white/[0.08]">
              <span className="text-xs text-zinc-400 font-medium">直连操作原生文档：</span>
              <span className="flex items-center gap-1.5 px-3 py-1 rounded-lg bg-emerald-500/10 text-emerald-300 text-xs font-semibold border border-emerald-500/20">
                <FileText className="w-3.5 h-3.5" /> Excel / WPS 表格
              </span>
              <span className="flex items-center gap-1.5 px-3 py-1 rounded-lg bg-orange-500/10 text-orange-300 text-xs font-semibold border border-orange-500/20">
                <FileText className="w-3.5 h-3.5" /> PPT / WPS 演示
              </span>
              <span className="flex items-center gap-1.5 px-3 py-1 rounded-lg bg-blue-500/10 text-blue-300 text-xs font-semibold border border-blue-500/20">
                <FileText className="w-3.5 h-3.5" /> Word / WPS 文字
              </span>
            </div>
          </div>

          {/* 4 Pillars of Collaboration */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mt-8 pt-6 border-t border-white/[0.06] text-xs">
            <div className="flex items-start gap-2.5 p-2.5 rounded-lg bg-white/[0.02]">
              <CheckCircle2 className="w-4 h-4 text-emerald-400 flex-shrink-0 mt-0.5" />
              <div>
                <strong className="text-white block mb-0.5">理解当前内容</strong>
                <span className="text-zinc-400">AI 读到的正是你当前打开并查看的窗口。</span>
              </div>
            </div>

            <div className="flex items-start gap-2.5 p-2.5 rounded-lg bg-white/[0.02]">
              <CheckCircle2 className="w-4 h-4 text-emerald-400 flex-shrink-0 mt-0.5" />
              <div>
                <strong className="text-white block mb-0.5">精准就地执行</strong>
                <span className="text-zinc-400">改动特定单元格、公式或特定单页 PPT。</span>
              </div>
            </div>

            <div className="flex items-start gap-2.5 p-2.5 rounded-lg bg-white/[0.02]">
              <CheckCircle2 className="w-4 h-4 text-emerald-400 flex-shrink-0 mt-0.5" />
              <div>
                <strong className="text-white block mb-0.5">结果当场呈现</strong>
                <span className="text-zinc-400">修改直接发生在眼前，无需下载与解压。</span>
              </div>
            </div>

            <div className="flex items-start gap-2.5 p-2.5 rounded-lg bg-white/[0.02]">
              <ShieldCheck className="w-4 h-4 text-blue-400 flex-shrink-0 mt-0.5" />
              <div>
                <strong className="text-white block mb-0.5">随时无缝交接</strong>
                <span className="text-zinc-400">你改完 AI 继续，AI 改完你马上接手。</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
};
