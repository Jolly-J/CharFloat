import React from 'react';
import { Terminal, Sparkles, Heart, ArrowUp } from 'lucide-react';

interface FooterProps {
  onOpenQuickStart: () => void;
}

export const Footer: React.FC<FooterProps> = ({ onOpenQuickStart }) => {
  const scrollToTop = () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  return (
    <footer className="bg-[#050507] text-zinc-400 border-t border-white/[0.08] relative overflow-hidden">
      {/* Big Closing CTA Strip */}
      <div className="max-w-5xl mx-auto px-4 sm:px-6 lg:px-8 py-20 text-center">
        <h2 className="text-3xl sm:text-5xl font-extrabold text-white tracking-tight mb-6">
          给你的 AI，装上真正的 Office 操作能力
        </h2>

        {/* 4 Virtues Pills */}
        <div className="flex flex-wrap items-center justify-center gap-4 sm:gap-6 text-sm sm:text-lg font-bold text-zinc-300 mb-8">
          <span className="px-4 py-2 rounded-xl bg-blue-500/10 text-blue-300 border border-blue-500/20">
            看得懂。
          </span>
          <span className="px-4 py-2 rounded-xl bg-emerald-500/10 text-emerald-300 border border-emerald-500/20">
            改得准。
          </span>
          <span className="px-4 py-2 rounded-xl bg-amber-500/10 text-amber-300 border border-amber-500/20">
            动得快。
          </span>
          <span className="px-4 py-2 rounded-xl bg-purple-500/10 text-purple-300 border border-purple-500/20">
            全程可控。
          </span>
        </div>

        <p className="text-lg sm:text-xl text-zinc-300 font-medium max-w-2xl mx-auto mb-10 leading-relaxed">
          让普通 AI，从“帮你想”升级到<br />
          <span className="text-white font-extrabold text-2xl sm:text-3xl underline decoration-blue-500 decoration-4 underline-offset-8">
            “和你一起把事情做完”
          </span>。
        </p>

        <button
          onClick={onOpenQuickStart}
          className="inline-flex items-center gap-2 px-8 py-4 rounded-xl bg-gradient-to-r from-blue-600 via-indigo-600 to-emerald-600 hover:from-blue-500 hover:to-emerald-500 text-white font-bold text-base shadow-2xl shadow-blue-500/30 hover:scale-[1.03] active:scale-[0.98] transition-all"
        >
          <Sparkles className="w-5 h-5 text-blue-200" />
          <span>立即开始使用</span>
        </button>
      </div>

      {/* Standard Bottom Credits Bar */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 border-t border-white/[0.06] flex flex-col sm:flex-row items-center justify-between gap-4 text-xs">
        <div className="flex items-center gap-2.5">
          <div className="w-6 h-6 rounded-md bg-blue-600 flex items-center justify-center text-white font-bold text-[10px]">
            OB
          </div>
          <span className="font-semibold text-zinc-200">Office Agent Bridge</span>
          <span>© 2026. 面向 Microsoft Office 与 WPS Office 的通用 AI 连接桥梁。</span>
        </div>

        <div className="flex items-center gap-6">
          <button 
            onClick={scrollToTop}
            className="flex items-center gap-1 text-zinc-400 hover:text-white transition-colors"
          >
            <span>回到顶部</span>
            <ArrowUp className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </footer>
  );
};
