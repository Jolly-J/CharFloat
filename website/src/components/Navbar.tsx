import React from 'react';
import { Terminal, Shield, Zap, Sparkles, ExternalLink } from 'lucide-react';

interface NavbarProps {
  onOpenQuickStart: () => void;
}

export const Navbar: React.FC<NavbarProps> = ({ onOpenQuickStart }) => {
  return (
    <header className="fixed top-0 left-0 right-0 z-50 bg-[#08080a]/80 backdrop-blur-md border-b border-white/[0.06]">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
        {/* Brand Logo */}
        <div className="flex items-center gap-3">
          <div className="relative flex items-center justify-center w-10 h-10 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-600 shadow-lg shadow-blue-500/20">
            <Terminal className="w-5 h-5 text-white" />
            <span className="absolute -top-1 -right-1 flex h-3 w-3">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-3 w-3 bg-emerald-500"></span>
            </span>
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="font-bold text-lg tracking-tight text-white">Office Agent Bridge</span>
              <span className="text-[10px] font-semibold uppercase tracking-wider px-2 py-0.5 rounded-full bg-blue-500/10 text-blue-400 border border-blue-500/20">
                v2.1
              </span>
            </div>
            <p className="text-xs text-zinc-400 hidden sm:block">让 AI 真正接管本地正在打开的 Office</p>
          </div>
        </div>

        {/* Navigation Links */}
        <nav className="hidden md:flex items-center gap-8 text-sm font-medium text-zinc-400">
          <a href="#interactive-demo" className="hover:text-blue-400 transition-colors">实景演练</a>
          <a href="#architecture" className="hover:text-blue-400 transition-colors">同台架构</a>
          <a href="#comparison" className="hover:text-blue-400 transition-colors">能力质变</a>
          <a href="#real-scenario" className="hover:text-blue-400 transition-colors">真实案例</a>
          <a href="#philosophy" className="hover:text-blue-400 transition-colors">产品理念</a>
        </nav>

        {/* Actions */}
        <div className="flex items-center gap-3">
          <button
            onClick={onOpenQuickStart}
            className="flex items-center gap-2 px-4 py-2 text-xs sm:text-sm font-semibold text-white bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 rounded-lg shadow-md shadow-blue-500/20 transition-all hover:scale-[1.02] active:scale-[0.98]"
          >
            <Sparkles className="w-4 h-4 text-blue-200" />
            <span>接入你的 AI</span>
          </button>
        </div>
      </div>
    </header>
  );
};
