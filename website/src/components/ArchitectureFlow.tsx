import React from 'react';
import { User, Cpu, FileText, ArrowRightLeft, Sparkles, CheckCircle2, ShieldCheck, Laptop } from 'lucide-react';

export const ArchitectureFlow: React.FC = () => {
  return (
    <section id="architecture" className="py-20 bg-slate-50 relative overflow-hidden">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center max-w-3xl mx-auto mb-14">
          <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-blue-50 text-blue-600 border border-blue-200/80 text-xs font-semibold mb-3">
            <ArrowRightLeft className="w-3.5 h-3.5" />
            <span>全新协同范式</span>
          </div>
          <h2 className="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight mb-4">
            和 AI 在同一个 Office 工作区
          </h2>
          <p className="text-slate-600 text-sm sm:text-base leading-relaxed">
            不再是人和 AI 轮流上传、下载、处理文件，而是在同一个工作区共同把事做完。
          </p>
        </div>

        {/* Visual Architecture Diagram Container */}
        <div className="max-w-4xl mx-auto p-6 sm:p-10 rounded-3xl bg-white border border-slate-200/90 shadow-xl relative">
          {/* Top Layer: 3 Nodes Connected */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6 relative z-10">
            {/* Node 1: You */}
            <div className="p-6 rounded-2xl bg-slate-50 border border-slate-200 flex flex-col items-center text-center shadow-sm hover:border-blue-300 transition-colors">
              <div className="w-12 h-12 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center mb-3">
                <User className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 mb-1">你 (人类操作者)</h3>
              <p className="text-xs text-slate-500">像平时一样说话 · 边看边提要求 · 随时接管微调</p>
            </div>

            {/* Node 2: Center Hub */}
            <div className="p-6 rounded-2xl bg-gradient-to-b from-blue-50/80 to-indigo-50/50 border-2 border-blue-500 flex flex-col items-center text-center shadow-lg relative">
              <span className="absolute -top-3 px-3 py-0.5 rounded-full bg-blue-600 text-white text-[10px] font-bold tracking-wider shadow">
                本地连接枢纽
              </span>
              <div className="w-12 h-12 rounded-full bg-blue-600 text-white flex items-center justify-center mb-3 shadow-md shadow-blue-500/25">
                <Sparkles className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 mb-1">字浮 CharFloat</h3>
              <p className="text-xs text-blue-700 font-medium">本地直驱 · 毫秒级响应</p>
              <span className="mt-2 text-[10px] text-slate-500 font-mono bg-white px-2 py-0.5 rounded-full border border-slate-200">
                100% 本地运算 · 零文档泄露
              </span>
            </div>

            {/* Node 3: AI Model */}
            <div className="p-6 rounded-2xl bg-slate-50 border border-slate-200 flex flex-col items-center text-center shadow-sm hover:border-indigo-300 transition-colors">
              <div className="w-12 h-12 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center mb-3">
                <Cpu className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 mb-1">你喜欢的任意 AI</h3>
              <p className="text-xs text-slate-500">豆包 / DeepSeek / Claude / ChatGPT / Cursor / 千问</p>
            </div>
          </div>

          {/* Bottom Layer: Office Target Document */}
          <div className="mt-8 pt-8 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4 bg-slate-50/80 p-5 rounded-2xl border border-slate-200">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center shadow-sm">
                <FileText className="w-5 h-5" />
              </div>
              <div>
                <h4 className="text-sm font-bold text-slate-900">
                  当前电脑正在打开的 WPS / Excel 表格
                </h4>
                <p className="text-xs text-slate-500">
                  直接在当前文档原地操作，保留原汁原味排版与公式
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 text-xs font-semibold text-emerald-700 bg-emerald-50 px-3 py-1.5 rounded-full border border-emerald-200">
              <CheckCircle2 className="w-4 h-4" />
              <span>无需另存导出 · 即刻完成</span>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
};
