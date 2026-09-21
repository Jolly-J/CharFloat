import React from 'react';
import { XCircle, CheckCircle2, ArrowRight, Sparkles } from 'lucide-react';

const COMPARISON_DATA = [
  {
    before: '告诉你怎么做（建议、步骤、代码）',
    after: '直接替你操作（当场在软件里改完）',
    highlight: true,
  },
  {
    before: '生成一个新文件（需重新下载命名）',
    after: '修改你当前的文件（保留上下文与进度）',
    highlight: true,
  },
  {
    before: '只能理解你单次上传的静态内容',
    after: '持续理解当前真实工作状态与选区',
    highlight: false,
  },
  {
    before: '擅长整体重新生成（容易推翻已有排版）',
    after: '可以精准修改一个细节（单元格/一段话）',
    highlight: false,
  },
  {
    before: '做完以后把结果一次性扔给你',
    after: '边看、边说、边改（人机实时同台对话）',
    highlight: true,
  },
  {
    before: '修改过程不可见（黑盒等待）',
    after: '每一步都发生在眼前（清晰可见、随时可停）',
    highlight: false,
  },
  {
    before: 'AI 和 Office 相互割裂（两边倒腾）',
    after: 'AI 真正进入 Office 工作区（浑然一体）',
    highlight: true,
  },
];

export const ComparisonTable: React.FC = () => {
  return (
    <section id="comparison" className="py-20 bg-[#08080c] relative">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Section Header */}
        <div className="text-center max-w-3xl mx-auto mb-14">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 text-xs font-semibold mb-3">
            <Sparkles className="w-3.5 h-3.5" />
            <span>核心质变对比</span>
          </div>
          <h2 className="text-3xl sm:text-4xl font-extrabold text-white tracking-tight mb-4">
            普通 AI + 我们，发生了什么变化？
          </h2>
          <p className="text-zinc-400 text-sm sm:text-base leading-relaxed">
            让普通 AI 从“会想、会说”，升级到“看得见、改得动、做得完”。
          </p>
        </div>

        {/* Comparison Matrix Table Card */}
        <div className="rounded-2xl overflow-hidden border border-white/[0.08] bg-[#121216]/80 shadow-2xl backdrop-blur-md">
          {/* Header Row */}
          <div className="grid grid-cols-1 md:grid-cols-2 bg-[#181820] border-b border-white/[0.08] text-sm font-bold">
            <div className="p-4 sm:p-5 flex items-center gap-2.5 text-zinc-400 border-b md:border-b-0 md:border-r border-white/[0.08]">
              <XCircle className="w-5 h-5 text-zinc-500" />
              <span>普通 AI / 传统方式</span>
            </div>
            <div className="p-4 sm:p-5 flex items-center gap-2.5 text-emerald-400 bg-emerald-500/[0.03]">
              <CheckCircle2 className="w-5 h-5 text-emerald-400" />
              <span>接入 Office Agent Bridge 之后</span>
            </div>
          </div>

          {/* Table Body Rows */}
          <div className="divide-y divide-white/[0.05]">
            {COMPARISON_DATA.map((item, idx) => (
              <div 
                key={idx} 
                className={`grid grid-cols-1 md:grid-cols-2 text-xs sm:text-sm transition-colors ${
                  item.highlight ? 'bg-white/[0.015]' : 'hover:bg-white/[0.01]'
                }`}
              >
                {/* Left: Ordinary AI */}
                <div className="p-4 sm:p-5 flex items-center gap-3 text-zinc-400 border-b md:border-b-0 md:border-r border-white/[0.06]">
                  <span className="w-1.5 h-1.5 rounded-full bg-zinc-600 flex-shrink-0" />
                  <span className="line-through decoration-zinc-600">{item.before}</span>
                </div>

                {/* Right: Connected AI */}
                <div className="p-4 sm:p-5 flex items-center gap-3 text-white font-medium bg-gradient-to-r from-transparent to-emerald-500/[0.02]">
                  <ArrowRight className="w-4 h-4 text-emerald-400 flex-shrink-0" />
                  <span className={item.highlight ? 'text-emerald-300 font-bold' : 'text-zinc-200'}>
                    {item.after}
                  </span>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Bottom Banner Callout */}
        <div className="mt-8 p-6 rounded-xl bg-gradient-to-r from-blue-900/20 via-indigo-900/20 to-purple-900/20 border border-blue-500/20 text-center">
          <p className="text-sm sm:text-base font-medium text-zinc-200">
            无需改变你的使用习惯：继续在 Cursor、Claude Code、ChatGPT 或自建网页里聊天，
            <br className="hidden sm:inline" />
            <strong className="text-white font-bold underline decoration-blue-400 decoration-2 underline-offset-4">
              AI 的所有回答，直接化作 Office 软件里的精准动作。
            </strong>
          </p>
        </div>
      </div>
    </section>
  );
};
