import React from 'react';
import { Calendar, ArrowRight, CheckCircle2, FileSpreadsheet, Presentation, Sparkles } from 'lucide-react';

const STEPS = [
  {
    step: '01',
    title: '读取上下文',
    desc: '自动识别你正在编辑的“2026年经营分析表.xlsx”，无需上传或向 AI 重新解释数据结构。',
    badge: '秒级感知'
  },
  {
    step: '02',
    title: '自动补入数据',
    desc: '将 8 月营收、成本、用户数等精准写入 E 列，单元格格式与原有数字排版 100% 一致。',
    badge: '精准就地'
  },
  {
    step: '03',
    title: '更新统计口径与图表',
    desc: '按历史 7 个月的计算公式自动扩展，已有折线图平滑延长至 8 月，无须重新插入图表。',
    badge: '逻辑继承'
  },
  {
    step: '04',
    title: '标出异常指标',
    desc: '自动筛选获客成本等 3 个未达标指标，就地标记醒目高亮背景。你说“不要红标黄”，当场改色。',
    badge: '边说边改'
  },
  {
    step: '05',
    title: '联动更新 PPT 第 12 页',
    desc: '自动唤起本地正在打开的“Q3复盘汇报.pptx”，精准写入第 12 页卡片，其余 27 页原稿绝不乱动。',
    badge: '跨文档协同'
  }
];

export const RealCaseTimeline: React.FC = () => {
  return (
    <section id="real-scenario" className="py-20 bg-[#09090d] relative">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Section Header */}
        <div className="text-center max-w-3xl mx-auto mb-14">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 text-xs font-semibold mb-3">
            <Calendar className="w-3.5 h-3.5" />
            <span>真实落地复盘</span>
          </div>
          <h2 className="text-3xl sm:text-4xl font-extrabold text-white tracking-tight mb-4">
            一个真实的工作场景
          </h2>
          <p className="text-zinc-400 text-sm sm:text-base leading-relaxed">
            月度经营复盘：Excel 已用了七个月，公式图表完备；PPT 也已做完大半。看 AI 如何当场接力！
          </p>
        </div>

        {/* Story Box */}
        <div className="mb-10 p-6 rounded-2xl bg-gradient-to-r from-blue-950/20 via-[#14141c] to-indigo-950/20 border border-white/[0.08] text-sm leading-relaxed text-zinc-300">
          <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <div>
              <span className="text-xs font-bold text-blue-400 uppercase tracking-wider block mb-1">
                你只对你的 AI 说了这一句话：
              </span>
              <p className="text-base sm:text-lg font-semibold text-white">
                “把 8 月数据补进去，按原来的口径更新趋势图，把三个未完成指标标出来，然后同步更新 PPT 第 12 页。”
              </p>
            </div>
            <div className="flex items-center gap-2 flex-shrink-0">
              <span className="px-3 py-1.5 rounded-lg bg-emerald-500/15 text-emerald-300 border border-emerald-500/30 text-xs font-bold flex items-center gap-1.5">
                <CheckCircle2 className="w-4 h-4" /> 零上传 · 零下载 · 当场完成
              </span>
            </div>
          </div>
        </div>

        {/* 5-Step Process Horizontal / Grid Cards */}
        <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
          {STEPS.map((item, idx) => (
            <div 
              key={idx}
              className="p-5 rounded-xl bg-[#14141a] border border-white/[0.06] hover:border-blue-500/30 transition-all flex flex-col justify-between group hover:bg-[#181822]"
            >
              <div>
                <div className="flex items-center justify-between mb-3">
                  <span className="font-mono text-2xl font-black text-white/20 group-hover:text-blue-500/60 transition-colors">
                    {item.step}
                  </span>
                  <span className="text-[10px] px-2 py-0.5 rounded-full bg-white/[0.05] text-zinc-400 font-semibold group-hover:text-blue-300">
                    {item.badge}
                  </span>
                </div>
                <h3 className="text-sm font-bold text-white mb-2">{item.title}</h3>
                <p className="text-xs text-zinc-400 leading-relaxed">{item.desc}</p>
              </div>
            </div>
          ))}
        </div>

        {/* Quick Micro-interaction quote */}
        <div className="mt-8 p-4 rounded-xl bg-white/[0.02] border border-white/[0.06] flex items-center justify-center gap-3 text-xs sm:text-sm text-zinc-300 text-center">
          <Sparkles className="w-4 h-4 text-amber-400 flex-shrink-0" />
          <span>
            过程中你说：<strong className="text-white">“这个指标不要标红，改成黄色。”</strong> AI 当场完成。无需推倒重来，无需重新解释！
          </span>
        </div>
      </div>
    </section>
  );
};
