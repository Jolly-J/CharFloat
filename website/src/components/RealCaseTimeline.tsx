import React from 'react';
import { Calendar, ArrowRight, CheckCircle2, FileSpreadsheet, Presentation, Sparkles } from 'lucide-react';

const STEPS = [
  {
    step: '01',
    title: '自动感知当前文件',
    desc: '自动识别你正在编辑的“2026年经营分析表.xlsx”，无需另存或向 AI 重新解释数据结构。',
    badge: '秒级感知'
  },
  {
    step: '02',
    title: '自动补入缺失数据',
    desc: '将 8 月营收、成本、用户转化率等精准写入 D 列，单元格格式与原有数字排版 100% 一致。',
    badge: '精准就地'
  },
  {
    step: '03',
    title: '更新统计口径与图表',
    desc: '按历史 7 个月的计算公式自动扩展同比，已有折线图平滑延长至 8 月，无须重新插入图表。',
    badge: '逻辑继承'
  },
  {
    step: '04',
    title: '标出异常指标',
    desc: '自动筛选获客成本等未达标指标，就地标记醒目高亮背景。你说“不要红标黄”，当场改色。',
    badge: '边说边改'
  },
  {
    step: '05',
    title: '联动更新汇报文档',
    desc: '自动唤起本地正在打开的“Q3复盘汇报.docx / pptx”，精准写入对应总结卡片，其余页面原稿绝不乱动。',
    badge: '跨文档协同'
  }
];

export const RealCaseTimeline: React.FC = () => {
  return (
    <section id="real-scenario" className="py-20 bg-white relative">
      <div className="max-w-5xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center max-w-3xl mx-auto mb-16">
          <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-blue-50 text-blue-600 border border-blue-200/80 text-xs font-semibold mb-3">
            <Calendar className="w-3.5 h-3.5" />
            <span>真实业务场景复盘</span>
          </div>
          <h2 className="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight mb-4">
            周五下午 5 点，它如何替你省下 2 小时
          </h2>
          <p className="text-slate-600 text-sm sm:text-base leading-relaxed">
            不用从头开始建表，不用自己手工核对几十个公式。几句口语，整个报表就地更新完毕。
          </p>
        </div>

        {/* Timeline Clean Cards */}
        <div className="space-y-4 max-w-4xl mx-auto">
          {STEPS.map((item, idx) => (
            <div
              key={idx}
              className="p-5 sm:p-6 rounded-2xl bg-slate-50 border border-slate-200/80 hover:border-blue-300 hover:bg-blue-50/20 transition-all flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4"
            >
              <div className="flex items-start gap-4">
                <span className="font-mono text-2xl font-black text-blue-600/40">
                  {item.step}
                </span>
                <div>
                  <h3 className="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
                    <span>{item.title}</span>
                    <span className="text-[11px] px-2 py-0.5 rounded-full bg-blue-50 text-blue-600 border border-blue-200/60 font-semibold">
                      {item.badge}
                    </span>
                  </h3>
                  <p className="text-xs sm:text-sm text-slate-600 leading-relaxed">
                    {item.desc}
                  </p>
                </div>
              </div>

              <div className="hidden sm:flex items-center text-emerald-600 text-xs font-bold gap-1 shrink-0">
                <CheckCircle2 className="w-4 h-4" />
                <span>自动生效</span>
              </div>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
};
