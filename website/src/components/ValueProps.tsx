import React from 'react';
import { Eye, Edit3, Zap, Shield, Sparkles, CheckCircle2 } from 'lucide-react';

const VALUES = [
  {
    icon: Eye,
    title: '看得懂。',
    subtitle: '秒级感知当前真实状态',
    desc: '自动识别你正在编辑的文件、活动工作表与光标当前选区。你不需要反复向 AI 解释背景，开口就是上下文。',
    color: 'blue'
  },
  {
    icon: Edit3,
    title: '改得准。',
    subtitle: '精准到单元格与段落',
    desc: '可以整体批量补齐公式，也可以只修改一个单元格的数值或底色。绝不擅自推翻你的母版与排版成果。',
    color: 'emerald'
  },
  {
    icon: Zap,
    title: '动得快。',
    subtitle: '本地毫秒级就地完成',
    desc: '基于轻量本地高能进程通信，毫秒级直接调用 WPS / Office 原生底层，拒绝上传等待与漫长导出。',
    color: 'amber'
  },
  {
    icon: Shield,
    title: '全程可控。',
    subtitle: '随时接管，随时撤销',
    desc: '每一步操作均自动生成快照。一句“撤回刚才那步”，秒级复原。你拥有至高无上的掌控权。',
    color: 'indigo'
  }
];

export const ValueProps: React.FC = () => {
  return (
    <section className="py-20 bg-slate-50 relative overflow-hidden">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center max-w-3xl mx-auto mb-16">
          <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-blue-50 text-blue-600 border border-blue-200/80 text-xs font-semibold mb-3">
            <Sparkles className="w-3.5 h-3.5" />
            <span>核心产品理念</span>
          </div>
          <h2 className="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight mb-4">
            真正的办公助理，应该是什么样？
          </h2>
          <p className="text-slate-600 text-sm sm:text-base leading-relaxed">
            不用把简单的事情搞复杂。坐在电脑前，你负责把控方向，它负责落地干活。
          </p>
        </div>

        {/* 4 Clean Value Cards */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 max-w-4xl mx-auto">
          {VALUES.map((val, idx) => {
            const Icon = val.icon;
            return (
              <div
                key={idx}
                className="p-8 rounded-3xl bg-white border border-slate-200/90 shadow-sm hover:shadow-md hover:border-blue-300 transition-all flex flex-col justify-between"
              >
                <div>
                  <div className="flex items-center justify-between mb-4">
                    <div className="w-12 h-12 rounded-2xl bg-blue-50 text-blue-600 flex items-center justify-center shadow-sm">
                      <Icon className="w-6 h-6" />
                    </div>
                    <span className="text-xs font-mono font-bold text-slate-300">
                      0{idx + 1}
                    </span>
                  </div>

                  <h3 className="text-2xl font-black text-slate-900 mb-1">
                    {val.title}
                  </h3>
                  <p className="text-sm font-bold text-blue-600 mb-3">
                    {val.subtitle}
                  </p>
                  <p className="text-sm text-slate-600 leading-relaxed">
                    {val.desc}
                  </p>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </section>
  );
};
