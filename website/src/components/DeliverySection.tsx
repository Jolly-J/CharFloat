import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';

interface DeliveryTab {
  id: 'formula' | 'highlight' | 'chart';
  title: string;
  description: string;
  image: string;
}

const TABS: DeliveryTab[] = [
  {
    id: 'formula',
    title: '智能公式与增量补齐',
    description: '识别已有计算口径，一句话将各列缺失数据精准补齐并继承复杂公式，绝不破坏原有数据结构。',
    image: '/delivery/delivery-formula.png'
  },
  {
    id: 'highlight',
    title: '条件标色与指标警示',
    description: '按照你口语化的业务规则（如“成本超标的标黄，不要标红”），精准定位对应单元格并柔和标色，保持原版面整洁。',
    image: '/delivery/delivery-highlight.png'
  },
  {
    id: 'chart',
    title: 'WPS 原生图表就地生成',
    description: '无需重新插入新文件。直接在当前工作表右下方生成绑好数据源的嵌入式原生图表，随时可双击微调。',
    image: '/delivery/delivery-chart.png'
  }
];

export const DeliverySection: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'formula' | 'highlight' | 'chart'>('formula');
  const [progress, setProgress] = useState(0);

  // Auto-play timer across tabs
  useEffect(() => {
    setProgress(0);
    const interval = 50; // update every 50ms
    const totalTime = 6000; // 6s per tab
    const step = (interval / totalTime) * 100;

    const timer = setInterval(() => {
      setProgress((prev) => {
        if (prev >= 100) {
          // Switch to next tab
          setActiveTab((current) => {
            const nextIdx = (TABS.findIndex((t) => t.id === current) + 1) % TABS.length;
            return TABS[nextIdx].id;
          });
          return 0;
        }
        return prev + step;
      });
    }, interval);

    return () => clearInterval(timer);
  }, [activeTab]);

  return (
    <section id="delivery" className="py-20 md:py-28 bg-[#f8fafc] relative scroll-mt-24">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Section Header */}
        <div className="max-w-4xl mb-10">
          <h2 className="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight leading-tight mb-3">
            改到哪儿，看到哪儿 · 原生操控更强大
          </h2>
          <p className="text-base text-slate-600 leading-relaxed font-normal max-w-3xl">
            赋予 AI 直接操控 WPS 与 Office 的底层双手。不再让 AI 在黑盒里盲猜重做文件——直接在眼前正打开的文档里精准落笔，复杂公式自动关联、原生图表就地生成，排版与数据 100% 完美继承。
          </p>
        </div>

        {/* Section Content: 3 Left Tabs perfectly aligned in height with Right 16:9 Showcase Image */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-stretch">
          {/* Left 3 Tabs */}
          <div className="lg:col-span-5 flex flex-col justify-between gap-3">
            {TABS.map((tab) => {
              const isActive = tab.id === activeTab;
              return (
                <motion.button
                  key={tab.id}
                  onClick={() => {
                    setActiveTab(tab.id);
                    setProgress(0);
                  }}
                  whileHover={{ scale: 1.012, x: 3 }}
                  whileTap={{ scale: 0.988 }}
                  transition={{ type: "spring", stiffness: 450, damping: 26 }}
                  className={`w-full flex-1 text-left px-5 py-3.5 rounded-2xl border transition-all relative overflow-hidden cursor-pointer flex flex-col justify-center ${
                    isActive
                      ? 'border-blue-400/90 shadow-md shadow-blue-500/10'
                      : 'bg-white/60 hover:bg-white border-slate-200/80 hover:border-slate-300'
                  }`}
                >
                  {/* Dynamic Background Fill: Deep blue at start (0%), fading from deep to light as countdown approaches 100% */}
                  {isActive && (
                    <div 
                      className="absolute inset-0 pointer-events-none transition-opacity duration-75 ease-linear -z-0"
                      style={{
                        backgroundColor: '#cae0f9',
                        opacity: Math.max(0.1, 0.85 - (progress / 100) * 0.75)
                      }}
                    />
                  )}

                  <div className="relative z-10">
                    <h3 className={`text-base font-bold mb-1 transition-colors ${
                      isActive ? 'text-blue-700' : 'text-slate-900'
                    }`}>
                      {tab.title}
                    </h3>
                    <p className={`text-xs sm:text-sm leading-relaxed transition-colors ${
                      isActive ? 'text-slate-700' : 'text-slate-500'
                    }`}>
                      {tab.description}
                    </p>
                  </div>
                </motion.button>
              );
            })}
          </div>

          {/* Right Showcase Image Panel: Exact 16:9 ratio with zero cropping */}
          <div className="lg:col-span-7 flex items-center">
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.8, ease: [0.16, 1, 0.3, 1] }}
              className="relative w-full aspect-video rounded-3xl overflow-hidden bg-white border border-slate-200/90 shadow-xl group"
            >
              <AnimatePresence mode="wait">
                <motion.div
                  key={activeTab}
                  initial={{ opacity: 0, scale: 0.98 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 1.02 }}
                  transition={{ duration: 0.35, ease: [0.16, 1, 0.3, 1] }}
                  className="w-full h-full relative"
                >
                  <img
                    src={TABS.find((t) => t.id === activeTab)?.image || '/hero-showcase.png'}
                    alt={TABS.find((t) => t.id === activeTab)?.title}
                    className="w-full h-full object-cover select-none"
                    loading="eager"
                  />
                  {/* Subtle inner border */}
                  <div className="absolute inset-0 rounded-3xl ring-1 ring-inset ring-black/5 pointer-events-none" />
                </motion.div>
              </AnimatePresence>
            </motion.div>
          </div>
        </div>
      </div>
    </section>
  );
};
