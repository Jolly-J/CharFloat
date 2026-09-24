import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Sparkles, ShieldCheck, Zap, Layers, RefreshCw } from 'lucide-react';

interface PluginTab {
  id: 'auth' | 'topology' | 'timemachine';
  badge: string;
  title: string;
  description: string;
  image: string;
}

const TABS: PluginTab[] = [
  {
    id: 'auth',
    badge: 'AI 助手授权中心',
    title: '一键授权，无需更换常用 AI',
    description: '100% 免费使用。继续用你习惯的 WorkBuddy、豆包、Kimi、Claude 或 Codex，一键完成 MCP 桥梁授权，零迁移成本。',
    image: '/screenshots/client-auth.jpg'
  },
  {
    id: 'topology',
    badge: '本地高保真拓扑',
    title: '原生打通 WPS 与 Microsoft Office',
    description: '通过本地安全加密管道双向交互，即装即用。无需把文件上传云端，正在桌面打开的文档当场实时协同。',
    image: '/screenshots/client-topology.jpg'
  },
  {
    id: 'timemachine',
    badge: '安全时光机',
    title: '自动快照保护，改错随时秒级撤回',
    description: '放心交给 AI 动手。每次对单元格值与公式的修改均自动生成安全快照，随时一键秒级还原，原文件万无一失。',
    image: '/screenshots/client-timemachine.jpg'
  }
];

export const McpPluginSection: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'auth' | 'topology' | 'timemachine'>('auth');
  const [progress, setProgress] = useState(0);

  // Auto-play timer across tabs (matches DeliverySection 6s cycle)
  useEffect(() => {
    setProgress(0);
    const interval = 50;
    const totalTime = 6000;
    const step = (interval / totalTime) * 100;

    const timer = setInterval(() => {
      setProgress((prev) => {
        if (prev >= 100) {
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
    <section id="mcp-plugin" className="py-20 md:py-28 bg-[#f1f6fc]/80 relative scroll-mt-24 border-y border-slate-200/60">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Section Header */}
        <div className="max-w-4xl mb-12">
          <div className="inline-flex items-center gap-2 px-3.5 py-1 rounded-full bg-blue-50 border border-blue-200/80 text-blue-700 text-xs font-bold mb-4 shadow-xs">
            <Zap className="w-3.5 h-3.5 text-blue-600 fill-blue-600" />
            <span>完全免费 · 赋能现有 AI 的通用 MCP 插件</span>
          </div>
          <h2 className="text-2xl sm:text-3xl lg:text-4xl font-black text-slate-900 tracking-tight leading-tight mb-3 sm:whitespace-nowrap">
            <span>不是又一个 AI 软件，而是给你现有 AI 的</span>
            <span 
              className="bg-clip-text text-transparent ml-1.5 font-black"
              style={{
                backgroundImage: 'linear-gradient(115deg, #2563eb 0%, #06b6d4 22%, #10b981 45%, #f59e0b 72%, #ef4444 100%)'
              }}
            >
              「外挂双手」
            </span>
          </h2>
          <p className="text-base text-slate-600 leading-relaxed font-normal max-w-3xl">
            无需额外花钱订阅新 AI，也不用改变你的日常工作习惯。装上字浮，一键为你电脑上已有的 AI 软件打通直接操控{' '}
            <span className="inline-flex items-center gap-1.5 font-semibold text-slate-800 bg-white border border-slate-200/90 px-2 py-0.5 rounded-lg shadow-2xs mx-0.5 text-xs sm:text-sm align-middle">
              <img src="/wps-logo.png" alt="WPS Office" className="w-3.5 h-3.5 object-contain inline-block shrink-0" />
              <span>WPS</span>
              <span className="bg-amber-100/90 text-amber-800 border border-amber-200/80 text-[10px] font-bold px-1 py-0.2 rounded leading-tight">
                推荐
              </span>
            </span>{' '}
            与{' '}
            <span className="inline-flex items-center gap-1 font-semibold text-slate-800 bg-white border border-slate-200/90 px-2 py-0.5 rounded-lg shadow-2xs mx-0.5 text-xs sm:text-sm align-middle">
              <img src="/office-logo.png" alt="Microsoft Office" className="w-3.5 h-3.5 object-contain inline-block shrink-0" />
              <span>Office</span>
            </span>{' '}
            的能力。
          </p>
        </div>

        {/* Section Content: Left Screenshot Image (7 cols) + Right 3 Interactive Tabs (5 cols) */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-stretch">
          {/* Left Screenshot Showcase Panel */}
          <div className="lg:col-span-7 flex items-center">
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.8, ease: [0.16, 1, 0.3, 1] }}
              className="relative w-full rounded-3xl overflow-hidden bg-white border border-slate-200/90 shadow-2xl shadow-slate-900/10 group"
            >
              <AnimatePresence mode="wait">
                <motion.div
                  key={activeTab}
                  initial={{ opacity: 0, scale: 0.98 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 1.02 }}
                  transition={{ duration: 0.35, ease: [0.16, 1, 0.3, 1] }}
                  className="w-full relative"
                >
                  <img
                    src={TABS.find((t) => t.id === activeTab)?.image || '/screenshots/client-auth.jpg'}
                    alt={TABS.find((t) => t.id === activeTab)?.title}
                    className="w-full h-auto object-contain select-none"
                    loading="eager"
                  />
                  {/* Subtle inner border */}
                  <div className="absolute inset-0 rounded-3xl ring-1 ring-inset ring-black/5 pointer-events-none" />
                </motion.div>
              </AnimatePresence>
            </motion.div>
          </div>

          {/* Right 3 Tabs */}
          <div className="lg:col-span-5 flex flex-col justify-between gap-3.5">
            {TABS.map((tab) => {
              const isActive = tab.id === activeTab;
              return (
                <motion.button
                  key={tab.id}
                  onClick={() => {
                    setActiveTab(tab.id);
                    setProgress(0);
                  }}
                  whileHover={{ scale: 1.012, x: -3 }}
                  whileTap={{ scale: 0.988 }}
                  transition={{ type: "spring", stiffness: 450, damping: 26 }}
                  className={`w-full flex-1 text-left px-5 py-4 rounded-2xl border transition-all relative overflow-hidden cursor-pointer flex flex-col justify-center ${
                    isActive
                      ? 'border-blue-400/90 bg-white shadow-lg shadow-blue-500/10'
                      : 'bg-white/70 hover:bg-white border-slate-200/80 hover:border-slate-300'
                  }`}
                >
                  {/* Dynamic Background Fill: Deep blue at start (0%), fading from deep to light as countdown approaches 100% */}
                  {isActive && (
                    <div 
                      className="absolute inset-0 pointer-events-none transition-opacity duration-75 ease-linear -z-0"
                      style={{
                        backgroundColor: '#cae0f9',
                        opacity: Math.max(0.12, 0.85 - (progress / 100) * 0.73)
                      }}
                    />
                  )}

                  <div className="relative z-10">
                    <div className="flex items-center gap-2 mb-1.5">
                      <span className={`text-[11px] font-bold px-2 py-0.5 rounded-md transition-colors ${
                        isActive ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-600'
                      }`}>
                        {tab.badge}
                      </span>
                    </div>
                    <h3 className={`text-base font-bold mb-1.5 transition-colors ${
                      isActive ? 'text-blue-700' : 'text-slate-900'
                    }`}>
                      {tab.title}
                    </h3>
                    <p className={`text-xs sm:text-[13px] leading-relaxed transition-colors ${
                      isActive ? 'text-slate-700' : 'text-slate-500'
                    }`}>
                      {tab.description}
                    </p>
                  </div>
                </motion.button>
              );
            })}
          </div>
        </div>
      </div>
    </section>
  );
};
