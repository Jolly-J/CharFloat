import React from 'react';
import { Presentation, CheckCircle, ArrowRight, Shield, Layers, Sparkles } from 'lucide-react';

interface OfficeSlidesProps {
  isUpdated: boolean;
  selectedAgentName: string;
}

export const OfficeSlides: React.FC<OfficeSlidesProps> = ({ isUpdated, selectedAgentName }) => {
  return (
    <div className="flex flex-col h-full bg-[#1e1e24] rounded-xl overflow-hidden border border-white/[0.08] shadow-2xl font-sans text-xs select-none">
      {/* PPT Title & Menu Bar */}
      <div className="bg-[#c43e1c] px-3 py-1.5 flex items-center justify-between text-white border-b border-black/20">
        <div className="flex items-center gap-2">
          <Presentation className="w-4 h-4 text-orange-100" />
          <span className="font-semibold text-xs tracking-wide">Q3战略复盘汇报.pptx - PowerPoint / WPS 演示</span>
          <span className="px-1.5 py-0.2 text-[10px] bg-white/20 rounded text-orange-100 font-mono">第 12 页精准定向</span>
        </div>
        <div className="flex items-center gap-2 text-[11px] text-orange-100/80">
          <span className="flex items-center gap-1">
            <span className="w-1.5 h-1.5 rounded-full bg-emerald-300 animate-pulse"></span>
            {selectedAgentName} 正在跨文档协同
          </span>
        </div>
      </div>

      {/* Ribbon Bar Mockup */}
      <div className="bg-[#2b2b36] px-3 py-1 flex items-center gap-4 text-zinc-300 border-b border-white/[0.06] text-[11px]">
        <span className="text-white font-medium border-b-2 border-[#c43e1c] pb-0.5">开始 (Home)</span>
        <span className="hover:text-white cursor-pointer">设计 (Design)</span>
        <span className="hover:text-white cursor-pointer">切换 (Transitions)</span>
        <span className="hover:text-white cursor-pointer">动画 (Animations)</span>
        <span className="hover:text-white cursor-pointer">放映 (Slide Show)</span>
      </div>

      {/* Main PPT Area: Thumbnails Left + Active Slide Canvas Right */}
      <div className="flex-1 flex overflow-hidden bg-[#141418]">
        {/* Left Thumbnails Strip */}
        <div className="w-28 bg-[#1a1a20] border-r border-white/[0.06] p-2 flex flex-col gap-2 overflow-y-auto">
          {/* Slide 10 */}
          <div className="p-1.5 rounded bg-[#23232c] border border-white/[0.04] text-[9px] text-zinc-400 opacity-60">
            <div className="flex justify-between items-center mb-1">
              <span>10</span>
              <span className="text-[8px] text-zinc-500">锁定未动</span>
            </div>
            <div className="h-9 bg-[#16161b] rounded flex items-center justify-center text-zinc-600">
              Q2 业务总结
            </div>
          </div>

          {/* Slide 11 */}
          <div className="p-1.5 rounded bg-[#23232c] border border-white/[0.04] text-[9px] text-zinc-400 opacity-60">
            <div className="flex justify-between items-center mb-1">
              <span>11</span>
              <span className="text-[8px] text-zinc-500">锁定未动</span>
            </div>
            <div className="h-9 bg-[#16161b] rounded flex items-center justify-center text-zinc-600">
              渠道投放分析
            </div>
          </div>

          {/* Slide 12 (Target Active Slide) */}
          <div className="p-1.5 rounded bg-blue-500/15 border-2 border-blue-400 text-[9px] text-blue-300 shadow-md">
            <div className="flex justify-between items-center mb-1">
              <span className="font-bold text-white">12</span>
              <span className="text-[8px] px-1 py-0.2 rounded bg-blue-500/30 text-blue-200 font-semibold">定向修改</span>
            </div>
            <div className="h-9 bg-[#1e2330] rounded flex flex-col items-center justify-center text-center p-1">
              <span className="font-semibold text-[8px] text-white">8月核心经营指标</span>
              <span className="text-[7px] text-emerald-400">● 实时同步</span>
            </div>
          </div>

          {/* Slide 13 */}
          <div className="p-1.5 rounded bg-[#23232c] border border-white/[0.04] text-[9px] text-zinc-400 opacity-60">
            <div className="flex justify-between items-center mb-1">
              <span>13</span>
              <span className="text-[8px] text-zinc-500">锁定未动</span>
            </div>
            <div className="h-9 bg-[#16161b] rounded flex items-center justify-center text-zinc-600">
              Q4 规划路线
            </div>
          </div>
        </div>

        {/* Right Active Slide Canvas */}
        <div className="flex-1 p-4 flex flex-col items-center justify-center overflow-auto bg-[#101014]">
          {/* Slide Paper Container (16:9 aspect) */}
          <div className="w-full max-w-[500px] aspect-video bg-gradient-to-br from-[#1c1d26] to-[#14141a] rounded-lg border border-white/[0.1] shadow-2xl p-4 flex flex-col justify-between relative overflow-hidden">
            {/* Slide Header */}
            <div>
              <div className="flex items-center justify-between border-b border-white/[0.08] pb-2 mb-3">
                <div>
                  <h3 className="text-sm font-bold text-white tracking-wide">
                    2026 Q3 月度经营成果与预警指标
                  </h3>
                  <p className="text-[10px] text-zinc-400">
                    来源：经营分析表.xlsx 同步更新 ｜ 范围：仅修改当前第 12 页
                  </p>
                </div>
                <span className="px-2 py-0.5 rounded bg-orange-500/20 text-orange-300 font-mono text-[9px]">
                  SLIDE 12
                </span>
              </div>

              {/* Slide Metric Cards Grid */}
              <div className="grid grid-cols-3 gap-2 mb-3">
                {/* Metric 1 */}
                <div className="bg-white/[0.03] p-2 rounded border border-white/[0.06]">
                  <span className="text-[9px] text-zinc-400 block mb-1">8月营收总计</span>
                  <div className="flex items-baseline gap-1">
                    <span className="text-sm font-extrabold text-emerald-400 font-mono">
                      {isUpdated ? '612.0 万' : '560.0 万'}
                    </span>
                    {isUpdated && <span className="text-[8px] text-emerald-400 font-semibold">+9.2%</span>}
                  </div>
                  <span className="text-[8px] text-zinc-500">已达成本月考核指标</span>
                </div>

                {/* Metric 2 */}
                <div className="bg-white/[0.03] p-2 rounded border border-white/[0.06]">
                  <span className="text-[9px] text-zinc-400 block mb-1">订单转化率</span>
                  <div className="flex items-baseline gap-1">
                    <span className="text-sm font-extrabold text-emerald-400 font-mono">
                      {isUpdated ? '3.8%' : '3.5%'}
                    </span>
                    {isUpdated && <span className="text-[8px] text-emerald-400 font-semibold">+0.3pp</span>}
                  </div>
                  <span className="text-[8px] text-zinc-500">转化漏斗健康</span>
                </div>

                {/* Metric 3 (Abnormal Warning) */}
                <div className={`p-2 rounded border transition-colors ${
                  isUpdated ? 'bg-amber-500/15 border-amber-500/30' : 'bg-white/[0.03] border-white/[0.06]'
                }`}>
                  <span className="text-[9px] text-zinc-400 block mb-1">获客成本 CAC</span>
                  <div className="flex items-baseline gap-1">
                    <span className={`text-sm font-extrabold font-mono ${
                      isUpdated ? 'text-amber-300' : 'text-zinc-200'
                    }`}>
                      {isUpdated ? '68 元 ⚠️' : '48 元'}
                    </span>
                  </div>
                  <span className={`text-[8px] ${isUpdated ? 'text-amber-400 font-semibold' : 'text-zinc-500'}`}>
                    {isUpdated ? '异常升高 (超标41%)' : '正常范围'}
                  </span>
                </div>
              </div>
            </div>

            {/* Bottom Target Modification Notice */}
            <div className="bg-blue-500/10 border border-blue-500/20 rounded px-2.5 py-1.5 flex items-center justify-between text-[9px] text-blue-200">
              <span className="flex items-center gap-1.5">
                <CheckCircle className="w-3.5 h-3.5 text-blue-400 flex-shrink-0" />
                <span>
                  {isUpdated 
                    ? 'AI 指令精准完成：仅同步更新第 12 页数字与卡片，其余 27 页原稿格式 100% 完整保留！'
                    : '等待执行指令：同步更新 PPT 第 12 页，其他页面不要动...'}
                </span>
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* PPT Bottom Status Bar */}
      <div className="bg-[#1c1c24] px-3 py-1 flex items-center justify-between border-t border-white/[0.06] text-[10px] text-zinc-400">
        <span>幻灯片 12 / 28 ｜ 中文(中国)</span>
        <span className="text-zinc-500">视图比例: 82% ｜ 锁定其他页面保护</span>
      </div>
    </div>
  );
};
