import React, { useState, useEffect, useRef } from 'react';
import { 
  Send, 
  Terminal, 
  Zap, 
  CheckCircle2, 
  Play, 
  Pause, 
  RotateCcw, 
  Sparkles, 
  CornerDownLeft, 
  FileSpreadsheet, 
  TrendingUp, 
  User, 
  Bot, 
  Clock, 
  ChevronRight,
  ShieldCheck,
  Check,
  Laptop
} from 'lucide-react';
import { OfficeSpreadsheet } from './OfficeSpreadsheet';

export interface TurnData {
  id: 'formula' | 'highlight' | 'chart' | 'rollback';
  tabName: string;
  badge: string;
  userPrompt: string;
  agentAck: string;
  toolCall: {
    name: string;
    desc: string;
    latency: string;
  };
  agentFinal: string;
  spreadsheetScenario: 'formula' | 'highlight' | 'chart' | 'rollback';
  spreadsheetFinalStep: number;
}

const TURNS: TurnData[] = [
  {
    id: 'formula',
    tabName: '① 补数据与算公式',
    badge: '数据补全 & 公式继承',
    userPrompt: '帮我把 8 月数据补齐，并按去年的计算口径把同比增长率算出来。',
    agentAck: '收到！正在读取当前表格历史公式口径，马上写入 8 月指标并计算同比...',
    toolCall: {
      name: 'excel_patch_cells',
      desc: '写入 8 月数据并自动继承同比公式',
      latency: '24ms'
    },
    agentFinal: '搞定了！8 月实际数据已补齐（营收 612 万元），同比增长率（+9.3%）已按同列口径就地计算完成。',
    spreadsheetScenario: 'formula',
    spreadsheetFinalStep: 3
  },
  {
    id: 'highlight',
    tabName: '② 筛异常就地标黄',
    badge: '智能条件标色',
    userPrompt: '把成本超标的和未达标指标标出来，不要标红，改成浅黄色。',
    agentAck: '收到！正在检索全表各项指标阈值，定位成本超标项并执行柔和浅黄标记...',
    toolCall: {
      name: 'excel_format_cells',
      desc: '将第 6 行获客成本及第 8 行客诉率设置为浅黄警示',
      latency: '18ms'
    },
    agentFinal: '已精准定位获客成本（68元）与客诉率（2.5%），背景已改为浅黄色提醒，未触发刺眼强红色。',
    spreadsheetScenario: 'highlight',
    spreadsheetFinalStep: 3
  },
  {
    id: 'chart',
    tabName: '③ 原地插原生折线图',
    badge: 'WPS 原生图表',
    userPrompt: '根据 6 到 8 月的营收走势画一张折线图，放在表格右下方。',
    agentAck: '收到！正在提取 6~8 月营收序列，调用本地原生图表引擎原地绘制...',
    toolCall: {
      name: 'excel_add_chart',
      desc: '在活动表格中就地生成嵌入式原生矢量折线图',
      latency: '36ms'
    },
    agentFinal: '折线走势图已在表格右下方原生生成，数据源已动态绑定，您可以随时在 WPS / Excel 里双击调整细节。',
    spreadsheetScenario: 'chart',
    spreadsheetFinalStep: 3
  },
  {
    id: 'rollback',
    tabName: '④ 一键撤回交接给人',
    badge: '毫秒级快照保护',
    userPrompt: '刚才那步标黄先撤回，这个我自己微调。',
    agentAck: '收到！正在调取操作快照 #snap-8021，执行毫秒级安全回退...',
    toolCall: {
      name: 'wps_rollback',
      desc: '还原单元格初始格式，保留已有数据',
      latency: '12ms'
    },
    agentFinal: '已撤回上一步的高亮标记，单元格恢复纯净状态。文件随时由您直接接管微调。',
    spreadsheetScenario: 'rollback',
    spreadsheetFinalStep: 2
  }
];

export const InteractiveChatWorkspace: React.FC = () => {
  const [currentTurnIdx, setCurrentTurnIdx] = useState<number>(0);
  const [phase, setPhase] = useState<number>(0);
  const [typedPrompt, setTypedPrompt] = useState<string>('');
  const [isPlaying, setIsPlaying] = useState<boolean>(true);
  const [inputValue, setInputValue] = useState<string>('');
  const timerRef = useRef<NodeJS.Timeout | null>(null);

  const currentTurn = TURNS[currentTurnIdx];

  const clearAllTimers = () => {
    if (timerRef.current) clearTimeout(timerRef.current);
  };

  useEffect(() => {
    clearAllTimers();
    setPhase(0);
    setTypedPrompt('');

    const fullText = currentTurn.userPrompt;
    let charIndex = 0;

    const typeInterval = setInterval(() => {
      charIndex++;
      setTypedPrompt(fullText.slice(0, charIndex));
      if (charIndex >= fullText.length) {
        clearInterval(typeInterval);

        // Phase 1: Agent ack
        timerRef.current = setTimeout(() => {
          setPhase(1);

          // Phase 2: Tool calling
          timerRef.current = setTimeout(() => {
            setPhase(2);

            // Phase 3: Done
            timerRef.current = setTimeout(() => {
              setPhase(3);

              if (isPlaying) {
                timerRef.current = setTimeout(() => {
                  setCurrentTurnIdx((prev) => (prev + 1) % TURNS.length);
                }, 3800);
              }
            }, 1100);
          }, 700);
        }, 400);
      }
    }, 28);

    return () => {
      clearInterval(typeInterval);
      clearAllTimers();
    };
  }, [currentTurnIdx, isPlaying]);

  const handleSelectTurn = (idx: number) => {
    clearAllTimers();
    setCurrentTurnIdx(idx);
    setPhase(0);
  };

  const getSpreadsheetStep = () => {
    if (phase < 2) return 1;
    if (phase === 2) return 2;
    return currentTurn.spreadsheetFinalStep;
  };

  const handleManualSend = (text?: string) => {
    const promptToSend = text || inputValue;
    if (!promptToSend.trim()) return;

    const matchIdx = TURNS.findIndex(t => t.userPrompt.includes(promptToSend.slice(0, 4)));
    if (matchIdx !== -1) {
      handleSelectTurn(matchIdx);
    } else {
      handleSelectTurn((currentTurnIdx + 1) % TURNS.length);
    }
    setInputValue('');
  };

  return (
    <section id="interactive-demo" className="py-16 md:py-24 bg-white relative overflow-hidden">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Section Header */}
        <div className="text-center max-w-3xl mx-auto mb-10">
          <div className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-full bg-blue-50 text-blue-600 border border-blue-200/80 text-xs font-semibold mb-3">
            <Sparkles className="w-3.5 h-3.5" />
            <span>实景演示 · 你说它就干</span>
          </div>
          <h2 className="text-3xl sm:text-4xl md:text-5xl font-black text-slate-900 tracking-tight mb-4">
            你一段一段说话，<br className="sm:hidden" />
            <span className="text-blue-600">AI 收到当场把活干完</span>
          </h2>
          <p className="text-slate-600 text-sm sm:text-base leading-relaxed max-w-2xl mx-auto">
            就像坐在你身边的熟手助理，看一眼你正在打开的表格，边听你说，边在表格里把数据补全、公式算好、图表画完。
          </p>
        </div>

        {/* Floating Capsule Scenario Selector */}
        <div className="flex flex-wrap items-center justify-between gap-3 mb-6 bg-slate-50 p-2 sm:p-2.5 rounded-full border border-slate-200 shadow-sm">
          {/* Turn Tabs Pills */}
          <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
            <span className="text-xs font-semibold text-slate-400 pl-3 hidden sm:inline">
              典型任务演练：
            </span>
            {TURNS.map((turn, idx) => {
              const isSelected = idx === currentTurnIdx;
              return (
                <button
                  key={turn.id}
                  onClick={() => handleSelectTurn(idx)}
                  className={`px-3.5 sm:px-4 py-1.5 rounded-full text-xs font-semibold transition-all duration-200 cursor-pointer ${
                    isSelected
                      ? 'bg-blue-600 text-white shadow-md shadow-blue-500/25 scale-[1.02]'
                      : 'bg-white text-slate-600 hover:text-slate-900 hover:bg-slate-100 border border-slate-200/60'
                  }`}
                >
                  <span>{turn.tabName}</span>
                </button>
              );
            })}
          </div>

          {/* Autoplay Controls */}
          <div className="flex items-center gap-2 pr-2 ml-auto">
            <button
              onClick={() => setIsPlaying(!isPlaying)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-white hover:bg-slate-100 border border-slate-200 text-xs font-semibold text-slate-700 transition-colors cursor-pointer shadow-2xl"
              title={isPlaying ? '暂停自动演练' : '开启连续演练'}
            >
              {isPlaying ? (
                <>
                  <Pause className="w-3.5 h-3.5 text-amber-500 fill-amber-500" />
                  <span className="hidden sm:inline">自动演练中</span>
                </>
              ) : (
                <>
                  <Play className="w-3.5 h-3.5 text-blue-600 fill-blue-600" />
                  <span className="hidden sm:inline">继续演练</span>
                </>
              )}
            </button>
            <button
              onClick={() => handleSelectTurn(currentTurnIdx)}
              className="p-1.5 rounded-full bg-white hover:bg-slate-100 border border-slate-200 text-slate-500 hover:text-slate-800 transition-colors cursor-pointer"
              title="重播当前步骤"
            >
              <RotateCcw className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

        {/* Clean Modern Workstation Frame */}
        <div className="rounded-3xl overflow-hidden bg-slate-50 border border-slate-200/90 shadow-[0_20px_50px_-15px_rgba(15,23,42,0.08)]">
          {/* Top Window Header (macOS Style) */}
          <div className="h-10 bg-slate-100/90 px-4 flex items-center justify-between border-b border-slate-200 select-none">
            <div className="flex items-center gap-2">
              <span className="w-3 h-3 rounded-full bg-[#ff5f56]" />
              <span className="w-3 h-3 rounded-full bg-[#ffbd2e]" />
              <span className="w-3 h-3 rounded-full bg-[#27c93f]" />
              <span className="ml-3 text-xs text-slate-600 font-semibold flex items-center gap-2">
                <Laptop className="w-3.5 h-3.5 text-slate-500" />
                <span>字浮 CharFloat · 本地办公协同视窗</span>
              </span>
            </div>

            <div className="flex items-center gap-3 text-xs text-slate-500 font-medium">
              <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 text-[11px] font-semibold">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping" />
                WPS / Excel 本地连接就绪
              </span>
            </div>
          </div>

          {/* Split Screen Container */}
          <div className="grid grid-cols-1 lg:grid-cols-12 min-h-[520px]">
            {/* LEFT 5 Cols: Chat Stream & Tool Actions */}
            <div className="lg:col-span-5 bg-white p-4 sm:p-6 flex flex-col justify-between border-b lg:border-b-0 lg:border-r border-slate-200">
              {/* Messages Container */}
              <div className="space-y-4 overflow-y-auto pr-1">
                {/* 1. User Message */}
                <div className="flex items-start gap-3">
                  <div className="w-8 h-8 rounded-full bg-blue-100 text-blue-700 font-bold flex items-center justify-center shrink-0 mt-0.5 text-xs">
                    你
                  </div>
                  <div className="flex-1">
                    <div className="text-[11px] font-semibold text-slate-400 mb-1">
                      发送指令
                    </div>
                    <div className="p-3.5 rounded-2xl rounded-tl-sm bg-blue-50/80 border border-blue-200/80 shadow-sm">
                      <p className="text-sm sm:text-base font-bold text-slate-900 leading-relaxed">
                        {typedPrompt}
                        {phase === 0 && (
                          <span className="inline-block w-1.5 h-4 ml-1 bg-blue-600 animate-pulse align-middle" />
                        )}
                      </p>
                    </div>
                  </div>
                </div>

                {/* 2. AI Acknowledgement & Tool Execution */}
                {phase >= 1 && (
                  <div className="flex items-start gap-3 animate-fadeIn">
                    <div className="w-8 h-8 rounded-full bg-gradient-to-tr from-blue-600 to-indigo-600 text-white flex items-center justify-center shrink-0 mt-0.5 shadow-md shadow-blue-500/20">
                      <Zap className="w-4 h-4 fill-white" />
                    </div>
                    <div className="flex-1">
                      <div className="text-[11px] font-semibold text-slate-400 mb-1">
                        Office 助手
                      </div>

                      {/* AI "收到" Bubble */}
                      <div className="p-3.5 rounded-2xl rounded-tl-sm bg-white border border-slate-200 text-xs sm:text-sm text-slate-800 leading-relaxed shadow-sm">
                        <p>{currentTurn.agentAck}</p>
                      </div>

                      {/* 3. Action Execution Tag (Clean & No AI Nerd Gimmicks) */}
                      {phase >= 2 && (
                        <div className="mt-3 p-3 rounded-xl bg-slate-50 border border-slate-200/90 space-y-2 animate-fadeIn shadow-sm">
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-2 text-xs">
                              <span className="w-2 h-2 rounded-full bg-blue-600 animate-pulse" />
                              <span className="font-bold text-slate-900">执行动作：</span>
                              <span className="text-slate-600">{currentTurn.toolCall.desc}</span>
                            </div>
                            <span className="text-[10px] font-mono text-slate-500 bg-white px-1.5 py-0.5 rounded border border-slate-200">
                              {phase === 2 ? '处理中...' : currentTurn.toolCall.latency}
                            </span>
                          </div>

                          <div className="flex items-center justify-between text-[11px] text-slate-500 pt-1 border-t border-slate-200/60">
                            <span className="font-mono text-slate-400 text-[10px]">
                              调用: {currentTurn.toolCall.name}
                            </span>
                            {phase >= 3 ? (
                              <span className="text-emerald-700 flex items-center gap-1 font-bold text-[11px]">
                                <Check className="w-3.5 h-3.5" /> 已完成回写并保存
                              </span>
                            ) : (
                              <span className="text-blue-600 flex items-center gap-1 text-[11px] animate-pulse font-medium">
                                正在向 WPS 提交修改...
                              </span>
                            )}
                          </div>
                        </div>
                      )}

                      {/* 4. Agent Final Response */}
                      {phase >= 3 && (
                        <div className="mt-3 p-3.5 rounded-2xl bg-emerald-50 border border-emerald-200/80 text-xs sm:text-sm text-emerald-900 leading-relaxed animate-fadeIn">
                          <div className="flex items-start gap-2">
                            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                            <p className="font-medium">{currentTurn.agentFinal}</p>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>

              {/* Bottom Clean Input Capsule (Doubao Work style) */}
              <div className="mt-4 pt-3 border-t border-slate-100">
                {/* Quick Prompts Capsules */}
                <div className="flex items-center gap-1.5 mb-2.5 overflow-x-auto pb-1 text-[11px]">
                  <span className="text-slate-400 shrink-0">点击体验：</span>
                  {TURNS.map((turn, i) => (
                    <button
                      key={turn.id}
                      onClick={() => handleSelectTurn(i)}
                      className="shrink-0 px-2.5 py-1 rounded-full bg-slate-100 hover:bg-blue-50 text-slate-600 hover:text-blue-700 border border-slate-200/60 transition-colors cursor-pointer font-medium"
                    >
                      {turn.userPrompt.slice(0, 11)}...
                    </button>
                  ))}
                </div>

                {/* Input Bar Pill */}
                <form 
                  onSubmit={(e) => {
                    e.preventDefault();
                    handleManualSend();
                  }}
                  className="flex items-center gap-2 bg-slate-100/90 rounded-full px-4 py-1.5 border border-slate-200 focus-within:border-blue-500 focus-within:bg-white shadow-sm transition-all"
                >
                  <input
                    type="text"
                    value={inputValue}
                    onChange={(e) => setInputValue(e.target.value)}
                    placeholder="说一句话，如：帮我把表格8月数据补齐并计算同比..."
                    className="flex-1 bg-transparent py-1.5 text-xs sm:text-sm text-slate-800 placeholder-slate-400 focus:outline-none"
                  />
                  <button
                    type="submit"
                    className="flex items-center justify-center w-7 h-7 rounded-full bg-blue-600 hover:bg-blue-700 text-white transition-all shadow-md shadow-blue-500/20 shrink-0 cursor-pointer"
                    title="发送"
                  >
                    <Send className="w-3.5 h-3.5" />
                  </button>
                </form>
              </div>
            </div>

            {/* RIGHT 7 Cols: Real-time Live Office Simulation */}
            <div className="lg:col-span-7 bg-[#f8fafc] p-3 sm:p-5 flex flex-col justify-center">
              <div className="flex-1 min-h-[460px] flex flex-col justify-center">
                <OfficeSpreadsheet
                  scenario={currentTurn.spreadsheetScenario}
                  step={getSpreadsheetStep()}
                />
              </div>

              {/* Status footer bar */}
              <div className="mt-3 px-4 py-2.5 rounded-2xl bg-white border border-slate-200 shadow-sm flex flex-wrap items-center justify-between text-xs text-slate-600 gap-2">
                <div className="flex items-center gap-2">
                  <ShieldCheck className="w-4 h-4 text-emerald-600" />
                  <span>安全与快照：每一步改动自动建立快照，说一句“撤销”即可秒级复原</span>
                </div>
                <div className="text-[11px] text-slate-400 font-medium">
                  零格式破坏 · 纯本地直驱
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
};
