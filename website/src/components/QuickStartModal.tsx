import React, { useState } from 'react';
import { X, Copy, Check, Sparkles, MessageSquare, Bot, ArrowRight, ShieldCheck, Terminal } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface QuickStartModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const QuickStartModal: React.FC<QuickStartModalProps> = ({ isOpen, onClose }) => {
  const [copied, setCopied] = useState(false);

  const AI_AUTO_CONFIG_PROMPT = `请帮我把「字浮 CharFloat」配置为你的 MCP 扩展工具，让我可以直接操控本地 WPS 与 Office。

MCP 服务配置信息如下：
服务名称: charfloat
启动命令: npx -y office-agent-bridge

请识别当前开发环境（如 Cursor、Windsurf、VS Code、Claude Code 等），自动将上述配置写入你的 MCP 配置文件（例如 .cursor/mcp.json 或对应客户端设置）中并激活。`;

  const handleCopy = () => {
    navigator.clipboard.writeText(AI_AUTO_CONFIG_PROMPT);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  return (
    <AnimatePresence>
      {isOpen && (
        <motion.div 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 0.22, ease: [0.16, 1, 0.3, 1] }}
          className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-slate-900/60 backdrop-blur-md overflow-y-auto"
          onClick={onClose}
        >
          <motion.div 
            initial={{ opacity: 0, scale: 0.94, y: 16 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.94, y: 16 }}
            transition={{ duration: 0.25, ease: [0.16, 1, 0.3, 1] }}
            className="w-full max-w-2xl bg-white border border-slate-200/90 rounded-3xl shadow-2xl text-slate-800 my-auto max-h-[90vh] flex flex-col overflow-hidden"
            onClick={(e) => e.stopPropagation()}
          >
            {/* Modal Header (Fixed Top) */}
            <div className="px-7 py-5 sm:px-8 sm:py-5.5 border-b border-slate-100 flex items-center justify-between bg-slate-50/80 shrink-0">
              <div className="flex items-center gap-3.5 sm:gap-4">
                <div className="w-11 h-11 sm:w-12 sm:h-12 rounded-2xl bg-gradient-to-br from-blue-50 to-indigo-50/80 border border-blue-100 p-2 flex items-center justify-center shrink-0 shadow-xs">
                  <img src="/logo.png" alt="字浮 Logo" className="w-full h-full object-contain select-none" />
                </div>
                <div>
                  <h3 className="text-base sm:text-lg font-black text-slate-900 tracking-tight leading-snug">
                    让 AI 为自己配置「字浮 MCP 插件」
                  </h3>
                  <p className="text-xs sm:text-[13px] text-slate-500 font-normal mt-0.5 leading-relaxed">
                    无需手动翻看文档写配置，一键复制提示词发给你的 AI 即可
                  </p>
                </div>
              </div>
              <motion.button 
                whileHover={{ scale: 1.1, rotate: 90 }}
                whileTap={{ scale: 0.9 }}
                transition={{ duration: 0.2 }}
                onClick={onClose}
                className="p-2 rounded-full text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 transition-colors cursor-pointer shrink-0 ml-2"
                aria-label="关闭弹窗"
              >
                <X className="w-5 h-5" />
              </motion.button>
            </div>

            {/* Modal Scrollable Body */}
            <div className="p-7 sm:p-8 overflow-y-auto flex-1 space-y-6">
              {/* 1. Copyable AI Prompt Card */}
              <div className="space-y-2.5">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                    <MessageSquare className="w-4 h-4 text-blue-600" />
                    <span>发给 AI 的一键自动配置指令：</span>
                  </span>
                  <span className="text-[11px] text-slate-400">支持 Cursor / Windsurf / Claude Code / Cline 等</span>
                </div>

                <div className="relative rounded-2xl bg-slate-900 border border-slate-800 p-4 font-mono text-xs text-slate-200 shadow-inner group">
                  <pre className="text-xs text-slate-200/90 leading-relaxed whitespace-pre-wrap font-sans select-all pr-4 font-normal">
                    {AI_AUTO_CONFIG_PROMPT}
                  </pre>
                </div>

                {/* Big Prominent Copy Button */}
                <motion.button
                  whileHover={{ scale: 1.015 }}
                  whileTap={{ scale: 0.985 }}
                  onClick={handleCopy}
                  className={`w-full py-3.5 px-6 rounded-2xl font-bold text-sm flex items-center justify-center gap-2 transition-all shadow-md cursor-pointer ${
                    copied
                      ? 'bg-emerald-600 text-white shadow-emerald-500/25'
                      : 'bg-blue-600 hover:bg-blue-700 text-white shadow-blue-500/25'
                  }`}
                >
                  {copied ? (
                    <>
                      <Check className="w-4 h-4 stroke-[3]" />
                      <span>已复制提示词！直接发给你常用的 AI 助手即可</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-4 h-4" />
                      <span>一键复制配置提示词（发给你的 AI）</span>
                    </>
                  )}
                </motion.button>
              </div>

              {/* 2. 3-Step Simple Flow */}
              <div className="pt-2 border-t border-slate-100">
                <p className="text-xs font-bold text-slate-800 mb-3 flex items-center gap-1.5">
                  <Sparkles className="w-3.5 h-3.5 text-blue-600" />
                  <span>3 步极速接入流程：</span>
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                  <div className="p-3.5 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
                    <div className="font-bold text-slate-900 mb-1 flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-700 text-[11px] font-black flex items-center justify-center">1</span>
                      <span>复制上方指令</span>
                    </div>
                    <p className="text-slate-500 text-[11.5px] leading-relaxed">
                      点击按钮一键复制这段专门调教好的 AI 配置提示词。
                    </p>
                  </div>

                  <div className="p-3.5 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
                    <div className="font-bold text-slate-900 mb-1 flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-700 text-[11px] font-black flex items-center justify-center">2</span>
                      <span>发给你的 AI</span>
                    </div>
                    <p className="text-slate-500 text-[11.5px] leading-relaxed">
                      直接粘贴发送给 Cursor、Windsurf 或 Claude 等任意智能体。
                    </p>
                  </div>

                  <div className="p-3.5 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
                    <div className="font-bold text-slate-900 mb-1 flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-700 text-[11px] font-black flex items-center justify-center">3</span>
                      <span>AI 自动完成配置</span>
                    </div>
                    <p className="text-slate-500 text-[11.5px] leading-relaxed">
                      AI 将自动修改配置并启用 MCP，当场获得操作 Office 的能力！
                    </p>
                  </div>
                </div>
              </div>

              {/* 3. Automatic Desktop Client Tip */}
              <div className="p-3.5 rounded-2xl bg-blue-50/70 border border-blue-100 flex items-start gap-2.5 text-xs text-blue-900">
                <ShieldCheck className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
                <p className="leading-relaxed">
                  <strong>更省心的方式</strong>：安装字浮桌面客户端，客户端打开时会自动一键配置并激活本机已安装的 AI 软件，完全无需手动复制。
                </p>
              </div>
            </div>

            {/* Modal Footer (Fixed Bottom) */}
            <div className="px-7 py-4 sm:px-8 bg-slate-50/90 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500 shrink-0">
              <span>完全免费 · 标准通用 MCP 协议</span>
              <motion.button
                whileHover={{ scale: 1.03 }}
                whileTap={{ scale: 0.97 }}
                onClick={onClose}
                className="px-5 py-2 rounded-full bg-slate-900 hover:bg-slate-800 text-white font-bold transition-colors cursor-pointer"
              >
                完成
              </motion.button>
            </div>
          </motion.div>
        </motion.div>
      )}
    </AnimatePresence>
  );
};
