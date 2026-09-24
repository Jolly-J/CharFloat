import React, { useState, useEffect, useRef } from 'react';
import { 
  X, 
  Download, 
  Terminal, 
  Check, 
  Copy, 
  ShieldCheck, 
  Info,
  ChevronRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { AppleLogo, WindowsLogo } from './BrandIcons';
import { getDetectedOS } from '../utils/os';

interface DownloadModalProps {
  isOpen: boolean;
  onClose: () => void;
  onOpenQuickStart: () => void;
  initialTab?: 'mac' | 'win' | 'cli';
}

export const DownloadModal: React.FC<DownloadModalProps> = ({ 
  isOpen, 
  onClose, 
  onOpenQuickStart,
  initialTab
}) => {
  const [activeTab, setActiveTab] = useState<'mac' | 'win' | 'cli'>(() => initialTab || getDetectedOS());
  const [copied, setCopied] = useState(false);
  const contentRef = useRef<HTMLDivElement>(null);
  const [height, setHeight] = useState<number | 'auto'>('auto');

  useEffect(() => {
    if (initialTab) setActiveTab(initialTab);
  }, [initialTab, isOpen]);

  // Measure content height dynamically to animate growth/shrinkage
  useEffect(() => {
    if (!contentRef.current) return;
    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        if (entry.contentRect) {
          setHeight(entry.contentRect.height);
        }
      }
    });

    observer.observe(contentRef.current);
    return () => observer.disconnect();
  }, [isOpen, activeTab]);

  const cliCommand = 'npx office-agent-bridge setup';

  const handleCopyCli = () => {
    navigator.clipboard.writeText(cliCommand);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <AnimatePresence>
      {isOpen && (
        <motion.div 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 0.25, ease: [0.16, 1, 0.3, 1] }}
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm"
          onClick={onClose}
        >
          <motion.div 
            initial={{ opacity: 0, scale: 0.92, y: 16 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.92, y: 16 }}
            transition={{ duration: 0.28, ease: [0.16, 1, 0.3, 1] }}
            className="w-full max-w-xl bg-white border border-slate-200 rounded-3xl shadow-2xl overflow-hidden text-slate-800"
            onClick={(e) => e.stopPropagation()}
          >
            {/* Modal Header with 打招呼.png Brand Image */}
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/70">
              <div className="flex items-center gap-3">
                <img 
                  src="/character/打招呼.png" 
                  alt="字浮" 
                  className="w-10 h-10 object-contain select-none shrink-0 drop-shadow-xs" 
                />
                <div>
                  <h3 className="font-bold text-base text-slate-900">获取 字浮 CharFloat 电脑版</h3>
                  <p className="text-xs text-slate-500">正式版 v2.2.0 · 免费纯本地运行</p>
                </div>
              </div>
              <motion.button 
                whileHover={{ scale: 1.1, rotate: 90 }}
                whileTap={{ scale: 0.9 }}
                transition={{ duration: 0.2 }}
                onClick={onClose}
                className="p-1.5 rounded-full text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </motion.button>
            </div>

            {/* Platform Selector Tabs with Smooth Sliding Line Indicator */}
            <div className="flex border-b border-slate-100 bg-white px-6 pt-3 gap-2 relative">
              <button
                onClick={() => setActiveTab('mac')}
                className={`relative flex items-center gap-2 pb-3 px-3 text-xs sm:text-sm font-semibold transition-colors cursor-pointer ${
                  activeTab === 'mac'
                    ? 'text-blue-600'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                <AppleLogo className="w-4 h-4" />
                <span>macOS 版</span>
                {activeTab === 'mac' && (
                  <motion.div 
                    layoutId="downloadActiveTabUnderline"
                    className="absolute bottom-0 left-0 right-0 h-[2px] bg-blue-600 rounded-full"
                    transition={{ type: "spring", stiffness: 500, damping: 35 }}
                  />
                )}
              </button>

              <button
                onClick={() => setActiveTab('win')}
                className={`relative flex items-center gap-2 pb-3 px-3 text-xs sm:text-sm font-semibold transition-colors cursor-pointer ${
                  activeTab === 'win'
                    ? 'text-blue-600'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                <WindowsLogo className="w-4 h-4" />
                <span>Windows 版</span>
                {activeTab === 'win' && (
                  <motion.div 
                    layoutId="downloadActiveTabUnderline"
                    className="absolute bottom-0 left-0 right-0 h-[2px] bg-blue-600 rounded-full"
                    transition={{ type: "spring", stiffness: 500, damping: 35 }}
                  />
                )}
              </button>

              <button
                onClick={() => setActiveTab('cli')}
                className={`relative flex items-center gap-2 pb-3 px-3 text-xs sm:text-sm font-semibold transition-colors cursor-pointer ${
                  activeTab === 'cli'
                    ? 'text-blue-600'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                <Terminal className="w-4 h-4" />
                <span>开发者 CLI</span>
                {activeTab === 'cli' && (
                  <motion.div 
                    layoutId="downloadActiveTabUnderline"
                    className="absolute bottom-0 left-0 right-0 h-[2px] bg-blue-600 rounded-full"
                    transition={{ type: "spring", stiffness: 500, damping: 35 }}
                  />
                )}
              </button>
            </div>

            {/* Height Animated Container: Smooth growth and shrinkage */}
            <motion.div
              animate={{ height: height || 'auto' }}
              transition={{ type: "spring", stiffness: 450, damping: 32 }}
              className="overflow-hidden"
            >
              <div ref={contentRef} className="p-6 space-y-5">
                <AnimatePresence mode="wait" initial={false}>
                  {activeTab === 'mac' && (
                    <motion.div 
                      key="mac"
                      initial={{ opacity: 0, y: 8 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -8 }}
                      transition={{ duration: 0.18, ease: [0.16, 1, 0.3, 1] }}
                      className="space-y-4"
                    >
                      <div className="p-4 rounded-2xl bg-slate-50 border border-slate-200/80 flex items-center justify-between gap-4">
                        <div>
                          <div className="flex items-center gap-2 font-bold text-slate-900 text-sm">
                            <AppleLogo className="w-4 h-4 text-slate-700" />
                            <span>字浮-CharFloat-2.2.0-mac-arm64.zip</span>
                          </div>
                          <p className="text-xs text-slate-500 mt-1">
                            适用系统: macOS 12+ (Apple Silicon M1/M2/M3/M4 & Intel)
                          </p>
                        </div>
                        <motion.a
                          href="/release/2.2.0/mac/字浮-CharFloat-2.2.0-mac-arm64.zip"
                          download
                          whileHover={{ scale: 1.03 }}
                          whileTap={{ scale: 0.97 }}
                          className="px-4 py-2 rounded-full bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-md shadow-blue-500/20 shrink-0 cursor-pointer"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>立即下载</span>
                        </motion.a>
                      </div>

                      <div className="p-3.5 rounded-xl bg-blue-50/50 border border-blue-100 text-xs text-slate-600 space-y-1">
                        <div className="font-semibold text-blue-900 flex items-center gap-1.5">
                          <Info className="w-3.5 h-3.5 text-blue-600" />
                          <span>安装提示：</span>
                        </div>
                        <p>1. 解压后将应用拖入 <code className="text-slate-800 font-semibold">Applications (应用程序)</code> 目录。</p>
                        <p>2. 启动后自动常驻系统菜单栏，随时与 WPS 和 Excel 自动联通。</p>
                      </div>
                    </motion.div>
                  )}

                  {activeTab === 'win' && (
                    <motion.div 
                      key="win"
                      initial={{ opacity: 0, y: 8 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -8 }}
                      transition={{ duration: 0.18, ease: [0.16, 1, 0.3, 1] }}
                      className="space-y-4"
                    >
                      <div className="p-4 rounded-2xl bg-slate-50 border border-slate-200/80 flex items-center justify-between gap-4">
                        <div>
                          <div className="flex items-center gap-2 font-bold text-slate-900 text-sm">
                            <WindowsLogo className="w-4 h-4 text-slate-700" />
                            <span>字浮-CharFloat-2.2.0-win-x64.zip</span>
                          </div>
                          <p className="text-xs text-slate-500 mt-1">
                            适用系统: Windows 10 / Windows 11 (64 位)
                          </p>
                        </div>
                        <motion.a
                          href="/release/2.2.0/win/字浮-CharFloat-2.2.0-win-x64.zip"
                          download
                          whileHover={{ scale: 1.03 }}
                          whileTap={{ scale: 0.97 }}
                          className="px-4 py-2 rounded-full bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-md shadow-blue-500/20 shrink-0 cursor-pointer"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>立即下载</span>
                        </motion.a>
                      </div>

                      <div className="p-3.5 rounded-xl bg-blue-50/50 border border-blue-100 text-xs text-slate-600 space-y-1">
                        <div className="font-semibold text-blue-900 flex items-center gap-1.5">
                          <Info className="w-3.5 h-3.5 text-blue-600" />
                          <span>运行提示：</span>
                        </div>
                        <p>1. 解压到本地目录，双击主程序即可运行。</p>
                        <p>2. 兼容原生 Microsoft Office 2016/2019/365 及 WPS Office 2019+。</p>
                      </div>
                    </motion.div>
                  )}

                  {activeTab === 'cli' && (
                    <motion.div 
                      key="cli"
                      initial={{ opacity: 0, y: 8 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -8 }}
                      transition={{ duration: 0.18, ease: [0.16, 1, 0.3, 1] }}
                      className="space-y-4"
                    >
                      <p className="text-xs text-slate-600 leading-relaxed">
                        无需下载完整桌面壳层，可在终端直接执行命令快速启动与配置：
                      </p>
                      <div className="p-3.5 rounded-xl bg-slate-100 border border-slate-200 font-mono text-xs flex items-center justify-between">
                        <code className="text-blue-700 font-semibold">{cliCommand}</code>
                        <motion.button
                          whileHover={{ scale: 1.04 }}
                          whileTap={{ scale: 0.96 }}
                          onClick={handleCopyCli}
                          className="flex items-center gap-1.5 px-3 py-1 rounded-lg bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 transition-colors cursor-pointer shadow-xs"
                        >
                          {copied ? (
                            <>
                              <Check className="w-3.5 h-3.5 text-emerald-600" />
                              <span className="text-emerald-600 font-bold">已复制</span>
                            </>
                          ) : (
                            <>
                              <Copy className="w-3.5 h-3.5" />
                              <span>复制</span>
                            </>
                          )}
                        </motion.button>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>

                {/* Compatibility list footer */}
                <div className="pt-4 border-t border-slate-100 flex flex-wrap items-center justify-between text-xs text-slate-500 gap-3">
                  <div className="flex items-center gap-1.5">
                    <ShieldCheck className="w-4 h-4 text-emerald-600" />
                    <span>纯本地架构 · 零数据上云 · 兼容 WPS & Office</span>
                  </div>
                  <motion.button
                    whileHover={{ x: 2 }}
                    onClick={() => {
                      onClose();
                      onOpenQuickStart();
                    }}
                    className="text-blue-600 hover:text-blue-700 font-semibold flex items-center gap-0.5 cursor-pointer"
                  >
                    <span>查看 MCP 配置</span>
                    <ChevronRight className="w-3.5 h-3.5" />
                  </motion.button>
                </div>
              </div>
            </motion.div>
          </motion.div>
        </motion.div>
      )}
    </AnimatePresence>
  );
};
