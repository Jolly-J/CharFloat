import { Sparkles, ArrowUp, Download } from 'lucide-react';

interface FooterProps {
  onOpenQuickStart: () => void;
  onOpenDownload?: (tab?: 'mac' | 'win' | 'cli') => void;
}

export const Footer: React.FC<FooterProps> = ({ onOpenQuickStart, onOpenDownload }) => {
  const scrollToTop = () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  return (
    <footer className="bg-white text-slate-600 border-t border-slate-200 relative overflow-hidden">
      {/* Big Closing CTA Strip (Doubao Work Style) */}
      <div className="max-w-5xl mx-auto px-4 sm:px-6 lg:px-8 py-20 text-center">
        <h2 className="text-3xl sm:text-5xl font-black text-slate-900 tracking-tight mb-5">
          给你的 AI，开个 Office 外挂
        </h2>
        <p className="text-base sm:text-lg text-slate-600 max-w-2xl mx-auto mb-10 leading-relaxed font-normal">
          不用上传下载，不再生成一份。现在就让你最常用的 AI 直接进入桌面 WPS 与 Excel，跟你一起改同一份文件。
        </p>

        {/* Big Action Button Group */}
        <div className="flex flex-col sm:flex-row items-center justify-center gap-3.5 max-w-md mx-auto mb-10">
          <button
            onClick={() => onOpenDownload ? onOpenDownload() : onOpenQuickStart()}
            className="w-full sm:w-auto flex-1 flex items-center justify-center gap-2.5 px-8 py-4 rounded-full font-bold text-white bg-blue-600 hover:bg-blue-700 shadow-xl shadow-blue-500/25 transition-all hover:scale-[1.02] active:scale-[0.98] cursor-pointer text-base"
          >
            <Download className="w-5 h-5" />
            <span>免费下载电脑版</span>
          </button>

          <button
            onClick={onOpenQuickStart}
            className="w-full sm:w-auto flex items-center justify-center gap-2 px-6 py-4 rounded-full font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 transition-all hover:scale-[1.02] active:scale-[0.98] cursor-pointer text-sm"
          >
            <Sparkles className="w-4 h-4 text-blue-600" />
            <span>开发者配置</span>
          </button>
        </div>

        {/* 4 Virtues Pills */}
        <div className="flex flex-wrap items-center justify-center gap-2.5 sm:gap-4 text-xs sm:text-sm font-bold text-slate-700">
          <span className="px-3.5 py-1.5 rounded-full bg-slate-100 border border-slate-200/80">
            当前文件直接改
          </span>
          <span className="px-3.5 py-1.5 rounded-full bg-slate-100 border border-slate-200/80">
            边做边看所见即所得
          </span>
          <span className="px-3.5 py-1.5 rounded-full bg-slate-100 border border-slate-200/80">
            原格式公式 100% 继承
          </span>
          <span className="px-3.5 py-1.5 rounded-full bg-slate-100 border border-slate-200/80">
            安全快照随时撤回
          </span>
        </div>
      </div>

      {/* Standard Bottom Credits Bar */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-500">
        <div className="flex items-center gap-2.5">
          <img 
            src="/logo.png" 
            alt="字浮 CharFloat" 
            className="w-5 h-5 rounded-md object-contain"
          />
          <span className="font-bold text-slate-900">字浮 CharFloat</span>
          <span>© 2026. 面向 Microsoft Office 与 WPS Office 的本地桌面桥梁。</span>
        </div>

        <div className="flex items-center gap-6">
          <button 
            onClick={scrollToTop}
            className="flex items-center gap-1 text-slate-500 hover:text-slate-900 transition-colors cursor-pointer"
          >
            <span>回到顶部</span>
            <ArrowUp className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </footer>
  );
};
