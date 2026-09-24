import React, { useState, useEffect, useRef } from 'react';
import { 
  Download, 
  ChevronDown
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { AppleLogo, WindowsLogo } from './BrandIcons';

interface HeroProps {
  onOpenQuickStart?: () => void;
  onOpenDownload?: (tab?: 'mac' | 'win' | 'cli') => void;
}

export const Hero: React.FC<HeroProps> = ({ 
  onOpenDownload
}) => {
  const [downloadMenuOpen, setDownloadMenuOpen] = useState(false);
  const [isPlayingVideo, setIsPlayingVideo] = useState(false);
  const videoRef = useRef<HTMLVideoElement>(null);

  // Show cover image first for 2.5s, then smoothly crossfade into autoplay video
  useEffect(() => {
    const timer = setTimeout(() => {
      setIsPlayingVideo(true);
      if (videoRef.current) {
        videoRef.current.play().catch(() => {
          // Autoplay fallback
        });
      }
    }, 2500);

    return () => clearTimeout(timer);
  }, []);

  return (
    <section id="download" className="relative pt-32 pb-16 md:pt-36 md:pb-24 overflow-hidden">
      {/* Sky Blue Gradient Starting from the Very Top (Logo-matched sky blue) */}
      <div 
        className="absolute top-0 left-0 right-0 h-[720px] pointer-events-none -z-10"
        style={{
          background: 'linear-gradient(180deg, #cae0f9 0%, #d8ebfc 25%, #ebf5fe 60%, #f8fafc 100%)'
        }}
      />

      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        {/* Brand App Character: Friendly companion beside headline without pushing showcase off-screen */}
        <motion.div 
          initial={{ opacity: 0, scale: 0.85, y: 15 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          transition={{ duration: 0.8, ease: [0.16, 1, 0.3, 1] }}
          className="flex justify-center mb-4"
        >
          <motion.img 
            src="/character-work.png" 
            alt="字浮 CharFloat" 
            animate={{ y: [0, -8, 0] }}
            transition={{ duration: 4.5, repeat: Infinity, ease: "easeInOut" }}
            className="w-24 h-24 sm:w-28 sm:h-28 object-contain select-none drop-shadow-md cursor-pointer"
            whileHover={{ scale: 1.08, rotate: [-1, 2, -1, 0] }}
          />
        </motion.div>

        {/* Main Title - Scaled down to refined size with uniform line spacing */}
        <motion.h1 
          initial={{ opacity: 0, y: 24 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.75, delay: 0.1, ease: [0.16, 1, 0.3, 1] }}
          className="text-3xl sm:text-5xl lg:text-[52px] font-black text-slate-900 tracking-tight leading-tight mb-6"
        >
          <span>给你的 AI，开个 Office 外挂。</span>
          <span className="text-2xl sm:text-4xl lg:text-[42px] text-blue-600 block mt-3.5">你说，它当场改。</span>
        </motion.h1>

        {/* Crisp Subtitle - Balanced lines with uniform rhythm */}
        <motion.div 
          initial={{ opacity: 0, y: 18 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.75, delay: 0.2, ease: [0.16, 1, 0.3, 1] }}
          className="text-base sm:text-lg text-slate-600 max-w-3xl mx-auto mb-6 space-y-2 font-normal leading-relaxed"
        >
          <p>ChatGPT、Workbuddy、豆包、千问……你用谁都行。</p>
          <p>装上字浮，让它们直接进入当前正打开的 WPS 与 Excel。不用上传下载，就在眼前这一份里跟你一起改完。</p>
        </motion.div>

        {/* Core Differentiation Badges */}
        <motion.div
          initial={{ opacity: 0, y: 12 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, delay: 0.25, ease: [0.16, 1, 0.3, 1] }}
          className="flex flex-wrap items-center justify-center gap-2 sm:gap-2.5 text-xs sm:text-[13px] font-semibold text-slate-600 mb-8"
        >
          <span className="px-3 py-1 rounded-full bg-white/80 border border-slate-200/80 shadow-xs">不用上传下载</span>
          <span className="text-slate-300">·</span>
          <span className="px-3 py-1 rounded-full bg-white/80 border border-slate-200/80 shadow-xs">不另存多余副本</span>
          <span className="text-slate-300">·</span>
          <span className="px-3 py-1 rounded-full bg-white/80 border border-slate-200/80 shadow-xs">原文档增量修改</span>
          <span className="text-slate-300">·</span>
          <span className="px-3 py-1 rounded-full bg-white/80 border border-slate-200/80 shadow-xs">改到哪儿看到哪儿</span>
        </motion.div>

        {/* Main Download Button with System Dropdown (Doubao Work 1:1) with Spring Warp Effects */}
        <motion.div 
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.75, delay: 0.3, ease: [0.16, 1, 0.3, 1] }}
          className="relative inline-block mb-12"
        >
          <div className="flex items-center">
            <motion.button
              onClick={() => onOpenDownload ? onOpenDownload() : undefined}
              whileHover={{ scale: 1.025, boxShadow: '0 20px 35px -8px rgba(37,99,235,0.35)' }}
              whileTap={{ scale: 0.975 }}
              transition={{ type: "spring", stiffness: 450, damping: 25 }}
              className="flex items-center justify-center gap-2.5 px-8 py-4 rounded-full font-bold text-white bg-blue-600 hover:bg-blue-700 shadow-xl shadow-blue-500/25 cursor-pointer text-base"
            >
              <Download className="w-5 h-5" />
              <span>下载字浮电脑端</span>
            </motion.button>

            <motion.button
              onClick={() => setDownloadMenuOpen(!downloadMenuOpen)}
              whileHover={{ scale: 1.05 }}
              whileTap={{ scale: 0.95 }}
              transition={{ type: "spring", stiffness: 450, damping: 25 }}
              className="ml-1 p-4 rounded-full bg-blue-600 hover:bg-blue-700 text-white shadow-xl shadow-blue-500/25 cursor-pointer"
              title="选择系统版本"
            >
              <motion.div
                animate={{ rotate: downloadMenuOpen ? 180 : 0 }}
                transition={{ duration: 0.3, ease: [0.16, 1, 0.3, 1] }}
              >
                <ChevronDown className="w-4 h-4" />
              </motion.div>
            </motion.button>
          </div>

          {/* Download System Select Dropdown with Warp-Speed Spring */}
          <AnimatePresence>
            {downloadMenuOpen && (
              <motion.div 
                initial={{ opacity: 0, y: -10, scale: 0.92 }}
                animate={{ opacity: 1, y: 0, scale: 1 }}
                exit={{ opacity: 0, y: -10, scale: 0.92 }}
                transition={{ duration: 0.25, ease: [0.16, 1, 0.3, 1] }}
                className="absolute top-full left-1/2 -translate-x-1/2 mt-2 w-56 bg-white/95 backdrop-blur-2xl border border-slate-200/90 rounded-2xl shadow-2xl py-2 z-30 text-left text-xs font-semibold text-slate-700 origin-top"
              >
                <button
                  onClick={() => {
                    setDownloadMenuOpen(false);
                    if (onOpenDownload) onOpenDownload('mac');
                  }}
                  className="w-full px-4 py-2.5 hover:bg-slate-50 flex items-center justify-between cursor-pointer transition-colors"
                >
                  <div className="flex items-center gap-2">
                    <AppleLogo className="w-4 h-4" />
                    <span>Mac 版 (Apple Silicon)</span>
                  </div>
                  <span className="text-[10px] text-blue-600 font-bold bg-blue-50 px-1.5 py-0.5 rounded">推荐</span>
                </button>

                <button
                  onClick={() => {
                    setDownloadMenuOpen(false);
                    if (onOpenDownload) onOpenDownload('mac');
                  }}
                  className="w-full px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 cursor-pointer transition-colors"
                >
                  <AppleLogo className="w-4 h-4" />
                  <span>Mac 版 (Intel 芯片)</span>
                </button>

                <div className="border-t border-slate-100 my-1" />

                <button
                  onClick={() => {
                    setDownloadMenuOpen(false);
                    if (onOpenDownload) onOpenDownload('win');
                  }}
                  className="w-full px-4 py-2.5 hover:bg-slate-50 flex items-center justify-between cursor-pointer transition-colors"
                >
                  <div className="flex items-center gap-2">
                    <WindowsLogo className="w-4 h-4" />
                    <span>Windows 版 (64 位)</span>
                  </div>
                </button>
              </motion.div>
            )}
          </AnimatePresence>
        </motion.div>

        {/* Hero App Showcase Video Player with Initial Cover Poster Delay */}
        <motion.div 
          initial={{ opacity: 0, y: 36, scale: 0.97 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          transition={{ duration: 0.85, delay: 0.35, ease: [0.16, 1, 0.3, 1] }}
          whileHover={{ y: -4, scale: 1.008 }}
          className="relative rounded-3xl overflow-hidden bg-slate-900/5 border border-slate-200/90 shadow-[0_24px_60px_-15px_rgba(0,106,255,0.14)] group max-w-5xl mx-auto transition-all aspect-[16/9] w-full"
        >
          {/* 1. Base Cover Poster Image */}
          <img
            src="/hero-showcase.png"
            alt="字浮 CharFloat · 给你的 AI 开个 Office 外挂，你说它当场改"
            className="absolute inset-0 w-full h-full object-cover rounded-3xl select-none"
            loading="eager"
          />

          {/* 2. Seamless Autoplay Showcase Video */}
          <video
            ref={videoRef}
            src="/hero-showcase.mp4"
            muted
            loop
            playsInline
            preload="auto"
            className={`absolute inset-0 w-full h-full object-cover rounded-3xl transition-opacity duration-1000 ease-out ${
              isPlayingVideo ? 'opacity-100' : 'opacity-0 pointer-events-none'
            }`}
          />

          {/* 3. Subtle Inner Glass Ring */}
          <div className="absolute inset-0 rounded-3xl ring-1 ring-inset ring-black/5 pointer-events-none" />
        </motion.div>
      </div>
    </section>
  );
};
