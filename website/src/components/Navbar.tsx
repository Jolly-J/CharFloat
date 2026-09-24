import React, { useState } from 'react';
import { Download, Menu, X } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface NavbarProps {
  onOpenQuickStart: () => void;
  onOpenDownload?: (tab?: 'mac' | 'win' | 'cli') => void;
}

export const Navbar: React.FC<NavbarProps> = ({ onOpenQuickStart, onOpenDownload }) => {
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  return (
    <header className="fixed top-5 left-0 right-0 z-50 flex justify-center px-4 pointer-events-none">
      {/* Floating Capsule Bar - High-End Frosted Glass with Motion */}
      <motion.div 
        initial={{ y: -30, opacity: 0, scale: 0.96 }}
        animate={{ y: 0, opacity: 1, scale: 1 }}
        transition={{ duration: 0.6, ease: [0.16, 1, 0.3, 1] }}
        className="w-full max-w-5xl bg-white/65 backdrop-blur-2xl backdrop-saturate-150 border border-white/80 shadow-[0_12px_36px_rgba(15,23,42,0.06),inset_0_1px_2px_rgba(255,255,255,0.95)] rounded-full px-5 py-2.5 flex items-center justify-between pointer-events-auto transition-colors"
      >
        {/* Brand Logo & Name */}
        <motion.a 
          href="#" 
          className="flex items-center gap-2.5 group shrink-0"
          whileHover={{ scale: 1.02 }}
          whileTap={{ scale: 0.98 }}
          transition={{ type: "spring", stiffness: 400, damping: 25 }}
        >
          <img 
            src="/logo.png" 
            alt="字浮 CharFloat" 
            className="w-8 h-8 object-contain select-none transition-transform"
          />
          <span className="font-extrabold text-base tracking-tight text-slate-900">
            字浮 CharFloat
          </span>
        </motion.a>

        {/* Center Nav Links (Desktop) */}
        <nav className="hidden md:flex items-center gap-10 lg:gap-14 text-[13.5px] font-normal text-slate-600/90 tracking-wide">
          <motion.a whileHover={{ y: -1, color: '#2563eb' }} transition={{ duration: 0.2 }} href="#delivery" className="hover:text-blue-600 transition-colors py-1">就地交付</motion.a>
          <motion.a whileHover={{ y: -1, color: '#2563eb' }} transition={{ duration: 0.2 }} href="#comparison" className="hover:text-blue-600 transition-colors py-1">传统对比</motion.a>
          <motion.a whileHover={{ y: -1, color: '#2563eb' }} transition={{ duration: 0.2 }} href="#models" className="hover:text-blue-600 transition-colors py-1">AI 外挂生态</motion.a>
          <motion.a whileHover={{ y: -1, color: '#2563eb' }} transition={{ duration: 0.2 }} href="#use-cases" className="hover:text-blue-600 transition-colors py-1">真实用例</motion.a>
        </nav>

        {/* Right Action Button */}
        <div className="hidden sm:flex items-center gap-3 shrink-0">
          <motion.button
            onClick={() => onOpenDownload ? onOpenDownload() : undefined}
            whileHover={{ scale: 1.03, boxShadow: '0 10px 25px -5px rgba(37,99,235,0.4)' }}
            whileTap={{ scale: 0.97 }}
            transition={{ type: "spring", stiffness: 400, damping: 25 }}
            className="flex items-center gap-2 px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-full shadow-md shadow-blue-500/20 cursor-pointer"
          >
            <Download className="w-3.5 h-3.5" />
            <span>下载电脑客户端</span>
          </motion.button>
        </div>

        {/* Mobile menu trigger */}
        <button
          onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
          className="md:hidden p-1.5 rounded-full text-slate-600 hover:bg-slate-100/80 transition-colors"
          aria-label="菜单"
        >
          {mobileMenuOpen ? <X className="w-5 h-5" /> : <Menu className="w-5 h-5" />}
        </button>
      </motion.div>

      {/* Mobile Drawer */}
      <AnimatePresence>
        {mobileMenuOpen && (
          <motion.div 
            initial={{ opacity: 0, y: -15, scale: 0.96 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: -15, scale: 0.96 }}
            transition={{ duration: 0.25, ease: [0.16, 1, 0.3, 1] }}
            className="md:hidden fixed top-20 left-4 right-4 bg-white/80 backdrop-blur-2xl border border-white/80 rounded-3xl p-5 shadow-2xl pointer-events-auto space-y-4"
          >
          <div className="flex flex-col gap-2.5 text-sm font-semibold text-slate-700">
            <a 
              href="#delivery" 
              onClick={() => setMobileMenuOpen(false)}
              className="px-3 py-2 rounded-xl hover:bg-slate-100"
            >
              就地交付
            </a>
            <a 
              href="#comparison" 
              onClick={() => setMobileMenuOpen(false)}
              className="px-3 py-2 rounded-xl hover:bg-slate-100"
            >
              传统对比
            </a>
            <a 
              href="#models" 
              onClick={() => setMobileMenuOpen(false)}
              className="px-3 py-2 rounded-xl hover:bg-slate-100"
            >
              AI 外挂生态
            </a>
            <a 
              href="#use-cases" 
              onClick={() => setMobileMenuOpen(false)}
              className="px-3 py-2 rounded-xl hover:bg-slate-100"
            >
              真实用例
            </a>
          </div>

          <div className="pt-3 border-t border-slate-100 flex flex-col gap-2">
            <button
              onClick={() => {
                setMobileMenuOpen(false);
                if (onOpenDownload) onOpenDownload();
              }}
              className="w-full flex items-center justify-center gap-2 py-2.5 rounded-full bg-blue-600 text-white font-bold text-xs shadow-md shadow-blue-500/20 cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>下载电脑客户端 (Win / Mac)</span>
            </button>
          </div>
        </motion.div>
      )}
      </AnimatePresence>
    </header>
  );
};
