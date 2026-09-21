import React from 'react';

export const Hero: React.FC = () => {
  return (
    <section className="relative pt-16 pb-8 md:pt-20 md:pb-10 overflow-hidden">
      {/* Subtle background glow */}
      <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[600px] h-[250px] bg-gradient-to-tr from-blue-600/15 via-indigo-600/10 to-purple-600/5 blur-[120px] rounded-full pointer-events-none -z-10" />

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        {/* Main Headline - Direct, punchy and clean */}
        <h1 className="text-4xl sm:text-6xl lg:text-7xl font-extrabold text-white tracking-tight leading-[1.15] max-w-5xl mx-auto">
          让普通 AI，直接变成你的 <br />
          <span className="text-transparent bg-clip-text bg-gradient-to-r from-blue-400 via-indigo-300 to-emerald-400">
            Office 真实操作助手
          </span>
        </h1>
      </div>
    </section>
  );
};
