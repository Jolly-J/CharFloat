import React, { useState, useRef } from 'react';
import { 
  Play, 
  Pause, 
  Volume2, 
  VolumeX, 
  Maximize2, 
  Sparkles, 
  FileSpreadsheet, 
  CheckCircle2,
  Laptop
} from 'lucide-react';

interface VideoShowcaseProps {
  // 可直接传入自定义视频地址，默认读取 public 目录下的 demo.mp4
  videoSrc?: string;
  posterSrc?: string;
}

export const VideoShowcase: React.FC<VideoShowcaseProps> = ({
  videoSrc = '/demo.mp4',
  posterSrc = ''
}) => {
  const videoRef = useRef<HTMLVideoElement>(null);
  const [isPlaying, setIsPlaying] = useState(false);
  const [isMuted, setIsMuted] = useState(true);
  const [hasError, setHasError] = useState(false);

  const togglePlay = () => {
    if (!videoRef.current) return;
    if (isPlaying) {
      videoRef.current.pause();
      setIsPlaying(false);
    } else {
      videoRef.current.play().then(() => {
        setIsPlaying(true);
      }).catch(() => {
        // 如果浏览器阻止自动播放或无源文件，平滑降级
        setIsPlaying(true);
      });
    }
  };

  const toggleMute = () => {
    if (!videoRef.current) return;
    videoRef.current.muted = !isMuted;
    setIsMuted(!isMuted);
  };

  const handleFullScreen = () => {
    if (!videoRef.current) return;
    if (videoRef.current.requestFullscreen) {
      videoRef.current.requestFullscreen();
    }
  };

  return (
    <section id="video-demo" className="py-12 md:py-20 bg-gradient-to-b from-[#f8fafc] to-white relative overflow-hidden">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Section Pill Badge & Title */}
        <div className="text-center max-w-3xl mx-auto mb-10">
          <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-blue-50 text-blue-600 border border-blue-200/80 text-xs font-semibold mb-3">
            <Sparkles className="w-3.5 h-3.5" />
            <span>实机演示 · 1 分钟了解真实效果</span>
          </div>
          <h2 className="text-3xl sm:text-4xl md:text-5xl font-black text-slate-900 tracking-tight mb-4">
            你说一句话，<span className="text-blue-600">它当场在 Office 里把活干完</span>
          </h2>
          <p className="text-slate-600 text-sm sm:text-base leading-relaxed max-w-2xl mx-auto">
            全程真实录制：自动读取已打开的表格，边听你说，边把数据补全、公式算准、异常标黄、原生图表画好。
          </p>
        </div>

        {/* Video Player Showcase Window (Doubao Work Style Container) */}
        <div className="relative rounded-3xl overflow-hidden bg-slate-900 border border-slate-200/90 shadow-[0_25px_60px_-15px_rgba(15,23,42,0.15)] group">
          {/* Top Window Bar */}
          <div className="h-10 bg-slate-100 px-4 flex items-center justify-between border-b border-slate-200 select-none">
            <div className="flex items-center gap-2">
              <span className="w-3 h-3 rounded-full bg-[#ff5f56]" />
              <span className="w-3 h-3 rounded-full bg-[#ffbd2e]" />
              <span className="w-3 h-3 rounded-full bg-[#27c93f]" />
              <span className="ml-3 text-xs text-slate-600 font-semibold flex items-center gap-2">
                <Laptop className="w-3.5 h-3.5 text-slate-500" />
                <span>字浮 CharFloat · 真实操作视频展示</span>
              </span>
            </div>

            <div className="flex items-center gap-2 text-xs text-slate-500">
              <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 text-[11px] font-semibold">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping" />
                4K 原画演示
              </span>
            </div>
          </div>

          {/* Video Container (16:9 Aspect Ratio) */}
          <div className="relative aspect-video w-full bg-slate-950 flex items-center justify-center overflow-hidden">
            {/* Real HTML5 Video Tag */}
            <video
              ref={videoRef}
              src={videoSrc}
              poster={posterSrc}
              className="w-full h-full object-cover"
              loop
              playsInline
              muted={isMuted}
              onError={() => setHasError(true)}
              onPlay={() => setIsPlaying(true)}
              onPause={() => setIsPlaying(false)}
            />

            {/* Video Placeholder State (当用户尚未放置真实 demo.mp4 时展示的高级海报与播放引导) */}
            {(!isPlaying || hasError) && (
              <div 
                onClick={togglePlay}
                className="absolute inset-0 bg-gradient-to-tr from-slate-950 via-slate-900 to-blue-950/80 flex flex-col items-center justify-center text-center p-6 cursor-pointer select-none"
              >
                {/* Simulated live preview screen graphics */}
                <div className="absolute inset-0 opacity-20 pointer-events-none bg-[radial-gradient(#3b82f6_1px,transparent_1px)] [background-size:24px_24px]" />

                {/* Big Floating Play Capsule */}
                <div className="relative z-10 w-20 h-20 sm:w-24 sm:h-24 rounded-full bg-white/90 backdrop-blur-md text-blue-600 flex items-center justify-center shadow-2xl shadow-blue-500/30 group-hover:scale-110 group-hover:bg-white transition-all duration-300">
                  <Play className="w-8 h-8 sm:w-10 sm:h-10 fill-blue-600 ml-1" />
                </div>

                <div className="relative z-10 mt-6 max-w-md">
                  <h3 className="text-white font-extrabold text-lg sm:text-xl mb-2">
                    点击播放产品实操演示
                  </h3>
                  <p className="text-slate-400 text-xs sm:text-sm">
                    支持放置任意 mp4 视频文件至项目 public 目录（默认路径：<code className="text-blue-400">/demo.mp4</code>）
                  </p>
                </div>
              </div>
            )}

            {/* Floating Glass Control Bar at Bottom */}
            <div className="absolute bottom-4 left-4 right-4 bg-slate-900/80 backdrop-blur-xl border border-white/10 rounded-2xl px-4 py-2.5 flex items-center justify-between text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity duration-300">
              <div className="flex items-center gap-3">
                <button
                  onClick={togglePlay}
                  className="p-1.5 rounded-lg hover:bg-white/10 transition-colors cursor-pointer"
                  title={isPlaying ? '暂停' : '播放'}
                >
                  {isPlaying ? <Pause className="w-4 h-4 fill-white" /> : <Play className="w-4 h-4 fill-white" />}
                </button>

                <button
                  onClick={toggleMute}
                  className="p-1.5 rounded-lg hover:bg-white/10 transition-colors cursor-pointer"
                  title={isMuted ? '取消静音' : '静音'}
                >
                  {isMuted ? <VolumeX className="w-4 h-4" /> : <Volume2 className="w-4 h-4" />}
                </button>

                <span className="text-[11px] text-slate-400 font-mono">
                  演示：本地表格就地生成公式与走势图
                </span>
              </div>

              <div className="flex items-center gap-3">
                <button
                  onClick={handleFullScreen}
                  className="p-1.5 rounded-lg hover:bg-white/10 transition-colors cursor-pointer"
                  title="全屏"
                >
                  <Maximize2 className="w-4 h-4" />
                </button>
              </div>
            </div>
          </div>

          {/* Bottom Video Meta Bar */}
          <div className="bg-white px-5 py-3.5 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-600">
            <div className="flex items-center gap-4">
              <span className="flex items-center gap-1.5 font-semibold text-slate-900">
                <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                原声无剪辑操作
              </span>
              <span className="hidden sm:inline text-slate-400">|</span>
              <span className="hidden sm:inline text-slate-500">
                支持 WPS 2019+ 及 Office 365 真实桌面环境
              </span>
            </div>
            <div className="text-[11px] text-slate-400">
              视频留空位：替换 <code className="bg-slate-100 text-slate-700 px-1 py-0.5 rounded">public/demo.mp4</code> 即可自动生效
            </div>
          </div>
        </div>
      </div>
    </section>
  );
};
