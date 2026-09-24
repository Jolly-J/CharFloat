import { Sparkles, ArrowUpRight, CheckCircle2 } from 'lucide-react';
import { motion } from 'motion/react';

interface AgentGridProps {
  onOpenQuickStart: () => void;
}

const SUPPORTED_AGENTS = [
  {
    name: 'WorkBuddy',
    logo: '/agents/workbuddy.svg',
    category: '腾讯助手',
    tag: '推荐',
    desc: '腾讯智能助手原生联动，一键接入后在日常对话中唤醒本地 Office 驱动，指哪改哪。',
    status: '已就绪'
  },
  {
    name: '豆包工作',
    logo: '/agents/doubao-color.svg',
    category: '字节生态',
    tag: '官方适配',
    desc: '字节跳动 AI 工作平台，字浮让它打破网页边界，直接在正在运行的 WPS 中就地标色与修表。',
    status: '已就绪'
  },
  {
    name: '千问办公',
    logo: '/agents/qwen-color.svg',
    category: '阿里通义',
    tag: '一键接入',
    desc: '阿里巴巴通义千问桌面端，本地通道秒级接入，让长文本分析与复杂计算直接增量注入表格。',
    status: '支持接入'
  },
  {
    name: 'Kimi',
    logo: '/agents/kimi.webp',
    category: '月之暗面',
    tag: '长文推理',
    desc: 'Moonshot Kimi 客户端，配置已就绪，直接在客户端对话提取与回写眼前正打开的工作表。',
    status: '已就绪'
  },
  {
    name: 'Claude Code',
    logo: '/agents/claude-color.svg',
    category: 'Anthropic',
    tag: '深度代码',
    desc: 'Anthropic 官方智能助手与 CLI，原生支持标准 MCP 协议，代码级精准控制 Office 内部对象。',
    status: '已就绪'
  },
  {
    name: 'Codex / ChatGPT',
    logo: '/agents/openai.svg',
    category: 'OpenAI',
    tag: '全能办公',
    desc: 'OpenAI 官方客户端与生态工具，秒级打通桌面通道，告别静态文件反复上传下载。',
    status: '已就绪'
  }
];

export const AgentGrid: React.FC<AgentGridProps> = ({ onOpenQuickStart }) => {
  return (
    <section id="models" className="py-20 md:py-28 bg-[#f8fafc] relative scroll-mt-24">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Section Header */}
        <div className="text-center max-w-3xl mx-auto mb-14">
          <h2 className="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight mb-4">
            你习惯用谁，就给谁装上字浮
          </h2>
          <p className="text-slate-600 text-sm sm:text-base leading-relaxed">
            不用离开你熟悉的客户端。<br className="hidden sm:inline" />
            给豆包、WorkBuddy、通义千问或 Claude 装上字浮外挂，它们就能直接动手改你桌面上正打开的文档。
          </p>
        </div>

        {/* Grid of 6 supported AI with Real Icons */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5 mb-10">
          {SUPPORTED_AGENTS.map((item, idx) => (
            <motion.div 
              key={idx}
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.6, delay: idx * 0.08, ease: [0.16, 1, 0.3, 1] }}
              whileHover={{ y: -6, scale: 1.015, boxShadow: '0 20px 30px -10px rgba(0,0,0,0.07)' }}
              className="p-6 rounded-3xl bg-white border border-slate-200/90 shadow-sm transition-colors hover:border-blue-300 flex flex-col justify-between cursor-default"
            >
              <div>
                {/* Card Top: Real Brand Logo + Category Tag */}
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-3">
                    <img 
                      src={item.logo} 
                      alt={item.name} 
                      className="w-9 h-9 object-contain shrink-0 select-none"
                      loading="lazy"
                    />
                    <div>
                      <h3 className="text-base font-bold text-slate-900 leading-tight">
                        {item.name}
                      </h3>
                      <span className="text-[11px] font-medium text-slate-400">
                        {item.category}
                      </span>
                    </div>
                  </div>

                  <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-bold border ${
                    item.tag === '推荐' 
                      ? 'bg-amber-50 text-amber-700 border-amber-200' 
                      : 'bg-blue-50 text-blue-700 border-blue-200/60'
                  }`}>
                    {item.tag}
                  </span>
                </div>

                <p className="text-xs text-slate-600 leading-relaxed font-normal">
                  {item.desc}
                </p>
              </div>

              {/* Card Footer: Ready Status */}
              <div className="mt-5 pt-4 border-t border-slate-100 flex items-center justify-between text-xs text-slate-400">
                <span>客户端一键授权</span>
                <span className="text-blue-600 font-semibold flex items-center gap-1">
                  <CheckCircle2 className="w-3.5 h-3.5 text-blue-600" />
                  <span>{item.status}</span>
                </span>
              </div>
            </motion.div>
          ))}
        </div>

        {/* Quick config button banner */}
        <div className="text-center">
          <motion.button
            onClick={onOpenQuickStart}
            whileHover={{ scale: 1.03 }}
            whileTap={{ scale: 0.97 }}
            transition={{ type: "spring", stiffness: 450, damping: 25 }}
            className="inline-flex items-center gap-2.5 px-6 py-3 rounded-full bg-white hover:bg-slate-50 border border-slate-200 shadow-sm text-xs sm:text-sm font-bold text-slate-700 hover:text-blue-600 transition-colors cursor-pointer"
          >
            {/* Overlapping Brand Avatars Stack (WorkBuddy, Doubao, Qwen) */}
            <div className="flex items-center -space-x-1.5 shrink-0">
              <div className="w-5 h-5 rounded-full bg-white border border-slate-200 shadow-2xs flex items-center justify-center p-0.5 z-30">
                <img src="/agents/workbuddy.svg" alt="WorkBuddy" className="w-full h-full object-contain" />
              </div>
              <div className="w-5 h-5 rounded-full bg-white border border-slate-200 shadow-2xs flex items-center justify-center p-0.5 z-20">
                <img src="/agents/doubao-color.svg" alt="豆包" className="w-full h-full object-contain" />
              </div>
              <div className="w-5 h-5 rounded-full bg-white border border-slate-200 shadow-2xs flex items-center justify-center p-0.5 z-10">
                <img src="/agents/qwen-color.svg" alt="通义千问" className="w-full h-full object-contain" />
              </div>
            </div>
            <span>让 AI 助手为自己配置字浮外挂（支持 Cursor / Windsurf / VS Code 等）</span>
            <ArrowUpRight className="w-3.5 h-3.5 ml-0.5" />
          </motion.button>
        </div>
      </div>
    </section>
  );
};
