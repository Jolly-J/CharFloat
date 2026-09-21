import React from 'react';
import { Cpu, Terminal, Sparkles, Check, ArrowUpRight } from 'lucide-react';

interface AgentGridProps {
  onOpenQuickStart: () => void;
}

const SUPPORTED_MODELS = [
  {
    name: 'Claude 3.7 / Claude Code',
    category: 'Anthropic',
    desc: '原生支持 stdio MCP 协议，通过代码级精准思考控制复杂表格与公式。',
    tag: '深度推理'
  },
  {
    name: 'Cursor & VS Code',
    category: 'IDE / Editor',
    desc: '配置一键生效，在 Composer / Chat 侧边栏直接对话操控打开中的本地文档。',
    tag: '开发者首选'
  },
  {
    name: 'ChatGPT Plus / 4o',
    category: 'OpenAI',
    desc: '通过标准 MCP / 本地 HTTP 代理通道连接，告别传统文件反复上传下载。',
    tag: '全能办公'
  },
  {
    name: '字节 豆包 / 扣子 Coze',
    category: '本土大模型',
    desc: '深度接入中文本土智能体，无缝结合企业内部知识库与本地桌面办公。',
    tag: '本土极速'
  },
  {
    name: '阿里 通义千问 Qwen',
    category: '超长上下文',
    desc: '处理十万行复杂业务表与多工作簿关联，极速执行结构化提取与回写。',
    tag: '海量数据'
  },
  {
    name: '企业自建 Agent / Dify',
    category: '定制生态',
    desc: '开放标准 RESTful / SSE / WebSocket 接口与 Python SDK，私有化部署零门槛。',
    tag: '企业可控'
  }
];

export const AgentGrid: React.FC<AgentGridProps> = ({ onOpenQuickStart }) => {
  return (
    <section className="py-20 bg-[#0b0b10] relative">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Section Header */}
        <div className="text-center max-w-3xl mx-auto mb-14">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-purple-500/10 text-purple-400 border border-purple-500/20 text-xs font-semibold mb-3">
            <Cpu className="w-3.5 h-3.5" />
            <span>无缝生态接入</span>
          </div>
          <h2 className="text-3xl sm:text-4xl font-extrabold text-white tracking-tight mb-4">
            继续使用你喜欢的 AI
          </h2>
          <p className="text-zinc-400 text-sm sm:text-base leading-relaxed">
            你不需要为了获得 Office 操作能力再换一个新的 AI。
            <br />
            <strong className="text-white">你负责选择最聪明的 AI，我们让它真正会用 Office。</strong>
          </p>
        </div>

        {/* Grid of supported AI */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 mb-12">
          {SUPPORTED_MODELS.map((item, idx) => (
            <div 
              key={idx}
              className="p-6 rounded-2xl bg-[#13131a] border border-white/[0.06] hover:border-blue-500/30 transition-all flex flex-col justify-between group hover:bg-[#161622] hover:-translate-y-1 shadow-lg"
            >
              <div>
                <div className="flex items-center justify-between mb-3">
                  <span className="text-[11px] font-semibold text-zinc-500 uppercase tracking-wider">
                    {item.category}
                  </span>
                  <span className="px-2 py-0.5 rounded-full bg-blue-500/10 text-blue-300 border border-blue-500/20 text-[10px] font-semibold">
                    {item.tag}
                  </span>
                </div>
                <h3 className="text-base font-bold text-white mb-2 flex items-center gap-1.5 group-hover:text-blue-300 transition-colors">
                  <span>{item.name}</span>
                </h3>
                <p className="text-xs text-zinc-400 leading-relaxed">{item.desc}</p>
              </div>

              <div className="mt-5 pt-4 border-t border-white/[0.04] flex items-center justify-between text-xs text-zinc-500">
                <span className="flex items-center gap-1 text-emerald-400 font-medium">
                  <Check className="w-3.5 h-3.5" /> 即插即用
                </span>
                <button 
                  onClick={onOpenQuickStart}
                  className="text-blue-400 hover:text-blue-300 flex items-center gap-0.5 font-medium transition-colors"
                >
                  <span>查看接入</span>
                  <ArrowUpRight className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>
          ))}
        </div>

        {/* Bottom CTA Box */}
        <div className="text-center">
          <button
            onClick={onOpenQuickStart}
            className="inline-flex items-center gap-2 px-8 py-3.5 rounded-xl bg-gradient-to-r from-blue-600 via-indigo-600 to-purple-600 hover:from-blue-500 hover:to-purple-500 text-white font-semibold text-sm shadow-xl shadow-blue-500/20 hover:scale-[1.02] active:scale-[0.98] transition-all"
          >
            <Sparkles className="w-4 h-4 text-blue-200" />
            <span>只需 1 分钟，为你的 AI 装上真实 Office 操作能力</span>
          </button>
        </div>
      </div>
    </section>
  );
};
