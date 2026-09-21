import React, { useState } from 'react';
import { X, Copy, Check, Terminal, Sparkles, Layers } from 'lucide-react';

interface QuickStartModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const QuickStartModal: React.FC<QuickStartModalProps> = ({ isOpen, onClose }) => {
  const [activeTab, setActiveTab] = useState<'cursor' | 'claude' | 'vscode' | 'http'>('cursor');
  const [copied, setCopied] = useState(false);

  if (!isOpen) return null;

  const cursorConfig = `{
  "mcpServers": {
    "office-agent-bridge": {
      "command": "node",
      "args": ["/Users/Python/Office Agent Bridge/dist/bridge/cli.cjs"]
    }
  }
}`;

  const claudeCodeCommand = `claude mcp add office-agent-bridge node "/Users/Python/Office Agent Bridge/dist/bridge/cli.cjs"`;

  const vscodeConfig = `{
  "mcp": {
    "servers": {
      "office-agent-bridge": {
        "command": "node",
        "args": ["/Users/Python/Office Agent Bridge/dist/bridge/cli.cjs"]
      }
    }
  }
}`;

  const httpConfig = `# 本地后台守护进程模式 (默认端口 19890)
curl http://127.0.0.1:19890/status \\
  -H "Authorization: Bearer <local_token>"`;

  let currentCode = cursorConfig;
  if (activeTab === 'claude') currentCode = claudeCodeCommand;
  if (activeTab === 'vscode') currentCode = vscodeConfig;
  if (activeTab === 'http') currentCode = httpConfig;

  const handleCopy = () => {
    navigator.clipboard.writeText(currentCode);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-fadeIn">
      <div 
        className="w-full max-w-2xl bg-[#14141c] border border-white/[0.12] rounded-2xl shadow-2xl overflow-hidden text-zinc-200"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Modal Header */}
        <div className="px-6 py-4 border-b border-white/[0.08] flex items-center justify-between bg-[#191924]">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-600/20 border border-blue-500/30 flex items-center justify-center text-blue-400">
              <Terminal className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-base font-bold text-white">接入 Office Agent Bridge</h3>
              <p className="text-xs text-zinc-400">1 分钟配置，让你的 AI 立即获得 Office 实时控制权</p>
            </div>
          </div>
          <button 
            onClick={onClose}
            className="p-1.5 rounded-lg text-zinc-400 hover:text-white hover:bg-white/[0.06] transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Client Selection Tabs */}
        <div className="flex border-b border-white/[0.08] bg-[#101016] px-6 pt-3 gap-2 text-xs font-medium">
          <button
            onClick={() => setActiveTab('cursor')}
            className={`pb-3 px-3 border-b-2 transition-colors ${
              activeTab === 'cursor'
                ? 'border-blue-500 text-blue-400 font-bold'
                : 'border-transparent text-zinc-400 hover:text-white'
            }`}
          >
            Cursor (.cursor/mcp.json)
          </button>
          <button
            onClick={() => setActiveTab('claude')}
            className={`pb-3 px-3 border-b-2 transition-colors ${
              activeTab === 'claude'
                ? 'border-amber-500 text-amber-400 font-bold'
                : 'border-transparent text-zinc-400 hover:text-white'
            }`}
          >
            Claude Code CLI
          </button>
          <button
            onClick={() => setActiveTab('vscode')}
            className={`pb-3 px-3 border-b-2 transition-colors ${
              activeTab === 'vscode'
                ? 'border-indigo-500 text-indigo-400 font-bold'
                : 'border-transparent text-zinc-400 hover:text-white'
            }`}
          >
            VS Code / Roo / Cline
          </button>
          <button
            onClick={() => setActiveTab('http')}
            className={`pb-3 px-3 border-b-2 transition-colors ${
              activeTab === 'http'
                ? 'border-emerald-500 text-emerald-400 font-bold'
                : 'border-transparent text-zinc-400 hover:text-white'
            }`}
          >
            HTTP / SSE (Dify/自建)
          </button>
        </div>

        {/* Code Block Area */}
        <div className="p-6">
          <div className="relative rounded-xl overflow-hidden bg-[#0a0a0e] border border-white/[0.08]">
            <div className="px-4 py-2 bg-[#121218] border-b border-white/[0.06] flex items-center justify-between text-xs text-zinc-400">
              <span className="font-mono">{activeTab === 'claude' ? '终端一键执行' : '配置文件内容'}</span>
              <button
                onClick={handleCopy}
                className="flex items-center gap-1 px-2.5 py-1 rounded bg-white/[0.06] hover:bg-white/[0.12] text-zinc-200 transition-colors font-medium text-[11px]"
              >
                {copied ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-400" />
                    <span className="text-emerald-400">已复制到剪贴板！</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>一键复制配置</span>
                  </>
                )}
              </button>
            </div>
            <pre className="p-4 text-xs font-mono text-blue-300 overflow-x-auto leading-relaxed">
              <code>{currentCode}</code>
            </pre>
          </div>

          <div className="mt-4 p-3 rounded-lg bg-blue-500/10 border border-blue-500/20 text-xs text-blue-200 leading-relaxed">
            <span className="font-bold">提示：</span>配置完成后，在你的 AI 对话框中直接说：
            <em className="text-white">“帮我看下当前打开的 Excel 经营分析表，把 8 月数据补齐”</em> 即可体验实时就地操作！
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-4 bg-[#101016] border-t border-white/[0.08] flex items-center justify-between text-xs text-zinc-400">
          <span>无需暴露公网 ｜ 100% 运行于本地 127.0.0.1</span>
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-lg bg-white/[0.08] hover:bg-white/[0.15] text-white font-medium transition-colors"
          >
            完成并关闭
          </button>
        </div>
      </div>
    </div>
  );
};
