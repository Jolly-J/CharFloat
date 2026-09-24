import React, { useEffect, useRef, useState } from 'react';
import { 
  X, 
  Copy, 
  Check, 
  CheckCircle2, 
  AlertCircle,
  ChevronDown
} from 'lucide-react';

interface DoubaoGuideModalProps {
  open: boolean;
  port?: number;
  token?: string;
  onClose: () => void;
  onMarkConfigured?: () => void;
  onOpenDoubaoApp?: () => void;
  onCopyText: (text: string) => void;
}

export const DoubaoGuideModal: React.FC<DoubaoGuideModalProps> = ({
  open,
  port = 19890,
  token = '',
  onClose,
  onMarkConfigured,
  onOpenDoubaoApp,
  onCopyText,
}) => {
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const [activeOpen, setActiveOpen] = useState(open);
  const [isClosing, setIsClosing] = useState(false);
  const exitTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  // 同步外部 open 状态并处理平滑退场动画
  useEffect(() => {
    if (open) {
      clearTimeout(exitTimer.current);
      setIsClosing(false);
      setActiveOpen(true);
    } else if (activeOpen && !isClosing) {
      setIsClosing(true);
      exitTimer.current = setTimeout(() => {
        setActiveOpen(false);
        setIsClosing(false);
      }, 160);
    }
  }, [open]);

  // 处理主动触发关闭
  const handleRequestClose = () => {
    if (isClosing) return;
    setIsClosing(true);
    exitTimer.current = setTimeout(() => {
      setIsClosing(false);
      setActiveOpen(false);
      onClose();
    }, 160);
  };

  // 快捷键支持（Escape 触发平滑退场）
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && activeOpen) {
        handleRequestClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
      clearTimeout(exitTimer.current);
    };
  }, [activeOpen]);

  if (!activeOpen) return null;

  const serverName = '字浮 CharFloat';
  const serverUrl = `http://127.0.0.1:${port}/mcp`;

  const handleCopy = (key: string, text: string) => {
    onCopyText(text);
    setCopiedKey(key);
    setTimeout(() => {
      setCopiedKey(null);
    }, 1800);
  };

  return (
    <div 
      className={`modal-backdrop ${isClosing ? 'is-exiting' : 'is-entering'}`} 
      onMouseDown={(e) => e.target === e.currentTarget && handleRequestClose()}
    >
      <div 
        className={`modal-window doubao-guide-modal ${isClosing ? 'is-exiting' : 'is-entering'}`} 
        style={{ maxWidth: '640px' }}
      >
        {/* 顶部标题栏（无星星图标） */}
        <div className="modal-header">
          <div className="diff-header-left">
            <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 600 }}>豆包工作 · 连接器配置指引</h4>
          </div>
          <button className="icon-btn-ghost" onClick={handleRequestClose} aria-label="关闭">
            <X size={15} />
          </button>
        </div>

        {/* 内容区 */}
        <div className="doubao-guide-body" style={{ padding: '16px 20px', overflowY: 'auto' }}>
          {/* 状态提要横幅 */}
          <div className="doubao-status-banner">
            <div className="doubao-status-item success">
              <CheckCircle2 size={14} className="text-emerald-500 shrink-0" />
              <div>
                <strong>Skill 已自动注入</strong>
                <p>4 个专属 Office 提示词技能已同步至本地 <code>.user_skills</code> 目录，对话时可直接识别。</p>
              </div>
            </div>
            <div className="doubao-status-item warning">
              <AlertCircle size={14} className="text-amber-500 shrink-0" />
              <div>
                <strong>MCP 工具连接器需手动添加</strong>
                <p>只需粘贴服务器名称与地址，自定义 Headers 保持默认留空，直接点击保存即可连接成功。</p>
              </div>
            </div>
          </div>

          {/* 步骤提示 */}
          <div className="doubao-step-prompt">
            <span>操作路径：打开<strong>豆包工作</strong> ➔ 左侧进入<strong>【技能·连接器·伙伴】</strong> ➔ 点击<strong>【新建自定义连接器】</strong>，对照下方填入：</span>
          </div>

          {/* 1:1 复刻豆包“新建自定义连接器”窗口 */}
          <div className="doubao-mock-window">
            <div className="mock-window-header">
              <div>
                <h5 className="mock-title">新建自定义连接器</h5>
                <p className="mock-subtitle">自定义连接器仅支持在本地电脑中使用</p>
              </div>
              <span className="mock-close-hint">✕</span>
            </div>

            <div className="mock-form">
              {/* 第一行：服务器名称与传输类型 */}
              <div className="mock-form-row">
                <div className="mock-form-group" style={{ flex: 1.2 }}>
                  <label className="mock-label">
                    服务器名称 <span className="mock-req">*</span>
                  </label>
                  <div className="mock-input-wrap">
                    <input 
                      type="text" 
                      readOnly 
                      value={serverName} 
                      className="mock-input"
                    />
                    <button 
                      type="button"
                      className="mock-copy-btn"
                      onClick={() => handleCopy('name', serverName)}
                      title="点击复制"
                    >
                      {copiedKey === 'name' ? <Check size={12} className="text-emerald-500" /> : <Copy size={12} />}
                      <span>{copiedKey === 'name' ? '已复制' : '复制'}</span>
                    </button>
                  </div>
                </div>

                <div className="mock-form-group" style={{ flex: 0.8 }}>
                  <label className="mock-label">传输类型</label>
                  <div className="mock-select-box">
                    <span>HTTP</span>
                    <ChevronDown size={14} className="text-muted" />
                  </div>
                </div>
              </div>

              {/* 第二行：服务器 URL */}
              <div className="mock-form-group">
                <label className="mock-label">
                  服务器 URL <span className="mock-req">*</span>
                </label>
                <div className="mock-input-wrap">
                  <input 
                    type="text" 
                    readOnly 
                    value={serverUrl} 
                    className="mock-input highlight-url"
                  />
                  <button 
                    type="button"
                    className="mock-copy-btn"
                    onClick={() => handleCopy('url', serverUrl)}
                    title="点击复制"
                  >
                    {copiedKey === 'url' ? <Check size={12} className="text-emerald-500" /> : <Copy size={12} />}
                    <span>{copiedKey === 'url' ? '已复制' : '复制'}</span>
                  </button>
                </div>
              </div>

              {/* 第三行：自定义 Headers（完全免填） */}
              <div className="mock-form-group">
                <label className="mock-label">自定义 Headers</label>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '2px' }}>
                  <div 
                    style={{ 
                      display: 'inline-flex', 
                      alignItems: 'center', 
                      padding: '5px 14px', 
                      borderRadius: '6px', 
                      border: '1px dashed var(--border-subtle)', 
                      background: 'var(--surface-tertiary, #f8fafc)',
                      color: 'var(--text-muted)',
                      fontSize: '12px',
                      cursor: 'default',
                      userSelect: 'none'
                    }}
                  >
                    + 添加
                  </div>
                  <span style={{ fontSize: '12px', color: '#10b981', display: 'flex', alignItems: 'center', gap: '4px', fontWeight: 500 }}>
                    <Check size={13} className="text-emerald-500" />
                    <span>无需添加，直接留空即可</span>
                  </span>
                </div>
              </div>

              {/* 模拟保存按钮 */}
              <div className="mock-footer-row">
                <span className="mock-btn-cancel">取消</span>
                <span className="mock-btn-save">保存（在豆包中点击）</span>
              </div>
            </div>
          </div>
        </div>

        {/* 弹窗底部操作栏 */}
        <div className="modal-footer" style={{ justifyContent: 'space-between', alignItems: 'center', padding: '12px 20px' }}>
          <button 
            type="button" 
            className="dialog-btn secondary"
            onClick={() => {
              if (onMarkConfigured) onMarkConfigured();
              handleRequestClose();
            }}
            title="标记已在豆包中配置好，后续一键配置不再自动弹出此指引"
          >
            <Check size={14} className="text-emerald-500" />
            <span>我已配置成功</span>
          </button>

          <button 
            type="button" 
            className="dialog-btn primary"
            onClick={() => {
              if (onOpenDoubaoApp) onOpenDoubaoApp();
              handleRequestClose();
            }}
          >
            <span>打开豆包工作并前往配置</span>
          </button>
        </div>
      </div>
    </div>
  );
};
