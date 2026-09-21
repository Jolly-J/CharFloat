import React, { useEffect, useRef, useState } from 'react';
import { 
  X, 
  RotateCcw, 
  FileSpreadsheet, 
  ArrowRight, 
  LoaderCircle,
  ShieldAlert
} from 'lucide-react';

interface DiffModalProps {
  record: any;
  confirm: 'stop' | 'rollback' | 'clear-audit' | null;
  busy: string;
  onCloseRecord: () => void;
  onCloseConfirm: () => void;
  onRequestRollback: () => void;
  onConfirmAction: () => void;
}

export const DiffModal: React.FC<DiffModalProps> = ({
  record,
  confirm,
  busy,
  onCloseRecord,
  onCloseConfirm,
  onRequestRollback,
  onConfirmAction,
}) => {
  const modalRef = useRef<HTMLDivElement>(null);
  
  // 维护平滑进退场动画的缓存状态
  const [activeRecord, setActiveRecord] = useState<any>(record);
  const [activeConfirm, setActiveConfirm] = useState<'stop' | 'rollback' | 'clear-audit' | null>(confirm);
  const [isClosing, setIsClosing] = useState(false);
  const exitTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  // 同步外部传入状态并执行平滑退场生命周期
  useEffect(() => {
    if (record) {
      clearTimeout(exitTimer.current);
      setIsClosing(false);
      setActiveRecord(record);
    } else if (activeRecord && !isClosing) {
      // 外部触发关闭时先播放 160ms 退场动画
      setIsClosing(true);
      exitTimer.current = setTimeout(() => {
        setActiveRecord(null);
        setIsClosing(false);
      }, 160);
    }
  }, [record]);

  useEffect(() => {
    if (confirm) {
      clearTimeout(exitTimer.current);
      setIsClosing(false);
      setActiveConfirm(confirm);
    } else if (activeConfirm && !isClosing) {
      // 外部触发关闭时先播放 160ms 退场动画
      setIsClosing(true);
      exitTimer.current = setTimeout(() => {
        setActiveConfirm(null);
        setIsClosing(false);
      }, 160);
    }
  }, [confirm]);

  // 主动触发平滑关闭
  const handleRequestClose = (type: 'record' | 'confirm') => {
    if (isClosing) return;
    setIsClosing(true);
    exitTimer.current = setTimeout(() => {
      setIsClosing(false);
      if (type === 'record') {
        setActiveRecord(null);
        onCloseRecord();
      } else {
        setActiveConfirm(null);
        onCloseConfirm();
      }
    }, 160);
  };

  // 快捷键支持 (Escape 触发平滑退场)
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (activeConfirm) handleRequestClose('confirm');
        else if (activeRecord) handleRequestClose('record');
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
      clearTimeout(exitTimer.current);
    };
  }, [activeRecord, activeConfirm]);

  // 如果没有任何内容需要渲染
  if (!activeRecord && !activeConfirm) return null;

  // 二次危险操作确认弹窗 (带平滑双向动画)
  if (activeConfirm) {
    const isStop = activeConfirm === 'stop';
    const isClear = activeConfirm === 'clear-audit';

    return (
      <div 
        className={`modal-backdrop ${isClosing ? 'is-exiting' : 'is-entering'}`} 
        onMouseDown={(e) => e.target === e.currentTarget && handleRequestClose('confirm')}
      >
        <div 
          className={`modal-window confirm-dialog ${isClosing ? 'is-exiting' : 'is-entering'}`} 
          ref={modalRef} 
          role="dialog" 
          aria-modal="true"
        >
          <div className="modal-header">
            <div className="diff-badge-icon">
              <ShieldAlert size={16} className="text-amber-500" />
            </div>
            <div className="modal-title-box">
              <h3>
                {isStop
                  ? '停止 Bridge 桥梁服务？'
                  : isClear
                  ? '清空全部安全快照记录？'
                  : '确认撤销这次修改？'}
              </h3>
            </div>
            <button className="modal-close-btn" onClick={() => handleRequestClose('confirm')} aria-label="关闭">
              <X size={15} />
            </button>
          </div>

          <div className="modal-body">
            <p className="sub-caption" style={{ fontSize: '12px', lineHeight: '1.6' }}>
              {isStop
                ? '停止服务后，所有 AI 客户端将暂时断开与 WPS/Office 的连接。办公软件中的文档不会被关闭，您可以随时重新启动服务。'
                : isClear
                ? '此操作将永久抹除本地所有历史单元格修改快照与可回滚数据。此操作不可逆，请谨慎确认。'
                : '系统将严格恢复修改前的单元格数值与公式。如果该区域已被后续操作覆盖，系统将安全拦截以防止覆写新数据。'}
            </p>
          </div>

          <div className="modal-footer">
            <button className="dialog-btn secondary" onClick={() => handleRequestClose('confirm')}>
              取消
            </button>
            <button
              className={`dialog-btn ${isStop || isClear ? 'danger' : 'primary'}`}
              disabled={Boolean(busy)}
              onClick={onConfirmAction}
            >
              {busy === activeConfirm ? <LoaderCircle size={13} className="spin" /> : null}
              <span>{isStop ? '停止服务' : isClear ? '确认清空' : '确认一键撤销'}</span>
            </button>
          </div>
        </div>
      </div>
    );
  }

  // 单元格 Diff 对比详情弹窗 (带平滑双向动画)
  if (activeRecord) {
    const isRolledBack = activeRecord.status === 'rolled_back';
    const diffItems = activeRecord.diff || [];
    const displayDiffs = diffItems.slice(0, 150);

    return (
      <div 
        className={`modal-backdrop ${isClosing ? 'is-exiting' : 'is-entering'}`} 
        onMouseDown={(e) => e.target === e.currentTarget && handleRequestClose('record')}
      >
        <div 
          className={`modal-window diff-modal-window ${isClosing ? 'is-exiting' : 'is-entering'}`} 
          ref={modalRef} 
          role="dialog" 
          aria-modal="true"
        >
          {/* 弹窗头部 */}
          <div className="modal-header">
            <div className="diff-header-left">
              <div className="diff-badge-icon">
                <FileSpreadsheet size={16} className="text-emerald-500" />
              </div>
              <div>
                <h3>单元格变更前后对比</h3>
                <p className="sub-caption">
                  {activeRecord.workbookName} · {activeRecord.sheetName} · 选区 <code>{activeRecord.address}</code>
                </p>
              </div>
            </div>

            <button className="modal-close-btn" onClick={() => handleRequestClose('record')} aria-label="关闭">
              <X size={15} />
            </button>
          </div>

          {/* 弹窗主体：描述与表格 */}
          <div className="modal-body">
            <div className="diff-summary-card">
              <div className="summary-line">
                <span className="summary-label">修改说明:</span>
                <span className="summary-val">{activeRecord.description || 'AI 自动化批量写入'}</span>
              </div>
              <div className="summary-line">
                <span className="summary-label">来源与时间:</span>
                <span className="summary-val">
                  {new Date(activeRecord.timestamp).toLocaleString()} · 由【{activeRecord.clientName || 'AI'}】发起 · 共变动 {activeRecord.modifiedCount || diffItems.length} 个单元格
                </span>
              </div>
            </div>

            {/* 对比表格 */}
            <div className="diff-table-container">
              <table className="diff-table">
                <thead>
                  <tr>
                    <th style={{ width: '85px' }}>单元格</th>
                    <th>修改前原值</th>
                    <th style={{ width: '28px', textAlign: 'center' }}></th>
                    <th>AI 写入后新值</th>
                  </tr>
                </thead>
                <tbody>
                  {displayDiffs.map((d: any, index: number) => {
                    const oldStr = String(d.oldFormula?.startsWith?.('=') ? d.oldFormula : d.oldValue ?? '(空)');
                    const newStr = String(d.newFormula?.startsWith?.('=') ? d.newFormula : d.newValue ?? '(空)');
                    return (
                      <tr key={index}>
                        <td className="cell-pos">
                          <code>{d.cell}</code>
                        </td>
                        <td className="cell-old">
                          <span className="diff-tag del">{oldStr}</span>
                        </td>
                        <td className="cell-arrow">
                          <ArrowRight size={11} className="text-slate-400" />
                        </td>
                        <td className="cell-new">
                          <span className="diff-tag add">{newStr}</span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {diffItems.length > 150 && (
              <p className="sub-caption">
                仅展示前 150 项差异，完整数据均已安全存储在本地快照中。
              </p>
            )}
          </div>

          {/* 弹窗底栏 */}
          <div className="modal-footer">
            <button className="dialog-btn secondary" onClick={() => handleRequestClose('record')}>
              关闭
            </button>

            {!isRolledBack && (
              <button
                className="dialog-btn rollback-btn"
                disabled={Boolean(busy)}
                onClick={onRequestRollback}
              >
                <RotateCcw size={13} />
                <span>撤销这次修改（一键还原）</span>
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }

  return null;
};
