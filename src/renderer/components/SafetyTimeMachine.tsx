import React from 'react';
import { 
  ShieldCheck, 
  RotateCcw, 
  Trash2, 
  RefreshCw, 
  ChevronRight, 
  ChevronLeft, 
  Clock, 
  FileSpreadsheet, 
  History
} from 'lucide-react';
import appLogo from '../assets/app-logo.png';
import brandOk from '../assets/brand-ok.png';

interface SafetyTimeMachineProps {
  online: boolean;
  records: any[];
  totalRecords: number;
  offset: number;
  busy: string;
  onRefresh: () => void;
  onClearHistory: () => void;
  onSelectRecord: (id: string) => void;
  onPageChange: (newOffset: number) => void;
  onStartService: () => void;
}

export const SafetyTimeMachine: React.FC<SafetyTimeMachineProps> = ({
  online,
  records = [],
  totalRecords,
  offset,
  busy,
  onRefresh,
  onClearHistory,
  onSelectRecord,
  onPageChange,
  onStartService,
}) => {
  const isRefreshing = busy === 'audit-refresh';
  const pageSize = 20;
  const currentPage = Math.floor(offset / pageSize) + 1;
  const totalPages = Math.max(1, Math.ceil(totalRecords / pageSize));

  return (
    <section className="time-machine-card">
      {/* 头部标题与控制栏 */}
      <div className="time-machine-header">
        <div className="tm-title-group">
          <div className="tm-icon-wrap">
            <ShieldCheck size={16} className="text-emerald-500" />
          </div>
          <div>
            <div className="tm-title-line">
              <h3>安全时光机</h3>
              <span className="auto-protect-badge">
                <img src={brandOk} alt="自动快照保护中" className="brand-badge-icon" />
                <span>自动快照保护中</span>
              </span>
            </div>
            <p className="sub-caption">
              AI 每次对单元格值与公式的修改均会自动生成安全快照，随时可一键秒级还原
            </p>
          </div>
        </div>

        {/* 快捷操作 */}
        {online && (
          <div className="tm-actions">
            <button
              className="tm-btn secondary"
              disabled={Boolean(busy)}
              onClick={onRefresh}
              title="刷新最新快照记录"
            >
              <RefreshCw size={11} className={isRefreshing ? 'spin' : ''} />
              <span>刷新</span>
            </button>

            {totalRecords > 0 && (
              <button
                className="tm-btn danger"
                disabled={Boolean(busy)}
                onClick={onClearHistory}
                title="清空历史快照记录"
              >
                <Trash2 size={11} />
                <span>清空快照</span>
              </button>
            )}
          </div>
        )}
      </div>

      {/* 列表内容区：带平滑收缩过渡层 */}
      <div className="tm-stage-wrapper">
        {!online ? (
          <div className="tm-empty-state stage-fade-in" key="offline">
            <History size={32} className="text-slate-300 dark:text-slate-600" />
            <h4>Bridge 服务未启动</h4>
            <p className="sub-caption">启动服务后，AI 的每一次表格写入操作都会记录在此处并受保护。</p>
            <button className="tm-cta-btn" onClick={onStartService}>
              立即启动服务
            </button>
          </div>
        ) : records.length === 0 ? (
          <div className="tm-empty-state stage-fade-in" key="empty">
            <ShieldCheck size={32} className="text-emerald-300 dark:text-emerald-800" />
            <h4>安全快照已就绪，等待首次修改</h4>
            <p className="sub-caption">当 AI 开始为您修改表格单元格或公式时，修改对比记录将实时呈现。</p>
          </div>
        ) : (
          <div className="tm-records-list stage-fade-in" key={`list-${offset}`}>
            {records.map((rec) => {
              const isRolledBack = rec.status === 'rolled_back';
              const formattedTime = new Date(rec.timestamp).toLocaleString('zh-CN', {
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
              });

              return (
                <div
                  key={rec.id}
                  className={`tm-record-item ${isRolledBack ? 'rolled-back' : ''}`}
                  onClick={() => onSelectRecord(rec.id)}
                >
                  <div className="record-status-indicator">
                    {isRolledBack ? (
                      <RotateCcw size={13} className="text-slate-400" />
                    ) : (
                      <Clock size={13} className="text-emerald-500" />
                    )}
                  </div>

                  <div className="record-main-info">
                    <div className="record-headline">
                      <span className="record-desc">{rec.description || 'AI 修改单元格'}</span>
                      <span className={`record-state-tag ${isRolledBack ? 'gray' : 'green'}`}>
                        {isRolledBack ? '已撤销还原' : '已保护写入'}
                      </span>
                    </div>

                    <div className="record-subline">
                      <span className="record-target">
                        <FileSpreadsheet size={11} />
                        <strong>{rec.workbookName}</strong> / {rec.sheetName} · <code>{rec.address}</code>
                      </span>
                      <span className="record-meta">
                        {formattedTime} · 来自: {rec.clientName || 'AI'} · 变动 <strong>{rec.modifiedCount || 1}</strong> 个单元格
                      </span>
                    </div>
                  </div>

                  <div className="record-arrow">
                    <ChevronRight size={14} />
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* 分页控制栏 */}
      {totalRecords > pageSize && (
        <div className="tm-pagination">
          <button
            className="pager-btn"
            disabled={offset === 0 || Boolean(busy)}
            onClick={() => onPageChange(Math.max(0, offset - pageSize))}
          >
            <ChevronLeft size={12} />
            <span>上一页</span>
          </button>

          <span className="pager-info">
            第 {currentPage} 页 / 共 {totalPages} 页 <span className="sub-caption">（共 {totalRecords} 条记录）</span>
          </span>

          <button
            className="pager-btn"
            disabled={offset + pageSize >= totalRecords || Boolean(busy)}
            onClick={() => onPageChange(offset + pageSize)}
          >
            <span>下一页</span>
            <ChevronRight size={12} />
          </button>
        </div>
      )}
    </section>
  );
};
