import React, { useEffect, useRef, useState } from 'react';
import {
  AlertTriangle, ArrowDown, ArrowDownToLine, ArrowRight, ArrowUp, ChartNoAxesCombined, Check,
  CheckCheck, ChevronDown, CircleCheck, Copy, FileArchive, FileSpreadsheet,
  Focus, MessageSquare, Mic, Monitor, MousePointer2, Pause, Play, Plus, RotateCcw,
  ScanLine, ShieldCheck, Sparkles, Upload, WandSparkles, X, Zap,
} from 'lucide-react';
import { motion, AnimatePresence, useInView } from 'motion/react';
import './ComparisonDemo.css';

const REQUEST = '帮我把这个月的数据做成分析报表，并补一张趋势图，保留现有格式。';
const FINAL_STEP = 5;
const STEP_DURATION = 2200;

const FILE_STEPS = [
  { icon: Upload, step: '01', title: '上传文件', detail: '离开当前工作区' },
  { icon: FileArchive, step: '02', title: '拆包 / 解析', detail: '从文件中提取内容' },
  { icon: ScanLine, step: '03', title: '分析内容', detail: '理解上传时的版本' },
  { icon: WandSparkles, step: '04', title: '处理任务', detail: '按理解重新组织内容' },
  { icon: Copy, step: '05', title: '重新生成文件', detail: '每次修改，另存一份' },
  { icon: ArrowDownToLine, step: '06', title: '发回给你', detail: '下载、检查，再接着改' },
];

const OFFICE_STEPS = [
  { label: '读取', message: '正在读取当前工作表', detail: '识别已有数据、公式与表结构' },
  { label: '快照', message: '生成修改前的安全快照', detail: '为支持的操作保留恢复依据' },
  { label: '选区', message: '精准修改指定区域', detail: '选中 B2:D5，沿用原有格式与公式' },
  { label: '标色', message: '添加条件格式', detail: '营收增长标蓝，成本增长标黄' },
  { label: '图表', message: '插入营收与利润组合图', detail: '图表直接放入当前工作表' },
  { label: '完成', message: '当前文件已修改完成', detail: '分析结论已更新，继续在这里工作' },
];

const BENEFITS = [
  { icon: FileSpreadsheet, title: '当前文件直接改', text: '免去上传、下载与反复生成副本' },
  { icon: ChartNoAxesCombined, title: '复杂内容也能做', text: '表格、PPT、图表与矢量图形' },
  { icon: Focus, title: '精准到具体区域', text: '指定选区、页面或对象，定点修改' },
  { icon: Monitor, title: '所见即所得', text: '边看、边说、边改，就在 Office 里' },
];

const GENERATED_VERSIONS = [
  'v1.xlsx',
  'v2.xlsx',
  'v3_微调.xlsx',
  'v4_改公式.xlsx',
  'v5_最终版.xlsx',
  'v6_真最终版.xlsx',
  'v12_老板意见.xlsx',
  'v28_排版错乱重导.xlsx',
  'v45_再改一版.xlsx',
  'v68_彻底重来.xlsx',
  'v88_今晚必交.xlsx',
  'v99_最终不改版.xlsx',
  'v100_改回第一版.xlsx',
];

function FileFlow() {
  const [versionNum, setVersionNum] = useState(1);

  useEffect(() => {
    const timer = setInterval(() => {
      setVersionNum((prev) => (prev >= 100 ? 1 : prev + 1));
    }, 140);
    return () => clearInterval(timer);
  }, []);

  return (
    <article className="comparison-track comparison-track--file" aria-labelledby="file-track-title">
      <header className="comparison-track-heading">
        <h3 id="file-track-title">在文件外，重新生成</h3>
        <p>文件在来回走，当前工作区没有变。</p>
      </header>

      {/* Simulated AI Software Input Dialog (DeepSeek / ChatGPT style) */}
      <div className="comparison-dialog-box" aria-label="AI 软件输入框模拟">
        {/* Uploaded File Attachment Pill */}
        <div className="comparison-dialog-attachment">
          <span className="comparison-excel-icon">
            <FileSpreadsheet size={18} strokeWidth={1.8} />
          </span>
          <div className="comparison-dialog-file-info">
            <span className="comparison-dialog-filename">月度经营.xlsx</span>
            <span className="comparison-dialog-filesize">24.5 KB</span>
          </div>
        </div>

        {/* User Prompt Text */}
        <p className="comparison-dialog-prompt">
          “{REQUEST}”
        </p>

        {/* Footer Toolbar: +, Context, Model, Mic, Send */}
        <div className="comparison-dialog-footer">
          <div className="comparison-dialog-tools">
            <span className="comparison-dialog-tool-btn" aria-hidden="true">
              <Plus size={14} />
            </span>
            <span className="comparison-dialog-pill">
              <Check size={11} className="comparison-dialog-check" />
              <span>默认权限</span>
              <ChevronDown size={11} />
            </span>
          </div>
          <div className="comparison-dialog-actions">
            <span className="comparison-dialog-model">
              <Zap size={11} className="comparison-dialog-zap" />
              <span>快速</span>
              <ChevronDown size={11} />
            </span>
            <span className="comparison-dialog-mic">
              <Mic size={14} />
            </span>
            <span className="comparison-dialog-send">
              <ArrowUp size={13} strokeWidth={2.5} />
            </span>
          </div>
        </div>
      </div>
      <div className="comparison-file-flow" role="list" aria-label="传统文件处理流程">
        {FILE_STEPS.map(({ icon: Icon, step, title, detail }, index) => (
          <React.Fragment key={title}>
            <div className="comparison-flow-card" role="listitem">
              <div className="comparison-flow-main">
                <span className="comparison-flow-icon">
                  <Icon size={15} />
                </span>
                <div className="comparison-flow-text">
                  <strong>{title}</strong>
                  <span className="comparison-flow-step">STEP {step}</span>
                </div>
              </div>
              <span className="comparison-flow-detail">{detail}</span>
            </div>
            {index < FILE_STEPS.length - 1 && (
              <div className="comparison-flow-divider" aria-hidden="true">
                <ArrowDown size={11} />
              </div>
            )}
          </React.Fragment>
        ))}
      </div>
      <div className="comparison-new-file">
        {/* Infinite scrolling versions marquee from v1 to v100 */}
        <div className="comparison-versions-marquee-wrap" aria-label="每改一次，又多一个版本">
          <div className="comparison-versions-marquee">
            {[...GENERATED_VERSIONS, ...GENERATED_VERSIONS].map((ver, idx) => (
              <span key={idx}>{ver}</span>
            ))}
          </div>
        </div>
        <div className="comparison-file-card">
          <FileSpreadsheet size={27} strokeWidth={1.5} />
          <div>
            <span>得到一个新文件</span>
            <strong>分析报表_final_v{versionNum}.xlsx</strong>
          </div>
          <div className="comparison-file-card-badge">
            <ArrowDownToLine size={16} />
            <span>v{versionNum}</span>
          </div>
        </div>
        <p><Copy size={12} />每说一句话，本地又多一份文件（已堆积 {versionNum} 个副本）</p>
      </div>
      <div className="comparison-caveats" aria-label="传统文件重构模式的痛点">
        <div className="comparison-caveats-heading">
          <AlertTriangle size={15} className="comparison-icon-warning" />
          <span>传统重构模式的隐患</span>
        </div>
        <div className="comparison-caveats-tags">
          <span className="comparison-tag-danger">
            <X size={13} strokeWidth={2.5} /> 表结构与排版易被篡改
          </span>
          <span className="comparison-tag-danger">
            <X size={13} strokeWidth={2.5} /> 原公式 / 宏 / 元数据易丢失
          </span>
          <span className="comparison-tag-danger">
            <X size={13} strokeWidth={2.5} /> 复杂图表与矢量图形受限
          </span>
        </div>
      </div>
    </article>
  );
}

function TrendChart({ visible }: { visible: boolean }) {
  return (
    <div className={`comparison-chart ${visible ? 'is-visible' : ''}`}>
      {visible ? (
        <>
          <div className="comparison-chart-heading"><strong>营收与利润趋势</strong><span><i />营收<b />利润率</span></div>
          <svg viewBox="0 0 440 96" role="img" aria-label="组合图：7 至 9 月营收为 86、94、108 万元，利润率为 23.3%、23.4%、25.9%">
            {[20, 44, 68].map(y => <line key={y} x1="34" y1={y} x2="410" y2={y} stroke="#e8eef6" strokeDasharray="3 4" />)}
            <text x="5" y="12" fill="#64748b" fontSize="9">万元</text><text x="412" y="12" fill="#64748b" fontSize="9">%</text>
            <text x="6" y="29" fill="#94a3b8" fontSize="9">120</text><text x="412" y="29" fill="#94a3b8" fontSize="9">30</text>
            <text x="17" y="76" fill="#94a3b8" fontSize="9">0</text><text x="412" y="76" fill="#94a3b8" fontSize="9">0</text>
            <rect x="80" y="37.6" width="44" height="34.4" rx="3" fill="#bfdbfe" />
            <rect x="196" y="34.4" width="44" height="37.6" rx="3" fill="#93c5fd" />
            <rect x="312" y="28.8" width="44" height="43.2" rx="3" fill="#3b82f6" />
            <path d="M102 34.7 L218 34.6 L334 30.6" fill="none" stroke="#0f766e" strokeWidth="2" strokeLinecap="round" />
            {[[102, 34.7], [218, 34.6], [334, 30.6]].map(([x, y]) => <circle key={x} cx={x} cy={y} r="3" fill="white" stroke="#0f766e" strokeWidth="2" />)}
            <text x="96" y="27" fill="#64748b" fontSize="10">86</text><text x="212" y="26" fill="#64748b" fontSize="10">94</text><text x="325" y="20" fill="#2563eb" fontSize="10" fontWeight="600">108</text>
            <text x="94" y="90" fill="#64748b" fontSize="10">7 月</text><text x="210" y="90" fill="#64748b" fontSize="10">8 月</text><text x="326" y="90" fill="#64748b" fontSize="10">9 月</text>
          </svg>
        </>
      ) : (
        <div className="comparison-chart-placeholder"><ChartNoAxesCombined size={24} strokeWidth={1.3} /><span>趋势图将直接插入这里</span><small>当前工作表 · 数据区域下方</small></div>
      )}
    </div>
  );
}

function OfficeWindow({ step }: { step: number }) {
  const selected = step >= 2;
  const formatted = step >= 3;
  const complete = step === FINAL_STEP;
  const rows = [
    ['营业收入', '86', '94', '108'],
    ['运营成本', '66', '72', '80'],
    ['营业利润', '20', '22', '28'],
    ['利润率', '23.3%', '23.4%', '25.9%'],
  ];

  return (
    <div className="comparison-office-window" aria-label="WPS / Office 当前工作表模拟窗口">
      <div className="comparison-window-title">
        <span className="comparison-window-app"><FileSpreadsheet size={15} /></span><strong>月度经营.xlsx</strong>
        <span className="comparison-open-file">当前打开</span>
        <span className="comparison-window-dots" aria-hidden="true"><i /><i /><i /></span>
      </div>
      <div className="comparison-ribbon" aria-hidden="true"><b>开始</b><span>插入</span><span>页面布局</span><span>公式</span><span>数据</span><span>WPS / Office</span></div>
      <div className="comparison-formula"><span>{selected ? 'B2:D5' : 'D5'}<ChevronDown size={10} /></span><i>fx</i><code>=D4/D2</code><span className="comparison-formula-kept"><Check size={11} />原公式保留</span></div>
      <div className="comparison-sheet">
        <div className="comparison-sheet-title">
          <strong>月度经营分析</strong>
          {selected ? <span className="comparison-selection-label">{complete ? <Check size={11} /> : <MousePointer2 size={11} fill="currentColor" />}{complete ? '指定区域已更新' : 'AI 正在此处编辑'}</span> : <span>2026 · 09</span>}
        </div>
        <div className={`comparison-grid-wrap ${selected ? 'is-selected' : ''}`}>
          <table className={`comparison-sheet-grid ${formatted ? 'is-formatted' : ''}`}>
            <caption className="sr-only">当前工作表，金额单位：万元。数据与公式保持不变，仅修改指定区域的样式。</caption>
            <thead><tr><th aria-label="行号" /><th scope="col">A</th><th scope="col">B</th><th scope="col">C</th><th scope="col">D</th></tr></thead>
            <tbody>
              <tr className="comparison-data-header"><th scope="row">1</th><td>经营指标 / 万元</td><td>7 月</td><td>8 月</td><td>9 月</td></tr>
              {rows.map((row, index) => (
                <tr key={row[0]} className={index === 1 && formatted ? 'comparison-cost-row' : ''}>
                  <th scope="row">{index + 2}</th>
                  {row.map((value, column) => <td key={column} className={column === 3 && formatted ? 'comparison-changed-cell' : ''}>{value}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
          {selected && <div className="comparison-selection" aria-hidden="true"><i /></div>}
        </div>
        <TrendChart visible={step >= 4} />
        <div className={`comparison-insight ${complete ? 'is-complete' : ''}`}><Sparkles size={13} /><span>{complete ? '9 月营收环比 +14.9%，利润率提升至 25.9%。' : '分析结论将在这里更新，原始数据继续保留。'}</span></div>
      </div>
      <div className="comparison-sheet-tabs"><span>＋</span><strong>经营分析</strong><span>原始数据</span><span><Check size={11} />{complete ? '修改完成' : '当前工作表'}</span></div>
    </div>
  );
}

export const ComparisonTable: React.FC = () => {
  const workspaceRef = useRef<HTMLDivElement>(null);
  const inView = useInView(workspaceRef, { amount: 0.3 });
  const [reducedMotion, setReducedMotion] = useState(() =>
    typeof window !== 'undefined' && window.matchMedia('(prefers-reduced-motion: reduce)').matches
  );
  const [step, setStep] = useState(0);
  const [pageVisible, setPageVisible] = useState(true);
  const displayedStep = reducedMotion ? FINAL_STEP : step;
  const complete = displayedStep === FINAL_STEP;
  const current = OFFICE_STEPS[displayedStep];

  useEffect(() => {
    const preference = window.matchMedia('(prefers-reduced-motion: reduce)');
    const onChange = () => setReducedMotion(preference.matches);
    onChange();
    preference.addEventListener('change', onChange);
    return () => preference.removeEventListener('change', onChange);
  }, []);

  useEffect(() => {
    const onVisibility = () => setPageVisible(!document.hidden);
    onVisibility();
    document.addEventListener('visibilitychange', onVisibility);
    return () => document.removeEventListener('visibilitychange', onVisibility);
  }, []);

  // Seamless auto-looping execution steps
  useEffect(() => {
    if (!inView || !pageVisible || reducedMotion) return;
    const duration = step === FINAL_STEP ? 3000 : STEP_DURATION;
    const timer = window.setTimeout(() => {
      setStep(value => (value >= FINAL_STEP ? 0 : value + 1));
    }, duration);
    return () => window.clearTimeout(timer);
  }, [step, inView, pageVisible, reducedMotion]);

  return (
    <section id="comparison" className="comparison-demo scroll-mt-24" aria-labelledby="comparison-title">
      <div className="comparison-container">
        <header className="comparison-heading">
          <span className="comparison-eyebrow">同一个需求 · 两条工作路径</span>
          <h2 id="comparison-title">同一句话，<span>两种结果</span></h2>
          <p>区别不在 AI 会不会做，而在它到底是在文件外生成，<br className="hidden sm:block" />还是在当前 Office 里直接操作。</p>
        </header>
        <div className="comparison-request">
          <span className="comparison-request-icon"><MessageSquare size={19} /></span>
          <div><span>你的同一句需求</span><p>“{REQUEST}”</p></div>
          <span className="comparison-request-send" aria-hidden="true"><ArrowRight size={17} /></span>
        </div>

        {/* Branching Fork with Route Pills */}
        <div className="comparison-fork-container" aria-hidden="true">
          <div className="comparison-fork-stem" />
          <div className="comparison-fork-arms">
            <div className="comparison-fork-arm comparison-fork-arm--left">
              <div className="comparison-fork-pill-wrap">
                <span className="comparison-fork-pill comparison-fork-pill--file">
                  <span className="comparison-fork-dot" />
                  没有插件 · 传统 AI
                </span>
                <span className="comparison-fork-arrow">
                  <ArrowDown size={11} strokeWidth={2.5} />
                </span>
              </div>
            </div>
            <div className="comparison-fork-arm comparison-fork-arm--right">
              <div className="comparison-fork-pill-wrap">
                <span className="comparison-fork-pill comparison-fork-pill--office">
                  <img src="/logo.png" alt="字浮 Logo" className="comparison-fork-logo" />
                  装上「字浮」外挂插件
                </span>
                <span className="comparison-fork-arrow comparison-fork-arrow--office">
                  <ArrowDown size={11} strokeWidth={2.5} />
                </span>
              </div>
            </div>
          </div>
        </div>

        <div className="comparison-tracks">
          <FileFlow />
          <article className="comparison-track comparison-track--office" aria-labelledby="office-track-title">
            <header className="comparison-track-heading">
              <h3 id="office-track-title">连接你喜爱的 AI，当场操作 WPS / Office</h3>
              <p>不是重新做个 AI。字浮作为外挂桥梁，给现有 AI 装上直接操控桌面文档的双手。</p>
            </header>

            {/* Bridge Connection Flow Bar: Your AI -> CharFloat Plugin -> WPS/Office */}
            <div className="comparison-bridge-bar" aria-label="协同连接链路：你的 AI 软件 -> 字浮插件 -> 正在运行的 Office">
              <div className="comparison-bridge-node">
                <span className="comparison-bridge-icon comparison-bridge-ai">
                  <img src="/agents/workbuddy.svg" alt="WorkBuddy" className="comparison-bridge-brand-img" />
                </span>
                <div className="comparison-bridge-text">
                  <strong>你的 AI 软件</strong>
                  <span>WorkBuddy / 豆包 / ChatGPT</span>
                </div>
              </div>

              <div className="comparison-bridge-link" aria-hidden="true">
                <span className="comparison-bridge-line" />
                <ArrowRight size={13} className="comparison-bridge-arrow" />
              </div>

              <div className="comparison-bridge-node comparison-bridge-node--hub">
                <img src="/logo.png" alt="字浮 Logo" className="comparison-bridge-hub-logo" />
                <div className="comparison-bridge-text">
                  <strong>字浮插件</strong>
                  <span className="comparison-bridge-badge">已激活 · 外挂中</span>
                </div>
              </div>

              <div className="comparison-bridge-link" aria-hidden="true">
                <span className="comparison-bridge-line" />
                <ArrowRight size={13} className="comparison-bridge-arrow" />
              </div>

              <div className="comparison-bridge-node">
                <span className="comparison-bridge-icon comparison-bridge-office">
                  <img src="/wps-logo.png" alt="WPS Office" className="comparison-bridge-brand-img" />
                </span>
                <div className="comparison-bridge-text">
                  <strong>WPS / Office</strong>
                  <span>桌面打开的文档</span>
                </div>
              </div>
            </div>

            {/* AI Input Dialog with /字浮 CharFloat Skill Command (No file upload needed) */}
            <div className="comparison-dialog-box comparison-dialog-box--bridge" aria-label="带有字浮外挂的 AI 软件输入框">
              {/* Active Document Status Bar (No file upload needed) */}
              <div className="comparison-dialog-skill-bar">
                <span className="comparison-dialog-skill-status">
                  <span className="comparison-dialog-skill-dot" />
                  已识别桌面 WPS 打开的文档
                </span>
              </div>

              {/* User Prompt Text calling CharFloat with single Logo Skill Tag */}
              <p className="comparison-dialog-prompt">
                <span className="comparison-dialog-skill-tag">
                  <img src="/logo.png" alt="字浮" className="comparison-dialog-skill-logo" />
                  <span>/字浮 CharFloat</span>
                </span>
                ，看到我在 WPS 打开的月度经营表，帮我把这个月的数据做成分析报表并补一张趋势图，保留现有格式。
              </p>

              {/* Footer Toolbar */}
              <div className="comparison-dialog-footer">
                <div className="comparison-dialog-tools">
                  <span className="comparison-dialog-tool-btn" aria-hidden="true">
                    <Plus size={14} />
                  </span>
                  <span className="comparison-dialog-pill">
                    <Check size={11} className="comparison-dialog-check" />
                    <span>Office 原生权限</span>
                    <ChevronDown size={11} />
                  </span>
                </div>
                <div className="comparison-dialog-actions">
                  <span className="comparison-dialog-model">
                    <img src="/agents/workbuddy.svg" alt="WorkBuddy" className="w-3.5 h-3.5 object-contain" />
                    <span>WorkBuddy</span>
                    <ChevronDown size={11} />
                  </span>
                  <span className="comparison-dialog-mic">
                    <Mic size={14} />
                  </span>
                  <span className="comparison-dialog-send comparison-dialog-send--active">
                    <ArrowUp size={13} strokeWidth={2.5} />
                  </span>
                </div>
              </div>
            </div>

            <div className="comparison-workspace" ref={workspaceRef}>
              <OfficeWindow step={displayedStep} />
              <div className={`comparison-agent-status ${complete ? 'is-complete' : ''}`} role="status" aria-live="polite" aria-atomic="true">
                <AnimatePresence mode="wait" initial={false}>
                  {complete ? (
                    <motion.span
                      key="done"
                      initial={{ scale: 0.7, opacity: 0 }}
                      animate={{ scale: 1, opacity: 1 }}
                      exit={{ scale: 0.7, opacity: 0 }}
                      transition={{ duration: 0.22, ease: [0.16, 1, 0.3, 1] }}
                      className="comparison-agent-direct-check"
                    >
                      <Check size={26} strokeWidth={2.8} />
                    </motion.span>
                  ) : (
                    <motion.img
                      key="acting"
                      src="/logo.png"
                      alt="字浮"
                      initial={{ scale: 0.8, opacity: 0 }}
                      animate={{ scale: 1, opacity: 1 }}
                      exit={{ scale: 0.8, opacity: 0 }}
                      transition={{ duration: 0.22, ease: [0.16, 1, 0.3, 1] }}
                      className="comparison-agent-direct-logo"
                    />
                  )}
                </AnimatePresence>

                <div className="comparison-agent-text-wrap">
                  <AnimatePresence mode="wait" initial={false}>
                    <motion.div
                      key={displayedStep}
                      initial={{ opacity: 0, y: 5 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -5 }}
                      transition={{ duration: 0.22, ease: 'easeOut' }}
                    >
                      <strong>{current.message}</strong>
                      <span>{current.detail}</span>
                    </motion.div>
                  </AnimatePresence>
                </div>

                <div className="comparison-step-count-wrap">
                  <AnimatePresence mode="wait" initial={false}>
                    <motion.span
                      key={displayedStep}
                      initial={{ opacity: 0, y: 4 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -4 }}
                      transition={{ duration: 0.18, ease: 'easeOut' }}
                      className="comparison-step-count"
                    >
                      {String(displayedStep + 1).padStart(2, '0')} / 06
                    </motion.span>
                  </AnimatePresence>
                </div>
              </div>
              <ol className="comparison-progress" aria-label="Office 模拟操作进度">
                {OFFICE_STEPS.map((item, index) => <li key={item.label} className={index <= displayedStep ? 'is-reached' : ''} aria-current={index === displayedStep ? 'step' : undefined}><span />{item.label}</li>)}
              </ol>
            </div>
            <div className="comparison-preserved" aria-label="字浮原生协同的安全保障">
              <div className="comparison-preserved-heading">
                <ShieldCheck size={16} className="comparison-icon-shield" />
                <span>字浮原生协同保障</span>
              </div>
              <div className="comparison-preserved-tags">
                <span className="comparison-tag-safe">
                  <Check size={13} strokeWidth={2.5} /> 100% 沿用原文件（0 多余副本）
                </span>
                <span className="comparison-tag-safe">
                  <Check size={13} strokeWidth={2.5} /> 原有公式与排版样式全保留
                </span>
                <span className="comparison-tag-safe">
                  <Check size={13} strokeWidth={2.5} /> 安全快照 · 支持操作随时撤回
                </span>
              </div>
            </div>
          </article>
        </div>
        <div className="comparison-benefits">
          {BENEFITS.map(({ icon: Icon, title, text }) => <div key={title}><Icon size={19} strokeWidth={1.7} /><h3>{title}</h3><p>{text}</p></div>)}
        </div>
        <p className="comparison-footnote">以上为工作方式模拟，非实时操作；具体能力因宿主、文档类型与操作而异。</p>
      </div>
    </section>
  );
};
