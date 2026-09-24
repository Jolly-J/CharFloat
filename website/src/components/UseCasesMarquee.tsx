import React from 'react';
import { motion } from 'motion/react';

interface PromptCase {
  avatar: string;
  characterName: string;
  userPrompt: string;
  aiStatus: string;
  aiAction: string;
}

const ROW_1: PromptCase[] = [
  {
    avatar: '/character/工作.png',
    characterName: '工作字浮',
    userPrompt: '帮我把 8 月数据补齐，按去年的计算口径把同比增长率算出来，原有公式别动。',
    aiStatus: '收到，正在就地执行',
    aiAction: '调用 excel_apply_formulas，已精准继承去年计算公式并补全 G 列，其余列完好。'
  },
  {
    avatar: '/character/OK.png',
    characterName: '搞定字浮',
    userPrompt: '把获客成本超标和转化率未达标的行标出来，不要标红，改成浅黄色。',
    aiStatus: '收到，开始高亮筛选',
    aiAction: '调用 excel_format_cells，已定位 14 处异常行并柔和标为浅黄色。'
  },
  {
    avatar: '/character/点赞.png',
    characterName: '点赞字浮',
    userPrompt: '根据 6 到 8 月的月度营收数据画一张折线图，放在表格右下方。',
    aiStatus: '收到，正在生成图表',
    aiAction: '调用 excel_create_chart，已就地生成绑定数据源的原生折线走势图。'
  },
  {
    avatar: '/character/交文件.png',
    characterName: '交件字浮',
    userPrompt: '以员工工号为索引，把 Sheet2 的绩效评分匹配回写到主表的 F 列。',
    aiStatus: '收到，开始跨表匹配',
    aiAction: '调用 excel_vlookup_merge，已完成 128 条工号对齐并填入绩效分。'
  },
  {
    avatar: '/character/喝咖啡.png',
    characterName: '摸鱼字浮',
    userPrompt: '清洗当前选区的所有客户姓名和手机号，去除首尾空格并剔除重复项。',
    aiStatus: '收到，正在数据清洗',
    aiAction: '调用 excel_clean_range，已清除 36 处多余空格并去重 8 条记录。'
  }
];

const ROW_2: PromptCase[] = [
  {
    avatar: '/character/打气.png',
    characterName: '加油字浮',
    userPrompt: '刚才那步格式调整撤回，恢复刚才的初态，这个我自己来微调。',
    aiStatus: '收到，秒级快照还原',
    aiAction: '调用 time_machine_rollback，已当场恢复至 1 分钟前的原始状态。'
  },
  {
    avatar: '/character/想法.png',
    characterName: '灵感字浮',
    userPrompt: '统计今天华东各门店的总成交额与退款率，生成在表格最下方汇总行。',
    aiStatus: '收到，正在汇总分析',
    aiAction: '调用 excel_insert_summary，已在表格底部计算总成交额与加权退款率。'
  },
  {
    avatar: '/character/比心.png',
    characterName: '贴心字浮',
    userPrompt: '把左边正在打开的采购合同第 3 条付款条款金额，填入审批表 C5 单元格。',
    aiStatus: '收到，开始跨文档提取',
    aiAction: '调用 word_read_paragraph 与 excel_write_cell，已精准填入 ¥128,000。'
  },
  {
    avatar: '/character/思考.png',
    characterName: '沉思字浮',
    userPrompt: '将 D 列金额全部格式化为千分位并保留两位小数，负数用括号表示。',
    aiStatus: '收到，正在规范格式',
    aiAction: '调用 excel_set_number_format，已将 D 列应用统一财会千分位规范。'
  },
  {
    avatar: '/character/开心.png',
    characterName: '元气字浮',
    userPrompt: '读取当前表格的 8 月经营亮点，原地写入桌面打开中的汇报 PPT 第 6 页。',
    aiStatus: '收到，正在生成汇报',
    aiAction: '调用 ppt_insert_bullets，已将 4 条经营亮点结构化写入第 6 页幻灯片。'
  }
];

export const UseCasesMarquee: React.FC = () => {
  const renderCard = (item: PromptCase, idx: number) => {
    return (
      <div
        key={idx}
        className="w-[360px] sm:w-[400px] shrink-0 p-5 rounded-2xl bg-gradient-to-br from-white via-[#f4f9fd] to-[#d8ebfc]/45 border border-[#bfe0fa]/70 shadow-[0_4px_20px_rgba(186,215,245,0.25)] hover:shadow-md hover:border-blue-400 transition-all flex flex-col justify-between select-none"
      >
        {/* User Request Bubble */}
        <div className="mb-4">
          <div className="flex items-center justify-between mb-2">
            <span className="text-[11px] font-bold text-slate-400">
              你对字浮说：
            </span>
          </div>
          <p className="text-xs sm:text-[13px] font-semibold text-slate-800 leading-relaxed bg-white/70 backdrop-blur-sm rounded-xl p-2.5 border border-slate-200/60 shadow-xs">
            “{item.userPrompt}”
          </p>
        </div>

        {/* AI Action with Cloud Character */}
        <div className="pt-3 border-t border-[#d5e9fc] flex items-start gap-3">
          <img 
            src={item.avatar} 
            alt={item.characterName} 
            className="w-10 h-10 object-contain shrink-0 drop-shadow-xs"
            loading="lazy"
          />
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-2 mb-0.5">
              <span className="text-xs font-bold text-blue-600">
                {item.aiStatus}
              </span>
            </div>
            <p className="text-[11px] text-slate-600 leading-normal font-medium">
              {item.aiAction}
            </p>
          </div>
        </div>
      </div>
    );
  };

  return (
    <section id="use-cases" className="py-20 md:py-28 bg-[#f8fafc] overflow-hidden relative scroll-mt-24">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 text-center mb-12">
        <h2 className="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight mb-4">
          “帮我把这个表改好”，而不是“帮我做个表”
        </h2>
        <p className="text-base text-slate-600 max-w-2xl mx-auto leading-relaxed">
          告别毫无意义的虚构模版。真实办公需要的是在已有的复杂数据上精准动刀，一起改到满意为止。
        </p>
      </div>

      {/* Infinite Marquee Track (Row 1 - Faster: 32s) */}
      <div className="relative w-full overflow-hidden mb-5 py-2 group">
        {/* Left & Right Gradient Shadows for Seamless Look */}
        <div className="absolute left-0 top-0 bottom-0 w-28 bg-gradient-to-r from-[#f8fafc] via-[#f8fafc]/80 to-transparent z-10 pointer-events-none" />
        <div className="absolute right-0 top-0 bottom-0 w-28 bg-gradient-to-l from-[#f8fafc] via-[#f8fafc]/80 to-transparent z-10 pointer-events-none" />

        <div 
          className="flex gap-4 w-max animate-marquee group-hover:[animation-play-state:paused]"
          style={{ animationDuration: '32s' }}
        >
          {ROW_1.concat(ROW_1).concat(ROW_1).map((item, idx) => renderCard(item, idx))}
        </div>
      </div>

      {/* Infinite Marquee Track (Row 2 - Slower & Staggered Reverse: 48s) */}
      <div className="relative w-full overflow-hidden py-2 group">
        {/* Left & Right Gradient Shadows for Seamless Look */}
        <div className="absolute left-0 top-0 bottom-0 w-28 bg-gradient-to-r from-[#f8fafc] via-[#f8fafc]/80 to-transparent z-10 pointer-events-none" />
        <div className="absolute right-0 top-0 bottom-0 w-28 bg-gradient-to-l from-[#f8fafc] via-[#f8fafc]/80 to-transparent z-10 pointer-events-none" />

        <div 
          className="flex gap-4 w-max animate-marquee-reverse group-hover:[animation-play-state:paused]"
          style={{ animationDuration: '48s' }}
        >
          {ROW_2.slice(2).concat(ROW_2.slice(0, 2)).concat(ROW_2).concat(ROW_2).map((item, idx) => renderCard(item, idx + 100))}
        </div>
      </div>
    </section>
  );
};
