/**
 * Excel（表格）工具定义（P2.3 剩余项：逐类迁移自 gateway.getOpenAiTools）。
 *
 * 已发布顺序中的第 11–22 条与第 24–38 条，分类标签为 `excel`。
 * 两段之间夹着审计工具 `wps_rollback`（第 23 条，见 `audit.ts`），
 * 因此本模块导出两个函数，由 `definitions/index.ts` 按已发布位置拼接；
 * 把两段合并会改变工具顺序，契约快照会失败。
 *
 * `ctx.wbDesc` 提供 workbookName 的描述文本，`ctx.activeWbHint` 参与
 * `wps_get_workspace_summary` 的 description 分支——两者都必须原样透传。
 */
import type { DefinitionContext, GatewayToolDefinition } from './shared.js';

/** Excel（表格）前半区：已发布顺序第 11–22 条。定义顺序即契约顺序，不得重排。 */
export function excelToolDefinitions(ctx: DefinitionContext): GatewayToolDefinition[] {
  return [
    {
      type: "function",
      function: {
        name: "wps_get_workspace_summary",
        description: ctx.activeWbHint
          ? `【当前WPS已在线打开工作簿: ${ctx.activeWbHint}】获取当前在 WPS 中打开的 Excel 工作簿概览、已打开的全部文件列表、包含的工作表列表与当前鼠标光标选区坐标。当前已打开 [${ctx.activeWbHint}]，严禁在磁盘搜索文件！`
          : "获取当前在 WPS 中打开的 Excel 工作簿概览、已打开的全部文件列表、包含的工作表列表与当前鼠标光标选区坐标。严禁在磁盘搜索文件！",
        parameters: {
          type: "object",
          properties: {
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_get_sheet_outline",
        description: "按需获取指定工作表的数据边界(UsedRange)与前3行表头样本，类似查看代码大纲，不耗多余Token",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_create_sheet",
        description: "在指定工作簿中独立新建工作表并自动激活呈现",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "新工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["sheetName"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_delete_sheet",
        description: "从指定工作簿中安全删除不需要的工作表",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "要删除的工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["sheetName"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_read_range",
        description: "切片读取指定区域(如 A1:C10)的单元格值与公式",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "区域地址，例如 'A1:E20'" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            includeFormulas: { type: "boolean", description: "是否返回公式，默认 true" },
            includeNumberFormats: { type: "boolean", description: "是否返回数字格式，默认 false；按需开启以控制返回量" }
          },
          required: ["address"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_get_range_styles",
        description: "读取指定区域的单元格样式。默认 summary 仅返回区域级样式摘要；cells 模式逐格返回并受 maxCells 限制。",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "区域地址，例如 'A1:E20'" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            mode: { type: "string", enum: ["summary", "cells"], description: "返回模式，默认 summary" },
            include: {
              type: "array",
              items: { type: "string", enum: ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea", "borders"] },
              description: "只返回指定样式字段；不传时返回常用字段"
            },
            maxCells: { type: "number", description: "cells 模式最多展开的单元格数，默认 100，最大 500" }
          },
          required: ["address"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_search_cells",
        description: "在表格中快速搜索包含指定文本或公式的单元格坐标",
        parameters: {
          type: "object",
          properties: {
            query: { type: "string", description: "搜索关键词" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            maxResults: { type: "number", description: "最大返回条数，默认 50" }
          },
          required: ["query"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_patch_cells",
        description: "直接在当前打开的 WPS 表格中原地修改数值或公式，自动抓取快照并实时呈现在用户屏幕上",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "目标区域地址，例如 'C2:C10'" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            values: {
              type: "array",
              items: { type: "array", items: {} },
              description: "二维数组数值"
            },
            formulas: {
              type: "array",
              items: { type: "array", items: { type: "string" } },
              description: "二维数组公式，例如 [['=A2*1.1']]"
            },
            reason: { type: "string", description: "本次修改的意图描述，用于留痕审计" }
          },
          required: ["address"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_format_cells",
        description: "专业级单元格格式与商业排版美化引擎。支持设置文字字体、字号、加粗、文字色、底纹背景色、水平垂直对齐、行高、数字格式、边框以及单元格合并/取消合并。建议遵循麦肯锡商业报表规范排版。",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "需要格式化的目标区域（例如 'A1:E1' 用于标题合并，'B4:H4' 用于表头，'C5:D20' 用于数值列）" },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc },
            merge: { type: "boolean", description: "【重要】是否将该区域合并为单个大单元格。例如 address='A1:E1', merge=true 可将 5 列合并为整行主标题栏。常配合 horizontalAlignment='center' 实现合并居中。" },
            unmerge: { type: "boolean", description: "是否取消该选区内的合并单元格，将其恢复为独立单元格" },
            horizontalAlignment: {
              type: "string",
              enum: ["left", "center", "right"],
              description: "水平对齐方式: 'left'(靠左，适合文本/大纲)、'center'(居中，适合表头/日期/状态/代码)、'right'(靠右，适合所有纯数字与百分比金额)"
            },
            verticalAlignment: {
              type: "string",
              enum: ["center", "top", "bottom"],
              description: "垂直对齐方式，默认 'center' (垂直居中，美学效果最佳)"
            },
            fontSize: { type: "number", description: "字体大小磅值（例如：大标题设 15-18，副标题设 9-10，表格列头设 10-11，正文数据设 9.5-10）" },
            bold: { type: "boolean", description: "是否加粗文字。大标题、表头、小计合计行建议设为 true；正文数据建议设为 false" },
            fontColor: { type: "string", description: "文字十六进制颜色（例如：标准深灰黑 '#0F172A'，纯白 '#FFFFFF' 配合深色表头，辅助说明淡灰 '#64748B'）" },
            backgroundColor: { type: "string", description: "背景底纹十六进制颜色（例如：商务深蓝表头 '#0F172A'，斑马纹浅灰 '#F8FAFC'，合计行淡灰 '#F1F5F9'，异常警示浅红 '#FEE2E2'）" },
            numberFormat: { type: "string", description: "Excel 数字格式规范代码。例如：千分符整数 '#,##0'，百分比保留两位 '0.00%'，短日期 'yyyy-mm-dd' 或 'm/d'，金额 '¥#,##0.00'。严禁让日期显示为 46249 这类五位数字序列号！" },
            rowHeight: { type: "number", description: "行高磅值。建议：大标题 34-38pt，副标题 20-22pt，表头 26-28pt，普通数据行 20-24pt。严禁设 100pt 以上产生大空白框！" },
            wrapText: { type: "boolean", description: "文本较长时是否自动换行。结论建议区、长表头建议设为 true 配合自适应展开" },
            borders: { type: ["string", "boolean"], description: "边框十六进制颜色（如 '#CBD5E1' 极细浅灰边框）或 true（默认浅灰细边框）。若传 'none' 或 false 则去除边框" }
          },
          required: ["address"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_add_conditional_formatting",
        description: "为指定单元格区域添加智能条件格式。支持阈值告警高亮（如良率低于90%自动标红）、单元格内嵌微型数据条进度显示、双色/三色热力图色阶。",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "应用条件格式的目标区域（例如 'E5:E20' 针对每日良率列，或 'C5:C20' 针对产出量）" },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc },
            ruleType: {
              type: "string",
              enum: ["cell_value", "data_bar", "color_scale"],
              description: "条件格式类型: 'cell_value'(基于单元格数值的阈值比较高亮), 'data_bar'(在单元格内绘制横向条形微型数据条), 'color_scale'(渐变热力图色阶)"
            },
            operator: {
              type: "string",
              enum: ["less_than", "greater_than", "equal", "between"],
              description: "当 ruleType='cell_value' 时的比较关系: 'less_than'(小于指定值), 'greater_than'(大于指定值), 'equal'(等于), 'between'(在两个值之间)"
            },
            formula1: { type: "string", description: "比较阈值1。例如良率低于 90% 告警时填 '0.9'，小于均值 85% 时填 '0.85'" },
            formula2: { type: "string", description: "比较阈值2，仅在 operator='between' 时需要提供" },
            backgroundColor: { type: "string", description: "命中规则时的背景高亮颜色（例如异常浅红告警底色 '#FEE2E2'，优秀达成浅绿底色 '#DCFCE7'）" },
            fontColor: { type: "string", description: "命中规则时的文字高亮颜色（例如异常文字深红 '#991B1B'，优秀文字深绿 '#166534'）" },
            barColor: { type: "string", description: "当 ruleType='data_bar' 时的微型数据条填充颜色（例如科技蓝 '#3B82F6'，森林绿 '#10B981'）" },
            colorScaleMin: { type: "string", description: "当 ruleType='color_scale' 时最小值对应的端点颜色（如浅红 '#F87171'）" },
            colorScaleMax: { type: "string", description: "当 ruleType='color_scale' 时最大值对应的端点颜色（如浅绿 '#4ADE80'）" },
            clearExisting: { type: "boolean", description: "是否在添加前清除该区域现有的条件格式，默认为 false" }
          },
          required: ["address"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_freeze_panes",
        description: "锁定并冻结工作表窗口窗格，使表头在大数据量向下或向右滚动时始终吸顶悬浮可见，极大提升交互可读性。",
        parameters: {
          type: "object",
          properties: {
            freezeRowIndex: { type: "number", description: "冻结分割行号（从 1 开始）。例如 freezeRowIndex: 5 表示将第 1 至 4 行锁定吸顶，用户向下滚动到几千行时表头始终悬浮在最顶部" },
            freezeColumnIndex: { type: "number", description: "冻结分割列号（可选）。例如 freezeColumnIndex: 3 表示将第 1 至 2 列（如 A、B 列）固定吸左，向右滚动时不被移出" },
            unfreeze: { type: "boolean", description: "若设为 true，则解除当前工作表的所有窗口冻结状态" },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_modify_rows_columns",
        description: "对表格的整行或整列执行插入、删除、隐藏或显示操作。适用于在数据块之间插入呼吸空白行、或者隐藏用于中间计算的辅助列。",
        parameters: {
          type: "object",
          properties: {
            targetType: {
              type: "string",
              enum: ["row", "column"],
              description: "操作维度: 'row' (行操作) 或 'column' (列操作)"
            },
            action: {
              type: "string",
              enum: ["insert", "delete", "hide", "unhide"],
              description: "具体执行的动作: 'insert'(插入空行/空列), 'delete'(删除), 'hide'(隐藏不显示), 'unhide'(取消隐藏重新展示)"
            },
            index: { type: "number", description: "起始行号或列号（数字，从 1 开始）。例如在第 10 行插入一行填 index: 10" },
            count: { type: "number", description: "连续操作的行数或列数，默认为 1" },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["targetType", "action", "index"],
          additionalProperties: false
        }
      }
    }
  ];
}

/** Excel（表格）后半区：已发布顺序第 24–38 条。定义顺序即契约顺序，不得重排。 */
export function excelToolDefinitionsAfterAudit(ctx: DefinitionContext): GatewayToolDefinition[] {
  return [
    {
      type: "function",
      function: {
        name: "wps_auto_fit_columns",
        description: "自动调整指定列或全表列宽，防止中文字符被右边框遮挡或显示为'...'省略号",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            address: { type: "string", description: "需要自适应调整的单元格或列区域（例如 'A1:E5' 或 'A:E'），不传则自适应全表已用区域" },
            columnRules: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  colIndex: { type: "number", description: "列索引，从 1 开始（如 1=A, 2=B）" },
                  minWidth: { type: "number", description: "最小列宽，默认 12" },
                  maxWidth: { type: "number", description: "最大列宽，默认 30" },
                  wrapText: { type: "boolean", description: "超出最大宽度时是否自动换行" }
                },
                required: ["colIndex"]
              },
              description: "指定列的列宽调整规则列表"
            }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_capture_sheet_preview",
        description: "高保真捕获工作表或指定单元格区域的渲染截图（返回 Base64 图像），用于多模态视觉模型自检排版效果、检查文字是否截断、是否有过多空白框等",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "工作表名称" },
            address: { type: "string", description: "需要截图的单元格区域（如 'B2:M22'），不填则默认截取全部已用区域" },
            range: { type: "string", description: "address 的同义别名（如 'B2:M22'）" },
            chartName: { type: "string", description: "需要单独捕获截图的原生图表名称（如 'Chart 1'），不填则自动捕获首个图表或工作表区域" },
            name: { type: "string", description: "图表名称同义别名" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_add_chart",
        description: "在当前工作表中创建与数据源动态绑定的原生矢量图表（折线图、簇状柱状图、条形图、饼图、圆环图、柏拉图）。支持指定标题、图例、数据标签以及通过单元格左上角坐标精准排版。",
        parameters: {
          type: "object",
          properties: {
            dataRange: {
              type: "string",
              description: "图表绑定的连续数据源单元格区域（包含行列标题）。例如: 'A4:E19' 或 'B2:C10'。与 dataRanges 选其一传入。"
            },
            sourceAddress: {
              type: "string",
              description: "图表绑定的连续数据源单元格区域别名（同 dataRange）。例如: 'B1:D4'"
            },
            dataRanges: {
              type: "array",
              items: { type: "string" },
              description: "【非连续区域支持】图表绑定的多段非连续单元格区域列表。例如: ['B4:B10', 'E4:E10'] 分别代表日期列与指标列，无需制作辅助列即可直接跨列建图！"
            },
            replaceExisting: {
              type: "boolean",
              description: "是否自动清除该锚点位置上已存在的旧图表（默认 true）。能彻底避免因重复建图导致废图堆叠或报错冲突。"
            },
            chartType: {
              type: "string",
              enum: ["line", "column", "column_clustered", "bar", "bar_clustered", "pie", "doughnut", "pareto", "area", "scatter"],
              description: "图表类型: 'line'(折线图), 'column'/'column_clustered'(簇状柱状图), 'bar'/'bar_clustered'(条形图), 'pie'(饼图), 'doughnut'(圆环图), 'pareto'(柏拉图), 'area'(面积图), 'scatter'(散点图)"
            },
            left: { type: "number", description: "图表距离工作表左侧像素距离" },
            top: { type: "number", description: "图表距离工作表顶部像素距离" },
            width: { type: "number", description: "图表像素宽度，默认 480 像素" },
            height: { type: "number", description: "图表像素高度，默认 280 像素" },
            title: {
              type: "string",
              description: "图表主标题文本。例如: '2026年8月综合良率推移分析'，不传则使用系统默认标题"
            },
            position: {
              type: "object",
              properties: {
                leftCell: {
                  type: "string",
                  description: "图表左上角锚定的单元格坐标。例如: 'G4'，用于让图表与左侧数据表格并列整齐排版，避免遮挡数据"
                },
                width: { type: "number", description: "图表像素宽度，默认 480 像素" },
                height: { type: "number", description: "图表像素高度，默认 280 像素" }
              },
              required: ["leftCell"],
              description: "图表在工作表中的空间放置坐标与长宽尺寸"
            },
            hasLegend: {
              type: "boolean",
              description: "是否显示图例，默认为 true。单系列数据（如单一缺陷占比）可设为 false 提高清爽度"
            },
            hasDataLabels: {
              type: "boolean",
              description: "是否在图表各节点/柱状柱顶端直接标注具体数值，默认为 false"
            },
            smoothLine: {
              type: "boolean",
              description: "【视觉升级】是否启用平滑曲线（仅针对折线图）。例如: smoothLine: true 可将生硬折角转为优雅现代的贝塞尔圆弧曲线，极大提升高管看板审美体验。"
            },
            seriesColors: {
              type: "array",
              items: { type: "string" },
              description: "【色彩系统】按顺序指定各数据系列的十六进制颜色数组。例如: ['#3B82F6', '#10B981', '#F59E0B'] 分别作为第1主系列(如投产量深蓝)、第2系列(如合格量绿色)底色，杜绝系统随机五颜六色。"
            },
            yAxis: {
              type: "object",
              properties: {
                min: {
                  type: "number",
                  description: "数值轴下限最小值。例如良率在 90%~98% 之间波动时，必须传入 min: 0.85，打破 Excel 默认从 0 开始将数据压在顶部的严重可视化缺陷！"
                },
                max: { type: "number", description: "数值轴上限最大值。例如良率上限设为 max: 1.0" },
                step: { type: "number", description: "数值轴主刻度步长。例如 step: 0.05 (以 5% 为一档横向网格线)" },
                numberFormat: { type: "string", description: "坐标轴刻度数字显示格式。例如: '0.0%' 或 '0%' 或 '#,##0'" },
                title: { type: "string", description: "坐标轴标题。例如: '综合良率 (%)'" }
              },
              description: "【核心刻度控制】数值 Y 轴范围与显示格式控制。对于良率、温度等高位指标，必须配置 min 放大波动趋势。"
            },
            seriesSettings: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  seriesIndex: { type: "number", description: "目标数据系列序号（从 1 开始）" },
                  color: { type: "string", description: "该系列的指定十六进制颜色" },
                  smooth: { type: "boolean", description: "该系列是否单独开启平滑线" }
                },
                required: ["seriesIndex"]
              },
              description: "针对特定系列的单项高级定制规则"
            },
            cellRange: {
              type: "string",
              description: "【刚性单元格吸附】图表锚定的单元格范围。例如: 'I8:P20'。由 Office 渲染引擎底层直接将图表咬死在该区域内，实现 100% 完美的行级对齐与等高排版，杜绝像素漂移！"
            },
            startCell: {
              type: "string",
              description: "图表左上角锚定单元格。例如: 'I8'"
            },
            endCell: {
              type: "string",
              description: "图表右下角锚定单元格。例如: 'P20'"
            },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_get_charts",
        description: "读取工作表中的原生图表。默认返回紧凑列表；指定 shapeName/chartIndex/chartTitle 并设置 detail=true 可读取系列与坐标轴详情。",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "目标工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            shapeName: { type: "string", description: "按稳定 Shape 名称精确定位图表" },
            chartIndex: { type: "number", description: "按图表序号定位，从 1 开始，不受图片等非图表 Shape 影响" },
            chartTitle: { type: "string", description: "按图表标题关键词筛选" },
            detail: { type: "boolean", description: "是否读取系列与坐标轴详情，默认 false" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_update_chart",
        description: "更新工作表中已有图表的位置、标题或图例。支持传入 cellRange（如 'I8:P20'）实现实机单元格刚性重定位吸附。",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "目标工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            chartName: { type: "string", description: "图表名称或 ID（如 Chart 1）" },
            name: { type: "string", description: "图表名称别名" },
            shapeName: { type: "string", description: "图表 Shape 名称" },
            title: { type: "string", description: "更新后的图表标题" },
            legendPosition: { type: "string", enum: ["Top", "Bottom", "Left", "Right", "Corner"], description: "图例位置" },
            cellRange: { type: "string", description: "更新图表吸附的单元格范围，例如 'I8:P20'" },
            startCell: { type: "string", description: "图表起始单元格" },
            endCell: { type: "string", description: "图表结束单元格" },
            left: { type: "number", description: "左侧像素" },
            top: { type: "number", description: "顶部像素" },
            width: { type: "number", description: "像素宽度" },
            height: { type: "number", description: "像素高度" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_delete_chart",
        description: "从工作表中安全删除指定的原生图表或一键清空全部图表。彻底解决旧图表无法清除、留下空白残缺边框的痛点。",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "目标工作表名称，不传则默认为当前活动工作表" },
            chartName: { type: "string", description: "需要删除的图表名称或 ID（如 {GUID} 或 Chart 1）" },
            name: { type: "string", description: "图表名称同义别名" },
            shapeName: { type: "string", description: "需要删除的图表稳定 Shape 名称，推荐使用 wps_get_charts 返回值" },
            chartTitle: { type: "string", description: "需要删除的图表标题关键词匹配。例如: '8月投产与良率推移趋势'" },
            leftCell: { type: "string", description: "图表左上角锚定的单元格坐标。例如: 'M57' 或 'I4'，用于精准删除特定位置的废图或空图表" },
            chartIndex: { type: "number", description: "图表序号（从 1 开始）" },
            clearAll: { type: "boolean", description: "是否清空当前工作表中的所有图表，默认为 false" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_create_pivot_table",
        description: "直接针对数千行原始明细数据，一键聚合生成多维交叉数据透视表。无需手工编写复杂的 SUMIFS/COUNTIFS 公式即可完成快速交叉分析。",
        parameters: {
          type: "object",
          properties: {
            sourceRange: {
              type: "string",
              description: "原始明细数据区域（必须包含第一行表头）。例如: 'A1:K5422' 或配合 sourceSheetName 跨表指定"
            },
            destCell: {
              type: "string",
              description: "透视表左上角起始放置的单元格地址。例如: 'B4'，建议留出顶部和左侧边距"
            },
            sourceSheetName: {
              type: "string",
              description: "数据源所在工作表名称。例如: '明细数据'，不传则默认当前工作表"
            },
            destSheetName: {
              type: "string",
              description: "透视表放置的目标工作表名称。例如: '交叉透视分析'，建议新建独立工作表放置"
            },
            rowFields: {
              type: "array",
              items: { type: "string" },
              description: "放置在行维度的字段名数组。例如: ['测试日期', '组件等级']"
            },
            columnFields: {
              type: "array",
              items: { type: "string" },
              description: "放置在列维度的字段名数组。例如: ['功率档位']"
            },
            dataFields: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  fieldName: { type: "string", description: "要汇总计算的源字段名，如 '实测功率(W)' 或 '条形码'" },
                  summaryFunction: {
                    type: "string",
                    enum: ["sum", "count", "average", "max", "min"],
                    description: "聚合函数: 'sum'(求和), 'count'(计数), 'average'(求平均), 'max'(最大值), 'min'(最小值)"
                  },
                  caption: { type: "string", description: "透视表表头显示的自定义名称，例如 '平均功率(W)'" }
                },
                required: ["fieldName"]
              },
              description: "需要聚合统计的数据指标字段列表"
            },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["sourceRange", "destCell"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_set_filter_and_sort",
        description: "为数据表格启用/关闭自动筛选下拉三角漏斗，并按指定列执行单列或多列升降序排列。大幅提升终端用户的查阅与交互体验。",
        parameters: {
          type: "object",
          properties: {
            range: {
              type: "string",
              description: "目标表格区域（包含表头）。例如: 'A4:G50' 或 'A1:E20'"
            },
            enableAutoFilter: {
              type: "boolean",
              description: "是否开启表头筛选下拉三角按钮。设为 true 开启，false 关闭"
            },
            sortRules: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  colIndex: { type: "number", description: "排序基准列索引（数字，从 1 开始，1=第1列/A列）" },
                  order: {
                    type: "string",
                    enum: ["asc", "desc"],
                    description: "排序方向: 'asc'(升序，从小到大/A到Z), 'desc'(降序，从大到小/Z到A)"
                  }
                },
                required: ["colIndex", "order"]
              },
              description: "排序规则列表。支持多级排序，列表第一项为主排序列，第二项为次排序列"
            },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["range"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_set_data_validation",
        description: "为指定单元格区域设置数据有效性验证（如下拉选择菜单、数值区间限制）。防止人为录入错误，并可配置选中时的提示气泡与输入非法报错弹窗。",
        parameters: {
          type: "object",
          properties: {
            address: {
              type: "string",
              description: "目标单元格区域。例如: 'E5:E50' 针对状态列，或 'C5:C20' 针对合格率输入"
            },
            validationType: {
              type: "string",
              enum: ["list", "number_range"],
              description: "验证类型: 'list'(单元格下拉选择菜单列表), 'number_range'(数值范围限定)"
            },
            listItems: {
              type: "array",
              items: { type: "string" },
              description: "当 validationType='list' 时的可用候选项列表。例如: ['已通过', '待复测', '已报废']"
            },
            operator: {
              type: "string",
              enum: ["between", "greater_than", "less_than", "equal"],
              description: "当 validationType='number_range' 时的比较条件，默认为 'between'"
            },
            minVal: { type: "number", description: "数值范围下限。例如: 0 或 1" },
            maxVal: { type: "number", description: "数值范围上限。例如: 100" },
            promptTitle: { type: "string", description: "用户选中单元格时浮动提示框的标题。例如: '状态选择提示'" },
            promptMessage: { type: "string", description: "用户选中单元格时浮动的提示文本。例如: '请从下拉菜单中选择审核结论'" },
            errorTitle: { type: "string", description: "输入非法值时的警告弹窗标题" },
            errorMessage: { type: "string", description: "输入非法值时的警告错误提示内容" },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["address", "validationType"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_manage_sheet",
        description: "工作表综合管理引擎。支持对工作表进行重命名、左右移动调整标签顺序、设置底部标签颜色高亮（如将重要汇总表标红）以及锁定/解锁工作表保护。",
        parameters: {
          type: "object",
          properties: {
            sheetName: {
              type: "string",
              description: "需要操作的目标工作表原名称。例如: 'Sheet1' 或 '临时表'"
            },
            action: {
              type: "string",
              enum: ["rename", "move", "tab_color", "protect", "unprotect"],
              description: "执行的管理操作: 'rename'(重命名), 'move'(调整位置顺序), 'tab_color'(设置工作表标签底色), 'protect'(锁定保护工作表), 'unprotect'(解除锁定保护)"
            },
            newName: {
              type: "string",
              description: "重命名时的新名称（仅当 action='rename' 时需要）。例如: '8月核心经营看板'"
            },
            targetIndex: {
              type: "number",
              description: "移动调整的目标位置序号（从 1 开始，仅当 action='move' 时需要）。例如: 1 表示移到最前第一张标签"
            },
            color: {
              type: "string",
              description: "标签底色十六进制颜色代码（仅当 action='tab_color' 时需要）。例如: '#EF4444'(醒目红)，'#3B82F6'(业务蓝)，'#10B981'(通过绿)"
            },
            password: {
              type: "string",
              description: "保护/解除保护工作表时的密码（可选）。不填则为无密码保护"
            },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["sheetName", "action"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_manage_rows_and_columns",
        description: "在工作表中批量插入、删除、隐藏、取消隐藏整行或整列，或精准设置行高/列宽。例如在表头下方插入汇总空白行、折叠隐藏明细列等。",
        parameters: {
          type: "object",
          properties: {
            targetType: {
              type: "string",
              enum: ["row", "column"],
              description: "操作维度: 'row'(行), 'column'(列)"
            },
            action: {
              type: "string",
              enum: ["insert", "delete", "hide", "unhide", "set_size"],
              description: "执行操作: 'insert'(插入), 'delete'(删除), 'hide'(隐藏), 'unhide'(取消隐藏), 'set_size'(设置行高或列宽)"
            },
            index: {
              description: "起始行号（数字，如 5 表示第5行）或列标识（数字 2 或字母 'B' 表示第B列）"
            },
            count: {
              type: "number",
              description: "连续操作的行数或列数，默认为 1"
            },
            size: {
              type: "number",
              description: "具体的尺寸值（仅当 action='set_size' 时需要）。若 targetType='row' 为磅值高度（如 24）；若 targetType='column' 为字符列宽（如 15）"
            },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["targetType", "action", "index"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_manage_cell_comments",
        description: "单元格原生批注与审阅备注管理引擎。支持为单元格添加、读取、删除或清空黄色气泡批注。非常适合 AI 作为质检/财务审核员在异常单元格留下审核依据与批注。",
        parameters: {
          type: "object",
          properties: {
            address: {
              type: "string",
              description: "目标单元格或区域坐标。例如: 'C5' 或 'E5:E20'"
            },
            action: {
              type: "string",
              enum: ["add", "read", "delete", "clear_all"],
              description: "操作类型: 'add'(添加/更新批注), 'read'(读取指定区域的全部批注), 'delete'(删除指定单元格批注), 'clear_all'(清空全表所有批注)"
            },
            text: {
              type: "string",
              description: "批注正文内容（仅当 action='add' 时需要）。例如: '审核说明: 8月20日良率低于 90%，经核实为设备温控报警，建议复测'"
            },
            author: {
              type: "string",
              description: "批注作者签名，默认为 'AI 智能审核'"
            },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["action"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_find_and_replace",
        description: "工作表全局高速查找与精准定位替换。支持在海量数据中瞬间检索出所有包含特定错误码（如 #VALUE!、#N/A）、特定状态（如 '待复测'、'未通过'）或关键词的单元格坐标列表，并支持一键批量替换。",
        parameters: {
          type: "object",
          properties: {
            searchQuery: {
              description: "要查找的目标关键词、数值或错误标识。例如: '#VALUE!'、'待复测'、0"
            },
            replaceText: {
              type: "string",
              description: "替换后的新内容（可选）。若传入则同时执行替换修改，不传则仅执行纯查找定位"
            },
            matchCase: {
              type: "boolean",
              description: "是否严格区分大小写，默认为 false"
            },
            matchEntireCell: {
              type: "boolean",
              description: "是否全字完全匹配单元格，默认为 false（默认包含即匹配）"
            },
            searchRange: {
              type: "string",
              description: "限定检索的单元格区域（例如 'A1:K500'）。不传则自动搜索当前工作表的有效使用数据区域"
            },
            maxResults: {
              type: "number",
              description: "最多返回的结果数量上限，默认为 50"
            },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["searchQuery"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_duplicate_sheet",
        description: "工作表完整克隆与模板复制。将现有工作表（100% 完整保留所有复杂三线表样式、公式、条件格式与图表）克隆复制出一张新表。常用于基于《月度模板》一键派生《9月报表》。",
        parameters: {
          type: "object",
          properties: {
            sourceSheetName: {
              type: "string",
              description: "被克隆的源工作表名称。例如: '综合看板模板' 或 '8月数据'"
            },
            newSheetName: {
              type: "string",
              description: "克隆生成的新工作表名称。例如: '9月综合分析看板'"
            },
            position: {
              type: "string",
              enum: ["after", "before", "end"],
              description: "新工作表的放置位置: 'after'(紧随源工作表之后，默认), 'before'(源工作表之前), 'end'(工作簿最末尾)"
            },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["sourceSheetName", "newSheetName"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_save_workbook",
        description: "保存当前打开的工作簿（支持 WPS 表格与 Microsoft Excel），避免改动仅停留在内存中未落盘。",
        parameters: {
          type: "object",
          properties: {
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          additionalProperties: false
        }
      }
    }
  ];
}
