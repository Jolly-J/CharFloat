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
        description: (ctx.activeWbHint
          ? `【当前WPS已在线打开工作簿: ${ctx.activeWbHint}】`
          : "") +
          "获取在 WPS 中打开的 Excel 工作簿概览、已打开文件清单(openWorkbooks)、工作表列表与选区坐标。" +
          (ctx.activeWbHint ? `当前已打开 [${ctx.activeWbHint}]。` : "") +
          "严禁在磁盘搜索文件！hasOpenWorkbook=false 只表示未解析到目标工作簿（锁目标可能已关闭），不代表没有工作簿打开——判断“打开了哪些文件”请读 openWorkbooks。选型：wps_* 只走 WPS 表格（不传 host）；同名 excel_* 是跨宿主统一入口，必传 host='wps'|'microsoft'。目标已在 WPS 中打开时二者等价，选一个调用即可、不要重复调用；目标在 Microsoft Excel 时必须用 excel_*。",
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
        description: "按需获取指定工作表的数据边界(UsedRange)与前 3 行表头样本，用于在读取整表前判断结构与数据量。返回体还含 **sheetState**（工作表级读回，用于写完自检）：protection(保护状态)、tabColor(标签色)、autoFilter(筛选模式与实际范围)、freezePanes(冻结行列，**仅目标表处于活动状态时可读**，否则 available:false 并说明)、conditionalFormats(已用区域上的条件格式规则摘要，最多 50 条)。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
        description: "在指定工作簿中新建工作表并激活。新建后需 wps_save_workbook 才会落盘。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
        description: "删除指定工作表：破坏性且不可回滚，删除前先确认表名。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
        description: "切片读取指定区域(如 A1:C10)的单元格值与公式。只读，不修改工作簿。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
        description: "读取指定区域的单元格样式。默认 summary 仅返回区域级样式摘要；cells 模式逐格返回并受 maxCells 限制。写入格式后用本工具读回验证。可通过 include 追加：**validation**（数据有效性读回：类型名/操作符/公式1-2/是否忽略空值/是否显示下拉/提示与报错文案）——检查下拉或范围校验是否真的生效；**validationViolations**（数据有效性**违规定位**）——逐格比对规则并列出**取值越界的单元格**（地址+值+命中的规则+期望），同时给出规则清单与无法判定的单元格；区域大时只扫描前 maxCells 个（默认 300，上限 2000），截断时返回 truncated=true 并在 warnings 说明。判定不了的一律进 unevaluated，**不会当作通过**。**宿主差异**：validation / validationViolations 仅在 WPS 表格侧实现（host=microsoft 会忽略 include，只返回基础样式，不要据此判断校验是否生效）。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "区域地址，例如 'A1:E20'" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            mode: { type: "string", enum: ["summary", "cells"], description: "返回模式，默认 summary" },
            include: {
              type: "array",
              items: { type: "string", enum: ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea", "borders", "validation", "validationViolations"] },
              description: "只返回指定样式字段；不传时返回常用字段。validation=只读回规则；validationViolations=规则+违规定位（两者可同时给）"
            },
            maxCells: { type: "number", description: "cells 模式最多展开的单元格数，默认 100，最大 500；include 含 validationViolations 时同时作为违规扫描上限（默认 300，上限 2000）" }
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
        description: "按单元格的显示值与**公式文本**搜索包含指定关键字的单元格坐标（命中项用 matchedIn 标明是 value 还是 formula）。可选 address 限定检索范围，不传则搜整表已用区域。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
        description: "原地修改表格中的数值或公式，返回 auditId（可用 wps_rollback 回滚值/公式）。约束：可同时传 values 与 formulas——values 先写、formulas 中的**非空项**随后覆盖对应单元格；formulas 中的空项(''/null)表示**不动该单元格**（不会清空）。清空单元格请传 values 的 null 矩阵（本工具没有 clear_range）。字符串日期如 '2026-01' 会被宿主转成日期序列号，需要保持文本时另行设置 numberFormat。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "目标区域地址，例如 'C2:C10'" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            values: {
              type: "array",
              items: { type: "array", items: {} },
              description: "二维数值矩阵（按行给出，形状必须与 address 一致）；清空某格传 null"
            },
            formulas: {
              type: "array",
              items: { type: "array", items: { type: ["string", "null"] } },
              description: "二维公式矩阵，例如 [['=A2*1.1']]；空项 '' 或 null 表示该单元格不改公式（同批 values 写入的值保留），不会清空单元格"
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
        description: "设置单元格字体、字号、加粗、文字色、底色、对齐、行高、数字格式、边框与合并/取消合并，写入后可用 wps_get_range_styles 读回。注意：borders 传 'none' 或 false 当前不会去除边框（只接受颜色值或 true 画细边框）。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
              description: "垂直对齐方式: 'center'(垂直居中，默认), 'top'(顶端对齐), 'bottom'(底端对齐)"
            },
            fontSize: { type: "number", description: "字体大小磅值（例如：大标题设 15-18，副标题设 9-10，表格列头设 10-11，正文数据设 9.5-10）" },
            bold: { type: "boolean", description: "是否加粗文字。大标题、表头、小计合计行建议设为 true；正文数据建议设为 false" },
            fontColor: { type: "string", description: "文字十六进制颜色（例如：标准深灰黑 '#0F172A'，纯白 '#FFFFFF' 配合深色表头，辅助说明淡灰 '#64748B'）" },
            backgroundColor: { type: "string", description: "背景底纹十六进制颜色（例如：商务深蓝表头 '#0F172A'，斑马纹浅灰 '#F8FAFC'，合计行淡灰 '#F1F5F9'，异常警示浅红 '#FEE2E2'）" },
            numberFormat: { type: "string", description: "Excel 数字格式规范代码。例如：千分符整数 '#,##0'，百分比保留两位 '0.00%'，短日期 'yyyy-mm-dd' 或 'm/d'，金额 '¥#,##0.00'。不要留下 46249 这类五位序列号显示。" },
            rowHeight: { type: "number", description: "行高磅值。建议：大标题 34-38pt，副标题 20-22pt，表头 26-28pt，普通数据行 20-24pt；不要超过 100pt，会产生大空白框。" },
            wrapText: { type: "boolean", description: "文本较长时是否自动换行。结论建议区、长表头建议设为 true 配合自适应展开" },
            borders: { type: ["string", "boolean"], description: "边框：十六进制颜色（如 '#CBD5E1'）或 true（默认浅灰细边框）。当前实现跳过 'none'/false，传它们不会去除已有边框。" }
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
        description: "为指定区域添加条件格式。可用 ruleType：cell_value（阈值高亮）、data_bar（数据条）、color_scale（双色色阶）、**icon_set（图标集，如红黄绿灯）**、**top10（前/后 N 名或百分比）**、**duplicate_values / unique_values（重复值/唯一值）**、**formula（公式规则）**、**text_contains（文字包含，内部用 SEARCH 公式实现）**、clear（清除该区域全部规则）。返回体带 appliedConditionCount 与 conditions（逐条规则的类型/启用/优先级/公式），**写完即可读回核对**，不需要另用脚本探测。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["apply", "read"], description: "apply（默认）写入并读回核对；read 只读回当前状态、不做任何修改" },
            address: { type: "string", description: "应用条件格式的目标区域（例如 'E5:E20' 针对每日良率列，或 'C5:C20' 针对产出量）" },
            sheetName: { type: "string", description: "工作表名称，不传则默认为当前活动工作表" },
            workbookName: { type: "string", description: ctx.wbDesc },
            ruleType: {
              type: "string",
              enum: ["cell_value", "data_bar", "color_scale", "icon_set", "top10", "duplicate_values", "unique_values", "formula", "text_contains", "clear"],
              description: "条件格式类型: 'cell_value' 阈值高亮 / 'data_bar' 数据条 / 'color_scale' 色阶 / 'icon_set' 图标集（红黄绿灯等）/ 'top10' 前N名 / 'duplicate_values' 重复值 / 'unique_values' 唯一值 / 'formula' 公式规则 / 'text_contains' 文字包含 / 'clear' 清除全部规则"
            },
            iconSet: {
              type: "string",
              description: "ruleType=icon_set 时用。常用：'3_traffic_lights'(红黄绿灯，默认)、'3_arrows'、'3_flags'、'3_symbols'、'4_arrows'、'5_arrows'、'5_quarters'、'5_ratings' 等 18 种"
            },
            iconThresholds: {
              type: "array",
              description: "图标集的阈值（可选），从第 2 档起逐档给；元素可为数字（只设值）或对象 {type,operator,value}"
            },
            topBottom: { type: "number", description: "ruleType=top10 时用：1=前 N，2=后 N" },
            topRank: { type: "number", description: "ruleType=top10 时用：N 的取值" },
            topPercent: { type: "boolean", description: "ruleType=top10 时用：true 表示按百分比而不是名次" },
            containsText: { type: "string", description: "ruleType=text_contains 时用：要高亮的文字" },
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
        description: "冻结工作表窗格，使表头在下滚/右滚时保持可见。freezeRowIndex、freezeColumnIndex、unfreeze 至少传一项，否则本调用不产生任何变化；行列用的是同一套-1 约定（freezeRowIndex=5 冻结前 4 行）。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["apply", "read"], description: "apply（默认）写入并读回核对；read 只读回当前状态、不做任何修改" },
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
        description: "对整行或整列执行插入、删除、隐藏、取消隐藏。与 wps_manage_rows_and_columns 的区别：本工具不支持设置行高/列宽(set_size)，参数集也不含 size；需要 set_size 请用 wps_manage_rows_and_columns。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
        description: "自动调整列宽，避免中文被右边框遮挡或显示为'...'。传 columnRules 时按规则逐列处理（含 minWidth/maxWidth 与超宽换行）；不传 columnRules 时按整表已用区域逐列自适应并返回实际列宽。address 当前被 WPS 宿主忽略，自适应范围不受它约束。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            address: { type: "string", description: "（兼容参数，当前 WPS 实现忽略）需要自适应调整的单元格或列区域，例如 'A1:E5'；实际范围由 columnRules 或整表已用区域决定" },
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
        description: "捕获工作表或指定区域的渲染截图，用于视觉自检。参数优先级：address 优先于别名 range，都不传取整表已用区域；chartName（别名 name）在 WPS 宿主被忽略，仍按区域截图。返回 {success, workbookName, sheetName, address, imageBase64, imageMimeType, imageSizeBytes}；截图失败时可能只有 message 而没有 imageBase64，用前先确认该字段存在。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "工作表名称" },
            address: { type: "string", description: "要截图的单元格区域（如 'B2:M22'）；与 range 同时传入时以本参数为准，都不传则截取已用区域" },
            range: { type: "string", description: "address 的同义别名（如 'B2:M22'），address 优先" },
            chartName: { type: "string", description: "图表名称（如 'Chart 1'）；当前 WPS 宿主忽略该参数，Microsoft 宿主可用" },
            name: { type: "string", description: "chartName 的同义别名" },
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
        description: "创建绑定数据源的原生矢量图表。dataRange（别名 sourceAddress）与 dataRanges 必传其一，都不传宿主直接报错。定位优先级：顶层 left/top/width/height（像素）> position.leftCell / startCell / cellRange（单元格锚点）> endCell（用锚点到 endCell 的矩形作尺寸）；都不传则落在默认 360/40，多图会重叠。seriesColors 对单系列是逐点染色（会得到彩虹柱），单系列请只传一个颜色。chartType 枚举外的值直接报错；建图后读回真实 ChartType，与请求不一致会返回 warnings，可用 wps_get_charts(detail=true) 复核。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            dataRange: {
              type: "string",
              description: "图表绑定的连续数据源单元格区域（含行列标题），例如 'A4:E19'。与 dataRanges 选其一（两处都传时以 dataRange 为准）。"
            },
            sourceAddress: {
              type: "string",
              description: "dataRange 的同义别名，例如 'B1:D4'"
            },
            dataRanges: {
              type: "array",
              items: { type: "string" },
              description: "多段非连续数据源区域列表，例如 ['B4:B10', 'E4:E10']；不传 dataRange 时生效"
            },
            replaceExisting: {
              type: "boolean",
              description: "是否先清除锚点 position.leftCell 附近（±25px）的旧图表，默认 true；不传 position 时不清理"
            },
            chartType: {
              type: "string",
              enum: ["line", "column", "column_clustered", "bar", "bar_clustered", "pie", "doughnut", "pareto", "area", "scatter"],
              description: "图表类型（必传其一的数据源之外的唯一必选项，不传默认 column_clustered）: 'line'(折线图), 'column'/'column_clustered'(簇状柱状图), 'bar'/'bar_clustered'(条形图), 'pie'(饼图), 'doughnut'(圆环图), 'pareto'(柏拉图), 'area'(面积图), 'scatter'(散点图)。枚举外的值会报错，不会静默降级。"
            },
            left: { type: "number", description: "图表距工作表左侧像素距离；当前 WPS 宿主忽略本参数，请改用 position.leftCell" },
            top: { type: "number", description: "图表距工作表顶部像素距离；当前 WPS 宿主忽略本参数，请改用 position.leftCell" },
            width: { type: "number", description: "图表像素宽度，默认 480；当前 WPS 宿主忽略本参数，请改用 position.width" },
            height: { type: "number", description: "图表像素高度，默认 280；当前 WPS 宿主忽略本参数，请改用 position.height" },
            title: {
              type: "string",
              description: "图表主标题文本，例如 '2026年8月综合良率推移分析'"
            },
            position: {
              type: "object",
              properties: {
                leftCell: {
                  type: "string",
                  description: "图表左上角锚定的单元格坐标，例如 'G4'；WPS 宿主按该单元格的 Left/Top 像素定位"
                },
                width: { type: "number", description: "图表像素宽度，默认 480" },
                height: { type: "number", description: "图表像素高度，默认 280" }
              },
              required: ["leftCell"],
              description: "图表放置位置（WPS 宿主唯一生效的定位参数）。不传时落在 360/40，多图会叠在一起。"
            },
            hasLegend: {
              type: "boolean",
              description: "是否显示图例，默认 true；单系列数据可设 false"
            },
            hasDataLabels: {
              type: "boolean",
              description: "是否在数据点/柱顶标注数值，默认 false"
            },
            smoothLine: {
              type: "boolean",
              description: "折线图是否启用平滑曲线，默认 false"
            },
            seriesColors: {
              type: "array",
              items: { type: "string" },
              description: "按顺序指定各数据系列的十六进制颜色。注意：宿主对单系列图表会把数组当作逐点颜色，传多个颜色会得到彩虹柱——单系列请只传一个或不传。"
            },
            yAxis: {
              type: "object",
              properties: {
                min: {
                  type: "number",
                  description: "数值轴下限，例如良率在 90%~98% 之间波动时传 0.85，避免默认从 0 起把差异压平"
                },
                max: { type: "number", description: "数值轴上限，例如 1.0" },
                step: { type: "number", description: "数值轴主刻度步长，例如 0.05" },
                numberFormat: { type: "string", description: "坐标轴刻度数字格式，例如 '0.0%'、'0%'、'#,##0'" },
                title: { type: "string", description: "坐标轴标题，例如 '综合良率 (%)'" }
              },
              description: "数值 Y 轴范围与显示格式"
            },
            seriesSettings: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  seriesIndex: { type: "number", description: "目标数据系列序号（从 1 开始）" },
                  color: { type: "string", description: "该系列的指定十六进制颜色" },
                  smooth: { type: "boolean", description: "该系列是否单独开启平滑线" },
                  type: { type: "string", description: "该系列的图表类型 → 做**组合图**（如柱+折线同图）：line / line_markers / column / area / scatter / bar，或直接传 xlChartType 整数" },
                  axisGroup: { type: "number", description: "坐标轴组：1 主轴（默认）/ 2 次轴（组合图里把折线放次轴）" }
                },
                required: ["seriesIndex"]
              },
              description: "针对特定系列的单项高级定制规则"
            },
            cellRange: {
              type: "string",
              description: "（当前 WPS 宿主忽略）期望图表锚定的单元格范围，例如 'I8:P20'。请改用 position.leftCell + position.width/height。"
            },
            startCell: {
              type: "string",
              description: "（当前 WPS 宿主忽略）图表左上角锚定单元格，例如 'I8'；请改用 position.leftCell"
            },
            endCell: {
              type: "string",
              description: "（当前 WPS 宿主忽略）图表右下角锚定单元格，例如 'P20'"
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
        description: "读取工作表中的原生图表。默认返回紧凑列表；指定 shapeName/chartIndex/chartTitle 并设置 detail=true 可读取系列与坐标轴详情。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "目标工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            shapeName: { type: "string", description: "按稳定 Shape 名称精确定位图表" },
            chartIndex: { type: "number", description: "按图表序号定位，从 1 开始，不受图片等非图表 Shape 影响" },
            chartTitle: { type: "string", description: "按图表标题关键词筛选" },
            detail: { type: "boolean", description: "是否读取系列与坐标轴详情，默认 false" },
            shapeSelfTest: { type: "boolean", description: "诊断用：在 Office.js 任务窗格内跑一遍矢量形状能力自检并返回逐项结果（返回体含 addon.version 与 capabilities，可用于确认窗格是否已换到最新代码）" },
            selfTestKeep: { type: "boolean", description: "诊断用：把自检明细留在 _cap08_shapetest 工作表里供读回" }
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
        description: "更新已有图表的位置、标题或图例。宿主限制：WPS 未实现该操作，调用会在执行前被拒绝且不改动文档；WPS 上请用 wps_execute_script 改 chart 对象，或改用 host=microsoft。定位必传其一：chartName、name、shapeName；只传 title/legendPosition 不会命中任何图表。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）；WPS 宿主上该操作会被告知未实现。",
        parameters: {
          type: "object",
          properties: {
            position: { type: "object", description: "移动/改尺寸（磅）：{ left, top, width, height } 绝对定位；{ topDelta, leftDelta } 相对位移（如往下挪 30 磅传 {'topDelta': 30}）；{ leftCell } 左上角对齐到单元格" },
            fontName: { type: "string", description: "图表字体（标题/图例/坐标轴/数据标签统一设置），如「微软雅黑」" },
            legendPosition: { type: "string", enum: ["right", "left", "top", "bottom"], description: "图例位置" },
            dataLabelColorMatchesSeries: { type: "boolean", description: "数据标签文字颜色自动匹配所属柱/点的颜色" },
            chartIndex: { type: "number", description: "或用图表序号定位（从 1 开始）；多个图表时与 chartName 二选一" },
            chartType: { type: "string", description: "改图表类型：column / bar / line / pie / scatter / area，或直接传 xlChartType 整数" },
            hasLegend: { type: "boolean", description: "是否显示图例" },
            showDataLabels: { type: "boolean", description: "是否在数据点/柱顶显示数值标签" },
            sheetName: { type: "string", description: "目标工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
            chartName: { type: "string", description: "图表名称或 ID（如 {GUID} 或 Chart 1），选其一作为定位参数" },
            name: { type: "string", description: "chartName 的同义别名" },
            shapeName: { type: "string", description: "图表 Shape 名称（取自 wps_get_charts），推荐用它精确定位" },
            title: { type: "string", description: "更新后的图表标题" },
            cellRange: { type: "string", description: "期望图表吸附的单元格范围，例如 'I8:P20'" },
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
        description: "删除工作表中的原生图表。必传其一用于定位：shapeName（推荐）、chartName/name、chartTitle、chartIndex、leftCell；或用 clearAll: true 显式清空全部图表。leftCell 是像素邻近匹配（±30px）：命中多于一张时本工具会拒绝执行并列出候选，一张都没命中也会报错，不会静默返回成功。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            sheetName: { type: "string", description: "目标工作表名称，不传则默认为当前活动工作表" },
            chartName: { type: "string", description: "图表名称或 ID（如 {GUID} 或 Chart 1）" },
            name: { type: "string", description: "chartName 的同义别名" },
            shapeName: { type: "string", description: "图表稳定 Shape 名称，推荐使用 wps_get_charts 返回值" },
            chartTitle: { type: "string", description: "按图表标题关键词匹配" },
            leftCell: { type: "string", description: "图表左上角锚点单元格，例如 'M57'。按像素邻近（±30px）匹配，命中多张时拒绝执行——不要用它删叠放的图表。" },
            chartIndex: { type: "number", description: "图表序号（从 1 开始，只数图表、不数图片）" },
            clearAll: { type: "boolean", description: "显式确认清空当前工作表的全部图表，默认 false；这是破坏性操作，仅在确实要全删时传 true" },
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
        description: "根据明细数据区域生成数据透视表。要求：sourceRange 必须含表头且与 destCell 一起必传；destSheetName 指定的表必须已存在（不存在时先 wps_create_sheet），否则直接报错；建表后如显示为空，需在宿主中刷新透视表。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["apply", "read"], description: "apply（默认）写入并读回核对；read 只读回当前状态、不做任何修改" },
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
          required: [],   // read 模式不需要 sourceRange/destCell；写模式缺参时宿主会明确报错
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_set_filter_and_sort",
        description: "为表格启用/关闭自动筛选，并按指定列排序。sortRules 的 colIndex 是**区域内相对列号**（1=区域第 1 列）。排序后会读回校验，未真正生效即报错（返回 sortApplied.attempts）。筛选范围不受 range 严格约束：宿主会扩展到相邻数据块，返回体 appliedFilterRange 是实际范围，与传入不一致时给出 warnings，需要严格范围请用空行隔离数据块。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["apply", "read"], description: "apply（默认）写入并读回核对；read 只读回当前筛选状态（筛选范围与各列条件），此时 range 可省略" },
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
          required: [],   // read 模式不需要 range；写模式缺参时宿主会明确报错
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_set_data_validation",
        description: "为指定区域设置数据有效性（下拉列表或数值区间），可配选中提示与非法输入报错。**先校验参数、后动手**：参数不合法（list 缺 listItems、number_range 缺 minVal/maxVal、operator 或 validationType 非法）会在**任何修改发生前**报错，原有校验保持不变；宿主写入失败时尽力写回原规则并如实报告是否恢复。写入成功后**读回核对**（类型/运算符/上下限/候选项/下拉开关），不一致则 success=false 并在 warnings 给出请求值与宿主读回值。读回规则用 wps_get_range_styles 的 include:['validation']；要找**存量数据里哪些单元格越界**用 include:['validationViolations']。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["apply", "read"], description: "apply（默认）写入并读回核对；read 只读回当前状态、不做任何修改" },
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
              description: "当 validationType='number_range' 时的比较条件: 'between'(介于 minVal 与 maxVal，默认), 'greater_than'(大于 minVal), 'less_than'(小于 maxVal), 'equal'(等于 minVal)"
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
          required: ["address"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_manage_sheet",
        description: "工作表管理：重命名、调整标签顺序、设置标签颜色、锁定/解锁工作表保护。sheetName 与 action 均必传。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）；但 move/tab_color/protect/unprotect 四个 action 在 Microsoft 宿主尚未实现。",
        parameters: {
          type: "object",
          properties: {
            sheetName: {
              type: "string",
              description: "需要操作的目标工作表原名称。例如: 'Sheet1' 或 '临时表'"
            },
            action: {
              type: "string",
              enum: ["rename", "move", "tab_color", "protect", "unprotect", "read"],
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
        description: "批量插入、删除、隐藏、取消隐藏整行或整列，或设置行高/列宽(action='set_size')。与 wps_modify_rows_columns 的区别：本工具多 set_size 与 size 参数。targetType、action、index 必传。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
              description: "执行操作: 'insert'(插入), 'delete'(删除，破坏性), 'hide'(隐藏), 'unhide'(取消隐藏), 'set_size'(设置行高或列宽，需 size)"
            },
            index: {
              type: ["number", "string"],
              description: "起始行号（数字，如 5）或列标识（数字 2 或字母 'B' 表示第 B 列）"
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
        description: "单元格批注管理：添加/更新、读取、删除指定批注，或清空全表批注。action 必传；'add' 需 address + text。注意：宿主批注作者恒为当前 WPS 用户，author 只作为签名拼入正文，不会写入作者字段。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
              description: "批注签名：会以 '作者:\\n正文' 的形式拼进批注内容。宿主批注的真实作者恒为当前 WPS 用户，不要指望本参数改变作者字段。"
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
        description: "在工作表中查找并可选批量替换。searchQuery 必传；同时匹配单元格显示值与**公式文本**（results[].matchedIn 标明命中在哪一侧；命中公式时写回公式位）。返回的 results[].row/col 是相对 searchRange 的偏移量，不是工作表绝对行列号。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            searchQuery: {
              type: ["string", "number"],
              description: "要查找的关键词、数值或错误标识，例如 '#VALUE!'、'待复测'、0"
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
        description: "克隆现有工作表（保留样式、公式、条件格式与图表）生成一张新表；源表名与新表名必传。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
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
        description: "保存当前打开的工作簿。保存是**工作簿级**操作：会把该工作簿当前内存状态整体落盘，包含其他会话尚未完成的中间结果。参数集只有 workbookName（不接受 sheetName，传了会报“未知参数”）；不传时按锁定目标/活动工作簿保存。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_clear_range",
        description: "清空指定区域的**内容**（值、公式、批注一并清除，保留格式）。这是唯一可靠的清空方式——用 patch_cells 写 null 矩阵在区域含公式或合并单元格时会残留。破坏性且不可回滚：先用 read_range 确认范围，并留意返回体里的 clearedCells。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "要清空的区域，如 'A1:C10'" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["address"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_get_style_token",
        description: "读取原表的**设计语言**，用于让新写的内容与既有表格风格一致（做看板/报表前先调它，不要自己臆造配色）。返回：① 取样格真实样式 `sampledCell`（字体/字号/粗斜体/字色/底色/数字格式/对齐/行高列宽/下框线）② `fonts` 字体层级 title/header/body/caption（**启发式**，判据见 `fonts.heuristic`，每级附代表单元格与出现次数）③ `palette` 配色（主题色 `themeColors`、工作簿 56 色调色板 `workbookPalette`、以及从实际用色统计出的 `observed` 排行）④ `tableStyles` 结构化表格样式名 ⑤ `conditionalFormatStyles` 条件格式风格 ⑥ `census` 样式普查分组明细 ⑦ `probes`/`unavailable`/`warnings` 探测记录。**读不到的字段一律 null 或进 unavailable，不会编默认值**（旧实现曾把读不到的字体/底色编成「微软雅黑 / #1E3A8A」）；字体层级是启发式推断，不是宿主读数。参数：`sampleAddress`（取样格，默认 'A3'）。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host；Microsoft 侧尚未实现）。",
        parameters: {
          type: "object",
          properties: {
            sampleAddress: { type: "string", description: "取样单元格，默认 'A3'" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_format_text_segment",
        description: "**单元格内局部格式（富文本）**：只给一个单元格里的一段文字加格式（如「整句里只把结论加粗变红」），不影响该格其余文字。定位方式二选一：① find 按文本找（可配 occurrence 指定第几处）② start（1 基起点）+ length。格式项至少要给一项：bold / italic / underline / fontColor / fontSize / fontName。**写完逐项读回核对**，任一项没落上直接报错，不做静默降级。若只想整格统一格式，用 format_cells 更省事。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            address: { type: "string", description: "目标单元格，如 'A1'（单格；多格请逐个调用）" },
            find: { type: "string", description: "要格式化的文本片段，按内容定位（优先于 start）" },
            occurrence: { type: "number", description: "find 命中第几处，默认 1（从 1 开始）" },
            start: { type: "number", description: "1 基起始字符位置（不传 find 时使用）" },
            length: { type: "number", description: "字符个数；不传则从 start 到单元格末尾" },
            bold: { type: "boolean", description: "该片段是否加粗" },
            italic: { type: "boolean", description: "该片段是否斜体" },
            underline: { type: "boolean", description: "该片段是否加下划线" },
            fontColor: { type: "string", description: "该片段字色，如 '#C00000'" },
            fontSize: { type: "number", description: "该片段字号（磅值）" },
            fontName: { type: "string", description: "该片段字体名" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["address"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_configure_print_layout",
        description: "**打印与分页设置**：打印区域、重复打印的标题行列、纸张、横竖向、缩放/适配页数、居中、网格线、页边距、页眉页脚（居中/左/右），以及手动分页符的增删。action='read' 只读回当前设置；默认 apply 先写后**逐项读回核对**（打印区域写错会直接报错）。做「可直接打印装订」的报表必用。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["apply", "read"], description: "apply（默认）先写后读回核对；read 只读回当前设置" },
            printArea: { type: "string", description: "打印区域，如 'A1:F40'；传空串或配合 clearPrintArea 清除" },
            printTitleRows: { type: "string", description: "每页重复的标题行，如 '$1:$2'（多页表格的必备项）" },
            printTitleColumns: { type: "string", description: "每页重复的标题列，如 '$A:$B'" },
            clearPrintArea: { type: "boolean", description: "清除打印区域与打印标题" },
            orientation: { type: "string", enum: ["portrait", "landscape"], description: "纸张方向：纵向/横向" },
            paperSize: { type: "string", enum: ["a4", "a3", "letter", "legal", "b5"], description: "纸张大小" },
            zoom: { type: "number", description: "缩放百分比（10-400）；与 fitToPages 互斥" },
            fitToPagesWide: { type: "number", description: "横向压缩到几页宽（1 = 一页宽）" },
            fitToPagesTall: { type: "number", description: "纵向压缩到几页高" },
            centerHorizontally: { type: "boolean", description: "水平居中" },
            centerVertically: { type: "boolean", description: "垂直居中" },
            printGridlines: { type: "boolean", description: "打印网格线" },
            leftMargin: { type: "number", description: "左边距（磅）" },
            rightMargin: { type: "number", description: "右边距（磅）" },
            topMargin: { type: "number", description: "上边距（磅）" },
            bottomMargin: { type: "number", description: "下边距（磅）" },
            headerMargin: { type: "number", description: "页眉边距（磅）" },
            footerMargin: { type: "number", description: "页脚边距（磅）" },
            centerHeader: { type: "string", description: "居中页眉文本" },
            leftHeader: { type: "string", description: "左页眉文本" },
            rightHeader: { type: "string", description: "右页眉文本" },
            centerFooter: { type: "string", description: "居中页脚文本" },
            leftFooter: { type: "string", description: "左页脚文本" },
            rightFooter: { type: "string", description: "右页脚文本" },
            addHorizontalPageBreak: { type: "number", description: "在第几行**之前**插入水平分页符" },
            addVerticalPageBreak: { type: "number", description: "在第几列**之前**插入垂直分页符（传列号数字）" },
            clearPageBreaks: { type: "boolean", description: "清除全部手动分页符" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_export_sheet_pdf",
        description: "**导出为 PDF**：scope='workbook' 导整个工作簿，'sheet' 只导指定工作表。返回 outputPath 与**落盘校验结果**（桥接侧用 existsSync 确认文件真的写出，宿主返回成功不代表写出）。交付链路的最后一步——生成 xlsx 之后直接出可发送的 PDF。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            scope: { type: "string", enum: ["workbook", "sheet"], description: "导出范围：整个工作簿（默认）或单个工作表" },
            quality: { type: "string", enum: ["standard", "minimum"], description: "导出质量，默认 standard" },
            sheetName: { type: "string", description: "scope='sheet' 时要导出的工作表名" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_add_shape",
        description: "在表格上画**矢量形状**：几何形状（矩形/椭圆/箭头/流程图形状等 30 种）、文本框、直线连接符、**艺术字（kind=wordart，WPS 独有）**。返回体带**读回的真实几何与样式**（位置/尺寸/旋转/填充/线条/文字），宿主没接受某个属性时写进 warnings——不做假成功。宿主差异（均已实测）：**WPS 表格全套可用**（矩形/直线/文本框/艺术字都成功）；Microsoft Excel 只有矩形与文本框可用，**直线报「当前对象不允许此操作」**，且没有 addSvg/getActiveShape。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            kind: { type: "string", enum: ["geometric", "textBox", "textbox", "line", "connector", "wordart"], description: "形状种类，默认 geometric；wordart 为艺术字（WPS 独有，需配 text/fontName/fontSize）" },
            shapeType: { type: "string", description: "kind=geometric 时的形状名（如 rectangle / ellipse / triangle / pentagon / arrow）" },
            text: { type: "string", description: "kind=textBox 时的文本内容；几何形状也可带文本" },
            left: { type: "number", description: "左边缘（磅）" },
            top: { type: "number", description: "上边缘（磅）" },
            width: { type: "number", description: "宽度（磅）" },
            height: { type: "number", description: "高度（磅）" },
            rotation: { type: "number", description: "旋转角度（度）" },
            x1: { type: "number", description: "kind=line 起点 x" },
            y1: { type: "number", description: "kind=line 起点 y" },
            x2: { type: "number", description: "kind=line 终点 x" },
            y2: { type: "number", description: "kind=line 终点 y" },
            fillColor: { type: "string", description: "填充色，如 '#2F6FEB'；也可用 fill 别名" },
            lineColor: { type: "string", description: "线条色" },
            lineWeight: { type: "number", description: "线条粗细（磅）" },
            name: { type: "string", description: "给形状起稳定名字，便于后续按名定位" },
            fontName: { type: "string", description: "文字字体名（对几何形状/文本框/艺术字都生效）" },
            fontColor: { type: "string", description: "文字颜色，如 '#1F3864'" },
            textAlign: { type: "string", enum: ["left", "center", "right"], description: "文字水平对齐；自选图形默认居中，文本框默认左对齐" },
            textVAlign: { type: "string", enum: ["top", "middle", "bottom"], description: "文字垂直对齐；自选图形默认垂直居中，文本框默认顶对齐" },
            marginLeft: { type: "number", description: "文字左边距（磅）" },
            marginRight: { type: "number", description: "文字右边距（磅）" },
            fontSize: { type: "number", description: "kind=wordart 时的字号（磅）" },
            bold: { type: "boolean", description: "kind=wordart 时是否加粗" },
            italic: { type: "boolean", description: "kind=wordart 时是否斜体" },
            wordArtPreset: { type: "number", description: "kind=wordart 的艺术字预设编号（msoTextEffect），默认 0" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_group_shapes",
        description: "把若干形状**组合**成一个组（按 shapeNames 指定）。**WPS 表格可用且能读回成员数**（走 Shapes.Range([...]).Group()，已实测）；⚠️ Microsoft Excel 侧尚未打通：宿主 API 本身可用（任务窗格自检在新建且激活的表上能组合成功），但经本工具对已存在的表调用会报「当前对象不允许此操作」，根因未定位——请把失败当作真实失败，不要重试绕过。",
        parameters: {
          type: "object",
          properties: {
            shapeNames: { type: "array", description: "要组合的形状名列表（用 list_shapes 或 add_shape 返回的 name）" },
            groupName: { type: "string", description: "组合后的名字，便于后续定位" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
          },
          required: ["shapeNames"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ungroup_shapes",
        description: "**解散分组**：把组内成员释放回工作表，返回释放后的形状名单与实际仍在表上的形状。仅对组合（Type=6）有效，传非组合会明确报错。WPS 表格可用；Microsoft Excel 侧依赖组合先成功（见 wps_group_shapes）。",
        parameters: {
          type: "object",
          properties: {
            shapeName: { type: "string", description: "要解散的组合名（组本身的名字）" },
            name: { type: "string", description: "要解散的组合名（与 shapeName 等价，二选一）" },
            shapeId: { type: "string", description: "或用形状 id 定位" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_set_shape_zorder",
        description: "调整形状的**层级**（bringToFront 置顶 / sendToBack 置底）。返回调整后所有形状的 zOrder 序列，可核对是否真的换了层。实测本机可用。",
        parameters: {
          type: "object",
          properties: {
            shapeName: { type: "string", description: "目标形状名" },
            name: { type: "string", description: "目标形状名（与 shapeName 等价，二选一）" },
            shapeId: { type: "string", description: "或用形状 id 定位" },
            zOrder: { type: "string", enum: ["bringToFront", "sendToBack", "bringForward", "sendBackward"], description: "层级调整方式" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
          },
          required: ["zOrder"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_export_shape_image",
        description: "把单个形状**导出为图片**（返回 base64）。做单图交付或把形状当图片复用时使用。实测本机可用。",
        parameters: {
          type: "object",
          properties: {
            shapeName: { type: "string", description: "要导出的形状名" },
            shapeId: { type: "string", description: "或用形状 id 定位" },
            format: { type: "string", enum: ["png", "jpeg", "gif", "bmp", "svg"], description: "图片格式，默认 png" },
            scale: { type: "number", description: "缩放倍数，默认 1" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc },
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_list_shapes",
        description: "**读回工作表上的全部形状**——绘图能力的验收入口。默认 detail=true 返回每个形状的名字/类型/位置/尺寸/旋转/可见性/填充色/线条色线宽/文字/所在单元格，组合还带成员数。写完形状后用它核对是否真的画上去了。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            detail: { type: "boolean", description: "是否返回完整属性，默认 true；false 只返回名字" },
            filterName: { type: "string", description: "只看某个名字的形状" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_update_shape",
        description: "**修改已有形状**：位置/尺寸/旋转/填充色/线条色线宽/文字/可见性，改名，或 action='delete' 删除。写完返回**读回的真实状态**。按 name 或 shapeIndex（从 1 开始）定位。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            name: { type: "string", description: "要修改的形状名（推荐）" },
            shapeIndex: { type: "number", description: "或按序号定位，从 1 开始" },
            action: { type: "string", enum: ["update", "delete"], description: "update（默认）或 delete" },
            left: { type: "number", description: "左边缘（磅）" },
            top: { type: "number", description: "上边缘（磅）" },
            width: { type: "number", description: "宽度（磅）" },
            height: { type: "number", description: "高度（磅）" },
            rotation: { type: "number", description: "旋转角度（度）" },
            fillColor: { type: "string", description: "填充色，如 '#2F6FEB'" },
            lineColor: { type: "string", description: "线条色" },
            lineWeight: { type: "number", description: "线条粗细（磅）" },
            text: { type: "string", description: "形状内文字" },
            visible: { type: "boolean", description: "是否可见" },
            newName: { type: "string", description: "改名" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_set_sheet_view",
        description: "设置工作表**视图**：是否显示网格线、行列标题、缩放比例。⚠️ **用矢量形状搭画布页（看板/信息图/封面页）前必须先隐藏网格线**，否则所有形状都浮在格线上，观感很乱。写后读回核对并返回 before/after，能确认是否真的生效。",
        parameters: {
          type: "object",
          properties: {
            showGridlines: { type: "boolean", description: "是否显示单元格网格线（画布页传 false）" },
            showHeadings: { type: "boolean", description: "是否显示行列标题（A/B/C 与 1/2/3）" },
            zoom: { type: "number", description: "缩放百分比，如 100" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_copy_range",
        description: "把一处区域的**值/公式/格式**复制到另一处（可跨工作表）。支持 all / values / formats / formulas 四种模式，写后读回目标左上角核对。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            copyType: { type: "string", enum: ["all", "values", "formats", "formulas"], description: "复制内容，默认 all" },
            sourceRange: { type: "string", description: "源区域，如 'A1:D10'" },
            destRange: { type: "string", description: "目标区域左上角或同尺寸区域，如 'F1'" },
            destSheetName: { type: "string", description: "目标工作表；省略则同表" },
            transpose: { type: "boolean", description: "是否转置粘贴" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["sourceRange", "destRange"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_manage_hyperlink",
        description: "单元格**超链接**增删查：加链接、列出全表链接（地址/显示文字/提示/锚点）、删除。写后从锚点读回核对。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["list", "add", "delete"], description: "list（默认）列出 / add 添加 / delete 删除（省略 address 则清空全表）" },
            address: { type: "string", description: "锚点单元格，如 'B2'" },
            url: { type: "string", description: "链接地址（add 必填）" },
            displayText: { type: "string", description: "显示文字" },
            tooltip: { type: "string", description: "悬停提示" },
            targetAddress: { type: "string", description: "同文档内的子地址（如某工作表/命名区域）" },
            sheetName: { type: "string", description: "工作表名称" },
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
        name: "wps_manage_named_range",
        description: "⚠️ 两条宿主限制（已实测）：① **`comment` 宿主不存储**——写入后读回恒为空，工具会给出 warnings；② **名字不能看起来像单元格地址**（如 `fz1`、`AB12`），宿主会拒绝，请用 `销量_2026` 这类名字。**命名区域**增删查：列出全部名称与引用位置、新增（如 图表源=Sheet1!$A$1:$B$10）、删除。写后读回引用位置核对。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["list", "add", "delete"], description: "list（默认）/ add / delete" },
            name: { type: "string", description: "名称（add/delete 必填）" },
            refersTo: { type: "string", description: "引用位置，如 'Sheet1!$A$1:$B$10'（add 必填）" },
            comment: { type: "string", description: "备注" },
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
        name: "wps_manage_document_properties",
        description: "读写工作簿**文档属性**：Title/Subject/Author/Keywords/Comments/Category/Company/Manager 八个内置字段 + 任意自定义属性。apply 后自动读回核对。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["read", "apply", "delete"], description: "read（默认）只读回 / apply 写入并核对 / delete 删除自定义属性（内置属性只能清空值，不能删）" },
            propertyNames: { type: "array", description: "delete 时要删除的自定义属性名数组" },
            properties: { type: "object", description: "要写入的键值对，如 {Title: 2026年报, Author: 财务部}；非内置键写入自定义属性" },
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
        name: "wps_manage_table",
        description: "**结构化表格**（Excel Table / ListObject）增删查：创建（带表头/汇总行/样式）、列出（名称/范围/行列数/列名）、删除（数据保留）。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["list", "apply", "delete"], description: "list（默认）/ apply 创建或更新 / delete 删除（仅删表格对象，数据保留）" },
            tableName: { type: "string", description: "表格名（create 时可指定，delete 必填）" },
            address: { type: "string", description: "创建时源区域，如 'A1:D20'" },
            styleName: { type: "string", description: "表格样式名" },
            newName: { type: "string", description: "改名" },
            hasHeaders: { type: "boolean", description: "是否含表头，默认 true" },
            totalsRow: { type: "boolean", description: "是否加汇总行，默认 false" },
            sheetName: { type: "string", description: "工作表名称" },
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
        name: "wps_manage_pictures",
        description: "工作表**图片**增删查：从本机路径插入图片（可指定位置尺寸）、列出全部图片（名称/位置/尺寸/所在单元格）、按名称或序号删除。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            action: { type: "string", enum: ["list", "insert", "delete"], description: "list（默认）/ insert / delete" },
            filePath: { type: "string", description: "本机图片绝对路径（insert 必填）" },
            pictureName: { type: "string", description: "图片名（插入时命名 / 删除时筛选）" },
            pictureIndex: { type: "number", description: "按形状序号删除" },
            left: { type: "number", description: "左边缘（磅）" },
            top: { type: "number", description: "上边缘（磅）" },
            width: { type: "number", description: "宽度（磅）" },
            height: { type: "number", description: "高度（磅）" },
            sheetName: { type: "string", description: "工作表名称" },
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
        name: "wps_export_chart_image",
        description: "把**图表导出为图片文件**（PNG/JPG/GIF/BMP）。⚠️ outputPath 必须落在 **WPS 可写目录**内（容器 tmp、用户文稿目录等）；写到 /tmp 之类不可达位置会「调用成功但不落盘」——桥接侧会用文件系统核对，返回值里的 fileWritten / fileSizeBytes 才是真实结果。工作表上有多个图表时须用 chartName 或 chartIndex 指定。选型：同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）。",
        parameters: {
          type: "object",
          properties: {
            chartName: { type: "string", description: "图表名（多个图表时必须指定其一）" },
            chartIndex: { type: "number", description: "或按图表序号，从 1 开始" },
            outputPath: { type: "string", description: "导出文件的绝对路径，须在 WPS 可写目录内" },
            format: { type: "string", enum: ["PNG", "JPG", "JPEG", "GIF", "BMP"], description: "图片格式，默认 PNG" },
            sheetName: { type: "string", description: "工作表名称" },
            workbookName: { type: "string", description: ctx.wbDesc }
          },
          required: ["outputPath"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_create_workbook",
        description: "**新建一个空白工作簿**（可同时保存到指定路径、给首张表命名）。写后读回工作簿名/表数/首表名核对。注意：新建的工作簿会成为活动工作簿，后续操作请显式带 workbookName 以免落到别的文件上。",
        parameters: {
          type: "object",
          properties: {
            savePath: { type: "string", description: "保存到的绝对路径（.xlsx）；省略则只新建不落盘" },
            sheetName: { type: "string", description: "给首张工作表起的名字" }
          },
          additionalProperties: false
        }
      }
    }
  ];
}
