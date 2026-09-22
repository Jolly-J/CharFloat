/**
 * PowerPoint（演示）工具定义（P2.3 剩余项：逐类迁移自 gateway.getOpenAiTools）。
 *
 * 已发布顺序中的第 51–59 条，分类标签为 `ppt`。
 * 注意：网关执行分支另有别名 `wps_ppt_add_chart`，它**没有**对外定义
 * （属"实现了但未注册"台账），不要顺手补进来。
 * 同理 `wps_ppt_save_presentation` 也**不能**在这里新增：网关 `HANDLERS` 注册表
 * （`src/bridge/gateway.ts`，不在本文件写区）没有该工具的处理器，注册后会命中
 * `tests/contract-consistency.test.ts` 的"注册工具缺少网关分支"断言。保存能力因此
 * 挂在 `wps_ppt_manage_slides` 的 `save` / `save_as` 动作上（宿主 `Presentation.Save`/`SaveAs`）。
 */
import type { GatewayToolDefinition } from './shared.js';

/** PowerPoint（演示）：已发布顺序第 51–59 条。定义顺序即契约顺序，不得重排。 */
export function pptToolDefinitions(): GatewayToolDefinition[] {
  return [
    {
      type: "function",
      function: {
        name: "wps_ppt_read_presentation",
        description: "读取 PowerPoint 演示文稿的大纲架构、幻灯片列表、各页文本要点及演讲者备注。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "演示文稿名称，不传则默认当前活动文稿" },
            includeNotes: { type: "boolean", description: "是否读取演讲者备注，默认 true" },
            maxSlides: { type: "integer", description: "最多读取的幻灯片页数，默认 50" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ppt_get_slide_shapes",
        description: "读取指定幻灯片中所有形状的几何坐标、尺寸、文本内容、表格元数据及图层层级。坐标为磅值(pt)。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "目标幻灯片页码(1-based)，不传默认当前活动页" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ppt_generate_deck",
        description: "按结构化 JSON 大纲在目标文稿中逐页生成幻灯片：title(封面) / cards_2、cards_3、cards_4(卡片页) / chart(原生图表页) / content(要点页) / end(结束页)。坐标为 720×405 设计基准，工具内部会按返回的 pageWidth/pageHeight 换算到实际页面尺寸：传 960×540 页面时实测每个几何量都会乘以 4/3（背景正好满页）；其他页面尺寸的缩放系数未实测。'content' 布局缺 bulletPoints 时本页只有标题，会记入 layoutWarnings；返回的 layoutWarnings 必须逐页预览核实。本工具是追加型：一次调用插入 N 页，不幂等，重复调用会继续插入，且 chart 布局失败时已插入的页面不会回滚。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            themeColor: { type: "string", description: "主题色十六进制值，如 '#0F4C81'(深蓝商务) 或 '#4B38B3'(科技紫)" },
            themePreset: { type: "string", enum: ["business_blue", "tech_purple", "clean_light"], description: "主题配色预设: 'business_blue'(商务蓝), 'tech_purple'(科技紫), 'clean_light'(浅色简洁)" },
            slides: {
              type: "array",
              description: "整套幻灯片规格清单",
              items: {
                type: "object",
                properties: {
                  layout: { type: "string", enum: ["title", "cards_2", "cards_3", "cards_4", "chart", "content", "end"], description: "页面版式: 'title'(封面), 'cards_2'/'cards_3'/'cards_4'(2/3/4 栏卡片), 'chart'(原生图表页), 'content'(要点列表页), 'end'(结束页)" },
                  title: { type: "string", description: "页面大标题" },
                  subtitle: { type: "string", description: "副标题（封面页使用）" },
                  bulletPoints: { type: "array", items: { type: "string" }, description: "普通内容页的要点列表；layout='content' 时必传，否则本页只有标题" },
                  cards: {
                    type: "array",
                    description: "商业卡片列表（cards_2 / cards_3 / cards_4 版式使用）",
                    items: {
                      type: "object",
                      properties: {
                        tag: { type: "string", description: "分类标签，如 'CORE', 'PHASE 1'" },
                        title: { type: "string", description: "卡片主标题" },
                        description: { type: "string", description: "卡片描述详情" },
                        accentColor: { type: "string", description: "卡片强调色" }
                      },
                      required: ["title", "description"]
                    }
                  },
                  chart: {
                    type: "object",
                    description: "原生图表规格（chart 版式使用）",
                    properties: {
                      chartType: { type: "string", enum: ["column", "line", "bar", "pie"], description: "图表类型: 'column'(柱状图), 'line'(折线图), 'bar'(条形图), 'pie'(饼图)" },
                      title: { type: "string", description: "图表标题" },
                      categories: { type: "array", items: { type: "string" }, description: "横轴分类标签" },
                      series: {
                        type: "array",
                        description: "数据系列数组，每个系列一组数值（长度应与 categories 对应）",
                        items: {
                          type: "object",
                          properties: {
                            name: { type: "string", description: "系列名称，显示在图例/图例项中" },
                            values: { type: "array", items: { type: "number" }, description: "该系列的数值列表，长度应与 categories 一致" }
                          },
                          required: ["name", "values"]
                        }
                      }
                    }
                  },
                  notes: { type: "string", description: "本页演讲者备注 (Speaker Notes)" }
                },
                required: ["layout", "title"]
              }
            }
          },
          required: ["slides"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ppt_manage_slides",
        description: "PowerPoint 幻灯片管理：新建演示文稿(new_presentation)、新增、删除、移动顺序、复制克隆、背景颜色设置与保存(save/save_as)。add/duplicate/new_presentation 属追加型操作、不幂等，重复调用会新增更多页；delete 是破坏性操作，先确认 slideIndex；越界页码会返回带有效范围的中文错误，不再回传宿主内部 JS 报错。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称；action='new_presentation' 时忽略（新建后请用返回的 presentationName 继续操作）" },
            action: { type: "string", enum: ["new_presentation", "add", "delete", "move", "duplicate", "set_background", "save", "save_as"], description: "操作: 'new_presentation'(新建演示文稿，可配 filePath 直接另存), 'add'(新增页，可配 layoutIndex), 'delete'(删除页，需 slideIndex), 'move'(移动页，需 slideIndex + targetIndex), 'duplicate'(克隆页，需 slideIndex), 'set_background'(设置背景色，需 slideIndex + backgroundColor), 'save'/'save_as'(保存；不传 filePath 走原地保存，未命名文稿会明确报告“没有文件路径”，不会静默成功)" },
            slideIndex: { type: "integer", description: "目标幻灯片页码(1-based)；越界时报错并给出当前总页数" },
            targetIndex: { type: "integer", description: "移动操作的目标页码(1-based)" },
            layoutIndex: { type: "integer", description: "新增页使用的 ppLayout 版式枚举，不是母版 CustomLayouts 的 1..N 序号。常用值：1=标题幻灯片(标题+副标题占位符)、2=标题和文本、7=标题和图示或组织结构图、12=空白(默认)。取值越界时宿主不报错但版式不可预期；要精确套用本模板的自定义版式，请改用 wps_execute_script 操作 slide.CustomLayout。" },
            backgroundColor: { type: "string", description: "set_background 时的十六进制颜色，如 '#0F172A'" },
            filePath: { type: "string", description: "保存/新建的目标完整路径（如 '/Users/.../方案.pptx' 或 '.pdf'）。注意：当前网关只向宿主转发 presentationName/action/slideIndex/targetIndex/layoutIndex/backgroundColor，filePath 尚未透传；在网关补齐前，带 filePath 的 save_as / new_presentation 只会退回原地保存或新建未保存文稿，需要指定落盘路径时请用 wps_execute_script 调 pres.SaveAs(path)。" },
            format: { type: "string", enum: ["pptx", "pdf"], description: "保存格式，默认 'pptx'(另存为 pptx)；'pdf' 走导出。同样受 filePath 未透传的限制。" }
          },
          required: ["action"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ppt_manage_table",
        description: "PowerPoint 表格操作：新建表格（可同时注入二维数据）、读取表格、改写单元格文字、套用表头/斑马纹/边框样式。宿主没有独立的“批量改写已有表格数据”操作：批量改数据请重新 create_table(data)，或逐格 set_cell_text。坐标与尺寸单位是 pt，不是像素或百分比。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "目标幻灯片页码(1-based)" },
            action: { type: "string", enum: ["create_table", "read_table", "set_cell_text", "style_table"], description: "表格操作: 'create_table'(新建，需 rows/columns，可用 data 同时填充), 'read_table'(读取表格数据), 'set_cell_text'(改写单元格，需 row/column/text), 'style_table'(应用表头/斑马纹/字号/边框样式)" },
            shapeId: { type: ["string", "integer"], description: "表格所在的形状 ID（数字）或名称（字符串）。数字先按 Shape.Id 匹配，未命中再按页内 1-based 索引兜底；不传则按 tableIndex 取第 N 个表格。" },
            tableIndex: { type: "integer", description: "页内第几个表格(默认 1)，仅在 shapeId 未命中时生效" },
            rows: { type: "integer", description: "新建表格行数" },
            columns: { type: "integer", description: "新建表格列数" },
            left: { type: "number", description: "距页面左边缘的磅值 pt；先读取实际 pageWidth" },
            top: { type: "number", description: "距页面上边缘的磅值 pt；先读取实际 pageHeight" },
            width: { type: "number", description: "宽度，单位 pt，不是像素或百分比" },
            height: { type: "number", description: "高度，单位 pt，不是像素或百分比" },
            data: { type: "array", items: { type: "array", items: { type: "string" } }, description: "二维表格数据数组，仅 create_table 时用于初始填充。每个单元格必须是字符串：数字会被 schema 拒绝并报 arguments.data[行][列]: 类型不正确，请自行转成字符串（如 '1000'）。" },
            row: { type: "integer", description: "目标行号(1-based)" },
            column: { type: "integer", description: "目标列号(1-based)" },
            text: { type: "string", description: "单元格文字内容" },
            fontSize: { type: "number", description: "字号，单位 pt；依据页面尺寸和文本框容量设置，编辑后读回并预览" },
            fontColor: { type: "string", description: "字体颜色十六进制" },
            fontBold: { type: "boolean", description: "是否加粗" },
            fillColor: { type: "string", description: "单元格填充底色十六进制" },
            headerFillColor: { type: "string", description: "表头背景底色十六进制(默认 '#0F4C81')" },
            headerFontSize: { type: "number", description: "表头字体字号大小(默认 16)" },
            bodyFontSize: { type: "number", description: "数据行正文字号大小(默认 14)" },
            borderColor: { type: "string", description: "表格边框线颜色十六进制" },
            columnWidths: { type: "array", items: { type: "number" }, description: "各列宽度数组(磅值)" },
            rowHeights: { type: "array", items: { type: "number" }, description: "各行高度数组(磅值)" },
            zebra: { type: "boolean", description: "是否启用交替斑马纹底色" }
          },
          required: ["action"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ppt_add_business_cards",
        description: "在指定幻灯片自动排版 2/3/4 栏信息卡片（圆角底卡、强调色顶线、分类 Tag、标题与正文）。卡片坐标按 720×405 设计基准换算到实际页面尺寸并限制在页面内。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "目标幻灯片页码(1-based)" },
            columnCount: { type: "integer", enum: [2, 3, 4], description: "卡片分栏数量(2, 3 或 4)" },
            cards: {
              type: "array",
              description: "卡片数据数组",
              items: {
                type: "object",
                properties: {
                  tag: { type: "string", description: "顶部小标签" },
                  title: { type: "string", description: "卡片标题" },
                  description: { type: "string", description: "卡片正文" },
                  accentColor: { type: "string", description: "强调色" }
                },
                required: ["title", "description"]
              }
            },
            topY: { type: "number", description: "卡片顶端 Y 坐标(默认 100)" },
            cardHeight: { type: "number", description: "卡片高度(默认 260)" }
          },
          required: ["cards"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ppt_insert_native_chart",
        description: "在 PPT 中插入原生矢量图表并绑定结构化数据。插入后必须用 wps_ppt_get_slide_shapes 或预览核实图表真的生成、数据真的写入（宿主存在插入成功但不建形状的情况）。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "目标幻灯片页码(1-based)" },
            chartType: { type: "string", enum: ["column", "line", "bar", "pie", "column_stacked", "bar_stacked", "bar_of_pie"], description: "图表类型: 'column'(柱状图), 'line'(折线图), 'bar'(条形图), 'pie'(饼图), 'column_stacked'(堆积柱形图), 'bar_stacked'(堆积条形图), 'bar_of_pie'(复合条饼图)" },
            title: { type: "string", description: "图表标题" },
            hasLegend: { type: "boolean", description: "是否显示图例" },
            showDataLabels: { type: "boolean", description: "是否显示数据标签" },
            categories: { type: "array", items: { type: "string" }, description: "分类横轴项目，例如 ['Q1', 'Q2', 'Q3', 'Q4']" },
            series: {
              type: "array",
              description: "数据系列",
              items: {
                type: "object",
                properties: {
                  name: { type: "string", description: "系列名称，例如 '营收 (万元)'" },
                  values: { type: "array", items: { type: "number" }, description: "数值列表" },
                  chartType: { type: "string", description: "单个系列的图表形态，例如 'line_markers', 'line', 'column'" },
                  axisGroup: { type: "integer", description: "坐标轴分组: 1为主坐标轴，2为次坐标轴" },
                  color: { type: "string", description: "系列十六进制颜色值，例如 '#FF7F00'" },
                  hasDataLabels: { type: "boolean", description: "该系列是否单独启用数据标签" }
                },
                required: ["name", "values"]
              }
            },
            left: { type: "number", description: "距页面左边缘的磅值 pt；先读取实际 pageWidth" },
            top: { type: "number", description: "距页面上边缘的磅值 pt；先读取实际 pageHeight" },
            width: { type: "number", description: "宽度，单位 pt，不是像素或百分比" },
            height: { type: "number", description: "高度，单位 pt，不是像素或百分比" }
          },
          required: ["categories", "series"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ppt_manage_shapes_and_media",
        description: "坐标和字号统一使用 pt。先读页面尺寸和现有形状；新增时显式提供 left/top/width/height，文本还需 fontSize，避免默认位置重叠。支持添加文本框、添加形状、修改已有形状的位置/尺寸/文本/填充、交换两个形状位置(swap_shapes)、设置图层层级、对齐分布(align_shapes)与删除形状(delete_shape)。两个对齐/交换动作的真实语义见对应参数说明，勿按通用 PPT 语义理解。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "目标幻灯片页码" },
            action: { type: "string", enum: ["add_textbox", "add_shape", "update_shape", "swap_shapes", "set_z_order", "align_shapes", "delete_shape"], description: "操作: 'add_textbox', 'add_shape', 'update_shape', 'swap_shapes', 'set_z_order', 'align_shapes', 'delete_shape'(删除形状，破坏性操作，需 shapeId)" },
            shapeType: { type: "string", enum: ["rectangle", "rounded_rectangle", "oval", "arrow"], description: "形状类型: 'rectangle'(矩形), 'rounded_rectangle'(圆角矩形), 'oval'(椭圆), 'arrow'(箭头)，action='add_shape' 时使用" },
            shapeId: { type: ["string", "integer"], description: "形状 ID（数字）或名称（字符串）。数字先按 Shape.Id 匹配，未命中再按页内 1-based 索引兜底，因此传 1、2、3 这类小数字可能落到索引而不是 Id。" },
            shapeId1: { type: ["string", "integer"], description: "互换位置或对齐时的第一个形状 ID（解析规则同 shapeId）。action='swap_shapes' 时与 shapeId2 只交换垂直位置(Top)、Left 不变：两个横向并排的形状交换后位置实际不变，需要水平换位请用 update_shape 显式设 left。" },
            shapeId2: { type: ["string", "integer"], description: "互换位置或对齐时的第二个形状 ID（解析规则同 shapeId）" },
            shapeIds: { type: "array", items: { type: ["string", "integer"] }, description: "批量对齐的形状 ID 列表。解析顺序同 shapeId；对齐基准是列表中第一个可解析到的形状，其余形状向它对齐。" },
            alignType: { type: "string", enum: ["left", "center", "right", "top", "middle", "bottom"], description: "对齐方式（以 shapeIds 首个形状为基准，不是页面或选区对齐）: 'left'/'center'/'right'(左/水平居中/右边缘), 'top'/'middle'/'bottom'(上/垂直居中/下边缘)" },
            zOrderAction: { type: "string", enum: ["bring_to_front", "send_to_back", "bring_forward", "send_backward"], description: "图层层级调整方式: 'bring_to_front'(置于顶层), 'send_to_back'(置于底层), 'bring_forward'(上移一层), 'send_backward'(下移一层)" },
            text: { type: "string", description: "文本内容" },
            fontSize: { type: "number", description: "字号，单位 pt；依据页面尺寸和文本框容量设置，编辑后读回并预览" },
            fontColor: { type: "string", description: "字体颜色十六进制" },
            fontBold: { type: "boolean", description: "是否加粗" },
            alignment: { type: "string", enum: ["left", "center", "right", "justify"], description: "段落水平对齐: 'left'(左), 'center'(居中), 'right'(右), 'justify'(两端对齐)" },
            left: { type: "number", description: "距页面左边缘的磅值 pt；先读取实际 pageWidth" },
            top: { type: "number", description: "距页面上边缘的磅值 pt；先读取实际 pageHeight" },
            width: { type: "number", description: "宽度，单位 pt，不是像素或百分比" },
            height: { type: "number", description: "高度，单位 pt，不是像素或百分比" },
            rotation: { type: "number", description: "旋转角度" },
            fillColor: { type: "string", description: "填充颜色十六进制" },
            lineColor: { type: "string", description: "边框线条颜色十六进制" }
          },
          required: ["action"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_ppt_capture_slide_preview",
        description: "将指定幻灯片导出为 PNG 图片并返回 base64，用于视觉核查排版是否重叠、文字是否溢出。失败时把宿主对每次导出尝试的原始报错放在 hostError / attempts 字段里回传（不再只给一句“PPT 未生成预览”）。已知可用绕行：wps_execute_script 里 call slide.Export(绝对路径, 'PNG', 1280, 720)。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "要截图的幻灯片页码(1-based)，不传默认第1页；越界时报错并给出当前总页数" },
            outputPath: { type: "string", description: "可选的额外导出落盘路径（如 '/Users/.../slide3.png'）。宿主会把图片同时写到该路径；但桥接返回的 base64 仍读取它自己生成的临时预览文件，因此指定本参数只用于你自己取图，不代表工具会自动返回该路径的图片。" }
          },
          required: [],
          additionalProperties: false
        }
      }
    }  ,
  {
    type: "function",
    function: {
      name: "wps_ppt_configure_layout",
      description: "演示文稿的**页面尺寸与母版版式**：改 16:9 / 4:3 / A4、切横纵方向、套用模板文件、列出并套用母版版式（CustomLayouts）。写后逐项读回核对，未生效写 warnings。action='read' 只读现状，action='list_layouts' 列出全部版式名。",
      parameters: {
        type: "object",
        properties: {
          action: { type: "string", enum: ["apply", "read", "list_layouts"], description: "apply（默认）写入并核对 / read 只读现状 / list_layouts 列出母版全部版式名" },
          preset: { type: "string", description: "页面尺寸预设：16:9（960x540）/ 4:3（720x540）/ a4 / a4_portrait" },
          slideWidth: { type: "number", description: "页宽（磅）；同时传 preset 与宽度时以显式宽高为准" },
          slideHeight: { type: "number", description: "页高（磅）" },
          orientation: { type: "string", enum: ["landscape", "portrait"], description: "页面方向" },
          templatePath: { type: "string", description: "要套用的模板文件绝对路径（.potx/.pptx）" },
          layoutName: { type: "string", description: "要套用到各页的母版版式名（见 list_layouts）" },
          layoutIndex: { type: "number", description: "或按版式序号（从 1 开始）" },
          presentationName: { type: "string", description: "演示文稿名称" }
        },
        additionalProperties: false
      }
    }
  }
]
}
