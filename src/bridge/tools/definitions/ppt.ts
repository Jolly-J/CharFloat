/**
 * PowerPoint（演示）工具定义（P2.3 剩余项：逐类迁移自 gateway.getOpenAiTools）。
 *
 * 已发布顺序中的第 51–59 条，分类标签为 `ppt`。
 * 注意：网关执行分支另有别名 `wps_ppt_add_chart`，它**没有**对外定义
 * （属"实现了但未注册"台账），不要顺手补进来。
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
        description: "【精准探测】深度获取指定幻灯片中所有形状的详细几何坐标、尺寸、文本内容、表格元数据及图层层级信息。",
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
        description: "根据结构化 JSON 大纲，按目标文稿真实页面尺寸生成演示胶片；返回 layoutWarnings，需逐页预览验证（封面页、2/3/4栏商业卡片页、原生图表页、结束页）。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            themeColor: { type: "string", description: "主题色十六进制值，如 '#0F4C81'(深蓝商务) 或 '#4B38B3'(科技紫)" },
            themePreset: { type: "string", enum: ["business_blue", "tech_purple", "clean_light"], description: "主题配色预设" },
            slides: {
              type: "array",
              description: "整套幻灯片规格清单",
              items: {
                type: "object",
                properties: {
                  layout: { type: "string", enum: ["title", "cards_2", "cards_3", "cards_4", "chart", "content", "end"], description: "页面版式布局" },
                  title: { type: "string", description: "页面大标题" },
                  subtitle: { type: "string", description: "副标题（封面页使用）" },
                  bulletPoints: { type: "array", items: { type: "string" }, description: "普通内容页的要点列表" },
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
                      chartType: { type: "string", enum: ["column", "line", "bar", "pie"], description: "图表类型" },
                      title: { type: "string", description: "图表标题" },
                      categories: { type: "array", items: { type: "string" }, description: "横轴分类标签" },
                      series: {
                        type: "array",
                        items: {
                          type: "object",
                          properties: {
                            name: { type: "string" },
                            values: { type: "array", items: { type: "number" } }
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
        description: "PowerPoint 幻灯片综合管理。支持单页幻灯片的增、删、移动顺序、复制克隆及背景颜色设置。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            action: { type: "string", enum: ["add", "delete", "move", "duplicate", "set_background"], description: "操作类型" },
            slideIndex: { type: "integer", description: "目标幻灯片页码(1-based)" },
            targetIndex: { type: "integer", description: "移动操作时的目标页码" },
            layoutIndex: { type: "integer", description: "添加幻灯片时的版式序号(默认 12 空白版式)" },
            backgroundColor: { type: "string", description: "设置背景时的十六进制颜色，如 '#0F172A'" }
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
        description: "【专精表格控制器】在 PowerPoint 中创建商业表格、批量注入二维数据、读取表格、读写单元格及应用专业主题色样式。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "目标幻灯片页码(1-based)" },
            action: { type: "string", enum: ["create_table", "read_table", "set_table_data", "set_cell_text", "style_table"], description: "表格操作类型" },
            shapeId: { description: "表格所在的形状 ID", oneOf: [{ type: "string" }, { type: "integer" }] },
            tableIndex: { type: "integer", description: "页内第几个表格(默认 1)" },
            rows: { type: "integer", description: "新建表格行数" },
            columns: { type: "integer", description: "新建表格列数" },
            left: { type: "number", description: "距页面左边缘的磅值 pt；先读取实际 pageWidth" },
            top: { type: "number", description: "距页面上边缘的磅值 pt；先读取实际 pageHeight" },
            width: { type: "number", description: "宽度，单位 pt，不是像素或百分比" },
            height: { type: "number", description: "高度，单位 pt，不是像素或百分比" },
            data: { type: "array", items: { type: "array", items: { type: "string" } }, description: "二维表格数据数组" },
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
        description: "在指定幻灯片自动计算排版 2 栏、3 栏、4 栏现代化商业信息卡片（圆角底卡、强调色顶线、分类 Tag、标题与正文）。",
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
        description: "在 PPT 中插入原生可交互矢量商业图表（柱状图、折线图、条形图、饼图）并自动绑定填充结构化业务数据。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "目标幻灯片页码(1-based)" },
            chartType: { type: "string", enum: ["column", "line", "bar", "pie", "column_stacked", "bar_stacked", "bar_of_pie"], description: "图表类型: 'column'(柱状图), 'line'(折线图), 'bar'(条形图), 'pie'(饼图), 'column_stacked'(堆积柱形), 'bar_of_pie'(复合条饼图)" },
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
        description: "坐标和字号统一使用 pt。先读页面尺寸和现有形状；新增时显式提供 left/top/width/height，文本还需 fontSize，避免默认位置重叠。添加文本框、形状、修改已有形状位置与尺寸、智能互换两个形状位置(swap_shapes)、设置图层层级及对齐分布。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "目标幻灯片页码" },
            action: { type: "string", enum: ["add_textbox", "add_shape", "update_shape", "swap_shapes", "set_z_order", "align_shapes", "delete_shape"], description: "操作类型" },
            shapeType: { type: "string", enum: ["rectangle", "rounded_rectangle", "oval", "arrow"], description: "形状类型" },
            shapeId: { description: "形状 ID / 标识", oneOf: [{ type: "string" }, { type: "integer" }] },
            shapeId1: { description: "互换位置或对齐时的第一个形状 ID", oneOf: [{ type: "string" }, { type: "integer" }] },
            shapeId2: { description: "互换位置或对齐时的第二个形状 ID", oneOf: [{ type: "string" }, { type: "integer" }] },
            shapeIds: { type: "array", items: { oneOf: [{ type: "string" }, { type: "integer" }] }, description: "批量对齐时的形状 ID 列表" },
            alignType: { type: "string", enum: ["left", "center", "right", "top", "middle", "bottom"], description: "对齐方式" },
            zOrderAction: { type: "string", enum: ["bring_to_front", "send_to_back", "bring_forward", "send_backward"], description: "图层层级调整方式" },
            text: { type: "string", description: "文本内容" },
            fontSize: { type: "number", description: "字号，单位 pt；依据页面尺寸和文本框容量设置，编辑后读回并预览" },
            fontColor: { type: "string", description: "字体颜色十六进制" },
            fontBold: { type: "boolean", description: "是否加粗" },
            alignment: { type: "string", enum: ["left", "center", "right", "justify"], description: "段落水平对齐" },
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
        description: "【多模态自检】将指定幻灯片导出为高保真图片，供多模态大模型视觉核查排版是否重叠、排版是否美观对齐。",
        parameters: {
          type: "object",
          properties: {
            presentationName: { type: "string", description: "目标文稿名称" },
            slideIndex: { type: "integer", description: "要截图的幻灯片页码(1-based)，不传默认第1页" }
          },
          required: [],
          additionalProperties: false
        }
      }
    }
  ];
}
