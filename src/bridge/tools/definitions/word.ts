/**
 * Word（文字）工具定义（P2.3 剩余项：逐类迁移自 gateway.getOpenAiTools）。
 *
 * 已发布顺序中的第 39–50 条，分类标签为 `word`。
 * 注意：网关执行分支另有 `wps_word_capture_preview`，它**没有**对外定义
 * （属 tests/contract-consistency.test.ts 的"实现了但未注册"台账），不要顺手补进来。
 */
import type { GatewayToolDefinition } from './shared.js';

/** Word（文字）：已发布顺序第 39–50 条。定义顺序即契约顺序，不得重排。 */
export function wordToolDefinitions(): GatewayToolDefinition[] {
  return [
    {
      type: "function",
      function: {
        name: "wps_word_create_document",
        description: "新建空白 Word 文档或基于指定模板创建文档。不传 templatePath 即新建空白文档；新建后需显式调用 wps_word_save_document 才会落盘。",
        parameters: {
          type: "object",
          properties: {
            templatePath: { type: "string", description: "可选模板文件完整路径 (.dotx/.dotm/.dot)" },
            isVisible: { type: "boolean", description: "是否显示文档窗口，默认 true" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_save_document",
        description: "保存当前 Word 文档、另存为指定路径或导出为 PDF 格式。传 filePath 时会校验文件是否真的落盘，未写出即报错（宿主返回成功不代表已写出）。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称，不传则默认当前活动文档" },
            filePath: { type: "string", description: "保存目标完整路径（例如 '/Users/.../文档.docx' 或 '/Users/.../文档.pdf'），不传则执行原地保存" },
            format: { type: "string", enum: ["docx", "pdf"], description: "保存格式: 'docx'(常规文档，默认), 'pdf'(导出为 PDF 格式)" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_close_document",
        description: "关闭指定的 Word 文档。saveChanges 默认 false：未保存的修改会直接丢失；不传 documentName 时关闭当前活动文档，多文档场景请显式指定目标。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称，不传则默认当前活动文档" },
            saveChanges: { type: "boolean", description: "关闭前是否保存修改，默认 false" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_manage_content",
        description: "物理删除 Word 中的段落、表格或一键清空重写整个文档。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称" },
            action: {
              type: "string",
              enum: ["delete_paragraph", "delete_table", "clear_all"],
              description: "操作类型: 'delete_paragraph'(物理删除指定段落或段落范围), 'delete_table'(删除指定表格), 'clear_all'(清空全文内容)"
            },
            paragraphIndex: { type: "integer", description: "要删除的段落索引序号(1-based)" },
            paragraphRange: {
              type: "array",
              items: { type: "integer" },
              description: "要删除的段落范围 [起始序号, 结束序号]，例如 [3, 5]"
            },
            tableIndex: { type: "integer", description: "要删除的表格索引序号(1-based)，默认 1" }
          },
          required: ["action"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_read_document",
        description: "读取 Word 文档连续段落内容 [P1, P2...]、标题大纲骨架、排版元数据（字体、字号、加粗、对齐等）、表格结构，以及只读读回：节与页面版式/页眉页脚/页码域/水印形状(scope='layout')、样式清单(scope='styles')、文档属性(scope='properties')、书签(scope='bookmarks')、内容控件(scope='content_controls')、域(scope='fields')。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "文档名称，例如 '关于召开年度总结大会的通知.docx'，不传则默认当前活动文档" },
            scope: { type: "string", enum: ["outline", "full", "selection", "paragraphs", "tables", "layout", "styles", "properties", "bookmarks", "content_controls", "fields"], description: "读取范围: 'outline'(仅标题大纲), 'full'(全文预览与大纲，默认), 'selection'(当前选区), 'paragraphs'(仅连续段落文本数组), 'tables'(仅表格结构与预览)；以下是只读读回类（只返回对应字段，不返回段落预览）: 'layout'(每节的页面尺寸/方向/页边距/页眉页脚文本与域/页码格式/水印形状), 'styles'(样式总数、启用中样式清单), 'properties'(标题/作者/主题/关键字/自定义属性), 'bookmarks'(书签名称与范围), 'content_controls'(内容控件 Title/Tag/类型/文本), 'fields'(文档域清单与域码)" },
            maxParagraphs: { type: "integer", description: "最多返回的段落数量，默认 200" },
            includeFormatting: { type: "boolean", description: "是否提取段落级排版元数据（是否加粗、字号、字体名等），默认 true" },
            includeTables: { type: "boolean", description: "是否返回文档内全部表格的尺寸与前三行预览，默认 true" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_write_content",
        description: "向 Word 文档结构化写入内容（标题、正文、列表、引用或代码块）。支持指定排版样式，并可在开头、结尾、指定段落后、光标处或书签所在段落之后插入。location='bookmark' 时书签必须存在，否则**直接报错拒绝**（不会静默落到文末）；写入后返回 insertedParagraphs 与实际落点，并核对读回长度，宿主吞字符时逐条报错而不是返回 success。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称，不传则默认当前文档" },
            content: {
              type: ["string", "array"],
              items: { type: "string" },
              description: "要写入的内容：单个字符串，或多行字符串数组（每个元素一段）"
            },
            type: {
              type: "string",
              enum: ["paragraph", "heading1", "heading2", "heading3", "bullet_list", "quote", "code_block"],
              description: "内容段落类型: 'paragraph'(普通正文), 'heading1'/'heading2'/'heading3'(一/二/三级标题), 'bullet_list'(项目符号列表), 'quote'(引用块), 'code_block'(代码块，等宽字体)"
            },
            location: {
              type: "string",
              enum: ["end", "start", "selection", "bookmark", "after_paragraph"],
              description: "写入位置: 'end'(文档末尾，默认), 'start'(文档最前), 'selection'(当前光标处), 'bookmark'(指定书签，实际写在书签**所在段落之后**的新段落里), 'after_paragraph'(指定段落后)"
            },
            paragraphIndex: { type: "integer", description: "当 location 为 'after_paragraph' 时的基准段落索引(1-based)，越界会报错" },
            targetBookmark: { type: "string", description: "当 location 为 bookmark 时的书签名称；不存在则报错（错误信息里会列出文档现有书签）" },
            formatting: {
              type: "object",
              properties: {
                bold: { type: "boolean", description: "是否加粗" },
                italic: { type: "boolean", description: "是否斜体" },
                fontSizePt: { type: "number", description: "字号磅值（如 16 为三号，14 为四号，12 为小四）" },
                fontName: { type: "string", description: "字体名称（如 '仿宋_GB2312'、'宋体'、'微软雅黑'）" },
                alignment: { type: "integer", description: "对齐方式: 0(左对齐), 1(居中), 2(右对齐), 3(两端对齐)" },
                firstLineIndentChars: { type: "number", description: "首行缩进字符数（如 2）" },
                lineSpacingPt: { type: "number", description: "固定行间距磅值（如 28）" },
                spaceBeforePt: { type: "number", description: "段前间距磅值" },
                spaceAfterPt: { type: "number", description: "段后间距磅值" }
              },
              description: "写入文本的精细化排版参数"
            }
          },
          required: ["content"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_format_document",
        description: "Word 文档排版：按预设或自定义参数设置指定范围的字体、字号、加粗、缩进与行间距。target 默认 'all'（整篇），只作用于段落级格式；既不传 preset 也不传任何自定义参数时不会产生实际变化。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称" },
            target: { type: "string", enum: ["all", "paragraph", "range", "selection"], description: "格式化目标范围: 'all'(全文，默认), 'paragraph'(单个段落), 'range'(段落范围), 'selection'(当前选区)" },
            paragraphIndex: { type: "integer", description: "当 target 为 'paragraph' 时的段落索引号" },
            paragraphRange: { type: "array", items: { type: "integer" }, description: "当 target 为 'range' 时的段落范围 [start, end]" },
            preset: {
              type: "string",
              enum: ["gov_standard", "business_modern", "academic", "custom"],
              description: "排版预设: 'gov_standard'(国家标准公文规范：仿宋三号+28磅行距+首行缩进2字符+公文字体分级), 'business_modern'(现代商务排版：微软雅黑/Segoe UI+段后间距), 'academic'(当前未实现独立规则，等同 'custom'：只应用下面显式传入的自定义参数), 'custom'(只用自定义参数)"
            },
            fontName: { type: "string", description: "自定义字体名称，如 '仿宋_GB2312' 或 '微软雅黑'" },
            fontSizePt: { type: "number", description: "自定义字号(磅值)，如 16(三号) 或 12" },
            bold: { type: "boolean", description: "是否加粗" },
            italic: { type: "boolean", description: "是否斜体" },
            lineSpacingPt: { type: "number", description: "自定义固定行间距(磅值)，如 28" },
            firstLineIndentChars: { type: "number", description: "首行缩进字符数，如 2" },
            spaceBeforePt: { type: "number", description: "段前磅值" },
            spaceAfterPt: { type: "number", description: "段后磅值" },
            alignment: { type: "integer", description: "对齐方式: 0(左对齐), 1(居中), 2(右对齐), 3(两端对齐)" },
            margins: {
              type: "object",
              properties: {
                topMm: { type: "number", description: "上边距(毫米)" },
                bottomMm: { type: "number", description: "下边距(毫米)" },
                leftMm: { type: "number", description: "左边距(毫米)" },
                rightMm: { type: "number", description: "右边距(毫米)" }
              },
              description: "页面边距(毫米)，整篇生效；单位是毫米不是磅，不传则沿用文档原设置"
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
        name: "wps_word_insert_table_of_contents",
        description: "在 Word 文档中生成目录：扫描全文 Heading 1-3 级标题，在指定位置插入带前导符与页码的目录域。标题必须已使用 Heading 样式才会被收录。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称" },
            upperHeadingLevel: { type: "integer", description: "目录包含的最高标题级别，默认 1" },
            lowerHeadingLevel: { type: "integer", description: "目录包含的最低标题级别，默认 3" },
            insertLocation: { type: "string", enum: ["start", "selection"], description: "目录插入位置: 'start'(文档最前，默认), 'selection'(当前光标处)" },
            includePageNumbers: { type: "boolean", description: "是否显示页码，默认 true" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_manage_table",
        description: "Word 专业表格管理。支持结构透视(inspect)、三线表插入、安全动态扩容二维写入(write_matrix)、单元格独立排版与行列增删。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称" },
            action: {
              type: "string",
              enum: ["inspect", "insert", "update_data", "write_matrix", "format_cell", "add_row", "delete_row", "merge_cells"],
              description: "操作类型: 'inspect'(透视行列数与全量数据), 'insert'(插入新三线表), 'update_data'/'write_matrix'(动态扩容安全写入二维矩阵), 'format_cell'(单元格底色/字体排版), 'add_row'(追加行), 'delete_row'(删除行), 'merge_cells'(合并单元格)"
            },
            tableIndex: { type: "integer", description: "目标表格索引序号(1-based)，默认 1" },
            rows: { type: "integer", description: "行数（插入时使用）" },
            columns: { type: "integer", description: "列数（插入时使用）" },
            data: { type: "array", items: { type: "array" }, description: "填充到表格的二维数据矩阵，行列数不足时按 write_matrix 规则动态扩容" },
            stylePreset: { type: "string", enum: ["mckinsey_three_line", "clean_minimal", "none"], description: "样式预设: 'mckinsey_three_line'(麦肯锡三线表，顶底粗线+表头细线+无内部竖线，默认), 'clean_minimal'(当前未实现独立样式，等同 'none'), 'none'(不套用预设样式)" },
            repeatHeader: { type: "boolean", description: "跨页时是否自动重复表头首行，默认 true" },
            cellRow: { type: "integer", description: "单元格行号(1-based)，action='format_cell' 时使用" },
            cellColumn: { type: "integer", description: "单元格列号(1-based)，action='format_cell' 时使用" },
            cellFormat: {
              type: "object",
              properties: {
                bold: { type: "boolean", description: "该单元格是否加粗" },
                fontSizePt: { type: "number", description: "该单元格字号(磅值)" },
                fontName: { type: "string", description: "该单元格字体名，如 '仿宋_GB2312'" },
                backgroundColor: { type: "string", description: "十六进制底色，例如 '#F1F5F9'" }
              },
              description: "单元格排版参数"
            },
            rowIndex: { type: "integer", description: "删除或操作的行号(1-based)" },
            mergeRange: {
              type: "object",
              properties: {
                startRow: { type: "integer", description: "合并起始行(1-based)" },
                startCol: { type: "integer", description: "合并起始列(1-based)" },
                endRow: { type: "integer", description: "合并结束行(1-based，含)" },
                endCol: { type: "integer", description: "合并结束列(1-based，含)" }
              },
              description: "合并单元格范围"
            }
          },
          required: ["action"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_review_and_comments",
        description: "Word 审阅与修订控制。支持开启/关闭修订记录模式（Track Changes）、一键接受/拒绝全部修订、插入与读取批注。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称" },
            action: {
              type: "string",
              enum: ["enable_track_changes", "disable_track_changes", "accept_all_revisions", "reject_all_revisions", "add_comment", "list_comments"],
              description: "审阅动作: 'enable_track_changes'/'disable_track_changes'(开启/关闭修订记录), 'accept_all_revisions'/'reject_all_revisions'(接受/拒绝全部修订，不可撤销), 'add_comment'(插入批注), 'list_comments'(读取批注列表)"
            },
            commentText: { type: "string", description: "添加批注时的批注内容，action='add_comment' 时必传" },
            author: { type: "string", description: "批注作者名称" }
          },
          required: ["action"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_page_layout_and_watermark",
        description: "Word 页面版式与水印设置：页眉页脚文本（支持奇偶页不同、首页不同）、页码域格式与倾斜半透明文字水印。页眉/页脚/页码**遍历文档全部节**逐节写入，返回 appliedSections 说明每节实际写了什么，失败项收进 warnings。水印：宿主 WPS for Mac 无法把形状放进页眉层（Headers.Shapes 的写入会静默落到正文层），因此水印仍是正文层浮动图形、只在第 1 页渲染，不是“每页可见”；跨页水印需在 Word 内手动插入（插入 → 水印）或改用 Windows/COM 通道，详情见返回的 warnings。传 pageNumberFormat 时会**覆盖**同一次调用里 footerText 写的页脚内容。三项至少传一项，否则直接报错。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称" },
            headerText: { type: "string", description: "页眉文本内容（覆盖各节原有页眉文本，会清掉页眉里已有的域）" },
            footerText: { type: "string", description: "页脚文本内容；同时传 pageNumberFormat 时会被页码域覆盖" },
            pageNumberFormat: { type: "string", enum: ["dash", "simple", "page_of_pages"], description: "页脚页码格式: 'simple'(仅页码，用真 PAGE 域实现，**当前宿主唯一可用的取值**), 'dash'(形如 - 1 -), 'page_of_pages'(形如 1 / 5)。后两者在 WPS for Mac 上写入会被宿主静默丢弃（页脚只支持纯页码域），工具会逐节返回 applied:false 与中文原因，不会静默降级成纯页码。" },
            watermarkText: { type: "string", description: "倾斜背景文字水印，例如 '内部机密 严禁外传'；落在正文层、只渲染第 1 页（见工具说明）" },
            watermarkColor: { type: "string", description: "水印文字颜色十六进制，默认 '#C0C0C0'" },
            differentFirstPage: { type: "boolean", description: "是否首页不同页眉页脚（文档级设置，不写首页页眉内容）" },
            differentOddEvenPages: { type: "boolean", description: "是否奇偶页不同页眉页脚（文档级设置，只写奇数页页眉/页脚条目）" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_word_find_and_replace",
        description: "Word 全局查找与精准替换。支持通配符、全字匹配、区分大小写以及指定替换文本的格式（防止加粗污染）。",
        parameters: {
          type: "object",
          properties: {
            documentName: { type: "string", description: "目标文档名称" },
            searchQuery: { type: "string", description: "要搜索的文本" },
            replaceText: { type: "string", description: "要替换为的新文本，不传则仅执行查找定位" },
            matchCase: { type: "boolean", description: "是否区分大小写，默认 false" },
            matchWholeWord: { type: "boolean", description: "是否全字匹配，默认 false" },
            useWildcards: { type: "boolean", description: "是否使用通配符，默认 false" },
            scope: { type: "string", enum: ["full", "selection"], description: "替换范围: 'full'(全文，默认), 'selection'(选区)" },
            replaceFormatting: {
              type: "object",
              properties: {
                bold: { type: "boolean", description: "替换后文本是否加粗（明确设为 false 避免继承前文加粗）" },
                italic: { type: "boolean", description: "替换后文本是否斜体" },
                fontSizePt: { type: "number", description: "替换后文本字号(磅值)" },
                fontName: { type: "string", description: "替换后文本字体名" }
              },
              description: "指定替换后新文字的格式"
            }
          },
          required: ["searchQuery"],
          additionalProperties: false
        }
      }
    }
  ];
}
