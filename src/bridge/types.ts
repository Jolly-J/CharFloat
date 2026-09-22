/**
 * WPS Bridge 类型定义中心
 */

export interface SheetInfo {
  index: number;
  name: string;
  visible: boolean;
}

export interface SelectionInfo {
  sheetName?: string;
  address: string;
  rowCount: number;
  columnCount: number;
}

export interface WorkspaceSummary {
  hasOpenWorkbook: boolean;
  workbookName?: string;
  fullName?: string;
  activeSheetName?: string;
  sheetCount?: number;
  sheets?: SheetInfo[];
  selection?: SelectionInfo | null;
  message?: string;
}

export interface SheetOutline {
  sheetName: string;
  isEmpty: boolean;
  usedRangeAddress: string;
  startRow?: number;
  startColumn?: number;
  rowCount: number;
  columnCount: number;
  headerPreview: any[][];
}

export interface RangeData {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  values: any[][];
  formulas: (string | null)[][] | null;
  numberFormats?: string[][] | null;
}

export type RangeStyleField =
  | "fontName"
  | "fontSize"
  | "bold"
  | "fontColor"
  | "backgroundColor"
  | "numberFormat"
  | "horizontalAlignment"
  | "verticalAlignment"
  | "wrapText"
  | "rowHeight"
  | "columnWidth"
  | "merged"
  | "mergeArea"
  | "borders";

export interface GetRangeStylesParams {
  sheetName?: string;
  workbookName?: string;
  address: string;
  mode?: "summary" | "cells";
  include?: RangeStyleField[];
  maxCells?: number;
}

export interface CellMatch {
  address: string;
  row: number;
  column: number;
  value: any;
}

export interface SearchResult {
  sheetName: string;
  query: string;
  totalFound: number;
  matches: CellMatch[];
}

export interface CellDiff {
  cell: string;
  oldValue: any;
  newValue: any;
  oldFormula: string | null;
  newFormula: string | null;
}

export interface RangeSnapshot {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  values: any[][];
  formulas: (string | null)[][];
}

export interface PatchResult {
  sheetName: string;
  address: string;
  modifiedCount: number;
  diff: CellDiff[];
  beforeSnapshot: RangeSnapshot;
  afterSnapshot: RangeSnapshot;
}

// 留痕记录对象
export interface AuditRecord {
  host?: "wps" | "microsoft";
  id: string;
  timestamp: number;
  clientName: string; // 例如 "Cursor", "Claude Code", "Agent", "Manual"
  /**
   * 产生该记录的会话标识（ISS-49）。
   *
   * `clientName` 可以相同（例如两个 stdio 客户端都叫 "stdio MCP"），
   * 会话标识才是"这条是谁改的"的稳定判据；HTTP /mcp 与 stdio 通道都会带上它。
   */
  sessionId?: string;
  /**
   * 操作类型。`update_values` / `update_formulas` 是唯一带值快照、可回滚的两类；
   * 其余是"只留痕、不可回滚"的操作事实（ISS-48）。
   */
  actionType:
    | "update_values" | "update_formulas" | "rollback"
    | "format" | "conditional_format" | "freeze_panes" | "insert_col" | "insert_row"
    | "delete_dimension" | "hide_dimension" | "sheet_structure" | "chart" | "data_validation"
    | "find_replace" | "pivot_table" | "filter_sort" | "comment" | "save" | "script";
  description: string;
  workbookName: string;
  sheetName: string;
  address: string;
  modifiedCount: number;
  diff: CellDiff[];
  /** 修改前快照；只有可回滚的记录才有（其余记录为 undefined，ISS-48）。 */
  beforeSnapshot?: RangeSnapshot;
  afterSnapshot?: RangeSnapshot;
  /**
   * 该记录能否用 `wps_rollback` 回滚。
   * 只有 patch_cells 的值/公式修改为 true；其余写操作只留痕不提供回滚（ISS-48）。
   */
  rollbackable?: boolean;
  status: "applied" | "rolled_back";
}

// RPC 协议定义
export interface RpcRequest {
  id: string;
  method: string;
  params?: any;
}

export interface RpcResponse {
  type: "rpc_response";
  id: string;
  result?: any;
  error?: string;
}

export interface BridgeEventPacket {
  type: "event";
  event: string;
  data: any;
}

// ==========================================
// 第一梯队：排版美学与结构刚需
// ==========================================

export interface FormatCellsParams {
  sheetName?: string;
  workbookName?: string;
  address: string;
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  fontColor?: string;
  backgroundColor?: string;
  horizontalAlignment?: "left" | "center" | "right";
  verticalAlignment?: "center" | "top" | "bottom";
  numberFormat?: string;
  rowHeight?: number;
  wrapText?: boolean;
  borders?: string | boolean;
  merge?: boolean;
  unmerge?: boolean;
}

export type ConditionalRuleType = "cell_value" | "data_bar" | "color_scale";
export type ConditionalOperator = "less_than" | "greater_than" | "equal" | "between";

export interface ConditionalFormattingParams {
  sheetName?: string;
  workbookName?: string;
  address: string;
  ruleType?: ConditionalRuleType;
  operator?: ConditionalOperator;
  formula1?: string | number;
  formula2?: string | number;
  backgroundColor?: string;
  fontColor?: string;
  barColor?: string;
  colorScaleMin?: string;
  colorScaleMax?: string;
  clearExisting?: boolean;
}

export interface FreezePanesParams {
  sheetName?: string;
  workbookName?: string;
  freezeRowIndex?: number;
  freezeColumnIndex?: number;
  unfreeze?: boolean;
}

export type DimensionTargetType = "row" | "column";
export type RowColumnAction = "insert" | "delete" | "hide" | "unhide";

export interface ModifyRowsColumnsParams {
  sheetName?: string;
  workbookName?: string;
  targetType: DimensionTargetType;
  action: RowColumnAction;
  index: number;
  count?: number;
}

export interface ColumnRule {
  colIndex: number;
  minWidth?: number;
  maxWidth?: number;
  wrapText?: boolean;
}

export interface AutoFitColumnsParams {
  sheetName?: string;
  workbookName?: string;
  columnRules?: ColumnRule[];
}

// ==========================================
// 第二梯队：数据可视化图表引擎
// ==========================================

export type ChartType =
  | "line"
  | "column_clustered"
  | "bar_clustered"
  | "pie"
  | "doughnut"
  | "pareto";

export interface ChartPosition {
  leftCell: string;
  width?: number;
  height?: number;
}

export interface ChartAxisSettings {
  min?: number;
  max?: number;
  step?: number;
  numberFormat?: string;
  title?: string;
}

export interface ChartSeriesSetting {
  seriesIndex: number; // 从 1 开始
  color?: string;
  smooth?: boolean;
  name?: string;
}

export interface AddChartParams {
  sheetName?: string;
  workbookName?: string;
  chartType: ChartType;
  dataRange?: string;
  dataRanges?: string[]; // 支持非连续区域数组，如 ['A4:A19', 'E4:E19']
  title?: string;
  position?: ChartPosition;
  hasLegend?: boolean;
  hasDataLabels?: boolean;
  smoothLine?: boolean;
  seriesColors?: string[];
  yAxis?: ChartAxisSettings;
  seriesSettings?: ChartSeriesSetting[];
  replaceExisting?: boolean; // 是否自动清理同一锚点的旧图表，默认 true
}

export interface GetChartsParams {
  sheetName?: string;
  workbookName?: string;
  shapeName?: string;
  chartIndex?: number;
  chartTitle?: string;
  detail?: boolean;
}

export interface DeleteChartParams {
  sheetName?: string;
  workbookName?: string;
  shapeName?: string;
  chartTitle?: string;
  leftCell?: string;
  chartIndex?: number;
  clearAll?: boolean;
}

// ==========================================
// 第三梯队：高阶数据分析与交互组件
// ==========================================

export type PivotSummaryFunction = "sum" | "count" | "average" | "max" | "min";

export interface PivotDataField {
  fieldName: string;
  summaryFunction?: PivotSummaryFunction;
  caption?: string;
}

export interface CreatePivotTableParams {
  workbookName?: string;
  sourceSheetName?: string;
  sourceRange: string;
  destSheetName?: string;
  destCell: string;
  rowFields?: string[];
  columnFields?: string[];
  dataFields?: PivotDataField[];
}

export interface SortRule {
  colIndex: number;
  order: "asc" | "desc";
}

export interface SetFilterAndSortParams {
  sheetName?: string;
  workbookName?: string;
  range: string;
  enableAutoFilter?: boolean;
  sortRules?: SortRule[];
}

export type ValidationType = "list" | "number_range";

export interface SetDataValidationParams {
  sheetName?: string;
  workbookName?: string;
  address: string;
  validationType: ValidationType;
  listItems?: string[];
  operator?: "between" | "greater_than" | "less_than" | "equal";
  minVal?: number;
  maxVal?: number;
  promptTitle?: string;
  promptMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
}

export type SheetAction = "rename" | "move" | "tab_color" | "protect" | "unprotect";

export interface ManageSheetParams {
  sheetName: string;
  workbookName?: string;
  action: SheetAction;
  newName?: string;
  targetIndex?: number;
  color?: string;
  password?: string;
}

// ==========================================
// 第五梯队：四大进阶杀手级能力
// ==========================================

// 1. 行列级综合结构控制
export interface ManageRowsAndColumnsParams {
  sheetName?: string;
  workbookName?: string;
  targetType: "row" | "column";
  action: "insert" | "delete" | "hide" | "unhide" | "set_size";
  index: number | string;
  count?: number;
  size?: number;
}

// 2. 单元格批注与审阅备注
export type CommentAction = "add" | "read" | "delete" | "clear_all";

export interface ManageCellCommentsParams {
  sheetName?: string;
  workbookName?: string;
  address: string;
  action: CommentAction;
  text?: string;
  author?: string;
}

// 3. 全局查找与快速定位
export interface FindAndReplaceParams {
  sheetName?: string;
  workbookName?: string;
  searchQuery: string;
  replaceText?: string;
  matchCase?: boolean;
  matchEntireCell?: boolean;
  searchRange?: string;
  maxResults?: number;
}

export interface FindResultItem {
  address: string;
  row: number;
  col: number;
  value: any;
  formula?: string | null;
  replaced?: boolean;
}

// 4. 工作表复制与模板克隆
export interface DuplicateSheetParams {
  sourceSheetName: string;
  newSheetName: string;
  position?: "after" | "before" | "end";
  workbookName?: string;
}

// ==========================================
// Word (文字) 模块类型定义 (8大高聚合工具)
// ==========================================

export interface WordReadParams {
  documentName?: string;
  scope?: "outline" | "full" | "selection" | "paragraphs" | "tables";
  maxParagraphs?: number;
}

export interface WordOutlineItem {
  level: number;
  text: string;
  paragraphIndex: number;
}

export interface WordReadResult {
  documentName: string;
  pageCount?: number;
  paragraphCount: number;
  wordCount?: number;
  outline: WordOutlineItem[];
  previewText?: string;
  selectionText?: string;
  tableCount: number;
}

export interface WordWriteParams {
  documentName?: string;
  location?: "end" | "start" | "selection" | "after_heading" | "bookmark";
  targetHeading?: string;
  targetBookmark?: string;
  type?: "paragraph" | "heading1" | "heading2" | "heading3" | "bullet_list" | "quote" | "code_block";
  content: string | string[];
}

export interface WordFormatParams {
  documentName?: string;
  preset?: "gov_standard" | "business_modern" | "academic" | "custom";
  fontName?: string;
  fontSizePt?: number;
  lineSpacingPt?: number;
  firstLineIndentChars?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  margins?: {
    topMm?: number;
    bottomMm?: number;
    leftMm?: number;
    rightMm?: number;
  };
}

export interface WordTocParams {
  documentName?: string;
  upperHeadingLevel?: number; // 默认 1
  lowerHeadingLevel?: number; // 默认 3
  insertLocation?: "start" | "end" | "selection";
  includePageNumbers?: boolean;
}

export interface WordTableParams {
  documentName?: string;
  action: "insert" | "update_data" | "format_style" | "merge_cells";
  tableIndex?: number;
  rows?: number;
  columns?: number;
  data?: any[][];
  stylePreset?: "mckinsey_three_line" | "modern_blue" | "clean_minimal" | "none";
  repeatHeader?: boolean;
  mergeRange?: { startRow: number; startCol: number; endRow: number; endCol: number };
}

export interface WordReviewParams {
  documentName?: string;
  action: "enable_track_changes" | "disable_track_changes" | "accept_all_revisions" | "reject_all_revisions" | "add_comment" | "list_comments";
  commentText?: string;
  author?: string;
}

export interface WordPageLayoutParams {
  documentName?: string;
  headerText?: string;
  footerText?: string;
  pageNumberFormat?: "dash" | "simple" | "page_of_pages"; // 如 "- 1 -" 或 "1"
  differentFirstPage?: boolean;
  differentOddEvenPages?: boolean;
  watermarkText?: string;
  watermarkColor?: string;
}

export interface WordFindReplaceParams {
  documentName?: string;
  searchQuery: string;
  replaceText?: string;
  matchCase?: boolean;
  matchWholeWord?: boolean;
  useWildcards?: boolean;
  scope?: "full" | "selection";
}

// ==========================================
// PowerPoint (演示) 模块类型定义 (7大高聚合工具)
// ==========================================

export interface PptReadParams {
  presentationName?: string;
  includeNotes?: boolean;
  maxSlides?: number;
}

export interface PptSlideOutline {
  index: number;
  title: string;
  layoutName: string;
  shapeCount: number;
  notes?: string;
  textSnippets: string[];
}

export interface PptReadResult {
  presentationName: string;
  slideCount: number;
  activeSlideIndex: number;
  slides: PptSlideOutline[];
}

export interface PptSlideSpec {
  layout: "title" | "cards_2" | "cards_3" | "cards_4" | "chart" | "content" | "end";
  title: string;
  subtitle?: string;
  bulletPoints?: string[];
  cards?: {
    tag?: string;
    title: string;
    description: string;
    accentColor?: string;
  }[];
  chart?: {
    chartType: "column" | "line" | "bar" | "pie";
    title?: string;
    categories: string[];
    series: { name: string; values: number[] }[];
  };
  notes?: string;
}

export interface PptGenerateDeckParams {
  presentationName?: string;
  themeColor?: string; // 默认深蓝商务 #0f4c81 或科技紫 #4b38b3
  themePreset?: "business_blue" | "tech_purple" | "clean_light" | "dark_elegant";
  slides: PptSlideSpec[];
}

export interface PptManageSlidesParams {
  presentationName?: string;
  action: "add" | "delete" | "move" | "duplicate" | "set_layout" | "set_background";
  slideIndex?: number;
  targetIndex?: number;
  layoutIndex?: number;
  backgroundColor?: string;
}

export interface PptBusinessCardsParams {
  presentationName?: string;
  slideIndex?: number;
  columnCount: 2 | 3 | 4;
  cards: {
    tag?: string;
    title: string;
    description: string;
    accentColor?: string;
  }[];
  topY?: number;
  cardHeight?: number;
}

export interface PptChartParams {
  presentationName?: string;
  slideIndex?: number;
  chartType: "column" | "line" | "bar" | "pie";
  title?: string;
  categories: string[];
  series: {
    name: string;
    values: number[];
  }[];
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

export interface PptShapeMediaParams {
  presentationName?: string;
  slideIndex?: number;
  action: "add_textbox" | "add_shape" | "insert_image" | "update_shape" | "delete_shape";
  shapeId?: string | number;
  shapeType?: "rectangle" | "rounded_rectangle" | "oval" | "arrow";
  text?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  fillColor?: string;
  lineColor?: string;
  imageBase64?: string;
  imageUrl?: string;
  zOrder?: "bring_to_front" | "send_to_back";
}

export interface PptCapturePreviewParams {
  presentationName?: string;
  slideIndex?: number;
}


