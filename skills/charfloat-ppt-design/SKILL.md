---
name: charfloat-ppt-design
description: 通过「字浮 CharFloat」在已打开的演示文稿中按页批量排版、编辑文本形状表格和图表，并验证字号、几何和真实预览；支持原生 JS 精细编辑。
---

# PPT 页面工作流

复用 [共享流程](../charfloat/SKILL.md) 的宿主、目标和工具上下文；只有缺失时初始化。

先按 [现稿样式协议](style-protocol.md) 采样用户参考页并记录约束，随后执行下述页面流程。

## 从结构到页面

先用 `wps_ppt_read_presentation` 获取页面尺寸与全稿摘要，只对目标页使用 `wps_ppt_get_slide_shapes` 读取形状、字号与几何。已有足够信息时不反复读取全稿。读取接口的文本摘要可能截断；需要完整改稿时通过目标对象脚本读取完整文本。

局部修改只处理指定对象。批量新建或统一排版时，先完成一个代表页并预览，通过后将相同规则应用到其他页；不同版式分别验证。使用已有批量生成工具或以一页为单位执行 JS，不逐个文字框发起 MCP 调用。页面复杂或操作风险高时拆成可核验的小批次。

## 布局规则

- 读取真实 pageWidth/pageHeight，几何和字号单位为 pt。16:9 不等于固定 720×405，截图像素不能直接作为文稿坐标。
- 根据内容划分页面区域，再计算位置和尺寸。新增文本框显式设置 left/top/width/height/fontSize，避免重复使用默认坐标。
- 字号与页面尺寸和容量一起确定。720×405 pt 页面可从标题 24–36 pt、正文 16–22 pt 起步，其他尺寸相应调整；这不是固定限制。
- 将换行、内边距、段落间距和 AutoSize 纳入布局。长文本优先调整可用区域；涉及精简内容或拆页时保持用户意图，不无限缩小字体。

精细设置使用 [原生脚本](../charfloat/references/native-scripting.md)；形状类型、图表类型与版式都是枚举，必须传整数，取值见 [枚举速查表](../charfloat/references/enumeration.md)。页面批次返回 slideIndex、形状 ID、实际几何、字号及失败步骤；避免返回全部文本或整个宿主对象。

## 空白背景圆角

macOS WPS 可用 `wps_ppt_set_corner_radius` 设置明确点名的顶层可见空白原生 roundRect 背景。先 `action:"plan"`，传 `presentationName`、`slideIndex`、`shapeNames`、`radiusPt`，审阅每个尺寸和计算半径；再用返回的同一 `operationId` 调 `action:"apply"` 并传同一组目标和半径。半径单位为 pt，不是 adj 或百分比，必须小于等于每个目标短边的一半。

该路径从独立 PPTX 参数副本导入原生形状，会生成新 ID；旧背景改名隐藏保留，原可见层级保持。工具拒绝文字、组合、旋转、动画/连接器身份引用等未支持对象。不要用 Adjustments getter 或“调用未报错”证明圆角生效。其他平台尚不支持该工具。

超时或部分失败先用同一文稿、页码、operationId 调 `action:"status"`，禁止换 ID 重放；`action:"rollback"` 仅在相关新旧对象没有后续修改时恢复本次对象并保存，不用整稿备份覆盖后续编辑。加载项重载后宿主回执丢失时拒绝自动回退，保留 `backupPath` 和桥接记录供人工审阅。没有自动删除备份或隐藏旧背景的行为。

分别核对 `liveApplied`、`saved`、`ooxmlVerified`，打开 `previewPath` 检查真实画面。`previewVerified:false` 表示仍需调用者做视觉检查；工具不会因导出 PNG 成功就宣称画面验收通过。PPT 保存会连同其他在途编辑一起落盘。

## 验证与继续

处理 `layoutWarnings`，再检查真实预览中的越界、遮挡、换行、留白和母版重复元素。每个修改页完成后检查一次，修正后复查受影响页，不在每个形状创建后截图。没有警告也不等于视觉通过。

图表需读回类型和系列数据；文字框或图片不能当作可编辑图表。创建后配置失败时先定位已创建对象，不重放整页。

复杂母版、动画与嵌入数据按实际宿主 API 验证。既有 macOS 经验提示图表数据编辑窗口可能导致崩溃；没有当前版本验证依据时，不调用 ChartData.ActivateChartDataWindow 或 ChartData.Workbook 试探。

预览失败或保存未执行时明确报告。PPT 没有统一审计回滚；圆角工具的局部回退仅覆盖自身回执中的新旧背景。
