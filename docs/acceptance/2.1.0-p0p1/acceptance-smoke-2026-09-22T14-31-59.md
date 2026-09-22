# 验收冒烟测试报告

- 运行时间：2026/9/22 22:31:59
- 目标工作簿：AI验收冒烟.xlsx（临时表 __终验__，结束时删除）
- 结果：**43/43 通过**

## 分组结果

| 组 | 通过 | 状态 |
|---|---|---|
| CAP-07 | 10/10 | ✔ |
| CAP-15 | 1/1 | ✔ |
| CAP-16 | 2/2 | ✔ |
| CAP-17 | 2/2 | ✔ |
| CAP-18 | 2/2 | ✔ |
| CAP-19 | 2/2 | ✔ |
| CAP-20 | 2/2 | ✔ |
| CAP-30 | 1/1 | ✔ |
| CAP-31 | 1/1 | ✔ |
| CAP-32 | 1/1 | ✔ |
| CAP-33 | 3/3 | ✔ |
| CAP-34 | 1/1 | ✔ |
| CAP-36 | 1/1 | ✔ |
| CAP-35 | 1/1 | ✔ |
| CAP-21 | 1/1 | ✔ |
| CAP-22 | 2/2 | ✔ |
| CAP-23 | 1/1 | ✔ |
| CAP-40 | 2/2 | ✔ |
| 清理 | 2/2 | ✔ |
| 自检 | 5/5 | ✔ |

## 逐项明细（含真机读数）

### CAP-07

- ✔ **add_shape 几何形状 + 读回填充/文字**
  - 读数：`{"name":"v_title","type":1,"autoShapeType":5,"left":400,"top":20,"width":160,"height":70,"rotation":0,"visible":true,"fillColor":"`
- ✔ **add_shape 文本框**
  - 读数：`name=v_text 文字="文本框"`
- ✔ **add_shape 直线（MS 侧不支持，WPS 支持）**
  - 读数：`name=v_line 类型=-2`
- ✔ **add_shape 艺术字（WPS 独有）**
  - 读数：`name=v_art 宽=97.85 高=44.45`
- ✔ **list_shapes 读回全部形状**
  - 读数：`形状数=4`
- ✔ **update_shape 改位置/填充并读回**
  - 读数：`{"name":"v_title","type":1,"autoShapeType":5,"left":420,"top":20,"width":160,"height":70,"rotation":0,"visible":true,"fi`
- ✔ **group_shapes 分组 + 读回成员**
  - 读数：`{"success":true,"workbookName":"AI验收冒烟.xlsx","sheetName":"__终验__","group":{"name":"v_group","type":6,"autoShapeType":-2,`
- ✔ **ungroup_shapes 解组**
  - 读数：`{"success":true,"workbookName":"AI验收冒烟.xlsx","sheetName":"__终验__","released":["v_title","v_text"],"releasedStillPresent"`
- ✔ **set_shape_zorder 层级**
  - 读数：`applied=bringToFront 形状序列=["v_text","v_line","v_art","v_title"]`
- ✔ **错误处理：未知形状类型列出可用值**
  - 读数：`不认识的形状类型 "不存在的形状"。可用：rectangle / rounded_rectangle / oval / ellipse / diamond / triangle / right_triangle / pe`

### CAP-15

- ✔ **copy_range 复制并读回左上角**
  - 读数：`{"success":true,"workbookName":"AI验收冒烟.xlsx","source":"__终验__!A1:C5","dest":"__终验__!E1","copyType":"all","readBackFirstC`

### CAP-16

- ✔ **超链接 添加 + 从锚点读回**
  - 读数：`读回文字=WPS官网`
- ✔ **超链接 列出（地址/文字/锚点）**
  - 读数：`count=1`

### CAP-17

- ✔ **命名区域 添加 + 读回 refersTo**
  - 读数：`refersTo=="__终验__!$B$2:$B$5"`
- ✔ **命名区域 列出**
  - 读数：`count=1`

### CAP-18

- ✔ **文档属性 写入 + 自动读回核对**
  - 读数：`{"Title":"终验-标题","Subject":"","Author":"终验-作者","Keywords":"","Comments":"","Category":"","Company":"","Manager":""}`
- ✔ **文档属性 读回 Title**
  - 读数：`Title=终验-标题 自定义="v1"`

### CAP-19

- ✔ **结构化表格 创建 + 读回列名**
  - 读数：`["产品","销量","单价"]`
- ✔ **结构化表格 列出**
  - 读数：`count=1`

### CAP-20

- ✔ **图片 插入 + 读回几何**
  - 读数：`{"name":"finProbePic","kind":"picture","left":700,"top":19.2,"width":60,"height":57.6,"topLeftCell":"$O$2"}`
- ✔ **图片 列出**
  - 读数：`count=1`

### CAP-30

- ✔ **条件格式读回（类型/运算符/填充色）**
  - 读数：`{"index":1,"type":"cell_value","typeCode":1,"operator":"greater_than","formula1":"=100","formula2":null,"priority":1,"stopIfTrue":`

### CAP-31

- ✔ **冻结窗格读回（frozen + 行列）**
  - 读数：`frozen=true 行=2 列=2`

### CAP-32

- ✔ **数据有效性读回（类型 + 候选项）**
  - 读数：`{"address":"$B$2:$B$5","type":"list","typeCode":3,"operator":"between","formula1":"进行中,已完成,未开始","formula2":"","inCellDropdown":true,"ignoreBlank":true`

### CAP-33

- ✔ **普通区域：筛选写入 + 读回范围**
  - 读数：`applied=$A$1:$C$5 告警=[]`
- ✔ **筛选状态读回（范围 + 列条件）**
  - 读数：`范围=$A$1:$C$5 条件列=0`
- ✔ **结构化表格上设筛选：如实告警不假成功**
  - 读数：`筛选**未生效**：工作表上存在结构化表格 finFilterTable，它自带筛选器，表级 AutoFilter 不会另外打开。请直接在表格上操作，或先把表格转为普通区域。`

### CAP-34

- ✔ **工作表状态读回（保护 + 标签色）**
  - 读数：`保护=true 标签色=#0F9D58`

### CAP-36

- ✔ **透视表读回（名称/数据源/字段）**
  - 读数：`{"n":"PivotTable_1790087519556","src":"=__终验__!R1C16:R3C17"}`

### CAP-35

- ✔ **Word 页眉页脚与水印读回**
  - 读数：`跳过：无打开的 Word 文档`

### CAP-21

- ✔ **update_chart 改标题 + 读回**
  - 读数：`{"name":"finProbeChart","chartType":51,"title":"终验图表","hasTitle":true,"hasLegend":true,"seriesCount":0,"dataLabels":null}`

### CAP-22

- ✔ **export_chart_image 落盘核对（应 true）**
  - 读数：`fileWritten=true 大小=638 字节`
- ✔ **不可达路径应如实报未落盘（不假成功）**
  - 读数：`fileWritten=false warn=宿主已执行导出，但文件不存在：/tmp/fin-should-not-exist.png。请确认 outputPath 落在 WPS 可写目`

### CAP-23

- ✔ **标签色 写入后读回新值**
  - 读数：`标签色=#E4572E`

### CAP-40

- ✔ **clear_range 清空并返回地址**
  - 读数：`clearedAddress=$E$1:$F$2 模式=all`
- ✔ **清空后读回确认为空**
  - 读数：`[[null,null],[null,null]]`

### 清理

- ✔ **删除本次写入的自定义文档属性**
  - 读数：`removed=["finProbeProp"] 剩余=3`
- ✔ **删除测试命名区域**
  - 读数：`stillExists=false 剩余=1`

### 自检

- ✔ **隔离工作表已删除**
  - 读数：`残留: `
- ✔ **无测试命名区域残留**
  - 读数：`残留: `
- ✔ **无测试自定义属性残留**
  - 读数：`残留: `
- ✔ **文档 Title/Author 已还原**
  - 读数：`Title=undefined Author=undefined`
- ✔ **临时图片/导出文件已删除**
  - 读数：`pic=false chart=false`
