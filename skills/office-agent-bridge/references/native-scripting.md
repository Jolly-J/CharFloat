# 原生脚本能力扩展

专用工具 → 只读探测 → 原生脚本 → 读回和预览。没有专用工具不能直接判定不支持。脚本是正常的扩展通道，无需仅因使用 JS 再次要求用户授权；操作仍须在用户任务范围内。

WPS 使用 `wps_execute_script`，明确传 component 和精确 workbookName/documentName/presentationName。上下文为 app、wb、doc、pres、wps、params 和 console；支持 async/await。Microsoft 应使用实际工具清单中的脚本入口和对应对象模型，不可照搬 WPS 示例。

可用 `wps_inspect_api` 检查 `wb.Worksheets.Item(1).Shapes` 或 `pres.PageSetup`。部分宿主对象无法完整枚举成员，空列表不是 API 不支持的证据，可针对具体属性进行只读访问。捕获具体异常，区分 API 缺失、目标错误、参数错误和连接问题。

## Excel 原生矢量绘图

Shapes 是独立于单元格和统计图表的绘图层，可用于流程图、标注、信息卡片。下例是 WPS 对象模型示例，使用前检查当前宿主的 Shapes API。坐标为 pt，依据目标单元格 Left/Top/Width/Height 定位，不是行列号或屏幕像素。

只读探测代码：

```js
const sheet = wb.Worksheets.Item(params.sheetName);
const anchor = sheet.Range(params.address);
return { workbookName: wb.Name, sheetName: sheet.Name,
  left: anchor.Left, top: anchor.Top, width: anchor.Width, height: anchor.Height,
  shapeCount: sheet.Shapes.Count };
```

绘制可编辑矩形的代码（params 传 sheetName、address、shapeName、text）：

```js
const sheet = wb.Worksheets.Item(params.sheetName);
const anchor = sheet.Range(params.address);
const shape = sheet.Shapes.AddShape(1, anchor.Left, anchor.Top, 180, 72);
shape.Name = params.shapeName;
shape.TextFrame.Characters().Text = params.text;
return { name: shape.Name, left: shape.Left, top: shape.Top,
  width: shape.Width, height: shape.Height, shapeCount: sheet.Shapes.Count };
```

连接线、自由曲线、组合可继续探测 AddConnector、BuildFreeform、Range(...).Group 等宿主 API。不要把截图当作可编辑矢量。样式、文字接口可能随宿主版本不同；创建后若后续设置失败，先检查已创建对象再继续，不能重新创建整套。

## 编辑已有对象

没有 update_chart 工具时，先读出现有对象名称和属性，再用脚本修改该对象。不能据此直接判定必须删除重建。返回普通 JSON：文档名、对象名称/ID、实际位置、字号、修改后的属性；不要返回整个宿主对象。

PPT 使用 pres.PageSetup.SlideWidth/SlideHeight 获取尺寸。按页面比例计算几何，所有结果转为 pt 后写入；字体与版面一起考虑。先读现有形状，设置文本框内边距、换行和字号，读回后检查真实预览。需要拆页或改动内容时结合用户目标判断，不用无限缩小字体掩盖溢出。

## 批量执行与返回

把一页、一张表或一组相关形状作为批次边界。用 params 传操作列表，在脚本内循环，不为每个对象单独发起 MCP 调用。所需 API 已在当前宿主验证时复用，无需每次反射。

执行前校验整批目标和输入；每步成功后记录对象标识。脚本可返回以下普通 JSON 结构（这是脚本自定义返回格式，不是新增 MCP 参数）：

```json
{"completed":[{"id":"实际对象标识","left":50,"top":30}],"failed":null,"pendingCount":0}
```

失败时记录当前对象、步骤和异常文本，并停止依赖该步骤的修改。completed 中使用宿主读回值；不返回完整宿主对象、全文或无关数据。Bridge 的脚本包装层可能仍返回 success，调用方必须同时检查 returnValue 中的 failed/pendingCount。响应超时导致整个结果丢失时，按目标标识重新读取，不依靠内存完成列表猜测。

不要为省调用合并本来需要人工决策的内容，也不要将需要审计回滚的单元格写入移到脚本。
