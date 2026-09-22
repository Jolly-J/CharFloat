# 11 PPT/Word 宿主实现修复台账（P5 / mcp-sweep）

- 执行者：PPT/Word 宿主实现修复员（子代理）
- 目标：ISS-44 / 55 / 57 / 58 / 67 / 70 / 72 / 73 / 76 / 78 / 79 / 85 / 86
- 写区（**只有这四个文件**）：
  - `wps-addon/src/ppt.js`
  - `wps-addon/src/word.js`
  - `src/bridge/tools/definitions/ppt.ts`
  - `src/bridge/tools/definitions/word.ts`
- 基线：`tools-snapshot.p6.json`（工具数 91）。
- ⚠️ 运行中的 WPS 加载项是**旧构建**，本轮所有宿主结论来自
  **`wps_execute_script` 隔离探针 + 源码复核**，不来自"调用新工具"。

---

## 0. 触达边界（决定了几条只能改说明）

网关处理器（`src/bridge/gateway/{ppt,word}.ts`）与网关注册表（`src/bridge/gateway.ts`）
**不在写区**，于是：

| 想要 | 能不能做 | 原因 |
|---|---|---|
| 改预览落盘判定 / `PPT 未生成预览` 兜底文案 | ❌ | 在 `gateway/ppt.ts:154`、`gateway/word.ts` |
| 新增工具名 `wps_ppt_save_presentation` | ❌ | 新工具必须在 `gateway.ts` 的 `HANDLERS` 注册，否则 `tests/contract-consistency.test.ts` 的"注册工具缺少网关分支"失败 |
| 给现有工具加参数 | ⚠️ 部分 | 只有网关**已转发**的字段才真正到达宿主；其他字段会被 `gateway/*.ts` 丢掉 |
| 给现有工具加 enum 取值 / 改说明 | ✅ | 纯 schema/描述，不涉及网关 |

**因此本轮的落地方式**：
- 保存能力挂到 `wps_ppt_manage_slides` 的 `save` / `save_as` 动作（网关已转发 action）；
- 新建能力挂到同工具的 `new_presentation` 动作；
- 页码格式、水印颜色加进 `wps_word_page_layout_and_watermark`（网关**早就**透传
  `pageNumberFormat` / `watermarkColor`，纯粹是 schema 漏字段）；
- Word 只读读回挂到 `wps_word_read_document` 的 `scope` 新枚举值（网关原样转发 scope）。

---

## 1. 逐条结果

| 编号 | 结论 | 改了什么 | 证据 |
|---|---|---|---|
| **ISS-57** | **已修（部分取值受宿主限制）** | schema 补 `pageNumberFormat` + `watermarkColor`；`word.js` 实现页码 | 实测：`Footers.Item(1).PageNumbers.Add(1,true)` 是真 PAGE 域（读回 `codes=["PAGE"]`、`text="1"`），两节文档逐节各得 `1`/`2`；但 `dash`/`page_of_pages` 需要页脚文字与 NUMPAGES 域，宿主**静默丢弃**（见 §2），故这两个取值**显式拒绝 + 中文原因**，不静默降级 |
| **ISS-58** | **修不动（宿主不允许），只改说明 + 如实告警** | 水印改为逐节写入并回传 placement；说明与 warnings 写明正文层只渲染第 1 页 | 实测：`Headers.Item(1).Shapes.AddTextEffect` **静默把形状加到正文层**（header.Shapes.Count 恒 0、doc.Shapes.Count +1）；`Range.ShapeRange.AddTextEffect` 不存在；`InsertXML` 注入页眉 story 无效（页眉 XML 长度不变）；无 `Document.Watermark` API。→ 页眉层水印在 WPS for Mac 上做不到 |
| **ISS-67** | **本轮未复现，已加写入后长度校验** | `wordWriteContent` 重写；写入后逐行读回长度，不符即报错 | 实测：transport 回显 26 字符全对；`Paragraphs.Add(range)+Range.Text`、`Range.InsertAfter` 在**当前宿主**都完整保留小写 `a`（`Xa a ab abc banana A Aa 啊阿` 逐字相等）。原报告环境与当前构建不同，**未能复现**；已加读回校验兜底并在说明写明 |
| **ISS-70** | **已修** | `location:"bookmark"` 改为"定位书签所在段落 → `InsertParagraphAfter` → `InsertAfter`"；书签不存在**报错拒绝**并列出已有书签 | 实测对照：旧写法 `Paragraphs.Add(bmRange)` 把文本写到**文首**、替换掉书签文字并**销毁该书签**；新写法文本落到书签所在段之后、书签保留（`BM_MID[3-6]` 不变）。探针 6/7/9 逐条对照 |
| **ISS-72** | **已修** | 页眉/页脚/水印遍历 `doc.Sections.Count` 全部节；返回 `appliedSections` + `warnings` | 实测：两节文档逐节写页码域各自成功；`appliedSections` 每节带 header/footer/pageNumberFormat/watermark 实际结果 |
| **ISS-76** | **部分修（受网关限制）** | addon 改为两条导出路径依次尝试，把每次尝试的宿主原始报错放进 `hostError` / `attempts`；说明标注可用绕行 | 实测：`slide.Export(path,"PNG",1280,720)` 与 `Export(path,"PNG")` 在宿主都成功落盘。**但**网关 `res?.error \|\| 'PPT 未生成预览'` 仍会覆盖成兜底文案——该文件不在写区 |
| **ISS-79** | **已修** | 新增 `pptRequireSlideIndex`：delete/move/duplicate/set_background/capture 全部前置校验，报中文并给有效范围 | 实测根因：`pres.Slides.Item(99)` **返回 null 而不抛错**，随后 `.Delete()` 才炸出 `Cannot read properties of null`。现在在取 Item 之前就拦下 |
| **ISS-86** | **部分修（受网关限制）** | 回传 `hostError`/`attempts`/`exportedPaths`；schema 加可选 `outputPath`（额外落盘点） | 见 ISS-76；`outputPath` 同时写给调用方自定义路径与桥接临时路径，但桥接仍按自己的临时路径读 base64 |
| **ISS-85** | **改为 `manage_slides` 动作（不能新增工具名）** | `save` / `save_as`：无 filePath 走 `Presentation.Save`，有 filePath 走 `SaveAs`，未命名文稿明确报"没有文件路径" | 实测：`SaveAs` 落盘成功（文件存在）、`Save()` 对已命名文稿读回 `Saved=-1`；未命名文稿的 `Save()` 实测**不抛错**（可能弹模态框），故改为不盲调、直接报错 |
| **ISS-78** | **已修（动作方式）** | `new_presentation` 动作（`Presentations.Add()`，可配 filePath）；追加型动作首行写明不幂等 | 实测：`Presentations.Add()` 成功、`Slides.Add(1,12)` 后 `CustomLayout.Name="空白"`；`add`/`duplicate` 返回体加 `appended:true, idempotent:false` |
| **ISS-55** | **只改说明** | `generate_deck` 说明写清"坐标为 720×405 设计基准，按返回的 pageWidth/pageHeight 换算；960×540 页面实测 ×4/3，其他尺寸未实测" | 依据 `docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/02-ppt.md` §1.6 的实测读数，未新增断言 |
| **ISS-44** | **已修（读回）** | Word：`read_document(scope='layout'/'styles'/'properties'/'bookmarks'/'content_controls'/'fields')`；PPT：`get_slide_shapes` 增加 `layout` 字段（版式名、layoutIndex、占位符清单） | 实测：节几何/页眉页脚文本与域/页码格式/水印形状全部可读回；书签、内容控件、样式(478 个/13 启用)、文档属性、域均可读 |
| **ISS-73** | **已修（优先读回类）** | 同上：样式清单、书签、内容控件、文档属性、域明细都补成了读回（不新增工具名，走 scope） | 同上 |

---

## 2. 本轮新增的宿主事实（探针原始结论，可直接复用）

Word（WPS for Mac 12.0 / 12.1.28496）：

| 事实 | 证据 |
|---|---|
| `Range.Format` **不存在**（`undefined`），段落属性必须走 `Range.ParagraphFormat` | 探针 22：`typeof r.Format === "undefined"`；旧代码 `newPara.Format.Alignment` 一直在静默失败 |
| `Paragraphs.Add(range)` 传书签范围会把内容写到文首、替换书签文字并销毁书签 | 探针 6/9 |
| `Range.InsertParagraphAfter()` + `Range.InsertAfter(text)` 才能在段后可靠新增段落（书签存活、文本完整） | 探针 9（W1/W2） |
| 页脚上 `Range.InsertAfter` / `InsertXML` 被**静默丢弃**；`doc.Fields.Add(range,33)` 传页脚范围**静默返回 null**；唯一可用是 `Footers.Item(1).PageNumbers.Add(1,true)` | 探针 23/24 |
| `Headers.Item(1).Shapes.AddTextEffect` 静默落到**正文层**；`Range.ShapeRange.AddTextEffect` 不存在；页眉 `InsertXML` 无效 | 探针 12/13/14/16/17 |
| `doc.Shapes.AddTextEffect` 落正文层，且只在第 1 页渲染 | 探针 20 + ISS-58 原始证据 |
| `ContentControls` / `BuiltInDocumentProperties` / `CustomDocumentProperties` / `Bookmarks` / `Styles.InUse` 均可安全读回 | 探针 10 |
| 写入路径本身**不再吞小写 `a`**（当前构建） | 探针 1/3 |

PPT（同一宿主）：

| 事实 | 证据 |
|---|---|
| `Slides.Item(越界)` **返回 null 而不抛错** → 后续 `.Delete()` 才报 `Cannot read properties of null` | 探针 24（`outOfRangeItem: "returned:null"`） |
| `slide.Export(path,"PNG",1280,720)` 与 `slide.Export(path,"PNG")` 都可用 | 探针 24 |
| `Presentations.Add()` 可用；`CustomLayout.Name` / `Layout` / `Shapes.Placeholders.Count` 可读 | 探针 24 |
| 未命名文稿 `pres.Save()` 不抛错（可能弹模态框），`pres.Path` 为空 | 探针 24 |

---

## 3. 快照差异（`npx tsx scripts/snapshot-tools.ts /tmp/snap-pptword.json`）

- **工具数 91 → 91**（未新增工具名）；只读 21、写 70 不变。
- **新增工具：无**。原因见 §0：新增工具名必须在网关 `HANDLERS` 注册，而 `src/bridge/gateway.ts`
  及其处理器不在写区；保存/新建改为 `wps_ppt_manage_slides` 的动作。
- 与 `tools-snapshot.p6.json` 的差异（8 项，其中 2 项**不是本轮改动**）：

| 工具 | 变化 | 归属 |
|---|---|---|
| `wps_word_read_document` | schema（scope 枚举 +6 个只读读回值）+ 描述 | 本任务（ISS-44/73） |
| `wps_word_write_content` | 描述 + 子字段描述（参数名未变、required 未变） | 本任务（ISS-67/70） |
| `wps_word_page_layout_and_watermark` | schema 新增 `pageNumberFormat`/`watermarkColor` + 描述 | 本任务（ISS-57/58/72） |
| `wps_ppt_generate_deck` | 仅描述 | 本任务（ISS-55/78） |
| `wps_ppt_manage_slides` | schema 新增 `filePath`/`format`、action 枚举 +4 + 描述 | 本任务（ISS-85/78/79） |
| `wps_ppt_capture_slide_preview` | schema 新增 `outputPath` + 描述 | 本任务（ISS-76/86） |
| `excel_add_chart` / `wps_add_chart` | 仅描述 | ⚠️ **不是本任务**：并行任务在改 `definitions/excel.ts`（图表定位优先级） |

- `required` 数组在所有工具上**均未变化**；无工具被新增/删除/重排。

---

## 4. 验证记录

| 检查 | 结果 |
|---|---|
| `npm run build:wps-addon` | ✅ 通过（word.js 1493 行、ppt.js 1503 行） |
| `node --check wps-addon/addon-core.js` | ✅ 通过 |
| `npm run typecheck` | ⚠️ 有 3 个错误，**全部在 `src/bridge/gateway/excel.ts`**（`Property 'ignoredParams' does not exist on type 'GatewayContext'`），属并行任务的中间态；本轮 4 个文件无错误 |
| `npm test` | ✅ 86/86 通过 |
| `npx tsx scripts/snapshot-tools.ts` | ✅ 91 个工具，差异见 §3 |
| 真实宿主探针 | ✅ 24 个只读/临时文稿探针（`/tmp/probe-*.mjs`，未入库） |

---

## 5. 需要决策 / 交接

1. **新增工具名需要网关写权限**：`wps_ppt_save_presentation` 若要成为独立工具，
   需同时在 `src/bridge/gateway.ts` 的 `HANDLERS` 加一行 + 在 `gateway/ppt.ts` 加处理器。
   目前用 `wps_ppt_manage_slides(action='save'/'save_as'/'new_presentation')` 替代。
2. **`filePath` 未透传**：`gateway/ppt.ts` 的 `manageSlides` 只转发
   `presentationName/action/slideIndex/targetIndex/layoutIndex/backgroundColor`，
   schema 里的 `filePath` / `format` **到不了宿主**。要么网关补两个字段，要么把说明改成"必须用
   `wps_execute_script` 保存到指定路径"。已在描述里如实标注。
3. **预览兜底文案**：`gateway/ppt.ts:154` 的 `res?.error || 'PPT 未生成预览'` 仍会盖掉 addon 的
   `hostError`。addon 侧已尽力回传，但调用方看不到。建议网关改成透传 `hostError`。
4. **跨页水印**：WPS for Mac 无可用页眉层写入路径（§2）。要么接受"只在第 1 页"，
   要么在 Windows/COM 通道实现，要么在文档里写清替代流程。
5. **`pageNumberFormat` 的 `dash` / `page_of_pages`**：`src/bridge/types.ts` 仍声明这两个取值，
   宿主实现不了。可选：网关层直接拒绝，或保持"接受但逐节返回 applied:false"。

---

## 6. 遗留清理（受沙箱限制，需使用者处理）

本轮探针在 `~/Downloads/` 留下临时文件（**沙箱禁止写工作区外的路径，我删不掉**）：

```
probe-iss67-70.docx / probe-wm.docx / probe-wm.pdf / probe-wm-page1.{docx,pdf}
probe-wm-anchorFirst_mark.{docx,pdf} / probe-wm-anchorLast_mark.{docx,pdf}
probe-ppt-save.pptx / probe-ppt-slide.png / probe-ppt-slide2.png
```

用户文档 `测试文字文稿.docx` 全程只读（仅 `word_read_document`），未做任何写入。
探针文稿已在 WPS 里全部关闭（未保存关闭）；`文字文稿7` 是其他任务的文稿，未动。
