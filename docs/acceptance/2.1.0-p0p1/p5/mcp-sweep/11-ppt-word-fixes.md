# 11 PPT/Word 宿主实现修复台账（P5 / mcp-sweep）

- 执行者：PPT/Word 宿主实现修复员（子代理）
- 目标：ISS-44 / 55 / 57 / 58 / 67 / 70 / 72 / 73 / 76 / 78 / 79 / 85 / 86
- 写区（**只有这四个文件**）：
  - `wps-addon/src/ppt.js`
  - `wps-addon/src/word.js`
  - `src/bridge/tools/definitions/ppt.ts`
  - `src/bridge/tools/definitions/word.ts`
- 基线：`tools-snapshot.p6.json`（工具数 91）。本文件边做边更新。
- ⚠️ 运行中的 WPS 加载项是**旧构建**，本文件的所有宿主行为结论以**源码复核 + 隔离探针**为准，
  未实测的一律标"未实测"。

## 触达边界（先说清楚，避免越界承诺）

网关处理器（`src/bridge/gateway/{ppt,word}.ts`）**不在写区**，所以：

- 能改：对外 schema（`definitions/*.ts`）+ 宿主 addon 实现（`wps-addon/src/*.js`）；
- 不能改：网关里 `outputPath = previewPath(...)` 与 `fs.existsSync(outputPath)` 的落盘判定、
  `PPT 未生成预览` 兜底文案、`res?.error || ...` 的取错顺序。
- 结论：凡是"错误信息在网关被兜底文案替换"的问题，本轮只能把 addon 的原始错误**带回调用方可见的响应体字段**
  或如实写进说明；真正的文案修复需网关侧配合（在"需要决策"里列出）。

## 逐条计划（编号 → 文件 → 改什么）

| 编号 | 文件 | 改什么 | 状态 |
|---|---|---|---|
| ISS-57 | `definitions/word.ts` + `wps-addon/src/word.js` | schema 补 `pageNumberFormat`（网关早已透传，纯 schema 缺口）并在 addon 实现页码域写入页脚；`watermarkColor` 一并补 schema + 实现（网关也早已透传） | 待做 |
| ISS-58 | `wps-addon/src/word.js` | 水印从正文层 `doc.Shapes.AddTextEffect` 改到**各节页眉层** `section.Headers.Item(1).Range.ShapeRange.AddTextEffect`，按页面尺寸居中，失败回退正文层并记 warnings | 待做 |
| ISS-67 | `wps-addon/src/word.js`（+ 说明） | 先隔离根因（探针：传输回显 / `Paragraphs.Add(targetRange)` / `Range.Text` / `InsertAfter` / 剪贴板搬移），能修则修，不能修在说明里如实告知并给规避写法 | 待做 |
| ISS-70 | `wps-addon/src/word.js` | `location:"bookmark"` 改为按书签 Range 原位插入（不再 `Paragraphs.Add(targetRange)`），书签不存在时**报错拒绝**而不是静默落到文末；插入后按需恢复书签 | 待做 |
| ISS-72 | `wps-addon/src/word.js` + `definitions/word.ts` | 页眉/页脚/水印遍历**所有节**，返回 `appliedSections` + `warnings`；说明同步 | 待做 |
| ISS-76 | `wps-addon/src/ppt.js` + `definitions/ppt.ts` | 预览导出改为多路径尝试 + **回传宿主原始错误**（`hostError`/`attempts`/`fileExists`），说明如实标注网关兜底文案会覆盖 | 待做 |
| ISS-79 | `wps-addon/src/ppt.js` | 幻灯片页码越界前先校验并抛中文错误（含有效范围），不再抛宿主 `Cannot read properties of null (reading 'Delete')` | 待做 |
| ISS-86 | `definitions/ppt.ts` + `wps-addon/src/ppt.js` | 回传原始返回（见 ISS-76）；schema 支持可选 `outputPath`，说明写明与桥接固定落盘路径的关系 | 待做 |
| ISS-85 | `definitions/ppt.ts` + `wps-addon/src/ppt.js` | 新增 `wps_ppt_save_presentation`（`Presentation.Save` / `SaveAs`，桥接侧已有落盘校验能力时复用；无 filePath 走原地保存） | 待做 |
| ISS-78 | `definitions/ppt.ts` + `wps-addon/src/ppt.js` | 新增「新建」路径（`Presentations.Add`）或在说明写明新建方式；追加型工具说明首行写"不幂等" | 待做 |
| ISS-55 | `definitions/ppt.ts` | `generate_deck` 说明写清"坐标为 720×405 设计基准，工具内部换算到真实页面尺寸"，不断言未验证行为 | 待做 |
| ISS-44 | `wps-addon/src/word.js` + `definitions/word.ts` | Word/PPT 侧补读回：页码格式、水印/节（`read_document` 的 `sections` 分支）、PPT 占位符读回 | 待做 |
| ISS-73 | `definitions/word.ts` + `wps-addon/src/word.js` | 挑宿主已实现且低风险的能力补成工具：样式清单、书签、内容控件、文档属性（**优先读回类**） | 待做 |

## 验证清单（收尾必跑）

- [ ] `npm run build:wps-addon`
- [ ] `node --check wps-addon/addon-core.js`
- [ ] `npm run typecheck`
- [ ] `npm test`
- [ ] `npx tsx scripts/snapshot-tools.ts /tmp/snap-pptword.json` 与 `tools-snapshot.p6.json` 比对，逐条说明新增工具
- [ ] 不提交（git 操作由使用者统一执行）

## 进度日志

- （待填）
