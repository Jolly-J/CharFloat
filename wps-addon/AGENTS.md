# WPS 宿主执行

## 文件地图

- [src/](src/)：**源码**（P3.2 逐类拆分）。`shared.js` 配置与公共工具、`connection.js` 连接握手、`ribbon.js` 功能区回调、`dispatch.js` RPC 分发、`excel.js` / `word.js` / `ppt.js` 三类文档操作、`ppt-layout.js` **纯布局计算（无宿主 API，可 Node 直接 import）**、`bootstrap.js` 启动引导。
- [addon-core.js](addon-core.js)：**构建生成物**（部署入口），由 `npm run build:wps-addon` 按固定顺序拼接 `src/**`。顶部有"请勿手改"声明。
- [index.html](index.html)：加载项入口。
- [manifest.xml](manifest.xml)：加载项元信息。
- [ribbon.xml](ribbon.xml)：功能区入口。
- [status.html](status.html)：状态页面。

## 定位与联动

先在 `src/dispatch.js` 定位 RPC method，再读对应类文件里的函数，不需通读全文件。**改代码改 `src/**`，不要改 `addon-core.js`**。入参来源见 [gateway](../src/bridge/gateway.ts)。

## 修改边界

精确定位文档，宿主对象不要直接序列化。修改批次保留部分失败信息，超时先读回。PPT 使用真实 PageSetup 尺寸和 pt；矢量绘图与统计图表分开判断。只修改目标对象，避免重建无关内容。

## 验证

`npm run build:wps-addon`（重建部署入口；**会用 esbuild 压缩**去注释与局部变量名；`--check` 只校验不写文件）→ `node --check wps-addon/addon-core.js` → `node --import tsx --test tests/addon.test.ts tests/ppt-layout.test.ts`；视觉变化还需真实宿主预览。

## 避坑

改了 `src/**` 但忘记重新构建 → 测试与部署仍跑旧 `addon-core.js`，改动看着"没生效"或"已生效"都是假的 → 生成物必须由构建刷新 → `npm run build:wps-addon --check` 漂移即报错；[ppt-layout.test.ts](../tests/ppt-layout.test.ts) 也断言生成物与源码一致。

PPT 在大页面集中左上角 → 固定 720×405 坐标直接用于所有页面 → 布局未换算真实尺寸 → 按实际宽高映射并检查字号容量 → [ppt-layout.test.ts](../tests/ppt-layout.test.ts) 加真实预览。公式结果为 0 → 用真假判断空值会丢数据 → 区分 null/undefined 与 0 → [addon.test.ts](../tests/addon.test.ts)。

写脚本读回 Word/表格文本、想清掉段落标记与单元格标记 → 用 `.replace(/[\r\n\a]/g, "")`——**JS 正则里 `\a` 是恒等转义、等于字母 `a`**（`\a` = BEL `0x07` 是 C/PCRE/.NET 的语义）→ 于是探针**先把读回文本里所有小写 `a` 删掉再比对**，得出"宿主吞掉了所有小写字母 `a`"的假报警：ISS-67 据此被记成"高严重度静默改坏内容"，实际宿主写入一直正确（原始输出 `strippedMatch:[148]`／共 149 段，恰好证明内容写到了文末），差点去改正常的写入路径 → 控制符一律写 `\x07`（本文件已有的读回代码就是这么写的），**看到"宿主行为诡异"先逐行回读探针自己的清洗/比较/断言逻辑** → 证据：[03-word.md §10 撤稿记录](../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/03-word.md)、[issues.md 撤稿记录](../docs/acceptance/2.1.0-p0p1/issues.md)；同类"测试工具出错却指向产品"本轮已 4 次。

`watermarkText` 与 `AddTextEffect` → 以为只是"水印落在正文层、只在第 1 页"这种效果问题 → 真机实测该调用会让 **WPS 主进程 SIGSEGV 崩溃**（崩溃报告已核实：`wpsapi` 栈帧连续重复 6 帧＝无限递归；Word/PPT/Excel 同时掉线，可能带走用户未保存数据）→ **该参数已在源码层硬禁用**：`wordPageLayoutAndWatermark` 在任何宿主写操作之前拒绝、已删除 `AddTextEffect` 实现、`action:"read"` 读回通路保留；解除禁用前必须换宿主版本并在**独立测试文档**上单独验证不再崩 → 证据：[issues.md ISS-125](../docs/acceptance/2.1.0-p0p1/issues.md)、[agent-tests/word-fixes.md §2](../docs/acceptance/2.1.0-p0p1/agent-tests/word-fixes.md)。

`app.Workbooks.Add()` 返回 null、工作簿数不变 → 以为是自己代码的 bug 去改代码 → 真机复现：**WPS 长会话（高频调用累积）后会进入拒绝新建文档的状态**，`Add()` 返回 null 且 `Workbooks.Count` 不变（`DisplayAlerts=true`，不是对话框阻塞）；**重启 WPS 后立即恢复** → 建簿路径必须**校验返回值**并明说「宿主拒绝 + 建议重启 WPS」，不要报成「已新建工作簿 [null]」（那文案会被误读成代码 bug，把排查方向带偏）；**此类问题的排查顺序是「先重启宿主、再怀疑代码」** → 证据见 [issues.md ISS-131](../docs/acceptance/2.1.0-p0p1/issues.md)。

构建产物**会做 esbuild 压缩**（`minify: true`，`target: es2018`）→ 用它去注释、缩短局部变量名，避免宿主适配实现随加载项明文外发；`addon-core.js` 从 12,020 行降到 84 行。**压缩不改动字符串字面量**，所以产品名、能力数量、错误文案等运行时可见文本不受影响。
⚠️ **构建指纹必须留在文件最前面**：桥接只读**前 4096 字节**匹配 `ADDON_BUILD_FINGERPRINT: <hex>`（见 [build-fingerprint.ts](../src/bridge/build-fingerprint.ts)）。压缩会剥掉头部注释，所以构建脚本在压缩后**重新注入**头部，并断言指纹串仍在、且确实以头部开头，否则中止写出。
⚠️ **压缩后必须真机验证**：WPS 的 JSA 引擎与浏览器有差异，压缩会改变变量名与函数提升结构，**构建通过完全看不出问题**，只有加载项在 WPS 里跑起来才知道。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
