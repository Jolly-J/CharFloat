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

`npm run build:wps-addon`（重建部署入口；`--check` 只校验不写文件）→ `node --check wps-addon/addon-core.js` → `node --import tsx --test tests/addon.test.ts tests/ppt-layout.test.ts`；视觉变化还需真实宿主预览。

## 避坑

改了 `src/**` 但忘记重新构建 → 测试与部署仍跑旧 `addon-core.js`，改动看着"没生效"或"已生效"都是假的 → 生成物必须由构建刷新 → `npm run build:wps-addon --check` 漂移即报错；[ppt-layout.test.ts](../tests/ppt-layout.test.ts) 也断言生成物与源码一致。

PPT 在大页面集中左上角 → 固定 720×405 坐标直接用于所有页面 → 布局未换算真实尺寸 → 按实际宽高映射并检查字号容量 → [ppt-layout.test.ts](../tests/ppt-layout.test.ts) 加真实预览。公式结果为 0 → 用真假判断空值会丢数据 → 区分 null/undefined 与 0 → [addon.test.ts](../tests/addon.test.ts)。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
