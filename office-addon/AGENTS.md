# Office.js 加载项

## 文件地图

- [src/](src/)：**源码**（P3.3 逐类拆分）。`state.js` / `ui.js` / `lifecycle.js` / `connection.js` / `rpc.js`（含 `dispatchExcelTool` 全部分支）/ `bootstrap.js`，以及 `src/excel/` 下按域细分的操作实现。
- [public/taskpane.js](public/taskpane.js)：**构建生成物**（部署入口），由 `npm run build:office-addon` 按固定顺序拼接 `src/**`。顶部有"请勿手改"声明。
- [public/taskpane.html](public/taskpane.html)：任务窗格入口。
- [public/taskpane.css](public/taskpane.css)：任务窗格样式。
- [excel/manifest.xml](excel/manifest.xml)：Excel 加载项清单。

## 定位与联动

先从 `src/rpc.js` 定位 method，再到 `src/excel/` 对应域文件查实现。**改代码改 `src/**`，不要改 `public/taskpane.js`**。参数转换在 [normalizer](../src/bridge/office/normalizer.ts)，连接和回退在 [adapter](../src/bridge/office/adapter.ts)。

## 修改边界

Office.js 与 WPS/COM 对象模型不同。读取先 load/sync，写入完成后 sync 再返回结果；不能把代理对象直接返回。API 是否支持按宿主验证。不要把已完成写入后的错误当成可安全整批重试。

## 验证

`npm run build:office-addon`（重建部署入口；`--check` 只校验不写文件）→ `node --check office-addon/public/taskpane.js` → `node --import tsx --test tests/office.test.ts`；加载、授权与 sync 效果需真实 Excel 验证。

## 避坑

照搬 WPS 脚本 → Office.js 使用 context/Excel 且需同步 → 使用当前执行器约定 → 核对 taskpane.js 的 execute_script 与 normalizer，并实机读回。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
