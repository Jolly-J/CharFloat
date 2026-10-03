# Office.js 加载项

## 文件地图

- [src/](src/)：**源码**（P3.3 逐类拆分）。`state.js` / `ui.js` / `lifecycle.js`（共享运行时与每工作簿自动启动）/ `connection.js` / `rpc.js`（含 `dispatchExcelTool` 全部分支）/ `bootstrap.js`，以及 `src/excel/` 下按域细分的操作实现。
- [public/taskpane.js](public/taskpane.js)：**构建生成物**（部署入口），由 `npm run build:office-addon` 按固定顺序拼接 `src/**`。顶部有"请勿手改"声明。
- [public/taskpane.html](public/taskpane.html)：任务窗格入口。
- [public/taskpane.css](public/taskpane.css)：Office / WPS 共用的侧栏样式源，构建时同步为 WPS 自包含的 panel.css。
- [excel/manifest.xml](excel/manifest.xml)：Excel 加载项清单，长生命周期共享运行时与侧栏使用同一入口。

## 定位与联动

先从 `src/rpc.js` 定位 method，再到 `src/excel/` 对应域文件查实现。**改代码改 `src/**`，不要改 `public/taskpane.js`**。参数转换在 [normalizer](../src/bridge/office/normalizer.ts)，连接和回退在 [adapter](../src/bridge/office/adapter.ts)。

## 修改边界

Office.js 与 WPS/COM 对象模型不同。读取先 load/sync，写入完成后 sync 再返回结果；不能把代理对象直接返回。API 是否支持按宿主验证。不要把已完成写入后的错误当成可安全整批重试。

## 验证

`npm run build:office-addon`（重建部署入口；`--check` 只校验不写文件）→ `node --check office-addon/public/taskpane.js` → `node --import tsx --test tests/office.test.ts`；加载、授权与 sync 效果需真实 Excel 验证。

## 避坑

图标 PNG 有效但显示默认占位图 → 对图标与脚本一起返回 no-store 会触发 Office 替换图标 → 微软要求图像允许缓存 → 仅 HTML/JS/CSS 禁止缓存，图片不加缓存禁用指令；更改图标 URL 刷新旧缓存 → [HTTP 回归](../tests/service.test.ts)、[微软说明](https://learn.microsoft.com/en-us/javascript/api/manifest/image?view=word-js-preview)。

侧栏图标正常但工具栏仍显示占位图，且 `/assets/icon-32.png` 实际返回 401 → 仅核对磁盘源码或重新构建不能证明后台已更新 → 独立后台会继续运行旧构建，侧栏与工具栏还使用不同的静态路径 → 核对 `/health` 的 PID 与进程启动时间，重启字浮后台后验证清单全部图标 URL 返回正确 PNG 且未禁止缓存；原生功能区仍未刷新时保存文件后重开 Excel → [静态路由实现](../src/bridge/ws-server.ts)、[HTTP 回归](../tests/service.test.ts)。

未打开侧栏就调用 Office.js → 宿主没有启动加载项运行时，CLI 无法收到连接 → 普通任务窗格的运行生命周期依附页面 → 配置长生命周期共享运行时，首次激活后为当前工作簿启用自动启动；不支持的版本明确提示保持侧栏 → [生命周期测试](../tests/addon-lifecycle.test.ts)、[范围与官方依据](../docs/addon-background.md)。

照搬 WPS 脚本 → Office.js 使用 context/Excel 且需同步 → 使用当前执行器约定 → 核对 taskpane.js 的 execute_script 与 normalizer，并实机读回。

构建产物**会做 esbuild 压缩**（`minify: true`，`target: es2018`）→ 去注释、缩短局部变量名，避免宿主适配实现随加载项明文外发；`taskpane.js` 从 152 KB 降到 82 KB。**字符串字面量不受影响**（产品名、能力数量等运行时文案照常）。
⚠️ **压缩后必须真机验证**：任务窗格跑在 Office 的 webview 里，压缩改变了变量名与函数提升结构，**构建通过看不出问题**，要在真实 Office 里确认加载项能连上。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
