# 官网

## 文件地图

- [package.json](package.json)：官网独立依赖及构建命令。
- [src/App.tsx](src/App.tsx)：页面编排。
- [src/index.css](src/index.css)：公共样式。
- [src/components/Navbar.tsx](src/components/Navbar.tsx)：悬浮胶囊顶栏导航与下载引导。
- [src/components/Hero.tsx](src/components/Hero.tsx)：首屏标题、下载下拉胶囊与产品视觉海报展台。
- [src/components/DeliverySection.tsx](src/components/DeliverySection.tsx)：文档交付核心章节（垂直轮播Tab + 大预览）。
- [src/components/ComparisonTable.tsx](src/components/ComparisonTable.tsx)：“同一句话，两种结果”双轨流程模拟（文件流转 / 当前工作区修改）；视口内单次播放、暂停、重播与减少动态效果适配。
- [src/components/ComparisonDemo.css](src/components/ComparisonDemo.css)：双轨模拟的局部样式及响应式布局，类名使用 comparison- 前缀。
- [src/components/UseCasesMarquee.tsx](src/components/UseCasesMarquee.tsx)：实用指令双排横向跑马灯卡片流。
- [src/components/AgentGrid.tsx](src/components/AgentGrid.tsx)：生态模型支持网格。
- [src/components/DownloadModal.tsx](src/components/DownloadModal.tsx)：全平台客户端下载与系统引导。
- [src/components/QuickStartModal.tsx](src/components/QuickStartModal.tsx)：MCP 接入引导（Cursor / Claude / VS Code）。
- [src/components/BrandIcons.tsx](src/components/BrandIcons.tsx)：标准 Apple 与 Windows 官方矢量图标。
- [src/components/OfficeSpreadsheet.tsx](src/components/OfficeSpreadsheet.tsx)：表格联动视窗。

## 定位与联动

先从 App.tsx 找对应组件。官网与桌面 renderer 是独立应用，不因页面相似改错目录。

## 修改边界

能力文案依据实现和验证证据，演示动画不能冒充真实运行状态。安装命令、平台与打包格式和根 package.json 保持一致。发布是独立操作，不因本地构建成功自动部署。

## 验证

根目录执行 `npm run build:website`，交互和响应式需真实浏览器核验。

## 避坑

将演示效果当作平台全量支持 → 用户预期与实际验证不符 → 对照 [验证文档](../docs/validation.md) 与运行能力 → 检查文案和实际入口。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
