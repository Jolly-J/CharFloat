# 官网

## 文件地图

- [package.json](package.json)：官网独立依赖及构建命令。
- [src/App.tsx](src/App.tsx)：页面编排。
- [src/index.css](src/index.css)：公共样式。
- [src/components/QuickStartModal.tsx](src/components/QuickStartModal.tsx)：接入引导。
- [src/components/ArchitectureFlow.tsx](src/components/ArchitectureFlow.tsx)：架构展示。
- [src/components/ComparisonTable.tsx](src/components/ComparisonTable.tsx)：能力对比。
- [src/components/OfficeSlides.tsx](src/components/OfficeSlides.tsx)：演示展示。
- [src/components/OfficeSpreadsheet.tsx](src/components/OfficeSpreadsheet.tsx)：表格展示。

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
