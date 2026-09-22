# Windows 原生 Office

## 文件地图

- [runner.ps1](runner.ps1)：JSON 输入与 action 分发。
- [common.ps1](common.ps1)：宿主和目标公共操作。
- [excel.ps1](excel.ps1)：Excel COM 结构化操作。

## 定位与联动

入口调用来自 [ms-office-driver.ts](../../src/bridge/office/ms-office-driver.ts)；先定位 action/method，再查脚本分支。

## 修改边界

参数通过 JSON stdin，避免拼接 shell。文档精确匹配且不唯一时停止。区分附着用户实例和测试自行创建实例，不退出用户 Office。stdout 保持协议 JSON，诊断不能污染响应。

## 验证

Windows 下检查 [native-contracts.ps1](../../tests/native-contracts.ps1)；实机要求见 [windows-office.ps1](../../tests/windows-office.ps1) 和 [验证文档](../../docs/validation.md)。先读脚本再运行。

## 避坑

模拟通过就承诺 Windows 兼容 → COM 行为依赖实际桌面宿主 → 在真实 Office 中验证 → 报告区分语法、模拟和实机结果。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
