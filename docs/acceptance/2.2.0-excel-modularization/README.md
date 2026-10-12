# WPS Excel 模块化第一批验收

范围：模块化方案 M0、M1。状态：代码、构建、静态与模拟验证完成；M1-H 本机 macOS WPS 回归验收完成，Windows 待验。Word、PPT、Office 加载项源码未迁移；真实 CLI 验收另修复两处 fs 动态导入兼容问题，Bridge 契约不变。

真实宿主结果、既有问题及产物见 [M1-H 本机验收报告](host/README.md)。代码搬迁和后续 CLI 修复的检查结果分别保留，下表为第一批搬迁收尾时的记录。

## 基线与变更

固定锚点及输入 SHA-256 见 [baseline.json](modularization/baseline.json)。开始时工作区干净，版本为 2.2.0。可恢复原文件的命令为 `git show <baseline.json 中的 commit>:<原路径>`，不要使用会变化的 HEAD。

Excel 原 5,930 行拆成 18 个领域片段，最大为 `chart-write.js` 的 703 行。88 个函数与 15 个初始化声明的正文逐字保持一致，初始化顺序保留 STYLE → VALIDATION → SHAPE → CHANGE_BASELINES；因此构建清单中 `shape.js` 在 `range.js` 前。根共享 helper、连接、RPC 分发、启动时机以及 Word/PPT 均保持原状。

WPS 构建器增加递归源码发现，仍使用显式清单决定拼接顺序；多层新增未登记、缺失及重复片段均拒绝。现有 esbuild 压缩、es2018 目标、部署入口和指纹协议保持不变。

## 验证与证据

| 检查 | 结果 | 证据 |
|---|---|---|
| 拆分前相关回归 | 34/34，退出码 0 | [baseline-tests.log](modularization/baseline-tests.log) |
| M0 构建检查 | 递归校验改造后，原生成物仍一致，退出码 0 | [m0-check.log](modularization/m0-check.log) |
| 递归清单夹具 | 合法多层、未登记、缺失、重复，4/4 | [m0-tests.log](modularization/m0-tests.log) |
| 函数及初始化等价 | 88 函数、15 初始化声明内容与顺序不变；非目标加载项源码不变 | [equivalence.json](modularization/equivalence.json)、[归属及跨片段绑定](modularization/ownership.json) |
| 新增纯函数与 RPC 回归 | 24/24，退出码 0 | [excel-tests.log](modularization/excel-tests.log) |
| 压缩产物前后 RPC 比对 | 21 组响应、写入值及宿主调用记录完全一致 | [rpc-equivalence.json](modularization/rpc-equivalence.json)、[迁移前](modularization/rpc-before.json)、[迁移后](modularization/rpc-after.json) |
| 工具契约 | 142 个工具、23 个只读工具；名字、说明、schema、顺序及非构建诊断能力一致 | [contract-equivalence.json](modularization/contract-equivalence.json)、[迁移前快照](modularization/tools-before.json)、[迁移后快照](modularization/tools-after.json) |
| 类型检查 | `npm run typecheck` 退出码 0 | [typecheck.log](modularization/typecheck.log) |
| 完整构建 | `npm run build` 退出码 0，含主进程、字节码、加载项和生成物品牌检查 | [build.log](modularization/build.log) |
| 干净重建与幂等 | 临时副本删除两类部署入口后 `build:addons` 成功；两次构建逐字节一致；真实构建 CLI 对三类漂移均返回 1，未覆写产物 | [clean-build.json](modularization/clean-build.json)、[clean-build.log](modularization/clean-build.log) |
| 完整回归 | `npm test` 152/152，失败 0，跳过 0，退出码 0 | [full-test-unsandboxed.log](modularization/full-test-unsandboxed.log) |
| 参数转发 | 84 个定义、81 个处理器可解析、3 个既有未解析；未发现丢参 | [params.log](modularization/params.log) |
| 收尾清单与发布保护回归 | 7/7，退出码 0；源码泄露反向夹具改为真实的新子目录路径 | [final-local-tests.log](modularization/final-local-tests.log) |

契约导出中的 `capabilities.build`、`capabilities.connection.build` 含时间戳、磁盘更新时间和产物 SHA，明确从契约比对中排除。其他内容逐字节一致；新产物 SHA 独立记录在 `clean-build.json`，不是要求新旧构建指纹相等。未为通过比对修改快照导出器。

首次 `npm test` 受沙箱本机监听限制（listen EPERM）影响，HTTP/stdio 服务相关用例失败，保留 [原始日志](modularization/full-test.log)。在允许临时本机服务的环境中重跑通过；测试使用隔离目录与端口，不附着真实办公宿主。验证环境见 [environment.json](modularization/environment.json)。

命令、退出码和日志哈希汇总见 [validation.json](modularization/validation.json)。日志仅将个人仓库绝对路径替换为 `<repo>`，保留错误类型与断言原文。技能枚举参考中的源码锚点同步到新路径，并去掉了与既有实现冲突的旧垂直对齐限制；没有改变业务实现。

模拟案例涵盖区域读写、混合值/公式、非法维度、样式与设计语言、验证读写、表视图、行隐藏、图表读改、形状清点、结构化表与透视表清点、排序越界、批注、保存、打印、宿主写入异常、未知方法、选区事件和心跳。这不是每个宿主分支或每种输入的全覆盖。

迁移前后模拟测试进程耗时分别为 220ms / 208ms，包含进程启动与编译，仅用于记录同环境观察，不作为真实宿主性能结论。

## 下个会话复核

从仓库根执行；证据脚本从自身位置解析仓库根，不依赖个人目录：

```bash
node docs/acceptance/2.2.0-excel-modularization/modularization/verify-migration.mjs
node docs/acceptance/2.2.0-excel-modularization/modularization/compare-rpc.mjs
node docs/acceptance/2.2.0-excel-modularization/modularization/verify-clean-build.mjs
node scripts/build-wps-addon.mjs --check
npm run check:agents
```

前两项需要本地可读取固定基线提交；迁移核验仅适用于本次搬迁收尾，不作为将来业务函数永远不得变化的测试门禁。RPC 比对与干净重建会刷新本目录对应证据，仅在临时目录恢复基线/移除产物，不部署或操作真实文件。

## 遗留与放行边界

- M1-H 本机 macOS WPS 项已完成：插件已备份、部署、重载并核对运行指纹，业务、保存关闭重开、预览与 PDF 有真实记录；Windows 待验，既有图表标题问题和视图使用条件见宿主报告。只操作新建测试文档。
- Word、PPT、Bridge 后续批次尚未执行。
- 未提交、推送、打包发布或签名。构建生成物变化不代表用户宿主已运行新版本。
