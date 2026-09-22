# macOS 加载项部署权限引导（体验修复）

## 问题（用户实测报错）
```
WPS  : EPERM: operation not permitted, open '.../com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml'
Office: cp: .../com.microsoft.Excel/Data/Documents/wef/wps-bridge-manifest.xml: Operation not permitted
```

## 根因
- WPS 侧：`AddonInstaller.install()` 的 catch 直接 `return { success:false, message: e.message }`，**裸 EPERM 原样进 UI**，无任何解释。
- Office 侧：只有一句请去系统设置开启的文案，且全仓无 `shell.openExternal`，**应用没有打开设置的能力**；提示停在 7.5 秒自动消失的 toast 里。
- 系统层面：目标路径是**别的 App 的沙盒容器**，受 TCC「App Data」保护；**FDA 是 TCC 里唯一既无申请 API、也无系统弹窗的类别**，Apple 要求手动勾选，因此不存在自动弹权限框的做法。

## 修复（方案 A，用户选定；不加手动导入兜底）
- `src/main/permission-detect.ts`：**纯识别逻辑**（无 electron 依赖，可单元测试）——识别 EPERM/EACCES、提取目标路径、非权限错误返回 null。
- `src/main/permissions.ts`：FDA 设置深链 + `appToAuthorize()`（dev → Electron.app；打包 → Office Agent Bridge.app）。
- `addon-installer.ts`：WPS 与 Office 两条路径的失败都返回结构化 `permissionIssue`，不再抛裸错误。
- `main/index.ts`：新增 IPC `permission:full-disk-access-info` / `permission:open-full-disk-access`；安装结果在此补 `appToAuthorize`/`isDev`（保持 installer 不依赖 electron）。
- `renderer/components/PermissionGuide.tsx`：**不自动消失**的引导弹窗——打开设置面板按钮、显示**当前运行模式**该授权的 App 路径（可复制）、提示授权后重启、重试按钮、可展开原始错误。

## 验证
- `tests/permission-guide.test.ts`（5 项）：真实两条报错都能识别并提取路径；EACCES/带 code 同样识别；**无关失败不得误判**；非 macOS 不产生引导；提取不到路径时留空不编造。全部通过。
- typecheck 0、check:agents 0、check:claims 0、test **63/63**、build 0、build:website 0、契约快照逐字节不变。
- 包内核验：重新打包后 `app.asar` 内命中 `Privacy_AllFiles` 与「完全磁盘访问权限」各 2 处，确认修复已进包。

## 未验证（不得视为已通过）
- **界面交互未实测**：没有真的点开弹窗、点按钮跳设置、走完授权流程（需要你在真实桌面操作）。
- 授权成功后重试安装是否真的写盘成功，取决于系统是否已放行，未实测。

## 已知限制
- 深链 `x-apple.systempreferences:...Privacy_AllFiles` 属未公开 scheme，若某系统版本失效，界面仍有完整文字步骤可照做。
- TCC 按代码签名记忆授权；重新构建/换签名后可能需要重新授权。
- 授权后通常需重启应用才生效（界面已提示）。

## 追加：方案 B 浮窗改造（用户实测「路子可行但界面不美观、交互不好」后）

保留：设置深链跳转已验证可用。改造：
- 新增 `src/main/permission-window.ts`：**置顶独立 BrowserWindow**（`alwaysOnTop: 'floating'`、无边框、透明），复用同一渲染包按 `location.hash === '#permission-guide'` 渲染，**不新增构建入口**。
- 新增 `src/renderer/components/PermissionFloat.tsx` + 一整套 `.pg-*` 样式：卡片式分步、应用图标**可拖拽**（`startDrag` 把当前该授权的 .app 交给设置面板）、路径退到折叠的「图标拖不动？查看位置」、技术详情默认折叠。
- 删除旧的应用内弹窗组件：**弹窗会被系统设置窗口盖住，用户无法一边看设置一边拖图标**，这正是必须用置顶浮窗的原因。
- 主界面不再重复弹窗：改为底部常驻细条（可点「打开引导」重新唤出浮窗），消除了此前弹窗 + 自动消失 toast同时出现的重复。
- 修掉旧弹窗的文字裁切（路径与按钮被右边缘截断）与纯工程术语文案。

验证：typecheck 0、test 63/63、check:agents 0、check:claims 0、build 0、契约快照逐字节不变；包内命中 `permission:open-guide`、浮窗文案与 `.pg-drag-card` 样式。

**仍未实测**：浮窗外观、拖拽落点、按钮跳转在本轮未由人操作验证（深链跳转在上一版已被用户验证可行）。
