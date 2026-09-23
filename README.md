# Office Agent Bridge

> **让你正在使用的 AI，真正会用 Office**  
> 看懂当前文件，直接操作当前文档，当场完成你的要求。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-macOS%20%7C%20Windows-blue.svg)](package.json)
[![Protocol](https://img.shields.io/badge/Protocol-MCP%20%2F%20HTTP-green.svg)](package.json)
[![Runtime](https://img.shields.io/badge/Node-%3E%3D22-lightgrey.svg)](package.json)

大多数 AI 助手在 Office 里仍停留在“教你怎么做”或“重新生成一个新文件”。  
**Office Agent Bridge** 为现有 AI 增加一层**原生执行基础设施**：连接你正在运行的桌面 Office / WPS，读懂当前状态，指哪改哪，当场把事情做完。

---

## 核心区别

| 维度 | 传统 AI 办公方案 | Office Agent Bridge |
|---|---|---|
| **操作对象** | 离线上载的静态文件 | **当前桌面打开的运行中文档** |
| **交互方式** | 告诉你怎么做，自己复制粘贴 | **直接在当前文档中当场执行** |
| **产出方式** | 全盘重写重新生成（易丢格式/宏） | **保留已有成果，只修改指定目标** |
| **操作粒度** | 粗粒度整篇替换 | **精确定位到指定单元格、公式、图表或页面** |
| **执行模式** | 盲盒式单向输出 | **Observe → Act 双向实时互动** |

### 为什么更靠谱？
- **比视觉点鼠标（GUI Agent）更精准**：不依赖截图和坐标猜测，直连宿主原生结构化 API，改动精准无漂移。
- **比离线脚本（如 openpyxl）更实时**：操作当前活体进程，修改毫秒级呈现在眼前，无需反复保存重开。
- **指哪改哪，无需重来**：已做好的 80% 格式与逻辑完整保留，只做你要求的增量修改。

---

## 运行架构

```text
你常用的 AI (ChatGPT / Claude / Cursor / 豆包 等)
                   │  (MCP / HTTP 协议)
                   ▼
          Office Agent Bridge (独立后台服务)
                   │  (WebSocket / COM / JXA)
                   ▼
        WPS Office  /  Microsoft Office
```

---

## 当前支持范围

| 宿主 | 支持能力 | 当前状态 |
|---|---|---|
| **macOS WPS** | 表格结构化工具、Word / PPT 基础操作 | 核心服务与测试通过 |
| **Windows WPS** | 加载项安装分支与同构工具协议 | 待实机最终回归 |
| **Windows Microsoft Excel** | 原生 COM 自动化与结构化表格工具 | 已实现，支持原生 Excel |
| **macOS Microsoft Office** | JXA 原生脚本通道 | 基础支持，持续迭代 |

---

## 快速上手

### 环境要求
- Node.js 22 LTS 或更高版本
- 本地安装有 WPS Office 或 Microsoft Office

### 源码启动

```sh
# 1. 安装依赖并构建
npm ci && npm run build

# 2. 启动桌面管理器
npm start
```

1. 打开桌面端，点击**“安装 / 修复加载项”**，重启 WPS 即可完成插件就绪。
2. 在**“AI 接入”**页一键复制或导出 MCP 配置至 Cursor、Claude Desktop 等智能体客户端。

### CLI 命令

```sh
node dist/bridge/cli.cjs --start          # 后台常驻运行
node dist/bridge/cli.cjs --status         # 查看连接状态
node dist/bridge/cli.cjs --doctor         # 运行诊断
node dist/bridge/cli.cjs --repair-addon    # 一键修复加载项
node dist/bridge/cli.cjs --stop           # 停止服务
```

> 无参数直接执行 `cli.cjs` 即作为 **stdio MCP Server** 运行。

---

## 参与开发

本项目采用模块化协作规范，详细开发文档参见 [AGENTS.md](AGENTS.md)：
- **MCP 路由与工具定义**：[src/bridge/AGENTS.md](src/bridge/AGENTS.md)
- **WPS 加载项**：[wps-addon/AGENTS.md](wps-addon/AGENTS.md)
- **Office.js 插件**：[office-addon/AGENTS.md](office-addon/AGENTS.md)
- **配套 Agent 技能**：[skills/office-agent-bridge/SKILL.md](skills/office-agent-bridge/SKILL.md)

---

## 许可证 (License)

本项目采用 [MIT License](LICENSE) 授权开源。
