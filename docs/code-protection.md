# 发布包的代码保护

面向"下一个会话的 AI / 新同事"：**这份文档说明我们怎么让发布包里的核心代码不可读，
以及为什么每一步是这么做的。** 改构建或排查"字节码加载失败"前先读这里。

相关入口：[构建字节码的脚本](../scripts/build-cli-bytecode.mjs) ·
[tsup 配置](../tsup.config.ts) · [包装配](../package.json)（`build.files`）

---

## 一、保护到什么程度（先对齐预期）

**没有任何本地手段能"锁死"代码。** 目标是抬高成本，不是做到不可能。

| 做法 | 攻击者要费多大劲 | 我们现在做了吗 |
|---|---|---|
| 什么都不做 | **复制文件就完事** | — |
| 排除 sourcemap | 拿不到原始 `.ts`，但 JS 仍可读 | ✅ |
| minify + mangle | 格式化一下，几小时 | ✅ |
| **V8 字节码** | 要找对版本的工具、处理反编译残缺，几天到几周 | ✅（`cli.cjs`） |
| 核心逻辑放服务端 | 几乎不可能 | ❌ 会失去离线/DLP 客户 |

**如实记录已知边界**：

- 字节码里**仍然保留**函数名、字符串常量、数值常量 —— V8 运行时需要它们。
  消失的是**注释、局部变量名、源码结构**。
- **Electron 升级后必须重跑 `build:bytecode`**（V8 版本变了）。
- **Windows 包尚未真机验证**字节码可加载。

---

## 二、已经做的两件事

### 1. 打包排除 sourcemap（`build.files` 加 `!**/*.map`）

**为什么**：`.map` 的 `sourcesContent` 字段**内嵌完整原始源码**。
实测改前 `app.asar` 里有 **42 个我们自己的 `.ts` 源文件、41.5 万字符的逐字原文** ——
**任何人解压就能拿到全部源码**。

| | 改前 | 改后 |
|---|---|---|
| `app.asar` | 46.0 MB | **24.0 MB** |
| 包内 `.map` | 3050 个 | **0** |
| 内嵌 `.ts` 源码 | 42 个 | **0** |

**注意**：`tsup.config.ts` 的 `sourcemap: true` **没有关** ——
开发时照常生成 map，只是**不进包**。

### 2. 核心 `cli.cjs` 编译为 V8 字节码

**做法**：`dist/bridge/cli.cjs` 从 1.27 MB 明文 JS → **287 字节加载器**；
真实实现编译成 `dist/bridge/cli.jsc`（字节码）。

**关键设计：就地替换 `cli.cjs` 为加载器。**
这样 `installer-engine.ts` / `index.ts` / `package.json` 里所有指向 `cli.cjs` 的引用
**一个字都不用改**，改动面和风险都最小。

---

## 三、目录约定（改构建时最容易搞错的地方）

```
dist/bridge/cli.cjs        加载器（287 B，无逻辑）  → 进包
dist/bridge/cli.jsc        V8 字节码（真实实现）    → 进包
dist/bridge/cli-full.cjs   minify 后的明文 JS       → **不进包**（build.files 已排除）
```

**为什么要有 `cli-full.cjs`**：本地开发与测试用**系统 node** 启动
（V8 版本和 Electron 不同），字节码加载不了。加载器在 V8 不匹配时回退到它，
本地才跑得起来。**发布包里没有这个文件，所以回退在客户机上不生效、不会泄漏源码。**

---

## 四、为什么必须用 **Electron 的 Node** 编译字节码

**V8 字节码与 V8 版本锁死。** 实测两边版本：

```
本机 node:        v25.9.0  → V8 14.1.146.11
Electron 的 Node: v22.16.0 → V8 13.4.114.21
```

用本机 node 编译出来的字节码，在 Electron 里**加载即失败**。

所以 `scripts/build-cli-bytecode.mjs` 调用
`node_modules/electron/dist/...`（`ELECTRON_RUN_AS_NODE=1`）来编译，
并把编译时的 V8 版本写进加载器；版本不符时报出**可操作**的错误，而不是一句 undefined。

---

## 五、构建流程

```bash
npm run build             # ← 正常构建，已自动包含字节码步骤
```

展开是：

```
build:renderer → build:main → build:bytecode → build:addons
                                  ↑ build:main 会覆盖 cli.cjs，所以字节码必须在其之后
```

单独跑：

```bash
npm run build:main        # tsup 出 dist/bridge/cli.cjs（明文，会被下一步替换）
npm run build:bytecode    # 另存 cli-full.cjs、编出 cli.jsc、把 cli.cjs 换成加载器
```

**改了 `src/bridge/**` 之后**：跑一次 `npm run build` 就够，**不需要手动做别的**。

---

## 六、怎么验证保护生效（每次打包后建议跑一遍）

```bash
# ① 包内不该有 map、不该有原始源码
npx asar list "<app>/Contents/Resources/app.asar" | grep -c '\.map$'          # 期望 0
npx asar extract "<app>/Contents/Resources/app.asar" /tmp/chk
find /tmp/chk -name '*.ts' -not -path '*/node_modules/*' | wc -l              # 期望 0
grep -rl "mcpServers = { ...config.mcpServers" /tmp/chk | wc -l               # 期望 0

# ② 明文包绝不能进包
npx asar list "<app>/Contents/Resources/app.asar" | grep cli-full             # 期望无输出

# ③ 功能仍然正常（生产路径 = 用 Electron 的 Node 启动）
node --import tsx scripts/check-cli-protection.mjs
```

**最有力的一条验证**：把 `dist/bridge/cli-full.cjs` **临时移走**再跑 MCP 握手 ——
仍然正常，才说明真的走的是字节码，而不是悄悄回退了。

---

## 七、避坑（都踩过）

**用本机 node 编译字节码** → 在 Electron 里加载即失败
→ V8 版本锁死，必须用目标运行时编译 → 见上面第四节

**忘了先删旧字节码** → `bytenode` 按输入文件名产出（`cli-full.cjs` → `cli-full.jsc`），
产物缺失时会被上一轮的陈旧文件蒙混过关，**打出来的包跑的是旧代码**
→ 编译前先删目标文件；见脚本里的注释

**只加 `minify` 就以为源码没了** → 注释和函数名确实没了，但**逻辑仍可读**，
且 `cli-full.cjs` 这类明文产物如果**没被排除**，等于白做
→ 明文产物必须同时进 `build.files` 的排除清单

**给 `minify` 加 `mangleProps`** → 会把 `args?.workbookName` 这类**属性名**也改掉，
宿主和桥接靠属性名对接，**直接功能全废**
→ esbuild 默认不重命名属性名，**不要开这个开关**

**改完构建不重新验证 V8 版本** → Electron 升级后产出能构建、但运行时加载失败
→ 加载器里带版本校验与可操作提示；升级 Electron 后**重跑并真机验证一次**

---

## 八、还没做的

| 项 | 现状 | 做的话要注意 |
|---|---|---|
| **加载项两个文件**（`wps-addon/addon-core.js`、`office-addon/public/taskpane.js`） | 仍是明文 + 注释 | 它们是**纯拼接**产物，要额外接 esbuild；WPS 的 JSA 引擎踩过多次坑，**压缩后必须真机验证** |
| **`skills/**` markdown** | 完全可读 | 那是积累的 know-how（图表规范、看板布局）。上云可按需下发，本地只能接受或另想办法 |
| **`resources/certs/localhost.key`** | 随包发布 | 实测服务**只绑 `127.0.0.1`**，跨机不可达，真实风险有限；但"所有安装共用一个私钥"仍不该长期保留。要修得引入纯 JS 证书库按机器生成（Windows 无 OpenSSL 时预置私钥是为兜底） |
