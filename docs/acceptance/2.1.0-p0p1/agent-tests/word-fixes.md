# Word 侧三项：ISS-67（吞字母 a）/ ISS-125（水印崩溃）/ CAP-24（关键词加粗）

- **候选标识**：2.1.0-p0p1
- **负责范围（独占）**：`wps-addon/src/word.js`、`src/bridge/tools/definitions/word.ts`、本文件
- **本轮日期**：2026-09-22（承接 `p5/mcp-sweep` 与 p0p1 台账）
- **本轮未做**：未构建（`build:wps-addon` / `build:main` / `setup`）、未重启服务、未调用任何真机工具、未做崩溃复现、未跑 `snapshot:tools`

## 结论速览

| 项 | 结论 | 本轮改了什么 | 验证层级 |
|---|---|---|---|
| **ISS-67** | **根因已定位，且是假报警**：宿主没有吞字符，是当时的**验证探针自己**用 `.replace(/[\r\a]/g, "")` 把读回文本里所有小写 `a` 删掉了（JS 里 `\a` = 字母 `a`，不是 BEL） | 读回校验从"只比长度"升级为**逐字比对**（等长改写也拦）+ 报错带出写入/读回原文 + 每行返回 `readBackMatches`；**没有动写入路径**（前提不成立，动它是无据的风险） | 根因：原始会话逐字取证（可直接复核）<br>代码：mock 宿主单测 8/8 |
| **ISS-125** | 选择**彻底禁用** `watermarkText`（不做崩溃复现，也不写一个无法在真机上验证的新实现） | 参数在任何宿主写操作**之前**硬拒绝 + 删除 `AddTextEffect` 实现 + `read` 通路不受影响 + schema/说明同步 | mock 宿主单测 9/9（含"拒绝时零宿主调用"）<br>**真机未验证**（Lead 串行） |
| **CAP-24** | 宿主侧早已实现，**只差 schema 暴露**；顺带补了 3 个真实缺陷（无上限查找循环、假成功文案、命中范围异常可能整篇加粗） | `wps_word_format_document` schema 增加 `searchQuery`/`searchQueries`；宿主侧加循环上限、逐处读回校验、异常命中跳过并告警、零命中如实报告 | `check:params` 通过（新参数确已被处理器读取）<br>mock 宿主单测 12/12<br>**真机未验证**（Lead 串行） |

---

## 1. ISS-67：`wps_word_write_content` 吞掉所有小写字母 `a`

### 1.1 结论

**宿主没有吞字符。这个现象是当时的验证探针自己制造的。**

子代理写的 `verify_tool.js` 把读回文本做了 `.replace(/[\r\a]/g, "")` 之后再比较。**在 JavaScript 正则里 `\a` 是恒等转义，等于字母 `a`**（只有 C/PCRE/.NET 里 `\a` 才是 BEL `0x07`；本项目内部一律写 `\x07`）。于是探针**先把读回文本里的每个小写 `a` 删掉**，再与原始串比较，必然得出"所有小写 `a` 消失"的结论——而文档里写进去的内容一直是对的。

这也解释了全部症状：小写 `a` 全消失、大写 `A` 与中文完好（因为删的就是字面 `a`）。

### 1.2 取证来源（原始会话，可直接复核）

现象出自子代理会话 `629ea463-d49c-4624-8441-83c9dbaf4c66`（2026-09-22 15:53）。DSH 会话记录是 zstd 压缩的：

```bash
zstdcat "$HOME/.dsh/sessions/--Users-Python-Office~0020Agent~0020Bridge--/629ea463-d49c-4624-8441-83c9dbaf4c66/session.v3.jsonl.zstd" > /tmp/sess-629.jsonl
grep -c '确认字符' /tmp/sess-629.jsonl   # → 5
```

### 1.3 原始证据（逐字引用，未改写）

第一轮"精确等值匹配"探针（会话里的 `verify_tool.js`）：

```js
const out = {};
const D = doc;
const want = params.want;
const stripped = want.replace(/a/g, "");
const found = { want: [], stripped: [] };
for (let i = 1; i <= D.Paragraphs.Count; i++) {
  const t = String(D.Paragraphs.Item(i).Range.Text).replace(/[\r\a]/g, "");   // ← 探针自己删掉了所有 a
  if (t === want) found.want.push(i);
  if (t === stripped) found.stripped.push(i);
}
out.r = JSON.stringify({ sentByPython: want, exactMatch: found.want, strippedMatch: found.stripped,
  strippedForm: stripped, total: D.Paragraphs.Count });
return out;
```

它的**原始输出**（会话里 tool/result 逐字，未改写）：

```
SENT to tool: '确认字符: a ab abc banana A Aa 啊阿'
tool success: True
{
  "sentByPython": "确认字符: a ab abc banana A Aa 啊阿",
  "exactMatch": [],
  "strippedMatch": [148],
  "strippedForm": "确认字符:  b bc bnn A A 啊阿",
  "total": 149
}
```

读法：`strippedMatch: [148]` 说明文本**确实**写进了第 148 段（共 149 段 → 落在文末，`location:"end"` 语义正确）。`exactMatch: []` 与"`a` 被删"完全由探针那行 `.replace(/[\r\a]/g, "")` 造成。

第二轮"隔离"探针（会话里的 `isolate_a.js`，四个写法**每个**都以同一行 `.replace(/[\r\a]/g, "")` 收尾）的**原始输出**：

```
SENT: 'Xa a ab abc banana A Aa 啊阿'
Paragraphs.Add(targetRange)+Range.Text     -> X  b bc bnn A A 啊阿
Paragraphs.Add()+Range.Text                -> 
Paragraphs.Add(targetRange)+InsertAfter    -> 
ContentEnd Range.InsertAfter               -> X  b bc bnn A A 啊阿
```

读法：
- 第 1、4 行的读回串**正好等于 `S` 去掉全部 `a`**（外加 `\r` 被同一正则删掉）→ 反过来说，**写进去的就是完整的 `S`**，写入正确。
- 第 2、3 行是**空串**（`Paragraphs.Add()` 无参那条在 WPS 上 `Range.Text` 赋值没落到读回范围），并不是报告里写的"✅ `S` 完整保留"。
- 因此报告表格"哪些 API 会吞 a"的结论是**误读**：没有一条能证明宿主吞字符。

### 1.4 已排除的可能（逐条给出方法）

| 假设 | 排除方法 | 结论 |
|---|---|---|
| 我方源码里有删 `a` 的代码 | 全量静态穷举：把 `src/**`、`wps-addon/src/**`、`office-addon/src/**`、`scripts/**` 里 **1161 条正则字面量 + `new RegExp`** 逐条当"逐字符清洗器"实测，看谁能把样本串的每个 `a` 删掉 | **0 条命中**；全仓库也没有裸 `\a` 转义（唯一出现处是本次新增的注释） |
| 构建时把转义改坏（`\x07` → 别的） | 逐行比对 `wps-addon/addon-core.js` 与 `src/**` 里的 `[\r\n\x07]` | 完全一致，无漂移 |
| 历史版本曾经有过 `\a` | `git log --all -S` 全历史搜 `[\r\n\a]` 与 `\a` | **从未出现过** |
| MCP/HTTP 传输吞字符 | 子代理自己的 transport 回显（`params.s` 原样返回）已证明传输完好；本轮复核部署指纹：磁盘上 15 份已部署 `addon-core.js` 与仓库产物**同一 sha256**（`2068b6b5…`） | 排除 |
| 读回工具（`wps_word_read_document` 等）删 `a` | 这些函数统一用 `[\r\n\x07]`，且上面的穷举已覆盖；`\x07` 在 JS 里是 BEL，不含 `a` | 排除 |
| 宿主 `Range.Text`/`InsertAfter` 在特定 API 路径吞 `a` | 原始输出第 1、4 行的读回串与 `S` 只差被探针删掉的 `a` → 写入本身完好 | **不成立**（原报告的核心猜测） |

### 1.5 本轮代码改动（保留并加固兜底）

前提不成立 ≠ 该删兜底：**宿主真的吞改字符**依然值得拦。改动只做"更会发现、更好定位"，不动写入路径（换了路径既无证据、又会引入位置语义风险）：

1. 新增 `wordNormalizeWrittenText()`：写入/读回两侧用**同一个**归一化（`\r\n\v\f\x07`），并写明"控制符必须写 `\x07`、不能写 `\a`"的坑。
2. 校验从"长度不一致"升级为**逐字不一致**——等长改写（如 `a`→`b`）以前能蒙过去，现在会被拦。
3. 报错带出**写入原文与读回原文**（各截断 60 字符）+ 行号，并提示"不要直接重试覆盖"。
4. 每行返回 `insertedParagraphs[].readBackMatches`（`true` / `null`=读回失败未核对），调用方不再只能盲信 `success`。
5. 删掉从未被使用的 `writtenRanges`（读代码时的干扰项）。

顺带修掉的旧实现小毛病：原来的 `try { ... throw ... } catch (e) { if (/…/.test(e.message)) throw e; }` 是"用报错文案做控制流"，现在读回与比较分离，不再有被文案改动削弱的隐患。

### 1.6 验证到什么程度

- ✅ **根因**：原始会话逐字取证（§1.2–1.3），可自行复核；静态穷举 §1.4 逐条给方法。
- ✅ **代码**：mock 宿主单测 8/8 —— 正常宿主写入返回 `readBackMatches:true`；mock 里模拟"吞 `a`"与"等长把 `a` 换成 `b`"两种宿主，**都被新护栏拦下**且报错含原文。
- ❌ **未做**：真机复跑原始复现串（不需要——前提已证伪，且会浪费独占宿主时间）。

### 1.7 台账订正（**已由 Lead 授权并执行完毕，2026-09-22 晚**）

假报警原先被当成"高严重度真缺陷"记在 4 处文档 + 缺一条避坑条目。**以下 5 处已全部改完**（授权范围仅此 5 处，未动其他文件）：

| 文件 | 位置 | 实际改法（已完成） |
|---|---|---|
| `docs/acceptance/2.1.0-p0p1/issues.md` | 状态口径 | 增加 `撤销·非缺陷`（经原始证据复核，原现象不成立）——遵守本文件"不删条目"的规则 |
| 同文件 | ISS-67 状态行（原"严重度**高** / 部分完成"） | 改为划线标题 + **撤销·非缺陷**、严重度「**撤销**（原判高是误判）」、状态「已澄清·非缺陷」、归属「**验证方法出错**（既非宿主实现，也非本仓库代码）」 |
| 同文件 | ISS-60~74 与 ISS-60~88 两张详述表的 ISS-67 行 + #8 澄清格 | 均改为撤稿说明并指回文末；两处"最重"改为"两条 / 四条（ISS-67 已撤稿）" |
| 同文件 | 文末 | 新增 `## 撤稿记录：ISS-67`：原结论、复核结论、**逐字引用的探针代码与原始输出**、排除清单表、现行处置、教训（同类"测试工具出错却指向产品"本轮第 4 次） |
| `docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/03-word.md` | §9 第 4 条 + §10 全节 | §9 划线撤回；§10 改为「【已撤回】…实为验证探针的正则坑」：10.1 当年错误结论（含"空串被读成完整保留"的反向误读）、10.2 复核结论与原始证据、10.3 排除清单、10.4 可复用教训 |
| `docs/acceptance/2.1.0-p0p1/total-acceptance-2026-09-22.md` | 5.3 部分完成表 | ISS-67 移出该表（3 条→2 条），新增 **5.3b 撤销·非缺陷（1 条）**；并在 5.4 后补一条"水印从'不修'升级为'禁用'（ISS-125）"的补充说明（该报告出具于 ISS-125 登记之前） |
| `wps-addon/AGENTS.md` | 避坑段 | 新增两条：**`\a` 陷阱**（触发条件→错误做法→已证实原因→正确做法→证据，按仓库格式）与**水印崩溃/已禁用**条目 |
| `skills/office-agent-bridge/references/native-scripting.md` | 第 284 行 | 改写过期表述："页眉/页脚**逐节写入全部节**（ISS-72 已修）" + "**水印写入已禁用**（ISS-125），需水印手动插入或走 COM" |

**残留（未授权、也没动）**：`issues.md` 第 155–156 行的 ISS-127 / ISS-128 两行表格少一列（5 个 `|`，同表其余行为 6 个）——这是先前提交遗留的格式问题，与 ISS-67 无关，建议由 Lead 顺手补一列。

检查：`npm run check:agents` ✅（12 个 AGENTS.md，路径有效）；`npm run check:claims` ✅（29 个文件，无与元数据源矛盾的数量表述）。

---

## 2. ISS-125：`watermarkText` 导致 WPS SIGSEGV

### 2.1 选择：彻底禁用（不是"修实现"）

两条路里选了禁用，理由：

1. 崩溃代价不对称：宿主崩溃会带走用户未保存的数据，而且 Word/PPT/Excel 一起掉线（code 1001）。
2. 任何"替代实现"都只能靠真机验证，而我被明确禁止触发水印写入路径（也不能做崩溃复现）。**写一个无法验证、且专门用于触发已确认崩溃点的实现，等于把风险留在线上。**
3. 原实现即使不崩也是错的定位：`Headers.Shapes` 的写入被宿主静默落到正文层，只在第 1 页渲染，本来就不是"每页可见"——这个功能本来就没达到它承诺的效果。

### 2.2 改法

1. **拒绝位置**：`watermarkText` 的检查放在 `wordPageLayoutAndWatermark` 里**所有宿主写操作之前**（在 `action:"read"` 分支与任何 `Headers/Footers` 写之前）。传了 `watermarkText` 就**整体拒绝**，页眉/页脚/页码一律不写 → **不留半成品**。
2. **报错内容**：说明崩溃现象与证据、明确"本次调用未修改文档"、给三条替代路径（WPS 内手动「插入 → 水印」/ Windows COM 通道 `Section.Headers.Shapes` / 确认宿主已修复后按流程解禁）。
3. **删除实现**：整个 `AddTextEffect` 实现（`wordAddWatermarkToSection`）已从源码删除，不留任何可被调用的水印写入路径；原位置留证据块（崩溃报告路径、栈帧、解禁条件）。原 `if (watermarkText)` 分支改为"不可达即抛错"的显式断言，避免有人绕过守卫后静默什么都不做。
4. **`watermarkColor`**：水印本体已禁用后它是死参数，传了会得到一条"已忽略"告警（不再静默忽略）。
5. **`read` 通路不受影响**：`action:"read"` 只读页眉页脚与现有水印现状（CAP-35），仍然可用；但若同时传 `watermarkText` 会先报错（自相矛盾的入参）。
6. **schema/说明同步**（`word.ts`）：工具说明、`watermarkText`、`watermarkColor` 三处都改成"已禁用 + 原因 + 替代路径"。

### 2.3 验证到什么程度

- ✅ mock 宿主单测 9/9：`watermarkText` 抛错且**宿主调用记录为 0**（证明拒绝发生在任何写之前）；与 `headerText` 同传时也是整体拒绝、零宿主调用；只传 `headerText` 正常写入且**全程没有 `AddTextEffect`**；只传 `watermarkColor` 不再假成功；`action:"read"` 正常且不碰写接口。
- ✅ `node --check wps-addon/src/word.js` 通过；`npm run typecheck` 通过；`scripts/check-param-forwarding.ts` 通过（schema 与处理器仍一致）。
- ✅ 静态确认：全仓库已无 `wordAddWatermarkToSection` 调用残留，无 `Shapes.AddTextEffect` 残留。
- ❌ **真机未验证**（按分工由 Lead 串行执行；**不要**用真机去试水印，只验"传了必须报错且不崩"）。

### 2.4 解除禁用的条件（写进代码注释）

宿主换版本 → 在**独立测试文档**上单独验证 `AddTextEffect` 不再崩溃 → 才能恢复实现（恢复时须同时确认跨页可见性，别把"只在第 1 页"的实现再放回来）。

---

## 3. CAP-24：Word 关键词加粗（按 `searchQueries`）

### 3.1 现状与缺口

- 宿主实现**早就存在**（`wordFormatDocument` 里的 `Range.Find` + `Font.Bold` 通路），`src/bridge/gateway/word.ts` 也一直在转发 `searchQuery`/`searchQueries`。
- 缺的是**工具 schema**：`wps_word_format_document` 的参数表里没有这两个字段 → AI 永远传不进来。这正是台账写的"宿主已实现，schema 未暴露，零成本"。

### 3.2 改法

**A. schema（`src/bridge/tools/definitions/word.ts`）**
- 新增 `searchQuery`（string）与 `searchQueries`（string[]），并在工具说明里写清关键词通路的语义与返回字段。
- `check:params` 复跑通过 → 新参数确实被处理器读取（不是只改描述）。

**B. 宿主侧加固（`wps-addon/src/word.js`）** —— 暴露前先把三个真实缺陷补掉：

1. **无上限查找循环**：原来是 `while (f.Execute())`，宿主一旦不推进查找位置就是**死循环 → WPS 卡死**。现在与 `wordFindAndReplace` 一致加了 5000 次上限，并返回"已达上限"告警。
2. **假成功文案**：原来无论命中几处都回 `已成功为 N 个关键词匹配项 (0 处) 应用排版`；零命中时调用方看不出来。现在返回 `matched`、`formattedMatches`、`verifiedMatches`、`queryStats`，零命中直接写"未找到这些关键词，未应用任何排版"。
3. **命中范围异常可能误伤全篇**：代码约定"宿主命中后把 Range 收缩到命中文本上"（`f.Parent`）。若某版本不收缩，`hit.Font.Bold = true` 会把**整个范围（极端情况是整篇）**加粗。现在先比对命中范围文本长度与关键词长度（容差 2，容忍段落标记），明显不符就**跳过不赋格式**并计入 `skippedUnexpected` + 告警。
4. **计数去重**：`doc.Content` 已经覆盖表格，再叠加 `Tables.Item(t).Range` 会把同一处查两遍；现在按 `Start:End` 去重，`formattedMatches` 才是"处数"。
5. **逐处读回校验**：请求了 `bold` 时逐处读回 `Font.Bold` 统计 `verifiedMatches`（"报了多少处但一处没生效"是这类工具最容易犯的错）。
6. **空/纯空白关键词**：全部为空则**报错且不改文档**（避免落进 `target:"all"` 通路把整篇格式化，也避免"给每个空格加粗"这种无意义动作）。

### 3.3 验证到什么程度

- ✅ `check:params`：`schema 有、处理器不读` = 0 条（新参数真被读取）。
- ✅ mock 宿主单测 12/12：命中 2 处 → `formattedMatches=2 / verifiedMatches=2 / queryStats` 正确且确实调用了 `Font.Bold`；零命中 → `matched:false` + 如实文案；**命中范围异常 → 跳过赋格式、零次 `Bold` 赋值、给出告警**；空/纯空白 → 报错。
- ✅ `npm run typecheck` 通过；`tests/contract-consistency.test.ts` 12/12；`tests/service.test.ts` 5/5。
- ❌ **真机未验证**：`searchQueries` 在 WPS 里是否真的命中、`f.Parent` 是否确实收缩到命中处、`Font.Bold` 读回值形态（`true` vs `-1`，代码用 `Boolean()` 兼容两者）——都要靠 Lead 的真机步骤确认。

### 3.4 待确认（真机验证时请重点看）

1. 「**待确认**」WPS 里 `Range.Find.Parent` 是否与 Word VBA 一样在命中后被收缩到匹配文本上。若**不**收缩，新护栏会跳过并告警（不会误伤全篇），但功能就等于没生效——这时需要改用按 index 逐个 `doc.Range(start,end)` 定位的写法。
2. 「**待确认**」`bold:true` 时 `Font.Bold` 读回是 `true` 还是 `-1`（`Boolean()` 两者都算通过；若读回 `0`/`null` 则 `verifiedMatches` 会小于 `formattedMatches` 并出现在告警里，属"如实降级"，不是崩溃）。
3. 「**待确认**」表格内命中是否与正文命中重复计数已被去重逻辑消除——真机上请对同一关键词核对 `formattedMatches` 与肉眼处数。

---

## 4. 本轮跑过/没跑过的检查

| 检查 | 结果 |
|---|---|
| `node --check wps-addon/src/word.js` | ✅ 通过 |
| `npm run typecheck`（tsc --noEmit） | ✅ 通过 |
| `npx tsx scripts/check-param-forwarding.ts` | ✅ 通过（82 个工具定义；未发现"schema 有、处理器不读"） |
| `node --import tsx --test tests/addon.test.ts` | ✅ 5/5 |
| `node --import tsx --test tests/contract-consistency.test.ts` | ✅ 12/12 |
| `node --import tsx --test tests/service.test.ts` | ✅ 5/5 |
| `node --import tsx --test tests/ppt-layout.test.ts tests/addon-freshness.test.ts` | ⚠️ 17/18：**唯一失败项是"生成物与源码一致，且不含 import/export"**（`build-wps-addon --check` 报漂移）——改了源码未重建，**预期之内**，未去改测试 |
| mock 宿主护栏单测（`/tmp/word-guard-test.mjs`，21 条断言） | ✅ 21/21（ISS-125 9 条 / ISS-67 等 8 条… 见 §1.6、§2.3、§3.3） |
| `npm run build:wps-addon` / `build:main` / `setup` / 真机调用 / `snapshot:tools` | ❌ **按分工未执行**（独占资源，Lead 串行处理） |

> mock 单测是**逻辑层验证**，不替代真机：它证明"守卫的触发时机、报错内容、统计与跳过逻辑"按设计工作，不证明宿主 API 行为。

---

## 5. 交给 Lead 的后续事项

1. **重建部署入口**：`npm run build:wps-addon` → `node --check wps-addon/addon-core.js` → 再跑 `tests/ppt-layout.test.ts`（生成物一致那条应恢复通过）。
2. **重建桥接**：`src/bridge/tools/definitions/word.ts` 改了 schema → 需要 `npm run build`（MCP 侧才会看到 `searchQueries`）。
3. **契约快照**：`wps_word_write_content` / `wps_word_format_document` / `wps_word_page_layout_and_watermark` 三个工具的描述与参数都变了 → `tools-snapshot` 需要由你按流程重生成（我未跑 `snapshot:tools`，避免动别人的产物）。
4. **真机验证建议顺序**（低风险优先，**不要**去试水印写入的旧行为）：
   - `wps_word_write_content`（含小写 `a` 的中文串）→ 期望 `success` 且 `insertedParagraphs[].readBackMatches === true`；
   - `wps_word_format_document` + `searchQueries:["<文档里真实存在的词>"]` + `bold:true` → 看 `formattedMatches/verifiedMatches/queryStats`，并 native 读回该处 `Font.Bold`；再跑一个不存在的关键词看是否 `matched:false`；
   - `wps_word_page_layout_and_watermark` 传 `watermarkText` → 期望**报错且 WPS 不崩**（确认无新崩溃报告）；再跑 `action:"read"` 确认读回通路仍在；
   - 看 `bridge_diagnose` 的构建指纹，确认跑的是新构建。
5. ~~**台账订正**~~ → **已完成**（2026-09-22 晚，Lead 授权）：见 §1.7 的执行结果表（5 处文档 + 2 条避坑条目 + 1 处 skill 过期表述）。
6. ~~**skill 文档**~~ → **已完成**：`skills/office-agent-bridge/references/native-scripting.md:284` 已按改写文本落地（页眉页脚逐节 + 水印禁用）。
7. **另一件别忘**：`task-2`（word-ppt-tester）的任务描述里仍写着"Word 正文写入是否**丢字符**（曾有'吞掉所有小写字母 a'的记录）"——**那条记录已撤稿**。若要让该子代理继续，建议先把这句换成"用 `[\r\n\x07]` 读回、逐字比对，注意 `\a` 恒等转义坑"，否则很可能再烧一轮去追同一个幽灵。

---

## 6. 未做 / 未验证清单（如实）

- ❌ 真机验证：本轮**一次真机工具都没调**（按硬约束）。
- ❌ 未构建：`wps-addon/addon-core.js` 仍是旧产物 → 现在跑真机**看到的是旧行为**，务必先重建。
- ❌ 未复现崩溃，也未尝试任何"修实现"的水印替代方案。
- ❌ 未跑 `npm test` 全量（改源后"生成物与源码一致"必然失败，且其他队友正在改 `excel.ts` 等文件，全量结果会混淆）；只跑了受影响且不依赖新产物的测试。
- 「**待确认**」§3.4 的三条宿主行为（`Range.Find.Parent` 是否收缩、`Font.Bold` 读回形态、表格命中去重）。
- 说明：ISS-67 的结论与台账当时的记述**相反**，属"推翻既有结论"，证据全部来自原始会话逐字引用（§1.3），可独立复核；台账 5 处已于 §1.7 订正完毕。
