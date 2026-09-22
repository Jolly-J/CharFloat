# 14 · 验收证据与文档缺口处理计划（ISS-06 / ISS-10 / ISS-16 / ISS-65）

开工时间：2026-09-22 · 执行者：验收证据与文档缺口处理员（子代理）
写区：仅 `docs/**`。不碰源码、`skills/**`、`tests/**`、`scripts/**`；不跑 `npm test` / `npm run build` / `npm run dist`；不做 git 操作。
ISS-06 比对脚本放 `.scratch/evidence/`，不放仓库根目录。

## 任务清单

| # | 任务 | 状态 | 产出 |
|---|---|---|---|
| 0 | 写本计划文件（防中断丢成果） | ✅ 完成 | 本文件 |
| 1 | 核对 `issues.md` 原始证据：ISS-06 / ISS-10 / ISS-16 / ISS-65 | ✅ 完成 | 本文件 §核对 |
| 2 | ISS-06-A：从 `baseline-prior.patch` 提取改造前文件，重跑 acorn 顶层语句多重集比对（目标 120/121） | ⏳ 进行中 | `evidence/iss-06-acorn-multiset.txt` |
| 3 | ISS-06-B：重跑 diff（目标：0 行删除/修改、仅新增 62 行） | ⏳ 待做 | `evidence/iss-06-diff-stat.txt` |
| 4 | ISS-06-C：重跑桩测试 20 条 RPC 序列快照比对 | ⏳ 待做 | `evidence/iss-06-rpc-sequence.txt` |
| 5 | ISS-06-D：`issues.md` ISS-06 条目引用改为指向证据文件 | ⏳ 待做 | `issues.md` |
| 6 | ISS-10：9 条未确认项逐条定状态，产出 `p5/open-questions.md` | ⏳ 待做 | `p5/open-questions.md` |
| 7 | ISS-16：补"实测场景 + 规避建议"到 `issues.md`（证据取自 `mcp-sweep/`） | ⏳ 待做 | `issues.md` |
| 8 | ISS-65：补台账记录（相对偏移语义 + 复现 + 规避） | ⏳ 待做 | `issues.md` |
| 9 | 自检：所有结论有可复核文件；链接可达；无越界改动 | ⏳ 待做 | 本文件 §自检 |

## 核对（任务 1）

`docs/acceptance/2.1.0-p0p1/issues.md`：

- **ISS-06**（第 36 行摘要 / 第 247–255 行正文）：声称 P3.2/P3.3 有 "acorn 顶层语句多重集 120/121 逐字节相同"、"diff 证明 0 行删除/修改、仅新增 62 行"、"桩测试 20 条 RPC 序列快照逐行一致"，但 `docs/acceptance/2.1.0-p0p1/evidence/` 下**无对应文件**（实测该目录只有 6 个 p1/p2 期前缀还原文件 + README）。结论：问题成立。
- **ISS-10**（第 40 行摘要 / 第 284–291 行正文）：指向 `tool-inventory.md` 第 180 行起的未确认项清单。结论：问题成立，需逐条定状态。
- **ISS-16**（第 46 行摘要 / 第 377–391 行正文）：正文已有现象与建议修法，**缺"实测场景 + 规避建议"**。结论：需补。
- **ISS-65**（第 95 行摘要）：**只有摘要表一行，无正文小节**。结论：需新建小节。

## 待续记录

（每完成一项在本节追加一行：时间 / 任务 / 结论 / 证据文件）
