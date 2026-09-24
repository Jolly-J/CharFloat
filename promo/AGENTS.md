# 宣传片剧本与录制环境

两条宣传片（片 A「无插件」/ 片 B「有插件」）所需的**剧本、录制环境、道具、施工脚本**。产物全部为 CLI 合成的**虚构数据**。

## 文件地图

- [README.md](README.md)：拍摄台本总览——两条片子的对照、三个作品、数据口径、边界声明。
- [00-开拍前检查单.md](00-开拍前检查单.md)：开拍前 / 中 / 后清单。
- [01-剧本A-无插件.md](01-剧本A-无插件.md)、[02-剧本B-有插件.md](02-剧本B-有插件.md)：分镜表（时间码 / 画面 / 施工步骤 / 台词字幕）。片 B 文末有**卖点→镜头对照表**。
- [03-话术卡.md](03-话术卡.md)：可直接粘贴的提示词，含"只说一句"收敛约束与追问卡。
- [04-救场台词.md](04-救场台词.md)：AI 跑偏、桥接失败、宿主异常的现场处理与降级方案。
- [05-道具清单.md](05-道具清单.md)：每个道具的路径与【真实】/【摆拍】标注。
- [施工/plan.json](施工/plan.json)：**施工单一真源**。三个作品共 33 步，每步含工具名、参数、镜头提示、预计时长。
- [施工/build.py](施工/build.py)：驱动器。`--probe`（只读链路自检）/ `--list` / `--dry-run`（参数契约自检）/ `--gate`（每步等回车）/ `--pace` / `--only`（单步重录）。
- [assets/make-props.py](assets/make-props.py)：道具生成器，幂等。
- [assets/verify-props.py](assets/verify-props.py)：道具验收，输出可核对报告。
- [assets/dump-tools.ts](assets/dump-tools.ts)：从运行时 `catalog` 导出真实工具契约（只读）。
- [assets/tool-schema.json](assets/tool-schema.json)：**生成物**，`dump-tools.ts` 的输出，供 `build.py` 做参数自检。
- [成品基线/](成品基线/)：三个成品的定稿文件 + [make-finals.py](成品基线/make-finals.py)。
- `env-A-无插件/`、`env-B-有插件/`：**生成物**，道具现场。

## 定位与联动

改演示口径时**四处必须同步**，否则片子自相矛盾：

1. `assets/make-props.py` 顶部的常量（7/8/9 月数据、分项、标题）
2. `施工/plan.json` 里的公式与文案（`sheet-03`、`sheet-04`、`sheet-10`）
3. `成品基线/make-finals.py` 的常量
4. 官网 [ComparisonTable.tsx](../website/src/components/ComparisonTable.tsx) 的双轨动画

工具升级后（工具名或参数变化），重新跑 `node --import tsx assets/dump-tools.ts`，再跑 `build.py --all --dry-run` 对齐。

## 修改边界

- `env-*/` 与 `成品基线/*.xlsx|.pptx|.docx` 是**生成物**：改脚本再重跑，不手改产物。
- 本目录**不参与** `npm run build` / `typecheck` / `test`，也不引入运行时依赖；只用 Python 标准库 + `openpyxl` / `python-docx` / `python-pptx`。
- **不往里放真实业务数据、真实客户名或真实凭据**。加密道具的口令 `Demo#2026` 是演示口令。
- 摆拍道具必须在 [05-道具清单.md](05-道具清单.md) 标注，不能混进"真实运行"的叙事。
- 录屏把 `env-B-有插件/月度经营.xlsx` 改脏是预期行为，用 `make-props.py` 复原。

## 验证

```bash
python3 promo/assets/make-props.py            # 幂等复原
python3 promo/assets/verify-props.py          # 道具验收（含加密真伪、成品结构）
python3 promo/施工/build.py --all --dry-run   # 参数契约自检，不连宿主
```

`--dry-run` 只校验工具名与参数名/必填项，**不校验语义**（例如区域地址是否落在表内）。
**静态自检不能替代真机录制验证**：33 步必须在一台装了 WPS 的机器上实跑过，才算这条片子可用。

## 避坑

改 `plan.json` 的参数名后直接开录 → 现场才发现宿主报"未知参数" → 参数名必须对运行时契约，凭记忆写必错 → 先 `build.py --all --dry-run` 再开录 → [tool-schema.json](assets/tool-schema.json) 是唯一口径。

`wps_word_page_layout_and_watermark` 的 `watermarkText` → 临场"加个水印更正式" → 真机实测该参数会让 WPS 主进程 SIGSEGV 崩溃，可能带走未保存内容 → 本套施工完全不用它，任何人临场提出也一律拒绝 → 见 [wps-addon/AGENTS.md](../wps-addon/AGENTS.md)。

WPS 长会话后新建文档失败（`Add()` 返回空、文档数不变）→ 以为脚本写错了 → 宿主进入"拒绝新建文档"状态，重启 WPS 立刻恢复 → **PPT 那一组开拍前必须先重启 WPS** → 已写进 [00-开拍前检查单.md](00-开拍前检查单.md)。

加密道具的明文底稿 → 顺手生成在 `env-A-无插件/03-加密文件/` 里 → "机密文件旁边就放着明文"直接穿帮 → 明文只写临时目录，生成完即删 → [make-props.py](assets/make-props.py) 的 `build_encrypted(out_dir, stage)` 签名就是这个约束。

按经验写"支持全部 Office 技能"这类话术 → 宣传超出实际验证范围 → 项目有 `npm run check:claims` 强制能力口径与运行时真值一致 → 字幕里出现数量必须有对应证据 → [check-capability-claims.ts](../scripts/check-capability-claims.ts)。

## 同步维护

剧本、检查单、剧本里的步骤 id、`plan.json` 的步数或数据口径变化时同步更新本页与 [README.md](README.md)；其余遵循根目录协作规范。
