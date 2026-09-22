# 15 Word 能力补强：CAP-04（域更新与交叉引用）/ CAP-05（内容控件全类型）

- 执行者：Word 能力补强员（子代理）
- 台账条目：`capability-backlog.md` 的 **CAP-04** / **CAP-05**
- 写区（**只有这三个文件**）：
  - `wps-addon/src/word.js`
  - `src/bridge/tools/definitions/word.ts`
  - `src/bridge/tools/definitions/index.ts`（仅装配，预计不需要）
- 端点：`http://127.0.0.1:19890/api/v1/tool/call`（`Authorization: Bearer <~/.wps-bridge/token>`）
- 测试文档：`~/Downloads/测试文字文稿.docx`（175 段 / 2 表 / 2 节 / Fields 65 / TOC 2）

> ⚠️ 运行中的加载项是**当前部署版**，本轮所有宿主结论来自
> **`wps_execute_script` 隔离探针**（只读探测 + 临时新建文档），不来自"调用新工具"。
> 新工具必须 `npm run build:wps-addon` + `npm run setup -- --addon` + **彻底退出并重开 WPS** 才生效；
> 部署与重启由调用方执行，本代理不做。

---

## 进度

| 阶段 | 状态 | 说明 |
|---|---|---|
| 1. 建台账（本文件） | ✅ 完成 | |
| 2. 探测 CAP-04 宿主能力（Fields / TOC / REF / PAGEREF） | ✅ 完成 | 见 §1 |
| 3. 探测 CAP-05 宿主能力（ContentControls 全类型 / ListItems） | ✅ 完成 | 见 §2 |
| 4. 实现 `wps_word_update_fields`（word.js + schema） | ✅ 完成 | 见 §3 |
| 5. 实现 `wps_word_manage_content_controls`（word.js + schema） | ✅ 完成 | 见 §4 |
| 6. 快照差异 + 回归（build / typecheck / test / snapshot） | ✅ 完成 | 见 §5 |
| 7. 网关注册行（交调用方加） | ✅ 完成 | 见 §6 |
| 8. 真机读数 | ⚠️ 逻辑经探针验证；新工具端到端待部署后复测 | 见 §7 |

---

## 0. 待填（边做边更新）

（本节保留为过程记录，最终结论见各节。）

---

## 1. CAP-04 宿主探测（真实读数）

（待填）

---

## 2. CAP-05 宿主探测（真实读数）

（待填）

---

## 3. `wps_word_update_fields` 实现

（待填）

---

## 4. `wps_word_manage_content_controls` 实现

（待填）

---

## 5. 回归与快照差异

（待填）

---

## 6. 需要在 `src/bridge/gateway.ts` 补的注册行

（待填）

---

## 7. 真机验证与遗留

（待填）
