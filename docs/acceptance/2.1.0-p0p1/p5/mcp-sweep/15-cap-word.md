# 15 Word 能力补强：CAP-04（域更新与交叉引用）/ CAP-05（内容控件全类型）

- 执行者：Word 能力补强员（子代理）
- 台账条目：`capability-backlog.md` 的 **CAP-04** / **CAP-05**
- 写区（**只有这三个文件**）：
  - `wps-addon/src/word.js`
  - `src/bridge/tools/definitions/word.ts`
  - `src/bridge/tools/definitions/index.ts`（未改：新工具挂在既有 word 区块尾部，装配顺序无需调整）
- 端点：`http://127.0.0.1:19890/api/v1/tool/call`（`Authorization: Bearer <~/.wps-bridge/token>`）
- 测试文档：`~/Downloads/测试文字文稿.docx`（175 段 / 2 表 / 2 节 / `doc.Fields` 63 / `TablesOfContents` 2 / `ContentControls` 1）

> ⚠️ 运行中的加载项是**部署版（旧构建）**，本轮全部宿主读数来自
> **`wps_execute_script` 隔离探针**：只读探测 + 临时新建文档；
> **新增函数则是把 `wps-addon/src/word.js` 里的真实函数原文抽出来、在脚本沙箱里 shim 掉
> `getWordDocument`/`getWordApp`/`log` 后直接调用**（`/tmp/verify_cap.py` + `/tmp/v_cap04.js`、
> `/tmp/v_cap05.js`），验证的是**将要部署的那份源码**，不是另写一份原型。
> 新工具要 `npm run build:wps-addon`（已跑）+ `npm run setup -- --addon` + **彻底退出并重开 WPS** 才生效；
> 部署与重启由调用方执行，本代理未做（也做不了）。

---

## 进度

| 阶段 | 状态 |
|---|---|
| 1. 建台账（本文件） | ✅ |
| 2. 探测 CAP-04 宿主能力 | ✅ 见 §1 |
| 3. 探测 CAP-05 宿主能力 | ✅ 见 §2 |
| 4. 实现 + 真机验证 `wps_word_update_fields` | ✅ 见 §3 |
| 5. 实现 + 真机验证 `wps_word_manage_content_controls` | ✅ 见 §4 |
| 6. 构建 / 语法 / typecheck / test / 快照 | ✅ 见 §5（1 项失败=待注册网关分支，见 §6） |
| 7. 网关侧注册（**交调用方**） | ⏳ 精确代码见 §6 |
| 8. 部署后端到端复测 | ⏳ 待调用方部署 + 重启 WPS |

---

## 1. CAP-04 宿主探测（真实读数）

**问题复现（=台账说的"交付物一改目录页码就全错"）**：把 `测试文字文稿.docx` 的内容经
`Selection.InsertFile` 导入临时文档（**原文件未改动**），在目录后插入一个整页分页符：

| 步骤 | 读数 |
|---|---|
| 导入后 | 段落 177 / 目录 2 / 域 65 / 内容控件 1 |
| 目录 1 的 14 个 PAGEREF（更新前） | `["3"×6,"4"×3,"5"×5]` |
| 目录后插整页分页符 | `Content.End 442`（此处 4627） |
| 再读 PAGEREF（**未更新**） | `["3"×6,"4"×3,"5"×5]` —— **一模一样，陈旧** |
| `doc.Fields.Update()` | 100 ms，返回值 `0`，域数 65 → 63 |
| 更新后 PAGEREF | `["4"×6,"5"×3,"6"×5]` —— 逐项 +1 ✅ |

其它已实测事实：

| 事实 | 读数 |
|---|---|
| `doc.TablesOfContents.Item(n).Update()` | 可用（18 ms，本机实测） |
| `toc.UpdatePageNumbers()` | 可用（5 ms）；**不收录新增标题**（新增标题后只调它，目录里仍无该标题） |
| `doc.Fields.Update()` 是否收录新标题 | **会**（实测新增"第三章 仅字段更新"后只调它，目录里出现了该标题） |
| `doc.Fields.Update()` 返回值 | **恒为 0**（空文档 0、36 个域 0、63 个域 0）→ 不能当计数 |
| 更新是否弹窗/需交互 | **无**（临时文档多次调用均直接返回；63 个域 100 ms） |
| `doc.Fields` 是否含页眉页脚域 | **不含**（正文 59 + 第 2 节 4 = 63 = `doc.Fields.Count`；页眉 `Fields.Count` 单独为 0） |
| `section.Range.Fields.Update()` | 可用（返回 number） |
| `section.Headers/Footers.Item(k).Range.Fields.Update()` | 可用 |
| `Application.CrossReference` / `Document.CrossReference` | **`undefined`** —— 没有 Word 那套"引用类型 + 引用内容"选择器 |
| `Fields.Add(range, 3, "BM \\h", false)`（REF） | 返回非 null，**插入即有结果** |
| `Fields.Add(range, 37, "BM \\h", false)`（PAGEREF） | 同上；`PAGEREF CAP_ANCHOR` 立即得到 `"1"` |
| 书签不存在时的 REF/PAGEREF 结果 | `"错误！未定义书签。"`（真实文档里本来就有 3 个这样的历史域） |
| **空书签**（起止相同）的 REF | 结果为空字符串；PAGEREF 仍给页码 |
| 书签打在**目录条目**上 | ❗ 目录条目是 HYPERLINK/PAGEREF 域的**结果**，一次 `doc.Fields.Update()` 重建目录就把该书签**销毁**，交叉引用随即变成"错误！未定义书签。"（本轮踩到，已修：锚点搜索跳过目录/域结果内部） |
| 域后继续插入内容 | ❗ 按 `field.Result.End` 直接插会把内容插进**域内部**，下次 `Fields.Update()` 连同后面的域一起清掉（实测 advance=0 时 "（第 1 页）" 与 PAGEREF 域消失）；正确推进量是 `field.Result.End + 1`（advance=1 时模板完整存活、两个域都在；advance=2 直接崩） |
| `Paragraphs.Add(doc.Range(...))` | 仍然不可靠（把内容写到文首并毁掉目录）——沿用既有 `InsertParagraphAfter`+`InsertAfter` 路径 |

---

## 2. CAP-05 宿主探测（真实读数）

`doc.ContentControls` 存在；`doc.ContentControls.Add(Type, Range)` 存在。

| Type | 名称 | 新建结果 | 写值/读回 |
|---|---|---|---|
| 0 | richText | ✅ 创建成功 | `Range.Text` + `Range.Font.*` 可读回 |
| 1 | plainText | ✅ | `Range.Text`；`Title`/`Tag`/`LockContentControl`/`LockContents` 可写可读回 |
| 2 | picture | ✅ 创建（内容只是一个占位符号） | **无内容语义** |
| 3 | comboBox | ✅ | `Range.Text` 可写**任意自由文本**；选项用 `DropdownListEntries` |
| 4 | dropdownList | ✅ | 只能从选项里选；**赋非法值静默无效** |
| 5 | buildingBlockGallery | ✅ 创建（内容只是占位文案） | **无内容语义**（无构建基块通路） |
| 6 | date | ✅ | `Range.Text` + `DateDisplayFormat`（实测 `yyyy年M月d日` 可写可读回）+ `DateDisplayLocale`（2052） |
| 7 | group | ❌ **`Add` 返回 null（不抛错）** | 不可用 |
| 8 | checkBox | ✅ | `Checked`(bool)，`Range.Text` 读回 `☐`/`☒` |
| 9 | repeatingSection | ❌ **`Add` 返回 null（不抛错）** | 不可用 |

其它已实测事实：

| 事实 | 读数 |
|---|---|
| `cc.ListItems` | **`undefined`**（Word 桌面版的 ListItems 在本宿主不存在）→ 选项必须走 `cc.DropdownListEntries` |
| `cc.DropdownListEntries.Add(text, value)` | ✅ 可用；`Item(i).Delete()` 可删 |
| 新建下拉/组合框的选项表 | 自带 **1 条占位条目**（`Text=占位文案`、`Value=""`）；不清理它就会变成"第 1 个选项" |
| `DropdownListEntries.Item(k).Select()` | ✅ 可用，`Range.Text` 随即变成选中项 |
| `cc.Range.Text = "不在选项里的值"`（下拉） | **静默无效**（不报错、值不变） |
| `cc.PlaceholderText.Text = "..."` | **无效**（写了读回是空） |
| `cc.SetPlaceholderText(undefined, undefined, "…")` | ✅ 生效（占位文案真的变了） |
| 目标范围已在另一个内容控件内时 `Add` | 返回 **null**（本宿主不支持嵌套控件） |
| 下拉/复选框的包裹语义 | 在原范围**之前**插入控件，**原文本残留**（实测文档里多出一份 `[选择]`）→ 必须删残留 |
| plainText/date 的包裹语义 | 把原范围**包起来**，文本保留（不产生残留） |
| `ContentControls.Item("标题")` | 返回 **null**（不能按名字取）→ 只能遍历 Title/Tag 匹配 |
| 折叠范围（起止相同）`Add` | ✅ 可创建（plainText 会把占位文案插进文档成为初始内容） |
| 真实文档现状 | `测试文字文稿.docx` 有 1 个内容控件：type=0 richText / Title「OAB 实测控件」/ Tag「oab-test」 |

---

## 3. `wps_word_update_fields`（CAP-04）

**实现**：`wps-addon/src/word.js` 末尾 `wordUpdateFields()` + 一组域/目录快照与比较辅助函数。
**真机验证**（抽取真实函数在沙箱里跑，`/tmp/v_cap04.js`，全部 `ok:true`）：

| 用例 | 读数 |
|---|---|
| `scope:'all'` 内容未变 | `updatedFields 11 / changedFields 0 / tocChanged false`，逐节页眉页脚域更新 `Headers.Item(1)=ok, Footers.Item(1)=ok`，27 ms |
| 目录后插分页符 → `scope:'all'` | `changedFields 6`（按域码配对 1 个 TOC 域 + 目录条目域 5 个），`changedPages` 五条 `1→2`，`addedEntries/removedEntries` 现在正确（条目名不再带页码） |
| `scope:'toc', tocMode:'page_numbers'` | 再插一个分页符后页码 `["3"×5]`，`changedPageCount 5`，2 ms，**正文域未被改动** |
| `scope:'section', sectionIndex:1` | `bodyUpdate.sections[0] = {ok:true, fieldsCount:11}`，2 个 story 更新，目录未动 |
| 越界/非法参数 | `sectionIndex 越界：收到 9，当前文档共 1 个节`；`未知的 scope: nope…`；`未知的 tocMode: nope…`（**全部显式报错，无静默降级**） |
| `insert_cross_reference` + `anchorText` | 书签 `CAP_SEC2` 建在正文 `[413,421]="第二章 交付内容"`；**自动跳过目录里的同名命中并告警**；REF → `"第二章 交付内容"`、PAGEREF → `"3"`，`resolved:true`，`bookmarkStillExistsAfterUpdate:true` |
| 插入后再跑一次 `scope:'all'` | 文本仍是 `详见 第二章 交付内容（第 3 页）` —— **模板与两个域都没被 Fields.Update 清掉**（`advance=Result.End+1` 的效果） |
| 引用不存在的书签 | 抛错并列出可用书签/替代路径，**未插入任何域** |
| 锚点文本找不到 | 抛错；若只是命中了目录，错误里会说明"已跳过 N 处命中 + 原因" |
| 空书签 | REF → `""`、PAGEREF → `"3"`，并给出中文告警 |
| `insertLocation:'after_paragraph'`（段落 1 = 目录内部） | 插入点**自动移到容器之后**并告警，域结果 `"3"`、`resolved:true` |

**读回**：`fields{countBefore,countAfter,updatedFields,bodyFieldsAfter,headerFooterFieldsUpdated,comparedByCode,changedByCode,tocEntryFieldsChanged,changedFields,changes[]}`、
`tablesOfContents{countBefore,countAfter,pageNumbersChanged,comparison[{pageNumbersBefore,pageNumbersAfter,changedPageCount,changedPages[{entry,pageBefore,pageAfter}],addedEntries,removedEntries}]}`、
`bodyUpdate`、`tocUpdate`、`stories[]`、`warnings`、`errors`、`hostLimitations`。

> **`updatedFields` 的口径**：宿主 `Fields.Update()` 不返回计数（恒 0），所以本工具把它定义为
> **"更新后文档里实际存在的域数（正文 + 页眉页脚 story）"**，另给 `comparedByCode/changedByCode`
> 说明"按域码稳定配对的域有几个、其中几个结果变了"，目录条目域的变化单列 `tocEntryFieldsChanged`
> —— 目录重建后 `_Toc` 书签名整批变化，按域码根本配不上（实测 5 个条目域只配到 1 个 TOC 域）。

---

## 4. `wps_word_manage_content_controls`（CAP-05）

**实现**：`wps-addon/src/word.js` 末尾 `wordManageContentControls()` + `wordCc*` 辅助函数。
action：`list` / `add` / `set_value` / `delete`。
**真机验证**（`/tmp/v_cap05.js`，全部 `ok:true`）：

| 用例 | 读数 |
|---|---|
| `list`（空文档） | `count 0` + `hostSupport`（可用 6 类 / 能建无内容语义 2 类 / 不支持 2 类 + 中文原因） |
| `add plainText`（marker `____`） | 创建 `[4,16]`，`Title=编号 Tag=no`，值 `OAB-2026-001`，`tagApplied/titleApplied:true` |
| `add date` | `dateDisplayFormat yyyy年M月d日` 写入并读回，值 `2026年9月22日` |
| `add dropdownList`（`listItems:[{text,value}]`） | 占位条目被清掉，读回 `entries=[{同意,Y},{不同意,N}]`，值 `同意` |
| `add checkBox checked:true` | 读回 `Checked:true`、`Range.Text:"☒"`（`leftoverDeleted:""` 说明残留清理真的执行过） |
| `add dropdownList value:'丙'`（不在选项里） | `valueApplied:false` + 中文告警，**不假装成功** |
| `add comboBox value:'广州'`（自由文本） | `valueApplied:"广州"` |
| `add group` / `add repeatingSection` | **显式报错**："本宿主未实现…请改用 richText/plainText/comboBox/dropdownList/date/checkBox" |
| `add` 目标在已有控件内 | **显式报错**（说明嵌套不支持 + 提示先 list 看范围），控件数不变 |
| `add` marker 找不到 | 报错，未创建任何控件 |
| `set_value` by tag / title / index | 三种定位都成功；`tag=nope` 时报错并给出当前控件数 |
| `set_value` 下拉赋非法值 | `applied.value:null` + 告警，读回仍是原值 |
| `set_value` 复选框 | `Checked:false`，读回 `☐` |
| `delete index` / `delete all:true` | 删除成功并回报剩余数量（`deleteAll` 前后 6 → 0） |
| 非法 action / type | 显式报错 |
| 最终 `list` | 6 个控件全部带 `type/title/tag/value(+entries/checked)`，可完整重建表单语义 |

**最小可用闭环（实测跑通）**：模板里写 `____` → `add` + `location:'marker'` + `markerText:'____'`
把它变成控件 → `set_value` 填值 → `action:'list'` 读回核对。
合同填空位 / 表单场景即由此覆盖。

---

## 5. 回归与快照差异

| 检查 | 结果 |
|---|---|
| `npm run build:wps-addon` | ✅ `addon-core.js` 8032 行（word.js 2489+ 行） |
| `node --check wps-addon/addon-core.js` | ✅ 通过 |
| `npm run typecheck` | ✅ 0 错误 |
| `npm test` | 85 pass / **1 fail** —— 唯一失败项是 `contract-consistency.test.ts` 的「注册工具缺少网关分支必须为空」，报 `wps_word_manage_content_controls`、`wps_word_update_fields`；**这正是硬约束 5 预期的状态**，加 §6 的注册行后即消失（其余 85 项全过，含 addon 契约与边界守卫） |
| `npx tsx scripts/snapshot-tools.ts /tmp/snap-cap-word.json` | ✅ 写出，101 个工具（只读 23 / 写 78） |

**快照差异**（基线 `docs/acceptance/2.1.0-p0p1/tools-snapshot.p7.json`，97 个工具）：

- 工具数 **97 → 101（+4）**。其中 **我新增 2 个**：
  `wps_word_update_fields`（required `[]`，15 个属性，`readOnlyHint:false`，`schemaHash 0969d281ca5f05fa`）、
  `wps_word_manage_content_controls`（required `["action"]`，22 个属性，`readOnlyHint:false`，`schemaHash 7147f715b9f4fb0a`）。
- 另 **+2 个不是我改的**：`excel_format_text_segment`、`wps_format_text_segment`（Excel 侧并行任务）；
  同理 `wps_inspect_api` / `excel_patch_cells` / `excel_add_conditional_formatting` 等既有条目的
  `inputSchema`/`description` 变化也来自并行任务，**与本轮无关**。
- 顺序影响：把新增工具过滤掉后，工具顺序与 p7 **逐项一致**；我的 2 个工具插在 word 区块尾部
  （index 87、88），ppt 区块整体后移 2 位，**既有工具相对顺序未变**。
- 既有 word 工具：`schemaHash`/`descriptionHash` **零变化**。
- `capabilities` 段变化：有（同样来自并行任务的 Excel 侧改动）。
- 未改 `src/bridge/tools/definitions/index.ts`（装配顺序无需调整）。

---

## 6. 需要在写区外补的代码（**交调用方**，三处）

### 6.1 `wps-addon/src/dispatch.js`（第 172–175 行之后、PowerPoint 分发之前）

```js
        case "word_update_fields":
          result = wordUpdateFields(app, params);
          break;
        case "word_manage_content_controls":
          result = wordManageContentControls(app, params);
          break;
```

### 6.2 `src/bridge/gateway/word.ts`（追加两个处理器）

```ts
export const updateFields: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("word_update_fields", {
      documentName: args?.documentName,
      action: args?.action || "update",
      scope: args?.scope || "all",
      sectionIndex: args?.sectionIndex,
      tocMode: args?.tocMode || "full",
      includeHeadersFooters: args?.includeHeadersFooters ?? true,
      bookmarkName: args?.bookmarkName,
      targetBookmark: args?.targetBookmark,
      anchorText: args?.anchorText,
      anchorOccurrence: args?.anchorOccurrence,
      crossReferenceType: args?.crossReferenceType || "both",
      crossReferenceTemplate: args?.crossReferenceTemplate,
      insertLocation: args?.insertLocation || "end",
      paragraphIndex: args?.paragraphIndex,
      prefixText: args?.prefixText
    });
};

export const manageContentControls: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('list'|'add'|'set_value'|'delete')");
    return await callOffice("word_manage_content_controls", {
      documentName: args?.documentName,
      action: args?.action,
      type: args?.type,
      index: args?.index,
      tag: args?.tag,
      title: args?.title,
      value: args?.value,
      checked: args?.checked,
      dateDisplayFormat: args?.dateDisplayFormat,
      dateDisplayLocale: args?.dateDisplayLocale,
      listItems: args?.listItems,
      clearListItems: args?.clearListItems ?? false,
      placeholderText: args?.placeholderText,
      lockContentControl: args?.lockContentControl,
      lockContents: args?.lockContents,
      location: args?.location,
      markerText: args?.markerText,
      markerOccurrence: args?.markerOccurrence,
      paragraphIndex: args?.paragraphIndex,
      initialText: args?.initialText,
      all: args?.all ?? false,
      maxControls: args?.maxControls
    });
};
```

> 两段都刻意用 `args?.X` 逐个读取：`npm run check:params`
> （`scripts/check-param-forwarding.ts`）按"schema 属性 vs 处理器里 `args?.X`"判定静默丢参，
> 漏读任何一个属性都会被它抓出来。

### 6.3 `src/bridge/gateway.ts`（第 106 行 `"wps_word_capture_preview": word.capturePreview,` 之后）

```ts
  "wps_word_update_fields": word.updateFields,
  "wps_word_manage_content_controls": word.manageContentControls,
```

加完这 3 处后：`contract-consistency` 的「注册工具缺少网关分支」恢复为空，
`npm run check:params` 也能通过（新工具的 37 个 schema 属性全部被处理器读取）。

---

## 7. 实测宿主**不支持**的能力（如实列出）

**CAP-04**

1. **没有 `CrossReference` 对象**（`Application.CrossReference` / `Document.CrossReference` 均为
   `undefined`）→ 交叉引用**只能按书签**，无法按 Word 的"编号项/标题/书签/脚注/尾注/图表"引用类型插入。
2. `doc.Fields.Update()` **不返回更新计数**（恒 0）→ 只能自己比对快照。
3. `doc.Fields` **不含页眉/页脚 story 的域** → 必须逐节 `Headers/Footers.Item(k).Range.Fields.Update()`。
4. **书签不能放在目录/域结果内部**：域一重建就被销毁（不是"不支持"，但用错会静默失效，已内置规避）。
5. `doc.Fields.Add(range, 33)`（PAGE 域）传**页脚范围**时静默返回 null ——（上一轮既已实测，本轮沿用既有结论，
   页码仍只能走 `Footers.Item(1).PageNumbers.Add`）。

**CAP-05**

1. **`group`(7)、`repeatingSection`(9) 本宿主未实现**：`ContentControls.Add` 返回 `null`（不抛错）。
2. **`cc.ListItems` 不存在**（`undefined`）→ 只能用 `DropdownListEntries`；`listItems` 参数在桥接侧
   是统一入口，落到宿主是 `DropdownListEntries.Add(text, value)`。
3. **不支持嵌套内容控件**（目标范围在已有控件内时 `Add` 返回 null）。
4. **`picture`(2)、`buildingBlockGallery`(5) 建得出但没内容语义**（无法填图片/构建基块）。
5. `ContentControls.Item("标题")` 按名字取**返回 null**（只能遍历 Title/Tag）。
6. `cc.PlaceholderText.Text` 直接赋值**无效**，只能 `SetPlaceholderText(...)`。
7. 下拉列表赋不在选项里的值是**静默无效**（不报错）——工具改为核对选项并如实告警。

---

## 8. 遗留

- **端到端复测待做**：§6 的注册行 + 部署 + 彻底重启 WPS 之后，需要用真实工具名
  （`wps_word_update_fields` / `wps_word_manage_content_controls`）各跑一次，确认 RPC 通道与返回值形状。
  在此之前，"工具可被 AI 调用"这一环**尚未验证**（本轮验证到的是"宿主实现 + 读回逻辑正确"）。
- `skills/**` 与台账 `capability-backlog.md` 的状态更新不在本轮写区，需由调用方在部署复测后补
  （skill 说明建议补：目录页必须显式更新、交叉引用只能按书签且锚点要落正文、内容控件 6 类可用 + listItems 口径）。
- 本轮真机探针只在临时新建文档与 `~/Downloads/测试文字文稿.docx` 的**导入副本**上做，
  **原文件未被任何写操作改动**（`InsertFile` 导入 + 探针文档 `Close(0)` 不保存）。
