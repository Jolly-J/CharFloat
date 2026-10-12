  // workbook.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  function saveWorkbook(app, params) {
    const { workbookName } = params || {};
    const wb = getWorkbook(app, workbookName);
    const warnings = [];
    // 从未保存过的新工作簿（Path 为空）调 Save() 会失败或弹"另存为"——如实告知而非报成功（excel-tester L-4）
    let pathBefore = "";
    try { pathBefore = String(wb.Path || ""); } catch (e) {}
    // ⚠️ **必须关掉宿主对话框**：真机踩到——`Save()` 会弹模态框（格式兼容/云同步之类），
    // 没人点它就一直卡着，`Saved` 也不会变；使用者点了"取消"就表现为
    // 「Save() 没抛错，但 Saved 仍是 false、磁盘上没有文件」。
    // 删表（deleteWorksheet）与建簿（createWorkbook）的路径都关了对话框，**保存这条此前漏了**。
    const prevAlerts = (() => { try { return app.DisplayAlerts; } catch (e) { return null; } })();
    try { app.DisplayAlerts = false; } catch (e) {}
    try { wb.Save(); }
    catch (e) {
      return { success: false, workbookName: wb.Name, hostError: e.message, warnings: [`保存失败: ${e.message}`],
        message: `工作簿 [${wb.Name}] 保存失败${pathBefore === "" ? "（该工作簿从未保存过，请先用 save_as 指定路径）" : ""}` };
    }
    finally {
      // 一定要恢复，否则后续所有宿主提示都被静默吞掉
      if (prevAlerts !== null) { try { app.DisplayAlerts = prevAlerts; } catch (e) {} }
      else { try { app.DisplayAlerts = true; } catch (e) {} }
    }
    // 读回是否真的落盘（Saved 标志）
    let savedFlag = null;
    try { savedFlag = Boolean(wb.Saved); } catch (e) {}
    if (savedFlag === false) {
      warnings.push(pathBefore === "" ? "该工作簿从未保存过，Save() 未落盘；请用 save_as 指定路径" : "宿主报告保存后仍为未保存状态，请人工确认");
    }
    return {
      success: savedFlag !== false,
      workbookName: wb.Name,
      fullName: wb.FullName,
      saved: savedFlag,
      warnings,
      message: savedFlag === false ? `工作簿 [${wb.Name}] 未确认落盘` : `工作簿 [${wb.Name}] 已成功保存到磁盘`
    };
  }

  // ── CAP-53 新建工作簿
  //
  // 之前**没有这个工具**（Word 有 word_create_document、PPT 有 new_presentation，Excel 漏了）。
  // 但宿主一直支持：真机实测 `app.Workbooks.Add()` + `wb.SaveAs(路径)` 均可用。
  // 「工具缺失」不等于「宿主不能做」——这是补工具，不是补能力。
  function createWorkbook(app, params) {
    const { savePath, sheetName } = params || {};
    const before = Number(app.Workbooks.Count);
    const wb = app.Workbooks.Add();
    const created = wb && wb.Name ? String(wb.Name) : null;
    // ⚠️ 宿主 **可能拒绝新建**：长会话累积状态下 `Workbooks.Add()` 会**返回 null 且数量不变**
    // （真机复现过：连续两次都失败，重启 WPS 后恢复）。
    // 此时必须**直接说是宿主拒绝**，不要报成"已新建工作簿 [null]"——那文案会被误读成代码 bug。
    if (!wb || !created) {
      return {
        success: false,
        workbookName: null,
        hostRefused: true,
        workbooksBefore: before,
        workbooksAfter: Number(app.Workbooks.Count),
        warnings: ["宿主拒绝新建工作簿：app.Workbooks.Add() 返回空且工作簿数未增加"],
        message: "宿主拒绝了新建工作簿（Workbooks.Add() 返回空）。常见原因：WPS 会话运行过久进入异常状态——" +
                 "**重启 WPS 后重试**；若仍失败请检查是否有未关闭的模态对话框。"
      };
    }

    let saved = null, saveError = null;
    if (savePath) {
      // ⚠️ **必须关掉宿主对话框**：目标路径已存在时 `SaveAs` 会弹"是否覆盖"的**模态框**，
      // 它会阻塞整个 WPS——所有后续调用全部无响应，而且没人点它就一直卡着。
      // （真机踩到：create_workbook 超时，WPS 被模态框阻塞，四组件全部无响应。）
      const prevAlerts = (() => { try { return app.DisplayAlerts; } catch (e) { return null; } })();
      try { app.DisplayAlerts = false; } catch (e) {}
      try {
        wb.SaveAs(String(savePath), 51);
        saved = String(savePath);
        // ⚠️ **不能以"没抛异常"判定保存成功**：DisplayAlerts=false 时，
        // 宿主对**已存在的路径**既不抛异常也不落盘
        // （excel-tester H-4：返回"已保存"但磁盘 mtime 完全没变）。
        // 读回 `wb.Path` 与 `wb.Saved` 才是判据。
        const savedPathBack = (() => { try { return String(wb.Path || ""); } catch (e) { return ""; } })();
        const savedFlag = (() => { try { return Boolean(wb.Saved); } catch (e) { return null; } })();
        if (savedPathBack === "" || savedFlag === false) {
          saveError = `宿主未真正落盘（Path="${savedPathBack}" Saved=${savedFlag}）`;
          saved = null;
        }
      } catch (e) { saveError = String(e.message || e); }
      finally {
        // 无论成功失败都要恢复，否则后续所有操作都静默吞掉提示
        if (prevAlerts !== null) { try { app.DisplayAlerts = prevAlerts; } catch (e) {} }
        else { try { app.DisplayAlerts = true; } catch (e) {} }
      }
    }
    // 首张工作表改名（可选）
    let firstSheet = null;
    try {
      const ws = wb.Worksheets.Item(1);
      if (sheetName) { ws.Name = String(sheetName); }
      firstSheet = String(ws.Name);
    } catch (e) {}

    // 读回核对：工作簿确实存在、页数、首表名、是否已落盘
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const after = Number(app.Workbooks.Count);
    const warnings = [];
    if (after !== before + 1) warnings.push(`新建后工作簿数 ${after}，期望 ${before + 1}`);
    if (savePath && !saved) warnings.push(`保存失败：${saveError}`);

    return {
      success: after === before + 1 && (!savePath || !!saved),
      workbookName: g(() => String(wb.Name), created),
      sheetCount: g(() => Number(wb.Worksheets.Count), null),
      firstSheetName: firstSheet,
      savedPath: saved,
      warnings,
      message: `已新建工作簿 [${g(() => String(wb.Name), created)}]（${g(() => Number(wb.Worksheets.Count), "?")} 张表，首表 ${firstSheet}）${saved ? `并保存到 ${saved}` : ""}`
    };
  }

  function manageWorkbookViews(app, params) {
    const { workbookName, action = "list", viewName } = params || {};
    const wb = getWorkbook(app, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };

    // 富值：本机宿主没有这个对象模型，直接如实拒绝，不要假装
    if (action === "read_rich_values") {
      return {
        success: false, workbookName: wb.Name, action,
        warnings: ["本机 WPS 没有 LinkedDataTypes / HasRichDataType 对象模型，链接数据类型（富值）不可用"],
        message: "本机宿主不支持链接数据类型（富值）：未找到 wb.LinkedDataTypes。可用脚本取单元格**显示文本**作为替代。"
      };
    }

    const listViews = () => {
      const out = [];
      const cv = g(() => wb.CustomViews, null);
      if (!cv) return out;
      const n = Number(g(() => cv.Count, 0));
      for (let i = 1; i <= n; i++) {
        const v = g(() => cv.Item(i), null);
        if (v) out.push({ name: g(() => String(v.Name), null), index: i });
      }
      return out;
    };

    if (action === "list") {
      const views = listViews();
      const slicers = (() => {
        const out = [];
        const sc = g(() => wb.SlicerCaches, null);
        if (!sc) return out;
        for (let i = 1; i <= Number(g(() => sc.Count, 0)); i++) {
          const c = g(() => sc.Item(i), null);
          if (c) out.push({ name: g(() => String(c.Name), null), sourceName: g(() => String(c.SourceName), null), slicerCount: g(() => Number(c.Slicers.Count), null) });
        }
        return out;
      })();
      return { success: true, workbookName: wb.Name, action, viewCount: views.length, views, slicerCacheCount: slicers.length, slicerCaches: slicers,
        warnings: [], message: `工作簿 [${wb.Name}] 有 ${views.length} 个自定义视图、${slicers.length} 个切片器缓存` };
    }

    if (action === "add") {
      if (!viewName) throw new Error("add 需要 viewName（视图名）");
      const cv = g(() => wb.CustomViews, null);
      if (!cv) return { success: false, workbookName: wb.Name, action, warnings: ["本机宿主没有 wb.CustomViews"], message: "本机宿主不支持自定义视图" };
      let created = null;
      try { created = cv.Add(String(viewName)); }
      catch (e) { return { success: false, workbookName: wb.Name, action, hostError: e.message, warnings: [`创建失败：${e.message}（同名视图可能已存在）`], message: `自定义视图 [${viewName}] 创建失败` }; }
      const back = listViews().find(v => v.name === String(viewName)) || null;
      return { success: !!back, workbookName: wb.Name, action, viewName: String(viewName), readBack: back,
        warnings: back ? [] : ["创建后读不回该视图"], views: listViews(),
        message: back ? `已创建自定义视图 [${viewName}]` : `自定义视图 [${viewName}] 创建后读不回` };
    }

    if (action === "show") {
      if (!viewName) throw new Error("show 需要 viewName");
      const v = g(() => wb.CustomViews.Item(String(viewName)), null);
      if (!v) throw new Error(`找不到自定义视图 [${viewName}]。现有：${listViews().map(x => x.name).join(" / ") || "无"}`);
      try { v.Show(); } catch (e) { return { success: false, workbookName: wb.Name, action, hostError: e.message, warnings: [`切换失败：${e.message}`], message: `切换到视图 [${viewName}] 失败` }; }
      return { success: true, workbookName: wb.Name, action, viewName: String(viewName), warnings: [], views: listViews(), message: `已切换到自定义视图 [${viewName}]` };
    }

    if (action === "delete") {
      if (!viewName) throw new Error("delete 需要 viewName");
      const v = g(() => wb.CustomViews.Item(String(viewName)), null);
      if (!v) return { success: true, workbookName: wb.Name, action, warning: undefined, views: listViews(), warnings: [], message: `自定义视图 [${viewName}] 不存在，无需删除` };
      try { v.Delete(); } catch (e) { return { success: false, workbookName: wb.Name, action, hostError: e.message, warnings: [`删除失败：${e.message}`], message: `删除视图 [${viewName}] 失败` }; }
      const still = listViews().some(x => x.name === String(viewName));
      return { success: !still, workbookName: wb.Name, action, viewName: String(viewName), views: listViews(),
        warnings: still ? ["删除后仍能读到该视图"] : [], message: still ? `视图 [${viewName}] 删除后仍在` : `已删除自定义视图 [${viewName}]` };
    }

    throw new Error(`未知的视图操作 action: ${action}（支持 list / add / show / delete / read_rich_values）`);
  }

  function manageDocumentProperties(app, params) {
    const { workbookName, action = "read", properties } = params || {};
    const wb = getWorkbook(app, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const BUILTIN = ["Title", "Subject", "Author", "Keywords", "Comments", "Category", "Company", "Manager"];

    const readAll = () => {
      const builtin = {};
      for (const k of BUILTIN) {
        builtin[k] = g(() => {
          const p = wb.BuiltinDocumentProperties.Item(k);
          const v = p && p.Value !== undefined ? p.Value : p;
          return v === undefined || v === null ? "" : String(v);
        }, "");
      }
      const custom = {};
      const n = g(() => Number(wb.CustomDocumentProperties.Count), 0);
      for (let i = 1; i <= n; i++) {
        const p = (() => { try { return wb.CustomDocumentProperties.Item(i); } catch (e) { return null; } })();
        if (!p) continue;
        custom[g(() => String(p.Name), "prop" + i)] = g(() => String(p.Value), "");
      }
      return { builtin, custom };
    };

    if (action === "read") {
      const cur = readAll();
      return { success: true, workbookName: wb.Name, builtin: cur.builtin, custom: cur.custom, warnings: [],
        message: `工作簿 [${wb.Name}] 属性：标题「${cur.builtin.Title || "（空）"}」作者「${cur.builtin.Author || "（空）"}」，自定义 ${Object.keys(cur.custom).length} 项` };
    }

    // action = delete：按名字删除**自定义**属性（内置属性不能删，只能清空值）
    if (action === "delete") {
      const names = Array.isArray(params?.propertyNames) ? params.propertyNames.filter(Boolean).map(String) : [];
      if (!names.length) throw new Error("delete 需要 propertyNames（要删除的自定义属性名数组）");
      const removed = [], notFound = [];
      for (const want of names) {
        let hit = false;
        for (let i = g(() => Number(wb.CustomDocumentProperties.Count), 0); i >= 1; i--) {
          const pr = (() => { try { return wb.CustomDocumentProperties.Item(i); } catch (e) { return null; } })();
          if (pr && g(() => String(pr.Name), "") === want) { pr.Delete(); removed.push(want); hit = true; break; }
        }
        if (!hit) notFound.push(want);
      }
      const left = readAll();
      return {
        success: true, workbookName: wb.Name, action: "delete",
        removed, notFound,
        remainingCustom: Object.keys(left.custom).length,
        warnings: [],
        message: `已删除自定义属性 ${removed.length} 项${notFound.length ? `，${notFound.length} 项不存在` : ""}；剩余自定义属性 ${Object.keys(left.custom).length} 项`
      };
    }

    // action = apply
    const warnings = [];
    const applied = {};
    if (properties && typeof properties === "object") {
      for (const [k, v] of Object.entries(properties)) {
        if (BUILTIN.indexOf(k) >= 0) {
          try { wb.BuiltinDocumentProperties.Item(k).Value = String(v); applied[k] = String(v); }
          catch (e) { warnings.push(`设置内置属性 ${k} 失败: ${e.message}`); }
        } else {
          // 自定义属性：存在则改，不存在则加（type 4 = msoPropertyTypeString）
          let done = false;
          for (let i = 1; i <= g(() => Number(wb.CustomDocumentProperties.Count), 0); i++) {
            const p = (() => { try { return wb.CustomDocumentProperties.Item(i); } catch (e) { return null; } })();
            if (p && g(() => String(p.Name), "") === k) { try { p.Value = String(v); done = true; applied[k] = String(v); } catch (e) { warnings.push(`改自定义属性 ${k} 失败: ${e.message}`); } break; }
          }
          if (!done) {
            try { wb.CustomDocumentProperties.Add(k, false, 4, String(v)); applied[k] = String(v); }
            catch (e) { warnings.push(`新增自定义属性 ${k} 失败: ${e.message}`); }
          }
        }
      }
    }
    const after = readAll();
    const mismatched = Object.keys(applied).filter(k => {
      const got = BUILTIN.indexOf(k) >= 0 ? after.builtin[k] : after.custom[k];
      return String(got) !== String(applied[k]);
    });
    for (const k of mismatched) warnings.push(`读回与写入不一致: ${k}`);
    return { success: true, workbookName: wb.Name, action: "apply", applied, after, warnings,
      message: `已更新 ${Object.keys(applied).length} 项文档属性${mismatched.length ? `（${mismatched.length} 项读回不一致）` : ""}` };
  }
