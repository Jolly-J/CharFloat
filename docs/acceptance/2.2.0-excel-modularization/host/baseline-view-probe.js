function getWorkbook(){return wb;}
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
return manageWorkbookViews(app,{workbookName:wb.Name,action:"add",viewName:"M1H_Baseline_TableView"});
