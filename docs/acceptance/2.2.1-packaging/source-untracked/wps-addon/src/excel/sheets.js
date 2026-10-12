  // sheets.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // 3. 新建或激活工作表 (显式锁定 workbookName)
  function createWorksheet(app, params) {
    // 表名校验（excel-tester M-4）：非法字符或超长名字宿主会**静默创建 SheetN** 并返回成功，
    // 调用方以为表叫自己给的名字，后续按名定位全部失败。这里先校验再创建。
    if (params && params.sheetName !== undefined && params.sheetName !== null && String(params.sheetName) !== "") {
      const want = String(params.sheetName);
      if (want.length > 31) throw new Error(`工作表名过长（${want.length} 字符，上限 31）："${want}"`);
      if (/[\\\/\?\*\[\]:]/.test(want)) throw new Error(`工作表名含非法字符（\\ / ? * [ ] :）："${want}"`);
      if (/^'.*'$/.test(want)) throw new Error(`工作表名不能以单引号开头或结尾："${want}"`);
    }
    const { sheetName, workbookName } = params;
    if (!sheetName) throw new Error("缺少 sheetName 参数");

    const wb = getWorkbook(app, workbookName);

    let targetSheet = null;
    try {
      targetSheet = wb.Worksheets.Item(sheetName);
    } catch (e) {}

    let isNew = false;
    if (!targetSheet) {
      const count = wb.Worksheets.Count;
      targetSheet = wb.Worksheets.Add(null, wb.Worksheets.Item(count));
      targetSheet.Name = sheetName;
      isNew = true;
    }

    try {
      wb.Activate();
      targetSheet.Activate();
    } catch (e) {}

    return {
      success: true,
      workbookName: wb.Name,
      sheetName: targetSheet.Name,
      isNew: isNew,
      sheetCount: wb.Worksheets.Count
    };
  }

  // 4. 删除指定工作表
  function deleteWorksheet(app, params) {
    const { sheetName, workbookName } = params;
    if (!sheetName) throw new Error("缺少 sheetName 参数");

    const wb = getWorkbook(app, workbookName);
    let targetSheet = null;
    try {
      targetSheet = wb.Worksheets.Item(sheetName);
    } catch (e) {
      return { success: true, workbookName: wb.Name, message: `工作表 ${sheetName} 已不存在，无需删除` };
    }

    if (wb.Worksheets.Count <= 1) {
      throw new Error(`工作簿 [${wb.Name}] 仅剩 1 个工作表，不能删除最后一个工作表`);
    }

    try { app.DisplayAlerts = false; } catch (e) {}
    try {
      targetSheet.Delete();
    } finally {
      try { app.DisplayAlerts = true; } catch (e) {}
    }

    return {
      success: true,
      workbookName: wb.Name,
      message: `已成功删除工作簿 [${wb.Name}] 中的工作表 [${sheetName}]`,
      sheetCount: wb.Worksheets.Count
    };
  }

  // 6. 获取大纲
  function getSheetOutline(app, params) {
    const sheetName = typeof params === "string" ? params : params?.sheetName;
    const workbookName = typeof params === "object" ? params?.workbookName : null;
    const sheet = getWorksheet(app, sheetName, workbookName);
    const usedRange = sheet.UsedRange;

    if (!usedRange || usedRange.Rows.Count === 0 || usedRange.Columns.Count === 0) {
      return {
        sheetName: sheet.Name,
        isEmpty: true,
        usedRangeAddress: "",
        rowCount: 0,
        columnCount: 0,
        headerPreview: []
      };
    }

    const address = usedRange.Address();
    const rowCount = usedRange.Rows.Count;
    const colCount = usedRange.Columns.Count;

    const sampleRows = Math.min(rowCount, 3);
    const previewRange = sheet.Range(sheet.Cells.Item(usedRange.Row, usedRange.Column), sheet.Cells.Item(usedRange.Row + sampleRows - 1, usedRange.Column + colCount - 1));
    const rawValues = previewRange.Value2;

    let headerPreview = [];
    if (Array.isArray(rawValues)) {
      headerPreview = rawValues;
    } else if (rawValues !== undefined && rawValues !== null) {
      headerPreview = [[rawValues]];
    }

    return {
      sheetName: sheet.Name,
      isEmpty: false,
      usedRangeAddress: address,
      startRow: usedRange.Row,
      startColumn: usedRange.Column,
      rowCount: rowCount,
      columnCount: colCount,
      headerPreview: headerPreview,
      // ISS-94 / ISS-64：把"能写但读不回来"的工作表级状态一并带回，
      // 让 AI 写完保护/标签色/冻结/筛选/条件格式后能自检，而不是只能相信 success。
      sheetState: readSheetState(app, sheet)
    };
  }

  /**
   * 读回条件格式的"是否启用"。
   *
   * 真机实测：本机 WPS 的 `FormatCondition` **没有 `Enabled` 属性**（读到 `undefined`）。
   * 直接 `Boolean(rule.Enabled)` 会把它写成 `false`，让读回看起来像"规则被禁用了"——
   * **读回说谎比没有读回更糟**。属性不存在时如实返回 null（未知），不猜测。
   */
  function readConditionEnabled(rule) {
    try {
      const raw = rule.Enabled;
      if (raw === undefined || raw === null) return null;
      return Boolean(raw);
    } catch (e) {
      return null;
    }
  }

  /**
   * 工作表级读回（只读，逐项 try/catch）。
   *
   * 这些状态原本都只有"写"没有"读"（问题台账 ISS-94）：调用方 set 完拿不到任何证据。
   * 冻结窗格依赖 `app.ActiveWindow`，**只在目标表处于活动状态时才有意义**，
   * 因此非活动表返回 null 并说明原因，避免拿别人的窗口状态冒充本表状态。
   */
  function readSheetState(app, sheet) {
    const state = {};
    const safe = (fn, fallback) => { try { const v = fn(); return v === undefined ? fallback : v; } catch (e) { return fallback; } };

    state.protection = safe(() => ({
      protectContents: Boolean(sheet.ProtectContents),
      protectDrawingObjects: Boolean(sheet.ProtectDrawingObjects),
      protectionMode: Boolean(sheet.ProtectionMode)
    }), null);

    state.tabColor = safe(() => {
      const color = sheet.Tab && sheet.Tab.Color !== undefined ? Number(sheet.Tab.Color) : null;
      if (color === null || !Number.isFinite(color)) return null;
      return excelColorToHex(color);
    }, null);

    state.autoFilter = safe(() => ({
      filterMode: Boolean(sheet.AutoFilterMode),
      range: sheet.AutoFilter && sheet.AutoFilter.Range ? sheet.AutoFilter.Range.Address() : null
    }), null);

    state.freezePanes = safe(() => {
      const active = app.ActiveWindow;
      if (!active) return { available: false, reason: "宿主没有活动窗口" };
      let isTarget = false;
      try { isTarget = String(active.ActiveSheet && active.ActiveSheet.Name) === String(sheet.Name); } catch (e) { isTarget = false; }
      if (!isTarget) {
        return { available: false, reason: "冻结窗格属窗口状态，只有目标表处于活动状态时才能读；请先激活该表" };
      }
      return {
        available: true,
        freezePanes: Boolean(active.FreezePanes),
        splitRow: safe(() => Number(active.SplitRow), null),
        splitColumn: safe(() => Number(active.SplitColumn), null)
      };
    }, { available: false, reason: "读取活动窗口状态失败" });

    // 条件格式：按已用区域扫描，返回每个区域的规则摘要（类型/优先级/是否启用）
    state.conditionalFormats = safe(() => {
      const rules = [];
      const used = sheet.UsedRange;
      if (!used) return rules;
      const fc = used.FormatConditions;
      const count = fc && typeof fc.Count === "number" ? fc.Count : 0;
      for (let i = 1; i <= Math.min(count, 50); i++) {
        try {
          const rule = fc.Item(i);
          rules.push({
            index: i,
            type: safe(() => Number(rule.Type), null),
            enabled: readConditionEnabled(rule),
            priority: safe(() => Number(rule.Priority), null),
            formula1: safe(() => (rule.Formula1 === undefined ? null : String(rule.Formula1)), null),
            interiorColor: safe(() => {
              const c = rule.Interior && rule.Interior.Color !== undefined ? Number(rule.Interior.Color) : null;
              return c === null || !Number.isFinite(c) ? null : excelColorToHex(c);
            }, null)
          });
        } catch (e) { /* 单条规则读失败不影响整体 */ }
      }
      return { count, scanned: Math.min(count, 50), rules };
    }, null);

    return state;
  }

  // 10.6 工作表复制与模板克隆
  function duplicateSheet(app, params) {
    const { sourceSheetName, newSheetName, position = "after", workbookName } = params || {};
    if (!sourceSheetName || !newSheetName) {
      throw new Error("缺少必要参数: sourceSheetName (源工作表) 和 newSheetName (新工作表名称)");
    }
    const wb = getWorkbook(app, workbookName);
    const srcSheet = getWorksheet(app, sourceSheetName, workbookName);

    try {
      const exist = wb.Worksheets.Item(newSheetName);
      if (exist) throw new Error(`工作表 [${newSheetName}] 已存在，请指定不同的新工作表名称`);
    } catch (e) {
      if (e.message && e.message.includes("已存在")) throw e;
    }

    if (position === "before") {
      srcSheet.Copy(srcSheet);
    } else {
      srcSheet.Copy(null, srcSheet);
    }

    const newSheet = wb.ActiveSheet;
    newSheet.Name = newSheetName;

    // M-3：`position` 原先被**完全忽略**（只落在源表之后），但工具照样报成功。
    // 支持 end / before / after（before|after 需配 targetSheetName），并按请求重排。
    const pos = String(position || "after").toLowerCase();
    let positionApplied = "after";
    const warnings = [];
    try {
      if (pos === "end") {
        newSheet.Move(undefined, wb.Worksheets.Item(wb.Worksheets.Count));
        positionApplied = "end";
      } else if (pos === "before" || pos === "after") {
        const targetName = (params && (params.targetSheetName || params.relativeTo)) || sourceSheetName;
        const anchor = wb.Worksheets.Item(String(targetName));
        if (pos === "before") newSheet.Move(anchor);
        else newSheet.Move(undefined, anchor);
        positionApplied = `${pos}:${targetName}`;
      } else {
        warnings.push(`不认识的 position: ${position}（可用 end / before / after），已按 after 处理`);
      }
    } catch (e) {
      warnings.push(`按 position=${position} 重排失败: ${e.message}（新表已创建，位置可能不是请求的位置）`);
    }
    // 读回真实位置核对
    let actualIndex = null;
    try { actualIndex = Number(newSheet.Index); } catch (e) {}
    const totalSheets = wb.Worksheets.Count;
    if (pos === "end" && actualIndex !== null && actualIndex !== totalSheets) {
      warnings.push(`position='end' 未生效：新表在第 ${actualIndex} 位，共 ${totalSheets} 张`);
    }

    return {
      success: warnings.length === 0,
      workbookName: wb.Name,
      sourceSheetName: srcSheet.Name,
      newSheetName: newSheet.Name,
      positionRequested: pos,
      positionApplied,
      actualIndex,
      totalSheets,
      warnings,
      message: `已将工作表 [${sourceSheetName}] 克隆为 [${newSheetName}]（位置 ${positionApplied}，第 ${actualIndex} 位 / 共 ${totalSheets} 张）`
    };
  }

  // 13.9 工作表视图：网格线 / 行列标题 / 缩放
  //
  // 为什么必须是个独立能力：用矢量形状搭画布页时**必须先隐藏网格线**，
  // 否则形状浮在格线上、观感很乱。此前没有任何工具入口，
  // 只能让 AI 退回 wps_execute_script 手写（真机踩到过）。
  function setSheetView(app, params) {
    const { sheetName, workbookName, showGridlines, showHeadings, zoom } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}
    const win = app.ActiveWindow;
    const before = {
      showGridlines: (() => { try { return Boolean(win.DisplayGridlines); } catch (e) { return null; } })(),
      showHeadings: (() => { try { return Boolean(win.DisplayHeadings); } catch (e) { return null; } })(),
      zoom: (() => { try { return Number(win.Zoom); } catch (e) { return null; } })()
    };
    const warnings = [];
    if (showGridlines !== undefined) {
      try { win.DisplayGridlines = showGridlines ? true : false; } catch (e) { warnings.push(`设置网格线失败: ${e.message}`); }
    }
    if (showHeadings !== undefined) {
      try { win.DisplayHeadings = showHeadings ? true : false; } catch (e) { warnings.push(`设置行列标题失败: ${e.message}`); }
    }
    if (Number.isFinite(Number(zoom))) {
      try { win.Zoom = Number(zoom); } catch (e) { warnings.push(`设置缩放失败: ${e.message}`); }
    }
    // 读回核对：写没写进去必须能验证
    const after = {
      showGridlines: (() => { try { return Boolean(win.DisplayGridlines); } catch (e) { return null; } })(),
      showHeadings: (() => { try { return Boolean(win.DisplayHeadings); } catch (e) { return null; } })(),
      zoom: (() => { try { return Number(win.Zoom); } catch (e) { return null; } })()
    };
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      before, after, warnings,
      message: `工作表 [${sheet.Name}] 视图：网格线 ${after.showGridlines ? "显示" : "隐藏"}、行列标题 ${after.showHeadings ? "显示" : "隐藏"}、缩放 ${after.zoom}%`
    };
  }

  // 19. 工作表综合管理 (第四梯队)
  function manageSheet(app, params) {
    const { sheetName, workbookName, action, newName, targetIndex, color, password } = params || {};
    if (!sheetName) throw new Error("缺少必要参数: sheetName");
    if (!action) throw new Error("缺少必要参数: action (rename, move, tab_color, protect, unprotect, read)");

    const wb = getWorkbook(app, workbookName);
    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-34 工作表状态读回：保护状态与标签色此前只能写不能读，
    // 于是"这张表是不是被保护了""标签什么颜色"只能靠人看。
    if (action === "read") {
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      const prot = g(() => Boolean(sheet.ProtectContents), null);
      const colorRaw = g(() => Number(sheet.Tab.Color), null);
      // Tab.Color 是 BGR 整数；-4142/16777215 之类表示"无颜色"
      const hasColor = colorRaw !== null && colorRaw >= 0 && colorRaw !== 16777215;
      const toHex = (v) => "#" + [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff].map(n => n.toString(16).padStart(2, "0")).join("").toUpperCase();
      return {
        success: true,
        workbookName: wb.Name,
        sheetName: sheet.Name,
        index: g(() => Number(sheet.Index), null),
        visible: g(() => Number(sheet.Visible), null),
        protection: {
          protectContents: prot,
          protectDrawingObjects: g(() => Boolean(sheet.ProtectDrawingObjects), null),
          protectionMode: g(() => Boolean(sheet.ProtectionMode), null)
        },
        tabColor: hasColor ? toHex(colorRaw) : null,
        tabColorRaw: colorRaw,
        warnings: [],
        message: `工作表 [${sheet.Name}]：${prot ? "已保护" : "未保护"}，标签色 ${hasColor ? toHex(colorRaw) : "未设置（默认）"}`
      };
    }

    switch (action) {
      case "rename": {
        if (!newName) throw new Error("rename 操作必须提供 newName");
        const oldName = sheet.Name;
        // 目标名已存在时宿主会静默失败或改名成别的——先查存在性（excel-tester M-1）
        const exists = (() => {
          try {
            for (let i = 1; i <= Number(wb.Worksheets.Count); i++) {
              if (String(wb.Worksheets.Item(i).Name) === String(newName)) return true;
            }
          } catch (e) {}
          return false;
        })();
        if (exists) throw new Error(`工作表名 "${newName}" 已被占用，改名未执行（原表仍是 "${oldName}"）`);
        sheet.Name = newName;
        // 读回核对：宿主改名可能静默不生效（excel-tester M-1）
        const actual = (() => { try { return String(sheet.Name); } catch (e) { return null; } })();
        if (actual !== String(newName)) {
          return { success: false, workbookName: wb.Name, oldName, newName, actualName: actual, warnings: [`改名未生效：请求 "${newName}"，宿主读回 "${actual}"`], message: `工作表改名未生效` };
        }
        return { success: true, workbookName: wb.Name, oldName, newName, actualName: actual, warnings: [], message: `工作表 [${oldName}] 已重命名为 [${newName}]` };
      }
      case "activate": {
        // shape 相关操作**依赖工作表处于激活状态**（分组代码里已有 sheet.activate()，
        // 真机实测非活动表会报「当前对象不允许此操作」）。此前没有对外暴露的动作，
        // 使用者想先激活再操作却做不到。
        const wasActive = (() => { try { return wb.ActiveSheet.Name === sheet.Name; } catch (e) { return null; } })();
        sheet.Activate();
        const nowActive = (() => { try { return String(wb.ActiveSheet.Name); } catch (e) { return null; } })();
        const ok = nowActive === String(sheet.Name);
        return {
          success: ok, workbookName: wb.Name, action: "activate",
          sheetName: String(sheet.Name), wasActive, activeSheet: nowActive,
          warnings: ok ? [] : [`激活未生效：请求 ${sheet.Name}，宿主当前活动表为 ${nowActive}`],
          message: ok ? `已激活工作表 [${sheet.Name}]` : `工作表 [${sheet.Name}] 激活未生效`
        };
      }
      case "move": {
        if (!targetIndex) throw new Error("move 操作必须提供 targetIndex (1-indexed)");
        const total = wb.Worksheets.Count;
        const validIndex = Math.max(1, Math.min(Number(targetIndex), total));
        const refSheet = wb.Worksheets.Item(validIndex);
        sheet.Move(refSheet);
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, targetIndex: validIndex, message: `工作表 [${sheet.Name}] 已成功移动至位置 ${validIndex}` };
      }
      case "tab_color": {
        if (!color) throw new Error("tab_color 操作必须提供十六进制颜色 (例如 '#0F172A')");
        const bgr = hexToExcelColor(color);
        if (bgr !== null) {
          sheet.Tab.Color = bgr;
        }
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, tabColor: color, message: `工作表 [${sheet.Name}] 标签底色已设置为 ${color}` };
      }
      case "protect": {
        sheet.Protect(password || undefined);
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, protected: true, message: `工作表 [${sheet.Name}] 已锁定保护` };
      }
      case "unprotect": {
        // ⚠️ 密码错误时宿主**不抛异常也不解锁**，原实现照样返回 protected:false —— 谎报成功
        // （excel-tester M-2）。改为读回 ProtectContents 核对。
        try { sheet.Unprotect(password === undefined ? undefined : String(password)); }
        catch (e) { return { success: false, workbookName: wb.Name, sheetName: sheet.Name, hostError: e.message, warnings: [`解除保护失败: ${e.message}`], message: `工作表 [${sheet.Name}] 解除保护失败` }; }
        let stillProtected = null;
        try { stillProtected = Boolean(sheet.ProtectContents); } catch (e) {}
        if (stillProtected === true) {
          return { success: false, workbookName: wb.Name, sheetName: sheet.Name, protected: true,
            warnings: [password === undefined ? "工作表仍处于保护状态：可能设有密码，请传入 password" : "工作表仍处于保护状态：密码可能不正确"],
            message: `工作表 [${sheet.Name}] 仍处于保护状态（未解除）` };
        }
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, protected: false, warnings: [], message: `工作表 [${sheet.Name}] 已解除锁定保护` };
      }
      default:
        throw new Error(`未知的 Sheet 操作: ${action} (支持 rename, move, tab_color, protect, unprotect)`);
    }
  }
