  // filter-sort.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // 17. 自动筛选与数据排序 (第三梯队)
  function setFilterAndSort(app, params) {
    const { sheetName, workbookName, range, enableAutoFilter, sortRules, action = "apply" } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-33 筛选状态读回：不传 range 也能读（读的是"这张表当前有没有筛选、覆盖哪一块"）。
    if (action === "read") {
      const active = (() => { try { return Boolean(sheet.AutoFilterMode); } catch (e) { return null; } })();
      const addr = (() => {
        try { return sheet.AutoFilterMode && sheet.AutoFilter && sheet.AutoFilter.Range ? sheet.AutoFilter.Range.Address() : null; } catch (e) { return null; }
      })();
      // 逐个字段的筛选条件（Criteria1/Criteria2/Operator）
      const filters = [];
      try {
        if (active && addr) {
          const f = sheet.AutoFilter;
          const rng = f.Range;
          const cols = rng.Columns.Count;
          for (let i = 1; i <= cols; i++) {
            const flt = f.Filters.Item(i);
            let on = false;
            try { on = Boolean(flt.On); } catch (e) {}
            if (!on) continue;
            const one = (() => { try { return flt.Criteria1 === undefined ? null : flt.Criteria1; } catch (e) { return null; } })();
            const two = (() => { try { return flt.Criteria2 === undefined ? null : flt.Criteria2; } catch (e) { return null; } })();
            const op = (() => { try { return Number(flt.Operator); } catch (e) { return null; } })();
            const header = (() => { try { return String(rng.Cells(1, i).Value2); } catch (e) { return null; } })();
            filters.push({ columnIndex: i, header, criteria1: one, criteria2: two, operator: op });
          }
        }
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        autoFilterOn: active,
        filterRange: addr,
        columnCount: addr ? sheet.Range(addr).Columns.Count : 0,
        filters,
        warnings: [],
        message: active
          ? `工作表 [${sheet.Name}] 筛选范围 ${addr}，其中 ${filters.length} 列设有条件`
          : `工作表 [${sheet.Name}] 当前没有启用筛选`
      };
    }

    if (!range) throw new Error("缺少必要参数: range (例如 'A4:E20')");

    // 防御性激活：WPS 上部分 Range 级操作（AutoFilter / Sort）对非活动表可能静默 no-op。
    // 说明：曾把「筛选不生效」记成产品缺陷（ISS-117），复核后确认是**验证脚本没检查写入结果**——
    // 数据根本没写进表，空区域上 AutoFilter() 自然无效。激活保留作防御，但不是该问题的根因。
    try { sheet.Activate(); } catch (e) {}

    const targetRange = sheet.Range(range);
    const warnings = [];

    // ⚠️ 同一类问题（Lead 追修「先校验、后动手」）：sortRules 的越界校验原先排在 AutoFilter 之后，
    // 于是 `enableAutoFilter:true` + 越界 colIndex 会**先把筛选打开、再报错**。
    // 前移到任何写入之前（只读 Columns.Count，无副作用）；读不到列数时不拦，避免误拒。
    if (Array.isArray(sortRules) && sortRules.length > 0) {
      const colCountForCheck = Math.trunc(Number(safeRead(() => Number(targetRange.Columns.Count), 0)) || 0);
      if (colCountForCheck > 0) {
        for (const rule of sortRules) {
          const ci = Number(rule && rule.colIndex);
          if (!Number.isFinite(ci) || ci < 1 || ci > colCountForCheck) {
            throw new Error(
              `sortRules.colIndex=${rule && rule.colIndex} 越界：目标区域 ${range} 只有 ${colCountForCheck} 列（1..${colCountForCheck}）` +
              "；本次未做任何修改（未开启/关闭筛选，未排序）"
            );
          }
        }
      }
    }

    // 自动筛选控制
    let appliedFilterRange = null;
    if (enableAutoFilter !== undefined) {
      if (enableAutoFilter) {
        if (!sheet.AutoFilterMode) {
          targetRange.AutoFilter();
        }
      } else {
        if (sheet.AutoFilterMode) {
          sheet.AutoFilterMode = false;
        }
      }
      // 回读筛选实际覆盖范围：宿主会把筛选自动扩展到相邻的整块数据区，
      // 传入的 range 只是锚点而不是约束（问题台账 ISS-40）。这里把真实范围报出来，不再让调用方以为是自己传的那个。
      try {
        appliedFilterRange = sheet.AutoFilterMode && sheet.AutoFilter && sheet.AutoFilter.Range
          ? sheet.AutoFilter.Range.Address()
          : null;
      } catch (e) {
        appliedFilterRange = null;
      }
      if (appliedFilterRange && appliedFilterRange !== targetRange.Address()) {
        warnings.push(`筛选实际覆盖 ${appliedFilterRange}，与传入的 ${range} 不一致：宿主会把筛选扩展到相邻数据块。若需精确范围，请在目标区与其它数据之间留一个空行。`);
      }
      // **读回核对**：写入"不报错"≠"真的生效"。
      // 真机踩到：目标区域若已是**结构化表格**（ListObject），它自带筛选器，
      // 此时 `Range.AutoFilter()` 不生效、`sheet.AutoFilterMode` 仍为 false，
      // 但工具此前返回 success —— 调用方会以为筛选已开。
      if (enableAutoFilter && !appliedFilterRange) {
        let tableNames = [];
        try {
          const los = sheet.ListObjects;
          for (let i = 1; i <= Number(los.Count); i++) { try { tableNames.push(String(los.Item(i).Name)); } catch (e) {} }
        } catch (e) {}
        if (tableNames.length) {
          warnings.push(`筛选**未生效**：工作表上存在结构化表格 ${tableNames.join(" / ")}，它自带筛选器，表级 AutoFilter 不会另外打开。请直接在表格上操作，或先把表格转为普通区域。`);
        } else {
          warnings.push("筛选**未生效**：调用后 sheet.AutoFilterMode 仍为 false，请人工确认。");
        }
      }
    }

    // 数据排序：旧式 Range.Sort(...) 在本机 WPS 上会静默 no-op（问题台账 ISS-38），
    // 因此先走 SortFields，再读回校验，都无效时**报错而不是返回假成功**。
    let sortApplied = null;
    if (Array.isArray(sortRules) && sortRules.length > 0) {
      const rowCount = targetRange.Rows.Count;
      const colCount = targetRange.Columns.Count;
      // colIndex 越界时宿主静默 no-op（excel-tester M-6）。
      // 边界校验已在函数开头（任何写入之前）执行过，这里不再重复抛错，直接用列数做后续计算。
      const before = normalize2DArray(targetRange.Value2, rowCount, colCount);
      const attempts = [];
      let after = before;
      let sorted = false;

      const readCurrent = () => normalize2DArray(targetRange.Value2, rowCount, colCount);
      const sortKeyOf = (rule) => targetRange.Columns.Item(Number(rule.colIndex));
      const orderOf = (rule) => (rule.order === "desc" ? 2 : 1);

      // 本机 WPS 实测：`Range.Sort` 是**方法**而不是对象（访问 .SortFields 抛
      // "Cannot read properties of undefined (reading 'Clear')"），而 `sheet.Sort` 才是可用的 Sort 对象。
      // 因此按下面顺序逐条尝试，每一条都**读回校验**，全都不生效就报错（问题台账 ISS-38）。
      const sortPaths = [
        {
          name: "sheet.Sort.SortFields",
          run: () => {
            const s = sheet.Sort;
            s.SortFields.Clear();
            for (let i = 0; i < sortRules.length; i++) s.SortFields.Add(sortKeyOf(sortRules[i]), 0, orderOf(sortRules[i]));
            s.SetRange(targetRange);
            s.Header = 1; // 包含表头
            s.Apply();
          }
        },
        {
          name: "range.Sort.SortFields",
          run: () => {
            const s = targetRange.Sort;
            s.SortFields.Clear();
            for (let i = 0; i < sortRules.length; i++) s.SortFields.Add(sortKeyOf(sortRules[i]), 0, orderOf(sortRules[i]));
            s.SetRange(targetRange);
            s.Header = 1;
            s.Apply();
          }
        },
        {
          name: "range.Sort(旧式)",
          run: () => {
            const key1 = sortKeyOf(sortRules[0]), order1 = orderOf(sortRules[0]);
            let key2, order2;
            if (sortRules[1]) {
              key2 = sortKeyOf(sortRules[1]);
              order2 = orderOf(sortRules[1]);
            }
            targetRange.Sort(key1, order1, key2, order2, undefined, undefined, 1);
          }
        }
      ];

      for (let p = 0; p < sortPaths.length; p++) {
        const path = sortPaths[p];
        let error = null;
        try {
          path.run();
        } catch (e) {
          error = e.message;
        }
        // 即使抛错也读回：某条路径可能已经部分生效
        after = readCurrent();
        sorted = isSortedByRules(after, sortRules);
        attempts.push(path.name + (error ? ":异常(" + error + ")" : ":ok") + (sorted ? ":已生效" : ":未生效"));
        if (sorted) break;
      }

      if (!sorted) {
        throw new Error(
          `排序未生效：区域 ${range} 读回后的顺序不满足请求的排序规则。已尝试：${attempts.join("；")}。` +
          `请核对 colIndex 是否为**区域内相对列号**（1 = 区域第一列），或改用 wps_execute_script 的 sheet.Sort.SortFields 路径。`
        );
      }
      sortApplied = { changed: JSON.stringify(before) !== JSON.stringify(after), attempts };
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      range,
      enableAutoFilter: enableAutoFilter !== undefined ? !!enableAutoFilter : sheet.AutoFilterMode,
      appliedFilterRange,
      sortedRuleCount: Array.isArray(sortRules) ? sortRules.length : 0,
      sortApplied,
      warnings,
      message: `已成功在 [${sheet.Name}] ${range} 应用筛选与排序`
    };
  }
