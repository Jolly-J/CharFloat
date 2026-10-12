  // validation-write.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  /**
   * 写入前的**参数校验**（必须在 `Validation.Delete()` 之前跑完）。
   *
   * 为什么单独成函数：原实现是 `Validation.Delete()` → 再校验 `listItems` / `validationType`，
   * 调用方少传一个参数就会**先被清掉该区域原有的数据有效性、然后才收到报错**——
   * 比"静默 no-op"更具破坏性。原则：**先校验、后动手**，任何"校验失败还会留下副作用"
   * 的顺序都是错的。校验通过后返回可直接交给宿主 Add 的计划。
   */
  function buildDataValidationPlan(params) {
    const validationType = params.validationType === undefined || params.validationType === null ? "list" : params.validationType;
    const operator = params.operator === undefined || params.operator === null ? "between" : params.operator;
    const untouched = "；本次未做任何修改（原有校验保持不变）";

    if (validationType === "list") {
      const raw = params.listItems;
      const items = Array.isArray(raw)
        ? raw.map((x) => String(x).trim()).filter((x) => x !== "")
        : String(raw === undefined || raw === null ? "" : raw).split(",").map((x) => x.trim()).filter((x) => x !== "");
      if (items.length === 0) {
        throw new Error("validationType='list' 必须提供非空的 listItems（候选项数组），否则宿主会静默不设任何校验" + untouched);
      }
      return { kind: "list", listStr: items.join(","), requestedItems: items };
    }

    if (validationType === "number_range") {
      const opMap = { between: 1, greater_than: 5, less_than: 6, equal: 3 };
      const op = opMap[operator];
      if (op === undefined) {
        throw new Error(`未知的 operator: "${operator}"（支持 between / greater_than / less_than / equal）` + untouched);
      }
      const numeric = (label, v) => {
        const n = Number(v);
        if (!Number.isFinite(n)) throw new Error(`${label} 必须是有限数值，收到 "${v}"` + untouched);
        return String(n);
      };
      if (op === 1) {
        if (params.minVal === undefined || params.maxVal === undefined) {
          throw new Error("operator='between' 必须同时提供 minVal 与 maxVal（缺一个会写出没有上下限的规则）" + untouched);
        }
        return { kind: "number_range", op, f1: numeric("minVal", params.minVal), f2: numeric("maxVal", params.maxVal), boundSource: "minVal+maxVal" };
      }
      if (op === 6) {
        // schema 文档写的是"小于 maxVal"，旧实现取的却是 minVal。两者都接受，优先 maxVal（文档口径），
        // 用的是哪个如实回传，不再静默按 minVal 走。
        const useMax = params.maxVal !== undefined && params.maxVal !== null;
        const chosen = useMax ? params.maxVal : params.minVal;
        if (chosen === undefined || chosen === null) {
          throw new Error("operator='less_than' 必须提供 maxVal（上限，也兼容 minVal）" + untouched);
        }
        return { kind: "number_range", op, f1: numeric(useMax ? "maxVal" : "minVal", chosen), f2: undefined, boundSource: useMax ? "maxVal" : "minVal" };
      }
      if (params.minVal === undefined || params.minVal === null) {
        throw new Error(`operator='${operator}' 必须提供 minVal` + untouched);
      }
      return { kind: "number_range", op, f1: numeric("minVal", params.minVal), f2: undefined, boundSource: "minVal" };
    }

    throw new Error(`未知的 validationType: "${validationType}" (支持 list, number_range)` + untouched);
  }

  /** Add 抛错后的**尽力回滚**：把写入前读到的规则写回去，并读回核对是否真的回去了。 */
  function restoreValidationRule(targetRange, previous) {
    if (!previous || previous.ok === false) {
      return { attempted: false, restored: false, reason: previous && previous.ok === false ? "写入前未能读到原规则（" + previous.error + "）" : "读取原规则失败" };
    }
    if (!previous.rule) return { attempted: false, restored: false, reason: "写入前该区域本就没有校验" };
    const rule = previous.rule;
    try {
      try { targetRange.Validation.Delete(); } catch (e) {}
      const alertStyle = Number.isFinite(rule.alertStyle) ? rule.alertStyle : 1;
      if (rule.type === 3) {
        targetRange.Validation.Add(3, alertStyle, 1, rule.formula1 === null ? "" : rule.formula1);
      } else {
        targetRange.Validation.Add(rule.type, alertStyle, Number.isFinite(rule.operator) ? rule.operator : 1, rule.formula1, rule.formula2 === null ? undefined : rule.formula2);
      }
    } catch (e) {
      return { attempted: true, restored: false, reason: "回滚调用失败: " + e.message };
    }
    const after = readCellValidationRule(targetRange);
    if (after.ok && after.rule && after.rule.signature === rule.signature) return { attempted: true, restored: true };
    return { attempted: true, restored: false, reason: "回滚后读回与原规则不一致" };
  }

  /**
   * 写后读回核对：把宿主**真实状态**与请求值逐项比较。
   * 返回 { ok, warnings, readBack }；ok=false 表示核心规则没落上（调用方已为此返回失败）。
   */
  function verifyDataValidationRule(targetRange, sheet, plan, requested) {
    const warnings = [];
    const readBack = readCellValidationRule(targetRange);
    if (readBack.ok === false) {
      return { ok: false, warnings: ["写后读回失败：" + readBack.error], readBack: null };
    }
    const rule = readBack.rule;
    if (!rule) {
      return { ok: false, warnings: ["写后读回：该区域没有任何校验规则（宿主未落上）"], readBack: null };
    }
    const summary = {
      type: rule.type, typeName: rule.typeName,
      operator: rule.operator, operatorName: rule.operatorName,
      formula1: rule.formula1, formula2: rule.formula2,
      inCellDropdown: rule.inCellDropdown, ignoreBlank: rule.ignoreBlank
    };
    let ok = true;

    if (plan.kind === "list") {
      if (rule.type !== 3) {
        ok = false;
        warnings.push(`写后读回类型不符：请求 list(xlValidateList=3)，宿主读回 ${rule.typeName}(${rule.type})`);
      } else {
        const parsed = parseValidationListFormula(rule.formula1);
        let got = null;
        if (parsed.ok && parsed.kind === "literal") got = parsed.items;
        else if (parsed.ok && parsed.kind === "reference") {
          const resolved = readValidationListReference(sheet, parsed.reference);
          if (resolved.ok) got = resolved.items;
          else warnings.push("写后读回的候选项引用无法解析：" + resolved.error);
        } else if (!parsed.ok) {
          warnings.push("写后读回的候选项无法解析：" + parsed.reason);
        }
        if (got === null) {
          // 解析不了就不敢说"一致"，如实降级为未核对
          warnings.push(`候选项未能逐项核对：请求 ${plan.requestedItems.length} 项，宿主 Formula1 = ${rule.formula1 === null ? "null" : rule.formula1}`);
        } else {
          const norm = (arr) => arr.map((x) => String(x).trim()).sort();
          const a = norm(plan.requestedItems), b = norm(got);
          if (a.length !== b.length || a.some((x, i) => x !== b[i])) {
            ok = false;
            warnings.push(`候选项与请求不一致：请求 [${a.join(", ")}]，宿主读回 [${b.join(", ")}]`);
          }
        }
        if (rule.inCellDropdown === false) warnings.push("写后读回：InCellDropdown 为 false，下拉箭头未生效");
      }
    } else {
      if (rule.type !== 2) {
        ok = false;
        warnings.push(`写后读回类型不符：请求 decimal(xlValidateDecimal=2)，宿主读回 ${rule.typeName}(${rule.type})`);
      }
      if (rule.operator !== plan.op) {
        ok = false;
        warnings.push(`写后读回运算符不符：请求 ${plan.op}，宿主读回 ${rule.operator}`);
      }
      const sameNumber = (a, b) => {
        if (a === null || a === undefined || b === null || b === undefined) return false;
        const na = Number(String(a).replace(/^=/, "")), nb = Number(String(b));
        return Number.isFinite(na) && Number.isFinite(nb) && na === nb;
      };
      if (!sameNumber(plan.f1, rule.formula1)) {
        ok = false;
        warnings.push(`写后读回下限不符：请求 ${plan.f1}，宿主读回 ${rule.formula1 === null ? "null" : rule.formula1}`);
      }
      if (plan.op === 1 && !sameNumber(plan.f2, rule.formula2)) {
        ok = false;
        warnings.push(`写后读回上限不符：请求 ${plan.f2}，宿主读回 ${rule.formula2 === null ? "null" : rule.formula2}`);
      }
    }

    // 提示/报错文案：属于附加项，落不上只报警告（核心规则已核对通过）
    const g = (fn) => { try { const x = fn(); return x === undefined ? null : x; } catch (e) { return null; } };
    if (requested.promptMessage) {
      const got = g(() => targetRange.Validation.InputMessage);
      if (String(got === null ? "" : got) !== String(requested.promptMessage)) {
        warnings.push(`提示文案未落上：请求 "${requested.promptMessage}"，宿主读回 ${got === null ? "null" : `"${got}"`}`);
      }
    }
    if (requested.errorMessage) {
      const got = g(() => targetRange.Validation.ErrorMessage);
      if (String(got === null ? "" : got) !== String(requested.errorMessage)) {
        warnings.push(`报错文案未落上：请求 "${requested.errorMessage}"，宿主读回 ${got === null ? "null" : `"${got}"`}`);
      }
    }
    return { ok: ok, warnings: warnings, readBack: summary };
  }

  // 18. 单元格下拉验证菜单 (第三梯队)
  function setDataValidation(app, params) {
    // 写入参数（validationType / listItems / operator / minVal / maxVal）**不在这里解构**：
    // 它们统一由 buildDataValidationPlan(params) 校验后再用，避免"读了却没校验"或
    // "解构出默认值后误以为已校验"这两类问题。
    const {
      sheetName,
      workbookName,
      address,
      promptTitle,
      promptMessage,
      errorTitle,
      errorMessage
    } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-32 数据有效性读回：读整张表**所有**带校验的区域，或指定 address 的那一块。
    // 只写不读时 AI 无法回答"这列现在允许哪些值"，复核只能靠人看。
    if ((params || {}).action === "read") {
      const VALTYPE = { 1: "whole_number", 2: "decimal", 3: "list", 4: "date", 5: "time", 6: "text_length", 7: "custom" };
      const OP = { 1: "between", 2: "not_between", 3: "equal", 4: "not_equal", 5: "greater_than", 6: "less_than", 7: "greater_equal", 8: "less_equal" };
      const readOne = (rng) => {
        const v = rng.Validation;
        let type = null;
        try { type = Number(v.Type); } catch (e) { return null; }
        // Type = -4142 (xlValidateInputOnly) 表示没设校验，跳过
        if (type === -4142 || !Number.isFinite(type)) return null;
        const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
        return {
          address: g(() => rng.Address(), null),
          type: VALTYPE[type] || `unknown(${type})`,
          typeCode: type,
          operator: OP[g(() => Number(v.Operator), null)] || null,
          formula1: g(() => v.Formula1, null),
          formula2: g(() => v.Formula2, null),
          inCellDropdown: g(() => Boolean(v.InCellDropdown), null),
          ignoreBlank: g(() => Boolean(v.IgnoreBlank), null),
          showError: g(() => Boolean(v.ShowError), null),
          errorTitle: g(() => v.ErrorTitle, null),
          errorMessage: g(() => v.ErrorMessage, null),
          promptTitle: g(() => v.InputTitle, null),
          promptMessage: g(() => v.InputMessage, null)
        };
      };
      let items = [];
      if (address) {
        const one = readOne(sheet.Range(address));
        items = one ? [one] : [];
      } else {
        // 逐行扫描已用区域（宿主没有"枚举所有校验区域"的 API）
        const ur = sheet.UsedRange;
        const r0 = ur.Row, c0 = ur.Column;
        const rows = Math.min(ur.Rows.Count, 2000), cols = Math.min(ur.Columns.Count, 100);
        const colName = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
        for (let r = r0; r < r0 + rows; r++) {
          for (let c = c0; c < c0 + cols; c++) {
            const one = readOne(sheet.Range(`${colName(c)}${r}`));
            if (one) items.push(one);
          }
        }
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        count: items.length,
        validations: items,
        warnings: address ? [] : [],
        message: `工作表 [${sheet.Name}] 上共 ${items.length} 处数据有效性设置`
      };
    }

    if (!address) throw new Error("缺少必要参数: address (例如 'E5:E20')");

    // ⚠️ 顺序铁律（Lead 追修）：**先校验、后动手**。
    // 原实现是 `Validation.Delete()` → 再校验 listItems / validationType，
    // 少传一个参数会先把该区域原有的数据有效性清掉、然后才报错——破坏性比"静默 no-op"更强。
    const plan = buildDataValidationPlan(params || {});

    const targetRange = sheet.Range(address);

    // 写入前先记录原规则：① 失败时能告诉调用方被清掉了什么；② Add 抛错时尽力回滚。
    const previousRuleProbe = readCellValidationRule(targetRange);
    const previousRule = previousRuleProbe.ok && previousRuleProbe.rule ? {
      type: previousRuleProbe.rule.type,
      typeName: previousRuleProbe.rule.typeName,
      operator: previousRuleProbe.rule.operator,
      formula1: previousRuleProbe.rule.formula1,
      formula2: previousRuleProbe.rule.formula2
    } : null;

    try {
      targetRange.Validation.Delete();
    } catch (e) {}

    let addError = null;
    try {
      if (plan.kind === "list") {
        // Type: 3 (xlValidateList), AlertStyle: 1 (xlValidAlertStop), Operator: 1 (xlBetween)
        targetRange.Validation.Add(3, 1, 1, plan.listStr);
        targetRange.Validation.InCellDropdown = true;
      } else {
        // Type: 2 (xlValidateDecimal)
        targetRange.Validation.Add(2, 1, plan.op, plan.f1, plan.f2);
      }
    } catch (e) {
      addError = e;
    }

    if (addError) {
      // 宿主 Add 失败时原规则已被 Delete 掉 —— 尽力写回并读回核对，结果如实上报，不假装无事发生。
      const rollback = restoreValidationRule(targetRange, previousRuleProbe);
      return {
        success: false,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address,
        validationType: plan.kind,
        hostError: addError.message,
        previousRule: previousRule,
        rollback: rollback,
        warnings: [
          `宿主写入数据有效性失败: ${addError.message}`,
          rollback.restored
            ? "写入前的原有校验已尽力写回并读回核对通过"
            : `原有校验未能恢复（${rollback.reason}）——该区域现在可能没有校验，请重新设置`
        ],
        message: `在 [${sheet.Name}] ${address} 配置数据有效性失败（原校验${rollback.restored ? "已恢复" : "未恢复"}）`
      };
    }

    if (params.promptMessage) {
      targetRange.Validation.InputTitle = params.promptTitle || "选择提示";
      targetRange.Validation.InputMessage = params.promptMessage;
      targetRange.Validation.ShowInput = true;
    }

    if (params.errorMessage) {
      targetRange.Validation.ErrorTitle = params.errorTitle || "输入无效";
      targetRange.Validation.ErrorMessage = params.errorMessage;
      targetRange.Validation.ShowError = true;
    }

    // 写后读回核对：状态文案已承诺"写入并读回核对"，就必须真的核对，不能只回 success。
    const verify = verifyDataValidationRule(targetRange, sheet, plan, params || {});
    const warnings = verify.warnings.slice();
    if (plan.boundSource === "minVal") {
      warnings.push("operator='less_than' 未传 maxVal，本次按 minVal 作为上限写入（schema 文档口径是 maxVal，建议显式传 maxVal）");
    }

    return {
      success: verify.ok,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address,
      validationType: plan.kind === "list" ? "list" : "number_range",
      requested: plan.kind === "list"
        ? { listItems: plan.requestedItems }
        : { operator: plan.operator === undefined ? params.operator : params.operator, operatorCode: plan.op, formula1: plan.f1, formula2: plan.f2 === undefined ? null : plan.f2 },
      readBack: verify.readBack,
      replacedPreviousRule: previousRule,
      warnings: warnings,
      message: verify.ok
        ? `已在 [${sheet.Name}] ${address} 配置数据有效性并读回核对通过`
        : `已在 [${sheet.Name}] ${address} 调用宿主写入数据有效性，但**读回与请求不一致**（见 warnings），请勿当作已生效`
    };
  }
