  // validation-read.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // ---------------------------------------------------------------------------
  // CAP-10：数据验证**违规定位**
  // ---------------------------------------------------------------------------
  // CAP-32 已经能读回"区域上挂了什么校验规则"，但读不回"**存量数据里哪些单元格越界了**"，
  // 于是 AI 写完下拉/范围校验后仍然只能盲信 success。这里在读取区域时逐格比对规则，
  // 列出违规单元格（值 + 命中的规则 + 期望）。
  //
  // 判定原则（与仓库"success:true 不算数"的规矩一致）：
  //   - 规则读失败、规则值解析不了（自定义公式、无法解析的日期写法）→ 进 `unevaluated` 并附原因，
  //     **绝不当成"通过"**；
  //   - 扫描有上限，截断时 `truncated: true` 并在 warnings 里说清扫了多少 / 还剩多少；
  //   - 只用已证实可用的宿主 API（`Range.Value2` / `Range.Validation`）。
  //     整表枚举校验区域的 `SpecialCells(xlCellTypeAllValidation)` 未在本机验证过，
  //     且 dispatch.js 的反射护栏把它列为保守跳过项，故本轮**不采用**，改用有上限的逐格扫描。
  const VALIDATION_TYPE_NAMES = { 1: "whole_number", 2: "decimal", 3: "list", 4: "date", 5: "time", 6: "text_length", 7: "custom" };
  const VALIDATION_OPERATOR_NAMES = { 1: "between", 2: "not_between", 3: "equal", 4: "not_equal", 5: "greater_than", 6: "less_than", 7: "greater_equal", 8: "less_equal" };
  const VALIDATION_TYPE_NONE = -4142; // xlValidateInputOnly：宿主用它表示"该格没有校验"
  const VALIDATION_SCAN_DEFAULT = 300;
  const VALIDATION_SCAN_MAX = 2000;
  const VALIDATION_REPORT_MAX = 200;
  const VALIDATION_RULE_CELLS_MAX = 20;

  /**
   * 读取单格的校验规则。**三态返回**，"读失败"与"没有规则"必须分开：
   *   { ok: true,  rule: null }   宿主明确表示该格没有校验（Type = -4142）
   *   { ok: true,  rule: {...} }  读到规则
   *   { ok: false, error }        读取失败 → 调用方记 unevaluated，不能当"没规则"
   */
  function readCellValidationRule(cell) {
    let v;
    try { v = cell.Validation; } catch (e) { return { ok: false, error: "读取 Validation 失败: " + e.message }; }
    if (!v) return { ok: false, error: "Validation 对象为空" };
    let type;
    try { type = Number(v.Type); } catch (e) { return { ok: false, error: "读取 Validation.Type 失败: " + e.message }; }
    if (!Number.isFinite(type) || type === VALIDATION_TYPE_NONE) return { ok: true, rule: null };
    const g = (fn, fallback) => { try { const x = fn(); return x === undefined ? fallback : x; } catch (e) { return fallback; } };
    const operator = g(() => Number(v.Operator), null);
    const rule = {
      type: type,
      typeName: VALIDATION_TYPE_NAMES[type] || ("unknown(" + type + ")"),
      operator: Number.isFinite(operator) ? operator : null,
      operatorName: Number.isFinite(operator) ? (VALIDATION_OPERATOR_NAMES[operator] || ("unknown(" + operator + ")")) : null,
      formula1: g(() => (v.Formula1 === undefined || v.Formula1 === null ? null : String(v.Formula1)), null),
      formula2: g(() => (v.Formula2 === undefined || v.Formula2 === null ? null : String(v.Formula2)), null),
      ignoreBlank: g(() => Boolean(v.IgnoreBlank), null),
      inCellDropdown: g(() => Boolean(v.InCellDropdown), null),
      alertStyle: g(() => Number(v.AlertStyle), null)
    };
    rule.signature = [rule.type, rule.operator, rule.formula1, rule.formula2, rule.ignoreBlank].join("|");
    return { ok: true, rule: rule };
  }

  /** 拆字面量候选项：半角逗号优先（Excel 的列表分隔符），没有半角逗号时退回全角逗号。 */
  function splitValidationListLiteral(text) {
    const separator = text.indexOf(",") >= 0 ? "," : (text.indexOf("，") >= 0 ? "，" : ",");
    return text.split(separator).map((s) => s.trim()).filter((s) => s !== "");
  }

  /**
   * 解析 list 规则的 Formula1。Excel/WPS 读回时有三种形态：
   *   1. 字面量列表（读回常带外层双引号）："通过,不通过"
   *   2. 区域引用：=$D$1:$D$5 或 =Sheet1!$D$1:$D$5
   *   3. 裸区域引用（个别宿主不带等号）：$D$1:$D$5
   */
  function parseValidationListFormula(formula) {
    if (formula === null || formula === undefined) return { ok: false, reason: "list 规则没有给出 Formula1" };
    const text = String(formula).trim();
    if (text === "") return { ok: false, reason: "list 规则的 Formula1 为空" };
    if (text.charAt(0) === "=") return { ok: true, kind: "reference", reference: text };
    const unquoted = text.replace(/^"/, "").replace(/"$/, "");
    if (unquoted !== text) return { ok: true, kind: "literal", items: splitValidationListLiteral(unquoted) };
    if (/^'?[^!']*'?![$A-Za-z]/.test(text) || /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(text)) {
      return { ok: true, kind: "reference", reference: text };
    }
    return { ok: true, kind: "literal", items: splitValidationListLiteral(text) };
  }

  /** 读取 list 规则引用的区域，把其中的非空值作为候选列表。 */
  function readValidationListReference(sheet, reference) {
    let text = String(reference).replace(/^=/, "").trim();
    let targetSheet = sheet;
    const bang = text.lastIndexOf("!");
    if (bang >= 0) {
      const namePart = text.slice(0, bang).trim().replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'");
      text = text.slice(bang + 1).trim();
      try { targetSheet = sheet.Parent.Worksheets.Item(namePart); } catch (e) {
        return { ok: false, error: `候选项引用的工作表 "${namePart}" 不存在` };
      }
    }
    let rng;
    try { rng = targetSheet.Range(text); } catch (e) {
      return { ok: false, error: `候选项引用 ${reference} 无法解析: ${e.message}` };
    }
    let values;
    try { values = rng.Value2; } catch (e) {
      return { ok: false, error: `读取候选项引用 ${reference} 的值失败: ${e.message}` };
    }
    const items = [];
    const push = (v) => { if (v !== null && v !== undefined && String(v).trim() !== "") items.push(String(v)); };
    if (Array.isArray(values)) {
      values.forEach((row) => { if (Array.isArray(row)) row.forEach(push); else push(row); });
    } else {
      push(values);
    }
    return { ok: true, items: items, reference: reference };
  }

  function excelSerialFromTimestamp(ms) {
    return ms / 86400000 + 25569; // Excel 序列号 25569 = 1970-01-01（UTC）
  }

  /** 把规则里的比较值解析成可比较的数字（日期/时间统一转 Excel 序列号）。 */
  function parseValidationBound(text) {
    if (text === null || text === undefined) return { ok: false, reason: "规则未给出比较值" };
    let s = String(text).trim().replace(/^=/, "").trim();
    if (s === "") return { ok: false, reason: "规则的比较值为空" };
    if (/^-?\d+(\.\d+)?$/.test(s)) return { ok: true, value: Number(s) };
    const d = s.match(/^DATE\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
    if (d) return { ok: true, value: excelSerialFromTimestamp(Date.UTC(Number(d[1]), Number(d[2]) - 1, Number(d[3]))) };
    const t = s.match(/^TIME\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
    if (t) return { ok: true, value: (Number(t[1]) * 3600 + Number(t[2]) * 60 + Number(t[3])) / 86400 };
    const plain = s.replace(/^"|"$/g, "");
    const parsed = Date.parse(plain);
    if (!Number.isNaN(parsed)) return { ok: true, value: excelSerialFromTimestamp(parsed), fromText: plain };
    return { ok: false, reason: `无法把规则比较值 "${text}" 解析为数值/日期` };
  }

  /** 单元格值 → 可比较的数字；非数值文本不算"跳过"，而是明确不满足数值校验。 */
  function coerceValidationNumber(value) {
    if (typeof value === "number") return { ok: true, value: value };
    if (typeof value === "boolean") return { ok: false, reason: "布尔值不是数值" };
    const raw = String(value).trim();
    if (raw === "") return { ok: false, reason: "空值" };
    const cleaned = raw.replace(/,/g, "").replace(/^[¥$€£]\s*/, "");
    if (/^-?\d+(\.\d+)?$/.test(cleaned)) return { ok: true, value: Number(cleaned), fromText: raw };
    const pct = cleaned.match(/^(-?\d+(\.\d+)?)%$/);
    if (pct) return { ok: true, value: Number(pct[1]) / 100, fromText: raw };
    const parsed = Date.parse(cleaned.replace(/^"|"$/g, ""));
    if (!Number.isNaN(parsed)) return { ok: true, value: excelSerialFromTimestamp(parsed), fromText: raw };
    return { ok: false, reason: `"${raw}" 不是可比较的数值/日期` };
  }

  /** 规则的自然语言"期望"，直接进违规明细，AI 不用自己翻译运算符。 */
  function describeValidationRule(rule, listInfo) {
    const t = rule.typeName;
    const op = rule.operatorName;
    const f1 = rule.formula1 === null ? "(空)" : rule.formula1;
    const f2 = rule.formula2 === null ? "(空)" : rule.formula2;
    if (t === "list") {
      if (listInfo && listInfo.kind === "reference" && listInfo.ok) {
        return `必须是 ${listInfo.reference} 中的一项（共 ${listInfo.items.length} 项）`;
      }
      if (listInfo && listInfo.kind === "literal") {
        const shown = listInfo.items.slice(0, 12);
        return `必须是列表中的一项: ${shown.join(" / ")}${listInfo.items.length > shown.length ? ` …（共 ${listInfo.items.length} 项）` : ""}`;
      }
      return `必须是 ${f1} 中的一项`;
    }
    if (t === "custom") return `自定义公式规则: ${f1}（宿主外无法求值）`;
    const unit = t === "date" ? "日期" : (t === "time" ? "时间" : (t === "text_length" ? "文本长度" : "数值"));
    if (t === "text_length") {
      const map = { between: `长度必须介于 ${f1} 与 ${f2} 之间`, not_between: `长度必须不在 ${f1}~${f2} 之间`, equal: `长度必须等于 ${f1}`, not_equal: `长度必须不等于 ${f1}`, greater_than: `长度必须大于 ${f1}`, less_than: `长度必须小于 ${f1}`, greater_equal: `长度必须不小于 ${f1}`, less_equal: `长度必须不大于 ${f1}` };
      return map[op] || `长度校验（运算符 ${op === null ? "未知" : op}，比较值 ${f1}）`;
    }
    const map = {
      between: `${unit}必须介于 ${f1} 与 ${f2} 之间（含端点）`,
      not_between: `${unit}必须不在 ${f1}~${f2} 之间`,
      equal: `${unit}必须等于 ${f1}`,
      not_equal: `${unit}必须不等于 ${f1}`,
      greater_than: `${unit}必须大于 ${f1}`,
      less_than: `${unit}必须小于 ${f1}`,
      greater_equal: `${unit}必须不小于 ${f1}`,
      less_equal: `${unit}必须不大于 ${f1}`
    };
    return map[op] || `${unit}校验（运算符 ${op === null ? "未知" : op}，比较值 ${f1}${rule.formula2 === null ? "" : " / " + f2}）`;
  }

  /**
   * 解析（并按签名缓存）list 规则的候选项来源。同一签名只解析一次，
   * 避免每个单元格都去重读引用的区域。
   */
  function resolveValidationListInfo(rule, listCache) {
    if (listCache[rule.signature]) return listCache[rule.signature];
    const parsed = parseValidationListFormula(rule.formula1);
    let info;
    if (!parsed.ok) {
      info = { kind: "error", ok: false, reason: parsed.reason };
    } else if (parsed.kind === "reference") {
      const referenced = readValidationListReference(listCache.__sheet, parsed.reference);
      info = referenced.ok
        ? { kind: "reference", ok: true, items: referenced.items, reference: parsed.reference }
        : { kind: "reference", ok: false, reason: referenced.error, reference: parsed.reference };
    } else {
      info = { kind: "literal", ok: true, items: parsed.items };
    }
    listCache[rule.signature] = info;
    return info;
  }

  /**
   * 逐格比对规则。返回:
   *   { status: "pass" } / { status: "violate", reason, expectation } / { status: "unevaluated", reason }
   */
  function evaluateValidationRule(rule, value, listCache) {
    if (rule.type === 7) {
      return { status: "unevaluated", reason: "自定义公式规则（xlValidateCustom）无法在宿主外求值" };
    }
    if (rule.type === 3) {
      const info = resolveValidationListInfo(rule, listCache);
      const expectation = describeValidationRule(rule, info);
      if (!info.ok) return { status: "unevaluated", reason: info.reason, expectation: expectation };
      const actual = String(value).trim();
      const hit = info.items.find((item) => item === actual)
        || info.items.find((item) => item.toLowerCase() === actual.toLowerCase());
      if (hit !== undefined) return { status: "pass", expectation: expectation };
      return { status: "violate", reason: "值不在允许的候选项中", expectation: expectation, allowed: info.items.slice(0, 30) };
    }
    if (rule.type === 4 || rule.type === 5 || rule.type === 6) {
      // date / time / text_length 都按数值比较（date/time 用序列号，text_length 用字符数）
      const left = rule.type === 6 ? { ok: true, value: String(value).length } : coerceValidationNumber(value);
      const expectation = describeValidationRule(rule);
      if (!left.ok) return { status: "violate", reason: left.reason, expectation: expectation };
      const f1 = parseValidationBound(rule.formula1);
      const f2 = parseValidationBound(rule.formula2);
      if (!f1.ok) return { status: "unevaluated", reason: "规则下限无法解析：" + f1.reason, expectation: expectation };
      const isBetween = rule.operator === 1 || rule.operator === 2;
      if (isBetween && !f2.ok) return { status: "unevaluated", reason: "规则上限无法解析：" + f2.reason, expectation: expectation };
      const verdict = compareByOperator(rule.operator, left.value, f1.value, isBetween ? f2.value : null);
      if (verdict === null) return { status: "unevaluated", reason: "未知运算符 " + rule.operator, expectation: expectation };
      return verdict ? { status: "pass", expectation: expectation } : { status: "violate", reason: "不满足 " + expectation, expectation: expectation };
    }
    if (rule.type === 1 || rule.type === 2) {
      const expectation = describeValidationRule(rule);
      const left = coerceValidationNumber(value);
      if (!left.ok) return { status: "violate", reason: left.reason, expectation: expectation };
      const f1 = parseValidationBound(rule.formula1);
      const f2 = parseValidationBound(rule.formula2);
      if (!f1.ok) return { status: "unevaluated", reason: "规则下限无法解析：" + f1.reason, expectation: expectation };
      const isBetween = rule.operator === 1 || rule.operator === 2;
      if (isBetween && !f2.ok) return { status: "unevaluated", reason: "规则上限无法解析：" + f2.reason, expectation: expectation };
      const verdict = compareByOperator(rule.operator, left.value, f1.value, isBetween ? f2.value : null);
      if (verdict === null) return { status: "unevaluated", reason: "未知运算符 " + rule.operator, expectation: expectation };
      return verdict ? { status: "pass", expectation: expectation } : { status: "violate", reason: "不满足 " + expectation, expectation: expectation };
    }
    return { status: "unevaluated", reason: `暂不支持的校验类型 ${rule.typeName}`, expectation: describeValidationRule(rule) };
  }

  /** 运算符判定；返回 null 表示运算符未知（→ unevaluated，不猜通过）。 */
  function compareByOperator(operator, value, f1, f2) {
    switch (operator) {
      case 1: return f2 === null ? null : (value >= f1 && value <= f2);
      case 2: return f2 === null ? null : (value < f1 || value > f2);
      case 3: return value === f1;
      case 4: return value !== f1;
      case 5: return value > f1;
      case 6: return value < f1;
      case 7: return value >= f1;
      case 8: return value <= f1;
      default: return null;
    }
  }

  /**
   * 区域级违规定位。扫描范围内每个单元格自己的规则（同一区域可能有多种规则），
   * 用一次性 `Range.Value2` 取回区块值，逐格判定。
   */
  function collectValidationViolations(sheet, range, options) {
    const config = options || {};
    const warnings = [];
    const unevaluated = [];
    const violations = [];
    const rules = [];
    const rulesByKey = {};
    const ruleObjectsByKey = {};
    const listCache = { __sheet: sheet };
    let violationsOverflow = 0;
    const maxScanCells = Math.max(1, Math.min(VALIDATION_SCAN_MAX, Math.trunc(Number(config.maxCells) || VALIDATION_SCAN_DEFAULT)));

    const rows = Math.max(1, Math.trunc(Number(safeRead(() => Number(range.Rows.Count), 1)) || 1));
    const cols = Math.max(1, Math.trunc(Number(safeRead(() => Number(range.Columns.Count), 1)) || 1));
    // range.Row / range.Column 是区域左上角的 1 基坐标；读不到就置 0，改用 cell.Address() 兜底。
    const firstRow = Math.trunc(Number(safeRead(() => Number(range.Row), 0)) || 0);
    const firstCol = Math.trunc(Number(safeRead(() => Number(range.Column), 0)) || 0);
    const totalCells = rows * cols;
    const scannedCells = Math.min(totalCells, maxScanCells);
    const truncated = totalCells > scannedCells;
    if (truncated) {
      warnings.push(`范围内共 ${totalCells} 个单元格，本次只扫描前 ${scannedCells} 个（含表头行优先）；如需覆盖其余部分请缩小 address 或调大 maxCells`);
    }

    // 值一次性取回（1 次宿主调用）；超上限时逐格读，避免为少数单元格拉整块大矩阵。
    let blockValues = null;
    let blockValuesError = null;
    if (!truncated) {
      try {
        blockValues = normalize2DArray(range.Value2, rows, cols);
      } catch (e) {
        blockValuesError = "一次性读取区域值失败: " + e.message;
      }
    }

    let validatedCells = 0;
    let skippedBlankCells = 0;
    let blankWithUnknownIgnoreBlank = 0;
    let readErrorCells = 0;
    for (let i = 0; i < scannedCells; i++) {
      const r = Math.floor(i / cols);
      const c = i % cols;
      // 地址：range.Row/Column 是区域左上角的 **1 基**坐标，r/c 是 0 基偏移，
      // 所以是 firstRow + r 而不是 firstRow + r + 1（多 +1 会把违规定位到邻居格上——模拟宿主已复现）。
      const knownOrigin = firstRow > 0 && firstCol > 0;
      let address = knownOrigin ? (excelColumnName(firstCol + c) + (firstRow + r)) : null;

      let cell;
      try {
        cell = range.Cells.Item(r + 1, c + 1);
      } catch (e) {
        readErrorCells++;
        unevaluated.push({ address: address || `#${i + 1}`, reason: "定位单元格失败: " + e.message });
        continue;
      }
      if (!address) address = safeRead(() => cell.Address(), null) || `#${i + 1}`;

      let value;
      if (blockValues && blockValuesError === null) {
        const row = blockValues[r];
        if (Array.isArray(row)) value = c < row.length ? row[c] : null;
        else if (blockValues.length === 1 && r === 0) value = c === 0 ? blockValues[0] : null;
        else value = null;
      } else {
        value = safeRead(() => cell.Value2, null);
      }

      const readRes = readCellValidationRule(cell);
      if (!readRes.ok) {
        readErrorCells++;
        unevaluated.push({ address: address, reason: readRes.error });
        continue;
      }
      if (!readRes.rule) continue; // 该格没有校验 → 与"越界"无关

      validatedCells++;
      const rule = readRes.rule;
      if (!rulesByKey[rule.signature]) {
        const group = {
          ruleKey: "R" + (rules.length + 1),
          type: rule.type, typeName: rule.typeName,
          operator: rule.operator, operatorName: rule.operatorName,
          formula1: rule.formula1, formula2: rule.formula2,
          ignoreBlank: rule.ignoreBlank, inCellDropdown: rule.inCellDropdown,
          cellCount: 0, sampleCells: []
        };
        rulesByKey[rule.signature] = group;
        ruleObjectsByKey[group.ruleKey] = rule;
        rules.push(group);
      }
      const group = rulesByKey[rule.signature];
      group.cellCount++;
      if (group.sampleCells.length < VALIDATION_RULE_CELLS_MAX) group.sampleCells.push(address);

      const isBlank = value === null || value === undefined || (typeof value === "string" && value.trim() === "");
      if (isBlank) {
        // 规则明确不允许空值（Excel 的"忽略空值"未勾选，IgnoreBlank=false）时，
        // 空单元格本身就越界——Excel 的"圈释无效数据"也是这么算的。
        if (rule.ignoreBlank === false) {
          if (violations.length < VALIDATION_REPORT_MAX) {
            violations.push({
              address: address,
              value: null,
              ruleKey: group.ruleKey,
              ruleType: rule.typeName,
              operator: rule.operatorName,
              formula1: rule.formula1,
              formula2: rule.formula2,
              expectation: describeValidationRule(rule, rule.type === 3 ? resolveValidationListInfo(rule, listCache) : null),
              reason: "规则未允许空值（IgnoreBlank=false），该单元格为空"
            });
          } else {
            violationsOverflow++;
          }
        } else {
          skippedBlankCells++;
          if (rule.ignoreBlank === null) blankWithUnknownIgnoreBlank++;
        }
        continue;
      }

      const verdict = evaluateValidationRule(rule, value, listCache);
      if (verdict.status === "pass") continue;
      if (verdict.status === "unevaluated") {
        unevaluated.push({ address: address, ruleKey: group.ruleKey, value: value, reason: verdict.reason });
        continue;
      }
      if (violations.length < VALIDATION_REPORT_MAX) {
        const entry = {
          address: address,
          value: value,
          ruleKey: group.ruleKey,
          ruleType: rule.typeName,
          operator: rule.operatorName,
          formula1: rule.formula1,
          formula2: rule.formula2,
          expectation: verdict.expectation,
          reason: verdict.reason
        };
        if (verdict.allowed) entry.allowedValues = verdict.allowed;
        violations.push(entry);
      } else {
        violationsOverflow++;
      }
    }

    // 规则清单补上"期望"文案（list 规则要按签名去解析引用/字面量，缓存在 listCache 里）
    rules.forEach((group) => {
      const rule = ruleObjectsByKey[group.ruleKey];
      if (!rule) return;
      const info = rule.type === 3 ? resolveValidationListInfo(rule, listCache) : null;
      group.expectation = describeValidationRule(rule, info);
    });

    const violationsTruncated = violationsOverflow > 0;
    if (violationsTruncated) {
      warnings.push(`违规条目已达上报上限 ${VALIDATION_REPORT_MAX} 条，另有 ${violationsOverflow} 个越界单元格未逐条列出（validatedCells=${validatedCells}）`);
    }
    if (unevaluated.length > 0) {
      warnings.push(`有 ${unevaluated.length} 个单元格无法判定（规则读失败/自定义公式/无法解析的比较值），已列入 unevaluated，**未计入通过**`);
    }
    if (blankWithUnknownIgnoreBlank > 0) {
      warnings.push(`有 ${blankWithUnknownIgnoreBlank} 个空格没能读到 IgnoreBlank，无法确定"空值是否算越界"，已按跳过处理（可能漏报）`);
    }
    if (blockValuesError) warnings.push(blockValuesError);

    return {
      rangeAddress: safeRead(() => range.Address(), null),
      scannedCells: scannedCells,
      totalCells: totalCells,
      truncated: truncated,
      validatedCells: validatedCells,
      skippedBlankCells: skippedBlankCells,
      blankWithUnknownIgnoreBlank: blankWithUnknownIgnoreBlank,
      readErrorCells: readErrorCells,
      ruleCount: rules.length,
      rules: rules,
      violations: violations,
      violationCount: violations.length,
      violationsTruncated: violationsTruncated,
      unevaluated: unevaluated.slice(0, VALIDATION_REPORT_MAX),
      unevaluatedCount: unevaluated.length,
      warnings: warnings,
      message: `已比对 [${sheet.Name}] ${safeRead(() => range.Address(), "")}：${validatedCells} 个带校验的单元格中 ${violations.length} 个取值越界` +
        (unevaluated.length ? `，另有 ${unevaluated.length} 个无法判定` : "") +
        (truncated ? "（扫描被截断，见 warnings）" : "")
    };
  }
