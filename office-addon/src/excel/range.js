// ── 模块: src/excel/range.js — 区域读写、结构与查找替换 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 614-918 行（原样搬迁，未改写）。
  // 3. 区域核心读写
  async function handleReadRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      let range;
      if (params.address) {
        range = sheet.getRange(params.address);
      } else {
        range = sheet.getUsedRange();
      }

      range.load("address, rowCount, columnCount, values, formulas, text, numberFormat");
      await context.sync();

      if (params.metadata_only) {
        return {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount
        };
      }

      return {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        formulas: range.formulas,
        text: range.text
      };
    });
  }

  async function handleGetRangeStyles(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      range.load("address, format/font/name, format/font/size, format/font/bold, format/font/color, format/fill/color, format/horizontalAlignment, format/verticalAlignment, numberFormat");
      await context.sync();

      return {
        success: true,
        address: range.address,
        font: {
          name: range.format.font.name,
          size: range.format.font.size,
          bold: range.format.font.bold,
          color: range.format.font.color
        },
        fill: {
          color: range.format.fill.color
        },
        alignment: {
          horizontal: range.format.horizontalAlignment,
          vertical: range.format.verticalAlignment
        },
        numberFormat: range.numberFormat ? range.numberFormat[0]?.[0] : null
      };
    });
  }

  async function handleWriteRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      let range;
      if (params.address) {
        range = sheet.getRange(params.address);
      } else {
        throw new Error("write_range 必须指定 address 参数 (如 'A1' 或 'B2:D5')");
      }

      let values = params.values || params.data;
      let formulas = params.formulas;

      // 标准化 values 为规范二维数组
      if (values !== undefined && values !== null) {
        if (!Array.isArray(values)) {
          values = [[values]];
        } else if (!Array.isArray(values[0])) {
          values = [values];
        }
      }

      // 标准化 formulas 为规范二维数组
      if (formulas !== undefined && formulas !== null) {
        if (!Array.isArray(formulas)) {
          formulas = [[formulas]];
        } else if (!Array.isArray(formulas[0])) {
          formulas = [formulas];
        }
      }

      const primaryMatrix = formulas || values;
      if (!primaryMatrix) {
        throw new Error("write_range 必须提供 values 或 formulas 参数");
      }

      const rowCount = primaryMatrix.length;
      const columnCount = primaryMatrix[0].length;

      // 如果给的是单一左上角单元格且提供了二维数组，自动按矩阵尺寸扩展 range
      if (!params.address.includes(":") && (rowCount > 1 || columnCount > 1)) {
        range = range.getResizedRange(rowCount - 1, columnCount - 1);
      }

      // 如果仅有 values，但 values 中包含以 '=' 开头的字符串，自动提取为公式矩阵
      if (values && !formulas) {
        const hasFormulasInValues = values.some(row => Array.isArray(row) && row.some(cell => typeof cell === 'string' && cell.startsWith('=')));
        if (hasFormulasInValues) {
          formulas = values.map(row => (Array.isArray(row) ? row.map(cell => (typeof cell === 'string' && cell.startsWith('=')) ? cell : null) : [null]));
        }
      }

      // 赋值 values (先写入所有常规数值与文字)
      if (values) {
        range.values = values;
      }

      // 赋值 formulas (仅针对实际包含公式的单元格精准注入，杜绝空字符串抹除数据)
      if (formulas) {
        for (let r = 0; r < rowCount; r++) {
          for (let c = 0; c < columnCount; c++) {
            const f = formulas[r]?.[c];
            if (f && typeof f === 'string' && f.startsWith('=')) {
              range.getCell(r, c).formulas = [[f]];
            }
          }
        }
      }

      // 关键：在 sync 前必须显式 load 需要返回的属性，绝不裸读！
      range.load("address");
      sheet.load("name");
      await context.sync();

      return {
        success: true,
        address: range.address || params.address,
        sheetName: sheet.name,
        rowCount: rowCount,
        columnCount: columnCount,
        modifiedCount: rowCount * columnCount
      };
    });
  }

  /** 列号 → 列字母（1 → A，27 → AA）。 */
  function columnIndexToLetters(colIdx) {
    let temp = Math.floor(colIdx);
    let letter = '';
    while (temp > 0) {
      const mod = (temp - 1) % 26;
      letter = String.fromCharCode(mod + 65) + letter;
      temp = Math.floor((temp - mod) / 26);
    }
    return letter;
  }

  /**
   * 维度归一（ISS-93）：网关用 WPS 语义 `targetType`（'row' | 'column'），
   * 早期 Office.js 实现读的是 `dimension`（'rows' | 'columns'）且**默认 rows** ——
   * 结果"想插列却插行"。这里两套字段名都认，且都识别不了时**抛错**，绝不默认成行。
   */
  function resolveDimension(params) {
    const raw = params.targetType !== undefined ? params.targetType : params.dimension;
    if (raw === undefined || raw === null || raw === '') {
      throw new Error(
        '[Office.js 通道] manage_rows_and_columns 缺少 targetType：必须显式给出 "row" 或 "column"。' +
        '原实现默认按行处理 → 传 "column" 时静默插行，已阻断。'
      );
    }
    const s = String(raw).trim().toLowerCase();
    if (['row', 'rows', '行', 'r'].includes(s)) return 'rows';
    if (['column', 'columns', 'col', 'cols', '列', 'c'].includes(s)) return 'columns';
    throw new Error(`[Office.js 通道] manage_rows_and_columns 无法识别的 targetType: "${raw}"（支持 "row" | "column"）。`);
  }

  /** index 归一：行用数字（1 基），列允许数字或列字母（'B'）。 */
  function resolveStartIndex(params, dimension) {
    const raw = params.index !== undefined ? params.index : params.startIndex;
    if (raw === undefined || raw === null || raw === '') {
      throw new Error('[Office.js 通道] manage_rows_and_columns 缺少 index：请给出起始行号或列号（数字，从 1 开始；列也接受字母如 "B"）。');
    }
    if (dimension === 'columns' && typeof raw === 'string' && /^[A-Za-z]{1,3}$/.test(raw.trim())) {
      const letters = raw.trim().toUpperCase();
      let idx = 0;
      for (const ch of letters) idx = idx * 26 + (ch.charCodeAt(0) - 64);
      return idx;
    }
    const n = Number(raw);
    if (!Number.isFinite(n) || n < 1) {
      throw new Error(`[Office.js 通道] manage_rows_and_columns 的 index="${raw}" 非法：需要 ≥1 的数字或列字母。`);
    }
    return Math.floor(n);
  }

  async function handleUpdateRangeStructure(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const rawAction = String(params.action || 'insert').trim().toLowerCase();
      const actionAliases = {
        insert: 'insert', add: 'insert',
        delete: 'delete', remove: 'delete',
        hide: 'hide',
        unhide: 'unhide', show: 'unhide',
        set_size: 'set_size', set_row_height: 'set_size', set_column_width: 'set_size'
      };
      const action = actionAliases[rawAction];
      if (!action) {
        throw new Error(
          `[Office.js 通道] manage_rows_and_columns 无法识别的 action: "${params.action}"。` +
          '支持 insert | delete | hide | unhide | set_size（原实现对未知 action 返回 success 但什么都不做，已改为显式报错）。'
        );
      }

      const dimension = resolveDimension(params);
      const count = Math.floor(Number(params.count ?? 1));
      if (!Number.isFinite(count) || count < 1) {
        throw new Error(`[Office.js 通道] manage_rows_and_columns 的 count=${params.count} 非法：需要 ≥1 的整数。`);
      }

      let range;
      let targetAddress;
      if (params.address) {
        targetAddress = params.address;
        range = sheet.getRange(targetAddress);
      } else {
        const start = resolveStartIndex(params, dimension);
        if (dimension === 'columns') {
          targetAddress = `${columnIndexToLetters(start)}:${columnIndexToLetters(start + count - 1)}`;
        } else {
          targetAddress = `${start}:${start + count - 1}`;
        }
        range = sheet.getRange(targetAddress);
      }

      if (action === 'set_size') {
        const size = Number(params.size);
        if (params.size === undefined || params.size === null || !Number.isFinite(size) || size <= 0) {
          throw new Error(
            `[Office.js 通道] manage_rows_and_columns(action="set_size") 缺少合法 size：` +
            `targetType='row' 时为磅值行高（如 24），'column' 时为字符列宽（如 15）。未执行任何修改。`
          );
        }
        if (dimension === 'columns') {
          range.format.columnWidth = size;
        } else {
          range.format.rowHeight = size;
        }
        await context.sync();

        // 写后读回：不把"请求成功"当成"尺寸已变"。
        range.load('format/rowHeight, format/columnWidth');
        await context.sync();
        const applied = dimension === 'columns' ? range.format.columnWidth : range.format.rowHeight;
        return {
          success: true,
          action,
          targetType: dimension === 'columns' ? 'column' : 'row',
          address: targetAddress,
          requestedSize: size,
          appliedSize: applied,
          warnings: Math.abs(Number(applied) - size) > 0.01 ? [`设置后读回 ${applied}，与请求 ${size} 不一致（宿主可能按内容或缩放换算）。`] : []
        };
      }

      if (action === 'hide' || action === 'unhide') {
        range.hidden = action === 'hide';
        await context.sync();
        range.load('hidden');
        await context.sync();
        return { success: true, action, targetType: dimension === 'columns' ? 'column' : 'row', address: targetAddress, hidden: range.hidden };
      }

      const shift = params.shift || (dimension === 'columns' ? 'Right' : 'Down');
      if (action === 'insert') {
        range.insert(shift);
      } else {
        range.delete(dimension === 'columns' ? 'Left' : 'Up');
      }
      // 关键：属性必须显式 load 再 sync 才能读；直接读未 load 的 sheet.name 会抛
      // 「属性"name"不可用」——且此时**写入已经执行**，调用方会误判为整体失败（实机踩到）。
      sheet.load('name');
      await context.sync();
      return {
        success: true,
        action,
        targetType: dimension === 'columns' ? 'column' : 'row',
        address: targetAddress,
        count,
        shift: action === 'insert' ? shift : undefined,
        sheetName: sheet.name
      };
    });
  }

  async function handleClearRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const targetRange = params.address || params.range;
      if (!targetRange) {
        throw new Error('[Office.js 通道] clear_range 缺少必要参数: address（如 "A1:E20"）。未执行任何修改。');
      }
      const range = sheet.getRange(targetRange);
      const applyTo = params.applyTo || "All";
      range.clear(applyTo);
      await context.sync();
      return { success: true, address: targetRange, cleared: applyTo };
    });
  }

  async function handleCopyRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const source = params.sourceAddress || params.sourceRange || params.from;
      const destination = params.destinationAddress || params.destinationRange || params.to;
      if (!source || !destination) {
        throw new Error('[Office.js 通道] copy_range 需要 sourceAddress 与 destinationAddress（别名 sourceRange/destinationRange）。未执行任何修改。');
      }
      const sourceRange = sheet.getRange(source);
      const destRange = sheet.getRange(destination);
      destRange.copyFrom(sourceRange, params.copyType || "All");
      await context.sync();
      return { success: true, sourceAddress: source, destinationAddress: destination, copyType: params.copyType || "All" };
    });
  }

  async function handleSetHyperlink(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const address = params.address || params.cell || params.range;
      if (!address || !params.url) {
        throw new Error('[Office.js 通道] set_hyperlink 需要 address 与 url。未执行任何修改。');
      }
      const range = sheet.getRange(address);
      range.hyperlink = {
        address: params.url,
        textToDisplay: params.textToDisplay || params.displayText || params.url,
        screenTip: params.screenTip || ""
      };
      await context.sync();
      return { success: true, address, url: params.url };
    });
  }

  /** WPS 语义 validationType → Office.js DataValidationRule 的规则键。 */
  const VALIDATION_TYPE_MAP = {
    list: 'list',
    number_range: 'wholeNumber',
    number: 'wholeNumber',
    whole_number: 'wholeNumber',
    integer: 'wholeNumber',
    decimal: 'decimal',
    date: 'date',
    time: 'time',
    text_length: 'textLength',
    textlength: 'textLength',
    custom: 'custom'
  };

  /** 网关 operator（WPS 语义）→ Excel.DataValidationOperator 的字符串值（用字符串以免旧宿主缺枚举）。 */
  const VALIDATION_OPERATOR_MAP = {
    between: 'Between',
    not_between: 'NotBetween',
    greater_than: 'GreaterThan',
    greater_than_or_equal: 'GreaterThanOrEqualTo',
    less_than: 'LessThan',
    less_than_or_equal: 'LessThanOrEqualTo',
    equal: 'EqualTo',
    equals: 'EqualTo',
    not_equal: 'NotEqualTo'
  };

  /** 数值/日期公式统一转 number；日期字符串保持字符串（ISO）。 */
  function toValidationFormula(value, ruleKey, label) {
    if (value === undefined || value === null || value === '') {
      throw new Error(`[Office.js 通道] set_data_validation 的 ${label} 缺失：当前验证类型需要该阈值。`);
    }
    if (ruleKey === 'date' || ruleKey === 'time') return String(value);
    const n = Number(value);
    if (!Number.isFinite(n)) {
      throw new Error(`[Office.js 通道] set_data_validation 的 ${label}="${value}" 不是合法数值。`);
    }
    return n;
  }

  /** 去掉 Excel 内联列表来源外层的成对引号（桥接侧可能已加引号）。 */
  function stripOuterQuotes(source) {
    const s = String(source).trim();
    return s.length >= 2 && s.startsWith('"') && s.endsWith('"') ? s.slice(1, -1) : s;
  }

  /**
   * 构造数据有效性规则（ISS-93-a）。
   *
   * 网关（WPS 语义）发的是 validationType / listItems / operator / minVal / maxVal / prompt* / error*；
   * 原 Office.js 实现只读 `params.rule` → 先 `clear()` 再什么都不设，**静默清空既有校验**。
   * 现在：两套字段名都认（WPS 语义优先），且**构造失败时在 clear() 之前抛错**，不清空任何东西。
   */
  function buildDataValidationRule(params) {
    // 兼容路径：调用方（含桥接侧 normalizer）直接给 Office.js 原生规则对象。
    if (params.rule !== undefined && params.rule !== null && !params.validationType) {
      if (typeof params.rule !== 'object') {
        throw new Error('[Office.js 通道] set_data_validation 的 rule 必须是 Office.js DataValidationRule 对象。');
      }
      const rule = { ...params.rule };
      if (rule.list && rule.list.source !== undefined) {
        rule.list = { ...rule.list, source: stripOuterQuotes(rule.list.source) };
      }
      return { rule, typeLabel: 'rule(原生对象)' };
    }

    const rawType = params.validationType !== undefined && params.validationType !== null && params.validationType !== ''
      ? String(params.validationType)
      : 'list';
    const ruleKey = VALIDATION_TYPE_MAP[rawType.trim().toLowerCase()];
    if (!ruleKey) {
      throw new Error(
        `[Office.js 通道] set_data_validation 无法识别的 validationType: "${rawType}"。` +
        '支持 list | number_range | decimal | date | time | text_length | custom（另兼容直接传 Office.js 原生 rule 对象）。'
      );
    }

    if (ruleKey === 'list') {
      const items = params.listItems ?? params.items ?? params.source;
      let source = null;
      if (Array.isArray(items)) {
        const flat = items.map(v => String(v)).filter(v => v.length > 0);
        if (flat.length === 0) {
          throw new Error('[Office.js 通道] set_data_validation(validationType="list") 的 listItems 为空数组：下拉列表至少需要一项。');
        }
        source = flat.join(',');
      } else if (typeof items === 'string' && items.trim()) {
        source = stripOuterQuotes(items);
      }
      if (!source) {
        throw new Error(
          '[Office.js 通道] set_data_validation 缺少 listItems：validationType="list" 必须给出候选项数组（如 ["已通过","待复测"]）。' +
          '原实现在缺少 rule 时会先 clear() 再什么都不设，静默清空既有校验，已阻断。'
        );
      }
      return { rule: { list: { inCellDropDown: true, source } }, typeLabel: 'list', source };
    }

    if (ruleKey === 'custom') {
      const formula = params.formula || params.customFormula || params.minVal;
      if (!formula || typeof formula !== 'string' || !formula.trim().startsWith('=')) {
        throw new Error('[Office.js 通道] set_data_validation(validationType="custom") 需要 formula，且必须以 "=" 开头（如 "=ISNUMBER(A1)"）。');
      }
      return { rule: { custom: { formula: formula.trim() } }, typeLabel: 'custom' };
    }

    const opRaw = params.operator !== undefined && params.operator !== null && params.operator !== '' ? String(params.operator) : 'between';
    const operator = VALIDATION_OPERATOR_MAP[opRaw.trim().toLowerCase()];
    if (!operator) {
      throw new Error(
        `[Office.js 通道] set_data_validation 无法识别的 operator: "${opRaw}"。` +
        '支持 between | not_between | greater_than | greater_than_or_equal | less_than | less_than_or_equal | equal | not_equal。'
      );
    }

    const isTernary = operator === 'Between' || operator === 'NotBetween';
    const primary = isTernary ? (params.minVal ?? params.formula1) : (params.minVal ?? params.maxVal ?? params.formula1);
    const label = isTernary ? 'minVal' : (params.minVal !== undefined ? 'minVal' : 'maxVal');
    const body = { operator, formula1: toValidationFormula(primary, ruleKey, label) };
    if (isTernary) {
      body.formula2 = toValidationFormula(params.maxVal ?? params.formula2, ruleKey, 'maxVal');
    }

    const rule = {};
    rule[ruleKey] = body;
    return { rule, typeLabel: ruleKey };
  }

  async function handleSetDataValidation(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const address = params.address || params.range;
      if (!address) {
        throw new Error('[Office.js 通道] set_data_validation 缺少必要参数: address（如 "E5:E50"）。未执行任何修改。');
      }

      // 关键顺序（ISS-93-a）：先构造规则，构造失败就抛错 —— 此时**还没有 clear()**，既有校验不被破坏。
      const built = buildDataValidationRule(params);
      const range = sheet.getRange(address);

      range.dataValidation.clear();
      range.dataValidation.rule = built.rule;
      if (params.ignoreBlanks !== undefined) range.dataValidation.ignoreBlanks = Boolean(params.ignoreBlanks);

      // 提示与报错文案：桥接侧 normalizer 会把它们列在 unsupportedFields 里但**不下发值**，
      // 这里两处都收：顶层参数（WPS 语义）+ unsupportedFields 里列出的同名字段。
      const extraListed = Array.isArray(params.unsupportedFields) ? params.unsupportedFields : [];
      const hasValue = (n) => params[n] !== undefined && params[n] !== null && String(params[n]) !== '';
      const picked = (...names) => {
        for (const n of names) if (hasValue(n)) return String(params[n]);
        return null;
      };
      const promptTitle = picked('promptTitle');
      const promptMessage = picked('promptMessage');
      if (promptTitle || promptMessage) {
        range.dataValidation.prompt = { showPrompt: true, title: promptTitle || '', message: promptMessage || '' };
      }
      const errorTitle = picked('errorTitle');
      const errorMessage = picked('errorMessage');
      if (errorTitle || errorMessage) {
        range.dataValidation.errorAlert = {
          showAlert: true,
          style: 'Stop',
          title: errorTitle || '',
          message: errorMessage || ''
        };
      }
      const auxiliaryCandidateFields = ['promptTitle', 'promptMessage', 'errorTitle', 'errorMessage'];
      const auxiliaryApplied = auxiliaryCandidateFields.filter(hasValue);
      const auxiliaryNotApplied = extraListed.filter(f => auxiliaryCandidateFields.includes(f) && !hasValue(f));

      await context.sync();

      // 写后读回：不把"请求成功"当成"校验已设上"。
      range.load('address, dataValidation/type, dataValidation/rule, dataValidation/prompt, dataValidation/errorAlert');
      sheet.load('name');
      await context.sync();

      const dv = range.dataValidation;
      const readType = dv ? dv.type : null;
      const warnings = [];
      if (!readType || String(readType).toLowerCase() === 'none') {
        warnings.push(
          `设置后读回 dataValidation.type=${readType || 'null'}：宿主没有保留该规则，` +
          '请用 read_range/execute_script 复核，不要直接重试。'
        );
      }
      if (built.typeLabel === 'list' && built.source && dv && dv.rule && dv.rule.list && dv.rule.list.source) {
        const readSource = String(dv.rule.list.source);
        if (!readSource.includes(built.source.replace(/^=/, '').split(',')[0])) {
          warnings.push(`读回的下拉来源为 "${readSource}"，与请求 "${built.source}" 不一致（Excel 可能把逗号列表改写为区域引用）。`);
        }
      }

      return {
        success: true,
        address: range.address || address,
        sheetName: sheet.name,
        validationType: built.typeLabel,
        rule: dv ? dv.rule : null,
        readBackType: readType,
        prompt: dv ? dv.prompt : null,
        errorAlert: dv ? dv.errorAlert : null,
        auxiliaryFieldsApplied: auxiliaryApplied,
        auxiliaryFieldsNotApplied: auxiliaryNotApplied,
        warnings
      };
    });
  }

  async function handleFindReplace(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      // 网关（WPS 语义）传的是 searchRange；Office.js 侧原来只认 address（ISS-93-c）。
      const searchRange = params.searchRange || params.address || params.range;
      const range = searchRange ? sheet.getRange(searchRange) : sheet.getUsedRange();
      // 字段名兼容：网关（WPS 语义）传的是 searchQuery，Office.js 侧原来只认 text/query/findText。
      // 名字对不上 → text 落成空串 → `replaceAll("", …)` 会命中整片区域并可能破坏内容（问题台账 ISS-93）。
      const text = params.text || params.query || params.findText || params.searchQuery || "";
      if (!text) {
        throw new Error(
          "find_and_replace 缺少搜索文本（Office.js 通道识别 text / query / findText / searchQuery，当前都为空）。" +
          "空搜索串会命中整个区域并可能破坏内容，已阻断；请显式提供要查找的文本。"
        );
      }
      const replaceText = params.replaceText;
      const matchCase = !!params.matchCase;
      const matchEntireCell = !!params.matchEntireCell;
      // maxResults（网关默认 50；桥接侧 normalizer 也会做一次截断）：<=0 视为不限量。
      const rawMax = Number(params.maxResults);
      const maxResults = Number.isFinite(rawMax) && rawMax > 0 ? Math.floor(rawMax) : Infinity;

      if (replaceText !== undefined) {
        // 替换前先确认范围可读，替换后回读文本以给出可核验的读数（不谎报命中数）。
        range.load("text, address");
        await context.sync();
        const beforeText = JSON.stringify(range.text || []);
        range.replaceAll(text, replaceText, { completeMatch: matchEntireCell, matchCase: matchCase });
        await context.sync();
        range.load("text");
        await context.sync();
        const afterText = JSON.stringify(range.text || []);
        return {
          success: true,
          address: range.address,
          searchRange: searchRange || "(usedRange)",
          replacedWith: replaceText,
          changed: beforeText !== afterText,
          textChanged: beforeText !== afterText
        };
      }

      // 跨平台安全搜索：macOS Office.js 缺少 range.findAll，采用纯 JS 内存矩阵检索
      range.load("values, text, rowIndex, columnIndex, address");
      await context.sync();

      function colToLetters(c) {
        let temp = c + 1;
        let letter = "";
        while (temp > 0) {
          const mod = (temp - 1) % 26;
          letter = String.fromCharCode(65 + mod) + letter;
          temp = Math.floor((temp - mod) / 26);
        }
        return letter;
      }

      const baseRow = range.rowIndex || 0;
      const baseCol = range.columnIndex || 0;
      const matches = [];
      const queryStr = matchCase ? String(text) : String(text).toLowerCase();
      const vals = range.text || range.values || [];
      let truncated = false;

      for (let r = 0; r < vals.length; r++) {
        const row = vals[r] || [];
        for (let c = 0; c < row.length; c++) {
          const cellVal = row[c];
          if (cellVal === null || cellVal === undefined) continue;
          const str = String(cellVal);
          const targetStr = matchCase ? str : str.toLowerCase();
          const matched = matchEntireCell ? targetStr === queryStr : targetStr.includes(queryStr);
          if (matched) {
            if (matches.length >= maxResults) { truncated = true; break; }
            const cellAddr = `${colToLetters(baseCol + c)}${baseRow + r + 1}`;
            matches.push({
              address: cellAddr,
              text: str,
              row: baseRow + r + 1,
              column: baseCol + c + 1
            });
          }
        }
        if (truncated) break;
      }

      return {
        success: true,
        address: range.address,
        searchRange: searchRange || "(usedRange)",
        count: matches.length,
        truncated,
        maxResults: maxResults === Infinity ? null : maxResults,
        matches: matches
      };
    });
  }

