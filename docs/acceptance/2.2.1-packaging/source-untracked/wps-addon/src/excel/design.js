  // design.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。
  // 2. 提取原表设计语言 (Design Token) — CAP-41
  // ---------------------------------------------------------------------------
  // 原实现只读一个取样格的字体+底色，而且**读不到就编一个默认值**（"微软雅黑" / "#1E3A8A"）——
  // 调用方拿到的是臆造值而不是原表状态，正是本仓库在修的那类"看着成功其实没生效"。
  // 现在：
  //   1) 读得到什么报什么，读不到一律 null + warnings，**绝不编默认值**；
  //   2) 追加 主题色板 / 字体层级 / 表格样式 / 条件格式风格 四类设计语言；
  //   3) 宿主 API **先探测后使用**：每个候选逐个 try，成功与失败都记进 probes，
  //      整体失败进 unavailable（附宿主错误），不静默缺字段。
  //   4) 明确不碰已知危险成员：wb.Styles、SpecialCells、整列/整行范围（见 dispatch.js 的反射护栏），
  //      这些在本机未验证且可能让 WPS 崩溃，宁可不读也不赌。
  const STYLE_CENSUS_MAX_CELLS = 240;
  const STYLE_CENSUS_MAX_ROWS = 60;
  const STYLE_CENSUS_MAX_COLS = 20;
  const STYLE_BLOCK_ADDRESS = "A1:L40";
  const STYLE_CF_FORMAT_TYPES = {
    1: "cell_value", 2: "formula", 3: "color_scale", 4: "data_bar", 5: "top10",
    6: "icon_set", 8: "unique_values", 9: "text_contains", 10: "blanks", 11: "time_period",
    12: "above_average", 13: "no_blanks", 16: "duplicate_values"
  };

  /** 宿主属性安全读取：异常/undefined 都收敛到 fallback。 */
  function styleRead(fn, fallback) {
    try {
      const value = fn();
      return value === undefined ? fallback : value;
    } catch (e) {
      return fallback;
    }
  }

  /**
   * 单元格/区域的设计快照。
   * `hasFill` 用 Interior.Pattern 判定（xlNone = -4142）：无填充时 Interior.Color 仍返回白色
   * （16777215），只看 Color 会把"白底=无填充"当成"刻意设了白色"，所以两者都给。
   */
  function readDesignSnapshot(rng) {
    const pattern = styleRead(() => Number(rng.Interior.Pattern), null);
    const backgroundColor = excelColorToHex(styleRead(() => rng.Interior.Color, null));
    const hasFill = pattern === null ? null : pattern !== -4142;
    return {
      address: styleRead(() => rng.Address(), null),
      fontName: styleRead(() => rng.Font.Name, null),
      fontSize: styleRead(() => Number(rng.Font.Size), null),
      bold: styleRead(() => rng.Font.Bold, null),
      italic: styleRead(() => rng.Font.Italic, null),
      fontColor: excelColorToHex(styleRead(() => rng.Font.Color, null)),
      backgroundColor: backgroundColor,
      backgroundColorEffective: hasFill === false ? null : backgroundColor,
      hasFill: hasFill,
      fillPattern: pattern,
      numberFormat: styleRead(() => rng.NumberFormat, null),
      horizontalAlignment: alignmentName(styleRead(() => rng.HorizontalAlignment, null), false),
      rowHeight: styleRead(() => Number(rng.RowHeight), null),
      columnWidth: styleRead(() => Number(rng.ColumnWidth), null),
      borderBottom: styleRead(() => Number(rng.Borders.Item(9).LineStyle), null)
    };
  }

  /** 设计指纹：把"同一套字体+字色+底色+数字格式"归成一组。 */
  function designSignature(snapshot) {
    return [
      snapshot.fontName, snapshot.fontSize, snapshot.bold, snapshot.italic,
      snapshot.fontColor,
      snapshot.hasFill === false ? "no-fill" : String(snapshot.backgroundColor),
      snapshot.numberFormat
    ].join("|");
  }

  /**
   * 设计语言普查：扫描已用区域左上角的有界窗口，按设计指纹统计出现次数，
   * 给出每套样式的代表单元格与出现次数——这是"字体层级"的原始读数。
   * 有界（默认 240 格）是为了控制只读探测的宿主调用量，截断时如实上报。
   */
  function censusDesignStyles(sheet, usedRange) {
    const warnings = [];
    const rows = Math.max(1, Math.min(styleRead(() => Math.trunc(Number(usedRange.Rows.Count)) || 1, 1), STYLE_CENSUS_MAX_ROWS));
    const cols = Math.max(1, Math.min(styleRead(() => Math.trunc(Number(usedRange.Columns.Count)) || 1, 1), STYLE_CENSUS_MAX_COLS));
    const firstRow = Math.max(1, styleRead(() => Math.trunc(Number(usedRange.Row)) || 1, 1));
    const firstCol = Math.max(1, styleRead(() => Math.trunc(Number(usedRange.Column)) || 1, 1));
    const blockCells = rows * cols;
    const scanned = Math.min(blockCells, STYLE_CENSUS_MAX_CELLS);
    if (scanned < blockCells) {
      warnings.push(`设计普查只扫了已用区域左上角 ${scanned}/${blockCells} 个单元格（上限 ${STYLE_CENSUS_MAX_CELLS}），统计是抽样而非全量`);
    }

    // 一次调用取回取值矩阵（用于识别"这一格有没有内容"和取代表性文本），避免逐格读值
    let values = null;
    if (blockCells <= STYLE_CENSUS_MAX_CELLS * 4) {
      values = styleRead(() => normalize2DArray(usedRange.Value2, rows, cols), null);
    }
    const valueAt = (r, c) => {
      if (!values) return null;
      const row = values[r];
      if (Array.isArray(row)) return c < row.length ? row[c] : null;
      if (values.length === 1 && r === 0) return c === 0 ? values[0] : null;
      return null;
    };

    const groups = {};
    const order = [];
    for (let i = 0; i < scanned; i++) {
      const r = Math.floor(i / cols);
      const c = i % cols;
      const cell = styleRead(() => sheet.Cells.Item(firstRow + r, firstCol + c), null);
      if (!cell) continue;
      const snapshot = readDesignSnapshot(cell);
      const rawValue = valueAt(r, c);
      const hasContent = rawValue !== null && rawValue !== undefined && String(rawValue).trim() !== "";
      const key = designSignature(snapshot);
      if (!groups[key]) {
        groups[key] = Object.assign({}, snapshot, {
          signature: key,
          cellCount: 0,
          contentCellCount: 0,
          sampleCells: [],
          sampleText: null,
          address: undefined
        });
        order.push(key);
      }
      const group = groups[key];
      group.cellCount++;
      if (group.sampleCells.length < 3) group.sampleCells.push(snapshot.address);
      if (hasContent) {
        group.contentCellCount++;
        if (group.sampleText === null) group.sampleText = String(rawValue).slice(0, 40);
      }
    }
    const list = order.map((key) => groups[key]);
    return { rows: rows, cols: cols, firstRow: firstRow, firstCol: firstCol, scannedCells: scanned, blockCells: blockCells, truncated: scanned < blockCells, styles: list, warnings: warnings };
  }

  /**
   * 字体层级（title / header / body / caption）。
   * 这是**启发式**，判据写进 heuristic 字段，避免调用方把猜测当读数：
   *   title   = 字号最大的一组；header = "加粗且有底色"里出现最多的那一组；
   *   body    = 出现次数最多的一组；caption = 字号最小的一组（与 body 相同则为 null）。
   */
  function deriveDesignHierarchy(census) {
    const heuristic = "启发式判据：title=字号最大的一组；header=加粗且有底色中出现最多的一组；body=出现最多的一组；caption=字号最小的一组（等于 body 时为 null）";
    const groups = (census.styles || []).filter((g) => g.cellCount > 0);
    if (groups.length === 0) return { title: null, header: null, body: null, caption: null, heuristic: heuristic };
    const brief = (g, role) => ({
      role: role,
      signature: g.signature,
      sampleAddress: g.sampleCells[0] || null,
      sampleCells: g.sampleCells,
      sampleText: g.sampleText,
      cellCount: g.cellCount,
      contentCellCount: g.contentCellCount,
      fontName: g.fontName,
      fontSize: g.fontSize,
      bold: g.bold,
      italic: g.italic,
      fontColor: g.fontColor,
      backgroundColor: g.backgroundColorEffective,
      hasFill: g.hasFill,
      numberFormat: g.numberFormat,
      horizontalAlignment: g.horizontalAlignment
    });
    const bySizeDesc = groups.slice().sort((a, b) => (b.fontSize || 0) - (a.fontSize || 0) || b.contentCellCount - a.contentCellCount || b.cellCount - a.cellCount);
    const byCountDesc = groups.slice().sort((a, b) => b.contentCellCount - a.contentCellCount || b.cellCount - a.cellCount);
    const title = bySizeDesc[0];
    const body = byCountDesc[0];
    const header = groups.filter((g) => g.bold === true && g.hasFill === true)
      .sort((a, b) => b.contentCellCount - a.contentCellCount || b.cellCount - a.cellCount)[0] || null;
    const caption = bySizeDesc[bySizeDesc.length - 1];
    return {
      title: brief(title, "title"),
      header: header ? brief(header, "header") : null,
      body: brief(body, "body"),
      caption: caption && caption !== body && caption.fontSize !== body.fontSize ? brief(caption, "caption") : null,
      heuristic: heuristic
    };
  }

  /**
   * 主题色板探测。宿主对"主题色"没有统一入口，候选逐个试，成功与失败都记进 probes。
   * 已知可靠的兜底是**实际用到的颜色**（observed），它才是"原表长什么样"的直接证据。
   */
  function probeDesignPalette(app, wb, census) {
    const probes = [];
    const palette = { themeColors: null, themeColorsSource: null, workbookPalette: null, workbookPaletteSource: null, observed: [] };

    const candidates = [
      { key: "themeColors", source: "Workbook.ThemeColorScheme", read: () => wb.ThemeColorScheme },
      { key: "themeColors", source: "Workbook.Theme", read: () => wb.Theme },
      { key: "themeColors", source: "Application.Theme", read: () => app.Theme }
    ];
    candidates.forEach((candidate) => {
      if (palette.themeColors !== null) return;
      const toHex = (v) => {
        const n = Number(v);
        return Number.isFinite(n) && n >= 0 ? excelColorToHex(n) : null;
      };
      try {
        const raw = candidate.read();
        if (raw === null || raw === undefined) {
          probes.push({ target: candidate.source, ok: false, detail: "返回 " + String(raw) });
          return;
        }
        if (Array.isArray(raw)) {
          const entries = [];
          raw.slice(0, 12).forEach((item) => {
            const hex = toHex(item && item.Color !== undefined ? item.Color : item);
            if (hex) entries.push(hex);
          });
          if (entries.length === 0) {
            probes.push({ target: candidate.source, ok: false, detail: "数组存在但取不到颜色" });
            return;
          }
          palette.themeColors = entries;
          palette.themeColorsSource = candidate.source;
          probes.push({ target: candidate.source, ok: true, detail: `读到 ${entries.length} 个主题色槽` });
          return;
        }
        if (typeof raw === "object") {
          const entries = [];
          const count = styleRead(() => Math.trunc(Number(raw.Count)) || 0, 0);
          for (let i = 1; i <= Math.min(count, 12); i++) {
            const item = styleRead(() => raw.Item(i), null);
            if (item === null) continue;
            const hex = toHex(styleRead(() => Number(item.Color), null));
            entries.push(hex || String(item).slice(0, 40));
          }
          if (entries.length === 0) {
            probes.push({ target: candidate.source, ok: false, detail: "对象存在但取不到颜色条目（Count=" + count + "）" });
            return;
          }
          palette.themeColors = entries;
          palette.themeColorsSource = candidate.source;
          probes.push({ target: candidate.source, ok: true, detail: `读到 ${entries.length} 个主题色槽` });
          return;
        }
        probes.push({ target: candidate.source, ok: false, detail: "返回值不是颜色集合" });
      } catch (e) {
        probes.push({ target: candidate.source, ok: false, detail: e.message });
      }
    });

    // 56 色工作簿调色板（VBA 通用入口）。整体数组读不到就退化为按下标读。
    try {
      let colors = null;
      const toHex = (v) => {
        const n = Number(v);
        return Number.isFinite(n) && n >= 0 ? excelColorToHex(n) : null;
      };
      const whole = styleRead(() => wb.Colors, null);
      if (Array.isArray(whole)) {
        colors = whole.map(toHex);
      } else {
        const entries = [];
        for (let i = 1; i <= 56; i++) {
          const v = styleRead(() => wb.Colors(i), null);
          if (v === null) break;
          entries.push({ index: i, color: toHex(v) });
        }
        if (entries.length > 0) colors = entries;
      }
      if (colors && colors.length > 0) {
        palette.workbookPalette = colors;
        palette.workbookPaletteSource = "Workbook.Colors";
        probes.push({ target: "Workbook.Colors", ok: true, detail: `读到 ${colors.length} 个调色板颜色` });
      } else {
        probes.push({ target: "Workbook.Colors", ok: false, detail: "读取为空（宿主可能不暴露该属性）" });
      }
    } catch (e) {
      probes.push({ target: "Workbook.Colors", ok: false, detail: e.message });
    }

    // 实际用到的颜色：来自设计普查，按出现次数排序 —— 这是最可靠的"原表配色"
    const observed = {};
    (census.styles || []).forEach((style) => {
      const add = (color, kind) => {
        if (!color) return;
        const key = kind + ":" + color;
        if (!observed[key]) observed[key] = { color: color, kind: kind, cellCount: 0, sampleCells: [] };
        observed[key].cellCount += style.cellCount;
        style.sampleCells.forEach((addr) => { if (observed[key].sampleCells.length < 3) observed[key].sampleCells.push(addr); });
      };
      add(style.fontColor, "font");
      if (style.hasFill !== false) add(style.backgroundColorEffective, "fill");
    });
    palette.observed = Object.keys(observed).map((k) => observed[k]).sort((a, b) => b.cellCount - a.cellCount).slice(0, 20);
    probes.push({ target: "observed(设计普查统计)", ok: palette.observed.length > 0, detail: `统计到 ${palette.observed.length} 个实际使用的颜色` });
    return { palette: palette, probes: probes };
  }

  /** 结构化表格（ListObject）样式：与 manage_table 读回同一条已验证路径。 */
  function readTableStyleTokens(sheet) {
    let listObjects;
    try {
      listObjects = sheet.ListObjects;
    } catch (e) {
      return { ok: false, count: null, items: [], error: e.message, limit: 20 };
    }
    const count = styleRead(() => Math.trunc(Number(listObjects.Count)) || 0, 0);
    const items = [];
    for (let i = 1; i <= Math.min(count, 20); i++) {
      const lo = styleRead(() => listObjects.Item(i), null);
      if (!lo) continue;
      items.push({
        index: i,
        name: styleRead(() => String(lo.Name), null),
        range: styleRead(() => lo.Range.Address(), null),
        tableStyleName: styleRead(() => (lo.TableStyle ? String(lo.TableStyle.Name) : null), null),
        showHeaderRow: styleRead(() => Boolean(lo.ShowHeaderRow), null),
        showTotals: styleRead(() => Boolean(lo.ShowTotals), null)
      });
    }
    return { ok: true, count: count, tableCount: count, returnedTables: items.length, items: items, truncated: count > items.length, limit: 20 };
  }

  /**
   * 条件格式风格。**只扫有界区域**：已用区域、左上角 40×12 区块、以及该区块的前若干行。
   * 不碰整行/整列范围——dispatch.js 的反射护栏已取证"整表范围对象求值会令 WPS 崩溃"。
   */
  function readConditionalFormatStyleTokens(sheet, blockRows) {
    const probes = [];
    const areas = [];
    const seenAddresses = {};
    const attempts = [];
    const usedRange = styleRead(() => sheet.UsedRange, null);
    if (usedRange) attempts.push({ label: "usedRange", range: usedRange });
    const block = styleRead(() => sheet.Range(STYLE_BLOCK_ADDRESS), null);
    if (block) {
      attempts.push({ label: STYLE_BLOCK_ADDRESS, range: block });
      for (let r = 1; r <= Math.min(blockRows, 12); r++) {
        const rowRange = styleRead(() => sheet.Range("A" + r + ":L" + r), null);
        if (rowRange) attempts.push({ label: "A" + r + ":L" + r, range: rowRange });
      }
    }
    let readErrors = 0;
    attempts.forEach((attempt) => {
      let fcs = null;
      try {
        fcs = attempt.range.FormatConditions;
      } catch (e) {
        readErrors++;
        return;
      }
      const count = styleRead(() => Math.trunc(Number(fcs.Count)) || 0, 0);
      if (count <= 0) return;
      const address = styleRead(() => attempt.range.Address(), attempt.label);
      if (seenAddresses[address]) return;
      seenAddresses[address] = true;
      const rules = [];
      for (let i = 1; i <= Math.min(count, 10); i++) {
        const fc = styleRead(() => fcs.Item(i), null);
        if (!fc) continue;
        const typeCode = styleRead(() => Number(fc.Type), null);
        rules.push({
          index: i,
          type: STYLE_CF_FORMAT_TYPES[typeCode] || "unknown(" + typeCode + ")",
          typeCode: typeCode,
          operator: styleRead(() => Number(fc.Operator), null),
          formula1: styleRead(() => (fc.Formula1 === undefined ? null : fc.Formula1), null),
          fillColor: excelColorToHex(styleRead(() => Number(fc.Interior.Color), null)),
          fontColor: excelColorToHex(styleRead(() => Number(fc.Font.Color), null)),
          fontBold: styleRead(() => Boolean(fc.Font.Bold), null)
        });
      }
      if (rules.length > 0) areas.push({ address: address, probe: attempt.label, ruleCount: count, returnedRules: rules.length, rules: rules });
    });
    probes.push({
      target: "FormatConditions(有界区域)",
      ok: areas.length > 0,
      detail: areas.length > 0
        ? `在 ${areas.length} 个区域读到条件格式规则`
        : (readErrors > 0 ? `尝试的 ${attempts.length} 个区域中有 ${readErrors} 个读取失败，其余无规则` : `尝试的 ${attempts.length} 个区域都没有条件格式规则`)
    });
    return { probes: probes, areas: areas, scannedAreas: attempts.length, readErrors: readErrors };
  }

  function getStyleToken(app, params) {
    const { sheetName, sampleAddress = "A3", workbookName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const wb = sheet.Parent;
    const warnings = [];
    const unavailable = [];
    const probes = [];

    const sampleRange = styleRead(() => sheet.Range(sampleAddress), null);
    if (!sampleRange) throw new Error(`取样地址无法解析: "${sampleAddress}"`);
    const sample = readDesignSnapshot(sampleRange);
    const unreadable = [];
    if (sample.fontName === null) unreadable.push("fontName");
    if (sample.fontSize === null) unreadable.push("fontSize");
    if (sample.fontColor === null) unreadable.push("fontColor");
    if (sample.backgroundColor === null || sample.hasFill === null) unreadable.push("backgroundColor");
    if (unreadable.length > 0) {
      warnings.push(`取样格 ${sample.address || sampleAddress} 的 ${unreadable.join(" / ")} 读不到（空单元格或宿主未返回该属性）；对应字段为 **null 而不是默认值**，请勿当成原表风格`);
    }

    // 大标题：A1 若与取样格不同就一并给出，读不到就 null（旧实现在这里编了"微软雅黑"）
    const titleCell = styleRead(() => sheet.Range("A1"), null);
    const titleSnapshot = titleCell ? readDesignSnapshot(titleCell) : null;

    // 设计普查 → 字体层级
    const usedRange = styleRead(() => sheet.UsedRange, null);
    let census = null;
    let fonts = null;
    if (usedRange) {
      census = censusDesignStyles(sheet, usedRange);
      census.warnings.forEach((w) => warnings.push(w));
      fonts = deriveDesignHierarchy(census);
      probes.push({ target: "设计普查(已用区域左上角有界窗口)", ok: census.styles.length > 0, detail: `扫描 ${census.scannedCells} 格，识别出 ${census.styles.length} 套不同样式` });
    } else {
      unavailable.push("设计普查: 读不到 UsedRange");
      probes.push({ target: "设计普查(已用区域左上角有界窗口)", ok: false, detail: "读不到 UsedRange" });
    }

    // 主题色板 + 实际用色
    const paletteResult = probeDesignPalette(app, wb, census || { styles: [] });
    paletteResult.probes.forEach((p) => probes.push(p));
    if (paletteResult.palette.themeColors === null) {
      unavailable.push("themeColors: 宿主的主题色入口（Workbook.ThemeColorScheme / Workbook.Theme / Application.Theme）在本机全部探测失败，见 probes；色板请参考 workbookPalette 与 observed");
    }

    // 表格样式
    const tableStyles = readTableStyleTokens(sheet);
    probes.push({
      target: "ListObjects(结构化表格样式)",
      ok: tableStyles.ok,
      detail: tableStyles.ok ? `工作表中 ${tableStyles.tableCount} 个结构化表格` : ("读取 ListObjects 失败: " + tableStyles.error)
    });
    if (!tableStyles.ok) unavailable.push("tableStyles: 读不到 ListObjects（" + tableStyles.error + "）");

    // 条件格式风格
    const cf = readConditionalFormatStyleTokens(sheet, census ? census.rows : 10);
    cf.probes.forEach((p) => probes.push(p));
    if (cf.areas.length === 0) {
      unavailable.push("conditionalFormatStyles: 有界扫描未发现条件格式规则（工作表中仍可能存在，本工具不做全表枚举）");
    }

    return {
      success: true,
      workbookName: wb.Name,
      sheetName: sheet.Name,
      sampleAddress: sample.address || sampleAddress,
      sampledCell: sample,
      // 兼容旧字段名：值全部来自真实读回，读不到就是 null（旧实现在这里编造 "微软雅黑"/"#1E3A8A"）
      fontName: sample.fontName,
      titleFontName: titleSnapshot ? titleSnapshot.fontName : null,
      sampleFontSize: sample.fontSize,
      sampleBold: sample.bold,
      fontColor: sample.fontColor,
      headerBackgroundColor: sample.backgroundColorEffective,
      titleCell: titleSnapshot,
      fonts: fonts,
      palette: paletteResult.palette,
      census: census ? { scannedCells: census.scannedCells, blockCells: census.blockCells, truncated: census.truncated, styleGroups: census.styles.length, styles: census.styles.map((g) => ({
        signature: g.signature,
        cellCount: g.cellCount,
        contentCellCount: g.contentCellCount,
        sampleCells: g.sampleCells,
        sampleText: g.sampleText,
        fontName: g.fontName,
        fontSize: g.fontSize,
        bold: g.bold,
        fontColor: g.fontColor,
        backgroundColor: g.backgroundColorEffective,
        hasFill: g.hasFill,
        numberFormat: g.numberFormat
      })) } : null,
      tableStyles: tableStyles,
      conditionalFormatStyles: { scannedAreas: cf.scannedAreas, areaCount: cf.areas.length, areas: cf.areas },
      probes: probes,
      unavailable: unavailable,
      warnings: warnings,
      message: `已读取 [${sheet.Name}] 的设计语言：取样格 ${sample.address || sampleAddress}` +
        (fonts && fonts.body ? `，主样式 ${fonts.body.fontName || "?"} ${fonts.body.fontSize === null ? "?" : fonts.body.fontSize}pt` : "") +
        `；${probes.filter((p) => p.ok).length}/${probes.length} 项探测成功` +
        (unavailable.length ? `，${unavailable.length} 项不可用（见 unavailable/probes）` : "")
    };
  }
