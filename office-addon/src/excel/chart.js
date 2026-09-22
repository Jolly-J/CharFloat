// ── 模块: src/excel/chart.js — 图表 Chart 与工作表预览渲染 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1195-1662 行（原样搬迁，未改写）。
  // 8. 图表

  // ── ISS-96（高，伪渲染）───────────────
  // 旧实现的 capture_sheet_preview 在"无原生图表"时用 Canvas 按 cellW=110 / rowH=28 硬编码
  // 合成一张"看起来像表格"的图，并返回 success:true + "高保真渲染图" —— AI 的视觉自查会据此
  // 得出与真实文件不符的结论。**这比报错更危险：报错会让人停下来，假图不会。**
  // 现在的契约只有两条：
  //   ① 真实渲染路径：只有 Office.js 的 chart.getImage()（宿主原生导出）算真实渲染；
  //   ② 其余一律**抛错**，并在错误里写清替代路径（host=wps 的截图能力 / 在 Excel 里自行截图）。
  // 绝不再返回任何合成图、示意图或"已完成排版自检"之类无法核验的措辞。
  const CHART_NAME_PATTERN = /^(?:chart|图表|图)\s*(\d+)$/i;

  /** 图表名归一：把 "图表 2" / "chart2" 与宿主返回的 "Chart 2" 视为同一个。 */
  function normalizeChartName(raw) {
    const s = String(raw === undefined || raw === null ? '' : raw).trim().toLowerCase().replace(/\s+/g, '');
    if (!s) return '';
    const m = CHART_NAME_PATTERN.exec(s);
    if (m) return `chart${m[1]}`;
    return s;
  }

  /** 按名称（兼容 "Chart 2"/"图表 2"）或 id 查找图表。 */
  function findChartByName(items, wanted) {
    const target = normalizeChartName(wanted);
    if (!target) return null;
    return (items || []).find(c => normalizeChartName(c.name) === target || String(c.id || '').toLowerCase() === String(wanted).trim().toLowerCase()) || null;
  }

  function chartNamesOf(items) {
    return (items || []).map(c => {
      const title = c.title && c.title.text ? `「${c.title.text}」` : '';
      return `${c.name}${title}`;
    });
  }

  /** 渲染目标与尺寸归一，任一非法值都抛错（不再静默用默认值顶替用户的请求）。 */
  function readRenderSize(params) {
    const rawW = params.width === undefined || params.width === null || params.width === '' ? 800 : Number(params.width);
    const rawH = params.height === undefined || params.height === null || params.height === '' ? 450 : Number(params.height);
    if (!Number.isFinite(rawW) || !Number.isFinite(rawH) || rawW < 100 || rawH < 100 || rawW > 4000 || rawH > 4000) {
      throw new Error(`[Office.js 通道] capture_sheet_preview 的 width/height 非法（收到 ${params.width}×${params.height}）：需为 100..4000 的像素值。`);
    }
    return { width: Math.round(rawW), height: Math.round(rawH) };
  }

  async function handleGetCharts(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const isDetail = !!params.detail;
      const charts = sheet.charts.load(
        isDetail
          ? "items/name, items/id, items/title/text, items/chartType, items/top, items/left, items/width, items/height, items/legend/visible, items/series/items/name"
          : "items/name, items/id, items/title/text, items/chartType, items/top, items/left, items/width, items/height"
      );
      // 错误分支里要用 sheet.name 组文案：必须一起 load，否则报「属性"name"不可用」而盖掉真正的错误。
      sheet.load("name");
      await context.sync();

      const all = charts.items.map((c, idx) => {
        const item = {
          chartIndex: idx + 1,
          name: c.name,
          shapeName: c.name,
          id: c.id,
          title: c.title ? c.title.text : "",
          chartType: c.chartType,
          top: c.top,
          left: c.left,
          width: c.width,
          height: c.height,
          hasLegend: isDetail && c.legend ? !!c.legend.visible : true,
          seriesCount: isDetail && c.series && c.series.items ? c.series.items.length : 1
        };
        if (isDetail && c.series && c.series.items) {
          item.series = c.series.items.map((s, sIdx) => ({
            seriesIndex: sIdx + 1,
            name: s.name || `系列 ${sIdx + 1}`
          }));
        }
        return item;
      });

      // ISS-93-d：网关传 shapeName / chartIndex / chartTitle 三个选择器，旧实现全部忽略、恒返回全表。
      // 现在按选择器过滤；选择器命中 0 张就**抛错并列出实际图表**，而不是悄悄返回别的图。
      const wantedName = params.shapeName || params.name || params.chartName;
      const wantedTitle = params.chartTitle;
      const wantedIndex = params.chartIndex === undefined || params.chartIndex === null || params.chartIndex === ''
        ? null
        : Number(params.chartIndex);
      if (wantedIndex !== null && (!Number.isFinite(wantedIndex) || wantedIndex < 1)) {
        throw new Error(`[Office.js 通道] get_charts 的 chartIndex=${params.chartIndex} 非法：需要 ≥1 的整数。`);
      }

      const hasSelector = Boolean(wantedName) || Boolean(wantedTitle) || wantedIndex !== null;
      let result = all;
      const selector = { shapeName: wantedName || null, chartTitle: wantedTitle || null, chartIndex: wantedIndex };

      if (hasSelector) {
        if (all.length === 0) {
          throw new Error(`[Office.js 通道] get_charts：工作表 [${sheet.name}] 中没有任何原生图表，选择器 ${JSON.stringify(selector)} 无对象可匹配。`);
        }
        if (wantedName) {
          const nameNorm = normalizeChartName(wantedName);
          result = result.filter(c => normalizeChartName(c.name) === nameNorm || normalizeChartName(c.id) === nameNorm);
        }
        if (wantedTitle) {
          const t = String(wantedTitle).toLowerCase();
          result = result.filter(c => String(c.title || '').toLowerCase().includes(t));
        }
        if (wantedIndex !== null) {
          result = result.filter(c => c.chartIndex === wantedIndex);
        }
        if (result.length === 0) {
          throw new Error(
            `[Office.js 通道] get_charts 选择器 ${JSON.stringify(selector)} 在 [${sheet.name}] 未匹配到任何图表（未返回其它图表充数）。` +
            `当前工作表实际有 ${all.length} 张：${chartNamesOf(all).join('、')}。`
          );
        }
      }

      return {
        success: true,
        sheetName: sheet.name,
        count: result.length,
        totalCount: all.length,
        filtered: hasSelector,
        selector: hasSelector ? selector : null,
        charts: result
      };
    });
  }

  function applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight) {
    const targetCellRange = params.cellRange || params.position?.cellRange;
    const targetStartCell = params.startCell || params.position?.startCell || params.leftCell || params.position?.leftCell;
    const targetEndCell = params.endCell || params.position?.endCell;

    if (targetCellRange && targetCellRange.includes(":")) {
      const parts = targetCellRange.split(":").map(x => x.trim());
      try {
        chart.setPosition(sheet.getRange(parts[0]), sheet.getRange(parts[1]));
        return "cellRange";
      } catch (e1) {
        try { chart.setPosition(parts[0], parts[1]); return "cellRange"; } catch (e2) {}
      }
    } else if (targetStartCell && targetEndCell) {
      try {
        chart.setPosition(sheet.getRange(targetStartCell), sheet.getRange(targetEndCell));
        return "start/endCell";
      } catch (e1) {
        try { chart.setPosition(targetStartCell, targetEndCell); return "start/endCell"; } catch (e2) {}
      }
    } else if (targetStartCell) {
      try {
        chart.setPosition(sheet.getRange(targetStartCell));
        return "startCell";
      } catch (e1) {
        try { chart.setPosition(targetStartCell); return "startCell"; } catch (e2) {}
      }
    }
    if (targetLeft !== undefined) chart.left = targetLeft;
    if (targetTop !== undefined) chart.top = targetTop;
    if (targetWidth !== undefined) chart.width = targetWidth;
    if (targetHeight !== undefined) chart.height = targetHeight;
    return "pixels";
  }

  async function handleCreateChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const sourceRange = params.dataRange || params.sourceAddress;
      const extraRanges = Array.isArray(params.dataRanges) ? params.dataRanges.filter(r => typeof r === 'string' && r.trim()) : [];
      if (!sourceRange && extraRanges.length === 0) {
        throw new Error('[Office.js 通道] add_chart 缺少数据源：请提供 dataRange（如 "A4:E19"）或 dataRanges（多段区域数组）。');
      }
      const source = sourceRange ? sheet.getRange(sourceRange) : sheet.getRange(extraRanges[0]);

      const targetTitle = params.title || "";
      const targetLeft = params.left !== undefined ? Number(params.left) : (params.position?.left ? Number(params.position.left) : 350);
      const targetTop = params.top !== undefined ? Number(params.top) : (params.position?.top ? Number(params.position.top) : 20);
      const targetWidth = params.width !== undefined ? Number(params.width) : (params.position?.width ? Number(params.position.width) : 480);
      const targetHeight = params.height !== undefined ? Number(params.height) : (params.position?.height ? Number(params.position.height) : 280);

      // 处理 replaceExisting：若开启，先清理重叠位置或同名旧图，彻底避免图表堆叠
      const replaceExisting = params.replaceExisting !== false;
      const warnings = [];
      const removedCharts = [];
      if (replaceExisting) {
        try {
          const existingCharts = sheet.charts.load("items/name, items/id, items/title/text, items/left, items/top");
          await context.sync();
          for (const c of existingCharts.items) {
            const titleMatch = targetTitle && c.title && c.title.text === targetTitle;
            const posMatch = Math.abs(c.left - targetLeft) < 40 && Math.abs(c.top - targetTop) < 40;
            if (titleMatch || posMatch) {
              removedCharts.push(c.name);
              c.delete();
            }
          }
          await context.sync();
        } catch (cleanErr) {
          warnings.push(`清理已有图表时出错（未阻断建图）：${cleanErr && cleanErr.message ? cleanErr.message : cleanErr}`);
        }
      }

      const typeMap = {
        column: "ColumnClustered",
        column_clustered: "ColumnClustered",
        columnclustered: "ColumnClustered",
        clustered_column: "ColumnClustered",
        bar: "BarClustered",
        bar_clustered: "BarClustered",
        barclustered: "BarClustered",
        clustered_bar: "BarClustered",
        line: "Line",
        line_marker: "LineMarkers",
        pie: "Pie",
        doughnut: "Doughnut",
        donut: "Doughnut",
        area: "Area",
        scatter: "Scatter",
        xy_scatter: "Scatter"
      };
      const rawType = String(params.chartType || "ColumnClustered").trim().toLowerCase().replace(/[-\s]/g, '_');
      // ISS-17 同类问题的防线：枚举外的类型**报错**，不静默降级。
      // pareto 在 Office.js 没有对应 ChartType —— 明确拒绝并给出可执行替代，绝不悄悄建成柱状图。
      if (rawType === 'pareto') {
        throw new Error(
          '[Office.js 通道] add_chart 不支持 chartType="pareto"：Office.js 没有 Pareto 图表类型，' +
          '不会静默降级成柱状图。替代路径：① 用 chartType="column_clustered" 建柱状图，另加一列累计占比系列（或改用 host="wps" 的 add_chart）；' +
          '② 需要柏拉图外观时，自行对数据降序排序后再建图。'
        );
      }
      const chartType = typeMap[rawType];
      if (!chartType) {
        throw new Error(
          `[Office.js 通道] add_chart 无法识别的 chartType: "${params.chartType}"。` +
          '支持 line | column | column_clustered | bar | bar_clustered | pie | doughnut | area | scatter。'
        );
      }

      const chart = sheet.charts.add(chartType, source, params.seriesBy || "Auto");
      const positionMode = applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight);

      if (targetTitle) chart.title.text = targetTitle;

      let legendApplied = null;
      if (params.hasLegend !== undefined) {
        chart.legend.visible = Boolean(params.hasLegend);
        legendApplied = Boolean(params.hasLegend);
      }

      // ── ISS-93-e：dataRanges（多段数据源）────────────────
      // 旧实现完全忽略 dataRanges；这里逐段追加，并在读回后核对系列数，不一致写入 warnings。
      let dataRangesApplied = [sourceRange || extraRanges[0]];
      if (extraRanges.length > 0) {
        dataRangesApplied = extraRanges.slice();
        try {
          for (const addr of extraRanges) {
            chart.setData(sheet.getRange(addr), params.seriesBy || "Auto");
          }
          await context.sync();
        } catch (multiErr) {
          warnings.push(
            `dataRanges 多段数据源设置失败（已保留首段）：${multiErr && multiErr.message ? multiErr.message : multiErr}。` +
            '可改用一段连续区域，或改用 host="wps" 的 add_chart。'
          );
        }
      }

      const seriesList = chart.series.load("items");
      await context.sync();
      const seriesItems = seriesList.items || [];
      if (extraRanges.length > 0 && seriesItems.length < extraRanges.length) {
        warnings.push(`请求 dataRanges 共 ${extraRanges.length} 段，宿主读回只有 ${seriesItems.length} 个系列（多段数据源可能未全部生效）。`);
      }

      // ── 系列着色 ────────────────────────────────
      let colors = params.seriesColors;
      if (typeof colors === "string") colors = [colors];
      if (!colors && params.seriesColor) colors = [params.seriesColor];
      if (!colors && params.color) colors = [params.color];
      if (!Array.isArray(colors)) colors = null;

      const isPieLike = rawType.includes("pie") || rawType.includes("doughnut") || rawType.includes("donut");
      try {
        if (isPieLike) {
          if (seriesItems.length > 0) {
            const points = seriesItems[0].points.load("items");
            await context.sync();
            if (colors && colors.length > 0) {
              // 饼/环图的"系列颜色"实际是逐点颜色：只给一个颜色就整圈同色，给 N 个就按点顺序取用。
              for (let pIdx = 0; pIdx < points.items.length; pIdx++) {
                points.items[pIdx].format.fill.setSolidColor(colors[pIdx % colors.length]);
              }
            } else {
              const piePalette = ["#046A38", "#00A854", "#2CFF73", "#52C41A", "#A3D4B6", "#145A32", "#7DCEA0"];
              for (let pIdx = 0; pIdx < points.items.length; pIdx++) {
                points.items[pIdx].format.fill.setSolidColor(piePalette[pIdx % piePalette.length]);
              }
            }
          }
        } else if (colors && colors.length > 0) {
          for (let sIdx = 0; sIdx < seriesItems.length; sIdx++) {
            seriesItems[sIdx].format.fill.setSolidColor(colors[sIdx % colors.length]);
          }
        }
      } catch (colorErr) {
        warnings.push(`图表系列颜色设置失败：${colorErr && colorErr.message ? colorErr.message : colorErr}`);
      }

      // ── ISS-93-e：hasDataLabels / smoothLine / yAxis / seriesSettings ──
      const labelsRequested = params.hasDataLabels;
      if (labelsRequested !== undefined) {
        try {
          const labels = chart.dataLabels.load("showValue");
          await context.sync();
          labels.showValue = Boolean(labelsRequested);
          await context.sync();
          labels.load("showValue");
          await context.sync();
          if (Boolean(labels.showValue) !== Boolean(labelsRequested)) {
            warnings.push(`hasDataLabels 请求 ${Boolean(labelsRequested)}，读回 ${Boolean(labels.showValue)}（宿主未接受）。`);
          }
        } catch (labelErr) {
          warnings.push(`hasDataLabels 设置失败：${labelErr && labelErr.message ? labelErr.message : labelErr}`);
        }
      }

      const smoothRequested = params.smoothLine;
      const isLineLike = /line|scatter/i.test(rawType);
      if (smoothRequested !== undefined && !isLineLike) {
        warnings.push(`smoothLine 仅对折线/散点图有意义，当前 chartType=${rawType}，已忽略。`);
      } else if (smoothRequested !== undefined) {
        let smoothApplied = 0;
        for (const s of seriesItems) {
          try { s.smooth = Boolean(smoothRequested); smoothApplied++; } catch (smoothErr) { /* 单系列不支持时跳过 */ }
        }
        if (smoothApplied === 0) {
          warnings.push('smoothLine 设置失败：宿主未接受任何系列的平滑属性。');
        } else {
          try {
            await context.sync();
          } catch (smoothSyncErr) {
            warnings.push(`smoothLine 写入未生效：${smoothSyncErr && smoothSyncErr.message ? smoothSyncErr.message : smoothSyncErr}`);
          }
        }
      }

      const seriesSettings = Array.isArray(params.seriesSettings) ? params.seriesSettings : [];
      for (const setting of seriesSettings) {
        const idx = Number(setting && setting.seriesIndex);
        if (!Number.isFinite(idx) || idx < 1) {
          throw new Error('[Office.js 通道] add_chart 的 seriesSettings.seriesIndex 非法：需要 ≥1 的整数（从 1 开始）。');
        }
        if (idx > seriesItems.length) {
          throw new Error(
            `[Office.js 通道] add_chart 的 seriesSettings.seriesIndex=${idx} 越界：本图只有 ${seriesItems.length} 个系列。` +
            '原实现会静默跳过，现改为显式报错（已建图不会回滚，请按需 delete_chart 后重建）。'
          );
        }
        const target = seriesItems[idx - 1];
        if (setting.color) {
          try { target.format.fill.setSolidColor(setting.color); } catch (cErr) { warnings.push(`seriesSettings[${idx}].color 设置失败：${cErr && cErr.message ? cErr.message : cErr}`); }
        }
        if (setting.smooth !== undefined) {
          try { target.smooth = Boolean(setting.smooth); } catch (smErr) { warnings.push(`seriesSettings[${idx}].smooth 设置失败：${smErr && smErr.message ? smErr.message : smErr}`); }
        }
      }
      if (seriesSettings.length > 0) {
        try { await context.sync(); } catch (ssErr) { warnings.push(`seriesSettings 写入未生效：${ssErr && ssErr.message ? ssErr.message : ssErr}`); }
      }

      const yAxisParam = params.yAxis || {};
      const yAxisRequested = ['min', 'max', 'step', 'numberFormat', 'title'].some(k => yAxisParam[k] !== undefined && yAxisParam[k] !== null && yAxisParam[k] !== '');
      let yAxisApplied = null;
      if (yAxisRequested) {
        try {
          const va = chart.axes.valueAxis;
          if (yAxisParam.min !== undefined && yAxisParam.min !== null) va.minimum = Number(yAxisParam.min);
          if (yAxisParam.max !== undefined && yAxisParam.max !== null) va.maximum = Number(yAxisParam.max);
          if (yAxisParam.step !== undefined && yAxisParam.step !== null) va.majorUnit = Number(yAxisParam.step);
          if (yAxisParam.numberFormat) va.numberFormat = String(yAxisParam.numberFormat);
          if (yAxisParam.title) va.title.text = String(yAxisParam.title);
          await context.sync();
          if (yAxisParam.numberFormat) va.format.numberFormat = String(yAxisParam.numberFormat);
          if (yAxisParam.title) va.title.text = String(yAxisParam.title);
          await context.sync();
          va.load("minimum, maximum, majorUnit, numberFormat, title/text");
          await context.sync();
          yAxisApplied = {
            minimum: va.minimum,
            maximum: va.maximum,
            majorUnit: va.majorUnit,
            numberFormat: va.numberFormat,
            title: va.title ? va.title.text : null
          };
          const mismatches = [];
          if (yAxisParam.min !== undefined && yAxisParam.min !== null && Math.abs(Number(yAxisApplied.minimum) - Number(yAxisParam.min)) > 1e-9) mismatches.push(`min 请求 ${yAxisParam.min} 读回 ${yAxisApplied.minimum}`);
          if (yAxisParam.max !== undefined && yAxisParam.max !== null && Math.abs(Number(yAxisApplied.maximum) - Number(yAxisParam.max)) > 1e-9) mismatches.push(`max 请求 ${yAxisParam.max} 读回 ${yAxisApplied.maximum}`);
          if (mismatches.length > 0) warnings.push(`yAxis 未完全生效：${mismatches.join('；')}`);
        } catch (axisErr) {
          warnings.push(`yAxis 设置失败：${axisErr && axisErr.message ? axisErr.message : axisErr}`);
        }
      }

      chart.load("name, id, top, left, width, height");
      sheet.load("name");
      await context.sync();

      return {
        success: true,
        id: chart.id,
        name: chart.name,
        shapeName: chart.name,
        chartIndex: 1,
        title: targetTitle,
        chartType: rawType,
        dataRange: sourceRange || extraRanges[0],
        dataRangesApplied,
        sheetName: sheet.name,
        left: chart.left !== undefined ? chart.left : targetLeft,
        top: chart.top !== undefined ? chart.top : targetTop,
        width: chart.width !== undefined ? chart.width : targetWidth,
        height: chart.height !== undefined ? chart.height : targetHeight,
        positionMode,
        hasLegendApplied: legendApplied,
        yAxisApplied,
        seriesCount: seriesItems.length,
        seriesSettingsApplied: seriesSettings.length,
        replaceExisting: replaceExisting,
        removedCharts,
        warnings,
        readBackNote: '左侧/顶部/宽高为宿主读回值；用 get_charts(detail=true) 复核标题、系列数与类型。',
        message: `已成功在 [${sheet.name}] 创建 ${rawType} 原生图表，数据源为 ${dataRangesApplied.join(' + ')}`
      };
    });
  }

  async function handleDeleteChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      if (params.clearAll) {
        const charts = sheet.charts.load("items");
        await context.sync();
        const count = charts.items.length;
        for (const c of charts.items) {
          c.delete();
        }
        await context.sync();
        return { success: true, count, message: `已成功清空当前工作表中的全部 ${count} 个图表` };
      }
      const wanted = params.chartName || params.name || params.id || params.shapeName;
      if (!wanted) {
        throw new Error('[Office.js 通道] delete_chart 需要 chartName（或 shapeName/id），否则不知道删哪一张；如需清空全部请显式传 clearAll=true。');
      }
      const chart = sheet.charts.getItem(wanted);
      chart.delete();
      await context.sync();
      return { success: true, deleted: wanted, message: `图表 [${wanted}] 已成功删除` };
    });
  }

  async function handleUpdateChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const wanted = params.name || params.chartName || params.id || params.shapeName;
      if (!wanted) {
        throw new Error('[Office.js 通道] update_chart 需要 name/chartName/shapeName/id 指定目标图表。');
      }
      const chart = sheet.charts.getItem(wanted);
      if (params.title) chart.title.text = params.title;
      if (params.legendPosition) chart.legend.position = params.legendPosition;
      const targetLeft = params.left !== undefined ? Number(params.left) : (params.position?.left ? Number(params.position.left) : undefined);
      const targetTop = params.top !== undefined ? Number(params.top) : (params.position?.top ? Number(params.position.top) : undefined);
      const targetWidth = params.width !== undefined ? Number(params.width) : (params.position?.width ? Number(params.position.width) : undefined);
      const targetHeight = params.height !== undefined ? Number(params.height) : (params.position?.height ? Number(params.position.height) : undefined);

      applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight);
      await context.sync();
      return { success: true, name: wanted, message: `图表 [${wanted}] 已成功更新位置与配置` };
    });
  }

  async function handleExportChartImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const wanted = params.name || params.chartName || params.shapeName;
      if (!wanted) {
        throw new Error('[Office.js 通道] export_chart_image 需要 chartName（或 name/shapeName）指定图表。');
      }
      const chart = sheet.charts.getItem(wanted);
      const size = readRenderSize({ width: params.width, height: params.height });
      const imageResult = chart.getImage(size.width, size.height);
      await context.sync();
      return { success: true, chartName: wanted, imageBase64: imageResult.value, imageMimeType: 'image/png', width: size.width, height: size.height, renderedBy: 'chart.getImage (Office.js 原生导出)' };
    });
  }

  /**
   * ISS-96：capture_sheet_preview —— 只做真实渲染，做不到就明确报错。
   *
   * 真实渲染的**唯一**来源是 `chart.getImage()`（宿主原生导出图表）。
   * 工作表区域截图在 Office.js 交付面上没有可用 API（macOS 桌面版既无 Range 截图，
   * 也没有可用的 Workbook 渲染导出），因此区域预览一律抛错并给出替代路径。
   */
  async function handleCaptureSheetPreview(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      const requestedMode = params?.mode === undefined || params?.mode === null || params?.mode === '' ? null : String(params.mode).toLowerCase();
      if (requestedMode && !['chart', 'sheet', 'auto'].includes(requestedMode)) {
        throw new Error(`[Office.js 通道] capture_sheet_preview 无法识别的 mode: "${params.mode}"（支持 chart | sheet | auto）。`);
      }

      const chartSelector = params?.chartName || params?.name || null;
      const address = params?.address || params?.range || null;
      if (requestedMode === 'sheet' && chartSelector) {
        throw new Error(
          `[Office.js 通道] capture_sheet_preview 参数冲突：mode="sheet" 与 chartName="${chartSelector}" 同时给出。` +
          '区域截图不受支持（见下），如需导出图表请去掉 mode 或传 mode="chart"。'
        );
      }

      const targetChartName = requestedMode === 'sheet' ? null : chartSelector;
      const wantsSheetArea = requestedMode === 'sheet' || (!targetChartName && Boolean(address));
      // 错误分支要用 sheet.name 组文案（本通道的报错文案本身就是交付物）：先 load。
      sheet.load("name");

      if (targetChartName) {
        const charts = sheet.charts.load("items/name, items/id, items/title/text, items/width, items/height");
        await context.sync();
        if (charts.items.length === 0) {
          throw new Error(`[Office.js 通道] capture_sheet_preview 无法渲染：工作表 [${sheet.name}] 没有任何原生图表，而你请求的是图表 "${targetChartName}"。`);
        }
        const target = findChartByName(charts.items, targetChartName);
        if (!target) {
          throw new Error(
            `[Office.js 通道] capture_sheet_preview 未找到图表 "${targetChartName}"（按名称或 id 精确匹配，兼容 "图表 2" 与 "Chart 2"）。` +
            `[${sheet.name}] 现有图表：${chartNamesOf(charts.items).join('、')}。`
          );
        }
        const size = readRenderSize(params || {});
        const imgResult = target.getImage(size.width, size.height);
        await context.sync();
        if (!imgResult || !imgResult.value) {
          throw new Error(`[Office.js 通道] capture_sheet_preview：宿主对图表 [${target.name}] 的 getImage 未返回图像数据，无法提供真实渲染图。`);
        }
        const title = target.title && target.title.text ? target.title.text : '';
        return {
          success: true,
          kind: 'chart',
          renderedBy: 'chart.getImage (Office.js 原生导出)',
          workbookName: params?.workbookName || null,
          sheetName: sheet.name,
          address: target.name,
          chartName: target.name,
          chartTitle: title,
          imageBase64: imgResult.value,
          imageMimeType: "image/png",
          width: size.width,
          height: size.height,
          message: `已导出 [${sheet.name}] 图表 [${target.name}] 的原生渲染图（${size.width}×${size.height}，由 Office.js chart.getImage 生成，非合成图）。`
        };
      }

      // 走到这里说明请求的是工作表区域（或既没给图表也没给区域）。
      const scope = wantsSheetArea ? `区域 ${address}` : '工作表已用区域';
      throw new Error(
        `[Office.js 通道] capture_sheet_preview 不支持真实渲染：无法为 [${sheet.name}] 的${scope}生成截图。` +
        'Office.js 交付面没有工作表/区域截图 API（只有图表有 chart.getImage），因此本通道不提供该图，' +
        '也不会用 Canvas 合成"示意图"充数（旧版本会返回假渲染图，已移除）。' +
        '替代路径：① 改用 host="wps" 的 wps_capture_sheet_preview（WPS 原生渲染，可截指定区域）；' +
        '② 在 Excel 里自行截图（选区后 Shift+Cmd+4/Ctrl+C 复制为图片）再把图交给 AI；' +
        '③ 需要"读回核对"时改用 read_range / get_range_styles / get_charts(detail=true) 做数据与格式自查。'
      );
    });
  }
