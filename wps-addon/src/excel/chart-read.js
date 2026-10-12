  // chart-read.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // ── CAP-22 图表导图（clear_range 早已存在，此处不重复实现）
  //
  // 真机探测确认（WPS 12.1.28496）：
  //   `Chart.Export(路径, "PNG")` 可用且**确实落盘**（实测导出 8040 字节 PNG）。
  // 注意：outputPath 必须落在 **WPS 可写目录**内，写到 /tmp 之类读不到的位置
  // 会"调用成功但不落盘"，所以桥接侧会再用 existsSync 核对一次。

  /** 把图表导出为图片文件。 */
  function exportChartImage(app, params) {
    const { sheetName, workbookName, chartName, chartIndex, outputPath, format = "PNG" } = params || {};
    if (!outputPath) throw new Error("缺少必要参数: outputPath（必须落在 WPS 可写目录内）");
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    const charts = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) {
      const sh = shapes.Item(i);
      let isChart = false;
      try { isChart = Boolean(sh.HasChart); } catch (e) {}
      if (isChart) charts.push(sh);
    }
    if (!charts.length) throw new Error(`工作表 [${sheet.Name}] 上没有图表`);

    let target = null;
    if (chartIndex !== undefined && chartIndex !== null && chartIndex !== "") {
      const idx = Number(chartIndex);
      if (!(idx >= 1) || idx > charts.length) throw new Error(`图表序号 ${idx} 越界（共 ${charts.length} 个）`);
      target = charts[idx - 1];
    } else if (chartName) {
      for (const sh of charts) { try { if (String(sh.Name) === String(chartName)) { target = sh; break; } } catch (e) {} }
      if (!target) {
        const names = charts.map(s => { try { return String(s.Name); } catch (e) { return "?"; } });
        throw new Error(`找不到名为 "${chartName}" 的图表。现有图表：${names.join(" / ")}`);
      }
    } else if (charts.length === 1) {
      target = charts[0];
    } else {
      const names = charts.map(s => { try { return String(s.Name); } catch (e) { return "?"; } });
      throw new Error(`工作表上有 ${charts.length} 个图表，请用 chartName 或 chartIndex 指定。现有：${names.join(" / ")}`);
    }

    const fmt = String(format).toUpperCase();
    const ALLOWED = ["PNG", "JPG", "JPEG", "GIF", "BMP"];
    if (ALLOWED.indexOf(fmt) < 0) throw new Error(`不支持的格式: ${format}（可用 ${ALLOWED.join(" / ")}）`);
    const exportFmt = fmt === "JPEG" ? "JPG" : fmt;

    target.Chart.Export(String(outputPath), exportFmt);

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      chartName: (() => { try { return String(target.Name); } catch (e) { return null; } })(),
      outputPath: String(outputPath),
      format: exportFmt,
      // 宿主侧无法探测文件系统，落盘由**桥接侧**用 existsSync 核对后回填 fileWritten/fileSizeBytes
      fileWritten: null,
      warnings: [],
      message: `已调用图表导出 → ${outputPath}（${exportFmt}）`
    };
  }

  function getChartOrdinal(sheet, shapeName) {
    let ordinal = 0;
    for (let i = 1; i <= sheet.Shapes.Count; i++) {
      const shape = sheet.Shapes.Item(i);
      if (!shape.HasChart) continue;
      ordinal++;
      if (shape.Name === shapeName) return ordinal;
    }
    return null;
  }

  function getCharts(app, params) {
    const { sheetName, workbookName, shapeName, chartIndex, chartTitle, detail = false } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const charts = [];
    let ordinal = 0;
    // ChartType 数字 → 可读枚举名（问题台账 ISS-23：detail 原来只给数字，调用方没法判类型）
    const CHART_TYPE_NAMES = {
      1: "xlArea", 4: "xlLine", 5: "xlPie", 51: "xlColumnClustered", 52: "xlColumnStacked",
      57: "xlBarClustered", 58: "xlBarStacked", 65: "xlLineMarkers", 68: "xlBarOfPie",
      122: "xlPareto", "-4120": "xlDoughnut", "-4169": "xlXYScatter"
    };

    for (let i = 1; i <= sheet.Shapes.Count; i++) {
      const shape = sheet.Shapes.Item(i);
      if (!shape.HasChart) continue;
      ordinal++;
      const chart = shape.Chart;
      const title = safeRead(() => chart.HasTitle && chart.ChartTitle ? chart.ChartTitle.Text : "", "");
      if (shapeName && shape.Name !== shapeName) continue;
      if (chartIndex && ordinal !== Number(chartIndex)) continue;
      if (chartTitle && String(title).indexOf(chartTitle) < 0) continue;

      const position = { left: safeRead(() => shape.Left, null), top: safeRead(() => shape.Top, null) };
      const item = {
        chartIndex: ordinal,
        shapeName: shape.Name,
        title,
        chartType: safeRead(() => chart.ChartType, null),
        chartTypeName: CHART_TYPE_NAMES[safeRead(() => chart.ChartType, null)] || null,
        hasLegend: !!safeRead(() => chart.HasLegend, false),
        left: position.left,
        top: position.top,
        width: safeRead(() => shape.Width, null),
        height: safeRead(() => shape.Height, null),
        // leftCell 是"形状左上角所在单元格"，**不是唯一锚点**：多张图叠在同一像素位时会报同一个值
        // （问题台账 ISS-23）。额外给出 isDefaultPosition，避免被当成唯一定位依据。
        leftCell: safeRead(() => shape.TopLeftCell.Address(), null),
        isDefaultPosition: Math.abs(Number(position.left) - 360) < 2 && Math.abs(Number(position.top) - 40) < 2,
        seriesCount: safeRead(() => chart.SeriesCollection().Count, 0)
      };

      if (detail) {
        const series = [];
        const seriesCollection = safeRead(() => chart.SeriesCollection(), null);
        const seriesCount = seriesCollection ? safeRead(() => seriesCollection.Count, 0) : 0;
        for (let s = 1; s <= seriesCount; s++) {
          const current = seriesCollection.Item(s);
          series.push({
            seriesIndex: s,
            name: safeRead(() => current.Name, null),
            formula: safeRead(() => current.Formula, null),
            smooth: safeRead(() => current.Smooth, null),
            lineColor: excelColorToHex(safeRead(() => current.Format.Line.ForeColor.RGB, null)),
            fillColor: excelColorToHex(safeRead(() => current.Format.Fill.ForeColor.RGB, null))
          });
        }
        item.series = series;
        item.yAxis = safeRead(() => {
          const axis = chart.Axes(2, 1);
          return {
            minimumScale: axis.MinimumScale,
            maximumScale: axis.MaximumScale,
            majorUnit: axis.MajorUnit,
            numberFormat: axis.TickLabels.NumberFormat,
            title: axis.HasTitle && axis.AxisTitle ? axis.AxisTitle.Text : ""
          };
        }, null);
      }
      charts.push(item);
    }

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      totalChartCount: ordinal,
      chartCount: charts.length,
      detail: !!detail,
      charts
    };
  }
