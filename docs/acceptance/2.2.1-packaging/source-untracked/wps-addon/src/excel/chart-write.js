  // chart-write.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // ── CAP-21 图表更新（wps_update_chart）
  //
  // 此前 `wps_update_chart` 只在路由表里声明，WPS 加载项**没有 RPC 分支**，
  // 调用会被明确拒绝（宿主未实现）。真机确认宿主 API 可用：
  // `shape.Chart.ChartTitle.Text` 可读可写、`ChartType` 可改、系列可读。
  /** 就地更新已有图表：标题、图表类型、图例、数据标签。写后逐项读回核对。 */
  function updateChart(app, params) {
    const { sheetName, workbookName, chartName, chartIndex, title, chartType, hasLegend, showDataLabels,
            fontName, legendPosition, dataLabelColorMatchesSeries, position } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    // 定位图表形状（HasChart 为真）
    let target = null;
    const charts = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) {
      const sh = shapes.Item(i);
      let isChart = false;
      try { isChart = Boolean(sh.HasChart); } catch (e) {}
      if (!isChart) continue;
      charts.push(sh);
    }
    if (!charts.length) throw new Error(`工作表 [${sheet.Name}] 上没有图表`);
    if (chartName || chartIndex !== undefined) {
      if (chartIndex !== undefined) {
        const idx = Number(chartIndex);
        if (idx < 1 || idx > charts.length) throw new Error(`图表序号 ${idx} 越界（共 ${charts.length} 个）`);
        target = charts[idx - 1];
      } else {
        for (const sh of charts) { try { if (String(sh.Name) === String(chartName)) { target = sh; break; } } catch (e) {} }
        if (!target) throw new Error(`找不到名为 "${chartName}" 的图表。现有图表：${charts.map(s => { try { return String(s.Name); } catch (e) { return "?"; } }).join(" / ")}`);
      }
    } else if (charts.length === 1) {
      target = charts[0];
    } else {
      throw new Error(`工作表上有 ${charts.length} 个图表，请用 chartName 或 chartIndex 指定。现有：${charts.map(s => { try { return String(s.Name); } catch (e) { return "?"; } }).join(" / ")}`);
    }

    const ch = target.Chart;
    const warnings = [];
    const applied = {};

    if (title !== undefined) {
      try { ch.HasTitle = true; ch.ChartTitle.Text = String(title); applied.title = String(title); }
      catch (e) { warnings.push(`设置标题失败: ${e.message}`); }
    }
    if (chartType !== undefined) {
      // 常用 xlChartType：柱状簇状=51 / 折线=4 / 饼图=5 / 条形簇状=57 / 散点=75 / 面积=1
      const TYPES = { column: 51, column_clustered: 51, bar: 57, bar_clustered: 57, line: 4, pie: 5, scatter: 75, area: 1 };
      const code = typeof chartType === "number" ? chartType : TYPES[String(chartType).toLowerCase()];
      if (code === undefined) { warnings.push(`不认识的 chartType: ${chartType}`); }
      else { try { ch.ChartType = code; applied.chartType = code; } catch (e) { warnings.push(`设置图表类型失败: ${e.message}`); } }
    }
    if (hasLegend !== undefined) {
      try { ch.HasLegend = Boolean(hasLegend); applied.hasLegend = Boolean(hasLegend); } catch (e) { warnings.push(`设置图例失败: ${e.message}`); }
    }
    // 图表字体：标题/图例/坐标轴/数据标签统一换成指定字体
    // （宿主默认走宋体，中英文混排观感差）
    if (fontName) {
      const f = String(fontName);
      const setFont = (obj, label) => { try { obj.Font.Name = f; } catch (e) { warnings.push(`设置${label}字体失败: ${e.message}`); } };
      try { setFont(ch.ChartArea, "图表区"); } catch (e) {}
      try { if (ch.HasTitle) setFont(ch.ChartTitle, "标题"); } catch (e) {}
      try { if (ch.HasLegend) setFont(ch.Legend, "图例"); } catch (e) {}
      // 坐标轴字体要走 **TickLabels**：WPS 里  是 undefined（设不上），
      // 刻度标签才挂得住 Font。真机探明：Axes().Item(n).TickLabels.Font 可读可写。
      const setTickFont = (idx) => {
        try {
          const ax = ch.Axes().Item(idx);
          ax.TickLabels.Font.Name = f;
          return ax.TickLabels.Font.Name === f;
        } catch (e) { return false; }
      };
      if (!setTickFont(1)) warnings.push("设置分类轴刻度字体失败");
      if (!setTickFont(2)) warnings.push("设置数值轴刻度字体失败");
      applied.fontName = f;
    }

    // 图例位置：xlLegendPositionRight=-4152 / Left=-4131 / Top=-4160 / Bottom=-4107
    if (legendPosition !== undefined && legendPosition !== null && legendPosition !== "") {
      const LP = { right: -4152, left: -4131, top: -4160, bottom: -4107 };
      const key = String(legendPosition).toLowerCase();
      const code = LP[key] !== undefined ? LP[key] : (Number.isFinite(Number(legendPosition)) ? Number(legendPosition) : undefined);
      if (code === undefined) warnings.push(`不认识的 legendPosition: ${legendPosition}（可用 right/left/top/bottom）`);
      else {
        try { ch.HasLegend = true; ch.Legend.Position = code; applied.legendPosition = key; }
        catch (e) { warnings.push(`设置图例位置失败: ${e.message}`); }
      }
    }

    // 数据标签文字颜色与所属柱/点颜色一致
    if (dataLabelColorMatchesSeries) {
      let matched = 0;
      try {
        const sc = ch.SeriesCollection();
        for (let i = 1; i <= Number(sc.Count); i++) {
          const s = sc.Item(i);
          let rgb = null;
          try { rgb = Number(s.Format.Fill.ForeColor.RGB); } catch (e) {}
          if (!Number.isFinite(rgb) || rgb < 0) { try { rgb = Number(s.Border.Color); } catch (e) {} }
          if (!Number.isFinite(rgb) || rgb < 0) continue;
          try {
            s.HasDataLabels = true;
            s.DataLabels().Font.Color = rgb;
            matched++;
          } catch (e) { warnings.push(`设置系列 ${i} 数据标签颜色失败: ${e.message}`); }
        }
      } catch (e) { warnings.push(`设置数据标签颜色失败: ${e.message}`); }
      applied.dataLabelColorMatchesSeries = matched;
      if (!matched) warnings.push("没有系列成功匹配到颜色，数据标签颜色未改");
    }

    // 移动/改尺寸（单位：磅）。支持绝对定位与相对位移——
    // 「把图表往下挪一点」这种调整本该一条指令完成，不必重建图表或跑脚本。
    if (position && typeof position === "object") {
      const moved = {};
      const setIf = (key, cur, next) => { if (Number.isFinite(Number(next))) { try { target[key] = Number(next); moved[key] = Number(next); } catch (e) { warnings.push(`设置 ${key} 失败: ${e.message}`); } } };
      setIf("Left", null, position.left);
      setIf("Top", null, position.top);
      setIf("Width", null, position.width);
      setIf("Height", null, position.height);
      // 相对位移：在当前位置基础上加减
      if (Number.isFinite(Number(position.leftDelta))) {
        try { target.Left = Number(target.Left) + Number(position.leftDelta); moved.left = Number(target.Left); } catch (e) { warnings.push(`左移失败: ${e.message}`); }
      }
      if (Number.isFinite(Number(position.topDelta))) {
        try { target.Top = Number(target.Top) + Number(position.topDelta); moved.top = Number(target.Top); } catch (e) { warnings.push(`下移失败: ${e.message}`); }
      }
      // 单元格锚点：左上角对齐到指定单元格
      if (position.leftCell) {
        try { const c = sheet.Range(String(position.leftCell)); target.Left = Number(c.Left) + 2; target.Top = Number(c.Top) + 2; moved.leftCell = String(position.leftCell); } catch (e) { warnings.push(`按单元格定位失败: ${e.message}`); }
      }
      applied.position = moved;
    }

    if (showDataLabels !== undefined) {
      // ⚠️ 不能用 `chart.ApplyDataLabels(0)` 关标签：真机实测它**关不掉**，
      // 但调用不报错——于是工具返回 success 而图表上标签仍在，属"假成功"。
      // 改为**逐个系列设置 HasDataLabels**，再读回核对，报**真实状态**而不是请求值。
      const want = Boolean(showDataLabels);
      let touched = 0;
      try {
        const sc = ch.SeriesCollection();
        const n = Number(sc.Count);
        for (let i = 1; i <= n; i++) { try { sc.Item(i).HasDataLabels = want; touched++; } catch (e) {} }
      } catch (e) { warnings.push(`设置数据标签失败: ${e.message}`); }
      if (!touched) {
        try { ch.ApplyDataLabels(want ? 2 : 0); touched = 1; } catch (e) { warnings.push(`设置数据标签失败: ${e.message}`); }
      }
      // 读回每个系列的真实状态
      const states = [];
      try {
        const sc = ch.SeriesCollection();
        for (let i = 1; i <= Number(sc.Count); i++) { try { states.push(Boolean(sc.Item(i).HasDataLabels)); } catch (e) {} }
      } catch (e) {}
      const actualAll = states.length ? states.every(Boolean) : null;
      const actualNone = states.length ? states.every(v => !v) : null;
      applied.showDataLabels = want;
      // 名字直说含义："实际状态与请求一致"，避免被读成"标签是开着的"
      applied.dataLabelsMatchRequest = want ? actualAll : actualNone;
      if (states.length && (want ? !actualAll : !actualNone)) {
        warnings.push(`数据标签**未按请求生效**：请求 ${want}，各系列实际为 [${states.join(", ")}]`);
      }
    }

    // 读回核对
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const actual = {
      name: g(() => String(target.Name), null),
      chartType: g(() => Number(ch.ChartType), null),
      title: g(() => String(ch.ChartTitle.Text), null),
      hasTitle: g(() => Boolean(ch.HasTitle), null),
      hasLegend: g(() => Boolean(ch.HasLegend), null),
      seriesCount: g(() => Number(ch.SeriesCollection().Count), null),
      // 数据标签的真实状态（逐系列读），让"写没写进去"可核对
      dataLabels: (() => {
        try {
          const sc = ch.SeriesCollection(); const out = [];
          for (let i = 1; i <= Number(sc.Count); i++) { try { out.push(Boolean(sc.Item(i).HasDataLabels)); } catch (e) {} }
          return out.length ? (out.every(Boolean) ? true : out.every(v => !v) ? false : out) : null;
        } catch (e) { return null; }
      })()
    };
    if (title !== undefined && actual.title !== String(title)) warnings.push(`标题读回为「${actual.title}」，与请求不一致`);

    return {
      success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name,
      chartName: actual.name, applied, chart: actual, warnings,
      message: `已更新图表「${actual.name}」${Object.keys(applied).length} 项${warnings.length ? `（${warnings.length} 条告警）` : ""}`
    };
  }

  // 15. 原生矢量图表生成与管理 (第二梯队 - 健壮事务版)
  function addChart(app, params) {
    const {
      sheetName,
      workbookName,
      chartType = "column_clustered",
      dataRange,
      dataRanges,
      title,
      position,
      hasLegend = true,
      hasDataLabels = false,
      smoothLine = false,
      seriesColors,
      yAxis,
      seriesSettings,
      replaceExisting = true,
      left: explicitLeft,
      top: explicitTop,
      width: explicitWidth,
      height: explicitHeight,
      startCell,
      endCell,
      cellRange
    } = params || {};

    // 1. 强制公式全局重算（彻底根治跨表引用公式未计算导致分类轴塌陷的严重时钟竞争 Bug！）
    try {
      if (app && app.Calculate) {
        app.Calculate();
      }
    } catch (calcErr) {
      log("公式重算提示: " + calcErr.message);
    }

    // 2. 解析数据源（支持非连续区域联合数据源，如 ['A4:A19', 'E4:E19']）
    let targetDataRange = dataRange;
    if (!targetDataRange && Array.isArray(dataRanges) && dataRanges.length > 0) {
      targetDataRange = dataRanges.join(",");
    }
    if (!targetDataRange) {
      throw new Error("缺少必要参数: dataRange (例如 'A4:E19') 或 dataRanges (例如 ['A4:A19', 'E4:E19'])");
    }

    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    /** 本次建图的告警集合（类型降级、单系列多色等），随返回体带出。 */
    const warnings = [];

    // 计算图表位置与尺寸
    let left = 360;
    let top = 40;
    let width = 480;
    let height = 280;

    if (position) {
      if (position.leftCell) {
        try {
          const anchorRange = sheet.Range(position.leftCell);
          left = anchorRange.Left;
          top = anchorRange.Top;
        } catch (e) {
          log("图表定位 leftCell 警告: " + e.message);
        }
      }
      if (position.width) width = Number(position.width);
      if (position.height) height = Number(position.height);
    }

    // 原实现只认 `position.leftCell/width/height`，而 startCell / cellRange / endCell 与
    // 顶层 left/top/width/height **全被静默忽略**，多图会叠在默认的 360/40（问题台账 ISS-18）。
    // 这里把它们真正接上，并让"锚点"统一走一套解析顺序：position.leftCell > startCell/cellRange > 顶层像素。
    const anchorRef = (position && position.leftCell)
      ? String(position.leftCell)
      : (startCell ? String(startCell) : (cellRange ? String(cellRange).split(":")[0] : null));
    if (anchorRef && !(position && position.leftCell)) {
      try {
        const anchorRange = sheet.Range(anchorRef.includes(":") ? anchorRef.split(":")[0] : anchorRef);
        left = anchorRange.Left;
        top = anchorRange.Top;
      } catch (e) {
        log("图表定位锚点警告: " + e.message);
      }
    }
    // 顶层像素参数直接生效（显式传入优先于上面的锚点推算）
    if (Number.isFinite(Number(explicitLeft))) left = Number(explicitLeft);
    if (Number.isFinite(Number(explicitTop))) top = Number(explicitTop);
    if (Number.isFinite(Number(explicitWidth))) width = Number(explicitWidth);
    if (Number.isFinite(Number(explicitHeight))) height = Number(explicitHeight);
    // endCell：用"锚点单元格 → endCell"的矩形尺寸作为图表宽高
    if (endCell) {
      try {
        const from = sheet.Range(anchorRef ? anchorRef.split(":")[0] : "A1");
        const to = sheet.Range(String(endCell));
        const w = to.Left + to.Width - from.Left;
        const h = to.Top + to.Height - from.Top;
        if (w > 0) width = w;
        if (h > 0) height = h;
      } catch (e) {
        log("图表 endCell 尺寸推算警告: " + e.message);
      }
    }

    // 3. 覆盖模式（如果开启 replaceExisting，先清理该锚点处的重叠旧图，杜绝废图堆叠）
    if (replaceExisting && position && position.leftCell) {
      try {
        const shapes = sheet.Shapes;
        const total = shapes.Count;
        for (let i = total; i >= 1; i--) {
          const shp = shapes.Item(i);
          if (shp.HasChart && Math.abs(shp.Left - left) < 25 && Math.abs(shp.Top - top) < 25) {
            log(`自动清理锚点 [${position.leftCell}] 处的旧图表: ${shp.Name}`);
            shp.Delete();
          }
        }
      } catch (cleanErr) {
        log("清理旧图表提示: " + cleanErr.message);
      }
    }

    // 映射 Excel 图表类型常量
    // xlLine=4, xlColumnClustered=51, xlBarClustered=57, xlPie=5, xlDoughnut=-4120, xlPareto=122,
    // xlArea=1, xlXYScatter=-4169
    const chartTypeMap = {
      line: 4,
      column: 51,
      column_clustered: 51,
      bar: 57,
      bar_clustered: 57,
      pie: 5,
      doughnut: -4120,
      pareto: 122,
      area: 1,
      scatter: -4169
    };
    // 原来未知类型会**静默退回柱状图**并返回 success（问题台账 ISS-17）：调用方拿到的图不是要的类型却毫无察觉。
    // 现在改为显式报错，并带上可用类型清单。
    if (chartTypeMap[chartType] === undefined) {
      throw new Error(
        `不支持的图表类型: ${chartType}（可用: ${Object.keys(chartTypeMap).join(" / ")}）`
      );
    }
    const xlChartType = chartTypeMap[chartType];

    let shape = null;
    try {
      if (sheet.Shapes.AddChart2) {
        shape = sheet.Shapes.AddChart2(201, xlChartType, left, top, width, height);
      } else {
        shape = sheet.Shapes.AddChart(xlChartType, left, top, width, height);
      }
    } catch (err) {
      // 降级为默认簇状柱状图
      shape = sheet.Shapes.AddChart(51, left, top, width, height);
    }

    if (!shape || !shape.Chart) {
      throw new Error("无法在当前工作表中创建原生图表对象");
    }

    // 4. 事务性配置（一旦发生异常，自动自毁半成品 Shape，绝不在工作表留下空白废图！）
    try {
      const chart = shape.Chart;
      // ⚠️ 数据源**支持跨表**（ISS-120）：宿主拒绝把带表名的字符串交给 AddChart2
      // （报 `Parameter type error source (arg 0)`），但 `chart.SetSourceData(跨表 Range 对象)`
      // **确实可用**（真机实测：跨表拿到 2 个系列）。
      // 所以这里把 `表名!A1:C3` 解析出来，从**那张表**取 Range 对象再传进去。
      let srcSheet = sheet;
      let srcAddr = String(targetDataRange);
      const m = srcAddr.match(/^\s*(?:'([^']+)'|([^!]+))!\s*(.+)$/);
      if (m) {
        const otherName = (m[1] || m[2] || "").trim();
        if (otherName) {
          // ⚠️ 本函数里**没有 `wb` 变量**（只有 `sheet`）——先前误用 `wb.Worksheets`
          // 会抛 ReferenceError，被 catch 吞掉后**静默回退到当前表**，
          // 于是拿一张空表的区域去 SetSourceData，报成 "Parameter type error source"，
          // 看起来像"宿主不支持跨表"，其实是变量名写错了。
          // 表名找不到时**必须报错**，不能回退到别的表——那会把数据源悄悄换掉。
          srcSheet = sheet.Parent.Worksheets.Item(otherName); // 找不到会抛，交给外层如实报错
          srcAddr = String(m[3]).trim();
        }
      }
      const srcRange = srcSheet.Range(srcAddr);
      chart.SetSourceData(srcRange);

      // 安全设置标题（防范 ChartTitle 空指针崩溃）
      if (title) {
        try {
          chart.HasTitle = true;
          if (chart.ChartTitle) {
            chart.ChartTitle.Text = title;
          }
        } catch (tErr) {
          try {
            if (chart.ChartTitle && chart.ChartTitle.Characters) {
              chart.ChartTitle.Characters.Text = title;
            }
          } catch (tErr2) {
            log("设置图表标题警告: " + tErr2.message);
          }
        }
      }

      if (hasLegend !== undefined) {
        chart.HasLegend = !!hasLegend;
      }

      if (hasDataLabels) {
        try {
          chart.ApplyDataLabels();
        } catch (e) {}
      }

      // 4.1 平滑线设置与系列颜色注入
      try {
        const seriesCol = chart.SeriesCollection();
        const seriesCount = seriesCol ? seriesCol.Count : 0;

        // 全局平滑线设置 (主要针对折线图)
        if (smoothLine) {
          for (let sIdx = 1; sIdx <= seriesCount; sIdx++) {
            try {
              seriesCol.Item(sIdx).Smooth = true;
            } catch (e) {}
          }
        }

        // 系列颜色注入 (seriesColors 数组)
        if (Array.isArray(seriesColors)) {
          // 单系列图表上传多个颜色，宿主会按"逐点染色"处理 → 应单色的柱图/条形图变成彩虹柱
          // （问题台账 ISS-24）。这里只应用第一个颜色，并把"其余被忽略"如实告知。
          if (seriesCount <= 1 && seriesColors.length > 1) {
            warnings.push(
              `seriesColors 传了 ${seriesColors.length} 个颜色，但该图只有 ${seriesCount || 1} 个数据系列；` +
              `已只应用第 1 个（宿主对单系列多色的处理是逐点染色，会出现彩虹柱）。`
            );
          }
          const colorLimit = seriesCount <= 1 ? Math.min(1, seriesColors.length) : seriesColors.length;
          seriesColors.slice(0, colorLimit).forEach((colorHex, idx) => {
            const bgr = hexToExcelColor(colorHex);
            const sIdx = idx + 1;
            if (bgr !== null && sIdx <= seriesCount) {
              try {
                const series = seriesCol.Item(sIdx);
                if (series.Format && series.Format.Line) {
                  series.Format.Line.ForeColor.RGB = bgr;
                } else if (series.Border) {
                  series.Border.Color = bgr;
                }
                if (series.Format && series.Format.Fill) {
                  series.Format.Fill.ForeColor.RGB = bgr;
                } else if (series.Interior) {
                  series.Interior.Color = bgr;
                }
              } catch (e) {
                log(`系列 ${sIdx} 颜色注入警告: ${e.message}`);
              }
            }
          });
        }

        // 高阶单系列独立设置 (seriesSettings)
        const seriesTypeApplied = [];
        const seriesTypeWarnings = [];
        const seriesAxisApplied = [];
        if (Array.isArray(seriesSettings)) {
          seriesSettings.forEach((ss) => {
            const sIdx = Number(ss.seriesIndex);
            if (sIdx >= 1 && sIdx <= seriesCount) {
              try {
                const series = seriesCol.Item(sIdx);
                if (ss.smooth !== undefined) {
                  series.Smooth = !!ss.smooth;
                }
                // 单系列图表类型 → 做**组合图**（柱 + 折线同图）。
                // xlChartType：折线=4 / 柱状簇状=51 / 折线带数据点=65 / 面积=1 / 散点=75
                if (ss.type !== undefined && ss.type !== null && ss.type !== "") {
                  const ST = { line: 4, line_markers: 65, column: 51, column_clustered: 51, area: 1, scatter: 75, bar: 57 };
                  const key = String(ss.type).toLowerCase();
                  const code = ST[key] !== undefined ? ST[key] : (Number.isFinite(Number(ss.type)) ? Number(ss.type) : undefined);
                  if (code === undefined) { seriesTypeWarnings.push(`系列 ${sIdx}: 不认识的 type=${ss.type}`); }
                  else { series.ChartType = code; seriesTypeApplied.push({ seriesIndex: sIdx, type: key, code }); }
                }
                // 副坐标轴：组合图里把折线放到次轴（xlSecondary=2）
                if (ss.axisGroup !== undefined) {
                  series.AxisGroup = Number(ss.axisGroup) === 2 ? 2 : 1;
                  seriesAxisApplied.push({ seriesIndex: sIdx, axisGroup: series.AxisGroup });
                }
                if (ss.color) {
                  const bgr = hexToExcelColor(ss.color);
                  if (bgr !== null) {
                    if (series.Format && series.Format.Line) series.Format.Line.ForeColor.RGB = bgr;
                    else if (series.Border) series.Border.Color = bgr;
                    if (series.Format && series.Format.Fill) series.Format.Fill.ForeColor.RGB = bgr;
                    else if (series.Interior) series.Interior.Color = bgr;
                  }
                }
              } catch (e) {}
            }
          });
        }
      } catch (e) {
        log("系列格式化提示: " + e.message);
      }

      // 4.2 Y 轴数值范围、刻度与数字格式设置 (核心排版防呆)
      if (yAxis) {
        try {
          // xlValue = 2 (数值坐标轴), xlPrimary = 1 (主坐标轴)
          const valAxis = chart.Axes(2, 1);
          if (valAxis) {
            if (yAxis.min !== undefined && yAxis.min !== null) {
              valAxis.MinimumScale = Number(yAxis.min);
            }
            if (yAxis.max !== undefined && yAxis.max !== null) {
              valAxis.MaximumScale = Number(yAxis.max);
            }
            if (yAxis.step !== undefined && yAxis.step !== null) {
              valAxis.MajorUnit = Number(yAxis.step);
            }
            if (yAxis.numberFormat) {
              // ⚠️ 真机实测（真实使用者反馈）：`TickLabels.NumberFormat` **写入即静默吞掉** ——
              // 赋值不抛错、读回仍是 `General`，且不给任何提示。改走 `NumberFormatLocal`
              // （本地化格式串，宿主上才真正生效），并且**无论走哪条路都必须读回核对**。
              const wantFmt = String(yAxis.numberFormat);
              try { valAxis.TickLabels.NumberFormat = wantFmt; } catch (e) {}
              const readFmt = () => {
                try {
                  const a = String(valAxis.TickLabels.NumberFormat);
                  if (a && a !== "General") return a;
                  const b = String(valAxis.TickLabels.NumberFormatLocal);
                  return b || a;
                } catch (e) { return null; }
              };
              if (String(readFmt()) !== wantFmt) {
                try { valAxis.TickLabels.NumberFormatLocal = wantFmt; } catch (e) {}
              }
              const gotFmt = readFmt();
              if (gotFmt === null || String(gotFmt) === "General") {
                warnings.push(`数值轴刻度格式未生效：请求 ${JSON.stringify(wantFmt)}，宿主读回 ${JSON.stringify(gotFmt)}` +
                  `（该属性在本机表现为"写入即静默吞掉"，已同时尝试 NumberFormat 与 NumberFormatLocal）`);
              }
            }
            if (yAxis.title) {
              // 写入后**读回核对**：真机实测 `AxisTitle` 不一定生效，且不生效时毫无提示。
              // `HasTitle` 在不同宿主上取 `true` 或 `-1`（xlChartTitleOn），两个都试。
              const wantTitle = String(yAxis.title);
              try { valAxis.HasTitle = -1; } catch (e) { try { valAxis.HasTitle = true; } catch (e2) {} }
              try { valAxis.AxisTitle.Text = wantTitle; } catch (e) {}
              const gotTitle = (() => { try { return String(valAxis.AxisTitle.Text); } catch (e) { return null; } })();
              if (gotTitle !== wantTitle) {
                warnings.push(`数值轴标题未生效：请求 ${JSON.stringify(wantTitle)}，宿主读回 ${JSON.stringify(gotTitle)}。` +
                  `可改用 wps_execute_script 直接设置 ax.HasTitle 与 ax.AxisTitle.Text 后读回核对。`);
              }
            }
          }
        } catch (e) {
          log("数值坐标轴设置提示: " + e.message);
        }
      }

      // 建图后读回真实 ChartType：宿主可能（在任何版本）把请求的类型落成别的类型，
      // 原实现对此毫无察觉（ISS-17）。这里把实际类型带回，不一致就给出 warning。
      let actualChartType = null;
      try { actualChartType = Number(shape.Chart.ChartType); } catch (e) {}
      // 组合图核对：逐系列读回真实 ChartType，确认"折线真的是折线"
      let seriesTypes = null;
      try {
        const sc = shape.Chart.SeriesCollection();
        seriesTypes = [];
        for (let i = 1; i <= Number(sc.Count); i++) { try { seriesTypes.push(Number(sc.Item(i).ChartType)); } catch (e) { seriesTypes.push(null); } }
      } catch (e) {}
      if (actualChartType !== null && Number.isFinite(actualChartType) && actualChartType !== xlChartType) {
        warnings.push(
          `请求的 chartType=${chartType}（xlChartType=${xlChartType}）实际落成 ChartType=${actualChartType}；` +
          `请用 wps_get_charts 复核，必要时改用 wps_execute_script 直接指定常量。`
        );
      }

      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        chartType,
        requestedChartType: xlChartType,
        seriesTypes,
        actualChartType,
        warnings,
        dataRange: targetDataRange,
        title: title || "",
        left,
        top,
        width,
        height,
        shapeName: shape.Name,
        chartIndex: getChartOrdinal(sheet, shape.Name),
        message: `已成功在 [${sheet.Name}] 创建 ${chartType} 原生图表，数据源为 ${targetDataRange}`
      };
    } catch (transactionErr) {
      // 事务回滚：立即自毁残缺半成品 Shape，绝不在工作表留下空白废图！
      try {
        if (shape) shape.Delete();
      } catch (delErr) {}
      throw new Error(`图表创建异常并已自动销毁半成品对象: ${transactionErr.message}`);
    }
  }

  // 15.1 原生矢量图表删除与清理引擎
  function deleteChart(app, params) {
    const {
      sheetName,
      workbookName,
      chartTitle,
      shapeName,
      leftCell,
      chartIndex,
      clearAll = false
    } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const count = shapes.Count;
    const chartOrdinals = {};
    let chartOrdinal = 0;
    for (let i = 1; i <= count; i++) {
      if (shapes.Item(i).HasChart) {
        chartOrdinal++;
        chartOrdinals[i] = chartOrdinal;
      }
    }

    let targetLeft = null;
    let targetTop = null;
    if (leftCell) {
      try {
        const anchor = sheet.Range(leftCell);
        targetLeft = anchor.Left;
        targetTop = anchor.Top;
      } catch (e) {}
    }

    // 先收集候选、再决定删不删。
    // `leftCell` 是按**像素邻近（±30px）**匹配的，多张图叠在同一位置时会**全部命中**——
    // 实测传 leftCell:"H2" 一次删掉了 14 张图（问题台账 ISS-19）。破坏性操作必须先设卡。
    const candidates = [];
    for (let i = 1; i <= count; i++) {
      const shp = shapes.Item(i);
      if (!shp.HasChart) continue;
      let reason = null;
      if (clearAll) {
        reason = "clearAll";
      } else if (shapeName && shp.Name === shapeName) {
        reason = "shapeName";
      } else if (chartIndex && chartOrdinals[i] === Number(chartIndex)) {
        reason = "chartIndex";
      } else if (targetLeft !== null && targetTop !== null) {
        if (Math.abs(shp.Left - targetLeft) < 30 && Math.abs(shp.Top - targetTop) < 30) {
          reason = "leftCell(±30px)";
        }
      } else if (chartTitle) {
        try {
          if (shp.Chart.HasTitle && shp.Chart.ChartTitle.Text.includes(chartTitle)) {
            reason = "chartTitle";
          }
        } catch (e) {}
      }
      if (reason) {
        candidates.push({ index: i, name: shp.Name, reason: reason, left: Math.round(shp.Left), top: Math.round(shp.Top) });
      }
    }

    if (candidates.length === 0) {
      throw new Error(
        `未匹配到任何图表：工作表 [${sheet.Name}] 上没有符合条件的图表。` +
        `请先用 wps_get_charts 读回图表清单（shapeName / 序号 / 标题 / 位置），再指定要删除的那一张。`
      );
    }

    if (!clearAll && candidates[0].reason === "leftCell(±30px)" && candidates.length > 1) {
      throw new Error(
        `leftCell 按像素邻近（±30px）匹配，本次命中 ${candidates.length} 张图表，**已拒绝批量删除**以免误伤。` +
        `命中清单：${candidates.map(c => `${c.name}(左${c.left},上${c.top})`).join("；")}。` +
        `请改用 shapeName 或 chartIndex 精确指定；确认要全删请显式传 clearAll: true。`
      );
    }

    const deletedNames = [];
    // 从后往前删，避免索引位移
    for (let k = candidates.length - 1; k >= 0; k--) {
      const shp = shapes.Item(candidates[k].index);
      deletedNames.push(shp.Name);
      shp.Delete();
    }
    const deletedCount = deletedNames.length;

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      deletedCount,
      deletedNames,
      matchedBy: candidates[0].reason,
      message: `已成功在工作表 [${sheet.Name}] 中删除 ${deletedCount} 张图表（匹配方式：${candidates[0].reason}）`
    };
  }
