function getWorksheet(app,name){return wb.Worksheets.Item(name);}function log(){}
function hexToExcelColor(hex) {
    if (!hex) return null;
    const cleanHex = hex.replace("#", "");
    if (cleanHex.length !== 6) return null;
    const r = parseInt(cleanHex.substring(0, 2), 16);
    const g = parseInt(cleanHex.substring(2, 4), 16);
    const b = parseInt(cleanHex.substring(4, 6), 16);
    return (b << 16) | (g << 8) | r;
  }
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
const r=addChart(app,{workbookName:wb.Name,sheetName:"无表对照",dataRange:"A1:B3",chartType:"column",title:"固定基线标题",position:{leftCell:"G15",width:300,height:180},replaceExisting:false});return r;
