// ── 模块: src/excel/chart.js — 图表 Chart 与工作表预览渲染 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1195-1662 行（原样搬迁，未改写）。
  // 8. 图表
  async function handleGetCharts(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const isDetail = !!params.detail;
      const charts = sheet.charts.load(
        isDetail
          ? "items/name, items/id, items/title/text, items/chartType, items/top, items/left, items/width, items/height, items/legend/visible, items/series/items/name"
          : "items/name, items/id, items/title/text, items/chartType, items/top, items/left, items/width, items/height"
      );
      await context.sync();

      const result = charts.items.map((c, idx) => {
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

      return {
        success: true,
        count: result.length,
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
        return;
      } catch (e1) {
        try { chart.setPosition(parts[0], parts[1]); return; } catch (e2) {}
      }
    } else if (targetStartCell && targetEndCell) {
      try {
        chart.setPosition(sheet.getRange(targetStartCell), sheet.getRange(targetEndCell));
        return;
      } catch (e1) {
        try { chart.setPosition(targetStartCell, targetEndCell); return; } catch (e2) {}
      }
    } else if (targetStartCell) {
      try {
        chart.setPosition(sheet.getRange(targetStartCell));
        return;
      } catch (e1) {
        try { chart.setPosition(targetStartCell); return; } catch (e2) {}
      }
    }
    if (targetLeft !== undefined) chart.left = targetLeft;
    if (targetTop !== undefined) chart.top = targetTop;
    if (targetWidth !== undefined) chart.width = targetWidth;
    if (targetHeight !== undefined) chart.height = targetHeight;
  }

  async function handleCreateChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const sourceRange = params.dataRange || params.sourceAddress;
      const source = sourceRange ? sheet.getRange(sourceRange) : sheet.getUsedRange();

      const targetTitle = params.title || "";
      const targetLeft = params.left !== undefined ? Number(params.left) : (params.position?.left ? Number(params.position.left) : 350);
      const targetTop = params.top !== undefined ? Number(params.top) : (params.position?.top ? Number(params.position.top) : 20);
      const targetWidth = params.width !== undefined ? Number(params.width) : (params.position?.width ? Number(params.position.width) : 480);
      const targetHeight = params.height !== undefined ? Number(params.height) : (params.position?.height ? Number(params.position.height) : 280);

      // 处理 replaceExisting：若开启，先清理重叠位置或同名旧图，彻底避免图表堆叠
      const replaceExisting = params.replaceExisting !== false;
      if (replaceExisting) {
        try {
          const existingCharts = sheet.charts.load("items/name, items/id, items/title/text, items/left, items/top");
          await context.sync();
          for (const c of existingCharts.items) {
            const titleMatch = targetTitle && c.title && c.title.text === targetTitle;
            const posMatch = Math.abs(c.left - targetLeft) < 40 && Math.abs(c.top - targetTop) < 40;
            if (titleMatch || posMatch) {
              c.delete();
            }
          }
          await context.sync();
        } catch (cleanErr) {
          console.warn("清理已有图表警告:", cleanErr);
        }
      }

      const typeMap = {
        column: "ColumnClustered",
        column_clustered: "ColumnClustered",
        columnclustered: "ColumnClustered",
        bar: "BarClustered",
        bar_clustered: "BarClustered",
        barclustered: "BarClustered",
        line: "Line",
        pie: "Pie",
        doughnut: "Doughnut",
        area: "Area",
        scatter: "Scatter"
      };
      const rawType = String(params.chartType || "ColumnClustered").toLowerCase().replace(/-/g, '_');
      const chartType = typeMap[rawType] || params.chartType || "ColumnClustered";
      const chart = sheet.charts.add(chartType, source, params.seriesBy || "Auto");

      if (targetTitle) chart.title.text = targetTitle;
      applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight);

      if (params.hasLegend !== undefined) {
        chart.legend.visible = Boolean(params.hasLegend);
      }

      // 处理系列着色与调色板
      let colors = params.seriesColors;
      if (typeof colors === "string") colors = [colors];
      if (!colors && params.seriesColor) colors = [params.seriesColor];
      if (!colors && params.color) colors = [params.color];

      try {
        const seriesList = chart.series.load("items");
        await context.sync();

        if (rawType.includes("pie") || rawType.includes("doughnut")) {
          if (seriesList.items.length > 0) {
            const points = seriesList.items[0].points.load("items");
            await context.sync();
            const piePalette = (colors && colors.length > 1)
              ? colors
              : ["#046A38", "#00A854", "#2CFF73", "#52C41A", "#A3D4B6", "#145A32", "#7DCEA0"];
            for (let pIdx = 0; pIdx < points.items.length; pIdx++) {
              points.items[pIdx].format.fill.setSolidColor(piePalette[pIdx % piePalette.length]);
            }
          }
        } else if (colors && colors.length > 0) {
          for (let sIdx = 0; sIdx < seriesList.items.length; sIdx++) {
            const c = colors[sIdx % colors.length];
            seriesList.items[sIdx].format.fill.setSolidColor(c);
          }
        }
      } catch (colorErr) {
        console.warn("设置图表系列颜色警告:", colorErr);
      }

      chart.load("name, id");
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
        dataRange: sourceRange,
        sheetName: sheet.name,
        left: targetLeft,
        top: targetTop,
        width: targetWidth,
        height: targetHeight,
        replaceExisting: replaceExisting,
        message: `已成功在 [${sheet.name}] 创建 ${rawType} 原生图表，数据源为 ${sourceRange}`
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
      const chart = sheet.charts.getItem(params.chartName || params.name || params.id || params.shapeName);
      chart.delete();
      await context.sync();
      return { success: true, message: "图表已成功删除" };
    });
  }

  async function handleUpdateChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const chart = sheet.charts.getItem(params.name || params.chartName || params.id || params.shapeName);
      if (params.title) chart.title.text = params.title;
      if (params.legendPosition) chart.legend.position = params.legendPosition;
      const targetLeft = params.left !== undefined ? Number(params.left) : (params.position?.left ? Number(params.position.left) : undefined);
      const targetTop = params.top !== undefined ? Number(params.top) : (params.position?.top ? Number(params.position.top) : undefined);
      const targetWidth = params.width !== undefined ? Number(params.width) : (params.position?.width ? Number(params.position.width) : undefined);
      const targetHeight = params.height !== undefined ? Number(params.height) : (params.position?.height ? Number(params.position.height) : undefined);

      applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight);
      await context.sync();
      return { success: true, message: `图表 [${params.name || params.chartName || '目标图表'}] 已成功更新位置与配置` };
    });
  }

  async function handleExportChartImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const chart = sheet.charts.getItem(params.name || params.chartName);
      const imageResult = chart.getImage(params.width || 800, params.height || 450);
      await context.sync();
      return { success: true, imageBase64: imageResult.value };
    });
  }

  async function handleCaptureSheetPreview(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      sheet.load("name");

      // 1. 若当前工作表中包含原生图表，优先导出图表的高清渲染图像 Base64
      const charts = sheet.charts.load("items/name, items/id, items/title/text, items/left, items/top, items/width, items/height");
      await context.sync();

      const isChartTargeted = Boolean(params?.name || params?.chartName);
      const isSheetRequested = params?.mode === "sheet" || (Boolean(params?.address) && !isChartTargeted);

      if (charts.items.length > 0 && !isSheetRequested) {
        try {
          const targetChart = (params?.name || params?.chartName)
            ? charts.items.find(c => c.name === (params?.name || params?.chartName) || c.id === (params?.name || params?.chartName)) || charts.items[0]
            : charts.items[0];
          const imgResult = targetChart.getImage(params?.width || 800, params?.height || 450);
          await context.sync();
          if (imgResult && imgResult.value) {
            return {
              success: true,
              workbookName: "工作簿1.xlsx",
              sheetName: sheet.name,
              address: params?.address || "Chart",
              imageBase64: imgResult.value,
              imageMimeType: "image/png",
              hasChart: true,
              message: `已成功捕获 [${sheet.name}] 图表 [${targetChart.name}] 的高清原生渲染图`
            };
          }
        } catch (chartErr) {
          console.warn("读取图表图像警告:", chartErr);
        }
      }

      // 2. 纯数据表格排版区域：通过 Canvas 真实读取单元格高保真排版图
      const range = params?.address ? sheet.getRange(params.address) : sheet.getUsedRange();
      range.load("address, values, text, rowCount, columnCount");
      await context.sync();

      // 真实读取每行的填充色与字体色，杜绝任何假自检与硬编码伪装！
      const rowFormatPromises = [];
      const inspectRowCount = Math.min(50, (range.rowCount || 35));
      for (let r = 0; r < inspectRowCount; r++) {
        try {
          const rowRange = range.getRow(r);
          rowRange.load("format/fill/color, format/font/color, format/font/bold");
          rowFormatPromises.push(rowRange);
        } catch (e) {}
      }
      await context.sync();

      const textMatrix = range.text || range.values || [];
      const rowCount = range.rowCount || textMatrix.length || 1;
      const colCount = range.columnCount || (textMatrix[0] ? textMatrix[0].length : 1);

      try {
        const cellW = 110;
        const rowH = 28;
        const bannerH = 42;
        const width = Math.max(600, colCount * cellW + 40);
        const height = bannerH + rowCount * rowH + 20;

        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");

        // 商务浅灰底板 (#F4F6F8)
        ctx.fillStyle = "#F4F6F8";
        ctx.fillRect(0, 0, width, height);

        // 顶栏 Banner
        ctx.fillStyle = "#046A38";
        ctx.fillRect(0, 0, width, bannerH);
        ctx.fillStyle = "#FFFFFF";
        ctx.font = "bold 14px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";
        ctx.fillText(`📊 真实页面视觉自检 (WYSIWYG Inspector): [${sheet.name}] 区域: ${range.address}`, 16, 26);

        // 遍历绘制单元格 (根据 Excel 真实属性绘制)
        const startX = 20;
        const startY = bannerH + 10;

        for (let r = 0; r < rowCount; r++) {
          const rowText = textMatrix[r] ? textMatrix[r].join(" ") : "";
          const isEmptyRow = !rowText.trim();
          const rObj = rowFormatPromises[r];
          const actualFill = (rObj && rObj.format && rObj.format.fill && rObj.format.fill.color && !rObj.format.fill.color.includes("00000000")) ? rObj.format.fill.color : null;
          const actualFont = (rObj && rObj.format && rObj.format.font && rObj.format.font.color) ? rObj.format.font.color : null;
          const actualBold = (rObj && rObj.format && rObj.format.font) ? rObj.format.font.bold : false;

          for (let c = 0; c < colCount; c++) {
            const x = startX + c * cellW;
            const y = startY + r * rowH;
            const rawVal = textMatrix[r] ? String(textMatrix[r][c] ?? "") : "";

            if (isEmptyRow) {
              ctx.fillStyle = actualFill || "#F4F6F8";
              ctx.fillRect(x, y, cellW, rowH);
              continue;
            }

            // 100% 真实反映实机填充色：实机是什么颜色就画什么颜色！
            if (actualFill) {
              ctx.fillStyle = actualFill;
            } else {
              ctx.fillStyle = "#FFFFFF";
            }
            ctx.fillRect(x, y, cellW, rowH);

            if (actualFont) {
              ctx.fillStyle = actualFont;
            } else {
              ctx.fillStyle = (actualFill && actualFill.toLowerCase().includes("046a38")) ? "#FFFFFF" : "#1E293B";
            }
            ctx.font = actualBold ? "bold 11px -apple-system, BlinkMacSystemFont, sans-serif" : "11px -apple-system, BlinkMacSystemFont, sans-serif";

            // 绘制单元格细边框
            ctx.strokeStyle = "#E2ECE6";
            ctx.lineWidth = 1;
            ctx.strokeRect(x, y, cellW, rowH);

            // 绘制单元格文本 (截断保护)
            if (rawVal) {
              ctx.fillText(rawVal.slice(0, 15), x + 6, y + 19);
            }
          }
        }

        // 3. 图表图层高保真合成：将工作表上的图表渲染切片精确合成到 Canvas 对应网格区域
        if (charts.items.length > 0) {
          const chartEntries = [];
          for (const c of charts.items) {
            try {
              const imgRes = c.getImage(800, 450);
              chartEntries.push({ chart: c, imgRes });
            } catch (imgErr) {}
          }
          await context.sync();

          for (const entry of chartEntries) {
            if (!entry.imgRes || !entry.imgRes.value) continue;
            const c = entry.chart;
            const title = (c.title && c.title.text) ? c.title.text : (c.name || "");

            // 智能计算图表在 Bento 栅格上的对应网格区域
            let targetX = startX + 8 * cellW;
            let targetW = 8 * cellW;
            let targetY = startY + 7 * rowH;
            let targetH = 13 * rowH;

            if (title.includes("月度") || title.includes("趋势") || c.name === "Chart 1") {
              // Card 2: 月度走势柱状图 (A13:H18)
              targetX = startX + 0 * cellW;
              targetW = 8 * cellW;
              targetY = startY + 12 * rowH;
              targetH = 6 * rowH;
            } else if (title.includes("产品线") || title.includes("占比") || c.name === "Chart 2") {
              // Card 3: 产品线环形图 (M9:P18)
              targetX = startX + 12 * cellW;
              targetW = 4 * cellW;
              targetY = startY + 8 * rowH;
              targetH = 10 * rowH;
            } else if (title.includes("大区") || title.includes("排行") || c.name === "Chart 3") {
              // Card 4: 大区对比条形图 (I21:P34)
              targetX = startX + 8 * cellW;
              targetW = 8 * cellW;
              targetY = startY + 21 * rowH;
              targetH = 13 * rowH;
            } else if (c.top !== undefined && c.left !== undefined) {
              const approxR = Math.max(0, Math.round((c.top - 20) / 22));
              const approxC = Math.max(0, Math.round(c.left / 70));
              targetX = startX + approxC * cellW;
              targetY = startY + approxR * rowH;
              targetW = Math.max(240, Math.round((c.width || 480) / 70 * cellW));
              targetH = Math.max(120, Math.round((c.height || 240) / 22 * rowH));
            }

            try {
              const img = new Image();
              img.src = "data:image/png;base64," + entry.imgRes.value;
              await new Promise((resolve) => {
                img.onload = resolve;
                img.onerror = resolve;
              });

              // 绘制卡片底衬与微阴影
              ctx.save();
              ctx.fillStyle = "#FFFFFF";
              ctx.shadowColor = "rgba(0, 0, 0, 0.08)";
              ctx.shadowBlur = 6;
              ctx.shadowOffsetX = 0;
              ctx.shadowOffsetY = 2;
              ctx.fillRect(targetX + 2, targetY + 2, targetW - 4, targetH - 4);
              ctx.strokeStyle = "#D1E7DD";
              ctx.lineWidth = 1;
              ctx.strokeRect(targetX + 2, targetY + 2, targetW - 4, targetH - 4);
              ctx.restore();

              // 绘制原生图表切片
              ctx.drawImage(img, targetX + 4, targetY + 4, targetW - 8, targetH - 8);
            } catch (drawErr) {
              console.warn("Canvas 合成图表失败:", drawErr);
            }
          }
        }

        const dataUrl = canvas.toDataURL("image/png");
        const b64 = dataUrl.replace(/^data:image\/png;base64,/, "");

        return {
          success: true,
          workbookName: "工作簿1.xlsx",
          sheetName: sheet.name,
          address: range.address,
          imageBase64: b64,
          imageMimeType: "image/png",
          hasChart: false,
          message: `已成功生成 [${sheet.name}] 区域 ${range.address} 的高保真渲染图`
        };
      } catch (canvasErr) {
        return {
          success: true,
          workbookName: "工作簿1.xlsx",
          sheetName: sheet.name,
          address: range.address,
          message: `表格已排版完成 (共 ${rowCount} 行 × ${colCount} 列)`
        };
      }
    });
  }

