  function pptGetSlideShapes(app, params) {
    const { presentationName, slideIndex } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }

    const slide = pres.Slides.Item(idx);
    const count = slide.Shapes ? slide.Shapes.Count : 0;
    const shapes = [];

    for (let s = 1; s <= count; s++) {
      try {
        const shp = slide.Shapes.Item(s);
        const hasText = Boolean(shp.HasTextFrame && shp.TextFrame.HasText);
        let textContent = "";
        if (hasText) {
          textContent = (shp.TextFrame.TextRange.Text || "").trim();
        }

        const hasTable = Boolean(shp.HasTable);
        let tableMeta = null;
        if (hasTable && shp.Table) {
          tableMeta = {
            rows: shp.Table.Rows ? shp.Table.Rows.Count : 0,
            columns: shp.Table.Columns ? shp.Table.Columns.Count : 0
          };
        }

        const hasChart = Boolean(shp.HasChart);

        shapes.push({
          shapeIndex: s,
          shapeId: shp.Id,
          name: shp.Name || `Shape_${s}`,
          typeCode: shp.Type,
          typeName: getShapeTypeName(shp.Type),
          left: Math.round(Number(shp.Left) * 10) / 10,
          top: Math.round(Number(shp.Top) * 10) / 10,
          width: Math.round(Number(shp.Width) * 10) / 10,
          height: Math.round(Number(shp.Height) * 10) / 10,
          rotation: shp.Rotation || 0,
          zOrderPosition: shp.ZOrderPosition || s,
          hasText,
          fontSize: hasText ? Number(shp.TextFrame.TextRange.Font.Size) : undefined,
          text: textContent ? (textContent.length > 200 ? textContent.slice(0, 200) + "..." : textContent) : undefined,
          hasTable,
          table: tableMeta || undefined,
          hasChart
        });
      } catch (e) {
        log(`读取第 ${idx} 页第 ${s} 个形状元数据失败: ${e.message}`);
      }
    }

    return {
      success: true,
      presentationName: pres.Name,
      slideIndex: idx,
      ...pptPageSize(pres),
      shapeCount: count,
      shapes,
      // 版式与占位符读回（ISS-44）：宿主目前无法区分矩形/圆角/椭圆（typeCode 都是 1），
      // 但占位符（typeCode 14）与所在版式名可读，用于核对"是否套用了正确版式"。
      layout: (() => {
        try {
          return {
            name: slide.CustomLayout ? slide.CustomLayout.Name : undefined,
            layoutIndex: Number(slide.Layout),
            placeholderCount: slide.Shapes.Placeholders ? slide.Shapes.Placeholders.Count : shapes.filter(s => s.typeCode === 14).length,
            placeholders: (() => {
              try {
                const list = [];
                const phCount = slide.Shapes.Placeholders.Count;
                for (let p = 1; p <= phCount; p++) {
                  const ph = slide.Shapes.Placeholders.Item(p);
                  list.push({
                    shapeId: ph.Id,
                    name: ph.Name,
                    placeholderType: (() => { try { return Number(ph.PlaceholderFormat.Type); } catch (e) { return undefined; } })(),
                    hasText: Boolean(ph.HasTextFrame && ph.TextFrame.HasText),
                    text: (() => { try { return ph.TextFrame.HasText ? String(ph.TextFrame.TextRange.Text).slice(0, 200) : ""; } catch (e) { return undefined; } })()
                  });
                }
                return list;
              } catch (e) { return []; }
            })()
          };
        } catch (e) {
          return { error: e.message };
        }
      })(),
      message: `已成功获取第 ${idx} 页幻灯片中全部 ${shapes.length} 个形状的几何与属性信息`
    };
  }

