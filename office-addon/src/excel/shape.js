// ── 模块: src/excel/shape.js — 形状与图片 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1688-1731 行（原样搬迁，未改写）。
  // 10. 形状与图片
  async function handleInsertImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = sheet.shapes.addImage(params.base64Image);
      if (params.left) shape.left = params.left;
      if (params.top) shape.top = params.top;
      if (params.width) shape.width = params.width;
      if (params.height) shape.height = params.height;

      shape.load("name, id");
      await context.sync();
      return { success: true, name: shape.name, id: shape.id };
    });
  }

  async function handleListShapes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shapes = sheet.shapes.load("items/name, items/id, items/type, items/left, items/top, items/width, items/height");
      await context.sync();
      return {
        shapes: shapes.items.map(s => ({ name: s.name, id: s.id, type: s.type, left: s.left, top: s.top, width: s.width, height: s.height }))
      };
    });
  }

  async function handleUpdateShape(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = sheet.shapes.getItem(params.name || params.shapeName);
      if (params.action === "delete") {
        shape.delete();
      } else {
        if (params.left !== undefined) shape.left = params.left;
        if (params.top !== undefined) shape.top = params.top;
        if (params.width !== undefined) shape.width = params.width;
        if (params.height !== undefined) shape.height = params.height;
      }
      await context.sync();
      return { success: true };
    });
  }

