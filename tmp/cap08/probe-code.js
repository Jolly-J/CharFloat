// CAP-08 真机探测：Office.js Excel Shape API 可用范围探测。
// 运行环境：Microsoft Excel 任务窗格（Office.js），由 bridge 的 run_script 分支执行。
// 返回：自包含 JSON（不引用宿主对象），由调用方读取。
(async () => {
  const out = {
    probe: "CAP-08-ms-shapes",
    at: new Date().toISOString(),
    host: {
      version: (typeof Office !== "undefined" && Office.context && Office.context.diagnostics && Office.context.diagnostics.version) || null,
      platform: (typeof Office !== "undefined" && Office.context && Office.context.diagnostics && Office.context.diagnostics.platform) || null,
      hostName: (typeof Office !== "undefined" && Office.context && Office.context.diagnostics && Office.context.diagnostics.host) || null,
      excelRun: typeof Excel !== "undefined" && typeof Excel.run === "function"
    },
    requirementSets: null,
    support: {},
    results: {},
    shapes: [],
    writtenShapeNames: [],
    errors: []
  };

  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported) {
      const sets = ["ExcelApi 1.9", "ExcelApi 1.10", "ExcelApi 1.11", "ExcelApi 1.12", "ExcelApi 1.13", "ExcelApi 1.14", "ExcelApi 1.15", "ExcelApi 1.16", "ExcelApi 1.17", "ExcelApi 1.18"];
      out.requirementSets = {};
      for (const s of sets) {
        try { out.requirementSets[s] = Office.context.requirements.isSetSupported(s.split(" ")[0], s.split(" ")[1]); }
        catch (e) { out.requirementSets[s] = "ERR:" + e.message; }
      }
    }
  } catch (e) { out.errors.push("reqsets: " + e.message); }

  // API 表面存在性（静态，不求值宿主）
  function probeSurface() {
    const s = {};
    const sheetProto = null;
    try {
      // 通过一次真实 run 拿到代理对象后判断成员是否存在
      s.excelNamespace = typeof Excel;
      s.geometricShapeTypeCount = (() => { try { return Object.keys(Excel.GeometricShapeType || {}).length; } catch (e) { return "ERR"; } })();
      s.shapeZOrderValues = (() => { try { return Object.keys(Excel.ShapeZOrder || {}); } catch (e) { return "ERR"; } })();
      s.shapeFillTypeValues = (() => { try { return Object.keys(Excel.ShapeFillType || {}); } catch (e) { return "ERR"; } })();
      s.geometricShapeTypeNames = (() => { try { return Object.keys(Excel.GeometricShapeType || {}); } catch (e) { return "ERR"; } })();
    } catch (e) { s.surfaceError = e.message; }
    return s;
  }
  out.support.surface = probeSurface();

  function rec(key, fn) {
    return (async () => {
      try {
        const t0 = Date.now();
        const value = await fn();
        out.results[key] = { ok: true, ms: Date.now() - t0, value };
      } catch (e) {
        out.results[key] = { ok: false, ms: null, error: String(e && e.message || e), code: e && e.code, debug: e && e.debugInfo && (e.debugInfo.errorLocation ? e.debugInfo.errorLocation.fullAddress || JSON.stringify(e.debugInfo.errorLocation) : undefined) };
      }
    })();
  }

  try {
    await rec("createSheet", () => Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const name = "_cap08_probe";
      const existing = sheets.items.filter(x => x.name === name);
      if (existing.length) existing[0].delete();
      await context.sync();
      const sheet = sheets.add(name);
      sheet.activate();
      sheet.getRange("A1").values = [["CAP-08 MS Shape API 探测"], ["key"], ["value"]];
      await context.sync();
      return { sheetName: sheet.name };
    }));

    // 1. 几何形状
    await rec("addGeometricShape", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle, { left: 20, top: 20, width: 160, height: 90 });
      sh.name = "cap08_rect";
      sh.fill.setSolidColor("#2F6FEB");
      sh.lineFormat.color = "#12305E";
      sh.lineFormat.weight = 2;
      sh.load("name,id,type,left,top,width,height,rotation");
      await context.sync();
      return { name: sh.name, id: sh.id, type: String(sh.type), left: sh.left, top: sh.top, width: sh.width, height: sh.height, rotation: sh.rotation };
    }));

    // 2. 直线 / 连接符
    await rec("addLine", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.addLine(200, 20, 360, 110);
      sh.name = "cap08_line";
      sh.lineFormat.color = "#E4572E";
      sh.lineFormat.weight = 3;
      sh.load("name,id,type,left,top,width,height");
      await context.sync();
      return { name: sh.name, id: sh.id, type: String(sh.type), left: sh.left, top: sh.top, width: sh.width, height: sh.height };
    }));

    // 3. SVG
    await rec("addSvg", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" viewBox="0 0 120 80"><rect width="120" height="80" fill="#0F9D58"/><circle cx="60" cy="40" r="26" fill="#fff"/></svg>';
      const b64 = (typeof btoa === "function" ? btoa(svg) : "");
      const sh = sheet.shapes.addSvg(b64);
      sh.name = "cap08_svg";
      sh.load("name,id,type,left,top,width,height");
      await context.sync();
      return { name: sh.name, id: sh.id, type: String(sh.type), left: sh.left, top: sh.top, width: sh.width, height: sh.height, b64len: b64.length };
    }));

    // 4. 文本框
    await rec("addTextBox", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.addTextBox("CAP-08 文本框");
      sh.name = "cap08_text";
      sh.left = 20; sh.top = 140; sh.width = 200; sh.height = 50;
      sh.textFrame.textRange.font.bold = true;
      sh.textFrame.textRange.font.color = "#1F2937";
      sh.textFrame.textRange.font.size = 14;
      sh.load("name,id,type,left,top,width,height,textFrame/textRange/text");
      await context.sync();
      return { name: sh.name, id: sh.id, type: String(sh.type), left: sh.left, top: sh.top, width: sh.width, height: sh.height, text: sh.textFrame.textRange.text, rotation: sh.rotation };
    }));

    // 5. 读回：列表 + 填充 + 线条 + 层叠顺序
    await rec("listShapes", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const shapes = sheet.shapes.load("items/name,items/id,items/type,items/left,items/top,items/width,items/height,items/rotation,items/fill/type,items/fill/foregroundColor,items/lineFormat/color,items/lineFormat/weight,items/zOrderPosition");
      await context.sync();
      return shapes.items.map(s => ({ name: s.name, type: String(s.type), left: s.left, top: s.top, width: s.width, height: s.height, rotation: s.rotation, fillType: String(s.fill.type), fill: s.fill.foregroundColor, line: s.lineFormat.color, lineWeight: s.lineFormat.weight, z: String(s.zOrderPosition) }));
    }));

    // 6. 旋转 / 缩放
    await rec("rotateAndScale", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.getItem("cap08_rect");
      sh.rotation = 30;
      sh.scaleHeight(1.5, Excel.ShapeScaleType.currentSize);
      sh.scaleWidth(1.5, Excel.ShapeScaleType.currentSize);
      sh.load("name,rotation,width,height");
      await context.sync();
      return { name: sh.name, rotation: sh.rotation, width: sh.width, height: sh.height };
    }));

    // 7. 层级调整
    await rec("setZOrder", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.getItem("cap08_rect");
      sh.setZOrder(Excel.ShapeZOrder.bringToFront);
      const sh2 = sheet.shapes.getItem("cap08_line");
      sh2.setZOrder(Excel.ShapeZOrder.sendToBack);
      await context.sync();
      const check = sheet.shapes.load("items/name,items/zOrderPosition");
      await context.sync();
      return check.items.map(s => ({ name: s.name, z: String(s.zOrderPosition) }));
    }));

    // 8. 激活形状
    await rec("getActiveShape", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.getActiveShape();
      sh.load("name,id,type");
      await context.sync();
      return { name: sh.name, id: sh.id, type: String(sh.type) };
    }));

    // 9. 分组
    await rec("addGroup", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const names = ["cap08_rect", "cap08_line", "cap08_svg"];
      const shapes = names.map(n => sheet.shapes.getItem(n));
      shapes.forEach(s => s.load("id"));
      await context.sync();
      const ids = shapes.map(s => s.id);
      const group = sheet.shapes.addGroup(ids);
      group.name = "cap08_group";
      group.load("name,id,type,left,top,width,height");
      await context.sync();
      return { name: group.name, id: group.id, type: String(group.type), left: group.left, top: group.top, width: group.width, height: group.height, memberIds: ids };
    }));

    // 10. 分组读回与解组
    await rec("groupReadback", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const g = sheet.shapes.getItem("cap08_group");
      const shapes = g.group.shapes.load("items/name,items/type,items/left,items/top,items/width,items/height");
      await context.sync();
      const members = shapes.items.map(s => ({ name: s.name, type: String(s.type), left: s.left, top: s.top, width: s.width, height: s.height }));
      return { memberCount: members.length, members };
    }));

    // 11. getAsImage
    await rec("getAsImage", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.getItem("cap08_text");
      const img = sh.getAsImage(Excel.PictureFormat.png, 2);
      await context.sync();
      const v = img.value;
      return { hasValue: !!v, length: v ? v.length : 0, prefix: v ? v.slice(0, 32) : null };
    }));

    // 12. 现有形状列表（含 chart 类型）
    await rec("listAllShapes", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const shapes = sheet.shapes.load("items/name,items/type,items/id");
      await context.sync();
      return { count: shapes.items.length, items: shapes.items.map(s => ({ name: s.name, type: String(s.type) })) };
    }));

    // 13. 文本读回（文本范围）
    await rec("textReadback", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.getItem("cap08_text");
      sh.load("textFrame/textRange/text,textFrame/textRange/font/bold,textFrame/textRange/font/size,textFrame/textRange/font/color,textFrame/horizontalAlignment,textFrame/verticalAlignment");
      await context.sync();
      return { text: sh.textFrame.textRange.text, bold: sh.textFrame.textRange.font.bold, size: sh.textFrame.textRange.font.size, color: sh.textFrame.textRange.font.color, hAlign: String(sh.textFrame.horizontalAlignment), vAlign: String(sh.textFrame.verticalAlignment) };
    }));

    // 14. 删除
    await rec("deleteShape", () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("_cap08_probe");
      const sh = sheet.shapes.getItem("cap08_svg");
      sh.delete();
      await context.sync();
      const rest = sheet.shapes.load("items/name");
      await context.sync();
      return { remaining: rest.items.map(s => s.name) };
    }));
  } catch (e) {
    out.errors.push("outer: " + String(e && e.message || e));
  }

  return out;
})()
