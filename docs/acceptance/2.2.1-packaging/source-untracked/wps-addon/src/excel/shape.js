  // shape.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // 10.27 表格矢量绘图（CAP-07）
  //
  // 依据真机探测（WPS 12.1.28496）：宿主 `Sheet.Shapes` 全套可用——
  //   AddShape / AddLine / AddTextbox / **AddTextEffect（艺术字）** / AddPicture / BuildFreeform，
  //   分组走 `Shapes.Range([名...]).Group()`（**WPS 表格能分组，Office.js 侧这条未打通**）。
  // 这是 WPS 相对 Microsoft Excel 的**能力优势**：艺术字与自由曲线在 Office.js 侧没有对应 API。
  // 此前我们一个表格绘图工具都没有，AI 只能用 wps_execute_script 手写。

  /** 几何形状枚举（msoAutoShapeType 常用值）。 */
  const AUTO_SHAPE_TYPES = {
    rectangle: 1, rounded_rectangle: 5, oval: 9, ellipse: 9, diamond: 4, triangle: 7,
    right_triangle: 8, pentagon: 51, hexagon: 10, arrow_right: 33, arrow_left: 34,
    arrow_up: 35, arrow_down: 36, chevron: 52, cross: 11, star_5: 12, star_6: 13,
    callout_rounded: 106, cloud: 179, heart: 21, lightning: 22, sun: 24, moon: 23,
    can: 13, cube: 17, donut: 18, flow_chart_process: 61, flow_chart_decision: 63,
    flow_chart_terminator: 68, flow_chart_data: 62, line_horizontal: 130, text_box: 17
  };

  /** 形状文字的默认字体。中文报表里微软雅黑比宿主默认的宋体更清晰、字重更统一。 */
  const DEFAULT_SHAPE_FONT = "微软雅黑";

  function resolveShapeType(raw) {
    if (raw === undefined || raw === null || raw === "") return AUTO_SHAPE_TYPES.rectangle;
    if (typeof raw === "number") return raw;
    const key = String(raw).toLowerCase().trim();
    if (AUTO_SHAPE_TYPES[key] !== undefined) return AUTO_SHAPE_TYPES[key];
    if (/^\d+$/.test(key)) return parseInt(key, 10);
    throw new Error(`不认识的形状类型 "${raw}"。可用：${Object.keys(AUTO_SHAPE_TYPES).join(" / ")}`);
  }

  /**
   * 判断某个十六进制底色偏亮还是偏暗，用于给形状文字选黑字还是白字。
   *
   * 为什么需要：WPS 给**自选图形默认白字**。如果调用方填了白底又不指定字色，
   * 就会得到"白字白底"——文字完全看不见，而且没有任何报错（真机踩到过：
   * 流程框全部隐形）。这里按底色亮度自动选，调用方显式给 fontColor 时以调用方为准。
   */
  function isLightFill(hex) {
    const rgb = hexToExcelColor(hex);
    if (rgb === null || rgb === undefined) return null;   // 无底色 → 交给宿主默认
    const r = rgb & 0xff, g = (rgb >> 8) & 0xff, b = (rgb >> 16) & 0xff;
    return (0.299 * r + 0.587 * g + 0.114 * b) > 150;
  }

  /** 单元格地址取值：宿主方法是原生实现，**必须保留接收者**直接调用（摘出来会丢 this）。 */
  function shapeAddressOf(cell) {
    try { return String(cell.Address()); } catch (e) {
      try { return String(cell.Address); } catch (e2) { return ""; }
    }
  }

  /** 读回单个形状的完整状态。 */
  function readShape(shape, sheet) {
    const safe = (fn, fallback) => { try { const v = fn(); return v === undefined ? fallback : v; } catch (e) { return fallback; } };
    let text = null;
    try { text = String(shape.TextFrame.Characters().Text); } catch (e) { text = null; }
    return {
      name: safe(() => String(shape.Name), null),
      type: safe(() => Number(shape.Type), null),
      autoShapeType: safe(() => Number(shape.AutoShapeType), null),
      left: safe(() => Math.round(Number(shape.Left) * 100) / 100, null),
      top: safe(() => Math.round(Number(shape.Top) * 100) / 100, null),
      width: safe(() => Math.round(Number(shape.Width) * 100) / 100, null),
      height: safe(() => Math.round(Number(shape.Height) * 100) / 100, null),
      rotation: safe(() => Number(shape.Rotation), null),
      visible: safe(() => Number(shape.Visible) !== 0, null),
      fillColor: safe(() => { const v = Number(shape.Fill.ForeColor.RGB); return Number.isFinite(v) && v >= 0 ? excelColorToHex(v) : null; }, null),
      lineColor: safe(() => { const v = Number(shape.Line.ForeColor.RGB); return Number.isFinite(v) && v >= 0 ? excelColorToHex(v) : null; }, null),
      lineWeight: safe(() => Number(shape.Line.Weight), null),
      text,
      // 真实对齐（Excel 对象模型：HorizontalAlignment/VerticalAlignment）。
      // 放进 readShape 而不是只在 addShape 里，这样 list_shapes 也能核对——
      // 「框居中了但字没居中」这类静默失败必须能被读回发现。
      textAlign: safe(() => { const h = Number(shape.TextFrame.HorizontalAlignment); return h === -4108 ? "center" : h === -4152 ? "right" : h === -4131 ? "left" : null; }, null),
      textVAlign: safe(() => { const v = Number(shape.TextFrame.VerticalAlignment); return v === -4108 ? "middle" : v === -4107 ? "bottom" : v === -4160 ? "top" : null; }, null),
      topLeftCell: safe(() => shapeAddressOf(shape.TopLeftCell), null),
      isGroup: safe(() => Number(shape.Type) === 6, false),
      groupItemCount: safe(() => (Number(shape.Type) === 6 && shape.GroupItems ? Number(shape.GroupItems.Count) : null), null)
    };
  }

  /** 找出形状：优先按名字，其次按序号（1 基）。 */
  function findShape(sheet, params) {
    const shapes = sheet.Shapes;
    if (params.name || params.shapeName) {
      const name = String(params.name || params.shapeName);
      for (let i = 1; i <= shapes.Count; i++) {
        const sh = shapes.Item(i);
        try { if (String(sh.Name) === name) return sh; } catch (e) {}
      }
      const available = [];
      for (let i = 1; i <= Math.min(shapes.Count, 30); i++) {
        try { available.push(String(shapes.Item(i).Name)); } catch (e) {}
      }
      throw new Error(`找不到名为 "${name}" 的形状。现有形状（最多 30 个）：${available.join(" / ") || "无"}`);
    }
    if (Number.isFinite(Number(params.shapeIndex))) {
      const idx = Number(params.shapeIndex);
      if (idx < 1 || idx > shapes.Count) throw new Error(`形状序号 ${idx} 越界（当前共 ${shapes.Count} 个）`);
      return shapes.Item(idx);
    }
    throw new Error("需要提供 name（形状名）或 shapeIndex（序号，从 1 开始）");
  }

  /**
   * 画矢量形状：几何形状 / 直线 / 文本框 / **艺术字**。
   * 写完**读回真实几何与样式**，宿主没接受某个属性时写进 warnings，不做假成功。
   */
  function addShape(app, params) {
    const { sheetName, workbookName, kind = "geometric", shapeType, text, name,
            left = 0, top = 0, width = 160, height = 80, rotation,
            x1, y1, x2, y2, fillColor, lineColor, lineWeight,
            wordArtPreset, fontName, fontSize, bold, italic, fontColor,
            textAlign, textVAlign, marginLeft, marginRight, marginTop, marginBottom } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const warnings = [];

    let shape = null;
    const k = String(kind).toLowerCase();
    if (k === "geometric" || k === "autoshape") {
      shape = shapes.AddShape(resolveShapeType(shapeType), Number(left), Number(top), Number(width), Number(height));
    } else if (k === "line" || k === "connector") {
      shape = shapes.AddLine(
        Number(x1 !== undefined ? x1 : left), Number(y1 !== undefined ? y1 : top),
        Number(x2 !== undefined ? x2 : Number(left) + Number(width)),
        Number(y2 !== undefined ? y2 : Number(top) + Number(height))
      );
    } else if (k === "textbox" || k === "text") {
      shape = shapes.AddTextbox(1, Number(left), Number(top), Number(width), Number(height));
      // 文本框默认**不画边框、不填底色**——否则页面上每段文字都套一个框。
      // 调用方显式传 fillColor/lineColor 时才打开对应可见性（见下方）。
      try { shape.Line.Visible = 0; } catch (e) {}
      try { shape.Fill.Visible = 0; } catch (e) {}
      // 关掉自动缩放：宿主默认会把文本框收缩到文字宽度，导致"按给定宽度算的居中"全部偏掉
      // （真机踩到：柱顶数据标签偏离柱心 1~2pt）。宽度按调用方给的固定，版面才可预测。
      try { shape.TextFrame.AutoSize = 0; } catch (e) {}
      try { shape.TextFrame.WordWrap = 0; } catch (e) {}
    } else if (k === "wordart" || k === "texteffect") {
      // 艺术字：WPS 相对 Office.js 的独有能力
      const preset = Number.isFinite(Number(wordArtPreset)) ? Number(wordArtPreset) : 0;
      shape = shapes.AddTextEffect(
        preset, String(text === undefined ? "" : text), String(fontName || "宋体"),
        Number(fontSize || 36), bold === false ? 0 : -1, italic ? -1 : 0, Number(left), Number(top)
      );
    } else {
      throw new Error(`不支持的形状种类: ${kind}（可用 geometric / line / textbox / wordart）`);
    }
    if (!shape) throw new Error("宿主没有返回形状对象（创建失败）");

    if (name) { try { shape.Name = String(name); } catch (e) { warnings.push(`设置名字失败: ${e.message}`); } }
    if (k !== "line" && Number.isFinite(Number(rotation))) {
      try { shape.Rotation = Number(rotation); } catch (e) { warnings.push(`设置旋转失败: ${e.message}`); }
    }
    if (fillColor) {
      const rgb = hexToExcelColor(fillColor);
      if (rgb === null) warnings.push(`fillColor 无法解析: ${fillColor}`);
      else { try { shape.Fill.Visible = -1; shape.Fill.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置填充失败: ${e.message}`); } }
    }
    if (lineColor) {
      const rgb = hexToExcelColor(lineColor);
      if (rgb === null) warnings.push(`lineColor 无法解析: ${lineColor}`);
      else { try { shape.Line.Visible = -1; shape.Line.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置线条色失败: ${e.message}`); } }
    }
    if (Number.isFinite(Number(lineWeight))) {
      try { shape.Line.Weight = Number(lineWeight); } catch (e) { warnings.push(`设置线宽失败: ${e.message}`); }
    }
    // 文字：几何形状/文本框都能写字（艺术字的文字在创建时给）。
    // 字号/加粗/斜体/字体/字色对**所有形状类型**生效——原来只对艺术字生效，
    // 于是"画个文本框并给它 22pt 加粗"这类最常见需求被静默忽略。
    if (k !== "wordart" && k !== "texteffect" && text !== undefined && text !== null && String(text) !== "") {
      try {
        const chars = shape.TextFrame.Characters();
        chars.Text = String(text);
        if (Number.isFinite(Number(fontSize))) chars.Font.Size = Number(fontSize);
        if (bold !== undefined) chars.Font.Bold = bold ? -1 : 0;
        if (italic !== undefined) chars.Font.Italic = italic ? -1 : 0;
        chars.Font.Name = String(fontName || DEFAULT_SHAPE_FONT);
        if (fontColor) {
          const fc = hexToExcelColor(fontColor);
          if (fc !== null) chars.Font.Color = fc;
        } else {
          // 未指定字色：按填充色亮度选黑/白，避免"白字白底"隐形
          const light = isLightFill(fillColor);
          if (light === true) chars.Font.Color = hexToExcelColor("#333333");
          else if (light === false) chars.Font.Color = hexToExcelColor("#FFFFFF");
        }
        // 对齐：文本框默认左对齐/顶对齐，自选图形默认居中/垂直居中。
        //
        // ⚠️ 这里是 **Excel 的对象模型**，不是 PowerPoint 的：
        //   `TextFrame.TextRange.ParagraphFormat.Alignment` 在 Excel 里**不存在**
        //   （TextRange=undefined、Characters().ParagraphFormat=undefined），
        //   正确属性是 `TextFrame.HorizontalAlignment` / `VerticalAlignment`。
        // 之前用错的那条路径被 try/catch 吞掉，于是"框居中了、框里的字还是居左"，
        // 既不报错也看不出来（使用者一眼就看出来了）——所以这里改成**失败即告警**，不再静默。
        const tf = shape.TextFrame;
        const H = { left: -4131, center: -4108, right: -4152 };   // xlHAlign*
        const V = { top: -4160, middle: -4108, bottom: -4107 };   // xlVAlign*
        const wantH = textAlign || ((k !== "textbox" && k !== "text") ? "center" : null);
        const wantV = textVAlign || ((k !== "textbox" && k !== "text") ? "middle" : null);
        if (wantH && H[wantH] !== undefined) {
          try { tf.HorizontalAlignment = H[wantH]; } catch (e) { warnings.push(`设置水平对齐失败: ${e.message}`); }
        }
        if (wantV && V[wantV] !== undefined) {
          try { tf.VerticalAlignment = V[wantV]; } catch (e) { warnings.push(`设置垂直对齐失败: ${e.message}`); }
        }
        for (const [key, val] of Object.entries({ MarginLeft: marginLeft, MarginRight: marginRight, MarginTop: marginTop, MarginBottom: marginBottom })) {
          if (Number.isFinite(Number(val))) { try { tf[key] = Number(val); } catch (e) { warnings.push(`设置 ${key} 失败: ${e.message}`); } }
        }
      } catch (e) { warnings.push(`写入文字失败: ${e.message}`); }
    }

    const actual = readShape(shape, sheet);

    // 核对请求与读回：不一致就如实告警（不抛错，因为部分属性宿主可能合法地做了归一）
    const requested = { kind: k, left: Number(left), top: Number(top), width: Number(width), height: Number(height) };
    if (k === "geometric" || k === "autoshape") {
      const tol = 1.5;
      for (const key of ["left", "top", "width", "height"]) {
        if (actual[key] !== null && Math.abs(Number(actual[key]) - requested[key]) > tol) {
          warnings.push(`${key} 请求 ${requested[key]} 读回 ${actual[key]}`);
        }
      }
    }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      kind: k,
      shape: actual,
      requested,
      shapeCount: shapes.Count,
      warnings,
      message: `已在 [${sheet.Name}] 创建 ${k} 形状「${actual.name || name || ""}」（当前共 ${shapes.Count} 个形状）`
    };
  }

  /** 读回工作表上的全部形状（绘图能力的验收入口）。 */
  function listShapes(app, params) {
    const { sheetName, workbookName, detail = true, filterName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const items = [];
    for (let i = 1; i <= shapes.Count; i++) {
      const sh = shapes.Item(i);
      if (filterName) {
        let nm = null;
        try { nm = String(sh.Name); } catch (e) {}
        if (nm !== String(filterName)) continue;
      }
      items.push(detail ? readShape(sh, sheet) : { name: (() => { try { return String(sh.Name); } catch (e) { return null; } })() });
    }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      count: items.length,
      shapes: items,
      message: `工作表 [${sheet.Name}] 上共有 ${items.length} 个形状`
    };
  }

  /** 修改形状：位置/尺寸/旋转/填充/线条/文字/可见性，或删除。 */
  function updateShape(app, params) {
    const { sheetName, workbookName, action, left, top, width, height, rotation,
            fillColor, lineColor, lineWeight, text, visible, newName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    const warnings = [];

    const act = String(action || "update").toLowerCase();
    if (act === "delete") {
      const nm = (() => { try { return String(shape.Name); } catch (e) { return null; } })();
      shape.Delete();
      return {
        success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name,
        action: "delete", deletedName: nm, shapeCountAfter: sheet.Shapes.Count,
        message: `已删除形状「${nm}」（剩余 ${sheet.Shapes.Count} 个）`
      };
    }

    if (Number.isFinite(Number(left))) shape.Left = Number(left);
    if (Number.isFinite(Number(top))) shape.Top = Number(top);
    if (Number.isFinite(Number(width))) shape.Width = Number(width);
    if (Number.isFinite(Number(height))) shape.Height = Number(height);
    if (Number.isFinite(Number(rotation))) shape.Rotation = Number(rotation);
    if (newName) { try { shape.Name = String(newName); } catch (e) { warnings.push(`改名失败: ${e.message}`); } }
    if (visible !== undefined) { try { shape.Visible = visible ? -1 : 0; } catch (e) { warnings.push(`设置可见性失败: ${e.message}`); } }
    if (fillColor) {
      const rgb = hexToExcelColor(fillColor);
      if (rgb === null) warnings.push(`fillColor 无法解析: ${fillColor}`);
      else { try { shape.Fill.Visible = -1; shape.Fill.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置填充失败: ${e.message}`); } }
    }
    if (lineColor) {
      const rgb = hexToExcelColor(lineColor);
      if (rgb === null) warnings.push(`lineColor 无法解析: ${lineColor}`);
      else { try { shape.Line.Visible = -1; shape.Line.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置线条色失败: ${e.message}`); } }
    }
    if (Number.isFinite(Number(lineWeight))) {
      try { shape.Line.Weight = Number(lineWeight); } catch (e) { warnings.push(`设置线宽失败: ${e.message}`); }
    }
    if (text !== undefined && text !== null) {
      try { shape.TextFrame.Characters().Text = String(text); } catch (e) { warnings.push(`写入文字失败: ${e.message}`); }
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      action: "update",
      shape: readShape(shape, sheet),
      warnings,
      message: `已更新形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」`
    };
  }

  /** 分组：`Shapes.Range([名...]).Group()`（WPS 表格可用；Office.js 侧未打通）。 */
  function groupShapes(app, params) {
    const { sheetName, workbookName, names, groupName } = params || {};
    const list = Array.isArray(names) ? names.filter(Boolean).map(String) : [];
    if (list.length < 2) throw new Error("group_shapes 需要至少 2 个形状（names 传形状名数组）");
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    // 先确认都存在，错误信息才有用
    const existing = [];
    for (let i = 1; i <= shapes.Count; i++) { try { existing.push(String(shapes.Item(i).Name)); } catch (e) {} }
    const missing = list.filter(n => existing.indexOf(n) < 0);
    if (missing.length) throw new Error(`这些形状不存在: ${missing.join(", ")}。现有形状：${existing.join(" / ") || "无"}`);

    const group = shapes.Range(list).Group();
    if (groupName) { try { group.Name = String(groupName); } catch (e) {} }
    const members = [];
    try {
      const gi = group.GroupItems;
      for (let i = 1; i <= gi.Count; i++) { try { members.push(String(gi.Item(i).Name)); } catch (e) {} }
    } catch (e) {}
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      group: readShape(group, sheet),
      requestedMembers: list,
      memberCount: members.length,
      members,
      verified: members.length === list.length,
      warnings: members.length === list.length ? [] : [`请求组合 ${list.length} 个，读回组内 ${members.length} 个`],
      message: `已把 ${list.length} 个形状组合为「${(() => { try { return String(group.Name); } catch (e) { return groupName || ""; } })()}」`
    };
  }

  /** 解散分组，返回释放后的形状名单。 */
  function ungroupShapes(app, params) {
    const { sheetName, workbookName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    const type = (() => { try { return Number(shape.Type); } catch (e) { return null; } })();
    if (type !== 6) {
      throw new Error(`形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」不是组合（Type=${type}，组合应为 6）`);
    }
    const before = [];
    try { const gi = shape.GroupItems; for (let i = 1; i <= gi.Count; i++) { try { before.push(String(gi.Item(i).Name)); } catch (e) {} } } catch (e) {}
    shape.Ungroup();
    const after = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) { try { after.push(String(shapes.Item(i).Name)); } catch (e) {} }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      released: before,
      releasedStillPresent: before.filter(n => after.indexOf(n) >= 0),
      shapeCountAfter: shapes.Count,
      message: `已解散组合，释放 ${before.length} 个形状`
    };
  }

  /** 调整层级。VBA ZOrder 常量：0=置顶 1=置底 2=上移一层 3=下移一层。 */
  function setShapeZorder(app, params) {
    const { sheetName, workbookName, zOrder = "bringToFront" } = params || {};
    const Z = { bringtofront: 0, sendtoback: 1, bringforward: 2, sendbackward: 3 };
    const code = Z[String(zOrder).toLowerCase()];
    if (code === undefined) throw new Error(`不支持的 zOrder: ${zOrder}（可用 bringToFront / sendToBack / bringForward / sendBackward）`);
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    shape.ZOrder(code);
    // 读回全部形状以核对层级（WPS 上 ZOrderPosition 不可读，改用"绘制顺序"近似：
    // 按 ZOrder(0) 逐个置顶来还原顺序代价太高，这里只如实返回当前形状与形状总数）
    const names = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) { try { names.push(String(shapes.Item(i).Name)); } catch (e) {} }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      applied: zOrder,
      shape: readShape(shape, sheet),
      shapeNames: names,
      warnings: ["WPS 表格读不到 ZOrderPosition，无法直接回读层级；shapeNames 的顺序是 Shapes 集合顺序，可作为近似参考"],
      message: `已把形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」调整为 ${zOrder}`
    };
  }

  /** 图片：插入 / 列出 / 删除。 */
  function managePictures(app, params) {
    const { sheetName, workbookName, action = "list", filePath, left, top, width, height, pictureName, pictureIndex } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const readOne = (sh, kind) => ({
      name: g(() => String(sh.Name), null), kind,
      left: g(() => Math.round(Number(sh.Left) * 100) / 100, null),
      top: g(() => Math.round(Number(sh.Top) * 100) / 100, null),
      width: g(() => Math.round(Number(sh.Width) * 100) / 100, null),
      height: g(() => Math.round(Number(sh.Height) * 100) / 100, null),
      topLeftCell: g(() => shapeAddressOf(sh.TopLeftCell), null)
    });

    if (action === "list") {
      const items = [];
      const n = g(() => Number(sheet.Shapes.Count), 0);
      for (let i = 1; i <= n; i++) {
        const sh = sheet.Shapes.Item(i);
        const isPic = g(() => Number(sh.Type) === 13, false);   // msoPicture = 13
        if (!isPic) continue;
        items.push(readOne(sh, "picture"));
      }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, count: items.length, pictures: items, warnings: [],
        message: `工作表 [${sheet.Name}] 共 ${items.length} 张图片（形状总数 ${g(() => Number(sheet.Shapes.Count), 0)}）` };
    }

    if (action === "delete") {
      const shapes = sheet.Shapes;
      let removed = 0;
      for (let i = shapes.Count; i >= 1; i--) {
        const sh = shapes.Item(i);
        if (g(() => Number(sh.Type) !== 13, true)) continue;
        const nm = g(() => String(sh.Name), "");
        if (pictureName && nm !== String(pictureName)) continue;
        if (Number.isFinite(Number(pictureIndex)) && i !== Number(pictureIndex)) continue;
        sh.Delete(); removed++;
      }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, removed, warnings: [], message: `已删除 ${removed} 张图片` };
    }

    if (!filePath) throw new Error("insert 需要 filePath（本机图片绝对路径）");
    const args = [String(filePath), false, true];
    if (Number.isFinite(Number(left))) args.push(Number(left));
    if (Number.isFinite(Number(top))) args.push(Number(top));
    if (Number.isFinite(Number(width))) args.push(Number(width));
    if (Number.isFinite(Number(height))) args.push(Number(height));
    const sh = sheet.Shapes.AddPicture.apply(sheet.Shapes, args);
    if (pictureName) { try { sh.Name = String(pictureName); } catch (e) {} }
    return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, action: "insert",
      picture: readOne(sh, "picture"), warnings: [], message: `已插入图片 ${filePath}` };
  }
