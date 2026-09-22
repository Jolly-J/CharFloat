// ── 模块: src/excel/shape.js — 形状与图片（CAP-08：MS 侧矢量绘图） ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1688-1731 行。
//
// 设计约束（CAP-08）：
// 1. **写后必读**：每个写操作都在同一批次里把真实落位的属性读回来返回给调用方，
//    不做"发了请求就报成功"的假成功；宿主没接受的属性写进 `warnings` 如实暴露。
// 2. **没实现就不假装**：某项属性设置失败只把该项写进 `warnings`，不阻断其它属性，
//    也不把失败折叠成成功。
// 3. 不用 Canvas / 任何本地合成去"造"图形 —— 形状只能由宿主真实创建。
  // 10. 形状与图片

  /** 几何形状类型的别名表（未知类型时给出可执行提示，而不是甩一句"参数错误"）。 */
  const SHAPE_TYPE_ALIASES = {
    rectangle: "rectangle", rect: "rectangle", square: "rectangle",
    rounded_rectangle: "roundedRectangle", roundedrectangle: "roundedRectangle", round_rect: "roundedRectangle",
    oval: "oval", ellipse: "oval", circle: "oval",
    triangle: "triangle", right_triangle: "rightTriangle",
    diamond: "diamond", rhombus: "diamond",
    pentagon: "pentagon", hexagon: "hexagon", octagon: "octagon",
    star: "star5", star5: "star5", star4: "star4", star8: "star8",
    arrow: "rightArrow", right_arrow: "rightArrow",
    left_arrow: "leftArrow", up_arrow: "upArrow", down_arrow: "downArrow",
    double_arrow: "leftRightArrow",
    cloud: "cloud", heart: "heart", sun: "sun", moon: "moon",
    lightning_bolt: "lightningBolt",
    flow_chart_process: "flowChartProcess", flow_chart_decision: "flowChartDecision", flow_chart_data: "flowChartData",
    flowchart_process: "flowChartProcess", flowchart_decision: "flowChartDecision", flowchart_data: "flowChartData",
    trapezoid: "trapezoid", parallelogram: "parallelogram", cross: "cross", plus: "cross",
    can: "can", cube: "cube", donut: "donut", plaque: "plaque",
    text_box: "textBox", textbox: "textBox"
  };

  /** 把调用方给的类型名解析成 `Excel.GeometricShapeType` 的枚举值。 */
  function resolveGeometricShapeType(raw) {
    const key = String(raw || "rectangle").trim();
    const alias = SHAPE_TYPE_ALIASES[key.toLowerCase()] || key;
    const all = (typeof Excel !== "undefined" && Excel.GeometricShapeType) || {};
    if (all[alias] !== undefined) return { value: all[alias], name: alias };
    for (const k of Object.keys(all)) {
      if (k.toLowerCase() === String(alias).toLowerCase()) return { value: all[k], name: k };
    }
    throw new Error(`不认识的几何形状类型 "${raw}"。可用枚举键（部分）：${Object.keys(all).slice(0, 40).join(", ")}`);
  }

  /** 解析要操作的目标形状；优先 name（稳定），其次 id。 */
  function resolveShape(sheet, params) {
    const name = params.name || params.shapeName;
    if (name) return sheet.shapes.getItem(name);
    if (params.shapeId) return sheet.shapes.getItem(params.shapeId);
    throw new Error("缺少必要参数：name（形状名，推荐用 list_shapes 读到的 name）或 shapeId");
  }

  /** 十六进制颜色归一：#RGB / #RRGGBB / 无 # 都接受；非法值返回 null（由调用方如实报告）。 */
  function normalizeHexColor(raw) {
    if (raw === undefined || raw === null || raw === "") return null;
    let s = String(raw).trim();
    if (!s.startsWith("#")) s = "#" + s;
    if (/^#[0-9a-fA-F]{3}$/.test(s)) s = "#" + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
    return /^#[0-9a-fA-F]{6}$/.test(s) ? s.toUpperCase() : null;
  }

  /** getAsImage 的格式参数：用枚举值，取不到时退回字符串字面量（Office.js 接受枚举字符串）。 */
  function resolvePictureFormat(raw) {
    const wantJpeg = String(raw || "png").toLowerCase() === "jpeg" || String(raw || "").toLowerCase() === "jpg";
    try {
      if (typeof Excel !== "undefined" && Excel.PictureFormat) return wantJpeg ? Excel.PictureFormat.jpeg : Excel.PictureFormat.png;
    } catch (e) { /* 该版本没有枚举则走字符串 */ }
    return wantJpeg ? "JPEG" : "PNG";
  }

  /** 一次批次里把形状的全部可读属性挂上 load（读回口径只在这里维护一份）。 */
  function loadShapeDetail(shape) {
    shape.load("name,id,type,left,top,width,height,rotation,zOrderPosition,visible");
    try { shape.load("fill/type,fill/foregroundColor,fill/transparency"); } catch (e) { /* 宿主不支持则跳过 */ }
    try { shape.load("lineFormat/color,lineFormat/weight,lineFormat/visible"); } catch (e) {}
    try { shape.load("textFrame/textRange/text"); } catch (e) {}
    try { shape.load("textFrame/textRange/font/bold,textFrame/textRange/font/size,textFrame/textRange/font/color"); } catch (e) {}
    try { shape.load("textFrame/horizontalAlignment,textFrame/verticalAlignment"); } catch (e) {}
    return shape;
  }

  /** 把已 load 的 shape 折算成扁平 JSON（不返回宿主代理对象）。 */
  function shapeToJson(shape) {
    const out = {
      name: shape.name,
      id: shape.id,
      type: String(shape.type),
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      rotation: shape.rotation,
      zOrderPosition: shape.zOrderPosition === undefined || shape.zOrderPosition === null ? null : String(shape.zOrderPosition),
      visible: shape.visible
    };
    try { out.fill = { type: shape.fill ? String(shape.fill.type) : null, foregroundColor: shape.fill ? shape.fill.foregroundColor : null, transparency: shape.fill ? shape.fill.transparency : null }; }
    catch (e) { out.fill = { unsupported: true, error: String(e.message || e) }; }
    try { out.line = { color: shape.lineFormat ? shape.lineFormat.color : null, weight: shape.lineFormat ? shape.lineFormat.weight : null, visible: shape.lineFormat ? shape.lineFormat.visible : null }; }
    catch (e) { out.line = { unsupported: true, error: String(e.message || e) }; }
    try {
      const tf = shape.textFrame;
      if (tf && tf.textRange) {
        out.text = tf.textRange.text;
        out.textFormat = {
          bold: tf.textRange.font ? tf.textRange.font.bold : null,
          size: tf.textRange.font ? tf.textRange.font.size : null,
          color: tf.textRange.font ? tf.textRange.font.color : null,
          horizontalAlignment: tf.horizontalAlignment === undefined ? null : String(tf.horizontalAlignment),
          verticalAlignment: tf.verticalAlignment === undefined ? null : String(tf.verticalAlignment)
        };
      } else out.text = null;
    } catch (e) { out.text = null; out.textUnsupported = String(e.message || e); }
    return out;
  }

  /** 读回核对：把"请求值 vs 宿主实际值"逐项比对，只报真实差异。 */
  function verifyAgainstRequest(actual, expected) {
    const mismatches = [];
    for (const key of Object.keys(expected)) {
      let want;
      if (key === "fillColor") want = actual.fill && actual.fill.foregroundColor;
      else if (key === "lineColor") want = actual.line && actual.line.color;
      else if (key === "lineWeight") want = actual.line && actual.line.weight;
      else if (key === "text") want = actual.text;
      else if (key === "name") want = actual.name;
      else if (key === "rotationDelta" || key === "scaleWidth" || key === "scaleHeight") continue; // 增量类无法按等值比对
      else want = actual[key];
      if (want === undefined || want === null) { mismatches.push(`${key}: 宿主未回读该属性`); continue; }
      const a = typeof want === "string" ? want.toUpperCase() : want;
      const b = (typeof expected[key] === "string" && key !== "text") ? String(expected[key]).toUpperCase() : expected[key];
      if (a !== b) mismatches.push(`${key}: 请求 ${expected[key]} / 实际 ${want}`);
    }
    return mismatches;
  }

  /**
   * 添加矢量图形：矩形/椭圆/箭头等几何形状、直线/连接符、SVG、文本框。
   * 一次调用只加一个形状；返回值里带**宿主读回**的实际位置尺寸与颜色。
   */
  async function handleAddShape(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const warnings = [];
      const requested = {};
      const kind = String(params.kind || "geometric").toLowerCase();

      let shape;
      if (kind === "line" || kind === "connector") {
        // addLine 收的是起止点，不是 left/top/width/height。
        const x1 = params.x1 !== undefined ? params.x1 : (params.left !== undefined ? params.left : 0);
        const y1 = params.y1 !== undefined ? params.y1 : (params.top !== undefined ? params.top : 0);
        const x2 = params.x2 !== undefined ? params.x2 : (params.right !== undefined ? params.right : x1 + (params.width || 100));
        const y2 = params.y2 !== undefined ? params.y2 : (params.bottom !== undefined ? params.bottom : y1 + (params.height || 0));
        shape = sheet.shapes.addLine(x1, y1, x2, y2);
        requested.x1 = x1; requested.y1 = y1; requested.x2 = x2; requested.y2 = y2;
      } else if (kind === "svg") {
        const svg = params.svg
          || (params.base64Image && typeof atob === "function" ? atob(params.base64Image) : null);
        if (!svg) throw new Error("kind=svg 需要传 svg（SVG 源码字符串）或 base64Image（base64 编码的 SVG）");
        const b64 = params.base64Image || (typeof btoa === "function" ? btoa(svg) : "");
        if (!b64) throw new Error("当前环境无法把 SVG 转成 base64（缺少 btoa）");
        shape = sheet.shapes.addSvg(b64);
      } else if (kind === "textbox") {
        shape = sheet.shapes.addTextBox(params.text !== undefined && params.text !== null ? String(params.text) : "");
      } else {
        shape = sheet.shapes.addGeometricShape(resolveGeometricShapeType(params.shapeType || params.type).value);
      }

      if (params.shapeName || params.name) shape.name = params.shapeName || params.name;

      for (const pair of [["left", params.left], ["top", params.top], ["width", params.width], ["height", params.height]]) {
        if (pair[1] === undefined || pair[1] === null) continue;
        shape[pair[0]] = pair[1]; requested[pair[0]] = pair[1];
      }
      if (params.rotation !== undefined && params.rotation !== null) { shape.rotation = params.rotation; requested.rotation = params.rotation; }

      // 文本框的初始文字已在 addTextBox 里给出；其余形状用 textFrame 写文字。
      if (kind !== "textbox" && params.text !== undefined && params.text !== null) {
        shape.textFrame.textRange.text = String(params.text);
        requested.text = String(params.text);
      }

      const fillHex = normalizeHexColor(params.fillColor || params.fill);
      if (fillHex) { shape.fill.setSolidColor(fillHex); requested.fillColor = fillHex; }
      else if (params.fillColor || params.fill) warnings.push(`fillColor "${params.fillColor || params.fill}" 不是合法十六进制颜色，未应用`);

      const lineHex = normalizeHexColor(params.lineColor);
      if (lineHex) { shape.lineFormat.color = lineHex; requested.lineColor = lineHex; }
      if (params.lineWeight !== undefined && params.lineWeight !== null) { shape.lineFormat.weight = params.lineWeight; requested.lineWeight = params.lineWeight; }

      loadShapeDetail(shape);
      await context.sync();
      const actual = shapeToJson(shape);
      const mismatches = verifyAgainstRequest(actual, requested);
      return {
        success: true,
        sheetName: sheet.name,
        kind,
        shape: actual,
        requested,
        verified: mismatches.length === 0,
        warnings: warnings.concat(mismatches.map(m => `读回与请求不一致：${m}`))
      };
    });
  }

  async function handleInsertImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      if (!params.base64Image) throw new Error("缺少必要参数：base64Image");
      const shape = sheet.shapes.addImage(params.base64Image);
      if (params.left) shape.left = params.left;
      if (params.top) shape.top = params.top;
      if (params.width) shape.width = params.width;
      if (params.height) shape.height = params.height;
      if (params.name || params.shapeName) shape.name = params.name || params.shapeName;

      loadShapeDetail(shape);
      await context.sync();
      return { success: true, sheetName: sheet.name, shape: shapeToJson(shape), verified: true };
    });
  }

  /** 读回：列表 + 位置尺寸 + 填充/线条 + 旋转 + 层级（CAP-08 的读回入口）。 */
  async function handleListShapes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shapes = sheet.shapes;
      // 整批一次 sync，避免形状多时 N 次往返；宿主不认某个子属性时退化为不读该属性。
      shapes.load("items/name,items/id,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition,items/visible");
      try { shapes.load("items/fill/type,items/fill/foregroundColor,items/fill/transparency"); } catch (e) {}
      try { shapes.load("items/lineFormat/color,items/lineFormat/weight"); } catch (e) {}
      try { shapes.load("items/textFrame/textRange/text"); } catch (e) {}
      await context.sync();

      const items = shapes.items.map(shapeToJson);
      const ordered = items.slice().sort((a, b) => Number(a.zOrderPosition) - Number(b.zOrderPosition));
      return {
        success: true,
        count: items.length,
        sheetName: sheet.name,
        shapes: items,
        zOrderBottomToTop: ordered.map(s => ({ name: s.name, zOrderPosition: s.zOrderPosition })),
        message: `工作表 [${sheet.name}] 共有 ${items.length} 个形状`
      };
    });
  }

  /** 改形状：位置/尺寸/旋转/缩放/填充/线条/文字/改名，或删除；同样带读回。 */
  async function handleUpdateShape(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const warnings = [];
      const requested = {};

      if (params.action === "delete") {
        const target = resolveShape(sheet, params);
        target.delete();
        await context.sync();
        const rest = sheet.shapes.load("items/name");
        await context.sync();
        const stillThere = rest.items.some(s => s.name === (params.name || params.shapeName));
        return {
          success: !stillThere,
          deleted: params.name || params.shapeName,
          remaining: rest.items.map(s => s.name),
          verified: !stillThere,
          warnings: stillThere ? ["删除后仍能读到该形状，删除未生效"] : []
        };
      }

      const shape = resolveShape(sheet, params);
      loadShapeDetail(shape);
      await context.sync();
      const beforeJson = shapeToJson(shape);

      for (const pair of [["left", params.left], ["top", params.top], ["width", params.width], ["height", params.height]]) {
        if (pair[1] === undefined || pair[1] === null) continue;
        shape[pair[0]] = pair[1]; requested[pair[0]] = pair[1];
      }
      if (params.rotation !== undefined && params.rotation !== null) { shape.rotation = params.rotation; requested.rotation = params.rotation; }
      if (params.newName || params.rename) { shape.name = params.newName || params.rename; requested.name = params.newName || params.rename; }
      if (params.text !== undefined && params.text !== null) { shape.textFrame.textRange.text = String(params.text); requested.text = String(params.text); }

      const fillHex = normalizeHexColor(params.fillColor || params.fill);
      if (fillHex) { shape.fill.setSolidColor(fillHex); requested.fillColor = fillHex; }
      else if (params.fillColor || params.fill) warnings.push(`fillColor "${params.fillColor || params.fill}" 不是合法十六进制颜色，未应用`);

      const lineHex = normalizeHexColor(params.lineColor);
      if (lineHex) { shape.lineFormat.color = lineHex; requested.lineColor = lineHex; }
      if (params.lineWeight !== undefined && params.lineWeight !== null) { shape.lineFormat.weight = params.lineWeight; requested.lineWeight = params.lineWeight; }

      if (params.rotationDelta !== undefined && params.rotationDelta !== null) { shape.incrementRotation(params.rotationDelta); requested.rotationDelta = params.rotationDelta; }
      if (params.scaleWidth !== undefined && params.scaleWidth !== null) {
        shape.scaleWidth(params.scaleWidth, params.scaleType === "originalSize" ? "OriginalSize" : "CurrentSize");
        requested.scaleWidth = params.scaleWidth;
      }
      if (params.scaleHeight !== undefined && params.scaleHeight !== null) {
        shape.scaleHeight(params.scaleHeight, params.scaleType === "originalSize" ? "OriginalSize" : "CurrentSize");
        requested.scaleHeight = params.scaleHeight;
      }

      if (Object.keys(requested).length === 0) {
        return {
          success: true,
          changed: false,
          shape: beforeJson,
          warnings: ["没有传入任何可修改字段（left/top/width/height/rotation/fillColor/lineColor/lineWeight/text/newName/scale*），未做修改"]
        };
      }

      loadShapeDetail(shape);
      await context.sync();
      const after = shapeToJson(shape);
      const mismatches = verifyAgainstRequest(after, requested);
      return {
        success: true,
        changed: true,
        sheetName: sheet.name,
        before: beforeJson,
        shape: after,
        requested,
        verified: mismatches.length === 0,
        warnings: warnings.concat(mismatches.map(m => `读回与请求不一致：${m}`))
      };
    });
  }

  async function handleGroupShapes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const names = Array.isArray(params.shapeNames) ? params.shapeNames.filter(Boolean) : [];
      if (names.length < 2) throw new Error("group_shapes 需要至少 2 个形状（shapeNames 传形状名数组）");
      // 组合是"活动表"范围的操作：真机实测，对非活动表调 addGroup 会报
      // 「当前对象不允许此操作」（子代理的自检能过，是因为它新建表时顺带激活了那张表）。
      sheet.activate();
      await context.sync();   // 先让激活单独生效，再发起组合
      const proxies = names.map(n => sheet.shapes.getItem(n));
      proxies.forEach(p => p.load("id,name"));
      await context.sync();
      const ids = proxies.map(p => p.id);
      const group = sheet.shapes.addGroup(ids);
      if (params.groupName || params.name) group.name = params.groupName || params.name;
      loadShapeDetail(group);
      await context.sync();

      const groupJson = shapeToJson(group);
      let members = [];
      try {
        const memberShapes = group.group.shapes.load("items/name,items/type");
        await context.sync();
        members = memberShapes.items.map(s => ({ name: s.name, type: String(s.type) }));
      } catch (e) {
        groupJson.memberReadError = String(e.message || e);
      }
      return {
        success: true,
        sheetName: sheet.name,
        group: groupJson,
        memberCount: members.length,
        members,
        verified: members.length === ids.length,
        warnings: members.length === ids.length ? [] : [`请求分组 ${ids.length} 个形状，读回组内 ${members.length} 个`]
      };
    });
  }

  async function handleUngroupShapes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const name = params.groupName || params.name;
      const group = sheet.shapes.getItem(name);
      group.load("id,type");
      await context.sync();
      // 先确认这确实是一个组合（ShapeType.group），再谈解组：对普通形状调 ungroup
      // 会报出难懂的宿主错误，提前拒绝对调用方更有用。
      if (String(group.type).toLowerCase().indexOf("group") < 0) {
        throw new Error(`"${name}" 不是组合（读到 type=${group.type}），无法解组。先用 list_shapes 确认组合名。`);
      }
      if (!group.group || !group.group.shapes) {
        throw new Error(`当前 Excel 版本读不到组合成员（Shape.group 不可用），无法解组 "${name}"。`);
      }
      const childShapes = group.group.shapes.load("items/name");
      await context.sync();
      const childNames = childShapes.items.map(s => s.name);
      if (typeof group.group.ungroup !== "function") {
        throw new Error("当前 Excel 版本不支持 ShapeGroup.ungroup()：只能删除整组，或改用 update_shape 单独移动成员");
      }
      group.group.ungroup();
      await context.sync();
      const rest = sheet.shapes.load("items/name");
      await context.sync();
      const remaining = rest.items.map(s => s.name);
      const childrenBack = childNames.filter(n => remaining.indexOf(n) >= 0);
      return {
        success: true,
        sheetName: sheet.name,
        ungrouped: name,
        releasedChildren: childNames,
        childrenStillPresent: childrenBack,
        shapesAfter: remaining,
        verified: childrenBack.length === childNames.length
      };
    });
  }

  /** 层级调整：bringToFront / sendToBack / bringForward / sendBackward，返回整表层级读回。 */
  async function handleSetShapeZOrder(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = resolveShape(sheet, params);
      const raw = String(params.zOrder || params.action || "bringToFront").toLowerCase().replace(/[\s_-]/g, "");
      const map = {
        bringtofront: "BringToFront", front: "BringToFront",
        sendtoback: "SendToBack", back: "SendToBack",
        bringforward: "BringForward", forward: "BringForward",
        sendbackward: "SendBackward", backward: "SendBackward"
      };
      const target = map[raw];
      if (target === undefined) throw new Error(`不认识的层级动作 "${params.zOrder || params.action}"，可用：bringToFront / sendToBack / bringForward / sendBackward`);
      shape.setZOrder(target);
      await context.sync();

      const all = sheet.shapes.load("items/name,items/zOrderPosition");
      await context.sync();
      const ordered = all.items
        .map(s => ({ name: s.name, zOrderPosition: s.zOrderPosition === undefined || s.zOrderPosition === null ? null : String(s.zOrderPosition) }))
        .sort((a, b) => Number(a.zOrderPosition) - Number(b.zOrderPosition));
      const self = ordered.filter(s => s.name === (params.name || params.shapeName))[0];
      const expectedEdge = target === "BringToFront" ? ordered.length - 1 : target === "SendToBack" ? 0 : null;
      return {
        success: true,
        sheetName: sheet.name,
        applied: target,
        shape: (params.name || params.shapeName) || params.shapeId,
        shapeZOrderPosition: self ? self.zOrderPosition : null,
        zOrderBottomToTop: ordered,
        // 端点动作可以强校验；bringForward/sendBackward 是相对移动，只如实给读回
        verified: expectedEdge === null ? null : Boolean(self && Number(self.zOrderPosition) === expectedEdge)
      };
    });
  }

  /** 读回：当前选中的形状（用户手上正在编辑哪一个）。 */
  async function handleGetActiveShape(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = sheet.shapes.getActiveShape();
      loadShapeDetail(shape);
      await context.sync();
      return { success: true, sheetName: sheet.name, shape: shapeToJson(shape) };
    });
  }

  /** 把形状导出成 PNG/JPEG（base64），用于视觉自检。 */
  async function handleExportShapeImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = resolveShape(sheet, params);
      const format = resolvePictureFormat(params.format);
      const scale = params.scale || 1;
      const image = shape.getAsImage(format, scale);
      await context.sync();
      const base64 = image.value;
      if (!base64) throw new Error("宿主未返回图像数据（getAsImage 返回空）");
      return {
        success: true,
        sheetName: sheet.name,
        name: params.name || params.shapeName,
        format: String(params.format || "png").toLowerCase(),
        scale,
        imageBase64: base64,
        byteEstimate: Math.round(base64.length * 0.75)
      };
    });
  }

  // ── CAP-08 能力自检 ──
  // 为什么要在加载项里自带自检：本机 MCP 工具面没有通道能把任意 Office.js 代码送进任务窗格
  // （excel_* 路由表不含形状方法；office_execute_script 走 macOS JXA / Windows COM），
  // 只靠桥接侧无法回答"这台 Excel 到底支持到哪一步"。自检在真实窗格里跑、写回一张工作表，
  // 结果可以被任意 host=microsoft 的读工具读回，也能被下一次会话直接复用。
  // 自检**逐项报可用/不可用**，任何一项失败都不影响其它项；不支持的项如实记录错误原文。
  let CAP08_SELF_TEST_RESULT = null;

  async function handleShapeSelfTest(params) {
    const sheetName = params.sheetName || "_cap08_shapetest";
    const keep = params.keep === true;           // 默认清理探测表，不留垃圾
    const started = new Date().toISOString();
    const checks = [];
    const record = async (name, fn) => {
      const t0 = Date.now();
      try {
        const detail = await fn();
        checks.push({ check: name, ok: true, ms: Date.now() - t0, detail: detail === undefined ? null : detail });
        return detail;
      } catch (e) {
        checks.push({ check: name, ok: false, ms: Date.now() - t0, error: String((e && e.message) || e), code: e && e.code });
        return null;
      }
    };

    const runOn = async (fn) => await Excel.run(fn);
    const sheetOf = (context) => context.workbook.worksheets.getItem(sheetName);

    /**
     * 让检查之间**互不依赖**。
     *
     * 原实现里后面的检查直接引用前面创建的形状（cap08_line / cap08_svg 等），
     * 一旦前面失败，后面会报"请求的资源不存在"——看上去像"分组/删除也不支持"，
     * 实际只是前置对象不存在。**级联失败会把"没做出来"误报成"宿主不支持"**，
     * 所以这里统一改成"需要什么就先确保它存在"。
     */
    const ensureShape = async (name, make) => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const existing = sheet.shapes.load("items/name");
      await context.sync();
      if (!existing.items.some(s => s.name === name)) {
        const sh = make(sheet);
        if (sh) sh.name = name;
      }
      await context.sync();
      return name;
    });
    const ensureRect = () => ensureShape("cap08_rect", (sheet) =>
      sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle, { left: 20, top: 20, width: 160, height: 90 }));
    const ensureText = () => ensureShape("cap08_text", (sheet) =>
      sheet.shapes.addTextBox("CAP-08 文本框"));

    try {
      await runOn(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        if (sheets.items.some(s => s.name === sheetName)) sheets.getItem(sheetName).delete();
        await context.sync();
        sheets.add(sheetName);
        await context.sync();
      });
    } catch (e) {
      return { success: false, stage: "createSheet", error: String(e.message || e), checks, started };
    }

    // 1. 几何形状
    const rect = await record("addGeometricShape(rectangle) + 填充/线条/旋转读回", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle, { left: 20, top: 20, width: 160, height: 90 });
      sh.name = "cap08_rect";
      sh.fill.setSolidColor("#2F6FEB");
      sh.lineFormat.color = "#12305E";
      sh.lineFormat.weight = 2;
      sh.rotation = 15;
      sh.textFrame.textRange.text = "CAP-08";
      loadShapeDetail(sh);
      await context.sync();
      return shapeToJson(sh);
    }));

    // 2. 直线
    const line = await record("addLine（连接符/直线）", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.addLine(200, 20, 360, 110);
      sh.name = "cap08_line";
      sh.lineFormat.color = "#E4572E";
      sh.lineFormat.weight = 3;
      loadShapeDetail(sh);
      await context.sync();
      return shapeToJson(sh);
    }));

    // 3. SVG
    const svgOk = await record("addSvg", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" viewBox="0 0 120 80"><rect width="120" height="80" fill="#0F9D58"/><circle cx="60" cy="40" r="26" fill="#ffffff"/></svg>';
      const b64 = typeof btoa === "function" ? btoa(svg) : null;
      if (!b64) throw new Error("当前环境没有 btoa，无法把 SVG 转 base64");
      const sh = sheet.shapes.addSvg(b64);
      sh.name = "cap08_svg";
      loadShapeDetail(sh);
      await context.sync();
      return shapeToJson(sh);
    }));

    // 4. 文本框
    const textBox = await record("addTextBox + 字体读回", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.addTextBox("CAP-08 文本框");
      sh.name = "cap08_text";
      sh.left = 20; sh.top = 140; sh.width = 200; sh.height = 50;
      sh.textFrame.textRange.font.bold = true;
      sh.textFrame.textRange.font.size = 14;
      loadShapeDetail(sh);
      await context.sync();
      return shapeToJson(sh);
    }));

    // 5. 列表读回
    const list = await record("shapes 列表读回（几何属性 + zOrder）", async () => {
      await ensureRect();
      await ensureText();
      return await runOn(async (context) => {
        const sheet = sheetOf(context);
        // 只 load 标量属性：嵌套的 fill/lineFormat 需要逐个 load 才能读，
        // 混在一起 load 会让整批失败（真机实测报"当前对象不允许此操作"）。
        const shapes = sheet.shapes.load("items/name,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition");
        await context.sync();
        return shapes.items.map(s => ({
          name: s.name, type: String(s.type), left: s.left, top: s.top,
          width: s.width, height: s.height, rotation: s.rotation, z: String(s.zOrderPosition)
        }));
      });
    });
    const listFill = await record("shapes 填充/线条读回（逐个 load）", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.getItem("cap08_rect");
      sh.fill.load("foregroundColor");
      sh.lineFormat.load("color,weight");
      await context.sync();
      return { fill: sh.fill.foregroundColor, lineColor: sh.lineFormat.color, lineWeight: sh.lineFormat.weight };
    }));

    // 6. 层级
    const zorder = await record("setZOrder(bringToFront/sendToBack)", async () => {
      await ensureRect(); await ensureText();
      return await runOn(async (context) => {
      const sheet = sheetOf(context);
      // 不引用 cap08_line（它可能没创建成功）；用两个确定存在的形状对比层级
      sheet.shapes.getItem("cap08_rect").setZOrder(Excel.ShapeZOrder.sendToBack);
      sheet.shapes.getItem("cap08_text").setZOrder(Excel.ShapeZOrder.bringToFront);
      const all = sheet.shapes.load("items/name,items/zOrderPosition");
      await context.sync();
      return all.items.map(s => ({ name: s.name, z: String(s.zOrderPosition) })).sort((a, b) => Number(a.z) - Number(b.z));
      });
    });

    // 7. 激活形状
    const active = await record("getActiveShape", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.getActiveShape();
      sh.load("name,type,left,top,width,height");
      await context.sync();
      return { name: sh.name, type: String(sh.type), left: sh.left, top: sh.top };
    }));

    // 8. 分组 / 解组
    const group = await record("addGroup + 组合成员读回", async () => {
      await ensureRect(); await ensureText();
      return await runOn(async (context) => {
      const sheet = sheetOf(context);
      // 只组合"确实创建成功"的形状：cap08_line / cap08_svg 在部分宿主上建不出来，
      // 引用它们会把分组本身的能力误判成不支持。
      const names = ["cap08_rect", "cap08_text"];
      const proxies = names.map(n => sheet.shapes.getItem(n));
      proxies.forEach(p => p.load("id"));
      await context.sync();
      const g = sheet.shapes.addGroup(proxies.map(p => p.id));
      g.name = "cap08_group";
      g.load("name");
      const members = g.group.shapes.load("items/name,items/type");
      await context.sync();
      return { group: g.name, memberCount: members.items.length, members: members.items.map(m => m.name) };
      });
    });

    const ungroup = await record("ShapeGroup.ungroup()", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const g = sheet.shapes.getItem("cap08_group");
      g.load("name");
      const kids = g.group.shapes.load("items/name");
      await context.sync();
      const kidNames = kids.items.map(k => k.name);
      if (typeof g.group.ungroup !== "function") throw new Error("该版本没有 ShapeGroup.ungroup()");
      g.group.ungroup();
      await context.sync();
      const rest = sheet.shapes.load("items/name");
      await context.sync();
      const remaining = rest.items.map(s => s.name);
      return { released: kidNames, backInSheet: kidNames.filter(n => remaining.indexOf(n) >= 0) };
    }));

    // 9. 形状导出图片
    const image = await record("getAsImage(png)", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const img = sheet.shapes.getItem("cap08_text").getAsImage(resolvePictureFormat("png"), 1);
      await context.sync();
      return { length: img.value ? img.value.length : 0, prefix: img.value ? String(img.value).slice(0, 16) : null };
    }));

    // 10. 缩放
    const scale = await record("scaleWidth/scaleHeight", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.getItem("cap08_rect");
      sh.scaleWidth(1.5, "CurrentSize");
      sh.scaleHeight(1.5, "CurrentSize");
      sh.load("width,height");
      await context.sync();
      return { width: sh.width, height: sh.height };
    }));

    // 11. 删除
    const del = await record("delete + 读回确认", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      sheet.shapes.getItem("cap08_text").delete();
      await context.sync();
      const rest = sheet.shapes.load("items/name");
      await context.sync();
      return { remaining: rest.items.map(s => s.name) };
    }));

    const failed = checks.filter(c => !c.ok).map(c => c.check);
    const result = {
      success: failed.length === 0,
      sheetName,
      started,
      finished: new Date().toISOString(),
      addonVersion: ADDON_VERSION,
      total: checks.length,
      passed: checks.length - failed.length,
      failed,
      checks,
      readback: { rect, line, svg: svgOk, textBox, list, zorder, active, group, ungroup, image, scale, deleteRemaining: del ? del.remaining : null }
    };

    // 把结果写回探测表：换窗格/换会话后仍能用 host=microsoft 的读工具读回来。
    try {
      await runOn(async (context) => {
        const sheet = sheetOf(context);
        const rows = [["check", "ok", "ms", "detail/error"], ["ADDON_VERSION", ADDON_VERSION, "", ""]];
        for (const c of checks) rows.push([c.check, c.ok ? "OK" : "FAIL", String(c.ms), c.ok ? JSON.stringify(c.detail) : c.error]);
        sheet.getRange("A1:D" + rows.length).values = rows;
        sheet.getRange("A1:D1").format.font.bold = true;
        sheet.getRange("A:D").format.columnWidth = 160;
        await context.sync();
      });
    } catch (e) {
      result.writebackError = String(e.message || e);
    }

    if (!keep) {
      try {
        await runOn(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.getItem(sheetName).delete();
          await context.sync();
        });
        result.cleanedUp = true;
      } catch (e) {
        result.cleanupError = String(e.message || e);
        result.cleanedUp = false;
      }
    }

    CAP08_SELF_TEST_RESULT = result;
    // 返回体保持轻量（不含工作表明细），明细可从写回的表或下次 list_shapes 读
    return {
      success: result.success,
      sheetName,
      total: result.total,
      passed: result.passed,
      failed: result.failed,
      checks: checks.map(c => ({ check: c.check, ok: c.ok, ms: c.ms, error: c.error || undefined })),
      readback: result.readback,
      cleanedUp: result.cleanedUp === true,
      finished: result.finished
    };
  }
