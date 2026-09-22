// ── 模块: src/excel/sheets.js — 工作表生命周期 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 532-613 行（原样搬迁，未改写）。
  // 2. 工作表操作
  async function handleListSheets() {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets.load("items/name, items/visibility, items/tabColor, items/position");
      await context.sync();
      return {
        sheets: sheets.items.map(s => ({ name: s.name, visibility: s.visibility, tabColor: s.tabColor, position: s.position }))
      };
    });
  }

  async function handleAddSheet(params) {
    return await Excel.run(async (context) => {
      const name = params.name || params.sheetName;
      const sheet = context.workbook.worksheets.add(name);
      sheet.load("name, position");
      await context.sync();
      return { success: true, name: sheet.name, position: sheet.position };
    });
  }

  async function handleUpdateSheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName || params.oldName);
      if (params.name || params.newName) sheet.name = params.name || params.newName;
      if (params.visibility) sheet.visibility = params.visibility;
      if (params.tabColor || params.color) sheet.tabColor = params.tabColor || params.color;
      if (params.activate) sheet.activate();
      await context.sync();
      return { success: true };
    });
  }

  async function handleDeleteSheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName || params.name);
      sheet.delete();
      await context.sync();
      return { success: true };
    });
  }

  async function handleCopySheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sourceSheet);
      const copied = sheet.copy(params.positionType || "After", sheet);
      if (params.newName) copied.name = params.newName;
      copied.load("name");
      await context.sync();
      return { success: true, name: copied.name };
    });
  }

  /** 取整数，失败返回 null（用于把 targetIndex / position / index 归一）。 */
  function readSheetNumber(...candidates) {
    for (const raw of candidates) {
      if (raw === undefined || raw === null || raw === "") continue;
      const n = Number(raw);
      if (Number.isFinite(n)) return n;
    }
    return null;
  }

  async function handleManageSheet(params) {
    return await Excel.run(async (context) => {
      const rawAction = String(params.action || "").trim();
      const action = rawAction || (params.newSheetName || params.newName ? "rename" : "");
      if (!action) {
        throw new Error(
          '[Office.js 通道] manage_sheet 缺少 action：支持 rename | move | tab_color | protect | unprotect。' +
          '原实现在缺失 action 时会返回 success 却什么都不做，已阻断。'
        );
      }

      const sheet = getTargetSheet(context, params.sheetName || params.oldName || params.sourceSheet);
      sheet.load("name");
      await context.sync();
      const sheetName = sheet.name;

      // 归一化 action 别名：网关用 WPS 语义（rename/move/tab_color/protect/unprotect），
      // 早期 Office.js 实现用 copy/hide/show/color —— 两套都认，避免任何一侧改名后静默 no-op（ISS-93）。
      const actionAliases = {
        copy: "copy", duplicate: "copy", clone: "copy",
        rename: "rename", set_name: "rename",
        move: "move", activate: "move", set_position: "move", reorder: "move",
        tab_color: "tab_color", color: "tab_color", set_tab_color: "tab_color",
        protect: "protect", unprotect: "unprotect",
        hide: "hide", show: "show", unhide: "show", visible: "show"
      };
      const normalized = actionAliases[action.toLowerCase()];

      const unsupported = {
        protect: '[Office.js 通道] manage_sheet(action="protect") 在 Microsoft 宿主未实现：Excel Office.js 的 WorksheetProtection 在 Excel on Mac 桌面版不可用/不可靠，本通道拒绝执行以免回报假成功。替代路径：① 改用 host="wps" 的 wps_manage_sheet(action="protect")；② 在 Excel 里用「审阅 → 保护工作表」手动加保护，然后 read_range 读回确认。',
        unprotect: '[Office.js 通道] manage_sheet(action="unprotect") 在 Microsoft 宿主未实现（同上）。替代路径：① 改用 host="wps" 的 wps_manage_sheet(action="unprotect")；② 在 Excel 里用「审阅 → 撤销工作表保护」手动解除。'
      };

      if (!normalized) {
        throw new Error(
          `[Office.js 通道] manage_sheet 无法识别的 action: "${rawAction}"。` +
          '支持 rename | move | tab_color | protect | unprotect（另兼容 copy/duplicate/activate/color/hide/show）。' +
          '原实现对未知 action 返回 success 但什么都不做，已改为显式报错。'
        );
      }
      if (unsupported[normalized]) throw new Error(unsupported[normalized]);

      if (normalized === "copy") {
        const copied = sheet.copy(Excel.WorksheetPositionType.after, sheet);
        const newName = params.newSheetName || params.newName;
        if (newName) copied.name = newName;
        copied.load("name, position");
        await context.sync();
        return { success: true, action: "copy", sheetName: copied.name, sourceSheetName: sheetName, position: copied.position };
      }

      if (normalized === "rename") {
        const newName = params.newName || params.newSheetName;
        if (!newName) {
          throw new Error('[Office.js 通道] manage_sheet(action="rename") 缺少 newName：重命名必须给出新名称，未执行任何修改。');
        }
        sheet.name = newName;
        sheet.load("name");
        await context.sync();
        return { success: true, action: "rename", sheetName: sheet.name, previousName: sheetName };
      }

      if (normalized === "move") {
        // targetIndex 从 1 开始；Office.js 的 Worksheet.position 从 0 开始。
        const targetIndex = readSheetNumber(params.targetIndex, params.position, params.index, params.targetPosition);
        if (targetIndex === null) {
          throw new Error(
            '[Office.js 通道] manage_sheet(action="move") 缺少 targetIndex：请给出目标位置序号（从 1 开始，1=最前）。' +
            '原实现会返回 success 但不移动工作表，已阻断。'
          );
        }
        const total = context.workbook.worksheets.getCount();
        await context.sync();
        if (targetIndex < 1 || targetIndex > total.value) {
          throw new Error(`[Office.js 通道] manage_sheet(action="move") 的 targetIndex=${targetIndex} 越界：当前工作簿共 ${total.value} 张工作表（合法范围 1..${total.value}）。`);
        }
        sheet.position = Math.round(targetIndex) - 1;
        sheet.load("name, position");
        await context.sync();
        return { success: true, action: "move", sheetName: sheet.name, position: sheet.position, targetIndex: Math.round(targetIndex) };
      }

      if (normalized === "tab_color") {
        const color = params.color || params.tabColor;
        if (!color) {
          throw new Error(
            '[Office.js 通道] manage_sheet(action="tab_color") 缺少 color：请给出十六进制标签底色（如 "#EF4444"）。' +
            '原实现只在 params.tabColor 存在时才设色，网关传的是 color → 静默 no-op，现已阻断。'
          );
        }
        sheet.tabColor = color;
        sheet.load("name, tabColor");
        await context.sync();
        return { success: true, action: "tab_color", sheetName: sheet.name, color: sheet.tabColor };
      }

      if (normalized === "hide" || normalized === "show") {
        sheet.visibility = normalized === "hide" ? Excel.SheetVisibility.hidden : Excel.SheetVisibility.visible;
        sheet.load("name, visibility");
        await context.sync();
        return { success: true, action: normalized, sheetName: sheet.name, visibility: sheet.visibility };
      }

      // 到这里说明别名表与分支不同步（防回归）。
      throw new Error(`[Office.js 通道] manage_sheet 内部错误：action "${rawAction}" 已归一为 "${normalized}"，但没有对应实现分支。`);
    });
  }
