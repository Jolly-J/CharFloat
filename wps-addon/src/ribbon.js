  // ---------------------------------------------------------------------------
  // ribbon.js — Ribbon 全局回调 window.OnAction* 与剪贴板工具
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  // ==========================================
  // Ribbon 全局回调顶置与原生状态呈现
  // ==========================================

  window.OnActionBridgeStatus = function () {
    try {
      const app = getApp();
      const host = detectHostComponent();
      const summary = app ? getWorkspaceSummary(app) : null;
      let msg = "【字浮 CharFloat 运行状态】\n\n";
      msg += `插件版本: v${currentVersion} (协议 v2)\n`;
      msg += `通信状态: ${isConnected ? `已连接 (127.0.0.1:${BRIDGE_PORT})` : `未连接 (${lastDisconnectReason})`}\n`;
      msg += `网关配置: 127.0.0.1:${BRIDGE_PORT} (Token: ${BRIDGE_TOKEN ? "已配置" : "缺失"})\n`;
      if (serverVersionNotice) {
        msg += `版本提示: 发现新版本 v${serverVersionNotice.latestVersion} (建议从客户端升级)\n`;
      }
      msg += `当前组件: ${host === "word" ? "WPS 文字 (Word)" : (host === "ppt" ? "WPS 演示 (PowerPoint)" : "WPS 表格 (Excel)")}\n\n`;

      if (host === "word") {
        if (summary && summary.hasOpenDocument) {
          msg += `当前文档: ${summary.documentName}\n`;
          msg += `段落总数: ${summary.paragraphCount} 段\n`;
          msg += `表格总数: ${summary.tableCount} 个\n`;
          msg += `字数统计: ${summary.wordCount} 字\n`;
        } else {
          msg += "当前未检测到打开的 Word 文档\n";
        }
      } else if (host === "ppt") {
        if (summary && summary.hasOpenPresentation) {
          msg += `当前演示文稿: ${summary.presentationName}\n`;
          msg += `幻灯片总数: ${summary.slideCount} 页\n`;
        } else {
          msg += "当前未检测到打开的 PPT 演示文稿\n";
        }
      } else {
        if (summary && summary.hasOpenWorkbook) {
          msg += `当前工作簿: ${summary.workbookName}\n`;
          msg += `活动工作表: ${summary.activeSheetName}\n`;
          if (summary.selection) {
            msg += `鼠标光标选区: ${summary.selection.address} (${summary.selection.rowCount}行 × ${summary.selection.columnCount}列)\n`;
          }
        } else {
          msg += "当前未检测到打开的 Excel 表格文件\n";
        }
      }

      if (!isConnected) {
        msg += "\n排查建议:\n1. 确认字浮 CharFloat 桌面客户端已启动后台服务；\n2. 若长期无法连通，可在客户端点击【安装 / 升级加载项】。";
      }
      showNativeAlert(msg);
    } catch (e) {
      showNativeAlert("获取 Bridge 状态异常: " + (e.message || String(e)));
    }
  };

  window.OnActionForceReconnect = function () {
    try {
      log("用户点击重新连接...");
      if (ws) {
        try { ws.close(); } catch (e) {}
      }
      initWebSocket();
      window.location.reload();
    } catch (e) {
      showNativeAlert("重连异常: " + (e.message || String(e)));
    }
  };

  window.OnRibbonLoaded = function (ribbonUI) {
    try {
      log("WPS 功能区载入，自动激活连接...");
      initWebSocket();
    } catch (e) {}
  };

  window.OnGetImage = function (control) {
    try {
      const id = typeof control === "object" && control ? (control.Id || control.id) : String(control);
      if (id === "btnBrandHero") {
        return "logo.png";
      }
      if (id === "btnAutoFitFormat") {
        return "table-format.png";
      }
    } catch (e) {
      log("[OnGetImage Error]", e.message);
    }
    return "";
  };

  window.OnActionUndoAiAction = async function () {
    try {
      if (detectHostComponent() !== "excel") throw new Error("审计撤销仅覆盖表格单元格的值与公式。");
      const summary = getWorkspaceSummary(getApp());
      if (!summary.workbookName) throw new Error("请先打开目标工作簿。");
      const call = async (name, args) => {
        const response = await fetch("http://127.0.0.1:" + (config.port || 19890) + "/api/v1/tool/call", {
          method: "POST", headers: { "Content-Type": "application/json", "Authorization": "Bearer " + config.token },
          body: JSON.stringify({ name: name, arguments: args, clientName: "WPS Ribbon", sessionId: "wps-ribbon" })
        });
        const result = await response.json();
        if (!response.ok || result.success === false) throw new Error(result.error || "Bridge 请求失败");
        return result.data;
      };
      const records = await call("wps_get_audit_history", { workbookName: summary.workbookName, sheetName: summary.activeSheetName, status: "applied", limit: 20 });
      const record = records.find(item => !item.host || item.host === "wps");
      if (!record) { showNativeAlert("当前工作表没有可撤销的单元格修改记录。"); return; }
      if (!confirm("撤销此修改？\n" + record.description + "\n" + record.workbookName + " / " + record.sheetName + " / " + record.address + "\n有后续修改时会拒绝覆盖。")) return;
      await call("wps_rollback", { auditId: record.id });
      showNativeAlert("已恢复记录中的单元格值与公式。");
    } catch (error) { showNativeAlert("撤销未完成：" + error.message); }
  };

  function copyTextToClipboard(text) {
    try {
      if (typeof navigator !== "undefined" && navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text);
        return true;
      }
    } catch (e) {}
    try {
      if (typeof document !== "undefined") {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.style.position = "fixed";
        textArea.style.opacity = "0";
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand("copy");
        document.body.removeChild(textArea);
        return true;
      }
    } catch (e) {}
    return false;
  }

  window.OnActionOpenDesktopApp = function () {
    try {
      showNativeAlert("【字浮 CharFloat 控制中心】\n\n请在屏幕顶部菜单栏或程序坞中切换至「字浮 CharFloat」桌面管理窗口。\n\n您可以在管理中心中配置 AI 助手（豆包、WorkBuddy、Kimi、Claude 等）、查看单元格修改快照以及进行系统诊断。");
    } catch (e) {
      showNativeAlert("打开控制中心提示: " + (e.message || String(e)));
    }
  };

  window.OnActionShowAuditHistory = async function () {
    try {
      if (detectHostComponent() !== "excel") {
        showNativeAlert("修改历史清单功能目前主要面向表格（Excel/WPS表格）单元格变更。");
        return;
      }
      const summary = getWorkspaceSummary(getApp());
      if (!summary.workbookName) throw new Error("请先打开目标工作簿。");
      const response = await fetch("http://127.0.0.1:" + (config.port || 19890) + "/api/v1/tool/call", {
        method: "POST", headers: { "Content-Type": "application/json", "Authorization": "Bearer " + config.token },
        body: JSON.stringify({ name: "wps_get_audit_history", arguments: { workbookName: summary.workbookName, limit: 10 }, clientName: "WPS Ribbon", sessionId: "wps-ribbon" })
      });
      const result = await response.json();
      const records = result?.data || [];
      if (!records.length) {
        showNativeAlert("【安全时光机 · 历史清单】\n\n当前工作簿尚未产生 AI 修改记录。\n当 AI 智能体执行表格写入时，系统将自动记录前后快照。");
        return;
      }
      let msg = `【安全时光机 · 最近 ${records.length} 条修改快照】\n\n`;
      records.forEach((r, idx) => {
        const time = r.timestamp ? new Date(r.timestamp).toLocaleTimeString() : "刚刚";
        const statusText = r.status === "rolled_back" ? "[已撤销]" : "[有效]";
        msg += `${idx + 1}. [${time}] ${statusText} ${r.sheetName || "Sheet1"}!${r.address || "区域"}\n   说明: ${r.description || "单元格写入"}\n`;
      });
      msg += "\n如需撤销最近修改，可直接点击功能区【撤销 AI 修改】。";
      showNativeAlert(msg);
    } catch (e) {
      showNativeAlert("获取修改历史异常: " + (e.message || String(e)));
    }
  };

  window.OnActionToggleSheetLock = function () {
    try {
      const app = getApp();
      if (!app) throw new Error("未检测到活动办公应用");
      const host = detectHostComponent();
      if (host === "excel") {
        const sheet = app.ActiveSheet;
        if (!sheet) throw new Error("未检测到活动工作表");
        if (sheet.ProtectContents) {
          sheet.Unprotect();
          showNativeAlert("【工作表安全锁已解除】\n\n当前工作表已允许编辑与 AI 智能体写入。");
        } else {
          sheet.Protect();
          showNativeAlert("【工作表已锁定保护】\n\n已开启工作表防误改保护！AI 智能体尝试写入时将受到安全防护拒绝。");
        }
      } else {
        showNativeAlert("锁定保护功能目前优先适配表格工作表。");
      }
    } catch (e) {
      showNativeAlert("锁定操作异常: " + (e.message || String(e)));
    }
  };

  window.OnActionAutoFitFormat = function () {
    try {
      const app = getApp();
      if (!app) throw new Error("未检测到活动办公应用");
      const host = detectHostComponent();
      if (host === "excel") {
        const sheet = app.ActiveSheet;
        if (!sheet) throw new Error("未检测到活动工作表");
        if (sheet.UsedRange && sheet.UsedRange.Columns) {
          sheet.UsedRange.Columns.AutoFit();
          showNativeAlert("【自适应排版完成】\n\n已自动根据内容长度优化并自适应整张工作表的所有列宽！");
        } else {
          showNativeAlert("当前工作表没有可用数据区域。");
        }
      } else {
        showNativeAlert("自适应排版功能优先适配表格（Excel）排版。");
      }
    } catch (e) {
      showNativeAlert("自适应排版失败: " + (e.message || String(e)));
    }
  };

  window.OnActionCopyAsMarkdown = function () {
    try {
      const app = getApp();
      if (!app) throw new Error("未检测到活动办公应用");
      const host = detectHostComponent();
      if (host === "excel") {
        const sel = app.Selection;
        if (!sel) throw new Error("请先用鼠标框选需要复制的单元格区域。");
        const rowCount = sel.Rows.Count;
        const colCount = sel.Columns.Count;
        if (rowCount === 0 || colCount === 0) throw new Error("选区为空");
        
        let md = "";
        for (let r = 1; r <= rowCount; r++) {
          let rowCells = [];
          for (let c = 1; c <= colCount; c++) {
            const cell = sel.Cells.Item(r, c);
            const val = cell.Text || cell.Value2 || "";
            rowCells.push(String(val).replace(/\|/g, "\\|").replace(/\n/g, " "));
          }
          md += "| " + rowCells.join(" | ") + " |\n";
          if (r === 1) {
            md += "| " + rowCells.map(() => "---").join(" | ") + " |\n";
          }
        }
        copyTextToClipboard(md);
        showNativeAlert(`【选区已复制为 Markdown】\n\n已成功将 ${rowCount} 行 × ${colCount} 列数据格式化为 Markdown 表格并写入剪贴板！\n可直接在任意 AI 助手（豆包、Kimi、ChatGPT 等）对话框中按 Ctrl+V / Cmd+V 粘贴提问。`);
      } else {
        showNativeAlert("复制为 Markdown 功能目前主要面向表格数据选区。");
      }
    } catch (e) {
      showNativeAlert("复制选区失败: " + (e.message || String(e)));
    }
  };

  window.OnActionClearEmptyRows = function () {
    try {
      showNativeAlert("【清除冗余空白】\n\n建议直接选中需要清理的行或列，按键盘 Delete 清除，或在 AI 助手中输入：“帮我检查并清理本表中的空行空列”。");
    } catch (e) {}
  };

  window.OnActionPromptFinance = function () {
    const prompt = "请分析当前打开的财务/业务数据表格，从核心营收、同比环比、毛利率以及异常波动点进行深度专业洞察，并指出潜在的经营风险与优化建议。";
    copyTextToClipboard(prompt);
    showNativeAlert("【已复制财务分析提示词】\n\n「" + prompt + "」\n\n提示词已复制到剪贴板，可直接粘贴发送给您的 AI 助手！");
  };

  window.OnActionPromptFormula = function () {
    const prompt = "请帮我检查当前表格中计算公式的逻辑与引用范围，指出潜在的 #N/A 或循环引用错误，并给出最简洁优雅的修复公式建议。";
    copyTextToClipboard(prompt);
    showNativeAlert("【已复制公式排错提示词】\n\n「" + prompt + "」\n\n提示词已复制到剪贴板，可直接粘贴发送给您的 AI 助手！");
  };

  window.OnActionPromptCleaning = function () {
    const prompt = "请帮我清洗当前表格数据：规范日期与手机号格式，剔除前后不可见空格，识别并标记重复项与缺失值。";
    copyTextToClipboard(prompt);
    showNativeAlert("【已复制数据清洗提示词】\n\n「" + prompt + "」\n\n提示词已复制到剪贴板，可直接粘贴发送给您的 AI 助手！");
  };

  window.OnActionPromptSummary = function () {
    const prompt = "请快速提炼当前文档/表格的核心关键数据与结论，按照高管汇报要点梳理出 3 条核心要点与下一步行动建议。";
    copyTextToClipboard(prompt);
    showNativeAlert("【已复制提炼摘要提示词】\n\n「" + prompt + "」\n\n提示词已复制到剪贴板，可直接粘贴发送给您的 AI 助手！");
  };

  window.OnActionShowGuide = function () {
    showNativeAlert("【字浮 CharFloat 快速使用指南】\n\n1. 确保字浮 CharFloat 桌面客户端处于「正常运行」状态；\n2. 在桌面端「AI 助手授权中心」一键绑定您的常用客户端（豆包 / WorkBuddy / Kimi / Claude 等）；\n3. 在 AI 客户端中直接对话即可实时读取、分析并修改当前打开的表格与文档！\n4. 任何时候均可点击上方【撤销 AI 修改】秒级恢复数据。");
  };
