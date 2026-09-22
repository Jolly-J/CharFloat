  // ---------------------------------------------------------------------------
  // dispatch.js — RPC 报文解析与方法分发（handleIncomingMessage）
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  async function handleIncomingMessage(raw) {
    let packet;
    try {
      packet = JSON.parse(raw);
    } catch (e) {
      log("无法解析报文: " + raw);
      return;
    }

    if (packet && packet.type === "version_notice") {
      serverVersionNotice = packet;
      log(`收到服务端版本协商提醒: ${packet.message} (最新可用: v${packet.latestVersion})`);
      const updateTipEl = document.getElementById("updateTip");
      if (updateTipEl) {
        updateTipEl.style.display = "block";
        updateTipEl.innerText = `💡 提示: 服务端检测到加载项有新版本 v${packet.latestVersion}，可从客户端升级`;
      }
      return;
    }

    const { id, method, params } = packet;
    log(`收到 RPC 调用 [${method}], id: ${id}`);

    const app = getApp();
    if (!app) {
      sendRpcResponse(id, null, "WPS 宿主对象未就绪");
      return;
    }

    try {
      let result = null;
      switch (method) {
        case "ping":
          result = { pong: true, time: Date.now() };
          break;
        case "get_workspace_summary":
          result = getWorkspaceSummary(app, params?.workbookName);
          break;
        case "get_sheet_outline":
          result = getSheetOutline(app, params);
          break;
        case "get_style_token":
          result = getStyleToken(app, params);
          break;
        case "create_sheet":
          result = createWorksheet(app, params);
          break;
        case "delete_sheet":
          result = deleteWorksheet(app, params);
          break;
        case "clear_range":
          result = clearRange(app, params);
          break;
        case "read_range":
          result = readRangeData(app, params);
          break;
        case "get_range_styles":
          result = getRangeStyles(app, params);
          break;
        case "search_cells":
          result = searchCells(app, params);
          break;
        case "patch_cells":
          result = patchCells(app, params);
          break;
        case "format_cells":
          result = formatCells(app, params);
          break;
        case "add_conditional_formatting":
          result = addConditionalFormatting(app, params);
          break;
        case "freeze_panes":
          result = freezePanes(app, params);
          break;
        case "modify_rows_columns":
          result = modifyRowsColumns(app, params);
          break;
        case "auto_fit_columns":
          result = autoFitColumns(app, params);
          break;
        case "rollback_cells":
          result = rollbackCells(app, params);
          break;
        case "insert_dimension":
          result = insertDimension(app, params);
          break;
        case "capture_sheet_preview":
          result = captureSheetPreview(app, params);
          break;
        case "add_chart":
          result = addChart(app, params);
          break;
        case "get_charts":
          result = getCharts(app, params);
          break;
        case "delete_chart":
          result = deleteChart(app, params);
          break;
        case "create_pivot_table":
          result = createPivotTable(app, params);
          break;
        case "set_filter_and_sort":
          result = setFilterAndSort(app, params);
          break;
        case "set_data_validation":
          result = setDataValidation(app, params);
          break;
        case "manage_sheet":
          result = manageSheet(app, params);
          break;
        case "manage_rows_and_columns":
          result = modifyRowsColumns(app, params);
          break;
        case "manage_cell_comments":
          result = manageCellComments(app, params);
          break;
        case "find_and_replace":
          result = findAndReplace(app, params);
          break;
        case "duplicate_sheet":
          result = duplicateSheet(app, params);
          break;
        case "save_workbook":
          result = saveWorkbook(app, params);
          break;

        // Word (文字) RPC 分发
        case "word_create_document":
          result = wordCreateDocument(app, params);
          break;
        case "word_save_document":
          result = wordSaveDocument(app, params);
          break;
        case "word_close_document":
          result = wordCloseDocument(app, params);
          break;
        case "word_manage_content":
          result = wordManageContent(app, params);
          break;
        case "word_read_document":
          result = wordReadDocument(app, params);
          break;
        case "word_write_content":
          result = wordWriteContent(app, params);
          break;
        case "word_format_document":
          result = wordFormatDocument(app, params);
          break;
        case "word_insert_table_of_contents":
          result = wordInsertTableOfContents(app, params);
          break;
        case "word_manage_table":
          result = wordManageTable(app, params);
          break;
        case "word_review_and_comments":
          result = wordReviewAndComments(app, params);
          break;
        case "word_page_layout_and_watermark":
          result = wordPageLayoutAndWatermark(app, params);
          break;
        case "word_find_and_replace":
          result = wordFindAndReplace(app, params);
          break;
        case "word_capture_preview":
          result = wordCapturePreview(app, params);
          break;

        // PowerPoint (演示) RPC 分发
        case "ppt_read_presentation":
          result = pptReadPresentation(app, params);
          break;
        case "ppt_get_slide_shapes":
          result = pptGetSlideShapes(app, params);
          break;
        case "ppt_generate_deck":
          result = pptGenerateDeck(app, params);
          break;
        case "ppt_manage_slides":
          result = pptManageSlides(app, params);
          break;
        case "ppt_manage_table":
          result = pptManageTable(app, params);
          break;
        case "ppt_add_business_cards":
          result = pptAddBusinessCards(app, params);
          break;
        case "ppt_insert_native_chart":
          result = pptInsertNativeChart(app, params);
          break;
        case "ppt_manage_shapes_and_media":
          result = pptManageShapesAndMedia(app, params);
          break;
        case "ppt_capture_slide_preview":
          result = pptCaptureSlidePreview(app, params);
          break;

        // 文档强隔离锁定 RPC 分发
        case "lock_target_document": {
          const { component, targetName } = params || {};
          if (component && lockedTargets.hasOwnProperty(component)) {
            lockedTargets[component] = targetName ? String(targetName).trim() : null;
          }
          result = {
            success: true,
            component,
            lockedTarget: lockedTargets[component],
            allLocks: lockedTargets,
            message: `WPS 内部已锁定 ${component.toUpperCase()} 文档为 [${lockedTargets[component]}]`
          };
          break;
        }
        case "unlock_target_document": {
          const { component } = params || {};
          if (component && lockedTargets.hasOwnProperty(component)) {
            lockedTargets[component] = null;
          } else {
            lockedTargets.word = null;
            lockedTargets.excel = null;
            lockedTargets.ppt = null;
          }
          result = { success: true, allLocks: lockedTargets, message: "WPS 内部已释放锁定" };
          break;
        }
        case "get_locked_status": {
          result = {
            locks: lockedTargets,
            openPptPresentations: (() => {
              try {
                const pptApp = getPptApp();
                const list = [];
                if (pptApp && pptApp.Presentations) {
                  for (let i = 1; i <= pptApp.Presentations.Count; i++) {
                    list.push(pptApp.Presentations.Item(i).Name);
                  }
                }
                return list;
              } catch (e) { return []; }
            })(),
            openWordDocuments: (() => {
              try {
                const wordApp = getWpsApp() || getWordApp();
                const list = [];
                if (wordApp && wordApp.Documents) {
                  for (let i = 1; i <= wordApp.Documents.Count; i++) {
                    list.push(wordApp.Documents.Item(i).Name);
                  }
                }
                return list;
              } catch (e) { return []; }
            })(),
            openExcelWorkbooks: (() => {
              try {
                const etApp = getEtApp();
                const list = [];
                if (etApp && etApp.Workbooks) {
                  for (let i = 1; i <= etApp.Workbooks.Count; i++) {
                    list.push(etApp.Workbooks.Item(i).Name);
                  }
                }
                return list;
              } catch (e) { return []; }
            })()
          };
          break;
        }

        case "eval":
        case "eval_code":
        case "execute_script": {
          const startTime = Date.now();
          const logs = [];
          const customConsole = {
            log: (...a) => logs.push(a.map(x => (typeof x === "object" ? JSON.stringify(x) : String(x))).join(" ")),
            warn: (...a) => logs.push("[WARN] " + a.map(x => (typeof x === "object" ? JSON.stringify(x) : String(x))).join(" ")),
            error: (...a) => logs.push("[ERROR] " + a.map(x => (typeof x === "object" ? JSON.stringify(x) : String(x))).join(" "))
          };

          let targetDoc = null;
          let targetWb = null;
          let targetPres = null;
          try { targetDoc = getWordDocument(app, params?.documentName); } catch (e) {}
          try { targetWb = getWorkbook(app, params?.workbookName); } catch (e) {}
          try { targetPres = getPptPresentation(app, params?.presentationName); } catch (e) {}

          const scriptCode = params?.code || params?.script || "";
          if (!scriptCode || typeof scriptCode !== "string") {
            throw new Error("缺少要执行的脚本代码: code");
          }

          // 封装具有返回值的异步/同步图灵执行沙箱
          const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
          const runner = new AsyncFunction(
            "app",
            "doc",
            "wb",
            "pres",
            "wps",
            "params",
            "console",
            `"use strict";\n${scriptCode}`
          );

          let evalResult;
          try {
            evalResult = await runner(
              app,
              targetDoc,
              targetWb,
              targetPres,
              typeof wps !== "undefined" ? wps : undefined,
              params?.params || {},
              customConsole
            );
          } catch (execErr) {
            throw new Error(`原生脚本执行异常: ${execErr.message}\n堆栈: ${execErr.stack || "无"}\n控制台输出: ${logs.join("\n")}`);
          }

          // 安全序列化返回值（防止 Office 原生对象循环引用）
          function safeSerialize(val, depth = 0) {
            if (val === null || val === undefined) return val;
            if (typeof val !== "object") return val;
            if (depth > 2) return String(val);
            if (Array.isArray(val)) return val.map(item => safeSerialize(item, depth + 1));
            const out = {};
            for (const k in val) {
              try {
                const v = val[k];
                if (typeof v === "function") continue;
                out[k] = safeSerialize(v, depth + 1);
              } catch (e) {}
            }
            return Object.keys(out).length > 0 ? out : String(val);
          }

          result = {
            success: true,
            executionTimeMs: Date.now() - startTime,
            returnValue: safeSerialize(evalResult),
            logs,
            message: `WPS 原生图灵脚本执行完毕（耗时 ${Date.now() - startTime}ms）`
          };
          break;
        }

        case "inspect_api": {
          const targetExpr = params?.expression || params?.path || "app";
          let targetDoc = null;
          let targetWb = null;
          let targetPres = null;
          try { targetDoc = getWordDocument(app, params?.documentName); } catch (e) {}
          try { targetWb = getWorkbook(app, params?.workbookName); } catch (e) {}
          try { targetPres = getPptPresentation(app, params?.presentationName); } catch (e) {}

          const resolver = new Function(
            "app", "doc", "wb", "pres", "wps",
            `try { return ${targetExpr}; } catch (e) { return { __inspect_error: e.message }; }`
          );

          const obj = resolver(
            app,
            targetDoc,
            targetWb,
            targetPres,
            typeof wps !== "undefined" ? wps : undefined
          );

          if (!obj || obj.__inspect_error) {
            result = {
              success: false,
              expression: targetExpr,
              error: obj ? obj.__inspect_error : "目标对象为空 (null/undefined)"
            };
            break;
          }

          const propList = [];
          const methodList = [];
          const memberMap = new Set();

          // 遍历对象属性与原型链
          let curr = obj;
          let depth = 0;
          while (curr && depth < 3) {
            try {
              const names = Object.getOwnPropertyNames(curr);
              for (const name of names) {
                if (memberMap.has(name) || name.startsWith("__")) continue;
                memberMap.add(name);
                try {
                  const val = obj[name];
                  const type = typeof val;
                  if (type === "function") {
                    methodList.push(name);
                  } else {
                    propList.push({
                      name,
                      type,
                      valueSample: type === "object" ? (val ? "[Object]" : "null") : String(val).slice(0, 80)
                    });
                  }
                } catch (e) {
                  propList.push({ name, type: "unknown", error: "无法读取" });
                }
              }
            } catch (e) {}
            try { curr = Object.getPrototypeOf(curr); } catch (e) { break; }
            depth++;
          }

          result = {
            success: true,
            expression: targetExpr,
            typeName: typeof obj,
            constructorName: obj.constructor ? obj.constructor.name : "Object",
            propertyCount: propList.length,
            methodCount: methodList.length,
            properties: propList.slice(0, 100),
            methods: methodList.sort(),
            message: `成功完成对 [${targetExpr}] 的运行时 API 反射探测`
          };
          break;
        }

        case "reload":
          window.location.reload();
          result = { reloading: true };
          break;
        default:
          throw new Error(`未知的 RPC 方法: ${method}`);
      }
      sendRpcResponse(id, result, null);
    } catch (err) {
      log(`执行方法 [${method}] 失败: ${err.message}`);
      sendRpcResponse(id, null, err.message);
    }
  }
