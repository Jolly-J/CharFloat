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
        case "configure_print_layout":
          result = configurePrintLayout(app, params);
          break;
        case "export_sheet_pdf":
          result = exportSheetPdf(app, params);
          break;
        case "add_shape":
          result = addShape(app, params);
          break;
        case "list_shapes":
          result = listShapes(app, params);
          break;
        case "update_shape":
          result = updateShape(app, params);
          break;
        case "group_shapes":
          result = groupShapes(app, params);
          break;
        case "ungroup_shapes":
          result = ungroupShapes(app, params);
          break;
        case "set_shape_zorder":
          result = setShapeZorder(app, params);
          break;
        case "format_text_segment":
          result = formatTextSegment(app, params);
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

        case "word_update_fields":
          result = wordUpdateFields(app, params);
          break;
        case "word_manage_content_controls":
          result = wordManageContentControls(app, params);
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

          // 安全序列化返回值（防止 Office 原生对象循环引用）。
          // 原实现 `depth > 2` 时直接 `String(val)` —— 三层以上的对象**属性被静默丢弃**，
          // 调用方拿到 undefined 或 "[object Object]" 却没有任何提示（问题台账 ISS-56）。
          // 现在：放宽到 6 层，超限节点写成**显式占位**并记录路径，随结果返回 truncatedPaths。
          const MAX_SERIALIZE_DEPTH = 6;
          const truncatedPaths = [];
          function safeSerialize(val, depth = 0, path = "$") {
            if (val === null || val === undefined) return val;
            if (typeof val !== "object") return val;
            if (depth > MAX_SERIALIZE_DEPTH) {
              truncatedPaths.push(path);
              return `[第 ${depth} 层超出上限 ${MAX_SERIALIZE_DEPTH}，属性已省略；需要完整数据请自行 JSON.stringify 后返回字符串]`;
            }
            if (Array.isArray(val)) {
              return val.map((item, i) => safeSerialize(item, depth + 1, `${path}[${i}]`));
            }
            const out = {};
            for (const k in val) {
              try {
                const v = val[k];
                if (typeof v === "function") continue;
                out[k] = safeSerialize(v, depth + 1, `${path}.${k}`);
              } catch (e) {
                out[k] = `<读取属性失败: ${e.message}>`;
              }
            }
            return Object.keys(out).length > 0 ? out : String(val);
          }

          const serialized = safeSerialize(evalResult);
          result = {
            success: true,
            executionTimeMs: Date.now() - startTime,
            returnValue: serialized,
            // 被截断就明确说出来，不再静默丢数据
            truncated: truncatedPaths.length > 0,
            truncatedPaths: truncatedPaths.length ? truncatedPaths.slice(0, 10) : undefined,
            logs,
            message: `WPS 原生图灵脚本执行完毕（耗时 ${Date.now() - startTime}ms）` +
              (truncatedPaths.length ? `；返回值有 ${truncatedPaths.length} 处超出深度上限被省略，见 truncatedPaths` : "")
          };
          break;
        }

        case "inspect_api": {
          const targetExpr = params?.expression || params?.path || "app";
          // ISS-89 护栏：反射**默认只列成员名，不对成员求值**。
          // 真机取证：逐成员 `obj[name]` 求值会撞进宿主原生层，实测令 WPS 进程崩溃
          //（3 份崩溃报告，2 份调用栈逐帧一致：kso → etcore → etapi → jsetapi → ksojscore；
          // 崩溃点定位在整表 `Worksheet.Cells` 那一类表达式）。列名字用
          // Object.getOwnPropertyNames 是安全的，求值才是危险动作。
          const evaluateMembers = params?.evaluate === true;
          const maxMembers = Number.isFinite(Number(params?.maxMembers)) ? Number(params.maxMembers) : 150;

          /**
           * 已证实/高度可疑的危险成员：即使用户显式要求求值也跳过。
           * - 一组：**已有崩溃取证**——整表/整列/整行范围与整表集合，
           *   求值会构造覆盖整表的原生对象，实测崩溃。
           * - 另一组：**未经证实但保守跳过**——会触达宿主内部集合的原生 getter，
           *   历来是最容易出问题的一类；如需探测请单独用 wps_execute_script 并自行承担风险。
           */
          const DANGEROUS_CONFIRMED = new Set(['Cells', 'Rows', 'Columns', 'UsedRange', 'EntireRow', 'EntireColumn']);
          const DANGEROUS_SUSPECT = new Set([
            'CurrentRegion', 'Precedents', 'Dependents', 'SpecialCells',
            'Comment', 'CommentThreaded', 'Comments', 'CommentsThreaded',
            'Sort', 'SortFields', 'AutoFilter', 'Filters',
            'FormatConditions', 'Validation', 'Names', 'QueryTables', 'Connections',
            'ChartObjects', 'PivotCaches', 'PivotTables', 'ListObjects', 'Styles', 'CommandBars'
          ]);

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
          const skipped = [];
          const memberMap = new Set();
          let evaluatedCount = 0;

          // 遍历对象属性与原型链
          let curr = obj;
          let depth = 0;
          while (curr && depth < 3) {
            try {
              const names = Object.getOwnPropertyNames(curr);
              for (const name of names) {
                if (memberMap.has(name) || name.startsWith("__")) continue;
                memberMap.add(name);

                // 不求值：连 typeof 都不取（取 typeof 同样会触发一次属性访问）
                if (!evaluateMembers) {
                  methodList.push(name);
                  continue;
                }

                if (DANGEROUS_CONFIRMED.has(name)) {
                  skipped.push({ name, reason: "已证实：求值会构造整表范围对象并令 WPS 进程崩溃（ISS-89）" });
                  continue;
                }
                if (DANGEROUS_SUSPECT.has(name)) {
                  skipped.push({ name, reason: "保守跳过：会触达宿主内部集合的原生 getter；如确需探测请用 wps_execute_script 自行承担风险" });
                  continue;
                }
                if (evaluatedCount >= maxMembers) {
                  skipped.push({ name, reason: `已达 maxMembers(${maxMembers}) 上限，未求值` });
                  continue;
                }
                evaluatedCount++;

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
            evaluatedMembers: evaluateMembers,
            propertyCount: propList.length,
            methodCount: methodList.length,
            properties: propList.slice(0, 100),
            methods: methodList.sort(),
            skippedCount: skipped.length,
            skipped: skipped.slice(0, 40),
            message: evaluateMembers
              ? `成功完成对 [${targetExpr}] 的运行时反射（已求值 ${evaluatedCount} 个成员，跳过 ${skipped.length} 个危险/超限成员）`
              : `已列出 [${targetExpr}] 的成员名（**未求值**）。默认不求值是为了避免触及宿主原生 getter 导致 WPS 崩溃（ISS-89）；需要类型/取值时传 evaluate:true，并接受危险成员会被跳过。`
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
