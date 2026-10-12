  // comments-links.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // 10.4 单元格批注与审阅备注管理
  function manageCellComments(app, params) {
    const { sheetName, workbookName, address, action = "read", text, author } = params || {};
    if (!address && action !== "clear_all") {
      throw new Error("缺少必要参数: address (如 'C5' 或 'D10:D20')");
    }
    const sheet = getWorksheet(app, sheetName, workbookName);
    const targetRange = address ? sheet.Range(address) : sheet.UsedRange;

    // 单元格地址取值。**必须保留接收者直接调用**：宿主方法是原生实现，
    // 摘出来再调（`const f = cell.Address; f()`）会丢 this，抛
    // "Address called on null or undefined"，被 catch 吞掉后表现为地址为空
    // （真机实测：消息变成"已在单元格  添加批注"，ISS-66 的根因就在这里）。
    const addressOf = (cell) => {
      try {
        return String(cell.Address());
      } catch (e) {
        try {
          // 少数宿主把 Address 暴露成属性而非方法
          return String(cell.Address);
        } catch (e2) {
          return "";
        }
      }
    };

    if (action === "add") {
      if (!text) throw new Error("添加批注时必须提供 text 参数");
      const cell = targetRange.Cells.Item(1, 1);
      // 地址必须在 AddComment **之前**取：真机实测 AddComment 之后该单元格的 Address 取不到，
      // 会让成功消息变成"已在单元格  添加批注"（地址为空）。
      const cellAddressText = addressOf(cell); // 先取地址，AddComment 后取不到
      try {
        if (cell.Comment) cell.Comment.Delete();
      } catch (e) {}
      // 不再把作者拼进正文（ISS-43）：先尝试设置真实 Author，设不上再退回"正文前缀"，
      // 并在返回体里**如实区分**"请求的作者"与"读回的作者"（原来直接回显入参，冒充读回，ISS-61）。
      const comment = cell.AddComment(text);
      comment.Visible = false;
      let authorApplied = false;
      if (author) {
        try {
          comment.Author = author;
          authorApplied = true;
        } catch (e) {
          authorApplied = false;
        }
      }
      let authorOnHost = null;
      try { authorOnHost = comment.Author || null; } catch (e) {}
      if (author && !authorApplied && authorOnHost !== author) {
        try {
          comment.Text(text ? `${author}:\n${text}` : text);
        } catch (e) {}
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: cellAddressText,
        action: "add",
        commentText: text,
        authorRequested: author || null,
        authorOnHost,
        authorApplied: authorOnHost === author,
        message: `已在单元格 ${cellAddressText} 添加批注${author && authorOnHost !== author ? `（宿主未接受自定义作者，已把作者写进正文；当前宿主作者为 ${authorOnHost ?? "未知"}）` : ""}`
      };
    } else if (action === "read") {
      const comments = [];
      const rowCount = targetRange.Rows.Count;
      const colCount = targetRange.Columns.Count;
      for (let r = 1; r <= rowCount; r++) {
        for (let c = 1; c <= colCount; c++) {
          const cell = targetRange.Cells.Item(r, c);
          try {
            if (cell.Comment) {
              comments.push({
                address: addressOf(cell),
                text: cell.Comment.Text(),
                author: cell.Comment.Author || undefined
              });
            }
          } catch (e) {}
        }
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: targetRange.Address,
        comments,
        totalComments: comments.length
      };
    } else if (action === "delete") {
      const cell = targetRange.Cells.Item(1, 1);
      try {
        if (cell.Comment) cell.Comment.Delete();
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: cell.Address,
        action: "delete",
        message: `已成功删除单元格 ${cell.Address} 上的批注`
      };
    } else if (action === "clear_all") {
      try {
        sheet.Cells.ClearComments();
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        action: "clear_all",
        message: `已清空工作表 [${sheet.Name}] 中的所有批注`
      };
    // ── CAP-12 线程化批注（真机探测结论）
    //   宿主支持：`AddCommentThreaded(文本)` 建线程、**`thread.AddReply(文本)` 回复**、
    //            `thread.Replies.Item(i).Text()` 读回复。
    //   ⚠️ 宿主**不支持**：`Replies.Add`（不是函数，回复要走 AddReply）、
    //            **`Resolved` 写不进去**（赋值后读回仍是 false）——按"如实告知"处理，不假装已解决。
    } else if (action === "thread" || action === "reply" || action === "read_threads" || action === "resolve") {
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      // ⚠️ WPS JSA 的 Range/Worksheet **都没有 `Cells()` 方法**（本仓库已踩过三次：
      // 截图范围、图表锚点、这里）。一律用**地址字符串**取单元格。
      const colName = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
      const cellAt = (baseAddr, r, c) => {
        const first = String(baseAddr).split(":")[0].replace(/\$/g, "");
        const mm = first.match(/^([A-Za-z]+)(\d+)$/);
        if (!mm) return sheet.Range(first);
        const col0 = mm[1].toUpperCase().split("").reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
        const row0 = Number(mm[2]);
        return sheet.Range(`${colName(col0 + c - 1)}${row0 + r - 1}`);
      };
      const readThread = (cell) => {
        const th = g(() => cell.CommentThreaded, null);
        if (!th) return null;
        const replies = [];
        try {
          const rp = th.Replies;
          for (let i = 1; i <= Number(rp.Count); i++) {
            const it = rp.Item(i);
            replies.push({ author: g(() => String(it.AuthorName), null), text: g(() => String(it.Text()), "") });
          }
        } catch (e) {}
        return {
          address: addressOf(cell),
          text: (() => { try { const rt = th.Text; return typeof rt === "function" ? String(rt()) : String(rt); } catch (e) { return null; } })(),
          author: g(() => String(th.AuthorName), null),
          resolved: g(() => Boolean(th.Resolved), null),
          replyCount: replies.length,
          replies
        };
      };

      if (action === "read_threads") {
        const threads = [];
        const used = targetRange;
        const base = String(g(() => used.Address(), address));
        const maxR = Math.min(Number(g(() => used.Rows.Count, 1)), 500), maxC = Math.min(Number(g(() => used.Columns.Count, 1)), 50);
        for (let r = 1; r <= maxR; r++) {
          for (let cc = 1; cc <= maxC; cc++) {
            const one = readThread(cellAt(base, r, cc));
            if (one) threads.push(one);
          }
        }
        return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, count: threads.length, threads,
          warnings: [], message: `工作表 [${sheet.Name}] 共 ${threads.length} 条线程批注` };
      }

      if (action === "thread") {
        if (!text) throw new Error("thread 操作必须提供 text（批注内容）");
        const cell = cellAt(String(g(() => targetRange.Address(), address)), 1, 1);
        const addr = addressOf(cell);
        // 已有线程则改为追加回复，避免宿主覆盖掉原有讨论
        const existing = g(() => cell.CommentThreaded, null);
        if (existing) {
          try { existing.AddReply(String(text)); }
          catch (e) { throw new Error(`该单元格已有线程批注，追加回复失败：${e.message}`); }
          const back = readThread(cell);
          return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, action: "reply",
            address: addr, thread: back, warnings: [], message: `单元格 ${addr} 已有线程批注，已追加为回复` };
        }
        const th = cell.AddCommentThreaded(String(text));
        if (author) { try { th.AuthorName = String(author); } catch (e) {} }
        const back = readThread(cell);
        const warnings = [];
        if (!back) warnings.push("创建后读不回线程批注，宿主可能未真正写入");
        // ⚠️ 真机实测（WPS 12.1.28496）：`AddCommentThreaded` **返回对象但不真的存内容**——
        // 读回 Text 是 null、回复文本固定为 "default"、Replies.Count 恒为 1、
        // `AddReply` 返回对象但条数不增、`Resolved` 写不进去。
        // 也就是**API 存在但功能是空壳**。这里必须**如实报失败**，不能因为"没抛异常"就报成功。
        if (back && back.text !== null && String(back.text) !== String(text)) {
          warnings.push(`宿主未保存批注正文：写入 ${JSON.stringify(String(text))}，读回 ${JSON.stringify(back.text)}`);
        }
        if (back && (back.text === null || back.text === undefined)) {
          warnings.push("宿主未保存批注正文（读回为 null）——本机 WPS 的线程批注 API 存在但不真正存储内容，请改用传统批注 action='add'");
        }
        return { success: !!back && warnings.length === 0, workbookName: sheet.Parent.Name, sheetName: sheet.Name, action: "thread",
          address: addr, thread: back, warnings,
          message: warnings.length
            ? `线程批注在单元格 ${addr} **未能真正写入**：${warnings[0]}`
            : `已在单元格 ${addr} 创建线程批注` };
      }

      if (action === "reply") {
        if (!text) throw new Error("reply 操作必须提供 text（回复内容）");
        const cell = cellAt(String(g(() => targetRange.Address(), address)), 1, 1);
        const addr = addressOf(cell);
        const th = g(() => cell.CommentThreaded, null);
        if (!th) throw new Error(`单元格 ${addr} 没有线程批注，无法回复；先用 action='thread' 创建`);
        const before = readThread(cell)?.replyCount ?? null;
        try { th.AddReply(String(text)); }
        catch (e) { throw new Error(`回复失败：${e.message}（真机上回复要走 thread.AddReply）`); }
        const back = readThread(cell);
        const after = back?.replyCount ?? null;
        const warnings = [];
        if (before !== null && after !== null && after <= before) warnings.push(`回复后条数未增加（${before} → ${after}），请人工确认`);
        return { success: warnings.length === 0, workbookName: sheet.Parent.Name, sheetName: sheet.Name, action: "reply",
          address: addr, replyCountBefore: before, thread: back, warnings,
          message: `已在单元格 ${addr} 的线程批注下回复（${before} → ${after}）` };
      }

      // action === "resolve"：宿主不支持写 Resolved，**如实告知**而不是假装成功
      const cell = cellAt(String(g(() => targetRange.Address(), address)), 1, 1);
      const addr = addressOf(cell);
      const th = g(() => cell.CommentThreaded, null);
      if (!th) throw new Error(`单元格 ${addr} 没有线程批注`);
      const want = params && params.resolved !== undefined ? Boolean(params.resolved) : true;
      let actual = null;
      try { th.Resolved = want; } catch (e) {}
      actual = g(() => Boolean(th.Resolved), null);
      const ok = actual === want;
      return {
        success: ok, workbookName: sheet.Parent.Name, sheetName: sheet.Name, action: "resolve",
        address: addr, requested: want, actual, warnings: ok ? [] : ["宿主读回与请求不一致"],
        message: ok
          ? `单元格 ${addr} 的线程批注已${want ? "标记解决" : "重新打开"}`
          : `本机宿主**不支持**修改线程批注的解决状态（请求 ${want}，读回 ${actual}）——请人工在 WPS 里操作`
      };
    }
    throw new Error(`未知的批注操作 action: ${action} (支持 'add' | 'read' | 'delete' | 'clear_all' | 'thread' | 'reply' | 'read_threads' | 'resolve')`);
  }

  /** 超链接增删查。 */
  function manageHyperlink(app, params) {
    const { sheetName, workbookName, action = "list", address, url, displayText, tooltip, emailSubject, targetAddress } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };

    if (action === "list") {
      const items = [];
      const n = g(() => Number(sheet.Hyperlinks.Count), 0);
      for (let i = 1; i <= n; i++) {
        const h = (() => { try { return sheet.Hyperlinks.Item(i); } catch (e) { return null; } })();
        if (!h) continue;
        items.push({
          index: i,
          address: g(() => String(h.Address), null),
          subAddress: g(() => String(h.SubAddress), null),
          text: g(() => String(h.TextToDisplay), null),
          tooltip: g(() => String(h.ScreenTip), null),
          anchor: g(() => String(h.Range.Address()), null),
          type: g(() => Number(h.Type), null)
        });
      }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, count: items.length, hyperlinks: items, warnings: [],
        message: `工作表 [${sheet.Name}] 共 ${items.length} 个超链接` };
    }

    if (action === "delete") {
      if (!address) throw new Error("delete 需要 address（锚点单元格）或省略 address 清空全部");
      let removed = 0;
      for (let i = g(() => Number(sheet.Hyperlinks.Count), 0); i >= 1; i--) {
        const h = (() => { try { return sheet.Hyperlinks.Item(i); } catch (e) { return null; } })();
        if (!h) continue;
        const a = g(() => String(h.Range.Address()), "");
        if (a === sheet.Range(address).Address() || !address) { h.Delete(); removed++; }
      }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, removed,
        remaining: g(() => Number(sheet.Hyperlinks.Count), null), warnings: [], message: `已删除 ${removed} 个超链接` };
    }

    // action = add
    if (!address) throw new Error("add 需要 address（锚点单元格）");
    // L-5：只传 targetAddress（文档内跳转，如 'Sheet2!A1'）也是合法用法，
    // 原先强制要求 url → 单独传 targetAddress 直接被拒，与 schema 暴露的参数不符。
    if (!url && !targetAddress) throw new Error("add 需要 url（外部链接）或 targetAddress（文档内跳转）");
    const anchor = sheet.Range(address);
    const linkAddr = String(url || "");
    const subAddr = String(targetAddress || "");
    const disp = String(displayText || url || targetAddress || "链接");
    // Hyperlinks.Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay)
    sheet.Hyperlinks.Add(anchor, linkAddr, subAddr, String(tooltip || ""), disp);
    // 读回核对
    let ok = false, readText = null;
    try { const h = anchor.Hyperlinks.Item(1); ok = true; readText = String(h.TextToDisplay); } catch (e) {}
    return {
      success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, address, url: linkAddr, targetAddress: subAddr,
      verified: ok, readBackText: readText,
      warnings: ok ? [] : ["写入后未能从锚点读到超链接，请人工确认"],
      message: `已在 [${sheet.Name}] ${address} 添加超链接 → ${url}`
    };
  }

  /** 命名区域增删查。 */
  function manageNamedRange(app, params) {
    const { workbookName, action = "list", name, refersTo, comment } = params || {};
    const wb = getWorkbook(app, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };

    if (action === "list") {
      const items = [];
      const n = g(() => Number(wb.Names.Count), 0);
      for (let i = 1; i <= n; i++) {
        const nm = (() => { try { return wb.Names.Item(i); } catch (e) { return null; } })();
        if (!nm) continue;
        items.push({ index: i, name: g(() => String(nm.Name), null), refersTo: g(() => String(nm.RefersTo), null),
                     visible: g(() => Boolean(nm.Visible), null), comment: g(() => String(nm.Comment), null) });
      }
      return { success: true, workbookName: wb.Name, count: items.length, names: items, warnings: [],
        message: `工作簿 [${wb.Name}] 共 ${items.length} 个命名区域` };
    }

    if (action === "delete") {
      if (!name) throw new Error("delete 需要 name");
      wb.Names.Item(String(name)).Delete();
      // ⚠️ 不能用 `Names.Item(name)` 探测是否还存在：删除后它返回 **null 而不抛错**，
      // 于是"没抛错＝还在"的判断**恒为 true**（真机踩到：删除明明成功，stillExists 却报 true，
      // 与实际自检结果矛盾）。改为**枚举全部名称**再比对，这是可靠的。
      const still = (() => {
        try {
          const n = Number(wb.Names.Count);
          for (let i = 1; i <= n; i++) { if (String(wb.Names.Item(i).Name) === String(name)) return true; }
        } catch (e) {}
        return false;
      })();
      const remaining = (() => { try { const out = []; const n = Number(wb.Names.Count); for (let i = 1; i <= n; i++) out.push(String(wb.Names.Item(i).Name)); return out; } catch (e) { return []; } })();
      return { success: !still, workbookName: wb.Name, deleted: name, stillExists: still, remainingNames: remaining.length, warnings: still ? [`删除后仍能枚举到 ${name}，删除未生效`] : [],
        message: still ? `命名区域 ${name} 删除失败（仍存在）` : `已删除命名区域 ${name}，剩余 ${remaining.length} 个` };
    }

    if (!name || !refersTo) throw new Error("add 需要 name 和 refersTo（如 'Sheet1!$A$1:$B$10'）");
    wb.Names.Add(String(name), String(refersTo), false, String(comment || ""));
    const back = g(() => String(wb.Names.Item(String(name)).RefersTo), null);
    // comment 是否真的写进去了（M-10：原先被静默忽略还不告警）
    const warnings = [];
    if (!back) warnings.push("写入后读不回该名称");
    if (comment) {
      const backComment = g(() => String(wb.Names.Item(String(name)).Comment), null);
      if (backComment !== String(comment)) {
        warnings.push(`comment 未生效：请求 ${JSON.stringify(String(comment))}，宿主读回 ${JSON.stringify(backComment)}`);
      }
    }
    return { success: !!back, workbookName: wb.Name, name, requested: refersTo, readBack: back,
      readBackComment: comment ? g(() => String(wb.Names.Item(String(name)).Comment), null) : undefined,
      verified: !!back, warnings,
      message: back ? `已添加命名区域 ${name} → ${back}` : `命名区域 ${name} 添加失败（读不回）` };
  }
