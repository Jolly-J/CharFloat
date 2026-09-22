// ── 模块: src/excel/comment.js — 审阅与批注 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1732-1762 行（原样搬迁，未改写）。
  // 11. 审阅与批注
  async function handleManageComments(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const action = params.action || "add";

      if (action === "add") {
        const targetRange = sheet.getRange(params.address || params.cellAddress || "A1");
        const comment = sheet.comments.add(targetRange, params.content || params.text || "");
        comment.load("id, content");
        await context.sync();
        return { success: true, id: comment.id, content: comment.content };
      } else if (action === "list" || action === "read") {
        const comments = sheet.comments.load("items/id, items/content, items/authorName, items/creationDate");
        await context.sync();
        return {
          success: true,
          comments: comments.items.map(c => ({ id: c.id, content: c.content, author: c.authorName, created: c.creationDate }))
        };
      } else if (action === "delete") {
        if (params.id || params.commentId) {
          const comment = sheet.comments.getItem(params.id || params.commentId);
          comment.delete();
        }
        await context.sync();
        return { success: true, message: "批注已删除" };
      }
      return { success: true };
    });
  }

