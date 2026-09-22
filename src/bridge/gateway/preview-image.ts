/**
 * 预览取图：**带 DLP（数据防泄漏）感知**。
 *
 * 背景：装了 DLP（如亿赛通）的机器上，**凡是 WPS/Office 写出的文件都会被包成加密容器**。
 * 宿主导出的预览图正是这种文件——读回来的字节**非空但不是图片**。
 * 只看"有没有字节"会把加密容器当成图片一路传下去，调用方拿到一堆无法解码的数据。
 *
 * 本模块是**纯工具**：不 import 任何执行器，只经形参拿 `GatewayContext`，
 * 因此不引入处理器之间的耦合，也不影响分层边界。
 */

/** 判断 base64 是不是**真的图片**（按 magic 字节）。 */
export function isRealImageBase64(b64: unknown): boolean {
  if (!b64) return false;
  try {
    const head = Buffer.from(String(b64).slice(0, 64), 'base64');
    if (head.length < 4) return false;
    const a = head[0], b = head[1], c = head[2], d = head[3];
    return (a === 0x89 && b === 0x50 && c === 0x4e && d === 0x47)   // PNG
      || (a === 0xff && b === 0xd8 && c === 0xff)                    // JPEG
      || (a === 0x42 && b === 0x4d)                                  // BMP
      || (a === 0x47 && b === 0x49 && c === 0x46);                   // GIF
  } catch { return false; }
}

export interface PreviewImageResult {
  imageBase64?: string;
  /** 宿主写的文件确实存在，但内容不是图片——几乎可以断定是 DLP 加密容器 */
  encryptedByDlp: boolean;
  /** 两条路都拿不到图时给调用方的可操作说明 */
  notice?: string;
}

/**
 * 取预览图：**先看宿主写的文件，不行再走剪贴板**。
 *
 * 顺序很关键——剪贴板是**内存通道**，不经过 DLP 的文件加密，
 * 在加密环境里往往是**唯一能用的那条路**。所以"文件里的字节不是图片"时
 * **必须继续尝试剪贴板**，而不是把加密容器当结果返回。
 *
 * 两条都不通时返回 `notice`，指引改用**几何回读**核验（形状/图表/区域读数），不要依赖截图。
 */
export async function resolvePreviewImage(
  ctx: { extractClipboardImageBase64: () => Promise<string> },
  filePath: string | undefined,
  readFile: (p: string) => string,
  exists: (p: string) => boolean
): Promise<PreviewImageResult> {
  let encryptedByDlp = false;
  if (filePath && exists(filePath)) {
    try {
      const fromFile = readFile(filePath);
      if (isRealImageBase64(fromFile)) return { imageBase64: fromFile, encryptedByDlp: false };
      encryptedByDlp = true;   // 文件在、字节非空、但不是图片 → 极可能是 DLP 加密容器
    } catch { /* 读不动也继续走剪贴板 */ }
  }
  try {
    const fromClip = await ctx.extractClipboardImageBase64();
    if (isRealImageBase64(fromClip)) return { imageBase64: fromClip, encryptedByDlp };
  } catch { /* 剪贴板也不可用，下面统一给说明 */ }
  return {
    imageBase64: undefined,
    encryptedByDlp,
    notice: encryptedByDlp
      ? '⚠️ 预览图被 DLP 加密，取不到。**视觉检查仍必须做**：① 先用你自己的截图能力截图检查；' +
        '② 没有截图能力就请使用者截图发给你；③ 不允许在没看到画面的情况下声称视觉验收通过。' +
        '（几何回读 list_shapes / get_charts(detail=true) / read_range 只能替代结构核验，**不能替代视觉检查**。）'
      : '⚠️ 预览图取不到（宿主未产出文件、剪贴板也无图）。**视觉检查仍必须做**：' +
        '先用你自己的截图能力，或请使用者截图发给你；不得凭结构数据声称视觉验收通过。'
  };
}
