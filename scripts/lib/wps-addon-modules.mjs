import fs from "node:fs";
import path from "node:path";

/** 发现全部源码片段；路径统一为 POSIX 相对路径，枚举顺序不作为拼接顺序。 */
export function listWpsModules(dir, prefix = "") {
  const files = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) files.push(...listWpsModules(path.join(dir, entry.name), relative));
    else if (entry.isFile() && entry.name.endsWith(".js")) files.push(relative);
  }
  return files.sort();
}

/** 只读校验：新增、缺失或重复片段都不能静默放行。 */
export function assertWpsModuleManifest(dir, modules) {
  const onDisk = listWpsModules(dir);
  const duplicates = modules.filter((name, index) => modules.indexOf(name) !== index);
  const missing = modules.filter((name) => !onDisk.includes(name));
  const unlisted = onDisk.filter((name) => !modules.includes(name));
  if (!duplicates.length && !missing.length && !unlisted.length) return;
  const lines = ["src 目录与 MODULES 清单不一致，构建已中止："];
  if (duplicates.length) lines.push(`  MODULES 中存在重复登记: ${[...new Set(duplicates)].join(", ")}`);
  if (missing.length) lines.push(`  MODULES 列出但文件不存在: ${[...new Set(missing)].join(", ")}`);
  if (unlisted.length) lines.push(`  src 中存在未登记模块（其代码不会进入部署入口）: ${unlisted.join(", ")}`);
  lines.push("  请同步 scripts/build-wps-addon.mjs 的 MODULES。");
  throw new Error(lines.join("\n"));
}
