/**
 * 网关原生工具定义的装配点（P2.3 剩余项：工具描述符逐类迁移）。
 *
 * `gatewayToolDefinitions()` 的输出顺序 = 迁移前 `UniversalGateway.getOpenAiTools()`
 * 的已发布顺序，逐段对应（括号内为已发布序号）：
 *
 *   script（1–2）→ microsoft（3–7）→ lock（8–10）→ excel（11–22）→
 *   audit（23，`wps_rollback`）→ excel（24–38）→ word（39–50）→ ppt（51–59）
 *
 * 两处**不能靠直觉调整**的地方：
 *   1. `wps_rollback` 属审计类，但已发布位置夹在 Excel 区间内部，故这里显式插回原位；
 *   2. Microsoft 区块按 `office_` 前缀聚合，与 `toolClassOf` 的 `lock`/`microsoft` 标签
 *      口径不同（标签拆分不改变顺序）。
 *
 * 顺序即 tools/list 与契约快照：改动后用 `npm run snapshot:tools` 比对，
 * 必须与 docs/acceptance/2.1.0-p0p1/tools-snapshot.p1.json 逐字节一致。
 *
 * 依赖方向：本模块只依赖同目录的定义模块（纯数据），不 import gateway / catalog /
 * contracts；`gateway.ts` 单向 import 本模块，反向依赖会形成回环。
 */
import { createDefinitionContext } from './shared.js';
import type { DefinitionContext, GatewayToolDefinition } from './shared.js';
import { scriptToolDefinitions } from './script.js';
import { microsoftToolDefinitions } from './microsoft.js';
import { lockToolDefinitions } from './lock.js';
import { excelToolDefinitions, excelToolDefinitionsAfterAudit } from './excel.js';
import { auditDefinitionTools } from './audit.js';
import { wordToolDefinitions } from './word.js';
import { pptToolDefinitions } from './ppt.js';

/**
 * 组装全部网关原生工具定义。
 *
 * `activeWbHint` 参与拼 description（`wps_get_workspace_summary`）与 `workbookName` 的描述文本，
 * 经 {@link createDefinitionContext} 透传到各定义模块；不传时使用通用文案。
 */
export function gatewayToolDefinitions(activeWbHint?: string): GatewayToolDefinition[] {
  const ctx = createDefinitionContext(activeWbHint);
  return [
    ...scriptToolDefinitions(),
    ...microsoftToolDefinitions(),
    ...lockToolDefinitions(),
    ...excelToolDefinitions(ctx),
    ...auditDefinitionTools(),
    ...excelToolDefinitionsAfterAudit(ctx),
    ...wordToolDefinitions(),
    ...pptToolDefinitions()
  ];
}

export { createDefinitionContext, type DefinitionContext, type GatewayToolDefinition };
