/**
 * 组装入口（composition root，P2.2）。
 *
 * 唯一知道"谁实现谁"的地方：创建协议层需要的工具服务，并注入到服务实例。
 * 入口点（`cli.ts --serve`、`scripts/inspect.ts`、集成测试）调用 `composeBridgeServer()` 完成装配；
 * 未装配就使用 HTTP/MCP 路由会得到明确错误，而不是静默的空清单。
 *
 * 依赖方向：compose → catalog（执行层）、compose → ws-server（协议层）。
 * 没有任何模块反向依赖本模块，因此不会引入新的回环。
 */
import { bridgeServer, WpsBridgeServer } from './ws-server.js';
import { createLocalToolService } from './tool-registry.js';

/** 装配单例服务，幂等；重复调用不会重复包装。 */
export function composeBridgeServer(server: WpsBridgeServer = bridgeServer): WpsBridgeServer {
  server.setToolService(createLocalToolService());
  return server;
}
