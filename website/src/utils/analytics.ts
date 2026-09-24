declare global {
  interface Window {
    LA?: {
      track: (eventName: string, data?: Record<string, string | number | boolean>) => void;
      [key: string]: any;
    };
  }
}

/**
 * 触发 51.la 自定义事件上报
 * @param eventName 事件名称（例如 'download_client'、'copy_cli'）
 * @param params 附加参数属性（例如 { os: 'mac', version: '2.2.0' }）
 */
export function trackEvent(eventName: string, params?: Record<string, string | number | boolean>) {
  try {
    if (typeof window !== 'undefined' && window.LA && typeof window.LA.track === 'function') {
      window.LA.track(eventName, params);
    }
  } catch (err) {
    // 埋点异常不阻断用户正常交互
    console.debug('[Analytics] track error:', err);
  }
}
