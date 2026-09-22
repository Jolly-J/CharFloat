/**
 * 目标文档锁存储（P3.1：从 gateway.ts 原样搬迁，语义未改）。
 *
 * 按 `sessionId:host` 分桶，保证不同会话/宿主的目标锁互不串扰。
 * 门面 `../gateway.ts` 继续以同名导出转发本类，既有调用方与测试无需改动。
 */
import { currentHost, currentSession } from "../context.js";

export class TargetLockStore {
  private static sessions = new Map<string, { word?: string; excel?: string; ppt?: string }>();
  private static get locks() {
    const key = currentSession() + ':' + currentHost();
    if (!this.sessions.has(key)) this.sessions.set(key, {});
    return this.sessions.get(key)!;
  }
  private static set locks(value: { word?: string; excel?: string; ppt?: string }) {
    this.sessions.set(currentSession() + ':' + currentHost(), value);
  }

  public static lock(component: "word" | "excel" | "ppt", targetName: string) {
    if (!targetName) {
      delete this.locks[component];
    } else {
      this.locks[component] = targetName.trim();
    }
  }

  public static unlock(component?: "word" | "excel" | "ppt") {
    if (component) {
      delete this.locks[component];
    } else {
      this.locks = {};
    }
  }

  public static getLocks() {
    return { ...this.locks };
  }

  public static resolve(component: "word" | "excel" | "ppt", provided?: string): string | undefined {
    if (provided && typeof provided === "string" && provided.trim()) {
      return provided.trim();
    }
    return this.locks[component];
  }
}
