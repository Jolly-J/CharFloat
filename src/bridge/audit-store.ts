import fs from "fs";
import path from "path";
import { runtimeHome, atomicWrite } from "./runtime.js";
import { v4 as uuidv4 } from "uuid";
import { AuditRecord, PatchResult, RangeSnapshot } from "./types.js";

type AuditListener = (record: AuditRecord) => void;

export interface AuditQuery {
  limit?: number;
  offset?: number;
  workbookName?: string;
  sheetName?: string;
  clientName?: string;
  actionType?: AuditRecord["actionType"];
  status?: AuditRecord["status"];
  fromTimestamp?: number;
  toTimestamp?: number;
}

export class AuditStore {
  private records: AuditRecord[] = [];
  private storageDir: string;
  private storageFile: string;
  private listeners: Set<AuditListener> = new Set();

  constructor(customDir?: string) {
    this.storageDir = customDir || runtimeHome();
    this.storageFile = path.join(this.storageDir, "audit_history.json");
    this.loadFromDisk();
  }

  private loadFromDisk() {
    try {
      if (fs.existsSync(this.storageFile)) {
        const raw = fs.readFileSync(this.storageFile, "utf-8");
        this.records = JSON.parse(raw);
      }
    } catch (e) {
      console.error("[AuditStore] 加载历史留痕失败:", e);
      this.records = [];
    }
  }

  private saveToDisk() {
    try {
      // records 使用最新记录在前的顺序，只保留最新 500 条。
      this.records = this.records.slice(0, 500);
      atomicWrite(this.storageFile, JSON.stringify(this.records, null, 2));
    } catch (e) {
      throw new Error("文档已修改，但审计持久化失败；请勿重试写入。" + String(e));
    }
  }

  /**
   * 记录一次 Agent 的修改动作
   */
  public addRecord(params: {
    clientName?: string;
    host?: "wps" | "microsoft";
    actionType: AuditRecord["actionType"];
    description: string;
    workbookName: string;
    sheetName: string;
    address: string;
    patchResult: PatchResult;
  }): AuditRecord {
    const record: AuditRecord = {
      id: uuidv4(),
      timestamp: Date.now(),
      clientName: params.clientName || "AI Agent",
      host: params.host || "wps",
      actionType: params.actionType,
      description: params.description,
      workbookName: params.workbookName,
      sheetName: params.sheetName,
      address: params.address,
      modifiedCount: params.patchResult.modifiedCount,
      diff: params.patchResult.diff,
      beforeSnapshot: params.patchResult.beforeSnapshot,
      afterSnapshot: params.patchResult.afterSnapshot,
      status: "applied"
    };

    this.records.unshift(record); // 最新记录排在前面
    this.saveToDisk();

    // 通知所有监听器 (例如 GUI 界面)
    for (const listener of this.listeners) {
      try {
        listener(record);
      } catch (err) {
        console.error("[AuditStore] 通知监听器失败:", err);
      }
    }

    return record;
  }

  /**
   * 获取留痕记录列表
   */
  public getRecords(query: number | AuditQuery = 50): AuditRecord[] {
    const options: AuditQuery = typeof query === "number" ? { limit: query } : query;
    const limit = Math.max(1, Math.min(100, Math.trunc(options.limit ?? 50)));
    const offset = Math.max(0, Math.trunc(options.offset ?? 0));

    return this.records
      .filter((record) => {
        if (options.workbookName && record.workbookName !== options.workbookName) return false;
        if (options.sheetName && record.sheetName !== options.sheetName) return false;
        if (options.clientName && record.clientName !== options.clientName) return false;
        if (options.actionType && record.actionType !== options.actionType) return false;
        if (options.status && record.status !== options.status) return false;
        if (options.fromTimestamp !== undefined && record.timestamp < options.fromTimestamp) return false;
        if (options.toTimestamp !== undefined && record.timestamp > options.toTimestamp) return false;
        return true;
      })
      .slice(offset, offset + limit);
  }

  public getRecordsWithTotal(query: number | AuditQuery = 50): { records: AuditRecord[]; total: number } {
    const options: AuditQuery = typeof query === "number" ? { limit: query } : query;
    const limit = Math.max(1, Math.min(100, Math.trunc(options.limit ?? 50)));
    const offset = Math.max(0, Math.trunc(options.offset ?? 0));

    const filtered = this.records.filter((record) => {
      if (options.workbookName && record.workbookName !== options.workbookName) return false;
      if (options.sheetName && record.sheetName !== options.sheetName) return false;
      if (options.clientName && record.clientName !== options.clientName) return false;
      if (options.actionType && record.actionType !== options.actionType) return false;
      if (options.status && record.status !== options.status) return false;
      if (options.fromTimestamp !== undefined && record.timestamp < options.fromTimestamp) return false;
      if (options.toTimestamp !== undefined && record.timestamp > options.toTimestamp) return false;
      return true;
    });

    return {
      records: filtered.slice(offset, offset + limit),
      total: filtered.length
    };
  }

  public clearAllRecords(): void {
    this.records = [];
    this.saveToDisk();
  }

  /**
   * 按 ID 查找某条留痕
   */
  public getRecordById(id: string): AuditRecord | undefined {
    return this.records.find((r) => r.id === id);
  }

  /**
   * 标记该记录已被回滚
   */
  public markRolledBack(id: string): boolean {
    const rec = this.getRecordById(id);
    if (!rec) return false;
    rec.status = "rolled_back";
    this.saveToDisk();
    return true;
  }

  /**
   * 注册修改留痕广播监听
   */
  public subscribe(listener: AuditListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
}

export const auditStore = new AuditStore();
