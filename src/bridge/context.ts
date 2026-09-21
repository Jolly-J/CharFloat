import { AsyncLocalStorage } from 'node:async_hooks';
export type Host = 'wps' | 'microsoft';
export interface RequestContext { sessionId: string; host: Host }
export const requestContext = new AsyncLocalStorage<RequestContext>();
export function currentHost(): Host { return requestContext.getStore()?.host || 'wps'; }
export function currentSession() { return requestContext.getStore()?.sessionId || 'local'; }
