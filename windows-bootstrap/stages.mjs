import { minimumStage } from './core.mjs';

// 完成状态只有操作真正成功且三个阶段的最低展示时间均满足后才可启动。
export async function runStages(operation, send, clock = {}) {
  const now = clock.now || Date.now;
  await send({ stage: 'hello' }); await minimumStage(now(), clock);
  await send({ stage: 'work', phase: 'verify', bytes: 0 }); const start = now();
  let result;
  try { result = await operation(); } finally { await minimumStage(start, clock); }
  await send({ stage: 'done', enabled: false }); await minimumStage(now(), clock);
  await send({ stage: 'done', enabled: true }); return result;
}
