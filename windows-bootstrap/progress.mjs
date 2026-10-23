// 跨文件合并界面进度；真实计数来自引擎，阶段切换立即发送，终态由主进程直接发送。
export function createProgressSender(send, { interval = 200, now = Date.now, setTimer = setTimeout, clearTimer = clearTimeout } = {}) {
  let latest, timer, lastPhase, lastAt = -Infinity, generation = 0, closed = false;
  function cancel() { generation++; if (timer !== undefined) clearTimer(timer); timer = undefined; }
  function emit() { const value = latest; latest = undefined; if (value) { lastAt = now(); lastPhase = value.phase; send(value); } }
  return {
    push(value) {
      if (closed) return;
      latest = value;
      if (value.phase !== lastPhase || now() - lastAt >= interval) { cancel(); emit(); }
      else if (timer === undefined) {
        const token = ++generation;
        timer = setTimer(() => { if (closed || token !== generation) return; timer = undefined; emit(); }, Math.max(0, interval - (now() - lastAt)));
      }
    },
    dispose() { closed = true; latest = undefined; cancel(); }
  };
}
