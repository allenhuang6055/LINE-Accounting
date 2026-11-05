'use strict';
// Process-local lifecycle only. It is not a durable inbox or a multi-writer lock.
function createLifecycle() {
  let closing = false;
  const active = new Set();
  function track(work) {
    if (closing) throw Error('WRITER_DRAINING');
    const task = Promise.resolve().then(work);
    active.add(task);
    task.then(() => active.delete(task), () => active.delete(task));
    return task;
  }
  async function idle() {
    while (active.size) await Promise.allSettled([...active]);
  }
  return { track, idle, stop() { closing = true; },
    get closing() { return closing; }, get active() { return active.size; } };
}
module.exports = { createLifecycle };
