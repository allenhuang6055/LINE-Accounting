'use strict';

// The same queue serializes immediate and background central mutations in one process.
// Deploy only one writer instance; Sheets has no conditional unique-key writes.
function createOutboxRunner({ read, execute, finish, around = (_book,_job,work) => work() }) {
  let tail = Promise.resolve();
  return function drain(book) {
    const task = tail.then(async () => {
      const failures = [];
      for (const job of await read(book)) {
        if (job.done) continue;
        try {
          if (!/^\d{6}月$/.test(job.payload.monthSheet) || job.payload.monthSheet < '202610月') {
            throw new Error('Outbox refuses historical months');
          }
          await around(book, job, async () => {
            const outcome = await execute(job.payload);
            await finish(book, job, outcome);
          });
        } catch (error) {
          failures.push({ id: job.payload.txId, error });
          // Preserve ordering: a pending create must not run after its delete.
          break;
        }
      }
      return failures;
    });
    tail = task.catch(() => {});
    return task;
  };
}
module.exports = { createOutboxRunner };
