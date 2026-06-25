const assert = require('assert/strict');
const { buildRetryQueueState_ } = require('../retry_queue_state');

function retryQueue(args) {
  return buildRetryQueueState_(
    args.snapshot,
    args.position,
    args.failed,
    args.stopMode
  ).retryQueue;
}

assert.deepEqual(
  retryQueue({ snapshot: [5, 8, 12, 20], position: 2, failed: [], stopMode: 'paused' }),
  [12, 20],
  'unprocessed retry indexes remain after stopping before 12'
);

assert.deepEqual(
  retryQueue({ snapshot: [5, 8, 12, 20], position: 2, failed: [5, 8], stopMode: 'paused' }),
  [5, 8, 12, 20],
  'failed and unprocessed indexes are saved when stopping before 12'
);

assert.deepEqual(
  retryQueue({ snapshot: [5, 8, 12, 20], position: 2, failed: [8], stopMode: 'paused' }),
  [8, 12, 20],
  'successfully retried indexes are removed while failed indexes remain'
);

assert.deepEqual(
  retryQueue({ snapshot: [5, 8, 12, 20], position: 4, failed: [8, 20], stopMode: 'round_completed' }),
  [8, 20],
  'completed retry round saves only indexes that failed again'
);

assert.deepEqual(
  buildRetryQueueState_([5, 8, 12, 20], 4, [8, 20], 'round_limit_reached'),
  { retryQueue: [8, 20], status: 'completed_with_errors' },
  'exhausted retry rounds complete with errors'
);

assert.deepEqual(
  retryQueue({ snapshot: [5, 8, 12, 20], position: 1, failed: [5], stopMode: 'paused' }),
  [5, 8, 12, 20],
  'no unprocessed index is lost when retry is paused'
);

console.log('rating_jobs_retry.test.js: all assertions passed');
