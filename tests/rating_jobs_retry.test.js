const assert = require('assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { buildRetryQueueState_ } = require('../retry_queue_state');

const repoRoot = path.resolve(__dirname, '..');
const parcingPath = path.join(repoRoot, 'parcing_rating.gs');
const parcing = fs.readFileSync(parcingPath, 'utf8');

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
  'scenario 1: unprocessed retry indexes remain after pausing before 12'
);

assert.deepEqual(
  retryQueue({ snapshot: [5, 8, 12, 20], position: 2, failed: [5, 8], stopMode: 'paused' }),
  [5, 8, 12, 20],
  'scenario 2: failed and unprocessed indexes are saved when pausing before 12'
);

assert.deepEqual(
  retryQueue({ snapshot: [5, 8, 12, 20], position: 2, failed: [8], stopMode: 'paused' }),
  [8, 12, 20],
  'scenario 3: successfully retried indexes are removed while failed indexes remain'
);

assert.deepEqual(
  retryQueue({ snapshot: [5, 8, 12], position: 3, failed: [8], stopMode: 'round_completed' }),
  [8],
  'scenario 4: completed retry round saves only indexes that failed again'
);

assert.deepEqual(
  buildRetryQueueState_([5, 8], 2, [5, 8], 'round_limit_reached'),
  { retryQueue: [5, 8], status: 'completed_with_errors' },
  'scenario 5: exhausted retry rounds complete with errors'
);

assert.match(parcing, /buildRetryQueueState_\(job\.retryQueueSnapshot, job\.retryQueuePosition, job\.retryFailedIndexes, 'paused'\)/, 'production code pauses retry queue through helper');
assert.match(parcing, /buildRetryQueueState_\(job\.retryQueueSnapshot, job\.retryQueueSnapshot\.length, job\.retryFailedIndexes, 'round_completed'\)/, 'production code completes retry queue through helper');
assert.doesNotMatch(parcing, /if\s*\([^)]*isRetryStage[^)]*\)\s*\{\s*job\.failedIndexes\s*=\s*\[\]/, 'retry round must not clear failedIndexes before snapshot is preserved');
assert.match(parcing, /retryQueueSnapshot: parseJsonArray_\(all\[prefix \+ 'retryQueueSnapshot'\]\)/, 'getRatingJob_ reads retryQueueSnapshot');
assert.match(parcing, /retryQueuePosition: Number\(all\[prefix \+ 'retryQueuePosition'\]/, 'getRatingJob_ reads retryQueuePosition');
assert.match(parcing, /retryFailedIndexes: parseJsonArray_\(all\[prefix \+ 'retryFailedIndexes'\]\)/, 'getRatingJob_ reads retryFailedIndexes');
assert.match(parcing, /job\.retryQueueSnapshot = job\.failedIndexes\.slice\(\)/, 'retry round starts from immutable failedIndexes snapshot');

for (const fn of ['processRatingJobBatch_', 'getRatingJob_', 'saveRatingJob_']) {
  const matches = parcing.match(new RegExp(`function\\s+${fn}\\s*\\(`, 'g')) || [];
  assert.equal(matches.length, 1, `${fn} must be defined exactly once`);
}

const configMatch = parcing.match(/const CONFIG = \{([\s\S]*?)\n\};/);
assert.ok(configMatch, 'CONFIG object exists');
const ratingJobsKeys = configMatch[1].match(/\bratingJobs\s*:/g) || [];
assert.equal(ratingJobsKeys.length, 1, 'CONFIG contains exactly one ratingJobs key');
assert.doesNotMatch(configMatch[1], /\bstoragePrefix\s*:/, 'CONFIG.ratingJobs does not keep storagePrefix');
assert.doesNotMatch(configMatch[1], /\bbatchMaxRuntimeMs\s*:/, 'CONFIG.ratingJobs does not keep batchMaxRuntimeMs');


assert.match(parcing, /function assertRatingJobHelpers_\(\)/, 'production code defines helper availability assertion');
assert.match(parcing, /assertRatingJobHelpers_\(\);/, 'production code checks retry_queue_state.js helper availability');
assert.match(parcing, /writeAndSaveReserveMs:\s*10000/, 'write-and-save reserve is configured');
assert.match(parcing, /stateSaveReserveMs:\s*3000/, 'state-save reserve is configured');

assert.match(parcing, /announcementWebhookUrl:\s*'https:\/\/n8n-x3\.tech\.temed\.ru\/webhook\//, 'production webhook URL is used');
assert.doesNotMatch(parcing, /announcementWebhookUrl:[^\n]*\/webhook-test\//, 'webhook-test URL is not used');

assert.match(parcing, /function startRatingJob_\(aggregatorKey\) \{\s*createOrContinueRatingJob_\(aggregatorKey\);\s*resumePendingRatingJobs\(\);\s*\}/, 'single-aggregator starts resume the common queue');
assert.match(parcing, /function resumePendingRatingJobs\(\) \{\s*var selected = selectPendingRatingJob_\(\);\s*if \(!selected\) \{\s*return;\s*\}/, 'empty queue returns without debug log');

const helperStart = parcing.indexOf('function mergeUniqueIndexes_');
const helperEnd = parcing.indexOf('function isParsedResultValid_', helperStart);
const pureHelpersSource = parcing.slice(helperStart, helperEnd);
const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(pureHelpersSource, sandbox);
assert.equal(sandbox.isDeferredFetchResult_({ ok: false, deferred: true, error: 'SAFE_TIME_LIMIT_NEAR' }), true, 'SAFE_TIME_LIMIT_NEAR is deferred');

assert.match(parcing, /buildRatingItemTransition_\(job, sourceIndex, result, isRetryStage\)/, 'production loop calls buildRatingItemTransition_ for doctor results');
assert.doesNotMatch(parcing, /function\s+applyRatingJobDecision_\s*\(/, 'unused applyRatingJobDecision_ helper is removed');

const runStats = sandbox.createRatingRunStats_();
const jobStats = { permanentErrors: 10, temporaryErrors: 20, preservedPrevious: 30 };
for (const result of [
  { status: 'success' },
  { status: 'temporary' },
  { status: 'permanent' },
  { status: 'empty' },
  { status: 'success' }
]) {
  const transition = sandbox.buildRatingItemTransition_(jobStats, 1, result, false);
  sandbox.applyRatingItemStats_(jobStats, runStats, transition.itemStats);
}
assert.deepEqual(JSON.parse(JSON.stringify(runStats)), { processed: 5, success: 2, preservedPrevious: 3, permanentErrors: 1, temporaryErrors: 1 }, 'run stats accumulate the whole batch');
assert.equal(jobStats.temporaryErrors, 21, 'job temporary errors increase by batch delta');
assert.equal(jobStats.permanentErrors, 11, 'job permanent errors increase by batch delta');
assert.equal(jobStats.preservedPrevious, 33, 'job preserved values increase by batch delta');

const deferredTransition = sandbox.buildRatingItemTransition_({ nextSourceIndex: 3, retryQueuePosition: 1 }, 12, { status: 'deferred', reason: 'SAFE_TIME_LIMIT_NEAR' }, false);
assert.equal(deferredTransition.nextStatus, 'pending', 'deferred doctor pauses job');
assert.equal(deferredTransition.shouldAdvance, false, 'deferred doctor does not advance position');
assert.deepEqual(Array.from(deferredTransition.addToFailedIndexes), [], 'deferred doctor is not added to failedIndexes');
assert.deepEqual(JSON.parse(JSON.stringify(deferredTransition.itemStats)), { processed: 0, success: 0, preservedPrevious: 0, permanentErrors: 0, temporaryErrors: 0 }, 'deferred doctor does not change stats');

const tempTransition = sandbox.buildRatingItemTransition_({}, 12, { status: 'temporary' }, false);
assert.equal(tempTransition.shouldAdvance, true, 'temporary result advances after HTTP result is received');
assert.deepEqual(Array.from(tempTransition.addToFailedIndexes), [12], 'temporary result is added to retry queue');
assert.deepEqual(JSON.parse(JSON.stringify(tempTransition.itemStats)), { processed: 1, success: 0, preservedPrevious: 1, permanentErrors: 0, temporaryErrors: 1 }, 'temporary result stats are captured');

const retryTempTransition = sandbox.buildRatingItemTransition_({}, 12, { status: 'temporary' }, true);
assert.deepEqual(Array.from(retryTempTransition.addToRetryFailedIndexes), [12], 'retry temporary result is added to retryFailedIndexes');

const statusSource = parcing.match(/function buildRatingJobStatusStage_[\s\S]*?\nfunction assertRatingJobHelpers_/)[0].replace(/\nfunction assertRatingJobHelpers_$/, '');
vm.runInContext(statusSource, sandbox);
assert.deepEqual(
  JSON.parse(JSON.stringify(sandbox.buildRatingJobStatusStage_({ retryQueueSnapshot: [5, 8, 12, 20], retryQueuePosition: 2, retryFailedIndexes: [5], retryRound: 1 }, 100))),
  { stage: 'повтор ошибок', retryRound: 1, retryPosition: 2, retryTotal: 4, retryRemaining: 2, retryFailed: 1 },
  'retry status displays retry queue progress'
);

const selectSource = parcing.match(/function isStaleJob_[\s\S]*?\nfunction ensureTodayRows_/)[0].replace(/\nfunction ensureTodayRows_$/, '');
function selectWithJobs(jobs) {
  const context = {
    CONFIG: { ratingJobs: { staleRunningAfterMs: 900000 } },
    Date,
    getRatingJob_: key => jobs[key] || {},
    saveRatingJob_: (key, job) => { jobs[key] = job; }
  };
  vm.createContext(context);
  vm.runInContext(selectSource, context);
  return context.selectPendingRatingJob_();
}
assert.equal(selectWithJobs({ pd: { status: 'pending', createdAt: '2026-06-25T06:00:00.000Z' }, np: { status: 'pending', createdAt: '2026-06-25T07:00:00.000Z' } }), 'pd', 'older PD job is selected before newer NP');
assert.equal(selectWithJobs({ pd: { status: 'completed', createdAt: '2026-06-25T06:00:00.000Z' }, np: { status: 'pending', createdAt: '2026-06-25T07:00:00.000Z' } }), 'np', 'NP is selected after PD completes');
assert.equal(selectWithJobs({ pd: { status: 'waiting_retry', createdAt: '2026-06-25T06:00:00.000Z', retryAfter: '2999-01-01T00:00:00.000Z' }, np: { status: 'pending', createdAt: '2026-06-25T07:00:00.000Z' } }), 'np', 'future waiting_retry does not block pending jobs');

console.log('rating_jobs_retry.test.js: all assertions passed');
