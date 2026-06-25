/**
 * Builds a retry queue snapshot without depending on Apps Script services.
 *
 * @param {number[]} retryQueueSnapshot Full retry queue at the beginning of the round.
 * @param {number} retryQueuePosition Zero-based position of the next queue item to process.
 * @param {number[]} retryFailedIndexes Indexes that failed again during the current retry round.
 * @param {string} stopMode One of: 'paused', 'round_completed', 'round_limit_reached'.
 * @returns {{retryQueue: number[], status: string}}
 */
function buildRetryQueueState_(
  retryQueueSnapshot,
  retryQueuePosition,
  retryFailedIndexes,
  stopMode
) {
  var snapshot = Array.isArray(retryQueueSnapshot) ? retryQueueSnapshot.slice() : [];
  var failed = Array.isArray(retryFailedIndexes) ? retryFailedIndexes.slice() : [];
  var position = Math.max(0, Math.min(Number(retryQueuePosition) || 0, snapshot.length));
  var mode = stopMode || 'paused';
  var nextQueue;

  if (mode === 'round_completed' || mode === 'round_limit_reached') {
    nextQueue = dedupeRetryIndexes_(failed);
  } else {
    nextQueue = mergeRetryIndexes_(failed, snapshot.slice(position), snapshot);
  }

  return {
    retryQueue: nextQueue,
    status: mode === 'round_limit_reached' && nextQueue.length > 0
      ? 'completed_with_errors'
      : 'pending_retry'
  };
}

function mergeRetryIndexes_(failed, unprocessed, originalOrder) {
  var included = {};
  var byOriginalOrder = [];
  var outsideOriginalOrder = [];
  var candidates = failed.concat(unprocessed);

  for (var i = 0; i < originalOrder.length; i += 1) {
    var index = originalOrder[i];
    if (candidates.indexOf(index) !== -1 && !included[index]) {
      byOriginalOrder.push(index);
      included[index] = true;
    }
  }

  for (var j = 0; j < candidates.length; j += 1) {
    var candidate = candidates[j];
    if (!included[candidate]) {
      outsideOriginalOrder.push(candidate);
      included[candidate] = true;
    }
  }

  return byOriginalOrder.concat(outsideOriginalOrder);
}

function dedupeRetryIndexes_(indexes) {
  var included = {};
  var result = [];
  for (var i = 0; i < indexes.length; i += 1) {
    var index = indexes[i];
    if (!included[index]) {
      result.push(index);
      included[index] = true;
    }
  }
  return result;
}

if (typeof module !== 'undefined') {
  module.exports = {
    buildRetryQueueState_: buildRetryQueueState_
  };
}
