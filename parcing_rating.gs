const CONFIG = {
  menuName: 'TEMED',
  sourceSheetName: 'Врачи',
  targetSheetName: 'Рейтинг',
  logSheetName: 'Log',
  announcementWebhookUrl: 'https://n8n-x3.tech.temed.ru/webhook-test/57353eb1-2f1c-4f4c-ab0a-995c84a617cf',

  sourceDoctorHeaders: ['Врач', 'ФИО', 'Доктор'],

  sourceHeaders: {
    pd: 'Ссылка ПД',
    np: 'Ссылка НП',
    sz: 'Ссылка СЗ'
  },

  targetHeaders: [
    'Дата',
    'Врач',
    'Рейтинг ПД',
    'Отзывы ПД',
    'Клиники ПД',
    'Рейтинг НП',
    'Отзывы НП',
    'Клиники НП',
    'Рейтинг СЗ',
    'Отзывы СЗ',
    'Клиники СЗ'
  ],

  aggregators: {
    pd: {
      title: 'ПД',
      sourceHeader: 'Ссылка ПД',
      ratingHeader: 'Рейтинг ПД',
      reviewsHeader: 'Отзывы ПД',
      clinicsHeader: 'Клиники ПД'
    },
    np: {
      title: 'НП',
      sourceHeader: 'Ссылка НП',
      ratingHeader: 'Рейтинг НП',
      reviewsHeader: 'Отзывы НП',
      clinicsHeader: 'Клиники НП'
    },
    sz: {
      title: 'СЗ',
      sourceHeader: 'Ссылка СЗ',
      ratingHeader: 'Рейтинг СЗ',
      reviewsHeader: 'Отзывы СЗ',
      clinicsHeader: 'Клиники СЗ'
    }
  },

  ratingJobs: {
    maxRetryRounds: {
      pd: 1,
      np: 1,
      sz: 1
    },
    storagePrefix: 'rating_job_',
    batchMaxRuntimeMs: 270000
  },

  fetchOptions: {
    maxAttempts: 3,
    baseDelayMs: 1200,
    maxDelayMs: 10000,
    jitterMs: 700,
    requestDelayMs: 350
  },

  ratingJobs: {
    safeExecutionLimitMs: 240000,
    staleRunningAfterMs: 900000,
    continuationIntervalMinutes: 5,
    batchSize: { pd: 5, np: 15, sz: 15 },
    maxRetryRounds: { pd: 2, np: 1, sz: 1 },
    retryCooldownMs: 900000,
    pd: {
      directAttempts: 1,
      reserveAttempts: 3,
      consecutiveDirect403Limit: 3,
      consecutiveFullFailureLimit: 3,
      retryCooldownMs: 900000,
      requestDelayMinMs: 1500,
      requestDelayMaxMs: 3000
    }
  }
};

/**
 * Меню при открытии.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.menuName)
    .addItem('Обновить ПД', 'startPdRatingUpdate')
    .addItem('Обновить НП', 'startNpRatingUpdate')
    .addItem('Обновить СЗ', 'startSzRatingUpdate')
    .addSeparator()
    .addItem('Обновить все рейтинги', 'updateAllRatings')
    .addItem('Проверить статус', 'showRatingJobsStatus')
    .addItem('Показать лог', 'showRatingLog')
    .addItem('Тест загрузки URL', 'testRatingUrlFetch')
    .addItem('Настроить расписание обновлений 6/7/8', 'setupRatingUpdateTriggers')
    .addItem('Отправить анонс', 'sendRatingAnnouncement')
    .addToUi();
}

function updatePdRatings() { startPdRatingUpdate(); }
function updateNpRatings() { startNpRatingUpdate(); }
function updateSzRatings() { startSzRatingUpdate(); }

function startPdRatingUpdate() { startRatingJob_('pd'); }
function startNpRatingUpdate() { startRatingJob_('np'); }
function startSzRatingUpdate() { startRatingJob_('sz'); }

function updateAllRatings() {
  createOrContinueRatingJob_('pd');
  createOrContinueRatingJob_('np');
  createOrContinueRatingJob_('sz');
  resumePendingRatingJobs();
}

function sendRatingAnnouncement() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.targetSheetName);

  if (!sheet) {
    ui.alert('Лист "' + CONFIG.targetSheetName + '" не найден.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    ui.alert('На листе "' + CONFIG.targetSheetName + '" нет данных для отправки.');
    return;
  }

  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(value) {
    return normalizeText_(value);
  });
  var dateColumnIndex = headers.indexOf('Дата');
  if (dateColumnIndex === -1) {
    ui.alert('На листе "' + CONFIG.targetSheetName + '" не найден столбец "Дата".');
    return;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  var today = new Date();
  var todayKey = toDateKey_(today);

  var latestDate = null;
  for (var i = 0; i < values.length; i++) {
    var rowDate = parseDateValue_(values[i][dateColumnIndex]);
    if (!rowDate) {
      continue;
    }

    if (!latestDate || rowDate.getTime() > latestDate.getTime()) {
      latestDate = rowDate;
    }
  }

  if (!latestDate) {
    ui.alert('На листе "' + CONFIG.targetSheetName + '" нет валидных дат в столбце "Дата".');
    return;
  }

  var latestDateKey = toDateKey_(latestDate);
  if (latestDateKey !== todayKey) {
    var button = ui.alert(
      'Внимание: дата не совпадает',
      'Дата запуска: ' + todayKey + '. Последняя дата на листе "' + CONFIG.targetSheetName + '": ' + latestDateKey + '. Продолжить отправку?',
      ui.ButtonSet.YES_NO
    );

    if (button !== ui.Button.YES) {
      ui.alert('Отправка отменена пользователем.');
      return;
    }
  }

  var weekStart = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 6);
  var weekEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59, 999);
  var rowsForWebhook = [];

  for (var j = 0; j < values.length; j++) {
    var itemDate = parseDateValue_(values[j][dateColumnIndex]);
    if (!itemDate) {
      continue;
    }

    if (itemDate.getTime() < weekStart.getTime() || itemDate.getTime() > weekEnd.getTime()) {
      continue;
    }

    rowsForWebhook.push(mapRowToWebhookObject_(headers, values[j]));
  }

  if (rowsForWebhook.length === 0) {
    ui.alert('За последнюю неделю нет данных для отправки.');
    return;
  }

  var payload = {
    sheetName: CONFIG.targetSheetName,
    generatedAt: new Date().toISOString(),
    period: {
      from: toDateKey_(weekStart),
      to: toDateKey_(today)
    },
    latestSheetDate: latestDateKey,
    rows: rowsForWebhook
  };

  var response = UrlFetchApp.fetch(CONFIG.announcementWebhookUrl, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    ui.alert('Ошибка отправки: HTTP ' + statusCode + '. Ответ: ' + response.getContentText());
    return;
  }

  ui.alert('Анонс успешно отправлен. Передано строк: ' + rowsForWebhook.length + '.');
}


/**
 * Основной обработчик новой очереди. aggregatorKeys сохранен только для обратной совместимости:
 * за одно выполнение обрабатывается первый переданный агрегатор.
 */
function updateRatings_(aggregatorKeys) {
  var key = aggregatorKeys && aggregatorKeys.length ? aggregatorKeys[0] : 'pd';
  startRatingJob_(key);
}

function startRatingJob_(aggregatorKey) {
  createOrContinueRatingJob_(aggregatorKey);
  processOneRatingJob_(aggregatorKey);
}

function resumePendingRatingJobs() {
  var selected = selectPendingRatingJob_();
  if (!selected) {
    var logSheet = getOrCreateLogSheet_();
    var logs = [];
    addLog_(logs, 'DEBUG', 'Нет незавершённых заданий рейтингов', {});
    flushLogs_(logSheet, logs);
    return;
  }
  processOneRatingJob_(selected);
}

function setupRatingUpdateTriggers() {
  var obsolete = {
    runSequentialUpdate: true,
    continueSequentialUpdate: true,
    retryFailedStep: true,
    updatePdRatings: true,
    updateNpRatings: true,
    updateSzRatings: true,
    updateAllRatingsScheduled: true,
    scheduledRatingUpdate: true,
    startPdRatingUpdate: true,
    startNpRatingUpdate: true,
    startSzRatingUpdate: true,
    resumePendingRatingJobs: true
  };

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    var handler = trigger.getHandlerFunction && trigger.getHandlerFunction();
    if (obsolete[handler]) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('startPdRatingUpdate').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('startNpRatingUpdate').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('startSzRatingUpdate').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('resumePendingRatingJobs').timeBased()
    .everyMinutes(CONFIG.ratingJobs.continuationIntervalMinutes).create();

  SpreadsheetApp.getUi().alert('Расписание обновлений настроено: ПД 06:00, НП 07:00, СЗ 08:00, продолжение каждые 5 минут.');
}

function createOrContinueRatingJob_(aggregatorKey) {
  assertAggregatorKey_(aggregatorKey);
  var job = getRatingJob_(aggregatorKey);
  var todayKey = toDateKey_(new Date());
  var nowIso = new Date().toISOString();

  if (job.date && job.date !== todayKey && job.status !== 'completed' && job.status !== 'completed_with_errors') {
    job.status = 'completed_with_errors';
    job.updatedAt = nowIso;
    saveRatingJob_(aggregatorKey, job);
  }

  if (job.date === todayKey && (job.status === 'completed' || job.status === 'completed_with_errors')) {
    return job;
  }

  if (job.date === todayKey) {
    if (job.status === 'running' && isStaleJob_(job)) {
      job.status = 'pending';
      job.updatedAt = nowIso;
      saveRatingJob_(aggregatorKey, job);
    }
    return job;
  }

  job = {
    status: 'pending',
    date: todayKey,
    nextSourceIndex: 0,
    failedIndexes: [],
    retryRound: 0,
    retryAfter: '',
    createdAt: nowIso,
    updatedAt: nowIso,
    processed: 0,
    total: 0,
    permanentErrors: 0,
    temporaryErrors: 0,
    preservedPrevious: 0
  };
  saveRatingJob_(aggregatorKey, job);
  return job;
}

function processOneRatingJob_(aggregatorKey) {
  assertAggregatorKey_(aggregatorKey);
  var lock = LockService.getScriptLock();
  var logSheet = getOrCreateLogSheet_();
  var logs = [];
  var lockAcquired = false;
  var runId = Utilities.getUuid();
  var executionStartedAt = Date.now();
  var ctx = createExecutionContext_(executionStartedAt, logs, runId, aggregatorKey);

  try {
    lockAcquired = lock.tryLock(5000);
    if (!lockAcquired) {
      addLog_(logs, 'INFO', 'Другое задание уже выполняется, повторит постоянный триггер', { runId: runId, aggregator: aggregatorKey });
      return;
    }

    var job = createOrContinueRatingJob_(aggregatorKey);
    if (job.status === 'completed' || job.status === 'completed_with_errors') {
      addLog_(logs, 'INFO', 'Задание за текущую дату уже завершено', { runId: runId, aggregator: aggregatorKey, date: job.date, status: job.status });
      return;
    }
    if (job.status === 'waiting_retry' && job.retryAfter && Date.now() < Date.parse(job.retryAfter)) {
      addLog_(logs, 'INFO', 'Задание ожидает времени повтора', { runId: runId, aggregator: aggregatorKey, retryAfter: job.retryAfter });
      return;
    }
    if (job.status === 'running' && isStaleJob_(job)) {
      addLog_(logs, 'WARN', 'Восстановление задания после аварийного завершения', { runId: runId, aggregator: aggregatorKey, updatedAt: job.updatedAt });
      job.status = 'pending';
    }

    job.status = 'running';
    job.updatedAt = new Date().toISOString();
    saveRatingJob_(aggregatorKey, job);

    processRatingJobBatch_(aggregatorKey, job, ctx);
  } catch (error) {
    var currentJob = getRatingJob_(aggregatorKey);
    currentJob.status = 'pending';
    currentJob.updatedAt = new Date().toISOString();
    saveRatingJob_(aggregatorKey, currentJob);
    addLog_(logs, 'ERROR', 'Критическая ошибка алгоритма обновления рейтингов', { runId: runId, aggregator: aggregatorKey, error: error && error.message ? error.message : String(error) });
    throw error;
  } finally {
    flushLogs_(logSheet, logs);
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function processRatingJobBatch_(aggregatorKey, job, ctx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(CONFIG.sourceSheetName);
  if (!sourceSheet) throw new Error('Не найден лист "' + CONFIG.sourceSheetName + '"');
  var targetSheet = getOrCreateTargetSheet_();
  var decimalSeparator = getDecimalSeparator_();
  var sourceObjects = getSheetObjects_(sourceSheet);
  var sourceHeaderRow = getHeaderRow_(sourceSheet);
  var doctorHeader = findFirstHeader_(sourceHeaderRow, CONFIG.sourceDoctorHeaders);
  if (!doctorHeader) throw new Error('На листе "' + CONFIG.sourceSheetName + '" не найдена колонка с именем врача.');

  var targetData = getExistingTargetData_(targetSheet);
  ensureTodayRows_(targetSheet, targetData, sourceObjects, doctorHeader, job.date);
  targetData = getExistingTargetData_(targetSheet);

  job.total = sourceObjects.length;
  var agg = CONFIG.aggregators[aggregatorKey];
  var columns = getAggregatorTargetColumns_(targetSheet, agg);
  var startIndex = Number(job.nextSourceIndex) || 0;
  var indexes = buildIndexesForCurrentStage_(job, sourceObjects.length);
  var isRetryStage = Number(job.retryRound) > 0;
  if (isRetryStage) {
    job.failedIndexes = [];
  }
  var processedInRun = 0;
  var success = 0;
  var preserved = 0;
  var permanentErrors = 0;
  var temporaryErrors = 0;
  var batchLimit = CONFIG.ratingJobs.batchSize[aggregatorKey] || 10;
  var nextPointer = startIndex;
  var failedThisRound = [];

  addLog_(ctx.logs, 'INFO', 'Старт порции обновления рейтингов', { runId: ctx.runId, date: job.date, aggregator: aggregatorKey, startIndex: startIndex, retryRound: job.retryRound });

  for (var p = startIndex; p < indexes.length; p++) {
    if (!hasExecutionTime_(ctx, 12000) || processedInRun >= batchLimit) {
      job.status = 'pending';
      job.nextSourceIndex = p;
      job.updatedAt = new Date().toISOString();
      saveRatingJob_(aggregatorKey, job);
      addLog_(ctx.logs, 'INFO', 'Порция завершена, задание будет продолжено', buildRunStats_(ctx, job, startIndex, p, processedInRun, success, preserved, permanentErrors, temporaryErrors, 'time_or_batch_limit'));
      return;
    }

    var sourceIndex = indexes[p];
    var sourceRow = sourceObjects[sourceIndex];
    var rowNumber = sourceIndex + 2;
    var doctorName = normalizeText_(sourceRow[doctorHeader]);
    nextPointer = p + 1;
    if (!doctorName) {
      job.nextSourceIndex = nextPointer;
      saveRatingJob_(aggregatorKey, job);
      continue;
    }

    var todayEntry = targetData.byDoctorDate[doctorName + '||' + job.date];
    if (!todayEntry) throw new Error('Не найдена строка текущего дня для врача: ' + doctorName);
    var url = normalizeUrl_(sourceRow[agg.sourceHeader]);
    var result = processAggregatorDoctor_(aggregatorKey, url, doctorName, rowNumber, decimalSeparator, ctx);
    processedInRun++;

    if (result.status === 'success') {
      writeAggregatorValues_(targetSheet, todayEntry.arrayIndex + 2, columns, result.values);
      success++;
    } else if (result.status === 'empty') {
      preserved++;
    } else if (result.status === 'permanent') {
      permanentErrors++;
      preserved++;
    } else {
      temporaryErrors++;
      preserved++;
      failedThisRound.push(sourceIndex);
      if (aggregatorKey === 'pd') {
        ctx.consecutiveFullFailures++;
        if (ctx.consecutiveFullFailures >= CONFIG.ratingJobs.pd.consecutiveFullFailureLimit) {
          job.status = 'waiting_retry';
          job.retryAfter = new Date(Date.now() + CONFIG.ratingJobs.pd.retryCooldownMs).toISOString();
          job.nextSourceIndex = nextPointer;
          job.failedIndexes = mergeUniqueIndexes_(job.failedIndexes, failedThisRound);
          job.updatedAt = new Date().toISOString();
          saveRatingJob_(aggregatorKey, job);
          addLog_(ctx.logs, 'WARN', 'Массовые технические ошибки ПД, задание поставлено на паузу', buildRunStats_(ctx, job, startIndex, nextPointer, processedInRun, success, preserved, permanentErrors, temporaryErrors, 'waiting_retry'));
          return;
        }
      }
    }
    if (result.status === 'success') ctx.consecutiveFullFailures = 0;

    job.nextSourceIndex = nextPointer;
    job.failedIndexes = isRetryStage ? failedThisRound.slice() : mergeUniqueIndexes_(job.failedIndexes, failedThisRound);
    job.processed = Math.max(Number(job.processed) || 0, sourceIndex + 1);
    job.permanentErrors = permanentErrors;
    job.temporaryErrors = temporaryErrors;
    job.preservedPrevious = preserved;
    job.updatedAt = new Date().toISOString();
    saveRatingJob_(aggregatorKey, job);

    if (aggregatorKey === 'pd' && hasExecutionTime_(ctx, CONFIG.ratingJobs.pd.requestDelayMaxMs + 5000)) {
      sleepMs_(randomInt_(CONFIG.ratingJobs.pd.requestDelayMinMs, CONFIG.ratingJobs.pd.requestDelayMaxMs));
    }
  }

  if (job.retryRound < (CONFIG.ratingJobs.maxRetryRounds[aggregatorKey] || 0) && (job.failedIndexes || []).length > 0) {
    job.retryRound = Number(job.retryRound) + 1;
    job.nextSourceIndex = 0;
    job.status = 'waiting_retry';
    job.retryAfter = new Date(Date.now() + (CONFIG.ratingJobs.retryCooldownMs || 900000)).toISOString();
    job.updatedAt = new Date().toISOString();
    saveRatingJob_(aggregatorKey, job);
    addLog_(ctx.logs, 'WARN', 'Основной список завершён, запланирован отдельный раунд повторов', buildRunStats_(ctx, job, startIndex, indexes.length, processedInRun, success, preserved, permanentErrors, temporaryErrors, 'retry_round_scheduled'));
    return;
  }

  job.status = (job.failedIndexes || []).length > 0 ? 'completed_with_errors' : 'completed';
  job.nextSourceIndex = indexes.length;
  job.updatedAt = new Date().toISOString();
  saveRatingJob_(aggregatorKey, job);
  addLog_(ctx.logs, 'INFO', 'Обновление агрегатора завершено', buildRunStats_(ctx, job, startIndex, indexes.length, processedInRun, success, preserved, permanentErrors, temporaryErrors, job.status));
}


function assertAggregatorKey_(key) {
  if (!CONFIG.aggregators[key]) throw new Error('Неизвестный агрегатор: ' + key);
}

function createExecutionContext_(startedAt, logs, runId, aggregatorKey) {
  return { startedAt: startedAt, logs: logs, runId: runId, aggregatorKey: aggregatorKey, consecutiveDirect403: 0, directDisabled: false, consecutiveFullFailures: 0 };
}

function hasExecutionTime_(ctx, reserveMs) {
  return Date.now() - ctx.startedAt + (Number(reserveMs) || 0) < CONFIG.ratingJobs.safeExecutionLimitMs;
}

function getRatingJob_(aggregatorKey) {
  var props = PropertiesService.getScriptProperties();
  var prefix = 'ratingJob.' + aggregatorKey + '.';
  var all = props.getProperties();
  return {
    status: all[prefix + 'status'] || '',
    date: all[prefix + 'date'] || '',
    nextSourceIndex: Number(all[prefix + 'nextSourceIndex'] || 0),
    failedIndexes: parseJsonArray_(all[prefix + 'failedIndexes']),
    retryRound: Number(all[prefix + 'retryRound'] || 0),
    retryAfter: all[prefix + 'retryAfter'] || '',
    createdAt: all[prefix + 'createdAt'] || '',
    updatedAt: all[prefix + 'updatedAt'] || '',
    processed: Number(all[prefix + 'processed'] || 0),
    total: Number(all[prefix + 'total'] || 0),
    permanentErrors: Number(all[prefix + 'permanentErrors'] || 0),
    temporaryErrors: Number(all[prefix + 'temporaryErrors'] || 0),
    preservedPrevious: Number(all[prefix + 'preservedPrevious'] || 0)
  };
}

function saveRatingJob_(aggregatorKey, job) {
  var prefix = 'ratingJob.' + aggregatorKey + '.';
  var props = {};
  Object.keys(job).forEach(function(key) {
    props[prefix + key] = Array.isArray(job[key]) ? JSON.stringify(job[key]) : String(job[key] === undefined || job[key] === null ? '' : job[key]);
  });
  PropertiesService.getScriptProperties().setProperties(props);
}

function parseJsonArray_(value) {
  try { var parsed = JSON.parse(value || '[]'); return Array.isArray(parsed) ? parsed : []; } catch (e) { return []; }
}

function isStaleJob_(job) {
  return job.updatedAt && Date.now() - Date.parse(job.updatedAt) > CONFIG.ratingJobs.staleRunningAfterMs;
}

function selectPendingRatingJob_() {
  var keys = ['pd', 'np', 'sz'];
  var jobs = [];
  keys.forEach(function(key) {
    var job = getRatingJob_(key);
    if (job.status === 'running' && isStaleJob_(job)) {
      job.status = 'pending';
      saveRatingJob_(key, job);
    }
    if (job.status === 'pending' || (job.status === 'waiting_retry' && (!job.retryAfter || Date.now() >= Date.parse(job.retryAfter)))) {
      jobs.push({ key: key, createdAt: job.createdAt || '9999-12-31T23:59:59.999Z' });
    }
  });
  jobs.sort(function(a, b) {
    if (a.createdAt < b.createdAt) return -1;
    if (a.createdAt > b.createdAt) return 1;
    return keys.indexOf(a.key) - keys.indexOf(b.key);
  });
  return jobs.length ? jobs[0].key : '';
}

function ensureTodayRows_(sheet, targetData, sourceObjects, doctorHeader, dateKey) {
  var rowsToAppend = [];
  var seen = {};
  for (var i = 0; i < sourceObjects.length; i++) {
    var doctor = normalizeText_(sourceObjects[i][doctorHeader]);
    if (!doctor || seen[doctor]) continue;
    seen[doctor] = true;
    if (targetData.byDoctorDate[doctor + '||' + dateKey]) continue;
    var obj = createEmptyTargetRowObject_();
    var latest = targetData.latestByDoctor[doctor];
    if (latest) copyTargetRow_(latest.rowObj, obj);
    obj['Дата'] = parseDateValue_(dateKey) || new Date();
    obj['Врач'] = doctor;
    rowsToAppend.push(CONFIG.targetHeaders.map(function(header) { return obj[header]; }));
  }
  if (rowsToAppend.length) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, CONFIG.targetHeaders.length).setValues(rowsToAppend);
}

function getAggregatorTargetColumns_(sheet, agg) {
  var idx = getHeaderIndexes_(sheet, [agg.ratingHeader, agg.reviewsHeader, agg.clinicsHeader]);
  return [idx[agg.ratingHeader], idx[agg.reviewsHeader], idx[agg.clinicsHeader]];
}

function writeAggregatorValues_(sheet, row, columns, values) {
  if (columns[1] === columns[0] + 1 && columns[2] === columns[1] + 1) {
    sheet.getRange(row, columns[0], 1, 3).setValues([[values.rating, values.reviews, values.clinics]]);
  } else {
    sheet.getRange(row, columns[0]).setValue(values.rating);
    sheet.getRange(row, columns[1]).setValue(values.reviews);
    sheet.getRange(row, columns[2]).setValue(values.clinics);
  }
}

function buildIndexesForCurrentStage_(job, total) {
  if (Number(job.retryRound) > 0) return (job.failedIndexes || []).slice();
  var result = [];
  for (var i = 0; i < total; i++) result.push(i);
  return result;
}

function mergeUniqueIndexes_(a, b) {
  var seen = {}, result = [];
  (a || []).concat(b || []).forEach(function(v) { var n = Number(v); if (!isNaN(n) && !seen[n]) { seen[n] = true; result.push(n); } });
  return result;
}

function buildRunStats_(ctx, job, startIndex, endIndex, processed, success, preserved, permanentErrors, temporaryErrors, reason) {
  return { runId: ctx.runId, date: job.date, aggregator: ctx.aggregatorKey, startIndex: startIndex, endIndex: endIndex, processed: processed, success: success, preservedPrevious: preserved, permanentErrors: permanentErrors, temporaryErrors: temporaryErrors, elapsedMs: Date.now() - ctx.startedAt, stopReason: reason, failedIndexes: job.failedIndexes || [] };
}

function processAggregatorDoctor_(key, url, doctorName, rowNumber, decimalSeparator, ctx) {
  var agg = CONFIG.aggregators[key];
  if (!url) {
    addLog_(ctx.logs, 'INFO', 'Пустая ссылка', { runId: ctx.runId, row: rowNumber, doctor: doctorName, aggregator: agg.title });
    return { status: 'empty' };
  }
  if (!isValidHttpUrl_(url)) {
    addLog_(ctx.logs, 'ERROR', 'Невалидный URL, сохранены предыдущие значения', { runId: ctx.runId, row: rowNumber, doctor: doctorName, aggregator: agg.title, url: url });
    return { status: 'permanent' };
  }
  var fetched = key === 'pd' ? fetchPdHtml_(url, ctx) : (key === 'np' ? fetchNpHtml_(url, ctx) : fetchSzHtml_(url, ctx));
  if (!fetched.ok) {
    var level = fetched.permanent ? 'ERROR' : 'WARN';
    addLog_(ctx.logs, level, (fetched.permanent ? 'Постоянная ошибка — сохранены предыдущие значения' : 'Техническая ошибка — сохранены предыдущие значения'), { runId: ctx.runId, row: rowNumber, doctor: doctorName, aggregator: agg.title, url: url, status: fetched.statusCode || '', error: fetched.error || '', needCheckSourceUrl: fetched.statusCode === 404 });
    return { status: fetched.permanent ? 'permanent' : 'temporary' };
  }
  var parsed = parseAggregatorData_(key, fetched.html);
  if (!isParsedResultValid_(key, fetched.html, parsed)) {
    addLog_(ctx.logs, 'ERROR', 'INVALID_HTML_OR_PARSE_FAILED — сохранены предыдущие значения', { runId: ctx.runId, row: rowNumber, doctor: doctorName, aggregator: agg.title, url: url });
    return { status: 'temporary' };
  }
  addLog_(ctx.logs, 'INFO', 'Данные извлечены', { runId: ctx.runId, row: rowNumber, doctor: doctorName, aggregator: agg.title, url: url, method: fetched.method, attempt: fetched.attempt });
  return { status: 'success', values: { rating: parsed.rating ? formatRatingForLocale_(parsed.rating, decimalSeparator) : '', reviews: parsed.reviews || '', clinics: parsed.clinics || '' } };
}

function isParsedResultValid_(key, html, parsed) {
  if (parsed && (parsed.rating || parsed.reviews || parsed.clinics)) return true;
  if (!html || /captcha|cloudflare|access denied|доступ ограничен|проверка безопасности/i.test(html)) return false;
  if (key === 'pd') return /prodoctorov|Отзывы|Рейтинг|b-doctor/i.test(html);
  if (key === 'np') return /napopravku|itemprop="ratingValue"|doctor-detail/i.test(html);
  if (key === 'sz') return /sberhealth|docdoc|doctor-page|reviewCount/i.test(html);
  return false;
}


function showRatingJobsStatus() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sourceSheetName);
  var total = sourceSheet ? Math.max(0, sourceSheet.getLastRow() - 1) : 0;
  var labels = { pd: 'ПроДокторов', np: 'НаПоправку', sz: 'СберЗдоровье' };
  var lines = ['Статус заданий обновления рейтингов', ''];
  ['pd', 'np', 'sz'].forEach(function(key) {
    var job = getRatingJob_(key);
    var failed = job.failedIndexes || [];
    lines.push(labels[key] + ':');
    lines.push('  дата задания: ' + (job.date || '—'));
    lines.push('  статус: ' + (job.status || '—'));
    lines.push('  обработано врачей: ' + (job.processed || 0));
    lines.push('  всего врачей: ' + (job.total || total));
    lines.push('  текущая позиция: ' + (job.nextSourceIndex || 0));
    lines.push('  количество ошибок: ' + ((job.permanentErrors || 0) + (job.temporaryErrors || 0)));
    lines.push('  врачей в очереди повторов: ' + failed.length);
    lines.push('  последнее изменение: ' + (job.updatedAt || '—'));
    lines.push('  следующий повтор: ' + (job.retryAfter || '—'));
    lines.push('');
  });
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

function showRatingLog() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(getOrCreateLogSheet_());
}

function testRatingUrlFetch() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Тест загрузки URL', 'Введите URL для проверки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var url = normalizeUrl_(response.getResponseText());
  var logs = [];
  var ctx = createExecutionContext_(Date.now(), logs, Utilities.getUuid(), 'pd');
  var result = fetchPdHtml_(url, ctx);
  flushLogs_(getOrCreateLogSheet_(), logs);
  ui.alert(result.ok ? ('URL загружен успешно через ' + result.method + ', символов: ' + result.html.length) : ('Ошибка загрузки: ' + (result.statusCode || '') + ' ' + (result.error || '')));
}

function showErrorsModal_(errors) {
  if (!errors || errors.length === 0) {
    return;
  }

  var lines = ['Во время обновления возникли ошибки (' + errors.length + '):', ''];

  for (var i = 0; i < errors.length; i++) {
    var item = errors[i];
    var suffix = item.oldValuesKept
      ? ' (HTTP 403: сохранены значения за текущую дату)'
      : '';

    lines.push(
      [
        i + 1 + '.',
        'Строка ' + item.row + ',',
        'врач: ' + item.doctor + ',',
        'агрегатор: ' + item.aggregator + ',',
        'ошибка: ' + item.message + suffix
      ].join(' ')
    );
  }

  var message = lines.join('\n');
  var html = HtmlService
    .createHtmlOutput('<pre style="white-space: pre-wrap; font-family: Arial, sans-serif;">' + escapeHtml_(message) + '</pre>')
    .setWidth(800)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Ошибки обновления рейтингов');
}

function escapeHtml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Парсер по агрегатору.
 */
function parseAggregatorData_(key, html) {
  if (key === 'pd') {
    return {
      rating: extractPdRating_(html),
      reviews: extractPdReviews_(html),
      clinics: extractPdClinics_(html)
    };
  }

  if (key === 'np') {
    return {
      rating: extractNpRating_(html),
      reviews: extractNpReviews_(html),
      clinics: extractNpClinics_(html)
    };
  }

  if (key === 'sz') {
    return {
      rating: extractSzRating_(html),
      reviews: extractSzReviews_(html),
      clinics: extractSzClinics_(html)
    };
  }

  throw new Error('Неизвестный агрегатор: ' + key);
}

/* =========================
   ПД
   ========================= */

/**
 * Рейтинг ПД: по уже действующей логике.
 */
function extractPdRating_(html) {
  if (!html) {
    return '';
  }

  var firstIdx = html.indexOf('Рейтинг');
  if (firstIdx === -1) {
    return '';
  }

  var secondIdx = html.indexOf('Рейтинг', firstIdx + 'Рейтинг'.length);
  if (secondIdx === -1) {
    return '';
  }

  var tail = html.slice(secondIdx);
  var match = tail.match(/text-h5\s+text--text\s+font-weight-medium\s+mr-2[^>]*>\s*([0-9]+(?:[\.,][0-9]+)?)/i);

  return match && match[1] ? match[1] : '';
}

/**
 * Отзывы ПД.
 */
function extractPdReviews_(html) {
  if (!html) {
    return '';
  }

  var match = html.match(/Отзывы\s*<\/div>\s*<div[^>]*class="[^"]*b-doctor-details__toc-num[^"]*"[^>]*>\s*([0-9]+)/i);
  return match && match[1] ? match[1] : '';
}

/**
 * Клиники ПД:

 */
function extractPdClinics_(html) {
  if (!html) {
    return '';
  }

  var matches = [];
  var match;

  function decodeHtmlAttrRaw_(text) {
    return String(text || '')
      .replace(/&quot;/gi, '"')
      .replace(/&#34;/gi, '"')
      .replace(/&#x22;/gi, '"')
      .replace(/&amp;/gi, '&')
      .replace(/&#38;/gi, '&')
      .replace(/&#x26;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&#60;/gi, '<')
      .replace(/&#x3c;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&#62;/gi, '>')
      .replace(/&#x3e;/gi, '>')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&#160;/gi, ' ')
      .replace(/&#xA0;/gi, ' ')
      .replace(/&#39;/gi, "'")
      .replace(/&#x27;/gi, "'")
      .replace(/&#92;/gi, '\\')
      .replace(/&#x5c;/gi, '\\')
      .replace(/&#(\d+);/g, function(_, code) {
        return String.fromCharCode(Number(code));
      })
      .replace(/&#x([0-9a-f]+);/gi, function(_, code) {
        return String.fromCharCode(parseInt(code, 16));
      });
  }

  // 1) Основной источник: :lpu-address-list
  var lpuAttrMatch =
    html.match(/:lpu-address-list\s*=\s*"([\s\S]*?)"\s*:synonym-price-list=/i) ||
    html.match(/:lpu-address-list\s*=\s*"([\s\S]*?)"\s*:is-appointment-on=/i) ||
    html.match(/:lpu-address-list\s*=\s*"([\s\S]*?)"/i);

  if (lpuAttrMatch && lpuAttrMatch[1]) {
    var rawAttr = lpuAttrMatch[1];
    var decodedAttr = decodeHtmlAttrRaw_(rawAttr);
    var parsed = null;

    try {
      parsed = JSON.parse(decodedAttr);
    } catch (error) {
      parsed = null;
    }

    if (parsed && parsed.length) {
      for (var i = 0; i < parsed.length; i++) {
        var item = parsed[i];
        if (item && item.lpu && item.lpu.name) {
          matches.push(cleanExtractedText_(item.lpu.name));
        }
      }
    } else {
      // Безопасный fallback: ищем именно lpu.name после блока town,
      // чтобы не схватить town.name = "Санкт-Петербург"
      var safeLpuNameRegex =
        /"lpu"\s*:\s*\{[\s\S]*?"town"\s*:\s*\{[\s\S]*?"translations"\s*:\s*\{\}\s*\}\s*,\s*"name"\s*:\s*"((?:\\.|[^"\\])*)"/gi;

      while ((match = safeLpuNameRegex.exec(decodedAttr)) !== null) {
        if (match[1]) {
          matches.push(cleanJsonText_(match[1]));
        }
      }
    }

    var lpuResult = uniqueJoin_(matches);
    if (lpuResult) {
      return lpuResult;
    }
  }

  // 2) Fallback: data-review-power-info-open
  matches = [];
  var reviewInfoRegex = /data-review-power-info-open\s*=\s*"([\s\S]*?)"/gi;
  while ((match = reviewInfoRegex.exec(html)) !== null) {
    var decodedReviewInfo = decodeHtmlAttrRaw_(match[1]);
    var reviewNameRegex = /name\s*:\s*'([^']+)'/gi;
    var nameMatch;

    while ((nameMatch = reviewNameRegex.exec(decodedReviewInfo)) !== null) {
      if (nameMatch[1]) {
        matches.push(cleanExtractedText_(nameMatch[1]));
      }
    }
  }

  var reviewInfoResult = uniqueJoin_(matches);
  if (reviewInfoResult) {
    return reviewInfoResult;
  }

  // 3) Fallback: адреса в отзывах
  matches = [];
  var addressRegex = /<[^>]*class="[^"]*b-review-card__address[^"]*"[^>]*>([\s\S]*?)<\/[^>]+>/gi;
  while ((match = addressRegex.exec(html)) !== null) {
    if (!match[1]) {
      continue;
    }

    var addressText = cleanExtractedText_(match[1]);
    var clinicName = normalizeText_(addressText.split(' - ')[0]);
    if (clinicName) {
      matches.push(clinicName);
    }
  }

  return uniqueJoin_(matches);
}

/* =========================
   НП
   ========================= */

/**
 * Рейтинг НП.
 */
function extractNpRating_(html) {
  if (!html) {
    return '';
  }

  var match = html.match(/itemprop="ratingValue"\s+content="([0-9]+(?:[\.,][0-9]+)?)"/i) ||
              html.match(/content="([0-9]+(?:[\.,][0-9]+)?)"\s+itemprop="ratingValue"/i);

  return match && match[1] ? match[1] : '';
}

/**
 * Отзывы НП.
 */
function extractNpReviews_(html) {
  if (!html) {
    return '';
  }

  var match = html.match(/itemprop="ratingCount"\s+content="([0-9]+)"/i) ||
              html.match(/content="([0-9]+)"\s+itemprop="ratingCount"/i);

  if (match && match[1]) {
    return match[1];
  }

  match = html.match(/>\s*([0-9]+)\s*отзыв/i);
  return match && match[1] ? match[1] : '';
}

/**
 * Клиники НП: из ссылок doctor-detail-workplace__title-text.
 */
function extractNpClinics_(html) {
  if (!html) {
    return '';
  }

  var matches = [];
  var regex = /<a[^>]*class="[^"]*doctor-detail-workplace__title-text[^"]*"[^>]*>\s*([\s\S]*?)\s*<\/a>/gi;
  var match;

  while ((match = regex.exec(html)) !== null) {
    if (match[1]) {
      matches.push(cleanExtractedText_(match[1]));
    }
  }

  return uniqueJoin_(matches);
}

/* =========================
   СЗ
   ========================= */

/**
 * Рейтинг СЗ.
 */
function extractSzRating_(html) {
  if (!html) {
    return '';
  }

  var match = html.match(/itemprop="ratingValue"[^>]*content="([0-9]+(?:[\.,][0-9]+)?)"/i) ||
              html.match(/content="([0-9]+(?:[\.,][0-9]+)?)"[^>]*itemprop="ratingValue"/i) ||
              html.match(/itemProp="ratingValue"[^>]*content="([0-9]+(?:[\.,][0-9]+)?)"/i) ||
              html.match(/content="([0-9]+(?:[\.,][0-9]+)?)"[^>]*itemProp="ratingValue"/i);

  return match && match[1] ? match[1] : '';
}

/**
 * Отзывы СЗ.
 */
function extractSzReviews_(html) {
  if (!html) {
    return '';
  }

  var match = html.match(/itemprop="reviewCount"[^>]*content="([0-9]+)"/i) ||
              html.match(/content="([0-9]+)"[^>]*itemprop="reviewCount"/i) ||
              html.match(/itemProp="reviewCount"[^>]*content="([0-9]+)"/i) ||
              html.match(/content="([0-9]+)"[^>]*itemProp="reviewCount"/i);

  return match && match[1] ? match[1] : '';
}

/**
 * Клиники СЗ:
 */
function extractSzClinics_(html) {
  if (!html) {
    return '';
  }

  var clinics = [];
  var match;

  // 1. Основной источник: только clinic chips.
  // Якоримся не на <label>, а на input с data-testid клиники.
  var clinicChipRegex = /<input[^>]*data-testid="doctor-page_filters-clinic-chip-\d+"[^>]*>[\s\S]{0,2000}?<span[^>]*class="[^"]*sdsClinicChip__t138vcdl[^"]*"[^>]*>\s*([\s\S]*?)\s*<\/span>/gi;

  while ((match = clinicChipRegex.exec(html)) !== null) {
    if (match[1]) {
      clinics.push(cleanExtractedText_(match[1]));
    }
  }

  var clinicResult = uniqueJoin_(clinics);
  if (clinicResult) {
    return clinicResult;
  }

  // 2. Если клиника одна и chips не нашли — берем текущее название клиники со страницы
  var currentClinicMatch = html.match(
    /<(?:a|p)[^>]*data-testid="doctor-page__clinic-name"[^>]*>\s*([\s\S]*?)\s*<\/(?:a|p)>/i
  );

  if (currentClinicMatch && currentClinicMatch[1]) {
    return cleanExtractedText_(currentClinicMatch[1]);
  }

  // 3. Fallback: practicesAt
  var practicesAtMatch = html.match(/"practicesAt":\[(.*?)\],"alumniOf"/);
  if (practicesAtMatch && practicesAtMatch[1]) {
    var practiceNameMatch = practicesAtMatch[1].match(/"name":"([^"]+)"/);
    if (practiceNameMatch && practiceNameMatch[1]) {
      return cleanJsonText_(practiceNameMatch[1]);
    }
  }

  // 4. Fallback: servicesClinics
  var servicesClinicMatch = html.match(/"servicesClinics":\[\{"id":[^{}]*?"name":"([^"]+)"/);
  if (servicesClinicMatch && servicesClinicMatch[1]) {
    return cleanJsonText_(servicesClinicMatch[1]);
  }

  return '';
}

/**
 * Очистка текста, извлеченного из JSON.
 */
function cleanJsonText_(value) {
  return normalizeText_(
    String(value || '')
      .replace(/\\"/g, '"')
      .replace(/\\\\/g, '\\')
      .replace(/\\u003c/gi, '<')
      .replace(/\\u003e/gi, '>')
      .replace(/\\u0026quot;/gi, '"')
      .replace(/\\u0026amp;/gi, '&')
      .replace(/\\u002F/gi, '/')
  );
}

/* =========================
   Общие утилиты
   ========================= */

function fetchHtml_(url) {
  return fetchGenericHtml_(url, null).html;
}

function fetchNpHtml_(url, ctx) {
  return fetchGenericHtml_(url, ctx, 'np', 3);
}

function fetchSzHtml_(url, ctx) {
  return fetchGenericHtml_(url, ctx, 'sz', 3);
}

function fetchGenericHtml_(url, ctx, key, maxAttempts) {
  var attempts = Math.max(1, Number(maxAttempts || (CONFIG.fetchOptions && CONFIG.fetchOptions.maxAttempts) || 3));
  var last = { ok: false, statusCode: '', error: '' };
  for (var attempt = 1; attempt <= attempts; attempt++) {
    if (ctx && !hasExecutionTime_(ctx, 10000)) return { ok: false, error: 'SAFE_TIME_LIMIT_NEAR', temporary: true };
    try {
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: buildFetchHeaders_() });
      var statusCode = response.getResponseCode();
      if (statusCode >= 200 && statusCode < 400) return { ok: true, html: response.getContentText(), method: 'direct', attempt: attempt };
      last = { ok: false, statusCode: statusCode, error: 'HTTP status: ' + statusCode, permanent: isPermanentHttpStatus_(statusCode) };
      if (isPermanentHttpStatus_(statusCode) || !isRetriableHttpStatus_(statusCode) || attempt >= attempts) return last;
    } catch (error) {
      last = { ok: false, error: error && error.message ? error.message : String(error), permanent: false };
      if (attempt >= attempts) return last;
    }
    var delay = Math.min(CONFIG.fetchOptions.maxDelayMs, CONFIG.fetchOptions.baseDelayMs * Math.pow(2, attempt - 1)) + randomInt_(0, CONFIG.fetchOptions.jitterMs);
    if (ctx && !hasExecutionTime_(ctx, delay + 8000)) return { ok: false, error: 'SAFE_TIME_LIMIT_NEAR', temporary: true };
    sleepMs_(delay);
  }
  return last;
}

function fetchPdHtml_(url, ctx) {
  var pdConfig = CONFIG.ratingJobs.pd;
  var last = { ok: false, statusCode: '', error: '' };
  if (!ctx.directDisabled) {
    if (!hasExecutionTime_(ctx, 12000)) return { ok: false, error: 'SAFE_TIME_LIMIT_NEAR' };
    try {
      var direct = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: buildFetchHeaders_() });
      var directStatus = direct.getResponseCode();
      if (directStatus >= 200 && directStatus < 400) {
        ctx.consecutiveDirect403 = 0;
        return { ok: true, html: direct.getContentText(), method: 'direct', attempt: 1 };
      }
      last = { ok: false, statusCode: directStatus, error: 'HTTP status: ' + directStatus, permanent: isPermanentHttpStatus_(directStatus) };
      if (directStatus === 403) {
        ctx.consecutiveDirect403++;
        addLog_(ctx.logs, 'WARN', 'direct вернул 403, переход к резервному способу', { runId: ctx.runId, url: url, consecutiveDirect403: ctx.consecutiveDirect403 });
        if (ctx.consecutiveDirect403 >= pdConfig.consecutiveDirect403Limit) {
          ctx.directDisabled = true;
          addLog_(ctx.logs, 'WARN', 'Прямые запросы ПД отключены до конца выполнения после серии 403', { runId: ctx.runId, limit: pdConfig.consecutiveDirect403Limit });
        }
      } else {
        ctx.consecutiveDirect403 = 0;
        if (isPermanentHttpStatus_(directStatus)) return last;
      }
    } catch (error) {
      ctx.consecutiveDirect403 = 0;
      last = { ok: false, error: error && error.message ? error.message : String(error), permanent: false };
    }
  }

  var reserves = [
    { method: 'allorigins_raw', url: 'https://api.allorigins.win/raw?url=' + encodeURIComponent(url) },
    { method: 'allorigins_get', url: 'https://api.allorigins.win/get?url=' + encodeURIComponent(url), unwrap: true },
    { method: 'codetabs', url: 'https://api.codetabs.com/v1/proxy?quest=' + encodeURIComponent(url) }
  ];
  var maxReserve = Math.min(pdConfig.reserveAttempts, reserves.length);
  for (var i = 0; i < maxReserve; i++) {
    var delay = i === 0 ? 0 : (i === 1 ? randomInt_(4000, 6000) : randomInt_(8000, 12000));
    if (delay > 0) {
      if (!hasExecutionTime_(ctx, delay + 12000)) return { ok: false, error: 'SAFE_TIME_LIMIT_NEAR' };
      sleepMs_(delay);
    }
    if (!hasExecutionTime_(ctx, 12000)) return { ok: false, error: 'SAFE_TIME_LIMIT_NEAR' };
    try {
      var reserve = reserves[i];
      var response = UrlFetchApp.fetch(reserve.url, { muteHttpExceptions: true, followRedirects: true, headers: buildFetchHeaders_() });
      var statusCode = response.getResponseCode();
      if (statusCode >= 200 && statusCode < 400) {
        var text = response.getContentText();
        if (reserve.unwrap) {
          try { text = JSON.parse(text).contents || text; } catch (e) {}
        }
        addLog_(ctx.logs, 'INFO', 'данные успешно получены через cors, попытка ' + (i + 2), { runId: ctx.runId, method: reserve.method, url: url });
        return { ok: true, html: text, method: reserve.method, attempt: i + 2 };
      }
      last = { ok: false, statusCode: statusCode, error: 'HTTP status: ' + statusCode, permanent: isPermanentHttpStatus_(statusCode) };
      if (isPermanentHttpStatus_(statusCode)) return last;
    } catch (error) {
      last = { ok: false, error: error && error.message ? error.message : String(error), permanent: false };
    }
  }
  return last;
}

function buildFetchHeaders_() {
  var userAgents = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 13_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
  ];

  return {
    'User-Agent': userAgents[randomInt_(0, userAgents.length - 1)],
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'Cache-Control': 'no-cache',
    'Pragma': 'no-cache'
  };
}

function isRetriableHttpStatus_(statusCode) {
  return statusCode === 408 || statusCode === 429 || statusCode === 500 || statusCode === 502 || statusCode === 503 || statusCode === 504 || statusCode === 520 || statusCode === 522;
}

function isPermanentHttpStatus_(statusCode) {
  return statusCode === 400 || statusCode === 401 || statusCode === 404;
}

function sleepMs_(ms) {
  var safeMs = Math.max(0, Math.min(300000, Math.floor(Number(ms) || 0)));
  if (safeMs > 0) {
    Utilities.sleep(safeMs);
  }
}

function randomInt_(min, max) {
  var from = Math.floor(Number(min) || 0);
  var to = Math.floor(Number(max) || 0);
  if (to < from) {
    var temp = from;
    from = to;
    to = temp;
  }

  return Math.floor(Math.random() * (to - from + 1)) + from;
}

function formatRatingForLocale_(rating, decimalSeparator) {
  var normalized = String(rating).replace(',', '.');
  return decimalSeparator === ','
    ? normalized.replace('.', ',')
    : normalized.replace(',', '.');
}

function getDecimalSeparator_() {
  var locale = SpreadsheetApp.getActive().getSpreadsheetLocale() || 'en_US';
  var localeTag = locale.replace('_', '-');

  try {
    var sample = (1.1).toLocaleString(localeTag);
    return sample.indexOf(',') !== -1 ? ',' : '.';
  } catch (error) {
    return /^ru|^uk|^be|^de|^fr|^es|^it|^pt|^tr|^pl|^cs|^sk|^sl|^lv|^lt|^et|^fi|^sv|^nl|^da|^no|^hu|^ro|^bg|^sr|^hr/i.test(locale)
      ? ','
      : '.';
  }
}

function isValidHttpUrl_(value) {
  if (!value) {
    return false;
  }

  var str = String(value).trim();

  if (!/^https?:\/\//i.test(str)) {
    return false;
  }

  if (/\s/.test(str)) {
    return false;
  }

  return true;
}

function normalizeUrl_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value)
    .replace(/\u00A0/g, ' ')
    .trim();
}

function normalizeText_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value)
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanExtractedText_(value) {
  return decodeHtmlEntities_(normalizeText_(stripTags_(value)));
}

function stripTags_(value) {
  return String(value || '').replace(/<[^>]*>/g, ' ');
}

function decodeHtmlEntities_(text) {
  var result = String(text || '');

  result = result
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&laquo;/gi, '«')
    .replace(/&raquo;/gi, '»')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>');

  return normalizeText_(result);
}

function mapRowToWebhookObject_(headers, row) {
  var result = {};

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i] || ('column_' + (i + 1));
    var value = row[i];

    if (value instanceof Date) {
      result[header] = value.toISOString();
    } else {
      result[header] = value;
    }
  }

  return result;
}

function uniqueJoin_(items) {
  var seen = {};
  var result = [];

  for (var i = 0; i < items.length; i++) {
    var value = normalizeText_(items[i]);
    if (!value) {
      continue;
    }
    if (!seen[value]) {
      seen[value] = true;
      result.push(value);
    }
  }

  return result.join(', ');
}

/* =========================
   Работа с листами
   ========================= */

function getOrCreateTargetSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.targetSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.targetSheetName);
    sheet.getRange(1, 1, 1, CONFIG.targetHeaders.length).setValues([CONFIG.targetHeaders]);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG.targetHeaders.length).setValues([CONFIG.targetHeaders]);
    return sheet;
  }

  var existingHeaders = getHeaderRow_(sheet);
  var isEmptyHeader = existingHeaders.every(function(v) { return !normalizeText_(v); });

  if (isEmptyHeader) {
    sheet.getRange(1, 1, 1, CONFIG.targetHeaders.length).setValues([CONFIG.targetHeaders]);
    return sheet;
  }

  var headerIndexes = {};
  for (var i = 0; i < existingHeaders.length; i++) {
    var name = normalizeText_(existingHeaders[i]);
    if (name) {
      headerIndexes[name] = i + 1;
    }
  }

  var missing = CONFIG.targetHeaders.filter(function(header) {
    return !headerIndexes[header];
  });

  if (missing.length > 0) {
    throw new Error(
      'На листе "' + CONFIG.targetSheetName + '" не найдены обязательные колонки: ' + missing.join(', ')
    );
  }

  return sheet;
}

function rewriteTargetSheet_(sheet, rows) {
  var maxRows = sheet.getMaxRows();
  var headersCount = CONFIG.targetHeaders.length;

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headersCount).clearContent();
  }

  if (rows.length === 0) {
    return;
  }

  sheet.getRange(2, 1, rows.length, headersCount).setValues(rows);
}

function getHeaderRow_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

function getSheetObjects_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  var objects = [];

  for (var i = 0; i < values.length; i++) {
    var rowObject = {};
    for (var j = 0; j < headers.length; j++) {
      rowObject[normalizeText_(headers[j])] = values[i][j];
    }
    objects.push(rowObject);
  }

  return objects;
}

function findFirstHeader_(headerRow, candidates) {
  var normalizedHeaders = headerRow.map(function(item) {
    return normalizeText_(item);
  });

  for (var i = 0; i < candidates.length; i++) {
    var candidate = candidates[i];
    for (var j = 0; j < normalizedHeaders.length; j++) {
      if (normalizedHeaders[j] === candidate) {
        return candidate;
      }
    }
  }

  return '';
}

function getExistingTargetData_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      rows: [],
      byDoctorDate: {},
      latestByDoctor: {}
    };
  }

  var values = sheet.getRange(2, 1, lastRow - 1, CONFIG.targetHeaders.length).getValues();
  var rows = [];
  var byDoctorDate = {};
  var latestByDoctor = {};

  for (var i = 0; i < values.length; i++) {
    rows.push(values[i]);

    var rowObj = {};
    for (var j = 0; j < CONFIG.targetHeaders.length; j++) {
      rowObj[CONFIG.targetHeaders[j]] = values[i][j];
    }

    var doctor = normalizeText_(rowObj['Врач']);
    if (!doctor) {
      continue;
    }

    var dateKey = toDateKey_(rowObj['Дата']);
    if (dateKey) {
      byDoctorDate[doctor + '||' + dateKey] = {
        rowObj: rowObj,
        arrayIndex: i
      };
    }

    var rowTimestamp = getDateTimestamp_(rowObj['Дата']);
    if (!latestByDoctor[doctor]) {
      latestByDoctor[doctor] = {
        rowObj: rowObj,
        rowIndex: i,
        timestamp: rowTimestamp
      };
      continue;
    }

    var currentLatest = latestByDoctor[doctor];
    if (
      rowTimestamp > currentLatest.timestamp ||
      (rowTimestamp === currentLatest.timestamp && i > currentLatest.rowIndex)
    ) {
      latestByDoctor[doctor] = {
        rowObj: rowObj,
        rowIndex: i,
        timestamp: rowTimestamp
      };
    }
  }

  return {
    rows: rows,
    byDoctorDate: byDoctorDate,
    latestByDoctor: latestByDoctor
  };
}

function createEmptyTargetRowObject_() {
  var obj = {};
  for (var i = 0; i < CONFIG.targetHeaders.length; i++) {
    obj[CONFIG.targetHeaders[i]] = '';
  }
  return obj;
}

function copyTargetRow_(fromObj, toObj) {
  for (var i = 0; i < CONFIG.targetHeaders.length; i++) {
    var header = CONFIG.targetHeaders[i];
    toObj[header] = fromObj[header];
  }
}

function toDateKey_(value) {
  var date = parseDateValue_(value);
  if (!date) {
    return '';
  }

  var timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
}

function getDateTimestamp_(value) {
  var date = parseDateValue_(value);
  return date ? date.getTime() : 0;
}

function parseDateValue_(value) {
  if (value instanceof Date) {
    return new Date(value.getTime());
  }

  if (value === null || value === undefined || value === '') {
    return null;
  }

  var parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return null;
  }

  return parsed;
}

function getHeaderIndexes_(sheet, requiredHeaders) {
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var indexes = {};

  for (var i = 0; i < headers.length; i++) {
    var name = normalizeText_(headers[i]);
    if (name) {
      indexes[name] = i + 1;
    }
  }

  var missing = requiredHeaders.filter(function(header) {
    return !indexes[header];
  });

  if (missing.length > 0) {
    throw new Error('Не найдены обязательные колонки: ' + missing.join(', '));
  }

  return indexes;
}

/* =========================
   Фоновые задания рейтингов
   ========================= */

function processRatingJobBatch_(jobId) {
  var job = getRatingJob_(jobId);
  if (!job) {
    throw new Error('Задание рейтингов не найдено: ' + jobId);
  }

  var aggregatorKey = job.aggregatorKey;
  var maxRetryRounds = getMaxRetryRounds_(aggregatorKey);
  var retryQueue = job.retryQueueSnapshot || job.failedIndexes || [];
  var retryPosition = Number(job.retryQueuePosition || 0);
  var retryFailedIndexes = job.retryFailedIndexes || [];
  var startedAt = Date.now();
  var maxRuntimeMs = Number(job.batchMaxRuntimeMs || CONFIG.ratingJobs.batchMaxRuntimeMs);

  if (!retryQueue.length) {
    job.status = 'completed';
    job.failedIndexes = [];
    clearRetryRoundState_(job);
    saveRatingJob_(job);
    return job;
  }

  job.retryQueueSnapshot = retryQueue.slice();
  job.retryQueuePosition = retryPosition;
  job.retryFailedIndexes = retryFailedIndexes.slice();
  job.status = 'retrying';

  for (var i = retryPosition; i < retryQueue.length; i++) {
    if (shouldStopRatingJobBatch_(startedAt, maxRuntimeMs)) {
      savePausedRetryRound_(job, retryQueue, i, retryFailedIndexes);
      return job;
    }

    var sourceIndex = retryQueue[i];
    job.retryQueuePosition = i;

    try {
      processRatingJobItem_(job, sourceIndex);
    } catch (error) {
      retryFailedIndexes.push(sourceIndex);
      job.lastError = error && error.message ? error.message : String(error);
    }

    if (shouldStopRatingJobBatch_(startedAt, maxRuntimeMs)) {
      savePausedRetryRound_(job, retryQueue, i + 1, retryFailedIndexes);
      return job;
    }
  }

  job.failedIndexes = uniqueIndexes_(retryFailedIndexes);
  job.retryRound = Number(job.retryRound || 0) + 1;
  clearRetryRoundState_(job);

  if (job.failedIndexes.length && job.retryRound >= maxRetryRounds) {
    job.status = 'completed_with_errors';
  } else if (job.failedIndexes.length) {
    job.status = 'pending_retry';
  } else {
    job.status = 'completed';
  }

  saveRatingJob_(job);
  return job;
}

function getRatingJob_(jobId) {
  var raw = PropertiesService.getScriptProperties().getProperty(CONFIG.ratingJobs.storagePrefix + jobId);
  if (!raw) {
    return null;
  }

  var job = JSON.parse(raw);
  job.failedIndexes = job.failedIndexes || [];
  job.retryQueueSnapshot = job.retryQueueSnapshot || null;
  job.retryQueuePosition = Number(job.retryQueuePosition || 0);
  job.retryFailedIndexes = job.retryFailedIndexes || [];

  return job;
}

function saveRatingJob_(job) {
  if (!job || !job.id) {
    throw new Error('Нельзя сохранить задание рейтингов без id');
  }

  var payload = {
    id: job.id,
    status: job.status,
    aggregatorKey: job.aggregatorKey,
    retryRound: Number(job.retryRound || 0),
    failedIndexes: job.failedIndexes || [],
    retryQueueSnapshot: job.retryQueueSnapshot || null,
    retryQueuePosition: Number(job.retryQueuePosition || 0),
    retryFailedIndexes: job.retryFailedIndexes || [],
    batchMaxRuntimeMs: job.batchMaxRuntimeMs || null,
    lastError: job.lastError || ''
  };

  PropertiesService
    .getScriptProperties()
    .setProperty(CONFIG.ratingJobs.storagePrefix + job.id, JSON.stringify(payload));
}

function savePausedRetryRound_(job, retryQueue, nextPosition, retryFailedIndexes) {
  var remainingIndexes = retryQueue.slice(nextPosition);

  job.failedIndexes = uniqueIndexes_(retryFailedIndexes.concat(remainingIndexes));
  job.retryFailedIndexes = uniqueIndexes_(retryFailedIndexes);
  job.retryQueueSnapshot = retryQueue.slice();
  job.retryQueuePosition = nextPosition;
  job.status = 'pending_retry';

  saveRatingJob_(job);
}

function clearRetryRoundState_(job) {
  job.retryQueueSnapshot = null;
  job.retryQueuePosition = 0;
  job.retryFailedIndexes = [];
}

function shouldStopRatingJobBatch_(startedAt, maxRuntimeMs) {
  return Date.now() - startedAt >= maxRuntimeMs;
}

function getMaxRetryRounds_(aggregatorKey) {
  var maxRetryRounds = CONFIG.ratingJobs.maxRetryRounds || {};
  return Number(maxRetryRounds[aggregatorKey] || 0);
}

function uniqueIndexes_(indexes) {
  var seen = {};
  var result = [];

  for (var i = 0; i < indexes.length; i++) {
    var index = Number(indexes[i]);
    if (isNaN(index) || seen[index]) {
      continue;
    }

    seen[index] = true;
    result.push(index);
  }

  return result;
}

/* =========================
   Логирование
   ========================= */

function getOrCreateLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.logSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.logSheetName);
    sheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Level', 'Message', 'Details']]);
  }

  return sheet;
}

function addLog_(buffer, level, message, details) {
  buffer.push([
    new Date(),
    level,
    message,
    details ? JSON.stringify(details) : ''
  ]);
}

function flushLogs_(logSheet, buffer) {
  if (!buffer || buffer.length === 0) {
    return;
  }

  var startRow = logSheet.getLastRow() + 1;
  logSheet.getRange(startRow, 1, buffer.length, 4).setValues(buffer);
}
