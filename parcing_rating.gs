const CONFIG = {
  menuName: 'TEMED',
  sourceSheetName: 'Врачи',
  targetSheetName: 'Рейтинг',
  logSheetName: 'Log',

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
  }
};

/**
 * Меню при открытии.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.menuName)
    .addItem('Обновить ПД', 'updatePdRatings')
    .addItem('Обновить НП', 'updateNpRatings')
    .addItem('Обновить СЗ', 'updateSzRatings')
    .addSeparator()
    .addItem('Обновить все рейтинги', 'updateAllRatings')
    .addToUi();
}

function updatePdRatings() {
  updateRatings_(['pd']);
}

function updateNpRatings() {
  updateRatings_(['np']);
}

function updateSzRatings() {
  updateRatings_(['sz']);
}

function updateAllRatings() {
  updateRatings_(['pd', 'np', 'sz']);
}

/**
 * Основной обработчик.
 */
function updateRatings_(aggregatorKeys) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = getOrCreateLogSheet_();
  var logs = [];
  var now = new Date();

  addLog_(logs, 'INFO', 'Старт обновления рейтингов', {
    sourceSheet: CONFIG.sourceSheetName,
    targetSheet: CONFIG.targetSheetName,
    aggregators: aggregatorKeys
  });

  try {
    var sourceSheet = ss.getSheetByName(CONFIG.sourceSheetName);
    if (!sourceSheet) {
      throw new Error('Не найден лист "' + CONFIG.sourceSheetName + '"');
    }

    var targetSheet = getOrCreateTargetSheet_();
    var decimalSeparator = getDecimalSeparator_();

    var sourceObjects = getSheetObjects_(sourceSheet);
    var sourceHeaderRow = getHeaderRow_(sourceSheet);
    var doctorHeader = findFirstHeader_(sourceHeaderRow, CONFIG.sourceDoctorHeaders);

    if (!doctorHeader) {
      throw new Error(
        'На листе "' + CONFIG.sourceSheetName + '" не найдена колонка с именем врача. Ожидается одна из: ' +
        CONFIG.sourceDoctorHeaders.join(', ')
      );
    }

    var existingTargetByDoctor = getExistingTargetRowsByDoctor_(targetSheet);
    var outputRows = [];

    for (var i = 0; i < sourceObjects.length; i++) {
      var rowNumber = i + 2;
      var sourceRow = sourceObjects[i];
      var doctorName = normalizeText_(sourceRow[doctorHeader]);

      if (!doctorName) {
        addLog_(logs, 'WARN', 'Строка пропущена: не указано имя врача', {
          row: rowNumber
        });
        continue;
      }

      var targetRow = createEmptyTargetRowObject_();

      if (existingTargetByDoctor[doctorName]) {
        copyTargetRow_(existingTargetByDoctor[doctorName], targetRow);
      }

      targetRow['Дата'] = now;
      targetRow['Врач'] = doctorName;

      for (var j = 0; j < aggregatorKeys.length; j++) {
        var key = aggregatorKeys[j];
        var agg = CONFIG.aggregators[key];
        var rawUrl = sourceRow[agg.sourceHeader];
        var url = normalizeUrl_(rawUrl);

        if (!url) {
          targetRow[agg.ratingHeader] = '';
          targetRow[agg.reviewsHeader] = '';
          targetRow[agg.clinicsHeader] = '';

          addLog_(logs, 'INFO', 'Пустая ссылка', {
            row: rowNumber,
            doctor: doctorName,
            aggregator: agg.title
          });
          continue;
        }

        if (!isValidHttpUrl_(url)) {
          targetRow[agg.ratingHeader] = '';
          targetRow[agg.reviewsHeader] = '';
          targetRow[agg.clinicsHeader] = '';

          addLog_(logs, 'WARN', 'Невалидный URL', {
            row: rowNumber,
            doctor: doctorName,
            aggregator: agg.title,
            url: url
          });
          continue;
        }

        try {
          var html = fetchHtml_(url);
          var parsed = parseAggregatorData_(key, html);

          targetRow[agg.ratingHeader] = parsed.rating
            ? formatRatingForLocale_(parsed.rating, decimalSeparator)
            : '';
          targetRow[agg.reviewsHeader] = parsed.reviews || '';
          targetRow[agg.clinicsHeader] = parsed.clinics || '';

          addLog_(logs, 'INFO', 'Данные извлечены', {
            row: rowNumber,
            doctor: doctorName,
            aggregator: agg.title,
            url: url,
            rating: parsed.rating || '',
            reviews: parsed.reviews || '',
            clinics: parsed.clinics || ''
          });
        } catch (error) {
          targetRow[agg.ratingHeader] = '';
          targetRow[agg.reviewsHeader] = '';
          targetRow[agg.clinicsHeader] = '';

          addLog_(logs, 'ERROR', 'Ошибка обработки ссылки', {
            row: rowNumber,
            doctor: doctorName,
            aggregator: agg.title,
            url: url,
            error: error && error.message ? error.message : String(error)
          });
        }
      }

      outputRows.push(
        CONFIG.targetHeaders.map(function(header) {
          return targetRow[header];
        })
      );
    }

    rewriteTargetSheet_(targetSheet, outputRows);

    addLog_(logs, 'INFO', 'Обновление завершено', {
      rowsWritten: outputRows.length,
      aggregators: aggregatorKeys
    });
  } catch (error) {
    addLog_(logs, 'ERROR', 'Критическая ошибка выполнения', {
      error: error && error.message ? error.message : String(error)
    });
    throw error;
  } finally {
    flushLogs_(logSheet, logs);
  }
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
 * Клиники ПД: name из блоков lpu.
 */
function extractPdClinics_(html) {
  if (!html) {
    return '';
  }

  var matches = [];
  var regex = /lpu\s*:\s*\{[\s\S]*?name\s*:\s*'([^']+)'/gi;
  var match;

  while ((match = regex.exec(html)) !== null) {
    if (match[1]) {
      matches.push(cleanExtractedText_(match[1]));
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
 * Клиники СЗ: ищем после блока "Выбор клиники" только названия клиник внутри label.
 */
function extractSzClinics_(html) {
  if (!html) {
    return '';
  }

  var startIdx = html.indexOf('aria-label="Выбор клиники"');
  if (startIdx === -1) {
    return '';
  }

  var tail = html.slice(startIdx);
  var clinics = [];
  var labelRegex = /<label\b[\s\S]*?<\/label>/gi;
  var labelMatch;
  var scanned = 0;
  var maxLabelsToScan = 30;

  while ((labelMatch = labelRegex.exec(tail)) !== null && scanned < maxLabelsToScan) {
    scanned++;
    var labelHtml = labelMatch[0];

    var clinicMatch = labelHtml.match(
      /<span[^>]*class="[^"]*sdsClinicChip__t138vcdl[^"]*"[^>]*>\s*([\s\S]*?)\s*<\/span>/i
    );

    if (clinicMatch && clinicMatch[1]) {
      clinics.push(cleanExtractedText_(clinicMatch[1]));
    }

    if (clinics.length > 0 && /<\/div>/.test(labelHtml)) {
      // просто продолжаем, ограничение maxLabelsToScan защищает от ухода слишком далеко
    }
  }

  return uniqueJoin_(clinics);
}

/* =========================
   Общие утилиты
   ========================= */

function fetchHtml_(url) {
  var response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript/1.0)'
    }
  });

  var statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 400) {
    throw new Error('HTTP status: ' + statusCode);
  }

  return response.getContentText();
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

function getExistingTargetRowsByDoctor_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  var headerIndexes = getHeaderIndexes_(sheet, CONFIG.targetHeaders);
  var values = sheet.getRange(2, 1, lastRow - 1, CONFIG.targetHeaders.length).getValues();
  var result = {};

  for (var i = 0; i < values.length; i++) {
    var rowObj = {};
    for (var j = 0; j < CONFIG.targetHeaders.length; j++) {
      rowObj[CONFIG.targetHeaders[j]] = values[i][j];
    }

    var doctor = normalizeText_(rowObj['Врач']);
    if (doctor) {
      result[doctor] = rowObj;
    }
  }

  return result;
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
