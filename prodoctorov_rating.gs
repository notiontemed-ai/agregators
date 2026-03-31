/**
 * Добавляет меню TEMED при открытии таблицы.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TEMED')
    .addItem('Обновить рейтинги (ПД)', 'parseRatingsFromLinks')
    .addToUi();
}

/**
 * Основная функция: обновляет рейтинги по ссылкам на активном листе.
 */
function parseRatingsFromLinks() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var headerIndexes = getHeaderIndexes_(sheet, ['Ссылка', 'Рейтинг']);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  var rowsCount = lastRow - 1;
  var links = sheet
    .getRange(2, headerIndexes['Ссылка'], rowsCount, 1)
    .getValues();

  var decimalSeparator = getDecimalSeparator_();
  var output = [];

  for (var i = 0; i < links.length; i++) {
    var rawUrl = links[i][0];
    var url = rawUrl ? String(rawUrl).trim() : '';

    if (!url) {
      output.push(['']);
      continue;
    }

    try {
      if (!isValidHttpUrl_(url)) {
        output.push(['']);
        continue;
      }

      var html = fetchHtml_(url);
      var rating = extractSecondRating_(html);

      if (!rating) {
        output.push(['']);
        continue;
      }

      output.push([formatRatingForLocale_(rating, decimalSeparator)]);
    } catch (error) {
      output.push(['']);
    }
  }

  sheet
    .getRange(2, headerIndexes['Рейтинг'], rowsCount, 1)
    .setValues(output);
}

/**
 * Возвращает индексы нужных колонок по названиям заголовков.
 */
function getHeaderIndexes_(sheet, requiredHeaders) {
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var indexes = {};

  for (var i = 0; i < headers.length; i++) {
    var name = headers[i] ? String(headers[i]).trim() : '';
    if (name) {
      indexes[name] = i + 1;
    }
  }

  var missing = requiredHeaders.filter(function(header) {
    return !indexes[header];
  });

  if (missing.length > 0) {
    throw new Error(
      'Не найдены обязательные колонки: ' + missing.join(', ')
    );
  }

  return indexes;
}

/**
 * Загружает HTML страницы по URL.
 */
function fetchHtml_(url) {
  var response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
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

/**
 * Ищет второе вхождение слова "Рейтинг" и извлекает рейтинг из соответствующего блока.
 */
function extractSecondRating_(html) {
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

  if (!match || !match[1]) {
    return '';
  }

  return match[1];
}

/**
 * Приводит рейтинг к разделителю, принятому в таблице.
 */
function formatRatingForLocale_(rating, decimalSeparator) {
  var normalized = String(rating).replace(',', '.');
  return decimalSeparator === ','
    ? normalized.replace('.', ',')
    : normalized.replace(',', '.');
}

/**
 * Определяет десятичный разделитель по локали таблицы.
 */
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

/**
 * Проверяет, что строка является валидным HTTP(S) URL.
 */
function isValidHttpUrl_(value) {
  try {
    var url = new URL(value);
    return url.protocol === 'http:' || url.protocol === 'https:';
  } catch (error) {
    return false;
  }
}
