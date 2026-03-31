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
  var logSheet = getOrCreateLogSheet_();
  var logs = [];
  addLog_(logs, 'INFO', 'Старт обновления рейтингов', {
    sheet: sheet.getName()
  });
  try {
    var headerIndexes = getHeaderIndexes_(sheet, ['Ссылка', 'Рейтинг']);
    addLog_(logs, 'INFO', 'Колонки найдены', {
      linkColumn: headerIndexes['Ссылка'],
      ratingColumn: headerIndexes['Рейтинг']
    });

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      addLog_(logs, 'INFO', 'Нет строк для обработки (только заголовки)');
      return;
    }

    var rowsCount = lastRow - 1;
    var links = sheet
      .getRange(2, headerIndexes['Ссылка'], rowsCount, 1)
      .getValues();

    var decimalSeparator = getDecimalSeparator_();
    addLog_(logs, 'INFO', 'Определен десятичный разделитель', {
      decimalSeparator: decimalSeparator
    });
    var output = [];

    for (var i = 0; i < links.length; i++) {
      var rowNumber = i + 2;
      var rawUrl = links[i][0];
      var url = rawUrl ? String(rawUrl).trim() : '';

      if (!url) {
        output.push(['']);
        addLog_(logs, 'INFO', 'Пустая ссылка, строка пропущена', {
          row: rowNumber
        });
        continue;
      }

      try {
        if (!isValidHttpUrl_(url)) {
          output.push(['']);
          addLog_(logs, 'WARN', 'Невалидный URL', {
            row: rowNumber,
            url: url
          });
          continue;
        }

        var html = fetchHtml_(url);
        var rating = extractSecondRating_(html);

        if (!rating) {
          output.push(['']);
          addLog_(logs, 'WARN', 'Рейтинг не найден во втором блоке', {
            row: rowNumber,
            url: url
          });
          continue;
        }

        var formattedRating = formatRatingForLocale_(rating, decimalSeparator);
        output.push([formattedRating]);
        addLog_(logs, 'INFO', 'Рейтинг успешно извлечен', {
          row: rowNumber,
          url: url,
          rating: formattedRating
        });
      } catch (error) {
        output.push(['']);
        addLog_(logs, 'ERROR', 'Ошибка обработки строки', {
          row: rowNumber,
          url: url,
          error: error && error.message ? error.message : String(error)
        });
      }
    }

    sheet
      .getRange(2, headerIndexes['Рейтинг'], rowsCount, 1)
      .setValues(output);

    addLog_(logs, 'INFO', 'Обновление завершено', {
      processedRows: rowsCount
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

/**
 * Возвращает лист Log или создает его при отсутствии.
 */
function getOrCreateLogSheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Log');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('Log');
    sheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Level', 'Message', 'Details']]);
  }

  return sheet;
}

/**
 * Добавляет запись в буфер логов.
 */
function addLog_(buffer, level, message, details) {
  buffer.push([
    new Date(),
    level,
    message,
    details ? JSON.stringify(details) : ''
  ]);
}

/**
 * Пакетно записывает логи на лист.
 */
function flushLogs_(logSheet, buffer) {
  if (!buffer || buffer.length === 0) {
    return;
  }

  var startRow = logSheet.getLastRow() + 1;
  logSheet.getRange(startRow, 1, buffer.length, 4).setValues(buffer);
}
