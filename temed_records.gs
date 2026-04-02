const REPORT_SPREADSHEET_ID = '1xjyk0eGgjDI2VLZpxDxGtJfQ26xU3SV9ULTRghjJLS8';
const CLINIC_MAPPING_SHEET = 'Соответствие клиник';
const CITY_MAPPING_SHEET = 'Соответствие городов';
const RAW_REPORT_SHEET = 'Все записи';
const COUPON_REPORT_SHEET = 'Все купоны';
const COUPON_SOURCE_SHEET = 'Клубные купоны';
const ERROR_SHEET_NAME = 'Ошибки купонов';
const MENU_NAME = 'TEMED';
const MENU_ITEM_NAME = 'Обработать записи';
const MENU_COUPON_ITEM_NAME = 'Обработать купоны';
const MENU_ANNOUNCEMENT_ITEM_NAME = 'Отправить анонс';
const ANNOUNCEMENT_WEBHOOK_URL =
  'https://n8n-x3.tech.temed.ru/webhook-test/57353eb1-2f1c-4f4c-ab0a-995c84a617cf';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem(MENU_ITEM_NAME, 'processTemedRecords')
    .addItem(MENU_COUPON_ITEM_NAME, 'processTemedCoupons')
    .addItem(MENU_ANNOUNCEMENT_ITEM_NAME, 'sendTemedAnnouncement')
    .addToUi();
}

function processTemedRecords() {
  const ui = SpreadsheetApp.getUi();
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reportSpreadsheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);

  const clinicMap = loadClinicMapping_(reportSpreadsheet);
  const cityMap = loadCityMapping_(reportSpreadsheet);
  const sourceSheets = sourceSpreadsheet
    .getSheets()
    .filter((sheet) => /^(Онлайн-запись|Экспресс-запись)/i.test(sheet.getName()));

  if (sourceSheets.length === 0) {
    ui.alert('Листы для обработки не найдены.');
    return;
  }

  const allRows = [];
  sourceSheets.forEach((sheet) => {
    allRows.push(...parseSourceSheet_(sheet, clinicMap, cityMap));
  });

  if (allRows.length === 0) {
    ui.alert('В подходящих листах нет строк с данными.');
    return;
  }

  const incomingDateKeys = collectDateKeys_(allRows);
  if (hasExistingDates_(reportSpreadsheet, incomingDateKeys)) {
    const button = ui.alert(
      'Найдены существующие записи',
      'На листе "Все записи" уже есть данные за даты из текущей загрузки. Перезаписать их?',
      ui.ButtonSet.YES_NO
    );

    if (button !== ui.Button.YES) {
      ui.alert('Обработка отменена пользователем. Данные не изменены.');
      return;
    }

    deleteRowsByDates_(reportSpreadsheet, incomingDateKeys);
  }

  appendRawRows_(reportSpreadsheet, allRows);

  sourceSheets.forEach((sheet) => sourceSpreadsheet.deleteSheet(sheet));

  ui.alert(
    `Обработка завершена. Перенесено записей: ${allRows.length}. Удалено листов: ${sourceSheets.length}.`
  );
}

function processTemedCoupons() {
  const ui = SpreadsheetApp.getUi();
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reportSpreadsheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
  const sourceSheet = sourceSpreadsheet.getSheetByName(COUPON_SOURCE_SHEET);

  if (!sourceSheet) {
    ui.alert(`Лист "${COUPON_SOURCE_SHEET}" не найден.`);
    return;
  }

  const clinicMap = loadClinicMapping_(reportSpreadsheet, 'Заголовок купоны');
  const parseResult = parseCouponSheet_(sourceSheet, clinicMap);

  if (parseResult.errors.length > 0) {
    writeCouponErrors_(sourceSpreadsheet, parseResult.errors);
    ui.alert(
      'Обработка купонов остановлена из-за ошибок. Подробности записаны на лист "Ошибки купонов".'
    );
    return;
  }

  if (parseResult.rows.length === 0) {
    ui.alert('На листе "Клубные купоны" не найдено строк для переноса.');
    return;
  }

  const incomingDateKeys = collectDateKeysByField_(parseResult.rows, 'recordDate');
  if (hasExistingCouponDates_(reportSpreadsheet, incomingDateKeys)) {
    const button = ui.alert(
      'Найдены существующие купоны',
      'На листе "Все купоны" уже есть данные за даты записи из текущей загрузки. Перезаписать их?',
      ui.ButtonSet.YES_NO
    );

    if (button !== ui.Button.YES) {
      ui.alert('Обработка отменена пользователем. Данные не изменены.');
      return;
    }

    deleteCouponRowsByRecordDates_(reportSpreadsheet, incomingDateKeys);
  }

  appendCouponRows_(reportSpreadsheet, parseResult.rows);
  clearCouponErrors_(sourceSpreadsheet);
  sourceSpreadsheet.deleteSheet(sourceSheet);

  ui.alert(`Обработка купонов завершена. Перенесено строк: ${parseResult.rows.length}.`);
}

function sendTemedAnnouncement() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ratingSheet = spreadsheet.getSheetByName('Рейтинг');

  if (!ratingSheet) {
    ui.alert('Лист "Рейтинг" не найден.');
    return;
  }

  const values = ratingSheet.getDataRange().getValues();
  if (values.length < 2) {
    ui.alert('На листе "Рейтинг" нет данных для отправки.');
    return;
  }

  const headers = values[0].map((value) => String(value || '').trim());
  const dateColumnIndex = headers.indexOf('Дата');
  if (dateColumnIndex === -1) {
    ui.alert('На листе "Рейтинг" не найден столбец "Дата".');
    return;
  }

  const today = new Date();
  const todayKey = toDateKeyLocal_(today);

  const rows = values.slice(1);
  let latestDate = null;

  rows.forEach((row) => {
    const rowDate = toDate_(row[dateColumnIndex]);
    if (!rowDate) {
      return;
    }

    if (!latestDate || rowDate.getTime() > latestDate.getTime()) {
      latestDate = rowDate;
    }
  });

  if (!latestDate) {
    ui.alert('На листе "Рейтинг" нет валидных дат в столбце "Дата".');
    return;
  }

  const latestDateKey = toDateKeyLocal_(latestDate);
  if (latestDateKey !== todayKey) {
    const userChoice = ui.alert(
      'Внимание: дата не совпадает',
      `Дата запуска: ${todayKey}. Последняя дата на листе "Рейтинг": ${latestDateKey}. Продолжить отправку?`,
      ui.ButtonSet.YES_NO
    );

    if (userChoice !== ui.Button.YES) {
      ui.alert('Отправка отменена пользователем.');
      return;
    }
  }

  const weekStart = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 6);
  const weekRows = rows
    .map((row) => {
      const rowDate = toDate_(row[dateColumnIndex]);
      return { row, rowDate };
    })
    .filter((item) => item.rowDate && item.rowDate.getTime() >= weekStart.getTime() && item.rowDate.getTime() <= today.getTime())
    .map((item) => mapRowToObject_(headers, item.row));

  if (weekRows.length === 0) {
    ui.alert('За последнюю неделю нет данных для отправки.');
    return;
  }

  const payload = {
    reportName: 'Рейтинг',
    generatedAt: new Date().toISOString(),
    period: {
      from: toDateKeyLocal_(weekStart),
      to: todayKey
    },
    latestSheetDate: latestDateKey,
    rows: weekRows
  };

  const response = UrlFetchApp.fetch(ANNOUNCEMENT_WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    ui.alert(`Ошибка отправки: HTTP ${status}. Ответ: ${response.getContentText()}`);
    return;
  }

  ui.alert(`Анонс успешно отправлен. Передано строк: ${weekRows.length}.`);
}

function mapRowToObject_(headers, row) {
  const obj = {};
  for (let i = 0; i < headers.length; i += 1) {
    const key = headers[i] || `column_${i + 1}`;
    const value = row[i];
    obj[key] = value instanceof Date ? value.toISOString() : value;
  }
  return obj;
}

function toDateKeyLocal_(value) {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    return '';
  }

  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function loadClinicMapping_(reportSpreadsheet, titleColumnName) {
  const sheet = reportSpreadsheet.getSheetByName(CLINIC_MAPPING_SHEET);
  if (!sheet) {
    throw new Error(`Не найден лист "${CLINIC_MAPPING_SHEET}" в книге отчетов.`);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return {};
  }

  const header = values[0].map((value) => String(value || '').trim());
  const clinicIdx = header.indexOf('Клиника');
  const titleIdx = header.indexOf(titleColumnName || 'Заголовок в актах');

  if (clinicIdx === -1 || titleIdx === -1) {
    throw new Error(
      `На листе "${CLINIC_MAPPING_SHEET}" должны быть столбцы "Клиника" и "${titleColumnName ||
        'Заголовок в актах'}".`
    );
  }

  const map = {};
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const clinic = String(row[clinicIdx] || '').trim();
    const title = String(row[titleIdx] || '').trim();
    if (clinic && title) {
      map[title] = clinic;
    }
  }

  return map;
}

function parseCouponSheet_(sheet, clinicMap) {
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return { rows: [], errors: [] };
  }

  const headerRowIndex = findCouponHeaderRowIndex_(values);
  if (headerRowIndex === -1) {
    throw new Error(
      'На листе "Клубные купоны" не найдена строка заголовков таблицы с колонками купонов.'
    );
  }

  const headerRow = values[headerRowIndex].map((cell) => normalizeHeader_(cell));
  const phoneIdx = headerRow.indexOf('телефон пациента');
  const recordDateIdx = headerRow.indexOf('дата записи');
  const appointmentDateIdx = headerRow.indexOf('дата приема');
  const infoIdx = headerRow.indexOf('информация');
  const statusIdx = headerRow.indexOf('статус');
  const priceIdx = headerRow.findIndex((name) => /^стоимость купона/.test(name));

  if ([phoneIdx, recordDateIdx, appointmentDateIdx, infoIdx, statusIdx, priceIdx].some((idx) => idx === -1)) {
    throw new Error(
      'На листе "Клубные купоны" не найдены обязательные столбцы: Телефон пациента, Дата записи, Дата приема, Информация, Статус, Стоимость купона.'
    );
  }

  const rows = [];
  const errors = [];
  let currentClinic = '';

  for (let i = headerRowIndex + 1; i < values.length; i += 1) {
    const row = values[i];
    const rowNum = i + 1;
    const firstCellRaw = row[0];
    const firstCell = String(firstCellRaw || '').trim();
    const isDataRow = isCouponDataRow_(firstCellRaw);
    const recordDate = toDate_(row[recordDateIdx]);
    const appointmentDate = toDate_(row[appointmentDateIdx]);
    const info = String(row[infoIdx] || '').trim();
    const phone = String(row[phoneIdx] || '').trim();
    const status = String(row[statusIdx] || '').trim();
    const price = toNumber_(row[priceIdx]);

    if (!isDataRow) {
      if (!recordDate && !appointmentDate && !info && !phone && !status && price === 0 && !firstCell) {
        continue;
      }

      if (/^сумма /i.test(firstCell)) {
        continue;
      }

      if (!recordDate && !appointmentDate && !info && !phone && !status && price === 0) {
        currentClinic = clinicMap[firstCell] || '';
        if (!currentClinic) {
          errors.push(
            `Строка ${rowNum}: не найдено соответствие клиники для заголовка "${firstCell}" (колонка "Заголовок купоны").`
          );
        }
      }
      continue;
    }

    if (!recordDate && !appointmentDate && !info && !phone && !status && price === 0 && !firstCell) {
      continue;
    }

    if (!recordDate && !appointmentDate && !info && !phone && !status && price === 0) {
      if (/^сумма /i.test(firstCell)) {
        continue;
      }

      currentClinic = clinicMap[firstCell] || '';
      if (!currentClinic) {
        errors.push(
          `Строка ${rowNum}: не найдено соответствие клиники для заголовка "${firstCell}" (колонка "Заголовок купоны").`
        );
      }
      continue;
    }

    if (!recordDate && !appointmentDate && !info && !phone && !status && price !== 0) {
      continue;
    }

    if (!currentClinic) {
      errors.push(`Строка ${rowNum}: не определена клиника для строки купона.`);
      continue;
    }

    if (!recordDate) {
      errors.push(`Строка ${rowNum}: не распознана дата записи.`);
      continue;
    }

    if (!appointmentDate) {
      errors.push(`Строка ${rowNum}: не распознана дата приема.`);
      continue;
    }

    const doctor = extractCouponDoctor_(info);
    if (!doctor) {
      errors.push(`Строка ${rowNum}: не удалось выделить врача из поля "Информация" ("${info}").`);
      continue;
    }

    rows.push({
      clinic: currentClinic,
      recordDate,
      appointmentDate,
      doctor,
      status,
      price,
      phone,
    });
  }

  return { rows, errors };
}

function isCouponDataRow_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value);
  }

  const text = String(value || '').trim();
  if (!text) {
    return false;
  }

  return /^\d+$/.test(text);
}

function findCouponHeaderRowIndex_(values) {
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i].map((cell) => normalizeHeader_(cell));
    if (
      row.includes('телефон пациента') &&
      row.includes('дата записи') &&
      row.includes('дата приема') &&
      row.includes('информация') &&
      row.includes('статус') &&
      row.some((name) => /^стоимость купона/.test(name))
    ) {
      return i;
    }
  }

  return -1;
}

function extractCouponDoctor_(info) {
  const text = String(info || '').trim();
  if (!text) {
    return '';
  }

  const withoutSpec = text.replace(/\s*\([^)]*\)\s*$/, '').trim();
  const parts = withoutSpec.split(/\s+/).filter(Boolean);
  if (parts.length < 2) {
    return '';
  }

  const lastName = parts[0];
  const firstInitial = parts[1] ? `${parts[1].charAt(0).toUpperCase()}.` : '';
  const middleInitial = parts[2] ? `${parts[2].charAt(0).toUpperCase()}.` : '';

  if (!firstInitial) {
    return '';
  }

  return [lastName, firstInitial, middleInitial].filter(Boolean).join(' ');
}

function writeCouponErrors_(spreadsheet, errors) {
  let sheet = spreadsheet.getSheetByName(ERROR_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(ERROR_SHEET_NAME);
  } else {
    sheet.clearContents();
  }

  const rows = [['Ошибка'], ...errors.map((errorText) => [errorText])];
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
  sheet.autoResizeColumn(1);
}

function clearCouponErrors_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(ERROR_SHEET_NAME);
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
}

function appendCouponRows_(reportSpreadsheet, rows) {
  const sheet = reportSpreadsheet.getSheetByName(COUPON_REPORT_SHEET);
  if (!sheet) {
    throw new Error(`Не найден лист "${COUPON_REPORT_SHEET}" в книге отчетов.`);
  }

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map((cell) => normalizeHeader_(cell));
  const rowValues = rows.map((row) =>
    header.map((name) => {
      if (name === 'дата записи') return row.recordDate;
      if (name === 'дата приема') return row.appointmentDate;
      if (name === 'врач') return row.doctor;
      if (name === 'статус') return row.status;
      if (/^стоимость купона/.test(name)) return row.price;
      if (name === 'телефон пациента') return row.phone;
      if (name === 'клиника') return row.clinic;
      return '';
    })
  );

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rowValues.length, rowValues[0].length).setValues(rowValues);
}

function collectDateKeysByField_(rows, fieldName) {
  const keys = {};
  rows.forEach((row) => {
    if (row[fieldName]) {
      keys[formatDateKey_(row[fieldName])] = true;
    }
  });
  return keys;
}

function hasExistingCouponDates_(reportSpreadsheet, dateKeys) {
  const sheet = reportSpreadsheet.getSheetByName(COUPON_REPORT_SHEET);
  if (!sheet) {
    throw new Error(`Не найден лист "${COUPON_REPORT_SHEET}" в книге отчетов.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return false;
  }

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map((cell) => normalizeHeader_(cell));
  const recordDateIdx = header.indexOf('дата записи');
  if (recordDateIdx === -1) {
    throw new Error(`На листе "${COUPON_REPORT_SHEET}" не найден столбец "Дата записи".`);
  }

  const dates = sheet.getRange(2, recordDateIdx + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < dates.length; i += 1) {
    const existingDate = toDate_(dates[i][0]);
    if (existingDate && dateKeys[formatDateKey_(existingDate)]) {
      return true;
    }
  }

  return false;
}

function deleteCouponRowsByRecordDates_(reportSpreadsheet, dateKeys) {
  const sheet = reportSpreadsheet.getSheetByName(COUPON_REPORT_SHEET);
  if (!sheet) {
    throw new Error(`Не найден лист "${COUPON_REPORT_SHEET}" в книге отчетов.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map((cell) => normalizeHeader_(cell));
  const recordDateIdx = header.indexOf('дата записи');
  if (recordDateIdx === -1) {
    throw new Error(`На листе "${COUPON_REPORT_SHEET}" не найден столбец "Дата записи".`);
  }

  const dates = sheet.getRange(2, recordDateIdx + 1, lastRow - 1, 1).getValues();
  for (let i = dates.length - 1; i >= 0; i -= 1) {
    const existingDate = toDate_(dates[i][0]);
    if (existingDate && dateKeys[formatDateKey_(existingDate)]) {
      sheet.deleteRow(i + 2);
    }
  }
}

function loadCityMapping_(reportSpreadsheet) {
  const sheet = reportSpreadsheet.getSheetByName(CITY_MAPPING_SHEET);
  if (!sheet) {
    throw new Error(`Не найден лист "${CITY_MAPPING_SHEET}" в книге отчетов.`);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return {};
  }

  const header = values[0].map((value) => String(value || '').trim());
  const cityIdx = header.indexOf('Город');
  const titleIdx = header.indexOf('Заголовок в актах');

  if (cityIdx === -1 || titleIdx === -1) {
    throw new Error(
      `На листе "${CITY_MAPPING_SHEET}" должны быть столбцы "Город" и "Заголовок в актах".`
    );
  }

  const map = {};
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const city = String(row[cityIdx] || '').trim();
    const title = String(row[titleIdx] || '').trim();

    if (!city || !title) {
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(map, title)) {
      throw new Error(
        `На листе "${CITY_MAPPING_SHEET}" найден дубликат заголовка "${title}" (строка ${i + 1}).`
      );
    }

    map[title] = city;
  }

  return map;
}

function parseSourceSheet_(sheet, clinicMap, cityMap) {
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return [];
  }

  const sheetName = sheet.getName();
  const isExpress = /^Экспресс-запись/i.test(sheetName);
  const title = String(values[0][0] || '').trim();

  let clinic = '';
  let city = '';
  if (isExpress) {
    city = cityMap[title];
    if (!city) {
      throw new Error(
        `Для листа "${sheetName}" не найден город: в "${CITY_MAPPING_SHEET}" отсутствует "Заголовок в актах" = "${title}".`
      );
    }
  } else {
    clinic = clinicMap[title] || title;
  }

  const headerRowIndex = findTableHeaderRowIndex_(values);

  if (headerRowIndex === -1) {
    return [];
  }

  const headerRow = values[headerRowIndex].map((cell) => normalizeHeader_(cell));
  const dateIdx = headerRow.indexOf('дата');
  const infoIdx = headerRow.indexOf('информация');
  const phoneIdx = headerRow.indexOf('телефон пациента');
  const priceIdx = headerRow.indexOf('цена');

  if ([dateIdx, infoIdx, phoneIdx, priceIdx].some((idx) => idx === -1)) {
    throw new Error(
      `На листе "${sheet.getName()}" не найдены обязательные столбцы: Дата, Информация, Телефон пациента, Цена.`
    );
  }

  const result = [];

  for (let i = headerRowIndex + 1; i < values.length; i += 1) {
    const row = values[i];
    const dateValue = row[dateIdx];
    const info = String(row[infoIdx] || '').trim();
    const phone = String(row[phoneIdx] || '').trim();
    const price = toNumber_(row[priceIdx]);

    if (!dateValue && !info && !phone && price === 0) {
      continue;
    }

    const date = toDate_(dateValue);
    if (!date) {
      continue;
    }

    const doctor = extractDoctor_(info);
    result.push({
      date,
      doctor,
      clinic,
      city,
      phone,
      price,
    });
  }

  return result;
}

function findTableHeaderRowIndex_(values) {
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i].map((cell) => normalizeHeader_(cell));
    if (
      row.includes('дата') &&
      row.includes('информация') &&
      row.includes('телефон пациента') &&
      row.includes('цена')
    ) {
      return i;
    }
  }

  return -1;
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function extractDoctor_(info) {
  const text = String(info || '').trim();
  const withoutPrefix = text.replace(/^(Онлайн-запись|Экспресс-запись)\s+/i, '').trim();
  const lastDotIndex = withoutPrefix.lastIndexOf('.');

  if (lastDotIndex <= 0) {
    return withoutPrefix;
  }

  return withoutPrefix.slice(0, lastDotIndex + 1).trim();
}

function toDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (!value) {
    return null;
  }

  const text = String(value).trim();
  const ruDateMatch = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (ruDateMatch) {
    const day = Number(ruDateMatch[1]);
    const month = Number(ruDateMatch[2]);
    const year = Number(ruDateMatch[3]);
    const parsedRu = new Date(year, month - 1, day);

    if (
      parsedRu.getFullYear() === year &&
      parsedRu.getMonth() === month - 1 &&
      parsedRu.getDate() === day
    ) {
      return parsedRu;
    }

    return null;
  }

  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function toNumber_(value) {
  if (typeof value === 'number') {
    return value;
  }

  const normalized = String(value || '')
    .replace(/\s/g, '')
    .replace(',', '.');

  const num = Number(normalized);
  return Number.isFinite(num) ? num : 0;
}

function appendRawRows_(reportSpreadsheet, rows) {
  const sheet = reportSpreadsheet.getSheetByName(RAW_REPORT_SHEET);
  if (!sheet) {
    throw new Error(`Не найден лист "${RAW_REPORT_SHEET}" в книге отчетов.`);
  }

  const values = rows.map((row) => [row.date, row.doctor, row.clinic, row.city, row.phone, row.price]);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);

  const totalRows = sheet.getLastRow();
  if (totalRows > 1) {
    sheet.getRange(2, 1, totalRows - 1, 6).sort({ column: 1, ascending: true });
    sheet.getRange(2, 1, totalRows - 1, 1).setNumberFormat('dd.mm.yyyy');
  }
}

function collectDateKeys_(rows) {
  const keys = {};
  rows.forEach((row) => {
    keys[formatDateKey_(row.date)] = true;
  });
  return keys;
}

function hasExistingDates_(reportSpreadsheet, dateKeys) {
  const sheet = reportSpreadsheet.getSheetByName(RAW_REPORT_SHEET);
  if (!sheet) {
    throw new Error(`Не найден лист "${RAW_REPORT_SHEET}" в книге отчетов.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return false;
  }

  const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < dates.length; i += 1) {
    const existingDate = toDate_(dates[i][0]);
    if (!existingDate) {
      continue;
    }

    if (dateKeys[formatDateKey_(existingDate)]) {
      return true;
    }
  }

  return false;
}

function deleteRowsByDates_(reportSpreadsheet, dateKeys) {
  const sheet = reportSpreadsheet.getSheetByName(RAW_REPORT_SHEET);
  if (!sheet) {
    throw new Error(`Не найден лист "${RAW_REPORT_SHEET}" в книге отчетов.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = dates.length - 1; i >= 0; i -= 1) {
    const existingDate = toDate_(dates[i][0]);
    if (!existingDate) {
      continue;
    }

    if (dateKeys[formatDateKey_(existingDate)]) {
      sheet.deleteRow(i + 2);
    }
  }
}

function formatDateKey_(date) {
  return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
}
