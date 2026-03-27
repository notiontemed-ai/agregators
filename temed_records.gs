const REPORT_SPREADSHEET_ID = '1xjyk0eGgjDI2VLZpxDxGtJfQ26xU3SV9ULTRghjJLS8';
const CLINIC_MAPPING_SHEET = 'Соответствие клиник';
const WEEKLY_REPORT_SHEET = 'Онлайн записи (недели)';
const MONTHLY_REPORT_SHEET = 'Онлайн записи (мес)';
const RAW_REPORT_SHEET = 'Все онлайн записи';
const MENU_NAME = 'TEMED';
const MENU_ITEM_NAME = 'Обработать записи';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem(MENU_ITEM_NAME, 'processTemedRecords')
    .addToUi();
}

function processTemedRecords() {
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reportSpreadsheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);

  const clinicMap = loadClinicMapping_(reportSpreadsheet);
  const sourceSheets = sourceSpreadsheet
    .getSheets()
    .filter((sheet) => /^(Онлайн-запись|Экспресс-запись)/i.test(sheet.getName()));

  if (sourceSheets.length === 0) {
    SpreadsheetApp.getUi().alert('Листы для обработки не найдены.');
    return;
  }

  const allRows = [];
  sourceSheets.forEach((sheet) => {
    allRows.push(...parseSourceSheet_(sheet, clinicMap));
  });

  if (allRows.length === 0) {
    SpreadsheetApp.getUi().alert('В подходящих листах нет строк с данными.');
    return;
  }

  appendRawRows_(reportSpreadsheet, allRows);
  upsertAggregate_(reportSpreadsheet, WEEKLY_REPORT_SHEET, buildWeeklyRows_(allRows));
  upsertAggregate_(reportSpreadsheet, MONTHLY_REPORT_SHEET, buildMonthlyRows_(allRows));

  sourceSheets.forEach((sheet) => sourceSpreadsheet.deleteSheet(sheet));

  SpreadsheetApp.getUi().alert(
    `Обработка завершена. Перенесено записей: ${allRows.length}. Удалено листов: ${sourceSheets.length}.`
  );
}

function loadClinicMapping_(reportSpreadsheet) {
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
  const titleIdx = header.indexOf('Заголовок в актах');

  if (clinicIdx === -1 || titleIdx === -1) {
    throw new Error(
      `На листе "${CLINIC_MAPPING_SHEET}" должны быть столбцы "Клиника" и "Заголовок в актах".`
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

function parseSourceSheet_(sheet, clinicMap) {
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return [];
  }

  const title = String(values[0][0] || '').trim();
  const clinic = clinicMap[title] || title;
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

  const parsed = new Date(value);
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

  const values = rows.map((row) => [row.date, row.doctor, row.clinic, row.phone, row.price]);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
}

function buildWeeklyRows_(rows) {
  const map = {};

  rows.forEach((row) => {
    const period = getWeekCode_(row.date);
    const key = `${period}|${row.doctor}`;

    if (!map[key]) {
      map[key] = { period, doctor: row.doctor, count: 0, amount: 0 };
    }

    map[key].count += 1;
    map[key].amount += row.price;
  });

  return Object.values(map).map((item) => [item.period, item.doctor, item.count, item.amount]);
}

function buildMonthlyRows_(rows) {
  const map = {};

  rows.forEach((row) => {
    const period = getMonthCode_(row.date);
    const key = `${period}|${row.doctor}`;

    if (!map[key]) {
      map[key] = { period, doctor: row.doctor, count: 0, amount: 0 };
    }

    map[key].count += 1;
    map[key].amount += row.price;
  });

  return Object.values(map).map((item) => [item.period, item.doctor, item.count, item.amount]);
}

function upsertAggregate_(reportSpreadsheet, sheetName, rows) {
  if (rows.length === 0) {
    return;
  }

  const sheet = reportSpreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Не найден лист "${sheetName}" в книге отчетов.`);
  }

  const existingRange = sheet.getDataRange();
  const existingValues = existingRange.getValues();
  const indexByKey = {};

  for (let i = 1; i < existingValues.length; i += 1) {
    const row = existingValues[i];
    const period = row[0];
    const doctor = String(row[1] || '').trim();
    if (period && doctor) {
      indexByKey[`${period}|${doctor}`] = i + 1;
    }
  }

  const toAppend = [];
  rows.forEach((row) => {
    const key = `${row[0]}|${String(row[1] || '').trim()}`;
    const existingRowNumber = indexByKey[key];

    if (existingRowNumber) {
      const currentCount = Number(sheet.getRange(existingRowNumber, 3).getValue()) || 0;
      const currentAmount = Number(sheet.getRange(existingRowNumber, 4).getValue()) || 0;
      sheet.getRange(existingRowNumber, 3, 1, 2).setValues([[currentCount + row[2], currentAmount + row[3]]]);
      return;
    }

    toAppend.push(row);
  });

  if (toAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
}

function getWeekCode_(date) {
  const year = date.getFullYear();
  const shortYear = year % 100;
  const firstDay = new Date(year, 0, 1);
  const diffDays = Math.floor((date - firstDay) / 86400000);
  const jan1MondayBasedDay = (firstDay.getDay() + 6) % 7;
  const daysToFirstMonday = (7 - jan1MondayBasedDay) % 7;

  let week;
  if (daysToFirstMonday === 0) {
    week = 1 + Math.floor(diffDays / 7);
  } else if (diffDays < daysToFirstMonday) {
    week = 1;
  } else {
    week = 2 + Math.floor((diffDays - daysToFirstMonday) / 7);
  }

  return shortYear * 1000 + week;
}

function getMonthCode_(date) {
  const year = date.getFullYear() % 100;
  const month = date.getMonth() + 1;
  return year * 100 + month;
}
