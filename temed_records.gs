const REPORT_SPREADSHEET_ID = '1xjyk0eGgjDI2VLZpxDxGtJfQ26xU3SV9ULTRghjJLS8';
const CLINIC_MAPPING_SHEET = 'Соответствие клиник';
const CITY_MAPPING_SHEET = 'Соответствие городов';
const RAW_REPORT_SHEET = 'Все записи';
const MENU_NAME = 'TEMED';
const MENU_ITEM_NAME = 'Обработать записи';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem(MENU_ITEM_NAME, 'processTemedRecords')
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
