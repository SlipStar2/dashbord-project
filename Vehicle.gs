/**
 * ПУЛЬТ — журнал обслуживания автомобиля.
 * Данные сохраняются в отдельные листы Google Таблицы, указанной в PULT_CONFIG.
 * Файл можно добавить в тот же Apps Script проект, где находится Code.gs.
 */

const VEHICLE_CONFIG = {
  SPREADSHEET_ID: '13Ob2dT7jj9zfJkjCT0zl1k1t5Sj5MRYW7r_zgVxoDpI',
  CAR_SHEET: 'Автомобиль',
  LOG_SHEET: 'Обслуживание авто',
  CAR_HEADERS: ['ID', 'Марка и модель', 'Год', 'Госномер', 'VIN', 'Двигатель', 'Текущий пробег', 'Данные JSON'],
  LOG_HEADERS: ['ID', 'Дата', 'Пробег, км', 'Категория', 'Работа / деталь', 'Стоимость, ₽', 'Следующая дата', 'Следующий пробег, км', 'Комментарий', 'Данные JSON']
};

function vehicleSetup() {
  const ss = SpreadsheetApp.openById(VEHICLE_CONFIG.SPREADSHEET_ID);
  const carSheet = vehicleEnsureSheet_(ss, VEHICLE_CONFIG.CAR_SHEET, VEHICLE_CONFIG.CAR_HEADERS);
  vehicleEnsureSheet_(ss, VEHICLE_CONFIG.LOG_SHEET, VEHICLE_CONFIG.LOG_HEADERS);
  if (carSheet.getLastRow() < 2) {
    const car = {
      id: 'main-car',
      name: 'Hyundai Solaris',
      year: 2017,
      plate: '',
      vin: '',
      engine: '1.6',
      odometer: ''
    };
    carSheet.appendRow([
      car.id, car.name, car.year, car.plate, car.vin, car.engine, car.odometer,
      JSON.stringify(car)
    ]);
  }
  return { ok: true, spreadsheetUrl: ss.getUrl(), message: 'Листы автомобиля подготовлены.' };
}

function vehicleEnsureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else {
    const existing = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (!headers.every((v, i) => existing[i] === v)) {
      throw new Error('Заголовки листа «' + name + '» не совпадают. Лист не изменён.');
    }
  }
  return sh;
}

function vehicleGetData() {
  const ss = SpreadsheetApp.openById(VEHICLE_CONFIG.SPREADSHEET_ID);
  const carSheet = ss.getSheetByName(VEHICLE_CONFIG.CAR_SHEET);
  const logSheet = ss.getSheetByName(VEHICLE_CONFIG.LOG_SHEET);
  if (!carSheet || !logSheet) {
    return { ok: false, setupRequired: true, message: 'Сначала запустите vehicleSetup().' };
  }
  const carRows = carSheet.getLastRow() > 1 ? carSheet.getRange(2, 1, carSheet.getLastRow() - 1, VEHICLE_CONFIG.CAR_HEADERS.length).getValues() : [];
  let car = carRows.length ? vehicleParse_(carRows[0][7]) : null;
  if (!car) car = { id: 'main-car', name: 'Автомобиль', year: '', plate: '', vin: '', engine: '', odometer: '' };

  const last = logSheet.getLastRow();
  const rows = last > 1 ? logSheet.getRange(2, 1, last - 1, VEHICLE_CONFIG.LOG_HEADERS.length).getValues() : [];
  const records = rows.map(function(row) {
    return vehicleParse_(row[9]) || {
      id: String(row[0]), date: vehicleDate_(row[1]), odometer: Number(row[2]) || '',
      category: String(row[3]), title: String(row[4]), cost: Number(row[5]) || 0,
      nextDate: vehicleDate_(row[6]), nextOdometer: Number(row[7]) || '', note: String(row[8])
    };
  }).filter(Boolean);
  records.sort(function(a, b) {
    return String(b.date || '').localeCompare(String(a.date || '')) || Number(b.odometer || 0) - Number(a.odometer || 0);
  });
  return { ok: true, car: car, records: records };
}

function vehicleSaveCar(car) {
  if (!car || typeof car !== 'object') throw new Error('Не переданы данные автомобиля.');
  const ss = SpreadsheetApp.openById(VEHICLE_CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(VEHICLE_CONFIG.CAR_SHEET) || vehicleEnsureSheet_(ss, VEHICLE_CONFIG.CAR_SHEET, VEHICLE_CONFIG.CAR_HEADERS);
  const clean = {
    id: 'main-car',
    name: vehicleText_(car.name, 120),
    year: vehicleNumber_(car.year, 1900, 2100, true),
    plate: vehicleText_(car.plate, 30),
    vin: vehicleText_(car.vin, 40),
    engine: vehicleText_(car.engine, 50),
    odometer: vehicleNumber_(car.odometer, 0, 10000000, true)
  };
  const row = [clean.id, clean.name, clean.year, clean.plate, clean.vin, clean.engine, clean.odometer, JSON.stringify(clean)];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (sh.getLastRow() < 2) sh.appendRow(row);
    else sh.getRange(2, 1, 1, row.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, car: clean };
}

function vehicleSaveRecord(record) {
  if (!record || typeof record !== 'object') throw new Error('Не передана запись обслуживания.');
  const clean = {
    id: vehicleText_(record.id, 80) || Utilities.getUuid(),
    date: vehicleDate_(record.date),
    odometer: vehicleNumber_(record.odometer, 0, 10000000, false),
    category: vehicleText_(record.category, 60) || 'Другое',
    title: vehicleText_(record.title, 180),
    cost: vehicleNumber_(record.cost, 0, 100000000, true) || 0,
    nextDate: vehicleDate_(record.nextDate),
    nextOdometer: vehicleNumber_(record.nextOdometer, 0, 10000000, true),
    note: vehicleText_(record.note, 2000)
  };
  if (!clean.date) throw new Error('Укажите дату обслуживания.');
  if (clean.odometer === '' || clean.odometer == null) throw new Error('Укажите пробег на момент обслуживания.');
  if (!clean.title) throw new Error('Укажите название работы или детали.');

  const ss = SpreadsheetApp.openById(VEHICLE_CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(VEHICLE_CONFIG.LOG_SHEET) || vehicleEnsureSheet_(ss, VEHICLE_CONFIG.LOG_SHEET, VEHICLE_CONFIG.LOG_HEADERS);
  const row = [
    clean.id, clean.date, clean.odometer, clean.category, clean.title, clean.cost,
    clean.nextDate, clean.nextOdometer, clean.note, JSON.stringify(clean)
  ];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const last = sh.getLastRow();
    let target = 0;
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === clean.id) { target = i + 2; break; }
      }
    }
    if (target) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, record: clean };
}

function vehicleDeleteRecord(id) {
  id = vehicleText_(id, 80);
  if (!id) throw new Error('Не указан ID записи.');
  const ss = SpreadsheetApp.openById(VEHICLE_CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(VEHICLE_CONFIG.LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, deleted: false };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === id) {
        sh.deleteRow(i + 2);
        return { ok: true, deleted: true };
      }
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true, deleted: false };
}

function vehicleParse_(value) {
  if (!value) return null;
  try { return typeof value === 'object' ? value : JSON.parse(String(value)); }
  catch (e) { return null; }
}

function vehicleText_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength);
}

function vehicleNumber_(value, min, max, optional) {
  if (value === '' || value == null) return optional ? '' : null;
  const n = Number(value);
  if (!isFinite(n) || n < min || n > max) throw new Error('Числовое значение вне допустимого диапазона.');
  return n;
}

function vehicleDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}
