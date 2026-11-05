/**
 * ПУЛЬТ — серверная часть Google Apps Script.
 * Каждый тип данных хранится в своём листе Google Таблицы.
 * Старый лист _Пульт_Данные сохраняется как источник миграции/резервная копия,
 * но после миграции новые записи в него не добавляются.
 */

const PULT_CONFIG = {
  SPREADSHEET_ID: '13Ob2dT7jj9zfJkjCT0zl1k1t5Sj5MRYW7r_zgVxoDpI',
  LEGACY_SHEET: '_Пульт_Данные',
  DRIVE_FOLDER: 'Пульт — фотографии',
  CHUNK_SIZE: 40000
};

// Карта отдельных листов. В колонке «Данные JSON» хранится полный объект,
// остальные колонки нужны для удобного просмотра и фильтрации прямо в таблице.
const PULT_TABLES = {
  tasks:       { name: 'Задачи',       headers: ['ID','Задача','Срок','Статус','Данные JSON'] },
  reminders:   { name: 'Напоминания',  headers: ['ID','Напоминание','Срок','Статус','Данные JSON'] },
  inboxIdeas:  { name: 'Входящие идеи', headers: ['ID','Идея','Дата','Статус','Данные JSON'] },
  drafts:      { name: 'Черновики',    headers: ['ID','Текст поста','Статус','Дата','Данные JSON'] },
  expenses:    { name: 'Расходы',      headers: ['ID','ID проекта','Наименование','Сумма, ₽','Дата','Данные JSON'] },
  ideas:       { name: 'Идеи',         headers: ['ID','Название','Категория','Статус','Описание','Данные JSON'] },
  projects:    { name: 'Проекты',      headers: ['ID','Название','Ссылка','Статус','Описание','Данные JSON'] },
  notes:       { name: 'Заметки',      headers: ['ID','ID идеи','Проект','Заметка','Статус','Дата','Данные JSON'] },
  photos:      { name: 'Фото',         headers: ['ID','ID идеи','Проект','Имя файла','Ссылка','ID файла','Данные JSON'] }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Пульт')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// Подключение отдельных частей клиентского JavaScript.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Создаёт недостающие листы, не очищая и не удаляя существующие вкладки. */
function setupPult() {
  const ss = SpreadsheetApp.openById(PULT_CONFIG.SPREADSHEET_ID);
  Object.keys(PULT_TABLES).forEach(function(key) { ensurePultSheet_(ss, PULT_TABLES[key]); });

  let folder;
  const folders = DriveApp.getFoldersByName(PULT_CONFIG.DRIVE_FOLDER);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PULT_CONFIG.DRIVE_FOLDER);
  PropertiesService.getScriptProperties().setProperty('PULT_DRIVE_FOLDER_ID', folder.getId());

  // Переносим старые данные только если новые отдельные листы ещё пусты.
  migrateLegacyState_();
  return {
    ok: true,
    spreadsheet: ss.getUrl(),
    sheets: Object.keys(PULT_TABLES).map(function(k) { return PULT_TABLES[k].name; }),
    driveFolder: folder.getUrl(),
    message: 'Отдельные листы подготовлены. Старый лист _Пульт_Данные сохранён.'
  };
}

function ensurePultSheet_(ss, config) {
  let sh = ss.getSheetByName(config.name);
  if (!sh) sh = ss.insertSheet(config.name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, config.headers.length).setFontWeight('bold');
  } else {
    const current = sh.getRange(1, 1, 1, config.headers.length).getDisplayValues()[0];
    const matches = config.headers.every(function(header, index) { return current[index] === header; });
    if (!matches) throw new Error('Лист «' + config.name + '» уже существует, но его заголовки отличаются от схемы Пульта. Данные не изменены. Нужна отдельная настройка сопоставления колонок.');
  }
  return sh;
}

function migrateLegacyState_() {
  const ss = SpreadsheetApp.openById(PULT_CONFIG.SPREADSHEET_ID);
  const legacy = ss.getSheetByName(PULT_CONFIG.LEGACY_SHEET);
  if (!legacy || legacy.getLastRow() < 2) return { migrated: false };
  const anyStructuredData = Object.keys(PULT_TABLES).some(function(key) {
    const sh = ss.getSheetByName(PULT_TABLES[key].name);
    return sh && sh.getLastRow() > 1;
  });
  if (anyStructuredData) return { migrated: false, reason: 'Отдельные листы уже содержат данные.' };

  const chunks = legacy.getRange(2, 2, legacy.getLastRow() - 1, 1).getDisplayValues()
    .map(function(r) { return r[0]; }).filter(Boolean);
  if (!chunks.length) return { migrated: false };
  const oldState = JSON.parse(chunks.join(''));
  pultSaveState(oldState);
  PropertiesService.getScriptProperties().setProperty('PULT_LEGACY_MIGRATED', '1');
  return { migrated: true };
}

/** Читает отдельные листы и собирает прежнюю структуру S для интерфейса. */
function pultLoadState() {
  const ss = SpreadsheetApp.openById(PULT_CONFIG.SPREADSHEET_ID);
  const state = { items: [], drafts: [], exp: [], proj: [], ideas: [] };
  // Настройки сайта (раздел «Настройки») хранятся в свойствах скрипта, а не в таблице.
  state.settings = JSON.parse(PropertiesService.getScriptProperties().getProperty('PULT_SETTINGS') || 'null');

  // Совместимость: если отдельные листы пока пусты, один раз переносим старое состояние.
  const migration = migrateLegacyState_();
  if (migration && migration.migrated) return pultLoadState();

  // Быстрая загрузка: получаем все листы одним запросом и читаем каждый лист целиком за один вызов.
  const byName = {};
  ss.getSheets().forEach(function(sh) { byName[sh.getName()] = sh; });
  let hasAnyRows = false;                                          // есть ли в таблице хоть одна запись
  const read = function(config) {                                  // прочитать лист и разобрать колонку «Данные JSON»
    const sh = byName[config.name];
    if (!sh) return [];
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return [];
    hasAnyRows = true;
    const jsonCol = config.headers.indexOf('Данные JSON');
    return values.slice(1).map(function(row, index) {
      const cell = row[jsonCol];
      if (cell === '' || cell == null) {
        throw new Error('Лист «' + config.name + '», строка ' + (index + 2) + ': отсутствует JSON. Загрузка остановлена, данные не изменены.');
      }
      let parsed;
      try { parsed = JSON.parse(cell); } catch (e) {
        throw new Error('Лист «' + config.name + '», строка ' + (index + 2) + ': JSON повреждён. Загрузка остановлена, данные не изменены.');
      }
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
        throw new Error('Лист «' + config.name + '», строка ' + (index + 2) + ': запись имеет неверный формат.');
      }
      return parsed;
    });
  };

  state.items = read(PULT_TABLES.tasks).concat(read(PULT_TABLES.reminders), read(PULT_TABLES.inboxIdeas));
  state.drafts = read(PULT_TABLES.drafts);
  const allExpenses = read(PULT_TABLES.expenses);
  state.exp = allExpenses.filter(function(x) { return !x.parentIdeaId; });
  state.proj = read(PULT_TABLES.projects);
  state.ideas = read(PULT_TABLES.ideas);

  const notes = read(PULT_TABLES.notes);
  const photos = read(PULT_TABLES.photos);
  state.ideas.forEach(function(idea) {
    idea.notes = notes.filter(function(n) { return String(n.parentIdeaId) === String(idea.id); });
    idea.photos = photos.filter(function(p) { return String(p.parentIdeaId) === String(idea.id); });
    idea.expenses = allExpenses.filter(function(x) { return String(x.parentIdeaId) === String(idea.id); });
  });
  return { ok: true, data: hasAnyRows ? JSON.stringify(state) : null, source: 'separate-sheets' };
}

function readObjects_(ss, config) {
  const sh = ss.getSheetByName(config.name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, config.headers.length).getValues();
  const jsonCol = config.headers.indexOf('Данные JSON');
  return values.map(function(row) {
    try { return JSON.parse(row[jsonCol]); } catch (e) { return null; }
  }).filter(function(x) { return x && typeof x === 'object'; });
}

/** Сохраняет каждый массив в предназначенный для него лист. */
function pultSaveState(state) {
  if (!state || typeof state !== 'object') throw new Error('Получены некорректные данные Пульта.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(PULT_CONFIG.SPREADSHEET_ID);
    Object.keys(PULT_TABLES).forEach(function(key) { ensurePultSheet_(ss, PULT_TABLES[key]); });

    const items = state.items || [];
    const tasks = items.filter(function(x) { return x.k === 'task'; });
    const reminders = items.filter(function(x) { return x.k === 'rem'; });
    const inboxIdeas = items.filter(function(x) { return x.k === 'idea'; });
    const ideas = state.ideas || [];
    const notes = [], photos = [], projectExpenses = [];
    ideas.forEach(function(idea) {
      (idea.notes || []).forEach(function(note) {
        const item = Object.assign({}, note, { parentIdeaId: idea.id, projectTitle: idea.title });
        if (!item.id) item.id = Utilities.getUuid();
        notes.push(item);
      });
      (idea.photos || []).forEach(function(photo) {
        const item = typeof photo === 'string' ? { url: photo } : Object.assign({}, photo);
        item.parentIdeaId = idea.id; item.projectTitle = idea.title;
        if (!item.id) item.id = item.fileId || Utilities.getUuid();
        photos.push(item);
      });
      (idea.expenses || []).forEach(function(expense) {
        const item = Object.assign({}, expense, { parentIdeaId: idea.id, projectTitle: idea.title });
        if (!item.id) item.id = Utilities.getUuid();
        projectExpenses.push(item);
      });
    });

    writeObjects_(ss, PULT_TABLES.tasks, tasks, function(x) { return [x.id, x.t, x.due || '', x.done ? 'Выполнено' : 'Ожидает', JSON.stringify(x)]; });
    writeObjects_(ss, PULT_TABLES.reminders, reminders, function(x) { return [x.id, x.t, x.due || '', 'Активно', JSON.stringify(x)]; });
    writeObjects_(ss, PULT_TABLES.inboxIdeas, inboxIdeas, function(x) { return [x.id, x.t, x.date || '', 'Новая', JSON.stringify(x)]; });
    writeObjects_(ss, PULT_TABLES.drafts, state.drafts || [], function(x) { return [x.id, x.t, x.st || 'draft', x.date || '', JSON.stringify(x)]; });
    writeObjects_(ss, PULT_TABLES.expenses, (state.exp || []).concat(projectExpenses), function(x) { return [x.id, x.parentIdeaId || '', x.label || x.l || '', Number(x.amount != null ? x.amount : x.a) || 0, x.d || '', JSON.stringify(x)]; });
    writeObjects_(ss, PULT_TABLES.ideas, ideas, function(x) { return [x.id, x.title, x.cat || '', x.st || 'idea', x.desc || '', JSON.stringify(x)]; });
    writeObjects_(ss, PULT_TABLES.projects, state.proj || [], function(x) { return [x.id || x.n, x.n || x.title || '', x.u || x.url || '', x.st || '', x.d || x.desc || '', JSON.stringify(x)]; });
    writeObjects_(ss, PULT_TABLES.notes, notes, function(x) { return [x.id, x.parentIdeaId, x.projectTitle, x.t || x.text || '', x.status === 'done' ? 'Выполнено' : 'Ожидает работы', x.d || '', JSON.stringify(x)]; });
    writeObjects_(ss, PULT_TABLES.photos, photos, function(x) { return [x.id, x.parentIdeaId, x.projectTitle, x.name || '', x.url || '', x.fileId || '', JSON.stringify(x)]; });

    if (state.settings) PropertiesService.getScriptProperties().setProperty('PULT_SETTINGS', JSON.stringify(state.settings)); // сохраняем настройки сайта
    return { ok: true, savedAt: new Date().toISOString(), destination: 'separate-sheets' };
  } finally {
    lock.releaseLock();
  }
}

function writeObjects_(ss, config, objects, rowMapper) {
  const sh = ensurePultSheet_(ss, config);
  const oldRows = sh.getLastRow();
  if (oldRows > 1) sh.getRange(2, 1, oldRows - 1, config.headers.length).clearContent();
  if (!objects.length) return;
  const rows = objects.map(rowMapper);
  sh.getRange(2, 1, rows.length, config.headers.length).setValues(rows);
}

/** Фото загружается в подпапку конкретной идеи/проекта. */
function pultUploadPhoto(dataUrl, originalName, ideaId, ideaTitle) {
  if (!dataUrl || String(dataUrl).indexOf('data:image/') !== 0) throw new Error('Файл не распознан как изображение.');
  const match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) throw new Error('Неверный формат изображения.');
  const folderId = PropertiesService.getScriptProperties().getProperty('PULT_DRIVE_FOLDER_ID');
  if (!folderId) throw new Error('Сначала запустите setupPult() в Apps Script.');
  const bytes = Utilities.base64Decode(match[2]);
  const safeName = String(originalName || 'photo').replace(/[\\/:*?"<>|]/g, '_').slice(0, 100);
  const blob = Utilities.newBlob(bytes, match[1], new Date().getTime() + '_' + safeName);
  const root = DriveApp.getFolderById(folderId);
  const projectPrefix = String(ideaId || 'без-проекта') + ' — ';
  const projectFolderName = projectPrefix + String(ideaTitle || 'Проект').replace(/[\\/:*?"<>|]/g, '_').slice(0, 70);
  const folders = root.getFolders(); let projectFolder = null;
  while (folders.hasNext()) { const candidate = folders.next(); if (candidate.getName().indexOf(projectPrefix) === 0) { projectFolder = candidate; break; } }
  if (!projectFolder) projectFolder = root.createFolder(projectFolderName);
  const file = projectFolder.createFile(blob);
  return { ok: true, fileId: file.getId(), name: file.getName(), url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000' };
}

function pultDeletePhoto(fileId) {
  if (!fileId) return { ok: true };
  const folderId = PropertiesService.getScriptProperties().getProperty('PULT_DRIVE_FOLDER_ID');
  if (!folderId) return { ok: true };
  const root = DriveApp.getFolderById(folderId), file = DriveApp.getFileById(String(fileId));
  const parents = file.getParents(); let belongs = false;
  while (parents.hasNext()) {
    const projectFolder = parents.next(), grandparents = projectFolder.getParents();
    while (grandparents.hasNext()) if (grandparents.next().getId() === root.getId()) belongs = true;
  }
  if (belongs) file.setTrashed(true);
  return { ok: true };
}

/** Gemini. Ключ GEMINI_API_KEY хранится в Script Properties. */
function pultGenerateText(kind, idea) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Не задан GEMINI_API_KEY в настройках проекта Apps Script.');
  const prompt = kind === 'desc'
    ? 'На русском языке аккуратно дополни описание проекта по исходному описанию, заметкам и комплектующим. Не выдумывай факты. Верни только готовый текст.\n\n' + JSON.stringify(idea)
    : (kind === 'summary' || kind === 'summary_all')
      ? 'Ты технический редактор и аналитик проектов. Собери заметки в единый грамотный рабочий документ на русском языке. Сохрани факты, даты, решения, проблемы, результаты и следующие шаги. Не выдумывай сведения.\n\n' + JSON.stringify(idea)
      : 'На русском языке подготовь грамотный пост о проекте для публикации. Используй только предоставленные факты и заметки, не выдумывай результаты. Верни только текст поста.\n\n' + JSON.stringify(idea);
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(key);
  const response = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }), muteHttpExceptions: true });
  const status = response.getResponseCode(), body = response.getContentText();
  if (status < 200 || status >= 300) throw new Error('Gemini API (' + status + '): ' + body.slice(0, 500));
  const result = JSON.parse(body), text = result && result.candidates && result.candidates[0] && result.candidates[0].content && result.candidates[0].content.parts && result.candidates[0].content.parts.map(function(p) { return p.text || ''; }).join('');
  if (!text) throw new Error('Gemini вернул пустой ответ.');
  return text.trim();
}

/** Фото из Google Диска → список файлов (blob) для отправки. Не больше 10 штук (лимит Telegram и ВК). */
function pultPhotoBlobs_(fileIds) {
  return (fileIds || []).slice(0, 10).map(function (id, n) {       // перебираем выбранные фото
    return DriveApp.getFileById(String(id)).getBlob().setName('photo' + n + '.jpg'); // берём файл с Диска по его ID
  });
}

/** Один запрос к Telegram. Если в payload есть файлы, UrlFetchApp сам отправит их как multipart. */
function pultTelegramCall_(token, method, payload) {
  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, { method: 'post', payload: payload, muteHttpExceptions: true }); // отправляем запрос
  const body = response.getContentText();                          // ответ Telegram как текст
  let result; try { result = JSON.parse(body); } catch (e) { result = null; } // превращаем в объект
  if (!result || !result.ok) throw new Error('Telegram не принял сообщение: ' + body.slice(0, 500)); // ошибка — показываем причину
  return result.result;                                            // успешный результат
}

/** Публикация в Telegram: только текст или текст + фото из Google Диска (fileIds — ID выбранных фото). */
function pultPublishTelegram(text, fileIds) {
  const props = PropertiesService.getScriptProperties();           // хранилище секретов
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');           // токен бота
  const chatId = props.getProperty('TELEGRAM_CHANNEL_ID');         // ID канала
  if (!token || !chatId) throw new Error('Не настроены TELEGRAM_BOT_TOKEN и TELEGRAM_CHANNEL_ID в Script Properties.');
  const message = String(text || '').trim();                       // текст поста без пробелов по краям
  if (!message) throw new Error('Текст публикации пустой.');
  const blobs = pultPhotoBlobs_(fileIds);                          // выбранные фото (может быть пусто)
  const caption = message.length <= 1024 ? message : '';           // Telegram: подпись к фото не длиннее 1024 знаков
  if (blobs.length === 1) {                                        // одно фото
    pultTelegramCall_(token, 'sendPhoto', Object.assign({ chat_id: chatId, photo: blobs[0] }, caption ? { caption: caption } : {}));
  } else if (blobs.length > 1) {                                   // несколько фото — альбомом
    const payload = { chat_id: chatId, media: JSON.stringify(blobs.map(function (b, n) {
      return Object.assign({ type: 'photo', media: 'attach://p' + n }, n === 0 && caption ? { caption: caption } : {}); // подпись только у первого фото
    })) };
    blobs.forEach(function (b, n) { payload['p' + n] = b; });      // прикладываем сами файлы p0, p1, …
    pultTelegramCall_(token, 'sendMediaGroup', payload);
  }
  if (!blobs.length || !caption) pultTelegramCall_(token, 'sendMessage', { chat_id: chatId, text: message }); // без фото или длинный текст — отдельным сообщением
  return { ok: true, photos: blobs.length };
}

/** Один запрос к API ВКонтакте. Ошибку ВК превращаем в понятное сообщение. */
function pultVkApi_(method, params) {
  const response = UrlFetchApp.fetch('https://api.vk.com/method/' + method, { method: 'post', payload: params, muteHttpExceptions: true });
  const body = response.getContentText();
  let result; try { result = JSON.parse(body); } catch (e) { result = null; }
  if (!result || result.error || !result.response) throw new Error('ВКонтакте (' + method + '): ' + (result && result.error ? result.error.error_msg : body.slice(0, 300)));
  return result.response;
}

/** Загрузка одного фото на стену сообщества. Возвращает строку вида photo-123_456 для поста. */
function pultVkUploadPhoto_(token, groupId, blob) {
  const server = pultVkApi_('photos.getWallUploadServer', { group_id: groupId, access_token: token, v: '5.199' }); // 1) адрес для загрузки
  const up = JSON.parse(UrlFetchApp.fetch(server.upload_url, { method: 'post', payload: { photo: blob }, muteHttpExceptions: true }).getContentText()); // 2) отправляем файл
  if (!up.photo || up.photo === '[]') throw new Error('ВКонтакте не принял фото при загрузке.');
  const saved = pultVkApi_('photos.saveWallPhoto', { group_id: groupId, photo: up.photo, server: up.server, hash: up.hash, access_token: token, v: '5.199' }); // 3) сохраняем фото
  return 'photo' + saved[0].owner_id + '_' + saved[0].id;           // 4) идентификатор для вложения
}

/**
 * Публикация в сообщество ВКонтакте: текст или текст + фото.
 * ВАЖНО: токен сообщества (VK_ACCESS_TOKEN) умеет только публиковать текст. Загружать фото ВК разрешает
 * лишь токену ПОЛЬЗОВАТЕЛЯ-администратора (права photos, wall, groups, offline). Такой токен храните
 * в Script Properties под именем VK_USER_TOKEN. Если его нет, пост уйдёт без фото с предупреждением.
 */
function pultPublishVK(text, fileIds) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('VK_ACCESS_TOKEN');              // токен сообщества
  const userToken = props.getProperty('VK_USER_TOKEN');            // необязательный токен администратора (для фото)
  const groupId = String(props.getProperty('VK_GROUP_ID') || '').replace(/^-/, ''); // ID группы без минуса
  if ((!token && !userToken) || !groupId) throw new Error('Не заданы VK_ACCESS_TOKEN и VK_GROUP_ID в Script Properties.');
  const message = String(text || '').trim();
  if (!message) throw new Error('Текст публикации пустой.');
  const blobs = pultPhotoBlobs_(fileIds);                          // выбранные фото
  const postToken = userToken || token;                            // чем публикуем: токен администратора, если он есть
  let attachments = '', warning = '';
  if (blobs.length) {
    try {
      attachments = blobs.map(function (b) { return pultVkUploadPhoto_(postToken, groupId, b); }).join(','); // загружаем фото
    } catch (e) {
      if (userToken) throw e;                                      // токен администратора есть, а фото не ушло — это настоящая ошибка
      warning = 'Фото во ВК не добавлены, пост опубликован текстом: ' + e.message + '. Чтобы публиковать с фото, добавьте VK_USER_TOKEN.';
    }
  }
  const params = { owner_id: '-' + groupId, from_group: 1, message: message, access_token: postToken, v: '5.199' };
  if (attachments) params.attachments = attachments;              // прикрепляем фото
  const post = pultVkApi_('wall.post', params);                    // публикуем
  return { ok: true, postId: post.post_id, warning: warning };
}
