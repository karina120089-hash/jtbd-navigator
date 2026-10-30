/**
 * Стартовый шаблон живой страницы: источник → Apps Script → страница.
 *
 *   sync()    забирает данные из источников в служебную Таблицу (по триггеру и по кнопке «Обновить»)
 *   doGet()   без параметров отдаёт страницу (index.html), с ?format=json — данные
 *   doPost()  запросы со страницы на Firebase / GitHub: POST text/plain {action, ...}
 *   api()     те же запросы со страницы внутри Apps Script (google.script.run)
 *
 * Секреты и ID — только в «Настройки проекта → Свойства скрипта», никогда в коде:
 *   SOURCE_SHEET_ID  ID исходной Google Таблицы (если источник — Таблица)
 *   BITRIX_HOOK      вебхук Битрикс24: https://портал.bitrix24.ru/rest/ID/КОД/
 *   API_KEY          ключ другой системы с API
 *   GEMINI_API_KEY   ключ Gemini (или LLM_PROVIDER = gigachat и GIGACHAT_AUTH_KEY)
 *   ALERT_EMAIL      кому писать, если загрузка упала (можно несколько через запятую)
 *   SHEET_ID         служебная Таблица — скрипт создаёт её сам при первом запуске
 *
 * Часовой пояс — в appsscript.json (timeZone), все даты форматируем в нём.
 * Ненужные блоки (Битрикс, запись, ИИ) удаляйте целиком.
 */

const TITLE = 'Название страницы';   // заголовок страницы, имя служебной Таблицы, тема писем об ошибках
const CACHE_KEY = 'payload_v1';      // смените, если поменялся формат данных — иначе страница покажет старый кэш
const TZ = Session.getScriptTimeZone();

function props_() { return PropertiesService.getScriptProperties(); }

function prop_(name, required) {
  const v = props_().getProperty(name);
  if (!v && required) throw new Error('Не задано свойство скрипта ' + name + ': «Настройки проекта» → «Свойства скрипта»');
  return v || '';
}

/** Дата в часовом поясе проекта. Без времени, если оно 00:00 */
function fmt_(d, pattern) {
  if (!(d instanceof Date)) return d == null ? '' : d;
  if (pattern) return Utilities.formatDate(d, TZ, pattern);
  const s = Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm');
  return s.slice(11) === '00:00' ? s.slice(0, 10) : s;
}

/** Значение для ячейки: числа как есть; текст — до 49 000 символов и без превращения «=…» в формулу */
function cell_(v) {
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  const s = String(v == null ? '' : v).slice(0, 49000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- служебная Таблица ---------- */

function book_() {
  const id = prop_('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create(TITLE + ' — данные');
  props_().setProperty('SHEET_ID', ss.getId());
  return ss;
}

function sheetFor_(name, head) {
  const ss = book_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
    if (def) ss.deleteSheet(def);
  }
  return sh;
}

/** Перезаписать лист целиком: шапка + строки */
function write_(name, head, rows) {
  const sh = sheetFor_(name, head);
  sh.clearContents();
  const data = [head].concat(rows);
  sh.getRange(1, 1, data.length, head.length).setValues(data);
}

/** Лист → список объектов {заголовок: значение}; даты — строками в поясе проекта */
function rows_(name) {
  const sh = book_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const head = v.shift().map(h => String(h).trim());
  return v.map(r => { const o = {}; head.forEach((h, i) => { if (h) o[h] = fmt_(r[i]); }); return o; });
}

function log_(started, status, text) {
  const sh = sheetFor_('Журнал', ['Время', 'Статус', 'Секунд', 'Итог']);
  sh.appendRow([fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss'), status, Math.round((new Date() - started) / 1000), cell_(text)]);
}

function notifyFail_(err) {
  const to = prop_('ALERT_EMAIL') || Session.getEffectiveUser().getEmail();
  if (to) MailApp.sendEmail(to, TITLE + ': загрузка данных не прошла', String(err && err.stack || err));
}

/* ---------- источник: Google Таблица ---------- */

/** Лист исходной Таблицы как список объектов. Столбцы ищем по названию, а не по букве */
function readSource_(sheetName) {
  const ss = SpreadsheetApp.openById(prop_('SOURCE_SHEET_ID', true));
  const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  if (!sh) throw new Error('В исходной Таблице нет листа «' + sheetName + '»');
  const v = sh.getDataRange().getValues();
  const head = v.shift().map(h => String(h).trim());
  return v.filter(r => r.some(c => c !== ''))
    .map(r => { const o = {}; head.forEach((h, i) => { if (h) o[h] = fmt_(r[i]); }); return o; });
}

/* ---------- источник: Битрикс24 ---------- */

function bitrix_(method, body) {
  const hook = prop_('BITRIX_HOOK', true).replace(/\/?$/, '/');
  for (let attempt = 0; attempt < 4; attempt++) {
    const res = UrlFetchApp.fetch(hook + method + '.json', {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(body || {}), muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 503 || code === 429) { Utilities.sleep(1500 * (attempt + 1)); continue; }
    const d = JSON.parse(res.getContentText() || '{}');
    if (d.error) throw new Error('Битрикс ' + method + ': ' + d.error + ' ' + (d.error_description || ''));
    return d;
  }
  throw new Error('Битрикс ' + method + ': не ответил после повторов');
}

/** Списочные методы: забираем все страницы по 50 (start / next) */
function bitrixAll_(method, body, pick) {
  const out = []; let start = 0;
  while (true) {
    const d = bitrix_(method, Object.assign({}, body, { start: start }));
    out.push.apply(out, pick ? pick(d.result) : d.result);
    if (d.next === undefined) break;
    start = d.next; Utilities.sleep(300);
  }
  return out;
}

/** Дата из Битрикса (ISO со смещением) → строка в поясе проекта */
function bxDate_(s) { return s ? fmt_(new Date(s), 'yyyy-MM-dd HH:mm') : ''; }

/* ---------- загрузка ---------- */

/** Забрать данные из источников. ПРИМЕР: оставьте нужные источники и поля, остальное удалите.
 *  Apps Script обрывает запуск через 6 минут: если данных много — грузите частями. */
function sync() {
  const started = new Date();
  props_().setProperty('LAST_ATTEMPT', started.toISOString());
  try {
    const done = [];
    if (prop_('SOURCE_SHEET_ID')) {
      const items = readSource_('');   // имя листа; пусто — первый лист
      const head = Object.keys(items[0] || {});
      write_('Данные', head.length ? head : ['Пусто'], items.map(o => head.map(h => cell_(o[h]))));
      done.push('строк из Таблицы ' + items.length);
    }
    if (prop_('BITRIX_HOOK')) {
      const tasks = bitrixAll_('tasks.task.list',
        { select: ['ID', 'TITLE', 'RESPONSIBLE_ID', 'STATUS', 'CREATED_DATE', 'DEADLINE', 'CLOSED_DATE'] }, r => r.tasks);
      write_('Задачи', ['ID', 'Название', 'Ответственный', 'Статус', 'Создана', 'Крайний срок', 'Закрыта'],
        tasks.map(t => [+t.id, cell_(t.title), +t.responsibleId, +t.status, bxDate_(t.createdDate), bxDate_(t.deadline), bxDate_(t.closedDate)]));
      done.push('задач из Битрикса ' + tasks.length);
    }
    log_(started, 'ok', done.join(', ') || 'источники не настроены');
    props_().setProperty('LAST_SYNC', new Date().toISOString());
    props_().deleteProperty('LAST_ERROR');
    cacheClear_();
    SpreadsheetApp.flush();
    payload_();   // сразу готовим кэш, чтобы страница открывалась быстро
  } catch (err) {
    const msg = String(err && err.message || err);
    try { log_(started, 'ошибка', msg); } catch (e) {}
    props_().setProperty('LAST_ERROR', JSON.stringify({ at: new Date().toISOString(), message: msg.slice(0, 500) }));
    cacheClear_();
    notifyFail_(err);
    throw err;
  }
}

/** Запуск sync() по расписанию: раз в N часов */
function installTrigger(hours) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sync').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sync').timeBased().everyHours(hours || 1).create();
}

/** Кнопка «Обновить»: не чаще раза в 2 минуты и без параллельных запусков */
function syncFromPage_() {
  const last = Date.parse(prop_('LAST_ATTEMPT') || 0);
  if (Date.now() - last < 120000) return { ok: true, status: 'recent', message: 'Данные обновлялись меньше двух минут назад' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { ok: true, status: 'busy', message: 'Обновление уже идёт' };
  try { sync(); return { ok: true, status: 'done' }; }
  catch (err) { return { ok: false, error: String(err && err.message || err) }; }
  finally { lock.releaseLock(); }
}

/* ---------- данные для страницы ---------- */

/* Кэш Google хранит до 100 КБ на ключ — режем ответ на части */
function cacheGet_() {
  const c = CacheService.getScriptCache(); const n = +(c.get(CACHE_KEY + '_n') || 0);
  if (!n) return null;
  const keys = []; for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
  const parts = c.getAll(keys);
  if (Object.keys(parts).length !== n) return null;
  return keys.map(k => parts[k]).join('');
}
function cachePut_(json) {
  const c = CacheService.getScriptCache(); const size = 90000, map = {};
  const n = Math.ceil(json.length / size);
  for (let i = 0; i < n; i++) map[CACHE_KEY + '_' + i] = json.slice(i * size, (i + 1) * size);
  map[CACHE_KEY + '_n'] = String(n);
  c.putAll(map, 21600);
}
function cacheClear_() { CacheService.getScriptCache().remove(CACHE_KEY + '_n'); }

/** Всё, что видит страница. Форма этого объекта = форма DEMO в index.html */
function payload_() {
  const hit = cacheGet_();
  if (hit) return hit;
  const out = {
    title: TITLE,
    tz: TZ,
    asOf: prop_('LAST_SYNC'),
    lastError: JSON.parse(prop_('LAST_ERROR') || 'null'),
    rows: rows_('Данные'),   // ← соберите здесь то, что нужно странице
    tasks: rows_('Задачи')
  };
  const json = JSON.stringify(out);
  cachePut_(json);
  return json;
}

/* ---------- запись со страницы ---------- */

/** ПРИМЕР: дописать строку в лист «Записи». Поля — под задачу */
function save_(req) {
  return locked_(() => {
    const sh = sheetFor_('Записи', ['Дата', 'Время', 'Кто', 'Текст']);
    const now = new Date();
    sh.appendRow([fmt_(now, 'yyyy-MM-dd'), fmt_(now, 'HH:mm:ss'), cell_(req.who), cell_(req.text)]);
    return { ok: true };
  });
}

/* ---------- ИИ (Gemini или GigaChat) ---------- */

/* Safari обрывает запрос, если сервер молчит около минуты. Поэтому ответ кладём в кэш
   под номером запроса reqId, а страница забирает его короткими запросами aiResult. */
function aiTracked_(req) {
  const id = /^[\w-]{6,64}$/.test(req.reqId || '') ? 'ai_' + req.reqId : '';
  const cache = CacheService.getScriptCache();
  if (id) cache.put(id, JSON.stringify({ ok: true, status: 'pending' }), 1800);
  let res;
  try {
    const text = String(req.text || '').trim();
    if (!text) throw new Error('Нет текста для разбора');
    if (text.length > 200000) throw new Error('Текст слишком длинный — сократите');
    res = { ok: true, status: 'done', text: callLlm_(AI_PROMPT.replace('{{TEXT}}', text)) };
  } catch (err) {
    res = { ok: false, error: String(err && err.message || err) };
  }
  if (id) { try { cache.put(id, JSON.stringify(res), 1800); } catch (e) { cache.remove(id); } }   // больше 100 КБ — отдаём только напрямую
  return res;
}

function aiResult_(req) {
  const id = /^[\w-]{6,64}$/.test(req.reqId || '') ? 'ai_' + req.reqId : '';
  const hit = id && CacheService.getScriptCache().get(id);
  return hit ? JSON.parse(hit) : { ok: true, status: 'unknown' };
}

/** schema — необязательная схема ответа: тогда Gemini вернёт строго JSON по ней */
function callLlm_(prompt, schema) {
  const provider = (prop_('LLM_PROVIDER') || 'gemini').toLowerCase();
  return provider === 'gigachat' ? gigachat_(prompt) : gemini_(prompt, schema);
}

function gemini_(prompt, schema) {
  const key = prop_('GEMINI_API_KEY', true);
  // Основная модель и запасные: если одна перегружена (503), упёрлась в лимит (429) или снята (404) —
  // идём к следующей. Полная flash на бесплатном тарифе быстро выбирает дневной лимит, поэтому по умолчанию lite.
  const cache = CacheService.getScriptCache();
  let models = [prop_('LLM_MODEL') || 'gemini-flash-lite-latest']
    .concat((prop_('LLM_FALLBACK_MODELS') || 'gemini-flash-latest').split(/\s*,\s*/));
  const good = cache.get('gemini_good_model');
  if (good && models.indexOf(good) > 0) models = [good].concat(models.filter(m => m !== good));
  const config = { temperature: 0.2, maxOutputTokens: 8192 };
  if (schema) { config.responseMimeType = 'application/json'; config.responseSchema = schema; }
  const req = {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, muteHttpExceptions: true,
    payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: config })
  };
  let body = {}, code = 0;
  for (let i = 0; i < models.length; i++) {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent', req);
    body = JSON.parse(res.getContentText() || '{}');
    code = res.getResponseCode();
    if (code === 200) { cache.put('gemini_good_model', models[i], 6 * 3600); break; }
    if ([404, 429, 503].indexOf(code) < 0) break;   // ключ или запрос сменой модели не лечатся
    console.warn('Gemini ' + models[i] + ' → ' + code + ', пробую следующую модель');
  }
  if (code !== 200) throw new Error('Gemini ' + code + ': ' + (body.error && body.error.message || ''));
  const parts = (((body.candidates || [])[0] || {}).content || {}).parts || [];
  const text = parts.map(p => p.text || '').join('').trim();
  if (!text) throw new Error('Модель вернула пустой ответ');
  return text;
}

function gigachat_(prompt) {
  const authKey = prop_('GIGACHAT_AUTH_KEY', true);
  // Сертификаты GigaChat выданы российским УЦ, которого нет в доверенных у Google
  const tls = { validateHttpsCertificates: false, muteHttpExceptions: true };
  const cache = CacheService.getScriptCache();
  let token = cache.get('gigachat_token');
  if (!token) {
    const auth = UrlFetchApp.fetch('https://ngw.devices.sberbank.ru:9443/api/v2/oauth', Object.assign({
      method: 'post', headers: { Authorization: 'Basic ' + authKey, RqUID: Utilities.getUuid() },
      payload: { scope: prop_('GIGACHAT_SCOPE') || 'GIGACHAT_API_PERS' }
    }, tls));
    const a = JSON.parse(auth.getContentText() || '{}');
    if (!a.access_token) throw new Error('GigaChat: не получил токен (' + auth.getResponseCode() + ')');
    token = a.access_token;
    cache.put('gigachat_token', token, 25 * 60);   // токен живёт 30 минут
  }
  const res = UrlFetchApp.fetch('https://gigachat.devices.sberbank.ru/api/v1/chat/completions', Object.assign({
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ model: prop_('LLM_MODEL') || 'GigaChat', temperature: 0.2, messages: [{ role: 'user', content: prompt }] })
  }, tls));
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) {
    if (res.getResponseCode() === 401) cache.remove('gigachat_token');
    throw new Error('GigaChat ' + res.getResponseCode() + ': ' + (body.message || ''));
  }
  const text = (((body.choices || [])[0] || {}).message || {}).content || '';
  if (!text.trim()) throw new Error('Модель вернула пустой ответ');
  return text.trim();
}

/** Проверка связи с моделью: запустить из редактора, ответ — в «Журнале выполнения» */
function testAgent() {
  Logger.log('Ответ модели: ' + callLlm_('Ответь одной фразой на русском: связь работает?'));
}

/** Какие модели Gemini доступны ключу: запустить из редактора, список — в «Журнале выполнения» */
function listModels() {
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200',
    { headers: { 'x-goog-api-key': prop_('GEMINI_API_KEY', true) }, muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('Gemini ' + res.getResponseCode() + ': ' + (body.error && body.error.message || ''));
  Logger.log((body.models || []).filter(m => (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0)
    .map(m => m.name.replace('models/', '')).join(', '));
}

/* ---------- входы ---------- */

const ACTIONS = {
  data: () => JSON.parse(payload_()),
  sync: () => syncFromPage_(),
  save: save_,
  ai: aiTracked_,
  aiResult: aiResult_
};

function handle_(req) {
  try {
    const fn = ACTIONS[req.action];
    if (!fn) return { ok: false, error: 'Неизвестное действие: ' + req.action };
    return fn(req);
  } catch (err) {
    console.error(err);
    return { ok: false, error: String(err && err.message || err) };
  }
}

function json_(o) {
  return ContentService.createTextOutput(typeof o === 'string' ? o : JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'sync') return json_(syncFromPage_());
  if (p.format === 'json') return json_(payload_());
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** POST со страницы на другом домене: Content-Type text/plain, чтобы браузер не слал предзапрос */
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents || '{}'); }
  catch (err) { return json_({ ok: false, error: 'Запрос не JSON' }); }
  return json_(handle_(req));
}

/** Для страницы внутри Apps Script: google.script.run.api(req) */
function api(req) { return JSON.stringify(handle_(req || {})); }

/** Запустить один раз из редактора: выдать разрешения и создать служебную Таблицу */
function setup() {
  Logger.log('Служебная Таблица: ' + book_().getUrl());
  Logger.log('Часовой пояс проекта: ' + TZ);
}
