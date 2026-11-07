/**
 * EduApp OP — API для дашборда.
 * Отдаёт ТОЛЬКО агрегаты (количества и суммы по дням, воронкам, этапам, источникам, акциям, продуктам, менеджерам).
 * Без ID сделок, названий, ФИО детей, контактов, телефонов, комментариев и тегов. Сама таблица остаётся закрытой.
 * Ключ не нужен: дашборд обращается по адресу веб-приложения напрямую.
 *
 * Обновление кода (URL при этом не меняется):
 *  1. Таблица → Расширения → Apps Script → файл EduApp_OP_API → заменить весь код этим → Сохранить.
 *  2. Вверху выбрать функцию zopSetup → «Выполнить» → разрешить доступ (один раз: включает автообновление каждые 10 минут).
 *  3. «Начать развёртывание» → «Управление развёртываниями» → карандаш → Версия: «Новая версия» → «Развернуть».
 * Ответ берётся из кэша, который обновляется каждые 10 минут; ?fresh=1 — пересчитать сразу.
 *
 * Карточки ежедневных отчётов (отдельная страница без сайта, ссылку можно пересылать):
 *   …/exec?card=leads&d=2026-10-08  — лиды за день (amoCRM таблица, лист «Сделки»)
 *   …/exec?card=sales&d=2026-10-07  — продажи за день (таблица продаж «Оқушы тізімі» + платформа; продукты и план — листы «Өнімдер» и «Жоспар» этой таблицы)
 *   без d — за вчера. …/exec?sales=1 — агрегаты продаж в JSON (без имён и телефонов).
 * Веб-приложение должно выполняться от вашего имени («Выполнять как: я»), чтобы читать таблицы продаж и платформы.
 */
// листы, которые заполняет синхронизация amoCRM (с 09.10.2026: «CRM Лиды — авто» и «CRM Журнал — авто» вместо «Сделки» и «Журнал»)
const ZOP = {LEADS: 'CRM Лиды — авто', PIPES: 'Воронки и этапы', LOG: 'CRM Журнал — авто', MAIL: 'Рассылки', COST: 'Расходы', KPI: 'KPI'};
// с 11.10.2026 amoCRM выгружается в отдельную таблицу, лист «Сделки» (только чтение). Листы ищутся по порядку.
// Сделки (LEADS) — ТОЛЬКО из этой таблицы (решение владельца 12.10.2026): прежний «CRM Лиды — авто» не читается;
// нет доступа или листа — ошибка, а не тихая подмена. Этапы и журнал — сначала там же, иначе прежний лист этой таблицы.
const ZOP_CRM = {ID: 'YOUR_CRM_SHEET_ID', LEADS: ['Сделки'], PIPES: ['Воронки и этапы'], LOG: ['Журнал', 'CRM Журнал — авто']};
// Главная таблица проекта (служебные листы: Расходы, Расходы журналы, Рассылки, KPI, Жоспар, Өнімдер, Өнім тізімі, Тарифтер, Гипотезалар).
// С 12.10.2026 — «All Info» (решение владельца). Переход сам: как только в «All Info» есть лист «Расходы» (перенос — zopMoveServiceSheets),
// главной становится она; до этого — AMO. Всё читается и пишется в главную таблицу явно — неважно, к какой таблице привязан проект.
const ZOP_INFO_ID = 'YOUR_MAIN_SHEET_ID', ZOP_HOME_ID = ZOP_INFO_ID, ZOP_HOME_OLD = ZOP_CRM.ID;
let ZOP_HOME_CACHE = null;
function zopHome_() {
  if (ZOP_HOME_CACHE) return ZOP_HOME_CACHE;
  let a = null; try { a = SpreadsheetApp.getActive(); } catch (_) {}
  let info = null; try { info = a && a.getId() === ZOP_HOME_ID ? a : SpreadsheetApp.openById(ZOP_HOME_ID); } catch (_) {}
  if (info && info.getSheetByName('Расходы')) return (ZOP_HOME_CACHE = info);
  try { return (ZOP_HOME_CACHE = a && a.getId() === ZOP_HOME_OLD ? a : SpreadsheetApp.openById(ZOP_HOME_OLD)); } catch (e) { if (a) return a; throw e; }
}
/** Pure: какой лист читать — первый найденный в новой таблице, иначе прежний в этой. has(where, name) → bool. */
function zopPickSheet_(has, list, fallback) {
  for (const n of list) if (has('crm', n)) return {where: 'crm', name: n};
  return fallback && has('own', fallback) ? {where: 'own', name: fallback} : null;
}
const ZOP_CRM_STRICT = {LEADS: true};   // без запасного листа
// «All Info» — собранная владельцем таблица (12.10.2026): «Сделки» (amoCRM), «ОП миникурс», «All Trainings». Основной источник данных;
// если лист там не готов (ошибка IMPORTRANGE, «Loading…», пусто) — тот же лист в AMO, затем оригинал. Этапы и журнал amoCRM — в AMO.
const ZOP_INFO = {ID: ZOP_INFO_ID, title: 'All Info'};
function zopInfoRows_(name, need) {
  try { const sh = SpreadsheetApp.openById(ZOP_INFO.ID).getSheetByName(name); if (!sh || sh.getLastRow() < 2) return null;
    const rows = sh.getDataRange().getValues(); return zopMirrorOk_(rows, need || []) ? rows : null; } catch (e) { return null; }
}
/** Строки листа CRM + откуда они взяты (для статуса источника). */
function zopCrmRead_(kind) {
  if (kind === 'LEADS') { const r = zopInfoRows_(ZOP_CRM.LEADS[0]); if (r) return {rows: r, src: ZOP_INFO.title + ' · ' + ZOP_CRM.LEADS[0], err: ''}; }
  const own = zopHome_(); let crm = null, err = '';
  try { crm = SpreadsheetApp.openById(ZOP_CRM.ID); } catch (e) { err = e.message; }
  const book = w => w === 'crm' ? crm : own, has = (w, n) => !!(book(w) && book(w).getSheetByName(n));
  const pick = zopPickSheet_(has, ZOP_CRM[kind], ZOP_CRM_STRICT[kind] ? null : ZOP[kind]);
  if (!pick) return {rows: [], src: null, err: err || 'amoCRM таблицасында «' + ZOP_CRM[kind].join('» / «') + '» парағы жоқ'};
  const sh = book(pick.where).getSheetByName(pick.name);
  return {rows: sh.getLastRow() > 0 ? sh.getDataRange().getValues() : [], src: (pick.where === 'crm' ? 'amoCRM таблица · ' : 'эта таблица · ') + pick.name, err};
}
// в «CRM Лиды — авто» причина отказа приходит только как ID; названия — как в amoCRM
// таблицы продаж: строки читаются только для подсчёта сумм и количеств; имена, телефоны, чеки не выходят из скрипта
// таблица продаж («ОП Миникурс») и платформа — ТОЛЬКО ЧТЕНИЕ; листы «Жоспар» и «Өнімдер» живут в главной таблице (AMO) (AmoCrm выгрузка)
const ZOP_USE_OP_SHEET = false;   // 16.10.2026 иесі: «полностью с ОП миникурс перейти в Платформа+СРМ» — «Оқушы тізімі» оқылмайды; ақша = CRM «Успешно» + платформа
const ZOP_SALES = {ID: 'YOUR_SALES_SHEET_ID', LIST: 'Оқушы тізімі', PLAN: 'Жоспар', PROD: 'Өнімдер'};
// платформа: лист «All Trainings and Challanges» (решение владельца 12.10) — собран формулой из «все тренинги и челленджи платформа» + «ОП арқылы»
// (копия ОП без invoice — не считается второй раз, по ней идёт построчная сверка). Исходный лист платформы — запасной вариант.
const ZOP_PLATFORM = {ID: 'YOUR_PLATFORM_SHEET_ID', LIST: 'All Trainings and Challanges', LIST_RAW: 'все тренинги и челленджи платформа'};
const ZOP_REASONS = {1000101: 'Слишком дорого', 1000102: 'Пропала потребность', 1000103: 'Не устроили условия', 1000104: 'Выбрали других'};
const ZOP_COST_HEAD = ['Дата', 'Подписка', 'Детали', 'Сумма, ₸', 'Оплата', 'Рассылка (если знаете)'];
// пример платежей сервиса WhatsApp-рассылок (демо-суммы; реальные вносятся в лист «Расходы»)
const ZOP_COST_SEED = [['07.10.2026', 'Рассылка', 'Subscription balance', '15000', 'партнёр', ''],
  ['07.10.2026', 'Рассылка', 'Subscription balance', '10000', 'Stripe', ''],
  ['06.10.2026', 'Рассылка', 'Subscription balance', '25000', 'партнёр', ''],
  ['06.10.2026', 'Рассылка', 'Subscription balance', '45000', 'Stripe', ''],
  ['05.10.2026', 'Рассылка', 'Subscription balance', '5000', 'Stripe', ''],
  ['05.10.2026', '7000', 'Subscription balance', '0', 'бонусы', ''],
  ['05.10.2026', 'Рассылка', 'Subscription balance', '0', 'бонусы', ''],
  ['05.10.2026', '7000', 'Subscription balance', '5000', 'партнёр', ''],
  ['05.10.2026', '7000', 'Subscription balance', '30000', 'партнёр', ''],
  ['05.10.2026', '7000', 'Subscription balance', '5000', 'партнёр', ''],
  ['05.10.2026', 'Рассылка', 'Subscription balance', '15000', 'партнёр', ''],
  ['04.10.2026', '7000', '1 channel on the "Waba" plan for 1 year', '300000', 'партнёр', ''],
  ['04.10.2026', 'Рассылка', 'Subscription balance', '5000', 'Stripe', ''],
  ['02.10.2026', 'Рассылка', 'Subscription balance', '30000', 'Stripe', ''],
  ['01.10.2026', 'Рассылка', 'Subscription balance', '10000', 'Stripe', ''],
  ['29.09.2026', 'Рассылка', '1 channel on the "Waba" plan for 1 month', '5000', 'партнёр', ''],
  ['28.09.2026', 'Рассылка', 'Subscription balance', '5000', 'Stripe', ''],
  ['26.09.2026', 'Рассылка', 'Subscription balance', '25000', 'Stripe', ''],
  ['26.09.2026', 'Рассылка', 'Subscription balance', '25000', 'Stripe', ''],
  ['26.09.2026', 'Рассылка', 'Subscription balance', '5000', 'Stripe', ''],
  ['25.09.2026', 'Рассылка', 'Subscription balance', '10000', 'Stripe', ''],
  ['10.08.2026', 'WhatsApp', '1 channel on the "Max" plan for 1 year', '150000', 'партнёр', '']];
const ZOP_MAIL_HEAD = ['Рассылка', 'Тема', 'Дата начала', 'Дата окончания', 'Воронка', 'Кому (сегмент)', 'Доп. расход, ₸ (кроме WhatsApp)', 'Отправлено, шт', 'Комментарий'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.card) return zopCardPage_(p);
  if (p.snap) return zopOut_(JSON.stringify(zopSnapGet_(String(p.snap).slice(0, 10))));   // бір күннің срезі (карточка өткен күнге)
  if (p.sales) return zopOut_((p.fresh ? null : zopCacheGet_('zops')) || zopSalesRefresh_());
  const hit = p.fresh ? null : zopCacheGet_();
  if (hit) return zopOut_(hit);
  try { return zopOut_(zopRefresh_()); }
  catch (err) { return zopOut_(JSON.stringify({error: 'crm-source', message: String(err.message || err)})); }   // Атлас көрсетеді, срезді қалдырады   // normally served from the cache that the 10-minute trigger keeps warm
}

/**
 * Запустите ОДИН раз из редактора (выбрать zopSetup → «Выполнить» → разрешить доступ):
 * включает автообновление каждые 10 минут и сразу считает данные. В журнале — сколько сделок прочитано и за сколько секунд.
 */
function zopSetup() {
  zopRequireOwner_();
  const created = zopCreateMailSheet_(), costs = zopCreateCostSheet_(), plan = zopCreatePlanSheet_(), kpi = zopCreateKpiSheet_();
  let prods = 0; try { prods = zopCreateProductSheet_(); } catch (err) { Logger.log('Лист «Өнімдер» не создан: ' + err.message); }
  ScriptApp.getProjectTriggers().filter(t => ['zopRefresh_', 'zopTick_'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('zopTick_').timeBased().everyMinutes(10).create();
  const t0 = Date.now(), payload = zopRefresh_(), d = JSON.parse(payload);
  if (costs) Logger.log('Создан лист «Расходы»: ' + costs + ' платежей — проверьте и дописывайте новые');
  if (kpi) Logger.log('Создан лист «KPI»: впишите цели на месяц (Ай · Бағыт · Көрсеткіш · Мақсат)');
  if (plan) Logger.log('Создан лист «Жоспар» (в этой таблице): планы и источники — проверьте и меняйте сами');
  if (prods) Logger.log('Создан лист «Өнімдер» (в этой таблице): ' + prods + ' строк — настоящие названия из платформы (выпадающий список из листа «Өнім тізімі»). Проверьте раздел, название и «Әрқашан көрсету»');
  try { const sd = JSON.parse(zopSalesRefresh_()); if (!sd.rows.platform) Logger.log('ВНИМАНИЕ: этот аккаунт не видит таблицу платформы — автообновление будет без платформы. Запустите zopSetup под аккаунтом с доступом.'); Logger.log('Продажи: строк «Оқушы тізімі» ' + sd.rows.sheet + ', платформы ' + sd.rows.platform + (sd.errors.length ? ' · ' + sd.errors.join('; ') : '')); }
  catch (err) { Logger.log('Продажи не прочитаны: ' + err.message); }
  if (created) Logger.log('Создан лист «Рассылки»: ' + created + ' рассылок найдено — впишите тему и расход');
  Logger.log('Готово: сделок ' + d.rows + ' · последняя синхронизация ' + d.snapshot + ' · расчёт ' + Math.round((Date.now() - t0) / 1000) + ' с · ' + payload.length + ' байт · автообновление каждые 10 минут включено');
}

/* =================== Срездер: amoCRM тек қазіргі күйді береді — күн сайын сақтаймыз (иесі 15.10.2026) ===================
   Лист «Срездер» (All Info): бір күн — бір жол, әр 10 минуттағы есептеу сол күннің жолын жаңартады → соңғы жазба = күн соңындағы күй.
   Атаулар мәтінмен сақталады (сөздік өзгерсе де срез дұрыс қалады). Тек агрегаттар: телефон, аты-жөн, чек жоқ. */
const ZOP_SNAP = 'Срездер', ZOP_SNAP_HEAD = ['Күні', 'Уақыты (дерек)', 'Жауап күткен / ашық сделкалар (менеджер · воронка · этап · күн · тапсырма · саны)', 'Ашық сделкалар этап бойынша (воронка · этап · менеджер · саны)', 'Төлем ↔ сделка', 'Қысқаша'];
/** Pure: есептелген жауаптан бір күннің срезі (атаулармен). */
function zopSnapOf_(out) {
  const D = out.dict || {}, MN = D.manager || [], FN = D.funnel || [], ST = D.stage || [], OUT = D.outcome || [], open = OUT.indexOf('В работе');
  const idle = (out.idle || []).map(r => [MN[r[0]] || '—', FN[r[1]] || '', ST[r[2]] || '', r[3], r[4], r[r.length - 1]]);
  const st = {}; (out.main || []).forEach(r => { if (r[3] !== open) return; const k = [FN[r[1]] || '', ST[r[2]] || '', MN[r[4]] || '—'].join('\u0001'), n = r.length >= 14 ? r[12] : r.length >= 13 ? r[11] : r[10]; st[k] = (st[k] || 0) + n; });
  const stages = Object.entries(st).map(([k, n]) => k.split('\u0001').concat([n])).sort((a, b) => b[3] - a[3]);
  const sum = {open: stages.reduce((t, r) => t + r[3], 0), late: idle.filter(r => r[4] === 2).reduce((t, r) => t + r[5], 0), idle3: idle.filter(r => r[3] >= 3 || r[3] < 0).reduce((t, r) => t + r[5], 0)};
  return {d: String(out.snapshot || '').slice(0, 10), t: String(out.snapshot || '').slice(11, 16), idle, stages, pm: out.paymatch || null, sum};
}
/** Pure: JSON ұяшыққа сыйғызу (Google Sheets шегі 50 000 таңба): көп болса — ең кіші жолдар алынады. */
function zopSnapCell_(rows, idx) { let r = rows.slice(); let j = JSON.stringify(r); while (j.length > 49000 && r.length) { r = r.sort((a, b) => b[idx] - a[idx]).slice(0, Math.floor(r.length * .85)); j = JSON.stringify(r); } return j; }
function zopSaveSnap_(out) {
  const s = zopSnapOf_(out); if (!/^\d{4}-\d{2}-\d{2}$/.test(s.d)) return;
  const ss = zopHome_(); let sh = ss.getSheetByName(ZOP_SNAP);
  if (!sh) { sh = ss.insertSheet(ZOP_SNAP); sh.getRange(1, 1, 1, ZOP_SNAP_HEAD.length).setValues([ZOP_SNAP_HEAD]).setFontWeight('bold').setBackground('#eef1f7'); sh.setFrozenRows(1); sh.getRange('A:B').setNumberFormat('@'); sh.getRange('A1').setNote('Автоматты: amoCRM күйі күн сайын (соңғы есептеу). Қолмен өзгертпеңіз.'); }
  const row = [s.d, s.t, zopSnapCell_(s.idle, 5), zopSnapCell_(s.stages, 3), JSON.stringify(s.pm), JSON.stringify(s.sum)];
  const last = sh.getLastRow(), dates = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(x => x[0]) : [], i = dates.lastIndexOf(s.d);
  sh.getRange(i >= 0 ? i + 2 : last + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function zopSnapGet_(day) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) return {error: 'snap-date'};
  const sh = zopHome_().getSheetByName(ZOP_SNAP); if (!sh || sh.getLastRow() < 2) return {d: day, none: true};
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, ZOP_SNAP_HEAD.length).getDisplayValues(), r = vals.filter(x => x[0] === day).pop();
  if (!r) return {d: day, none: true};
  const pj = t => { try { return JSON.parse(t); } catch (e) { return null; } };
  return {d: r[0], t: r[1], idle: pj(r[2]) || [], stages: pj(r[3]) || [], pm: pj(r[4]), sum: pj(r[5])};
}
/** Pure: «Облыс» → ресми өңір атауы (17 облыс + 3 қала). Жазылуы әр түрлі болса да (облысы / обл. / область / қаласы, орысша, латынша) бір атау.
    Тізімде жоқ мән — «Не указан» (бос мәтін сыртқа шықпайды). Python-да region_of — сол ереже. */
function zopRegion_(v) {
  const raw = String(v == null ? '' : v).trim(); if (!raw) return 'Не указан';
  const low = raw.toLowerCase().replace(/ё/g, 'е'), isObl = /облыс|обл\.|област|(^|\s)обл(\s|$)/.test(low);
  const t = low.replace(/облысы|облыс|область|обл\.?|қаласы|каласы|город|г\.|қ\./g, ' ').replace(/[^a-zа-яәғқңөұүһі]+/g, '');
  const R = [['астана', 'Астана'], ['нурсултан', 'Астана'], ['нұрсұлтан', 'Астана'], ['шымкент', 'Шымкент'], ['алматинск', 'Алматы обл.'], ['алматы', isObl ? 'Алматы обл.' : 'Алматы'],
    ['абай', 'Абай'], ['семей', 'Абай'], ['ақмола', 'Ақмола обл.'], ['акмол', 'Ақмола обл.'], ['көкшетау', 'Ақмола обл.'], ['кокшетау', 'Ақмола обл.'], ['ақтөбе', 'Ақтөбе'], ['актоб', 'Ақтөбе'], ['актюб', 'Ақтөбе'], ['атырау', 'Атырау'],
    ['батысқазақстан', 'Батыс Қазақстан'], ['батысказахстан', 'Батыс Қазақстан'], ['западноказахстан', 'Батыс Қазақстан'], ['бқо', 'Батыс Қазақстан'], ['зко', 'Батыс Қазақстан'], ['орал', 'Батыс Қазақстан'], ['уральск', 'Батыс Қазақстан'],
    ['жамбыл', 'Жамбыл'], ['тараз', 'Жамбыл'], ['жетісу', 'Жетісу'], ['жетысу', 'Жетісу'], ['талдықорған', 'Жетісу'], ['талдыкорган', 'Жетісу'], ['қарағанды', 'Қарағанды'], ['караганд', 'Қарағанды'],
    ['қостанай', 'Қостанай'], ['костанай', 'Қостанай'], ['қызылорда', 'Қызылорда'], ['кызылорд', 'Қызылорда'], ['маңғыстау', 'Маңғыстау'], ['мангист', 'Маңғыстау'], ['мангыст', 'Маңғыстау'], ['ақтау', 'Маңғыстау'], ['актау', 'Маңғыстау'],
    ['павлодар', 'Павлодар'], ['солтүстікқазақстан', 'Солтүстік Қазақстан'], ['североказахстан', 'Солтүстік Қазақстан'], ['сқо', 'Солтүстік Қазақстан'], ['ско', 'Солтүстік Қазақстан'], ['петропавл', 'Солтүстік Қазақстан'],
    ['түркістан', 'Түркістан'], ['туркестан', 'Түркістан'], ['ұлытау', 'Ұлытау'], ['улытау', 'Ұлытау'], ['жезқазған', 'Ұлытау'], ['жезказган', 'Ұлытау'],
    ['шығысқазақстан', 'Шығыс Қазақстан'], ['восточноказахстан', 'Шығыс Қазақстан'], ['шқо', 'Шығыс Қазақстан'], ['вко', 'Шығыс Қазақстан'], ['өскемен', 'Шығыс Қазақстан'], ['усткаменогорск', 'Шығыс Қазақстан'], ['устькаменогорск', 'Шығыс Қазақстан']];
  const hit = R.find(([k]) => t.indexOf(k) === 0) || R.find(([k]) => k.length >= 5 && t.indexOf(k) >= 0);
  return hit ? hit[1] : 'Не указан';
}
/* =================== Жауап уақыты (иесі 15.10: «мы не можем вычислить точное время?») ===================
   amoCRM жауап сәтін бермейді. Әр есептеуде рассылка тегі бойынша жауап санын (ОП-ға өткен + жауап этапы + Успешно) алдыңғымен салыстырамыз:
   өскені — CRM синхрондау уақытында келген жаңа жауаптар. Дәлдік = amoCRM → кесте синхрондауының жиілігі. Тек агрегат: тег, уақыт, сан. */
const ZOP_RT = 'Жауап уақыты', ZOP_RT_KEY = 'ZOP_RT_STATE', ZOP_RT_HEAD = ['Импорт тегі', 'Уақыты (CRM синхрондау)', 'Жаңа жауап'];
/** Pure: тег бойынша жауап саны — impgeo кубынан (жасыру импорт пен нәтижені сақтайды). */
function zopReplyCounts_(out) { const I = (out.dict || {}).imp || [], c = {}; (out.impgeo || []).forEach(r => { if (r[3] >= 1 && I[r[0]]) c[I[r[0]]] = (c[I[r[0]]] || 0) + r[r.length - 1]; }); return c; }
/** Pure: бұрын көрілген тег — өскені; жаңа тег (импорт уақыты жазу басталғаннан кейін) — бүкіл саны; ескі тег — тек бастапқы мән (ескі жауаптардың уақыты белгісіз). */
function zopReplyDelta_(state, cur, now) {
  const st = state && state.c ? {since: state.since, c: Object.assign({}, state.c)} : {since: now, c: {}}, add = [];
  Object.keys(cur).sort().forEach(t => { const prev = st.c[t], d = prev == null ? (t >= st.since ? cur[t] : 0) : cur[t] - prev; if (d > 0) add.push([t, now, d]); st.c[t] = Math.max(prev || 0, cur[t]); });
  Object.keys(st.c).sort().slice(0, Math.max(0, Object.keys(st.c).length - 400)).forEach(t => { delete st.c[t]; });   // ең ескі тегтер — шектеу (Properties 9 КБ)
  return {state: st, add};
}
function zopSaveReplies_(out) {
  const now = String(out.snapshot || '').slice(0, 16).replace('T', ' '); if (!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(now)) return;
  const P = PropertiesService.getScriptProperties(); let st = null; try { st = JSON.parse(P.getProperty(ZOP_RT_KEY) || 'null'); } catch (e) { st = null; }
  const r = zopReplyDelta_(st, zopReplyCounts_(out), now);
  if (r.add.length) {
    const ss = zopHome_(); let sh = ss.getSheetByName(ZOP_RT);
    if (!sh) { sh = ss.insertSheet(ZOP_RT); sh.getRange(1, 1, 1, 3).setValues([ZOP_RT_HEAD]).setFontWeight('bold').setBackground('#eef1f7'); sh.setFrozenRows(1); sh.getRange('A1').setNote('Автоматты: рассылка тегі бойынша жаңа жауаптар (CRM синхрондау уақыты). Қолмен өзгертпеңіз.'); }
    sh.getRange(sh.getLastRow() + 1, 1, r.add.length, 3).setNumberFormat('@').setValues(r.add.map(x => [x[0], x[1], String(x[2])]));
  }
  P.setProperty(ZOP_RT_KEY, JSON.stringify(r.state));
}
/** Pure: парақ жолдары → куб rt [импорт коды, күн индексі, сағат, минут, жаңа жауап]. */
function zopReplyCube_(vals, imps, base) {
  const b = Date.UTC(+base.slice(0, 4), +base.slice(5, 7) - 1, +base.slice(8, 10)), m = {};
  vals.forEach(([t, w, n]) => { const i = imps.indexOf(String(t)), s = String(w), v = +n; if (i < 0 || !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}/.test(s) || !(v > 0)) return;
    const k = [i, Math.round((Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) - b) / 864e5), +s.slice(11, 13), +s.slice(14, 16)].join(','); m[k] = (m[k] || 0) + v; });
  return Object.keys(m).sort().map(k => k.split(',').map(Number).concat([m[k]]));
}
function zopReadReplies_(out) {
  const sh = zopHome_().getSheetByName(ZOP_RT); if (!sh || sh.getLastRow() < 2) return [];
  return zopReplyCube_(sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues(), (out.dict || {}).imp || [], String(out.base));
}
/** Считает агрегаты и кладёт готовый ответ в кэш (кусками по 90 КБ — лимит кэша 100 КБ на ключ). */
function zopRefresh_() {
  const out = zopBuild_();
  try { zopSaveSnap_(out); } catch (e) { console.warn('Срез жазылмады: ' + e.message); }   // иесі 15.10: «срез жазып отыр, бәрін»
  try { zopSaveReplies_(out); out.rt = zopReadReplies_(out); } catch (e) { console.warn('Жауап уақыты: ' + e.message); }   // иесі 15.10: «точное время» ответа
  const payload = JSON.stringify(out);
  zopCachePut_('zop', payload);
  return payload;
}
/** Триггер каждые 10 минут: CRM и продажи (ошибка в продажах не мешает CRM). */
function zopTick_() { zopRefresh_(); try { const d = JSON.parse(zopSalesRefresh_()); try { zopWriteProductList_(d); } catch (e2) { console.warn('Өнім тізімі: ' + e2.message); } } catch (err) { console.warn('Продажи: ' + err.message); } }
function zopSalesRefresh_() {
  const out = zopSalesBuild_(), payload = JSON.stringify(out);
  // нет доступа к таблице платформы (скрипт запущен под другим аккаунтом) — не затираем прошлые полные данные
  if (!out.rows.platform && (out.errors || []).length) {
    const prev = zopCacheGet_('zops');
    if (prev) { try { const p = JSON.parse(prev); if (p.rows && p.rows.platform) { console.warn('Платформа не прочитана — оставлены данные от ' + p.snapshot);
      p.stale = {reason: out.errors.join('; '), attemptAt: out.snapshot, lastGoodAt: p.snapshot}; return JSON.stringify(p); return prev; } } catch (e) { /* пересчитать */ } }
  }
  zopCachePut_('zops', payload);
  return payload;
}
/** Бросает понятную ошибку, если этот аккаунт не видит таблицу платформы. */
function zopNeedPlatform_(d) {
  if (!d.rows || !d.rows.platform) throw new Error('Нет доступа к таблице платформы «All Trainings and Challanges» (' + (d.errors || []).join('; ') + '). Откройте Apps Script под аккаунтом, у которого есть доступ к этой таблице, и запустите снова.');
}
function zopCachePut_(pre, payload) {
  const parts = {}, size = 90000, n = Math.ceil(payload.length / size);
  for (let i = 0; i < n; i++) parts[pre + '-' + i] = payload.slice(i * size, (i + 1) * size);
  parts[pre + '-n'] = String(n);
  try { CacheService.getScriptCache().putAll(parts, 25 * 60); } catch (err) { /* без кэша тоже работает, просто медленнее */ }
}
function zopCacheGet_(pre) {
  pre = pre || 'zop';
  const c = CacheService.getScriptCache(), n = +(c.get(pre + '-n') || 0); if (!n) return null;
  const keys = Array.from({length: n}, (_, i) => pre + '-' + i), got = c.getAll(keys);
  return keys.every(k => got[k] != null) ? keys.map(k => got[k]).join('') : null;
}

function zopOut_(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function zopBuild_() {
  const ss = zopHome_(), tz = ss.getSpreadsheetTimeZone();
  const read = name => { const sh = ss.getSheetByName(name); return sh && sh.getLastRow() > 0 ? sh.getDataRange().getValues() : []; };
  // local date parts: the zone offset is read once (Kazakhstan has no daylight saving), then plain arithmetic —
  // calling Utilities.formatDate for ~24 000 dates made the web app too slow for Google's response window
  const z = Utilities.formatDate(new Date(), tz, 'Z'), off = (z[0] === '-' ? -1 : 1) * (+z.slice(1, 3) * 60 + +z.slice(3, 5)) * 60000;
  const parts = d => {
    if (!(d instanceof Date) || isNaN(d.getTime())) return null;
    const x = new Date(d.getTime() + off);
    return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours(), min: x.getUTCMinutes(), wd: (x.getUTCDay() + 6) % 7};
  };
  const L = zopCrmRead_('LEADS'), P = zopCrmRead_('PIPES'), G = zopCrmRead_('LOG');
  if (!L.src) throw new Error('amoCRM «Сделки» оқылмады: ' + L.err);   // нөлдер көрсетілмейді: кэштегі соңғы дұрыс жауап қалады
  let payRows = []; if (ZOP_USE_OP_SHEET) try { payRows = zopSourceRows_('sales', []).rows || []; } catch (e) { console.warn('ОП кестесі (сәйкестендіру): ' + e.message); }   // тек сделка «төленді» белгісі үшін — телефон/чек сыртқа шықпайды
  const out = zopAggregate_(L.rows, P.rows, G.rows, parts, read(ZOP.MAIL), read(ZOP.COST), read(ZOP.KPI), read(ZOP_TARIFF), read(ZOP_HYP), payRows);
  out.crmSource = {leads: L.src, pipes: P.src, log: G.src, error: L.src ? null : L.err};   // shown on «Деректер сапасы»
  return out;
}

/** Создаёт лист «Расходы» (если его нет) и заполняет известными платежами сервиса рассылок. Другие листы не трогает. */
function zopCreateCostSheet_() {
  const ss = zopHome_();
  if (ss.getSheetByName(ZOP.COST)) return 0;
  const sh = ss.insertSheet(ZOP.COST);
  sh.getRange(1, 1, 1, ZOP_COST_HEAD.length).setValues([ZOP_COST_HEAD]).setFontWeight('bold').setBackground('#eef1f7');
  const rows = ZOP_COST_SEED.map(r => { const m = r[0].match(/(\d{2})\.(\d{2})\.(\d{4})/); return [new Date(+m[3], +m[2] - 1, +m[1]), r[1], r[2], +r[3], r[4], r[5]]; });
  sh.getRange(2, 1, rows.length, ZOP_COST_HEAD.length).setValues(rows);
  sh.setFrozenRows(1); sh.getRange('A:A').setNumberFormat('dd.mm.yyyy'); sh.getRange('D:D').setNumberFormat('#,##0');
  [110, 130, 300, 120, 110, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  return rows.length;
}

/**
 * Создаёт лист «Рассылки» (если его ещё нет) и заполняет найденные рассылки: дни, когда в воронке рассылки
 * массово создавались сделки (40+ в день, соседние дни — одна рассылка). Тему и расход вписываете вы.
 * Другие листы не трогает.
 */
function zopCreateMailSheet_() {
  const ss = zopHome_();
  if (ss.getSheetByName(ZOP.MAIL)) return 0;
  const leads = zopCrmRead_('LEADS').rows; if (!leads.length) return 0;
  const head = leads[0].map(String), ix = {};
  head.forEach((h, i) => { ix[h] = i; });
  const tz = ss.getSpreadsheetTimeZone(), key = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const FUN = ['Рассылка', 'Рассылка 5-6 класс Алтын сақа'], byDay = {}, tags = {};
  leads.slice(1).forEach(r => {
    const f = r[ix['Воронка']], c = r[ix['Создана']]; if (FUN.indexOf(f) < 0 || !(c instanceof Date)) return;
    const k = f + '|' + key(c); byDay[k] = (byDay[k] || 0) + 1;
    String(r[ix['Теги']] || '').split(',').map(t => t.trim()).filter(t => t && !/\d{5,}/.test(t)).forEach(t => { tags[k] = tags[k] || {}; tags[k][t] = (tags[k][t] || 0) + 1; });
  });
  const waves = [];
  FUN.forEach(f => {
    const days = Object.keys(byDay).filter(k => k.startsWith(f + '|') && byDay[k] >= 40).map(k => k.split('|')[1]).sort();
    days.forEach(d => { const w = waves[waves.length - 1], prev = w && w.f === f ? new Date(w.b + 'T00:00:00') : null, cur = new Date(d + 'T00:00:00');
      if (prev && (cur - prev) / 864e5 <= 1) { w.b = d; w.days.push(d); } else waves.push({f, a: d, b: d, days: [d]}); });
  });
  waves.sort((x, y) => x.a < y.a ? -1 : 1);
  const rows = waves.map((w, i) => {
    const n = w.days.reduce((s, d) => s + byDay[w.f + '|' + d], 0), t = {};
    w.days.forEach(d => Object.entries(tags[w.f + '|' + d] || {}).forEach(([k, v]) => { t[k] = (t[k] || 0) + v; }));
    const seg = Object.entries(t).sort((a, b) => b[1] - a[1]).slice(0, 4).map(e => e[0]).join(', ');
    return ['Рассылка ' + (i + 1), '', new Date(w.a + 'T00:00:00'), w.a === w.b ? '' : new Date(w.b + 'T00:00:00'), w.f, seg, '', n, 'авто: ' + n + ' сделок создано'];
  });
  const sh = ss.insertSheet(ZOP.MAIL);
  sh.getRange(1, 1, 1, ZOP_MAIL_HEAD.length).setValues([ZOP_MAIL_HEAD]).setFontWeight('bold').setBackground('#eef1f7');
  if (rows.length) sh.getRange(2, 1, rows.length, ZOP_MAIL_HEAD.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange('C:D').setNumberFormat('dd.mm.yyyy'); sh.getRange('G:H').setNumberFormat('#,##0');
  [150, 260, 110, 120, 230, 240, 120, 120, 220].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('B2:B').setBackground('#fff8e1'); sh.getRange('G2:G').setBackground('#fff8e1');   // поля для заполнения
  return rows.length;
}

/* =================== KPI: цели на месяц из листа «KPI» =================== */
const ZOP_KPI_HEAD = ['Ай', 'Бағыт', 'Көрсеткіш', 'Мақсат', 'Арна (міндетті емес)', 'Акция (міндетті емес)', 'Өнім (міндетті емес)', 'Менеджер (міндетті емес)'];
/** Создаёт лист «KPI» в этой таблице (если его нет): заголовки, выпадающие списки, подсказки. Другие листы не трогает. */
function zopCreateKpiSheet_() {
  const ss = zopHome_();
  if (ss.getSheetByName(ZOP.KPI)) return 0;
  const sh = ss.insertSheet(ZOP.KPI), n = 200;
  sh.getRange(1, 1, 1, ZOP_KPI_HEAD.length).setValues([ZOP_KPI_HEAD]).setFontWeight('bold').setBackground('#eef1f7');
  sh.setFrozenRows(1); [110, 130, 130, 110, 150, 150, 230, 170].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const list = (col, vals) => sh.getRange(2, col, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(true).build());
  list(2, ['Маркетинг', 'Рассылка', 'ОП', 'Өнімдер']); list(3, ['Лидтер', 'Сатылым', 'Түсім', 'Конверсия']);
  sh.getRange('A2:A').setNumberFormat('@'); sh.getRange('D2:D').setNumberFormat('#,##0.##').setBackground('#fff8e1');
  sh.getRange('A1').setNote('Ай — мысалы 2026-10 (жыл-ай).');
  sh.getRange('C1').setNote('Лидтер — жаңа мәмілелер; Сатылым — сәтті жабылған; Түсім — сатылым сомасы, ₸; Конверсия — сатылым ÷ жабылған, % (мысалы 30).');
  sh.getRange('E1').setNote('Бос қалса — барлығы. Арна: Рассылка, WhatsApp, Таргет, Instagram, Сарафан. Акция: Предпродажа, 2+1 Акция, Жеңілдік 15%. Өнім мен менеджер — amoCRM-дегідей.');
  return 1;
}
/* Pure: rows of the «KPI» sheet → [{month:'YYYY-MM', area, metric, target, src, of, pr, m}] (raw words; the dashboard maps them) */
function zopKpi_(rows, parts) {
  if (!rows.length) return [];
  const h = rows[0].map(x => String(x || '').trim().toLowerCase()), c = re => h.findIndex(x => re.test(x));
  const C = {mo: c(/^ай/), a: c(/^бағыт/), k: c(/^көрсеткіш/), t: c(/^мақсат/), src: c(/^арна/), of: c(/^акция/), pr: c(/^өнім/), m: c(/^менеджер/)};
  const pad = n => String(n).padStart(2, '0'), str = (r, k) => C[k] >= 0 && r[C[k]] != null ? String(r[C[k]]).trim() : '';
  const month = v => {
    if (v instanceof Date) { const p = parts(v); return p ? `${p.y}-${pad(p.m)}` : ''; }
    const t = String(v == null ? '' : v).trim(); let m = t.match(/^(\d{4})[-./](\d{1,2})/); if (m) return `${m[1]}-${pad(+m[2])}`;
    m = t.match(/^(?:\d{1,2}[./])?(\d{1,2})[./](\d{4})$/); if (m) return `${m[2]}-${pad(+m[1])}`;
    return '';
  };
  const num = v => { if (typeof v === 'number') return v; const x = +String(v == null ? '' : v).replace(/[^\d.,-]/g, '').replace(',', '.'); return isFinite(x) && String(v).trim() !== '' ? x : null; };
  return rows.slice(1).map(r => ({month: C.mo >= 0 ? month(r[C.mo]) : '', area: str(r, 'a'), metric: str(r, 'k'), target: C.t >= 0 ? num(r[C.t]) : null,
    src: str(r, 'src'), of: str(r, 'of'), pr: str(r, 'pr'), m: str(r, 'm')})).filter(k => k.month && k.area && k.metric && k.target != null && k.target > 0);
}

/* «Жеңілдік» → акция: названия из amoCRM как есть; процент (0,15 в ячейке с форматом %, 15 или «15%») → «Жеңілдік 15%»; пусто → «Без акции»; прочее → «Басқа акция» */
function zopOffer_(v, OFFERS) {
  if (v === null || v === undefined || String(v).trim() === '') return 'Без акции';
  if (OFFERS[v]) return OFFERS[v];
  let x = typeof v === 'number' ? v : null;
  if (x === null) { const m = String(v).trim().match(/^(\d+(?:[.,]\d+)?)\s*%?$/); if (m) x = +m[1].replace(',', '.'); }
  if (x !== null && x > 0) { const pct = x < 1 ? Math.round(x * 1000) / 10 : x; if (pct > 0 && pct <= 100) return 'Жеңілдік ' + String(pct).replace('.', ',') + '%'; }
  return 'Басқа акция';
}

/* «Теги» сделки → рассылка по тегу импорта (решение владельца 13.10.2026: рассылки — по тегам, не по воронке).
   «импорт_ДДММГГГГ_ЧЧММ» (amoCRM ставит при загрузке) → 'YYYY-MM-DD HH:MM'; остальные короткие теги без длинных чисел — сегменты
   (Алтын сақа, Астана қ., 3 сынып…). Теги с 5+ цифрами (телефоны, WZ (номер)) не выходят наружу. */
function zopImportTags_(v) {
  const imps = [], segs = [];
  String(v == null ? '' : v).split(',').map(t => t.trim()).filter(Boolean).forEach(t => {
    const m = t.match(/^импорт_(\d{2})(\d{2})(\d{4})_(\d{2})(\d{2})$/i);
    if (m) { const k = m[3] + '-' + m[2] + '-' + m[1] + ' ' + m[4] + ':' + m[5]; if (imps.indexOf(k) < 0) imps.push(k); return; }
    if (/\d{5,}/.test(t) || /^импорт_/i.test(t) || /^wz\b/i.test(t) || t.length > 40) return;
    if (segs.indexOf(t) < 0) segs.push(t);
  });
  return {imps, segs};
}

/* Pure aggregation, mirrors инструменты/обновить_данные.py exactly (same dictionaries, order and cubes). */
/* Бюджет из таблицы: число, или текст вида «12 000», «12000,50 ₸». Пусто → 0 (нет бюджета). Нечитаемое → null (считается отдельно). */
function zopBudget_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (v === null || v === undefined) return 0;
  const t = String(v).trim(); if (!t) return 0;
  const c = t.replace(/[\s ₸тгTKZkzt]/gi, '').replace(/,(?=\d{1,2}$)/, '.').replace(/,/g, '');
  return /^-?\d+(\.\d+)?$/.test(c) ? +c : null;
}
function zopAggregate_(leads, pipes, logRows, parts, mailRows, costRows, kpiRows, tariffRows, hypRows, payRows) {
  const FUNNELS = ['ОП', 'Рассылка', 'Рассылка 5-6 класс Алтын сақа', 'Техподдержка'];
  const OUTCOMES = ['В работе', 'Успешно', 'Не реализовано'];
  const OFFERS = {'Предпродажа': 'Предпродажа', '2+1 Акция': '2+1 Акция'};
  const SOURCES = {'Рассылка': 'Рассылка', 'Органика WhatsApp': 'WhatsApp', 'Таргетированная реклама в Инстаграм или Фейсбук': 'Таргет',
    'Органика Instagram': 'Instagram', 'Рекомендация (Сарафан)': 'Сарафан', 'Самост. обращение': 'Самост. обращение'};
  const BASE = Date.UTC(2026, 7, 10);
  const dayIdx = p => p ? Math.round((Date.UTC(p.y, p.m - 1, p.d) - BASE) / 864e5) : -1;
  const iso = p => `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}T${String(p.h).padStart(2, '0')}:${String(p.min).padStart(2, '0')}`;
  const nonEmpty = v => v !== null && v !== undefined && v !== '';

  // stage order
  const stages = {}; FUNNELS.forEach(f => { stages[f] = []; });
  pipes.slice(1).forEach(r => { if (r && stages[r[1]]) stages[r[1]].push({name: r[4], order: r[5]}); });
  FUNNELS.forEach(f => stages[f].sort((a, b) => (a.order - b.order) || (String(a.name) < String(b.name) ? -1 : String(a.name) > String(b.name) ? 1 : 0)));

  const idx = {}, dictOrder = [];
  const code = (dim, v) => { if (!idx[dim]) { idx[dim] = new Map(); dictOrder.push(dim); } const m = idx[dim]; if (!m.has(v)) m.set(v, m.size); return m.get(v); };
  FUNNELS.forEach(f => code('funnel', f)); OUTCOMES.forEach(o => code('outcome', o));

  const head = (leads[0] || []).map(String), ix = {}; head.forEach((h, i) => { ix[h] = i; });
  const g = (r, k) => (ix[k] === undefined ? null : r[ix[k]]), idCol = ['ID', 'ID сделки', 'Id', 'id', 'Сделка ID'].find(h => ix[h] !== undefined) || null;
  const dh = new Map();   // күн × сағат: [0 құрылды | 1 «Успешно», күн, сағат, воронка, «Қайдан келді» = Рассылка] — рассылканың нақты уақыты үшін
  const dhi = new Map(), imppay = new Map();   // imppay: [импорт коды] → төленген сделка саны, төлем сомасы (ОП кестесі → сделка → тег; ТЗ 15.10 «1 000-ға»)
  const _dhiNote = 0;   // рассылка тегі бойынша сатып алу уақыты: [импорт коды, жабылған күн, сағат] — карточкадағы «Уақыт» тек осы рассылканың сатылымы (14.10.2026)
  const imp = new Map(), impseg = new Map();   // рассылка по тегу импорта: [импорт, день импорта, воронка, этап, итог, менеджер, день закрытия, продукт] · сегменты импорта
  const impgeo = new Map(), impstg = new Map();   // impstg: [өңір, этап, нәтиже] — тегі бар сделка бір рет   // рассылка × облыс × оқу тілі: [импорт, облыс, тіл, нәтиже 0 жауап жоқ · 1 жауап берді · 2 CRM «Успешно»] — карточкадағы облыс/тіл графиктері үшін
  const MAIL_REPLY = /жауап|первичный контакт|принимают решение/i;   // Атластағы MAIL_REPLY-мен бірдей
  const fill = new Map(), main = new Map(), hours = new Map(), buy = new Map(), prof = new Map(), supp = new Map(), reasons = new Map(), idle = new Map(), idleRaw = [], rsrc = new Map(), hsrc = new Map(), missRaw = [];
  const reg = new Map(), clsd = new Map();   // [воронка, нәтиже, өңір | сынып, құрылған күн] — «Аудитория» (15.10)
  const add = (m, key, n, b) => { const k = key.join(','); const e = m.get(k) || {key, n: 0, b: 0}; e.n += n; e.b += b || 0; m.set(k, e); };
  let nRows = 0, lastChange = null, lastChangeP = null, badBudget = 0;
  const PM = zopPayMatch_(leads, payRows || []);   // ОП кестесіндегі төлем → сделка (телефон «Доступ номер» / «Чек»); тек «төленді» белгісі шығады

  // иесі 15.10: «7000 — серверсіз»: 7000 сервері сделкаға utm жазады, amoCRM синхроны барлық өрісті «Сделки»-ге көшіреді → креатив × өтінім × CRM Успешно осы жерден
  const utmCol = k => Object.keys(ix).find(h => h.toLowerCase().replace(/[\s_-]/g, '') === k) || '', UC = {content: utmCol('utmcontent'), campaign: utmCol('utmcampaign'), source: utmCol('utmsource')};
  const utv = h => { if (!h) return ''; const t = String(g(rCur, h) == null ? '' : g(rCur, h)).trim().slice(0, 80); return !t || /\d{7,}/.test(t) ? '' : t; };   // ұзын сан (телефон сияқты) — шықпайды
  let rCur = null; const utm = new Map();   // [құрылған күн, кампания, креатив, дереккөз, нәтиже]
  leads.slice(1).forEach(r => {
    if (!r.some(nonEmpty)) return;
    const f = g(r, 'Воронка'); if (!stages[f]) return; rCur = r;
    nRows++;
    const pc = parts(g(r, 'Создана')), pcl = parts(g(r, 'Закрыта')), changed = g(r, 'Изменена');
    if (changed instanceof Date && !isNaN(changed.getTime()) && (lastChange === null || changed > lastChange)) { lastChange = changed; lastChangeP = parts(changed); }
    const fu = code('funnel', f), st = code('stage', f + '|' + String(g(r, 'Этап')));
    const ocv = g(r, 'Итог') || 'В работе', oc = code('outcome', ocv), mg = code('manager', (n => ZOP_MAN_ALIAS[n] || n)(String(g(r, 'Ответственный') || 'Не назначен').trim()));
    const rid = g(r, 'Причина отказа ID'), reason = nonEmpty(rid) ? (ZOP_REASONS[Math.trunc(+rid)] || 'Причина ' + Math.trunc(+rid)) : g(r, 'Причина отказа');
    const srcRaw = String(g(r, 'Қайдан келді') == null ? '' : g(r, 'Қайдан келді')).trim(), so = code('source', SOURCES[srcRaw] || srcRaw.slice(0, 60) || 'Не указан'),   // белгісіз мән де өз атымен (бұрын «Не указан» болып кететін; иесі, 14.10.2026)
    b0 = ix['Бюджет CRM'] !== undefined ? g(r, 'Бюджет CRM') : g(r, 'Бюджет');
    const bud = zopBudget_(b0); if (bud === null && nonEmpty(b0)) badBudget++;   // text «12 000» is parsed; an unreadable value is counted, not silently 0
    const of = code('offer', zopOffer_(g(r, 'Жеңілдік'), OFFERS));
    const prd = g(r, 'Продукт'), pr = code('product', nonEmpty(prd) ? String(prd).trim() : 'Не указан');
    const tgs = code('tagset', zopImportTags_(g(r, 'Теги')).segs.slice().sort().join('|'));   // amoCRM тегтері (5+ цифрлы, импорт, WZ — жоқ) — сүзгі үшін (иесі, 14.10.2026)
    add(main, [dayIdx(pc), fu, st, oc, mg, so, dayIdx(pcl), nonEmpty(reason) ? 1 : 0, of, pr, tgs, PM.paid.has(r) ? 1 : 0], 1, bud || 0);
    if (f !== 'Техподдержка') { const tg = zopImportTags_(g(r, 'Теги'));
      tg.imps.forEach(k => { const ic = code('imp', k); add(imp, [ic, dayIdx({y: +k.slice(0, 4), m: +k.slice(5, 7), d: +k.slice(8, 10)}), fu, st, oc, mg, dayIdx(pcl), pr], 1, bud || 0);
        tg.segs.forEach(sg => add(impseg, [ic, code('tagseg', sg)], 1));
        if (ocv === OUTCOMES[1] && pcl) add(dhi, [ic, dayIdx(pcl), pcl.h], 1);
        if (PM.sum.has(r)) add(imppay, [ic], 1, PM.sum.get(r)); }); }
    // для контроля заполнения ОП: закрытые сделки — есть ли сумма, продукт, «Қайдан келді», причина отказа
    // open commercial deals: how long without any change (by «Изменена»; empty → «Создана»), and the next task — counted after the snapshot time is known
    if (ocv === OUTCOMES[0] && f !== 'Техподдержка') { const ch = changed instanceof Date && !isNaN(changed.getTime()) ? changed : g(r, 'Создана'), tk = g(r, 'Ближайшая задача');
      idleRaw.push([mg, fu, st, ch instanceof Date && !isNaN(ch.getTime()) ? ch.getTime() : null, tk instanceof Date && !isNaN(tk.getTime()) ? tk.getTime() : null, idCol ? +String(g(r, idCol) == null ? '' : g(r, idCol)).replace(/\D/g, '') || 0 : 0, f === 'ОП' ? 1 : 0]); }
    // miss: жабылған сделканың бос өрістері — тек сделка ID (аты-жөнсіз), менеджер, күн, итог, белгі: 1 бюджет · 2 продукт · 4 «Қайдан келді» · 8 бас тарту себебі (ОП) — «кім не толтырмады» (иесі, 14.10.2026)
    if (pcl && f !== 'Техподдержка') { const won = ocv === OUTCOMES[1], lostOp = ocv === OUTCOMES[2] && f === 'ОП', lid = idCol ? String(g(r, idCol) == null ? '' : g(r, idCol)).replace(/\D/g, '') : '';
      const fl = won ? ((bud != null && bud > 0 ? 0 : 1) | (nonEmpty(prd) ? 0 : 2) | (nonEmpty(g(r, 'Қайдан келді')) ? 0 : 4) | (nonEmpty(g(r, 'Язык обучения')) ? 0 : 16) | (nonEmpty(g(r, 'Чек')) ? 0 : 32)) : lostOp ? (nonEmpty(reason) ? 0 : 8) : 0;   // 16 оқу тілі · 32 чек (тек толтырылғаны, мәні шықпайды)
      if (fl && lid) missRaw.push([+lid, dayIdx(pcl), mg, won ? 0 : 1, fl]); }
    if (pcl) add(fill, [dayIdx(pcl), fu, mg, oc, bud != null && bud > 0 ? 1 : 0, nonEmpty(prd) ? 1 : 0, nonEmpty(g(r, 'Қайдан келді')) ? 1 : 0, nonEmpty(reason) ? 1 : 0], 1, bud || 0);
    if (pc) { add(hours, [pc.wd, pc.h, fu], 1); add(hsrc, [pc.wd, pc.h, fu, so], 1); }   // hsrc — арна бойынша (маркетинг, 14.10.2026)
    const mfl = SOURCES[g(r, 'Қайдан келді')] === 'Рассылка' ? 1 : 0;
    if (pc) add(dh, [0, dayIdx(pc), pc.h, fu, mfl], 1);
    if (ocv === OUTCOMES[1] && pcl) add(dh, [1, dayIdx(pcl), pcl.h, fu, mfl], 1, bud || 0);
    if (ocv === OUTCOMES[1] && pcl) add(buy, [pcl.wd, pcl.h, fu, so], 1);
    let cls = g(r, 'Класс'); cls = typeof cls === 'number' && cls > 0 && cls < 13 ? Math.trunc(cls) : 0;
    let pk = g(r, 'Пакет'); pk = pk === 'AltynSaqa' ? 'AltynSaqa' : (nonEmpty(pk) ? 'Другое' : 'Не указан');
    const lang = ['Қазақша', 'Русский'].indexOf(g(r, 'Язык обучения')) >= 0 ? g(r, 'Язык обучения') : 'Не указан';
    const cl = {'Жаңа клиент': 'Новый', 'Ескі клиент': 'Повторный'}[g(r, 'Клиент')] || 'Не указан';
    add(reg, [fu, oc, code('region', zopRegion_(g(r, 'Облыс'))), dayIdx(pc)], 1); add(clsd, [fu, oc, cls, dayIdx(pc)], 1);   // өңір / сынып × сделка құрылған күн — «Аудитория» кезеңге бағынады (иесі 15.10)
    add(prof, [fu, oc, code('region', zopRegion_(g(r, 'Облыс'))), cls, code('package', pk), code('lang', lang), code('client', cl), of, so], 1);
    if (f !== 'Техподдержка') { const res = ocv === OUTCOMES[1] ? 2 : f === 'ОП' || MAIL_REPLY.test(String(g(r, 'Этап'))) ? 1 : 0, rg = code('region', zopRegion_(g(r, 'Облыс'))), lg = code('lang', lang);
      const ims = zopImportTags_(g(r, 'Теги')).imps; ims.forEach(k => add(impgeo, [code('imp', k), rg, lg, res], 1)); if (ims.length) add(impstg, [rg, st, oc], 1); }   // impstg: рассылка сделкасы өңір × қазіргі этап (иесі 15.10: «кім қандай этапта»)
    if (f !== 'Техподдержка') { const uc = utv(UC.campaign), ux = utv(UC.content), us = utv(UC.source); if (uc || ux || us) add(utm, [dayIdx(pc), code('utmc', uc || '—'), code('utmx', ux || '—'), code('utms', us || '—'), oc], 1); }
    if (f === 'Техподдержка') add(supp, [dayIdx(pc), code('parea', g(r, 'problem_area') || 'Не указано'), code('ptype', g(r, 'problem_type') || 'Не указано'), oc], 1);
    if (nonEmpty(reason)) { const rc = code('reason', String(reason)); add(reasons, [fu, rc], 1); add(rsrc, [fu, so, rc], 1); }   // rsrc — себеп арна бойынша (маркетинг, 14.10.2026)
  });

  // privacy: a profile cell of 1–2 children with one region/class/language combination could point to a specific child →
  // its quasi-identifiers are masked (region «Жасырылған (< 3)», class 0, package/language/client unknown); totals by funnel/outcome/offer/source are kept
  const PROF_MIN = 3, profSafe = new Map(), regSafe = new Map();
  const regTot = {}, clsTot = {}; reg.forEach(e => { regTot[e.key[2]] = (regTot[e.key[2]] || 0) + e.n; }); clsd.forEach(e => { clsTot[e.key[2]] = (clsTot[e.key[2]] || 0) + e.n; });   // жасыру — өңір/сынып бүкіл кезеңде 3-тен аз болса
  const clsSafe = new Map(); clsd.forEach(e => { const k = clsTot[e.key[2]] >= PROF_MIN ? e.key : [e.key[0], e.key[1], -1, e.key[3]], id = k.join(','), x = clsSafe.get(id) || {key: k, n: 0, b: 0}; x.n += e.n; clsSafe.set(id, x); });
  reg.forEach(e => { const k = regTot[e.key[2]] >= PROF_MIN ? e.key : [e.key[0], e.key[1], code('region', 'Жасырылған (< 3)'), e.key[3]], id = k.join(','), x = regSafe.get(id) || {key: k, n: 0, b: 0}; x.n += e.n; regSafe.set(id, x); });   // өңір графигі: тек өңір × нәтиже, 3-тен аз — жасырылады (15.10)
  prof.forEach(e => { if (e.n >= PROF_MIN) { profSafe.set(e.key.join(','), e); return; }
    const k = [e.key[0], e.key[1], code('region', 'Жасырылған (< 3)'), 0, code('package', 'Не указан'), code('lang', 'Не указан'), code('client', 'Не указан'), e.key[7], e.key[8]], id = k.join(',');
    const x = profSafe.get(id) || {key: k, n: 0, b: 0}; x.n += e.n; profSafe.set(id, x); });
  // сегмент импорта меньше 3 сделок → «Басқа (< 3)» (как профиль)
  const impgeoSafe = new Map();   // 1–2 сделка бір облыс/тіл/нәтижеде → облыс «Жасырылған (< 3)», тіл «Не указан» (профильдегідей)
  impgeo.forEach(e => { const k = e.n >= PROF_MIN ? e.key : [e.key[0], code('region', 'Жасырылған (< 3)'), code('lang', 'Не указан'), e.key[3]], id = k.join(','), x = impgeoSafe.get(id) || {key: k, n: 0, b: 0}; x.n += e.n; impgeoSafe.set(id, x); });
  const tagTot = {}, tagRep = {}; impgeo.forEach(e => { const t = e.key[0]; tagTot[t] = (tagTot[t] || 0) + e.n; if (e.key[3] >= 1) tagRep[t] = (tagRep[t] || 0) + e.n; });
  const impstgSafe = new Map(); impstg.forEach(e => { const k = e.n >= PROF_MIN ? e.key : [code('region', 'Жасырылған (< 3)'), e.key[1], e.key[2]], id = k.join(','), x = impstgSafe.get(id) || {key: k, n: 0, b: 0}; x.n += e.n; impstgSafe.set(id, x); });
  const impreg = new Map(), impregSafe = new Map();   // [өңір, нәтиже 0/1/2, тікелей ОП-ға импорт 0/1] — «Қай өңірге жіберген тиімді» (15.10); жасырылады тек 3-тен аз ұяшықтың өңірі
  impgeo.forEach(e => add(impreg, [e.key[1], e.key[3], tagRep[e.key[0]] / tagTot[e.key[0]] >= .9 ? 1 : 0], e.n));
  impreg.forEach(e => { const k = e.n >= PROF_MIN ? e.key : [code('region', 'Жасырылған (< 3)'), e.key[1], e.key[2]], id = k.join(','), x = impregSafe.get(id) || {key: k, n: 0, b: 0}; x.n += e.n; impregSafe.set(id, x); });
  const impsegSafe = new Map();
  impseg.forEach(e => { const k = e.n >= PROF_MIN ? e.key : [e.key[0], code('tagseg', 'Басқа (< 3)')], id = k.join(','), x = impsegSafe.get(id) || {key: k, n: 0, b: 0}; x.n += e.n; impsegSafe.set(id, x); });
  const cmp = (a, b) => { for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return a[i] - b[i]; return 0; };
  const rows = (m, withB) => [...m.values()].sort((x, y) => cmp(x.key, y.key)).map(e => withB ? e.key.concat([e.n, Math.round(e.b)]) : e.key.concat([e.n]));
  // журнал синхронизации: «Время» (+ «Режим» в старом листе «Журнал»; в «CRM Журнал — авто» режима нет)
  const lh = (logRows[0] || []).map(String), lt = Math.max(0, lh.indexOf('Время')), lm = lh.indexOf('Режим');
  const log = logRows.slice(1).filter(r => r && r[lt] instanceof Date).map(r => [r[lt], lm >= 0 ? r[lm] : '']).sort((a, b) => a[0] - b[0]);
  // idle: [manager, funnel, stage, days without change at the snapshot (60 = 60+; -1 = no date), task: 0 none · 1 planned · 2 overdue, n]
  const tariffs = zopTariffs_(tariffRows || [], parts, dayIdx);
  const snapMs = log.length ? log[log.length - 1][0].getTime() : null;
  const act = [];   // «Әрекет керек» (ревью 15.10): ОП-дағы ашық сделкалар — 3+ күн өзгеріссіз немесе тапсырма мерзімі өткен: [сделка ID, менеджер, этап, күн, тапсырма 2 = мерзімі өткен] — аты-жөнсіз
  if (snapMs != null) idleRaw.forEach(([mg, fu, st, ch, tk, lid, op]) => { if (!op || !lid) return; const dd = ch == null ? -1 : Math.min(60, Math.max(0, Math.floor((snapMs - ch) / 864e5))), late = tk != null && tk < snapMs; if (dd >= 3 || late) act.push([lid, mg, st, dd, late ? 2 : tk == null ? 0 : 1]); });
  if (snapMs != null) idleRaw.forEach(([mg, fu, st, ch, tk]) => add(idle, [mg, fu, st, ch == null ? -1 : Math.min(60, Math.max(0, Math.floor((snapMs - ch) / 864e5))), tk == null ? 0 : tk < snapMs ? 2 : 1], 1));
  const dict = {}; dictOrder.forEach(k => { dict[k] = [...idx[k].keys()]; });
  return {
    base: '2026-08-10', schemaVersion: 2, methodVersion: 'crm-2026-10-10', rows: nRows, quality: {badBudget},
    snapshot: log.length ? iso(parts(log[log.length - 1][0])) : null,
    lastChange: lastChangeP ? iso(lastChangeP) : null,
    sync: {runs: log.length, full: log.filter(r => r[1] === 'Полная').length, first: log.length ? iso(parts(log[0][0])) : null,
      gaps: log.slice(1).map((r, i) => Math.round((r[0] - log[i][0]) / 60000))},
    dict, stages: Object.fromEntries(FUNNELS.map(f => [f, stages[f].map(s => s.name)])),
    main: rows(main, true), imp: rows(imp, true), impseg: rows(impsegSafe), impgeo: rows(impgeoSafe), impreg: rows(impregSafe), impstg: rows(impstgSafe), clsd: rows(clsSafe), utm: rows(utm), utmCols: UC, fill: rows(fill, true), idle: rows(idle), hours: rows(hours), dh: rows(dh), dhi: rows(dhi), reg: rows(regSafe), act: act.sort((a, b) => b[3] - a[3] || b[0] - a[0]).slice(0, 3000), imppay: rows(imppay, true), buy: rows(buy), prof: rows(profSafe), supp: rows(supp), reasons: rows(reasons), rsrc: rows(rsrc), hsrc: rows(hsrc), miss: zopMissRows_(missRaw), paymatch: PM.stats,
    mailings: zopMailings_(mailRows || [], parts, dayIdx, FUNNELS),
    costs: zopCosts_(costRows || [], parts, dayIdx, tariffs.list), tariffs: tariffs.list, tariffWarn: tariffs.warn, hypotheses: zopHypotheses_(hypRows || []),
    kpi: zopKpi_(kpiRows || [], parts),
    source: 'sheets',
  };
}
/* Лист «Рассылки» → список рассылок (только данные, которые вы вносите сами: тема, даты, расход). */
function zopMailings_(rows, parts, dayIdx, FUNNELS) {
  if (!rows.length) return [];
  const head = rows[0].map(x => String(x || '').trim().toLowerCase()), col = re => head.findIndex(h => re.test(h));
  const c = {product: col(/өнім|продукт|product/), goal: col(/мақсат|цель|goal/), offer: col(/ұсыныс|предложен|offer/), planned: col(/жоспар|план|planned/), tariff: col(/тариф/), status: col(/статус|status/), limits: col(/шектеу|огранич|limit/),
    text: col(/мәтін|текст/), media: col(/медиа|сурет|видео|media/), btn: col(/батырма|кнопк|button/), fmt: col(/формат|format/), num: col(/нөмір|^номер|number/), hyp: col(/гипотез|hypothes/),
    name: col(/^рассылка|назв/), topic: col(/^тема/), time: col(/уақыт|время/), cls: col(/^сынып|^класс/), lang: col(/тіл тоб|^тіл$|язык/), a: col(/начал|^дата$/), b: col(/оконч/), f: col(/воронк/), seg: col(/кому|сегмент|аудитор/), cost: col(/расход|затрат|сумм/), sent: col(/отправ|охват|факт/), note: col(/коммент/)};
  const num = v => { if (typeof v === 'number') return v; const t = String(v || '').replace(/[^\d.,-]/g, '').replace(',', '.'); return t ? +t : null; };
  const day = v => { if (v instanceof Date) return dayIdx(parts(v)); const m = String(v || '').match(/(\d{1,2})[./](\d{1,2})[./](\d{4})/); return m ? dayIdx({y: +m[3], m: +m[2], d: +m[1]}) : null; };
  const fun = v => { const t = String(v || '').toLowerCase(); return /5\s*[-–]\s*6/.test(t) ? FUNNELS.indexOf('Рассылка 5-6 класс Алтын сақа') : /рассыл/.test(t) ? FUNNELS.indexOf('Рассылка') : -1; };
  const get = (r, k) => c[k] >= 0 ? r[c[k]] : '';
  const hm = v => { if (v instanceof Date) { const p = parts(v); return p ? ('0' + p.h).slice(-2) + ':' + ('0' + p.min).slice(-2) : ''; } return zopHm_(String(v == null ? '' : v).trim().replace('.', ':')); };
  return rows.slice(1).map(r => ({name: String(get(r, 'name') || '').trim(), topic: String(get(r, 'topic') || '').trim(), time: hm(get(r, 'time')), a: day(get(r, 'a')), b: day(get(r, 'b')), f: fun(get(r, 'f')),
    seg: String(get(r, 'seg') || '').trim(), cost: num(get(r, 'cost')), sent: num(get(r, 'sent')), note: String(get(r, 'note') || '').trim(),
    text: String(get(r, 'text') || '').trim().slice(0, 2000), media: zopMediaList_(get(r, 'media')), buttons: zopBtnList_(get(r, 'btn')),
    product: String(get(r, 'product') || '').trim(), goal: String(get(r, 'goal') || '').trim(), offer: String(get(r, 'offer') || '').trim(), planned: num(get(r, 'planned')), tariff: String(get(r, 'tariff') || '').trim(), status: String(get(r, 'status') || '').trim(), limits: String(get(r, 'limits') || '').trim(),
    fmt: String(get(r, 'fmt') || '').trim().slice(0, 40), num: String(get(r, 'num') || '').trim().replace(/\d{5,}/g, '…').slice(0, 30), hyp: String(get(r, 'hyp') || '').trim().slice(0, 300), cls: String(get(r, 'cls') || '').trim().slice(0, 60), lang: String(get(r, 'lang') || '').trim().slice(0, 60)}))   // нөмірде 5+ цифр шықпайды
    .filter(m => m.a != null && (m.name || m.topic));
}
/* =================== Расходы v2 (модуль Cost Control) ===================
   Лист «Расходы»: обязательные колонки — Дата, Подписка, Детали, Сумма, Оплата; дополнительные (по желанию, в любом порядке) —
   Категория, Поставщик, Валюта, Курс, Кезең басы / Кезең соңы (период), Кампания, Хабарлама саны, Тариф, Құжат, Тексерілді, ID.
   Старые строки без новых колонок читаются как раньше (категория определяется по «Детали»).
   Сумма неизвестна (нет суммы и нет тарифа, или валюта без курса) → sum = null: расход НЕ становится нулём.
   Дубли (тот же ID, или та же дата+категория+поставщик+сумма+кампания+документ) помечаются dup и в итоги не входят. */
const ZOP_COST_CATS = ['Wazzup top-up', 'Wazzup subscription', 'WhatsApp template charges', 'Provider service fees', 'Sipuni subscription', 'Marketing campaign costs', 'Other sales expenses', 'Wallet opening balance', 'Adjustment'];
const ZOP_COST_KIND = {'Wazzup top-up': 'balance', 'Wazzup subscription': 'plan', 'Sipuni subscription': 'plan', 'WhatsApp template charges': 'usage', 'Provider service fees': 'usage',
  'Marketing campaign costs': 'direct', 'Other sales expenses': 'direct', 'Wallet opening balance': 'opening', 'Adjustment': 'adjust'};
const ZOP_COST_HEAD_V2 = ['Дата', 'Подписка', 'Детали', 'Сумма, ₸', 'Оплата', 'Рассылка (если знаете)', 'Категория', 'Поставщик', 'Валюта', 'Курс', 'Кезең басы', 'Кезең соңы', 'Кампания', 'Хабарлама саны', 'Тариф', 'Құжат', 'Тексерілді', 'ID'];
/** Pure: «A=30000; B=20000» / «A:60%; B:40%» / «A» → [{name, sum}] + ошибка, если части не сходятся с итогом. */
function zopCostAlloc_(text, total) {
  const parts = String(text || '').split(';').map(x => x.trim()).filter(Boolean);
  if (!parts.length) return {items: [], err: null};
  if (parts.length === 1 && !/[=:]/.test(parts[0])) return {items: [{name: parts[0], sum: total}], err: null};
  const items = parts.map(x => { const m = x.match(/^(.+?)\s*[=:]\s*([\d\s.,]+)\s*(%?)$/); if (!m) return {name: x, sum: null, bad: true};
    const v = +m[2].replace(/\s/g, '').replace(',', '.'); return {name: m[1].trim(), sum: m[3] ? (total == null ? null : total * v / 100) : v, pct: m[3] ? v : null}; });
  if (items.some(x => x.bad)) return {items, err: 'кампания жолы оқылмады'};
  if (total == null) return {items, err: null};
  const sum = items.reduce((a, x) => a + (x.sum || 0), 0);
  return {items, err: Math.abs(sum - total) > 0.5 ? `бөлінген ${Math.round(sum)} ≠ сома ${Math.round(total)}` : null};
}
/** Pure: лист «Тарифтер» → [{prov, type, from, to, price, cur, src}] + предупреждения о пересечении периодов. */
function zopTariffs_(rows, parts, dayIdx) {
  if (!rows || rows.length < 2) return {list: [], warn: []};
  const head = rows[0].map(x => String(x || '').trim().toLowerCase()), col = re => head.findIndex(h => re.test(h));
  const c = {prov: col(/поставщ|провайдер|provider/), type: col(/түр|тип|type|категор/), from: col(/басталу|начал|from|с$/), to: col(/аяқталу|оконч|to|по$/), price: col(/баға|цена|price|тариф/), cur: col(/валют|currency/), src: col(/дерек|источ|source|құжат/)};
  const day = v => { if (v instanceof Date) return dayIdx(parts(v)); const m = String(v || '').match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/); return m ? dayIdx({y: m[3].length === 2 ? 2000 + +m[3] : +m[3], m: +m[2], d: +m[1]}) : null; };
  const get = (r, k) => c[k] >= 0 ? r[c[k]] : '';
  const list = rows.slice(1).map(r => ({prov: String(get(r, 'prov') || '').trim(), type: String(get(r, 'type') || '').trim(), from: day(get(r, 'from')), to: day(get(r, 'to')),
    price: typeof get(r, 'price') === 'number' ? get(r, 'price') : +String(get(r, 'price') || '').replace(/[^\d.,]/g, '').replace(',', '.') || null, cur: String(get(r, 'cur') || 'KZT').trim().toUpperCase() || 'KZT', src: String(get(r, 'src') || '').trim()}))
    .filter(t => t.prov && t.from != null && t.price != null);
  const warn = [];
  list.forEach((a, i) => list.slice(i + 1).forEach(b => { if (a.prov === b.prov && a.type === b.type && a.from <= (b.to == null ? 1e9 : b.to) && b.from <= (a.to == null ? 1e9 : a.to)) warn.push(`${a.prov} · ${a.type || '—'}: тариф кезеңдері қиылысады`); }));
  return {list, warn};
}
const zopTariffAt_ = (tariffs, prov, type, d) => (tariffs || []).filter(t => t.prov.toLowerCase() === String(prov || '').toLowerCase() && (!type || !t.type || t.type.toLowerCase() === String(type).toLowerCase()) && t.from <= d && (t.to == null || d <= t.to)).sort((a, b) => b.from - a.from)[0] || null;
function zopCosts_(rows, parts, dayIdx, tariffs) {
  if (!rows.length) return [];
  const head = rows[0].map(x => String(x || '').trim().toLowerCase()), col = re => head.findIndex(h => re.test(h));
  const c = {d: col(/^дата|^date/), sub: col(/подпис|subscr|канал/), det: col(/детал|detail|описан/), sum: col(/^сумм|^sum|сома/), pay: col(/^оплат|payment method|^төлем/), mail: col(/^рассылк/),
    cat: col(/категор|category/), prov: col(/поставщ|провайдер|provider/), cur: col(/валют|currency/), rate: col(/^курс|rate/), from: col(/кезең басы|начало периода|period from/), to: col(/кезең соңы|конец периода|period to/),
    camp: col(/кампан|campaign/), msgs: col(/хабарлама саны|сообщени|messages/), tariff: col(/^тариф$|tariff/), doc: col(/құжат|документ|document/), ver: col(/тексерілді|провер|verified/), id: col(/^id$/)};
  const num = v => { if (typeof v === 'number') return v; const t = String(v == null ? '' : v).replace(/[^\d.,-]/g, '').replace(',', '.'); return t ? +t : null; };
  const day = v => { if (v instanceof Date) return dayIdx(parts(v)); const m = String(v || '').match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/); return m ? dayIdx({y: m[3].length === 2 ? 2000 + +m[3] : +m[3], m: +m[2], d: +m[1]}) : null; };
  const get = (r, k) => c[k] >= 0 ? r[c[k]] : '', str = (r, k) => String(get(r, k) == null ? '' : get(r, k)).trim();
  const seen = {};
  return rows.slice(1).map(r => {
    const det = str(r, 'det'), sub = str(r, 'sub'), plan = /plan|тариф/i.test(det), d = day(get(r, 'd'));
    let cat = str(r, 'cat'); if (ZOP_COST_CATS.indexOf(cat) < 0) cat = cat ? cat : /sipuni/i.test(sub + ' ' + det) ? 'Sipuni subscription' : plan ? 'Wazzup subscription' : /balance|баланс/i.test(det) ? 'Wazzup top-up' : 'Other sales expenses';
    const kind = ZOP_COST_KIND[cat] || 'direct', cur = (str(r, 'cur') || 'KZT').toUpperCase(), rate = num(get(r, 'rate')), raw = num(get(r, 'sum')), msgs = num(get(r, 'msgs'));
    const prov = str(r, 'prov') || (/sipuni/i.test(sub) ? 'Sipuni' : 'Wazzup'), warn = [];
    let sum = raw, src = raw != null ? 'row' : null, tariff = str(r, 'tariff');
    if (sum == null && msgs != null && d != null) { const t = zopTariffAt_(tariffs, prov, tariff, d); if (t) { sum = msgs * t.price; src = 'tariff'; if (t.cur !== 'KZT') { sum = rate ? sum * rate : null; if (!rate) warn.push('тариф валютасы ' + t.cur + ', курс жоқ'); } } else warn.push('сома да, тариф та жоқ'); }
    else if (sum != null && cur !== 'KZT') { if (rate) sum = sum * rate; else { sum = null; warn.push('валюта ' + cur + ', курс жоқ'); } }
    else if (sum == null) warn.push('сома жоқ');
    const months = kind === 'plan' ? (/year|год|жыл/i.test(det) ? 12 : 1) : 0, from = day(get(r, 'from')), to = day(get(r, 'to'));
    const id = str(r, 'id'), key = id || [d, cat, prov, sub, raw, cur, str(r, 'pay'), str(r, 'doc'), str(r, 'camp')].join('|'), dup = !!seen[key]; seen[key] = 1;
    const al = zopCostAlloc_(str(r, 'camp'), sum); if (al.err) warn.push(al.err);
    return {d, sub, det, sum, pay: str(r, 'pay'), mail: str(r, 'mail'), kind, months, cat, prov, cur, rate, from, to, camp: str(r, 'camp'), alloc: al.items, msgs, tariff, doc: str(r, 'doc'),
      ver: /^(иә|да|yes|true|1|✓)$/i.test(str(r, 'ver')), id, src, dup, warn};
  }).filter(x => x.d != null && (x.sum != null || x.warn.length || x.msgs != null));
}
/* =================== продажи: «Оқушы тізімі» + платформа + лист «Жоспар» =================== */
const ZOP_PLAN_HEAD = ['Топ', 'Өнім (есептегі атауы)', 'Бағасы, ₸', 'Күндік жоспар, дана', 'Дана коэффициенті', 'Кілт сөз (атауында)', '', 'Қайдан келді (кестедегі мән)', 'Есептегі жол'];
// стартовые значения — из образца отчёта за 07.10.2026; дальше их меняют в листе «Жоспар»
const ZOP_PLAN_SEED = [
  ['Негізгі өнімдер', 'Математика — Мега қайталау', 13500, 10, 1, 'мега'],
  ['Негізгі өнімдер', 'Математика — Тоқсандық', 8400, 4, 1, 'тоқсан+айлық;четверть+айлық;математика+тоқсан'],
  ['Негізгі өнімдер', 'Математика — Күндік', 1190, 30, 1, 'тоқсан+күндік;четверть+күндік;математика+күнд'],
  ['Негізгі өнімдер', 'Ағылшын — Тоқсандық', 8400, 4, 1, 'english+айлық;ағылшын+айлық;ағылшын+тоқсан'],
  ['Негізгі өнімдер', 'Ағылшын — Күндік', 1190, 30, 1, 'english+күндік;ағылшын+күндік;ағылшын+күнд'],
  ['Алтын сақа', 'Барлығы', '', 67, '', ''],
  ['Алтын сақа', 'Бір тренинг', 5600, '', 1, ''],
  ['Алтын сақа', 'КОМБО', 24900, '', 1, ''],
  ['Алтын сақа', 'Тренажер', 13500, '', 1, ''],
  ['Алтын сақа', 'Мини-курс', 16900, '', 1, ''],
  ['Алтын сақа', 'Толық пакет (акция)', 13440, '', 1, ''],
  ['Алтын сақа', '2 тренинг −20% (рассылка)', 7840, '', 2, '']];
const ZOP_OLD_KEYS = {'математика+тоқсан': ZOP_PLAN_SEED[1][5], 'математика+күнд': ZOP_PLAN_SEED[2][5], 'ағылшын+тоқсан': ZOP_PLAN_SEED[3][5], 'ағылшын+күнд': ZOP_PLAN_SEED[4][5]};
const ZOP_SRC_SEED = [
  ['Региональный директор', 'Мектепке бару'],
  ['Органика WhatsApp', 'Маркетинг (органика)'], ['Органика Instagram', 'Маркетинг (органика)'], ['Органика Телеграмм', 'Маркетинг (органика)'], ['Рекомендация (Сарафан)', 'Маркетинг (органика)'],
  ['Таргетированная реклама в Инстаграм или Фейсбук', 'Маркетинг (таргет)'],
  ['Рассылка', 'Рассылка'],
  ['(платформа)', 'Платформадан өзі сатып алды'],
  ['Білім іс шаралары', 'EduApp орталықтары'],
  ['Вебинар', 'Вебинар'], ['база Алтын сақа', 'База Алтын сақа'], ['Тікелей жіберді', 'Тікелей жіберді'],
  ['(бос)', 'Басқа (Не указан)']];


/* Таблица продаж только читается. Если там остался лист с прошлой версии — берём его строки (чтение), чтобы ваши правки не пропали. */
function zopCopyFromSales_(name, sh) {
  let src = null; try { src = SpreadsheetApp.openById(ZOP_SALES.ID).getSheetByName(name); } catch (err) { return false; }
  if (!src || src.getLastRow() < 2) return false;
  const v = src.getDataRange().getValues();
  sh.getRange(1, 1, v.length, v[0].length).setValues(v);
  sh.getRange(1, 1, 1, v[0].length).setFontWeight('bold').setBackground('#eef1f7');
  return true;
}
/** Создаёт лист «Жоспар» в главной таблице (AMO), если его нет. Таблицу продаж не меняет. */
function zopCreatePlanSheet_() {
  const ss = zopHome_();
  if (ss.getSheetByName(ZOP_SALES.PLAN)) return 0;
  const sh = ss.insertSheet(ZOP_SALES.PLAN), n = Math.max(ZOP_PLAN_SEED.length, ZOP_SRC_SEED.length);
  if (zopCopyFromSales_(ZOP_SALES.PLAN, sh)) { sh.setFrozenRows(1); return 1; }
  const rows = Array.from({length: n}, (_, i) => (ZOP_PLAN_SEED[i] || ['', '', '', '', '', '']).concat([''], ZOP_SRC_SEED[i] || ['', '']));
  sh.getRange(1, 1, 1, ZOP_PLAN_HEAD.length).setValues([ZOP_PLAN_HEAD]).setFontWeight('bold').setBackground('#eef1f7');
  sh.getRange(2, 1, n, ZOP_PLAN_HEAD.length).setValues(rows);
  sh.setFrozenRows(1); sh.getRange('C:C').setNumberFormat('#,##0');
  [140, 230, 90, 130, 130, 170, 24, 300, 220].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('D2:D').setBackground('#fff8e1');
  sh.getRange(n + 3, 1, 6, 1).setValues([
    ['Как это работает:'],
    ['• «Күндік жоспар» — план на один день (дана). Для Алтын сақа план задаётся строкой «Барлығы».'],
    ['• Продажа попадает в строку, если её название содержит «Кілт сөз» (несколько слов через «+» — все должны быть; варианты через «;»), а цена совпадает с «Бағасы» (пустая цена — любая).'],
    ['• Продажи из «Оқушы тізімі» без совпадения по слову относятся к Алтын сақа по цене; другие цены — «Тренинг, басқа баға». Покупки платформы без совпадения — «Басқа өнімдер».'],
    ['• «Дана коэффициенті» — сколько штук считать за одну оплату (2 тренинга −20% = 2).'],
    ['• Справа: значение «Қайдан келді» → строка отчёта. (платформа) — покупки на платформе, (бос) — пустое поле.']]).setFontColor('#64748b');
  return 1;
}

/* =================== Все источники в одной таблице (AMO) ===================
   Каждому внешнему источнику — свой лист-зеркало в таблице AMO: формула IMPORTRANGE сама подтягивает изменения оригинала.
   Решение владельца 12.10: зеркало — ПОЛНАЯ копия листа (все столбцы). Таблица AMO закрытая; из скрипта наружу по-прежнему выходят только агрегаты.
   cols — только если понадобится выборка столбцов по заголовку (сейчас не используется).
   Скрипт читает зеркало; если оно ещё не готово (доступ не разрешён, «Loading…», нет столбцов) — читает оригинал напрямую.
   «Сделки» (amoCRM) уже живёт в таблице AMO — синхронизация пишет туда сама. */
const ZOP_MIRROR = {
  sales: {sheet: 'ОП миникурс', id: ZOP_SALES.ID, src: ZOP_SALES.LIST, title: 'ОП Миникурс',
    cols: null, need: ['Дата']},
  plat: {sheet: 'All Trainings', id: ZOP_PLATFORM.ID, src: ZOP_PLATFORM.LIST, alt: ZOP_PLATFORM.LIST_RAW, title: 'All Trainings and Challanges',
    cols: null, need: ['invoice_id']}
};
const ZOP_SRC_INDEX = 'Дереккөздер';
const ZOP_CHUNK = 2000;   // строк в одной IMPORTRANGE
/** Pure: номер столбца → буквы (1 → A, 27 → AA). */
function zopColLetter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
/** Pure: формула зеркала — весь лист источника, из него только нужные столбцы по названию заголовка (порядок в источнике может меняться).
 *  XMATCH по списку названий — только внутри ARRAYFORMULA, иначе находит лишь первое (вышел бы один столбец). */
function zopMirrorFormula_(id, sheet, lastCol, cols, lastRow) {
  const sh = "'" + String(sheet).replace(/'/g, "''") + "'!", L = zopColLetter_(Math.max(1, lastCol || 1)), rng = sh + 'A1:' + L;
  if (!cols || !cols.length) {   // полная копия: сама находит последнюю строку источника (по столбцу B) и берёт её частями по ZOP_CHUNK строк
    // (одна IMPORTRANGE — не больше ~10 МБ). Без жёсткого предела строк и без IFERROR — ошибка части видна, а не теряется молча.
    const q = x => '"' + String(x).replace(/"/g, '""') + '"';
    return '=LET(u,' + q(id) + ',s,' + q(sh) + ',k,' + ZOP_CHUNK + ',b,IMPORTRANGE(u,s&"B1:B"),n,MAX(FILTER(SEQUENCE(ROWS(b)),b<>"")),'
      + 'f,IMPORTRANGE(u,s&"A1:' + L + '"&MIN(n,k)),IF(n<=k,f,REDUCE(f,SEQUENCE(ROUNDUP(n/k)-1),LAMBDA(acc,i,VSTACK(acc,IMPORTRANGE(u,s&"A"&(i*k+1)&":' + L + '"&MIN(n,(i+1)*k)))))))';
  }

  return '=LET(s,IMPORTRANGE("' + id + '","' + rng.replace(/"/g, '""') + '"),h,INDEX(s,1),i,ARRAYFORMULA(XMATCH(SPLIT("' + cols.join('|').replace(/"/g, '""') + '","|"),h)),CHOOSECOLS(s,FILTER(i,ISNUMBER(i))))';
}
/** Pure: зеркало готово к чтению — есть строки, нужные заголовки, нет ошибки формулы. */
function zopMirrorOk_(rows, need) {
  if (!rows || rows.length < 2) return false;
  const h = rows[0].map(x => String(x == null ? '' : x).trim());
  if (!h.length || /^(#|Loading)/i.test(h[0])) return false;
  return (need || []).every(n => h.indexOf(n) >= 0);
}
/** Строки источника: зеркало в таблице AMO, иначе оригинал. via — откуда прочитано (для «Деректер сапасы»). */
function zopSourceRows_(kind, errors) {
  const M = ZOP_MIRROR[kind], info = zopInfoRows_(M.sheet, M.need);
  if (info) return {rows: info, via: ZOP_INFO.title + ' · «' + M.sheet + '»'};
  errors.push('«' + ZOP_INFO.title + '» · «' + M.sheet + '» дайын емес — AMO / түпнұсқадан оқылды');
  try { const sh = zopHome_().getSheetByName(M.sheet); if (sh && sh.getLastRow() > 1) { const rows = sh.getDataRange().getValues(); if (zopMirrorOk_(rows, M.need)) return {rows, via: 'AMO · «' + M.sheet + '»'}; errors.push('«' + M.sheet + '» парағы әлі дайын емес — түпнұсқадан оқылды'); } }
  catch (err) { errors.push('«' + M.sheet + '»: ' + err.message); }
  for (const name of [M.src, M.alt].filter(Boolean)) {
    try { const sh = SpreadsheetApp.openById(M.id).getSheetByName(name); if (!sh) { errors.push('нет листа «' + name + '»'); continue; }
      const rows = sh.getLastRow() > 0 ? sh.getDataRange().getValues() : []; if (zopMirrorOk_(rows, M.need)) return {rows, via: M.title + ' · «' + name + '» (тікелей)'}; errors.push('«' + name + '»: нужных столбцов нет'); }
    catch (err) { errors.push('«' + name + '»: ' + err.message); }
  }
  return {rows: [], via: null};
}
/** Запустить один раз (или из меню EduApp → «Дереккөздерді қосу»): создаёт листы-зеркала и лист «Дереккөздер». Другие листы не трогает.
 *  После запуска откройте каждый лист-зеркало и нажмите «Разрешить доступ» у ячейки A1 (Google спрашивает один раз на источник). */
function zopSetupSources() {
  try { SpreadsheetApp.getUi(); } catch (_) { zopRequireOwner_(); }
  const ss = zopHome_(), made = [];
  Object.keys(ZOP_MIRROR).forEach(k => {
    const M = ZOP_MIRROR[k]; let lastCol = 60, lastRow = 20000;
    try { const src = SpreadsheetApp.openById(M.id).getSheetByName(M.src); if (src) { lastCol = src.getLastColumn(); lastRow = src.getLastRow(); } } catch (err) { Logger.log('«' + M.title + '»: ' + err.message); }
    const ex = ss.getSheetByName(M.sheet); if (ex && ex.getLastRow() > 1) { made.push(M.sheet + ' (бұрыннан бар)'); return; }   // қолмен қойылған формуланы өшірмейміз
    const sh = ex || ss.insertSheet(M.sheet);
    sh.clear(); sh.getRange('A1').setFormula(zopMirrorFormula_(M.id, M.src, lastCol, M.cols, lastRow));
    sh.getRange('A1').setNote('Айна: «' + M.title + '» → «' + M.src + '». Өздігінен жаңарады. Қолмен өзгертпеңіз — түпнұсқада өзгертіңіз.');
    sh.setTabColor('#2EE6A6'); sh.setFrozenRows(1); made.push(M.sheet);
  });
  const ix = ss.getSheetByName(ZOP_SRC_INDEX) || ss.insertSheet(ZOP_SRC_INDEX);
  ix.clear();
  const rows = [['Дереккөз', 'Түпнұсқа кесте', 'Түпнұсқа парақ', 'Осы кестедегі парақ', 'Жол саны', 'Қалай жаңарады']];
  rows.push(['АМО (amoCRM)', 'https://docs.google.com/spreadsheets/d/' + ZOP_CRM.ID, ZOP_CRM.LEADS[0], ZOP_CRM.LEADS[0], "=COUNTA('" + ZOP_CRM.LEADS[0] + "'!A:A)-1", 'amoCRM синхрондауы (сағат сайын)']);
  Object.keys(ZOP_MIRROR).forEach(k => { const M = ZOP_MIRROR[k]; rows.push([M.title, 'https://docs.google.com/spreadsheets/d/' + M.id, M.src, M.sheet, "=COUNTA('" + M.sheet + "'!A:A)-1", 'IMPORTRANGE — түпнұсқа өзгерсе, өздігінен']); });
  ix.getRange(1, 1, rows.length, rows[0].length).setValues(rows); ix.getRange(1, 1, 1, rows[0].length).setFontWeight('bold'); ix.setFrozenRows(1); ix.autoResizeColumns(1, rows[0].length);
  Logger.log('Дайын: ' + made.join(', ') + ' + «' + ZOP_SRC_INDEX + '». Әр айна парағын ашып, A1 ұяшығында «Разрешить доступ» басыңыз.');
  return made;
}
function zopSetupSourcesMenu() { zopRequireUi_(); const m = zopSetupSources(); SpreadsheetApp.getUi().alert('Парақтар дайын: ' + m.join(', ') + '.\nӘр парақты ашып, A1 ұяшығында «Разрешить доступ» басыңыз (бір рет).'); }

/** Один раз (из редактора): переносит служебные листы из AMO в «All Info» — копии с данными; в AMO ничего не удаляет.
 *  «Расходы» копируется последним: после него главной таблицей становится «All Info». Ставит в «All Info» меню EduApp и выпадающие списки. */
function zopMoveServiceSheets() {
  zopRequireOwner_();
  const amo = SpreadsheetApp.openById(ZOP_HOME_OLD), info = SpreadsheetApp.openById(ZOP_HOME_ID), done = [], skip = [];
  const names = [ZOP_COST_LOG, ZOP.MAIL, ZOP.KPI, ZOP_SALES.PLAN, ZOP_PLIST, ZOP_SALES.PROD, ZOP_PSEC, ZOP_TARIFF, ZOP_HYP, ZOP.COST];   // «Расходы» — последним
  names.forEach(n => {
    const src = amo.getSheetByName(n); if (!src) return;
    if (info.getSheetByName(n)) { skip.push(n); return; }
    const c = src.copyTo(info); c.setName(n); if (src.isSheetHidden()) c.hideSheet(); done.push(n);
  });
  ZOP_HOME_CACHE = null;
  try { zopApplyProductSections_(); } catch (err) { Logger.log('Өнімдер тізімдері: ' + err.message); }
  ScriptApp.getProjectTriggers().filter(t => ['zopInfoOpen_', 'zopInfoEdit_'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('zopInfoOpen_').forSpreadsheet(ZOP_HOME_ID).onOpen().create();
  ScriptApp.newTrigger('zopInfoEdit_').forSpreadsheet(ZOP_HOME_ID).onEdit().create();
  try { CacheService.getScriptCache().removeAll(['zop', 'zops']); } catch (_) {}
  Logger.log('Көшірілді: ' + (done.join(', ') || '—') + (skip.length ? ' · бұрыннан бар: ' + skip.join(', ') : '') + ' · басты кесте: ' + zopHome_().getName());
  return {done, skip};
}
function zopInfoOpen_() { zopMenu_(); }
function zopInfoEdit_(e) { onEdit(e); }

function zopSalesBuild_() {
  const tz = zopHome_().getSpreadsheetTimeZone(), errors = [];
  const z = Utilities.formatDate(new Date(), tz, 'Z'), off = (z[0] === '-' ? -1 : 1) * (+z.slice(1, 3) * 60 + +z.slice(3, 5)) * 60000;
  const parts = d => { if (!(d instanceof Date) || isNaN(d.getTime())) return null; const x = new Date(d.getTime() + off);
    return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours(), min: x.getUTCMinutes()}; };
  const own = zopHome_(), prodSheet = own.getSheetByName(ZOP_SALES.PROD), planSheet = own.getSheetByName(ZOP_SALES.PLAN);   // служебные листы — в этой таблице
  const prodRows = prodSheet && prodSheet.getLastRow() > 1 ? prodSheet.getRange(1, 1, prodSheet.getLastRow(), ZOP_PROD_HEAD.length).getValues() : [];
  const plist = own.getSheetByName(ZOP_PLIST), mapRows = plist && plist.getLastRow() > 1 ? plist.getRange(2, 1, plist.getLastRow() - 1, 7).getValues().map(r => [r[0], r[5], r[6]]) : [];   // A нақты атауы, F есептегі жол, G бағыт
  const S1 = ZOP_USE_OP_SHEET ? zopSourceRows_('sales', errors) : {rows: [], via: 'қолданылмайды (CRM + платформа)'}, P1 = zopSourceRows_('plat', errors), sheetRows = S1.rows, platRows = P1.rows;   // зеркала в таблице AMO, иначе оригиналы
  const out = zopSalesAggregate_(sheetRows, planSheet && planSheet.getLastRow() > 0 ? planSheet.getDataRange().getValues() : [], platRows, parts, off, prodRows, mapRows);
  // построчная сверка нужна, только если в платформе есть копия ОП (строки без invoice); исходный лист платформы её не содержит
  try { if (sheetRows.length && platRows.length) { const rc = zopReconcileRows_(sheetRows, platRows); if (rc && rc.copy && rc.copy.rows) out.recon = rc; } } catch (err) { errors.push('сверка: ' + err.message); }
  const now = parts(new Date());
  out.snapshot = `${now.y}-${String(now.m).padStart(2, '0')}-${String(now.d).padStart(2, '0')}T${String(now.h).padStart(2, '0')}:${String(now.min).padStart(2, '0')}`;
  out.errors = errors; out.via = {sheet: S1.via, plat: P1.via};
  return out;
}


/* =================== Построчная сверка продаж (P0.1) ===================
   Внутри Apps Script: каждая строка «Оқушы тізімі» и каждая строка платформы без invoice превращается во внутренний ключ
   из ВСЕХ колонок, которые есть в обеих таблицах (дата, курс, суммы, источник и т. д.). Ключи сравниваются как мультимножества:
   совпало / только в листе / только в платформе. Наружу уходят только счётчики, суммы и названия колонок — не значения.
   Плюс проверка счетов платформы: уникальные invoice_id, счета из нескольких позиций, смешанные статусы, возвраты. */
function zopReconcileRows_(sheetRows, platRows) {
  const H = rows => (rows[0] || []).map(h => String(h == null ? '' : h).trim());
  const hs = H(sheetRows), hp = H(platRows), ip = {}; hp.forEach((h, i) => { if (h && ip[h] === undefined) ip[h] = i; });
  const is = {}; hs.forEach((h, i) => { if (h && is[h] === undefined) is[h] = i; });
  const common = hs.filter((h, i) => h && hs.indexOf(h) === i && ip[h] !== undefined);
  const norm = v => v instanceof Date ? (isNaN(v.getTime()) ? '' : v.toISOString()) : typeof v === 'number' ? String(Math.round(v * 100)) : String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' ');
  const filled = r => r && r.some(v => v !== '' && v !== null && v !== undefined);
  const sumCol = ['Төлем сомасы', 'Бағасы'].find(c => common.indexOf(c) >= 0) || null;
  const amt = (r, idx) => { if (!sumCol) return 0; const v = r[idx[sumCol]]; return typeof v === 'number' ? v : +String(v || '').replace(/[^\d.,-]/g, '').replace(',', '.') || 0; };
  const invI = ip['invoice_id'], stI = ip['invoice_status'];
  const bag = (rows, idx, pick) => { const m = new Map(); let n = 0; rows.slice(1).forEach(r => { if (!filled(r) || !pick(r)) return; n++; const k = common.map(c => norm(r[idx[c]])).join('\u0001'); const e = m.get(k) || {n: 0, s: 0}; e.n++; e.s += amt(r, idx); m.set(k, e); }); return {m, n}; };
  const S = bag(sheetRows, is, () => true), C = bag(platRows, ip, r => invI === undefined || String(r[invI] == null ? '' : r[invI]).trim() === '');
  let matched = 0, sheetOnly = 0, copyOnly = 0, sheetOnlySum = 0, copyOnlySum = 0, dupKeys = 0, dupRows = 0;
  new Set([...S.m.keys(), ...C.m.keys()]).forEach(k => { const a = S.m.get(k) || {n: 0, s: 0}, b = C.m.get(k) || {n: 0, s: 0}, mm = Math.min(a.n, b.n);
    matched += mm; if (a.n > mm) { sheetOnly += a.n - mm; sheetOnlySum += a.s / a.n * (a.n - mm); } if (b.n > mm) { copyOnly += b.n - mm; copyOnlySum += b.s / b.n * (b.n - mm); }
    if (a.n > 1) { dupKeys++; dupRows += a.n; } });
  // invoices
  const inv = {rows: 0, unique: 0, multiItem: 0, mixedStatus: 0, statuses: {}, refundLike: 0}, byInv = {};
  if (invI !== undefined) platRows.slice(1).forEach(r => { if (!filled(r)) return; const id = String(r[invI] == null ? '' : r[invI]).trim(); if (!id) return; inv.rows++;
    const st = stI === undefined ? '' : (r[stI] instanceof Date ? 'date-coded' : String(r[stI] == null ? '' : r[stI]).trim()); inv.statuses[st || '(бос)'] = (inv.statuses[st || '(бос)'] || 0) + 1;
    if (/refund|возврат|қайтар/i.test(st)) inv.refundLike++;
    const e = byInv[id] || (byInv[id] = {n: 0, st: {}}); e.n++; e.st[st] = 1; });
  Object.values(byInv).forEach(e => { inv.unique++; if (e.n > 1) inv.multiItem++; if (Object.keys(e.st).length > 1) inv.mixedStatus++; });
  return {method: 'rows-multiset-v1', commonCols: common, keyCols: common.length, sheet: {rows: S.n, duplicateKeys: dupKeys, rowsInDuplicateKeys: dupRows}, copy: {rows: C.n},
    matched, sheetOnly, copyOnly, sheetOnlySum: Math.round(sheetOnlySum), copyOnlySum: Math.round(copyOnlySum), invoices: inv};
}

/* invoice_status — число (2 = оплачен); в части строк ячейка в формате даты, и 2 приходит как 01.01.1900 → возвращаем число */
function zopStatus_(v, parts) {
  if (v instanceof Date) { const p = parts(v); return p ? String(Math.round((Date.UTC(p.y, p.m - 1, p.d) - Date.UTC(1899, 11, 30)) / 864e5)) : ''; }
  return String(v == null ? '' : v).trim();
}
/* Pure: rows → дневные агрегаты без персональных данных. day = 'YYYY-MM-DD'. */
function zopSalesAggregate_(sheetRows, planRows, platRows, parts, off, prodRows, mapRows) {
  const nonEmpty = v => v !== null && v !== undefined && String(v).trim() !== '';
  const pad = n => String(n).padStart(2, '0'), key = p => `${p.y}-${pad(p.m)}-${pad(p.d)}`;
  const num = v => { if (typeof v === 'number') return v; const t = String(v == null ? '' : v).replace(/[^\d.,-]/g, '').replace(/\s/g, ''); if (!t) return null;
    const x = +(t.indexOf(',') >= 0 && t.indexOf('.') < 0 ? t.replace(',', '.') : t.replace(/,/g, '')); return isFinite(x) ? x : null; };
  const day = v => {
    if (v instanceof Date) { const p = parts(v); return p ? key(p) : null; }
    const s = String(v == null ? '' : v).trim(); if (!s) return null;
    let m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})/); if (m) return `${m[3]}-${pad(+m[2])}-${pad(+m[1])}`;
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::\d{2}(?:\.\d+)?)?)?\s*(Z|[+-]\d{2}:?\d{2})?$/);
    if (m) { if (!m[6] || !m[4]) return `${m[1]}-${m[2]}-${m[3]}`; const p = parts(new Date(Date.parse(s.replace(' ', 'T')))); return p ? key(p) : null; }
    if (typeof v === 'number' && v > 30000 && v < 80000) { const p = parts(new Date(Math.round((v - 25569) * 864e5) - off)); return p ? key(p) : null; }   // серийная дата Sheets
    return null;
  };
  const dict = {}, code = (dim, v) => { const d = dict[dim] || (dict[dim] = []); let i = d.indexOf(v); if (i < 0) { d.push(v); i = d.length - 1; } return i; };
  const cols = rows => { const h = (rows[0] || []).map(x => String(x || '').trim()), ix = {}; h.forEach((x, i) => { if (x && ix[x] === undefined) ix[x] = i; });
    return (r, ...names) => { for (const n of names) if (ix[n] !== undefined && nonEmpty(r[ix[n]])) return r[ix[n]]; return ''; }; };
  const cube = new Map(), add = (k, n, sum) => { const id = k.join('|'); const e = cube.get(id) || {k, n: 0, s: 0}; e.n += n; e.s += sum; cube.set(id, e); };
  // иесі 15.10: «покупки по регионам и по классам» — [көзі, күн, сынып | өңір коды, саны, сомасы]; бағаны атауымен табамыз (ОП кестесінде «Сыныбы»)
  const pcls = new Map(), preg = new Map(), pAdd = (m, k, sum) => { const id = k.join('|'), e = m.get(id) || {k, n: 0, s: 0}; e.n++; e.s += sum; m.set(id, e); };
  const RE_CLS = /^(сынып|сыныбы|класс|class|grade)/i, RE_REG = /^(облыс|өңір|регион|область|region)/i, hcol = (rows, re) => (rows[0] || []).map(x => String(x || '').trim()).findIndex(x => re.test(x));
  const HS = (sheetRows[0] || []).map(x => String(x || '').trim()), HP = (platRows[0] || []).map(x => String(x || '').trim());
  const sC = hcol(sheetRows, RE_CLS), sR = hcol(sheetRows, RE_REG), pC = hcol(platRows, RE_CLS), pR = hcol(platRows, RE_REG);
  const clsOf = v => { const m = String(v == null ? '' : v).match(/\d{1,2}/), c = m ? +m[0] : 0; return c >= 1 && c <= 12 ? c : 0; };
  const sheetMan = new Map();   // «Оқушы тізімі» × «Енгізген қызметкер» (кто внёс оплату): [день, код сотрудника, количество, сумма] — для раздела ОП
  let nSheet = 0, nPlat = 0;
  // provenance codes kept in the cube (last two key columns): date source, amount source
  //   dsrc: 0 = payment date (paid_at / sheet «Дата»), 1 = only an invoice date (NOT a payment day), 2 = no date
  //   asrc: 0 = paid amount (item_total / «Төлем сомасы»), 1 = invoice_total (whole invoice), 2 = reference price (tariff / «Бағасы»), 3 = no amount
  const Q = {sheetNoAmount: 0, sheetListPrice: 0, platInvoiceDateOnly: 0, platNoDate: 0, platInvoiceTotal: 0, platInvoiceTotalRepeat: 0, platReferencePrice: 0, platNoAmount: 0, duplicateRows: 0};
  const seenRow = {}, seenInvTotal = {};

  // «Оқушы тізімі»: дата, источник, название курса, сумма оплаты (имя, телефон, чек не читаются)
  const gs = cols(sheetRows);
  sheetRows.slice(1).forEach(r => {
    const d = day(gs(r, 'Дата')); if (!d) return;
    const paid = num(gs(r, 'Төлем сомасы')), list = paid == null ? num(gs(r, 'Бағасы')) : null, asrc = paid != null ? 0 : list != null ? 2 : 3;
    if (asrc === 2) Q.sheetListPrice++; if (asrc === 3) Q.sheetNoAmount++;
    const price = paid ?? list ?? 0;   // the provenance code says whether this is a paid amount or only a list price
    nSheet++;
    add(['s', d, code('src', String(gs(r, 'Қайдан келді')).trim()), code('name', String(gs(r, 'Миникурс Атауы')).trim()), price, 0, 0, 0, 0, 0, asrc], 1, price);
    if (sC >= 0) pAdd(pcls, ['s', d, clsOf(r[sC])], price); if (sR >= 0) pAdd(preg, ['s', d, code('region', zopRegion_(r[sR]))], price);
    const ek = code('emp', String(gs(r, 'Енгізген қызметкер', 'Менеджер')).trim()), mk = d + '|' + ek, me = sheetMan.get(mk) || [d, ek, 0, 0]; me[2]++; me[3] += price; sheetMan.set(mk, me);
  });
  // платформа: дата оплаты, название продукта/тренинга/тарифа, сумма, тип, статус счёта
  const gp = cols(platRows);
  platRows.slice(1).forEach(r => {
    // date: the payment day only from paid_at (or the sheet «Дата»); an invoice date is kept but marked — it is not a payment day
    const dPaid = day(gp(r, 'paid_at', 'Дата')), dInv = dPaid ? null : day(gp(r, 'date_issued', 'invoice_created_at'));
    const d = dPaid || dInv, dsrc = dPaid ? 0 : dInv ? 1 : 2; if (!d) { Q.platNoDate++; return; }
    if (dsrc === 1) Q.platInvoiceDateOnly++;
    const inv = String(gp(r, 'invoice_id')).trim();
    const label = [gp(r, 'training_name', 'challenge_name', 'product_name', 'Миникурс Атауы'), gp(r, 'tariff_name')].map(x => String(x).trim()).filter((x, i, a) => x && a.indexOf(x) === i).join(' · ');
    // amount: item_total (or «Төлем сомасы») is the paid amount; invoice_total covers the whole invoice — counted once per invoice; tariff price is only a reference
    let amt = num(gp(r, 'item_total', 'Төлем сомасы')), asrc = 0;
    if (amt == null) { const it = num(gp(r, 'invoice_total')); if (it != null) { asrc = 1; Q.platInvoiceTotal++; if (inv && seenInvTotal[inv]) { amt = 0; Q.platInvoiceTotalRepeat++; } else { amt = it; if (inv) seenInvTotal[inv] = 1; } } }
    if (amt == null) { const rp = num(gp(r, 'tariff_price', 'Бағасы')); if (rp != null) { asrc = 2; amt = rp; Q.platReferencePrice++; } else { asrc = 3; amt = 0; Q.platNoAmount++; } }
    // the same invoice item twice (same invoice, product, amount, date) is one payment, not two
    if (inv) { const rk = [inv, label, amt, d].join('|'); if (seenRow[rk]) { Q.duplicateRows++; return; } seenRow[rk] = 1; }
    nPlat++;
    add(['p', d, code('src', String(gp(r, 'Қайдан келді')).trim()), code('name', label), amt, code('ptype', String(gp(r, 'Тип продукта')).trim()),
      code('pstatus', zopStatus_(gp(r, 'invoice_status'), parts)), inv ? 1 : 0, code('pub', String(gp(r, 'Жариялану статусы')).trim()), dsrc, asrc], 1, amt);
    if (inv && dsrc === 0 && zopPayStatus_(zopStatus_(gp(r, 'invoice_status'), parts)) === 'paid') { if (pC >= 0) pAdd(pcls, ['p', d, clsOf(r[pC])], amt); if (pR >= 0) pAdd(preg, ['p', d, code('region', zopRegion_(r[pR]))], amt); }   // тек есепке кіретін төлем (payWhy-мен бірдей)
  });
  // лист «Жоспар»: продукты и планы (слева), источники → строки отчёта (справа)
  const products = [], sources = [];
  if (planRows.length) {
    const h = planRows[0].map(x => String(x || '').trim().toLowerCase()), c = re => h.findIndex(x => re.test(x));
    const C = {g: c(/^топ/), n: c(/^өнім/), p: c(/^баға/), pl: c(/жоспар/), k: c(/коэфф/), w: c(/кілт|ключ/), s: c(/қайдан/), l: c(/есептегі жол/)};
    planRows.slice(1).forEach(r => {
      const g = C.g >= 0 ? String(r[C.g] || '').trim() : '', n = C.n >= 0 ? String(r[C.n] || '').trim() : '';
      let keys = C.w >= 0 ? String(r[C.w] || '').trim() : '';
      keys = ZOP_OLD_KEYS[keys] || keys;   // first version of the sheet had keywords that miss the platform names
      if (g && n) products.push({group: g, name: n, price: C.p >= 0 ? num(r[C.p]) : null, plan: C.pl >= 0 ? num(r[C.pl]) : null, coef: (C.k >= 0 ? num(r[C.k]) : null) || 1, keys});
      const sv = C.s >= 0 ? String(r[C.s] || '').trim() : '', sl = C.l >= 0 ? String(r[C.l] || '').trim() : '';
      if (sv && sl) sources.push([sv, sl]);
    });
  }
  // лист «Өнімдер»: каталог — раздел, название в отчёте, слово в названии, цена, «всегда показывать», план, коэффициент
  const catalog = [];
  if ((prodRows || []).length) {
    const h = prodRows[0].map(x => String(x || '').trim().toLowerCase()), c = re => h.findIndex(x => re.test(x));
    const C = {pp: c(/жоспар бағасы/), s: c(/^бөлім|^топ|^раздел/), r: c(/нақты|платформа/), n: c(/есептегі|^название/), k: c(/кілт|слово/), p: c(/^баға|^цена/), a: c(/әрқашан|всегда/), pl: c(/жоспар|план/), q: c(/коэфф/)};
    const val = (r, k) => C[k] >= 0 ? r[C[k]] : '';
    prodRows.slice(1).forEach(r => {
      const sec = String(val(r, 's') || '').trim(), real = String(val(r, 'r') || '').trim(), keys = String(val(r, 'k') || '').trim(), price = num(val(r, 'p'));
      // в отчёте: «Есептегі атауы», иначе короткое настоящее название, иначе слово + цена
      const name = String(val(r, 'n') || '').trim() || (real ? zopShortName_(real.split(/,\s+/)[0]) : keys ? keys.split(';')[0].split('+').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ') + (price != null ? ' · ' + String(Math.round(price)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' ₸' : '') : '');
      const plan = num(val(r, 'pl')), total = /^(барлығы|итого|всего)$/i.test(name);   // «Барлығы» — план раздела
      if (!name || !real && !keys && price == null && !(total && plan != null) && !String(val(r, 'n') || '').trim()) return;   // a row with only C can get products from «Өнім тізімі» F
      catalog.push({sec: sec || 'Басқа өнімдер', name, real, keys, price, always: /^(иә|ия|да|yes|y|1|true|✓|✔)$/i.test(String(val(r, 'a')).trim()),
        plan, coef: num(val(r, 'q')) || 1, pprice: num(val(r, 'pp'))});
    });
  }
  // «Өнім тізімі» F: продукт → строка отчёта (по «Есептегі атауы»); добавляется к «Нақты атауы» этой строки — как выбор нескольких
  (mapRows || []).forEach(([real, line]) => { real = String(real || '').trim(); line = String(line || '').trim(); if (!real || !line) return;
    const c = catalog.find(x => x.name === line); if (!c) return;
    const has = String(c.real || '').split(/,\s+/).some(x => x.trim().toLowerCase() === real.toLowerCase()); if (!has) c.real = c.real ? c.real + ', ' + real : real; });
  const rows = [...cube.values()].sort((a, b) => a.k.join('|') < b.k.join('|') ? -1 : 1).map(e => e.k.concat([e.n, Math.round(e.s)]));
  ['src', 'name', 'ptype', 'pstatus', 'pub', 'emp'].forEach(k => { dict[k] = dict[k] || []; });
  // бағыт — «Өнім тізімі» G (иесі түзетеді); dict.name-мен қатар, тізімде жоқ атау — '' (Атлас өзі болжайды және соны көрсетеді)
  const famOf = {}; (mapRows || []).forEach(r => { const real = String(r[0] || '').trim().toLowerCase(), f = String(r[2] || '').trim(); if (real && f) famOf[real] = f; });
  dict.fam = dict.name.map(n => famOf[String(n || '').trim().toLowerCase()] || '');
  const safe = (m, hid) => { const o = new Map(); m.forEach(e => { const k = e.n >= 3 ? e.k : [e.k[0], e.k[1], hid()], id = k.join('|'), x = o.get(id) || {k, n: 0, s: 0}; x.n += e.n; x.s += e.s; o.set(id, x); });   // 1–2 төлем — сынып/өңір жасырылады
    return [...o.values()].sort((a, b) => a.k.join('|') < b.k.join('|') ? -1 : 1).map(e => e.k.concat([e.n, Math.round(e.s)])); };
  const payProf = {cls: safe(pcls, () => -1), reg: safe(preg, () => code('region', 'Жасырылған (< 3)')), cols: {sheet: {cls: sC >= 0 ? HS[sC] : '', reg: sR >= 0 ? HS[sR] : ''}, plat: {cls: pC >= 0 ? HP[pC] : '', reg: pR >= 0 ? HP[pR] : ''}}};
  dict.region = dict.region || [];
  return {kind: 'sales', schemaVersion: 2, methodVersion: 'sales-2026-10-10', rows: {sheet: nSheet, platform: nPlat}, quality: Q, dict, cube: rows, payProf,
    sheetMan: [...sheetMan.values()].sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]).map(e => [e[0], e[1], e[2], Math.round(e[3])]), plan: {products: products.length ? products : ZOP_PLAN_SEED.map(r => ({group: r[0], name: r[1], price: r[2] === '' ? null : r[2], plan: r[3] === '' ? null : r[3], coef: r[4] || 1, keys: r[5]})),
    sources: sources.length ? sources : ZOP_SRC_SEED.slice(), fromSheet: products.length > 0, catalog}};
}

/* =================== лист «Өнімдер»: каталог продуктов для отчёта =================== */
const ZOP_PROD_HEAD = ['Бөлім', 'Нақты атауы (платформадан — тізімнен таңдаңыз)', 'Есептегі атауы (бос болса — нақты атауы)', 'Бағасы, ₸', 'Әрқашан көрсету', 'Күндік жоспар, дана', 'Дана коэффициенті', 'Кілт сөз («Оқушы тізімі» үшін)', 'Ескертпе · 90 күнде сатылды', 'Жоспар бағасы, ₸ (D бос болса)'];
const ZOP_PLIST = 'Өнім тізімі';   // справочник: настоящие названия из «All Trainings and Challanges» (для выпадающего списка)
// те же правила, что в карточке (zopCardClient_): нижний регистр без кавычек, «+» — все слова, «;» — варианты
const zopNoComma_ = s => String(s || '').replace(/\s*,\s*/g, '; ').trim();   // запятая — разделитель множественного выбора в Sheets
const zopLc_ = s => String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[«»"„“”]/g, '');
const zopKeyMatch_ = (keys, name) => { const n = zopLc_(name); return String(keys || '').split(';').map(x => x.trim()).filter(Boolean).some(alt => alt.split('+').map(w => zopLc_(w.trim())).filter(Boolean).every(w => n.indexOf(w) >= 0)); };
const zopShortName_ = s => {
  const parts = String(s || '').split(' · '), tariff = (parts[1] || '').trim();
  let t = parts[0].split(' / ')[0].replace(/\[[^\]]*\]/g, '').replace(/\d+\s*-?\s*(сынып|класс)\.?/gi, '')
    .replace(/\d+\s*-?\s*вариант[^\s)]*/gi, '').replace(/[«»"„“”]/g, '').replace(/\s+([,.;])/g, '$1').replace(/([,.;])(\s*[,.;])+/g, '$1').replace(/^[\s,.;|:-]+|[\s,.;|:-]+$/g, '').replace(/\s{2,}/g, ' ').trim();
  return (t || '—') + (tariff ? ' · ' + tariff : '');
};
const zopSecGuess_ = (name, from) => { const n = zopLc_(name);
  return /алтын\s*са[қк]а/.test(n) ? (from === 's' || /тренинг|дайындық|миникурс|нұсқа/.test(n) ? 'Алтын сақа' : 'Алтын сақа · пробный тест') : /алтын\s*түлек/.test(n) ? 'Алтын түлек' : /\bnis\b/.test(n) ? 'NIS'
    : /english|ағылшын|get ready/.test(n) ? 'Ағылшын' : /тоқсан|четверт|мега/.test(n) ? (/мега|(^|[^a-zа-я])(i|1)\s*[-–]\s*(iv|4)\b/.test(n) ? 'Математика · Мега қайталау' : 'Математика · Тоқсандық') : 'Басқа өнімдер'; };
// Статус счёта платформы → paid | unpaid | rejected | unknown. Подстрочный поиск запрещён: «unpaid» содержит «paid».
// «2» — код оплаченного счёта (в таблице часть ячеек в формате даты: 2 → «01.01.1900», zopStatus_ возвращает «2»).
// Пустой или незнакомый статус = unknown: такая строка не считается подтверждённой оплатой и показывается отдельно.
const ZOP_PAY = {
  paid: ['2', 'paid', 'success', 'succeeded', 'successful', 'completed', 'complete', 'approved', 'оплачен', 'оплачено', 'оплачена', 'төленді', 'төленген', 'сәтті'],
  unpaid: ['unpaid', 'not paid', 'notpaid', 'pending', 'awaiting', 'waiting', 'new', 'created', 'issued', 'не оплачен', 'не оплачено', 'ожидает оплаты', 'төленбеген', 'күтілуде'],
  rejected: ['failed', 'fail', 'error', 'declined', 'cancelled', 'canceled', 'refunded', 'refund', 'expired', 'отменен', 'отменён', 'отклонен', 'отклонён', 'возврат', 'қайтарылды', 'бас тартылды']
};
const zopPayNorm_ = v => String(v == null ? '' : v).trim().toLowerCase().replace(/ё/g, 'е').replace(/[_\-]+/g, ' ').replace(/\s+/g, ' ');
function zopPayStatus_(v) {
  let t = zopPayNorm_(v); if (!t) return 'unknown';
  if (/^mon jan 01 1900/.test(t) || /^1900 01 01/.test(t) || t === '01.01.1900') t = '2';   // legacy: code 2 shown as a date
  for (const k of ['paid', 'unpaid', 'rejected']) if (ZOP_PAY[k].indexOf(t) >= 0) return k;
  return 'unknown';
}
const zopPaid_ = st => zopPayStatus_(st) === 'paid';
// Строка куба = подтверждённая оплата этого дня? Лист «Оқушы тізімі» — да; платформа — только с invoice, статусом paid и датой оплаты (dsrc 0).
// Старые кубы (schemaVersion 1) без колонок происхождения: dsrc считается 0 — как было.
const zopDsrc_ = r => r.length >= 13 ? r[9] : 0, zopAsrc_ = r => r.length >= 13 ? r[10] : 0;
function zopSkipReason_(r, ST) {
  if (r[0] !== 'p') return '';
  if (r[7] !== 1) return 'invoice жоқ';
  const ps = zopPayStatus_(ST[r[6]] || '');
  if (ps !== 'paid') return {unpaid: 'төленбеген', rejected: 'бас тартылған / қайтарылған', unknown: 'статус белгісіз'}[ps];
  if (zopDsrc_(r) !== 0) return 'төлем күні жоқ (шот күні)';
  return '';
}
/** Pure: настоящие названия продуктов платформы (как в «All Trainings and Challanges»): название · тип · цены · последняя продажа · продано за 90 дней. */
function zopProductList_(data) {
  const N = data.dict.name, TY = data.dict.ptype, ST = data.dict.pstatus, last = String(data.snapshot || '').slice(0, 10);
  const from = last ? new Date(Date.parse(last + 'T00:00:00Z') - 90 * 864e5).toISOString().slice(0, 10) : '', m = {};
  data.cube.forEach(r => { if (r[0] !== 'p' || zopSkipReason_(r, ST)) return; const nm = String(N[r[3]] || '').trim(); if (!nm || last && r[1] > last) return;
    const e = m[nm] || (m[nm] = {nm, type: TY[r[5]] || '', prices: {}, last: '', n90: 0}); e.prices[r[4]] = 1; if (r[1] > e.last) e.last = r[1]; if (r[1] >= from) e.n90 += r[r.length - 2]; });
  return Object.values(m).sort((a, b) => (b.last > a.last ? 1 : b.last < a.last ? -1 : 0) || a.nm.localeCompare(b.nm))
    .map(e => [zopNoComma_(e.nm), e.type, Object.keys(e.prices).map(Number).sort((a, b) => a - b).join(', '), e.last, e.n90]);
}
/** Лист «Өнім тізімі» (в этой таблице): переписывается, только если список изменился. */
function zopWriteProductList_(data) {
  if (!data.rows || !data.rows.platform) return 0;   // без данных платформы список не трогаем — иначе он сотрётся
  const ss = zopHome_(), rows = zopProductList_(data), head = ['Нақты атауы (платформа)', 'Түрі', 'Бағалары, ₸', 'Соңғы сатылым', '90 күнде сатылды', 'Есепте қай жолға (тізімнен — «Өнімдер» C)', 'Бағыт (автомат; қате болса — түзетіңіз)'];
  let sh = ss.getSheetByName(ZOP_PLIST);
  if (!sh) { sh = ss.insertSheet(ZOP_PLIST); sh.setFrozenRows(1); }
  [520, 100, 160, 120, 130, 280, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const old = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues() : [], keep = {}, keepG = {};
  old.forEach(r => { if (String(r[5] || '').trim()) keep[String(r[0]).trim()] = r[5]; if (String(r[6] || '').trim()) keepG[String(r[0]).trim()] = r[6]; });   // ваш выбор в F и G сохраняется по названию
  const out = rows.map(r => r.concat([keep[String(r[0]).trim()] || '', keepG[String(r[0]).trim()] || zopSecGuess_(r[0], 'p')]));
  if (old.length > out.length) sh.getRange(out.length + 2, 1, old.length - out.length, 7).clearContent();
  sh.getRange(1, 1, 1, 7).setValues([head]).setFontWeight('bold').setBackground('#eef1f7').setWrap(true);
  if (out.length) sh.getRange(2, 1, out.length, 7).setValues(out);
  if (!sh.getRange('G2').getDataValidation()) sh.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ZOP_FAMILIES, true).setAllowInvalid(true).build());
  const prod = ss.getSheetByName(ZOP_SALES.PROD);
  if (prod && !sh.getRange('F2').getDataValidation()) sh.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(prod.getRange('C2:C'), true).setAllowInvalid(true).build()).setBackground('#fff8e1');
  try { zopApplyProductSections_(); } catch (err) { console.warn('Өнім бөлімдері: ' + err.message); }
  return out.length;
}
/** Pure: строки нового листа «Өнімдер» — план из «Жоспар» сверху, затем каждое настоящее название, которое покупали за 90 дней, затем цены из «Оқушы тізімі». */
function zopProductSeed_(data) {
  // Правила согласованы 10.10.2026. Порядок = приоритет: продажа попадает в первую подходящую строку сверху.
  //  • Алтын сақа — все продукты, кроме пробных тестов (пробные — отдельным разделом, в план Алтын сақа не входят)
  //  • Математика — Мега қайталау: курс «I-IV тоқсан» (и «Мегақайталау») за 13 500 ₸
  //  • Математика — Күндік: тоқсандық курсы и Мега қайталау за 1 190 ₸
  //  • Математика — Тоқсандық: тоқсандық курсы за 8 400 ₸; остальные цены математики — «Математика — басқа»
  const P = (data.plan && data.plan.products) || [], planOf = (name, def) => { const x = P.find(p => p.name === name); return x && x.plan != null && x.plan !== '' ? x.plan : def; };
  const fmt = n => String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  const AS_TR = 'алтын сақа тренинг;алтын сака тренинг', MATH = 'тоқсан;четверть;мега', ENG = 'english;ағылшын;get ready';
  const R = (sec, real, name, price, always, plan, coef, keys, note, pp) => [sec, real, name, price == null ? '' : price, always ? 'Иә' : 'Жоқ', plan == null ? '' : plan, coef && coef !== 1 ? coef : '', keys, note || '', pp == null ? '' : pp];
  const head = [
    R('Негізгі өнімдер', '', 'Ағылшын — Тоқсандық', null, true, planOf('Ағылшын — Тоқсандық', 4), 1, 'english+тоқсандық;english+айлық;ағылшын+тоқсандық;ағылшын+айлық;get ready+тоқсандық', 'ағылшын, тариф «Тоқсандық» / «3 Айлық», кез келген баға · математикадан жоғары тұруы керек', 8400),
    R('Негізгі өнімдер', '', 'Ағылшын — Күндік', null, true, planOf('Ағылшын — Күндік', 30), 1, 'english+күндік;ағылшын+күндік;get ready+күндік', 'ағылшын, тариф «Күндік», кез келген баға', 1190),
    R('Ағылшын — басқа', '', 'Ағылшын — басқа', null, false, null, 1, ENG + ';nis', 'қалған ағылшын өнімдері (NIS те)'),
    R('Негізгі өнімдер', '', 'Математика — Мега қайталау', 13500, true, planOf('Математика — Мега қайталау', 10), 1, 'i-iv тоқсан;i–iv тоқсан;1-4 тоқсан;мега', 'I-IV тоқсан курсы 13 500 ₸'),
    R('Негізгі өнімдер', '', 'Математика — Тоқсандық', 8400, true, planOf('Математика — Тоқсандық', 4), 1, 'тоқсан;четверть', 'тоқсандық курс 8 400 ₸'),
    R('Негізгі өнімдер', '', 'Математика — Күндік', 1190, true, planOf('Математика — Күндік', 30), 1, MATH, 'тоқсандық және Мега қайталау 1 190 ₸'),
    R('Математика — басқа', '', 'Математика — басқа баға', null, false, null, 1, MATH, 'математика басқа бағамен (500, 1 000, 5 000 …)'),
    R('Алтын түлек', '', 'Алтын түлек', null, false, null, 1, 'алтын түлек;алтын тулек', 'Алтын сақаға кірмейді'),
    R('Алтын сақа · пробный тест', '', 'Пробный тест — мектепішілік кезең', null, false, null, 1, 'пробный+мектепішілік', 'пробныйлар Алтын сақа жоспарына кірмейді'),
    R('Алтын сақа · пробный тест', '', 'Пробный тест — аудандық кезең', null, false, null, 1, 'пробный+аудандық', ''),
    R('Алтын сақа · пробный тест', '', 'Пробный тест — басқа', null, false, null, 1, 'пробный тест', '«9 нұсқа / 9 пробников» — бұл Алтын сақа тренажері, пробныйға кірмейді'),
    R('Алтын сақа', '', 'Барлығы', null, false, planOf('Барлығы', 67), 1, '', 'Алтын сақаның күндік жоспары'),
    R('Алтын сақа', '', 'Бір тренинг', 5600, true, null, 1, AS_TR, '«Оқушы тізімі» + платформа'),
    R('Алтын сақа', '', 'КОМБО', 24900, true, null, 1, AS_TR, ''),
    R('Алтын сақа', '', 'Тренажер', 13500, true, null, 1, AS_TR, ''),
    R('Алтын сақа', '', 'Мини-курс', 16900, true, null, 1, AS_TR, ''),
    R('Алтын сақа', '', 'Толық пакет (акция)', 13440, true, null, 1, AS_TR, ''),
    R('Алтын сақа', '', '2 тренинг −20% (рассылка)', 7840, true, null, 2, AS_TR, '1 төлем = 2 дана')];
  const rules = head.filter(r => r[7]).map(r => ({keys: r[7], price: r[3]}));
  // всё остальное, что покупали за 90 дней и что попадает в Алтын сақа: платформа — по настоящему названию, «Оқушы тізімі» — по цене
  const N = data.dict.name, ST = data.dict.pstatus, TY = data.dict.ptype;
  const last = String(data.snapshot || '').slice(0, 10), from = new Date(Date.parse(last + 'T00:00:00Z') - 90 * 864e5).toISOString().slice(0, 10);
  const realM = {}, sheetM = {};
  data.cube.forEach(r => {
    if (r[1] < from || r[1] > last) return;
    if (zopSkipReason_(r, ST)) return;
    const name = String(N[r[3]] || '').trim(), price = r[4], n = r[r.length - 2], hay = name + ' ' + (r[0] === 'p' ? TY[r[5]] || '' : '');
    if (!name || !/алтын\s*са[қк]а/i.test(zopLc_(name))) return;
    if (rules.some(x => zopKeyMatch_(x.keys, hay) && (x.price === '' || +x.price === +price))) return;
    if (r[0] === 'p') { const e = realM[name] || (realM[name] = {name, n: 0}); e.n += n; }
    else { const k = name + '|' + price, e = sheetM[k] || (sheetM[k] = {name, price, n: 0}); e.n += n; }
  });
  const real = Object.values(realM).sort((a, b) => b.n - a.n).map(x => R('Алтын сақа', zopNoComma_(x.name), '', null, false, null, 1, '', '90 күнде ' + x.n));
  const sheet = Object.values(sheetM).sort((a, b) => b.n - a.n).map(x => R('Алтын сақа', '', '', x.price, false, null, 1, x.name.toLowerCase(), `«Оқушы тізімі»: «${x.name}» ${fmt(x.price)} ₸ · 90 күнде ${x.n} — B-дан нақты өнімді таңдаңыз`));
  return head.concat(real, sheet, [R('Алтын сақа', '', 'Алтын сақа — басқа', null, false, null, 1, 'алтын сақа;алтын сака', 'жаңа Алтын сақа өнімдері (пробныйдан басқа)')]);
}
/** Запустите вручную, чтобы пересобрать «Өнімдер» по согласованным правилам. Текущий лист не удаляется — переименовывается в «Өнімдер (ескі …)». */


/* =================== Гипотезы (Strategy Repository) и служебные листы =================== */
const ZOP_MAN_ALIAS = {'Dinara': 'Айгерім'};   // amoCRM аккаунты → адам (иесі, 15.10.2026: «Динара — бұл Айгерім»); Python-да MAN_ALIAS, Атласта MAN_ALIAS
const ZOP_TARIFF = 'Тарифтер', ZOP_HYP = 'Гипотезалар', ZOP_COST_LOG = 'Расходы журналы';
const ZOP_TARIFF_HEAD = ['Поставщик', 'Хабарлама түрі', 'Басталуы', 'Аяқталуы', 'Бағасы', 'Валюта', 'Дереккөз (шот / бет)'];
const ZOP_HYP_HEAD = ['ID', 'Сегмент', 'Өнім', 'Ұсыныс', 'Неге бұл сегмент', 'Дәлел (дерек)', 'Эксперимент', 'Табыс критерийі', 'Тәуекел', 'Статус', 'Нәтиже', 'Шешім', 'Күні', 'Гипотеза', 'Рассылкалар'];
/** Pure: лист «Гипотезалар» → список (только текст сотрудников; персональные данные туда не пишут — см. подсказку на листе). */
function zopHypotheses_(rows) {
  if (!rows || rows.length < 2) return [];
  const head = rows[0].map(x => String(x || '').trim()), ix = {}; head.forEach((h, i) => { ix[h] = i; });
  const g = (r, h) => ix[h] === undefined ? '' : (r[ix[h]] instanceof Date ? r[ix[h]].toISOString().slice(0, 10) : String(r[ix[h]] == null ? '' : r[ix[h]]).trim());
  return rows.slice(1).map(r => Object.fromEntries(ZOP_HYP_HEAD.map(h => [h, g(r, h)]))).filter(h => h['Гипотеза'] || h['Сегмент'] || h['Өнім'] || h['Эксперимент']);
}
/** Pure: проверка новой/изменённой строки расхода из формы. Возвращает {ok, errors, row}. */
function zopCostValidate_(o, tariffs, existing) {
  const errors = [], d = /^\d{4}-\d{2}-\d{2}$/.test(o.date || '') ? o.date : null;
  if (!d) errors.push('Күні керек');
  if (ZOP_COST_CATS.indexOf(o.cat) < 0) errors.push('Санат тізімнен таңдалуы керек');
  const sum = o.sum === '' || o.sum == null ? null : +o.sum, msgs = o.msgs === '' || o.msgs == null ? null : +o.msgs, cur = String(o.cur || 'KZT').toUpperCase(), rate = o.rate === '' || o.rate == null ? null : +o.rate;
  if (sum != null && !(sum >= 0 || o.cat === 'Adjustment')) errors.push('Сома теріс болмауы керек');
  if (sum == null && msgs == null) errors.push('Сома немесе хабарлама саны керек');
  if (cur !== 'KZT' && !(rate > 0)) errors.push('Валюта ' + cur + ' — курс керек');
  if (o.camp) { const al = zopCostAlloc_(o.camp, sum == null ? null : sum * (cur === 'KZT' ? 1 : rate)); if (al.err) errors.push('Кампания: ' + al.err); }
  const key = [d, o.cat, o.prov, o.sub || '', sum, cur, o.pay || '', o.doc || '', o.camp || ''].join('|');
  const dup = (existing || []).some(e => e.key === key && e.id !== o.id);
  return {ok: !errors.length, errors, dup, key, row: {d, sum, msgs, cur, rate}};
}
/* =================== Комментарий к рассылке с сайта ===================
   Сайт (статичный, Vercel) шлёт POST text/plain JSON {action:'mailNote', code, day:'YYYY-MM-DD', funnel, note}.
   Пишется только столбец «Комментарий» листа «Рассылки» главной таблицы (строка того же дня и воронки; нет строки — добавляется).
   Защита: код доступа (свойство скрипта ZOP_NOTE_CODE; владелец задаёт через меню «EduApp → Сайттағы пікір коды»),
   после 20 неверных кодов за 10 минут запись закрыта; текст ≤ 500 символов, без формул. Наружу не отдаётся ничего, кроме ok/ошибки. */
const ZOP_NOTE_MAX = 500, ZOP_NOTE_FUNNELS = ['Рассылка', 'Рассылка 5-6 класс Алтын сақа'];
function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = b && b.action === 'mailNote' ? zopSaveMailNote_(b) : b && b.action === 'mailText' ? zopSaveMailText_(b) : b && b.action === 'mailName' ? zopSaveMailName_(b) : {ok: false, error: 'action'};
  } catch (err) { out = {ok: false, error: 'server', message: String(err && err.message || err).slice(0, 200)}; }
  return zopOut_(JSON.stringify(out));
}
/** Pure: код совпадает (сравнение без раннего выхода); пустой код в свойствах — запись выключена. */
function zopCodeOk_(code, want) {
  code = String(code == null ? '' : code); want = String(want == null ? '' : want);
  if (!want || want.length < 4 || code.length !== want.length) return false;
  let d = 0; for (let i = 0; i < want.length; i++) d |= code.charCodeAt(i) ^ want.charCodeAt(i);
  return d === 0;
}
/** Pure: текст комментария — без управляющих символов, ≤ 500, формула превращается в текст (апостроф). */
function zopNoteClean_(s) {
  const t = String(s == null ? '' : s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, '').replace(/\r\n?/g, '\n').trim().slice(0, ZOP_NOTE_MAX);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}
/** Pure: куда писать. values — лист «Рассылки», ymdOf(cell) → 'YYYY-MM-DD'. → {row (0-based, -1 = новая), cA, cF, cN, cName, addNoteCol}. */
function zopNoteTarget_(values, ymdOf, ymd, funnel, time, hmOf) {   // 16.10: time 'ЧЧ:ММ' — бір күнде бірнеше рассылка (импорт тегінің уақыты)
  const head = (values[0] || []).map(x => String(x || '').trim().toLowerCase()), col = re => head.findIndex(h => re.test(h));
  const cA = col(/начал|^дата$/), cF = col(/воронк/), cName = col(/^рассылка|назв/), cTm = col(/уақыт|время/);
  let cN = col(/коммент/), addNoteCol = false; if (cN < 0) { cN = Math.max(head.length, 1); addNoteCol = true; }
  const fun = v => { const t = String(v || '').toLowerCase(); return /5\s*[-–]\s*6/.test(t) ? ZOP_NOTE_FUNNELS[1] : /рассыл/.test(t) ? ZOP_NOTE_FUNNELS[0] : ''; };
  const hm = v => hmOf ? hmOf(v) : String(v == null ? '' : v).trim();
  const same = i => ymdOf(values[i][cA]) === ymd && (cF < 0 ? funnel === ZOP_NOTE_FUNNELS[0] : fun(values[i][cF]) === funnel);
  let row = -1, setTime = false;
  if (cA >= 0 && time && cTm >= 0) for (let i = 1; i < values.length; i++) if (same(i) && hm(values[i][cTm]) === time) { row = i; break; }
  if (cA >= 0 && row < 0) for (let i = 1; i < values.length; i++) if (same(i) && (!time || cTm < 0 || !hm(values[i][cTm]))) { row = i; setTime = !!time; break; }
  return time || cTm >= 0 ? {row, cA, cF, cN, cName, addNoteCol, cTm, setTime} : {row, cA, cF, cN, cName, addNoteCol};
}
/** Pure: 'ЧЧ:ММ' немесе ''. */
function zopHm_(v) { const m = String(v == null ? '' : v).match(/^(\d{1,2}):(\d{2})$/); return m && +m[1] < 24 && +m[2] < 60 ? ('0' + m[1]).slice(-2) + ':' + m[2] : ''; }
/** «Рассылки» парағындағы бір рассылканың жолы (күн + воронка + уақыт). Жоқ болса — жаңа жол; «Уақыты» бағаны жоқ болса — қосылады. → 1-ден басталатын нөмір. */
function zopMailRowFor_(sh, tz, ymd, funnel, time) {
  const values = sh.getDataRange().getValues(), m = ymd.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const ymdOf = v => { if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd'); const x = String(v || '').match(/(\d{1,2})[./](\d{1,2})[./](\d{4})/); return x ? `${x[3]}-${('0' + x[2]).slice(-2)}-${('0' + x[1]).slice(-2)}` : ''; };
  const hmOf = v => v instanceof Date ? Utilities.formatDate(v, tz, 'HH:mm') : zopHm_(String(v == null ? '' : v).trim().replace('.', ':'));
  const t = zopNoteTarget_(values, ymdOf, ymd, funnel, time, hmOf);
  let cTm = t.cTm == null ? -1 : t.cTm;
  if (time && cTm < 0) { cTm = Math.max(values[0].length, 1); sh.getRange(1, cTm + 1).setValue('Уақыты').setFontWeight('bold'); }
  if (t.row >= 0) { if (time && t.setTime) sh.getRange(t.row + 1, cTm + 1).setNumberFormat('@').setValue(time); return t.row + 1; }
  if (t.cA < 0) throw new Error('«Рассылки» парағында «Дата начала» бағаны жоқ');
  const width = Math.max(values[0].length, cTm + 1, 1), row = Array(width).fill('');
  row[t.cA] = new Date(+m[1], +m[2] - 1, +m[3]); if (t.cF >= 0) row[t.cF] = funnel; if (t.cName >= 0) row[t.cName] = 'Рассылка ' + m[3] + '.' + m[2] + '.' + m[1] + (time ? ' ' + time : '');
  if (time) row[cTm] = time;
  const r = sh.getLastRow() + 1; if (time) sh.getRange(r, cTm + 1).setNumberFormat('@'); sh.getRange(r, 1, 1, width).setValues([row]); return r;
}
/** Pure: ОП кестесінің төлемі → amoCRM сделкасы. Кілт: телефон (соңғы 10 цифр: сделкада «Доступ номер», кестеде «телефон/номер» бағаны) немесе «Чек».
   Әр төлем — сол кілті бар, төлем күнінен бұрын (не сол күні) құрылған ЕҢ СОҢҒЫ сделкаға. Нәтиже: төленген сделка жолдарының жиыны + жалпы сандар (кілт мәндері шықпайды). */
function zopPayMatch_(leads, payRows) {
  const out = {paid: new Set(), sum: new Map(), stats: {pays: 0, matched: 0, byPhone: 0, byCheck: 0, noKey: 0}};   // sum: сделка жолы → төлем сомасы (рассылка тегі бойынша ақша үшін)
  if (!leads.length || !payRows.length) return out;
  const lh = (leads[0] || []).map(x => String(x || '').trim()), ph = (payRows[0] || []).map(x => String(x || '').trim().toLowerCase());
  const li = n => lh.indexOf(n), pi = re => ph.findIndex(h => re.test(h));
  const cPh = li('Доступ номер'), cCk = li('Чек'), cCr = li('Создана'), pPh = ph.findIndex(h => /телефон|нөмір|номер|phone|whatsapp|ватсап/.test(h) && !/чек|сделк|^id|договор|шот/.test(h)), pCk = pi(/^чек|чек номер|номер чек/), pDt = pi(/^дата$|^күні$/), pAm = pi(/^төлем сомасы$/), pPr = pi(/^бағасы$/);
  const phone = v => { const d = String(v == null ? '' : v).replace(/\D/g, ''); return d.length >= 10 ? d.slice(-10) : ''; };
  const check = v => { const t = String(v == null ? '' : v).trim().toLowerCase(), runs = (t.match(/\d{6,}/g) || []).sort((a, b) => b.length - a.length); return runs.length ? 'd' + runs[0] : t.replace(/\s+/g, ''); };   // 15.10: «№ 123456», «QR-123456», сілтеме — бірдей кілт (ең ұзын 6+ цифр)
  const dnum = v => { if (v instanceof Date && !isNaN(v.getTime())) return Math.floor(v.getTime() / 864e5); if (typeof v === 'number' && v > 30000 && v < 80000) return Math.floor(v) - 25569;   // серийная дата Sheets
    const t = String(v == null ? '' : v); let m = t.match(/(\d{1,2})[./](\d{1,2})[./](\d{4})/); if (m) return Math.floor(Date.UTC(+m[3], +m[2] - 1, +m[1]) / 864e5);
    m = t.match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? Math.floor(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5) : null; };
  out.stats.cols = {phone: pPh >= 0 ? ph[pPh] : null, check: pCk >= 0 ? ph[pCk] : null, date: pDt >= 0 ? ph[pDt] : null, crmPhone: cPh >= 0, crmCheck: cCk >= 0};   // тек баған атаулары (дерек емес) — сәйкестендіруді тексеру үшін
  const byKey = {}; let c0 = null;   // кілт → [{row, d}]; c0 — CRM-дағы ең ерте сделка күні (одан бұрынғы төлемдер CRM-да болуы мүмкін емес)
  leads.slice(1).forEach(r => { const d = cCr >= 0 ? dnum(r[cCr]) : null; if (d == null) return; if (c0 == null || d < c0) c0 = d;
    const k1 = cPh >= 0 ? phone(r[cPh]) : '', k2 = cCk >= 0 ? check(r[cCk]) : '';
    if (k1) { out.stats.crmPhones = (out.stats.crmPhones || 0) + 1; (byKey['p' + k1] = byKey['p' + k1] || []).push({r, d}); } if (k2 && k2.length >= 4) { out.stats.crmChecks = (out.stats.crmChecks || 0) + 1; (byKey['c' + k2] = byKey['c' + k2] || []).push({r, d}); } });
  const pick = (list, d) => { let best = null, late = null; (list || []).forEach(x => { if (x.d <= d + 1) { if (!best || x.d > best.d) best = x; } else if (x.d <= d + 7 && (!late || x.d < late.d)) late = x; }); return best || late; };   // сделка төлемнен кейін 7 күнге дейін ашылса да
  out.stats.paysCov = 0; out.stats.matchedCov = 0;
  payRows.slice(1).forEach(p => { const d = pDt >= 0 ? dnum(p[pDt]) : null; if (d == null) return; out.stats.pays++; const cov = c0 != null && d >= c0; if (cov) out.stats.paysCov++;
    const k2 = pCk >= 0 ? check(p[pCk]) : '', k1 = pPh >= 0 ? phone(p[pPh]) : ''; if (k1) out.stats.payPhones = (out.stats.payPhones || 0) + 1; if (k2 && k2.length >= 4) out.stats.payChecks = (out.stats.payChecks || 0) + 1;
    if (!k1 && !(k2 && k2.length >= 4)) { out.stats.noKey++; return; }
    let hit = k2 && k2.length >= 4 ? pick(byKey['c' + k2], d) : null, how = 'byCheck';
    if (!hit && k1) { hit = pick(byKey['p' + k1], d); how = 'byPhone'; }
    const amt = v => { if (typeof v === 'number') return isFinite(v) ? v : 0; const t = String(v == null ? '' : v).replace(/[^\d.,]/g, '').replace(',', '.'); return t ? +t || 0 : 0; };
    if (hit) { out.paid.add(hit.r); out.stats.matched++; if (cov) out.stats.matchedCov++; out.stats[how]++; const a = (pAm >= 0 ? amt(p[pAm]) : 0) || (pPr >= 0 ? amt(p[pPr]) : 0); out.sum.set(hit.r, (out.sum.get(hit.r) || 0) + a); } });
  return out;
}
/** Pure: «кім не толтырмады» — соңғы 45 күн, жаңалары бірінші, ең көбі 3000 жол. */
function zopMissRows_(raw) { const mx = raw.reduce((m, r) => Math.max(m, r[1]), -1e9); return raw.filter(r => r[1] >= mx - 45).sort((a, b) => b[1] - a[1] || b[0] - a[0]).slice(0, 3000); }
/** Pure: медиа ұяшығы → тек https сілтемелер (ең көбі 4), бос орын / жол / үтір бойынша бөлінеді. */
function zopMediaList_(v) { return String(v || '').split(/[\s,]+/).map(x => x.trim()).filter(x => /^https:\/\/[^\s<>"]{4,500}$/.test(x)).slice(0, 4); }
/** Pure: батырмалар ұяшығы → [{t, u}] — әр жолға біреу: «Мәтін» немесе «Мәтін | https://…»; ең көбі 3, мәтін 40 таңбаға дейін. */
function zopBtnList_(v) {
  return String(v || '').split(/\r?\n/).map(l => { const [t, u] = l.split('|').map(x => String(x || '').trim()); const tt = zopTextClean_(t).replace(/^'/, '').slice(0, 40);
    return tt ? (/^https:\/\/[^\s<>"]{4,500}$/.test(u || '') ? {t: tt, u} : {t: tt}) : null; }).filter(Boolean).slice(0, 3);
}
/** Pure: рассылка мәтіні — басқару таңбаларынсыз, 2000 таңбаға дейін, формула болып кетпейді. */
function zopTextClean_(s) { let t = String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim().slice(0, 2000); if (/^[=+\-@]/.test(t)) t = "'" + t; return t; }
/** Сайттан рассылка мәтіні мен медиа сілтемесі → «Рассылки» парағы («Мәтін», «Медиа (сілтеме)» бағандары жоқ болса — қосылады). Код — пікірдікімен бірдей. */
function zopSaveMailText_(b) {
  const cache = CacheService.getScriptCache(), fails = +(cache.get('zop-note-fail') || 0);
  if (fails >= 20) return {ok: false, error: 'locked'};
  if (!zopCodeOk_(b.code, PropertiesService.getScriptProperties().getProperty('ZOP_NOTE_CODE'))) { cache.put('zop-note-fail', String(fails + 1), 600); Utilities.sleep(700); return {ok: false, error: 'code'}; }
  const ymd = String(b.day || ''), m = ymd.match(/^(\d{4})-(\d{2})-(\d{2})$/); if (!m) return {ok: false, error: 'day'};
  const funnel = ZOP_NOTE_FUNNELS.indexOf(b.funnel) >= 0 ? b.funnel : null; if (!funnel) return {ok: false, error: 'funnel'};
  const text = zopTextClean_(b.text), media = zopMediaList_(b.media).join('\n'), btns = zopBtnList_(b.buttons), btnCell = btns.map(x => x.u ? x.t + ' | ' + x.u : x.t).join('\n'), lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss = zopHome_(), tz = ss.getSpreadsheetTimeZone();
    let sh = ss.getSheetByName(ZOP.MAIL);
    if (!sh) { sh = ss.insertSheet(ZOP.MAIL); sh.getRange(1, 1, 1, ZOP_MAIL_HEAD.length).setValues([ZOP_MAIL_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); }
    let values = sh.getDataRange().getValues();
    const head = () => (values[0] || []).map(x => String(x || '').trim().toLowerCase());
    const colOf = (re, title) => { let i = head().findIndex(h => re.test(h)); if (i < 0) { i = Math.max(values[0].length, 1); sh.getRange(1, i + 1).setValue(title).setFontWeight('bold'); values = sh.getDataRange().getValues(); } return i; };
    const cT = colOf(/мәтін|текст/, 'Мәтін'), cM = colOf(/медиа|сурет|видео|media/, 'Медиа (сілтеме)'), cB = colOf(/батырма|кнопк|button/, 'Батырмалар'), cC = b.cost == null || b.cost === '' ? -1 : colOf(/расход|затрат|сумм/, 'Доп. расход, ₸ (кроме WhatsApp)');
    const r = zopMailRowFor_(sh, tz, ymd, funnel, zopHm_(b.time));   // 16.10: рассылка уақытымен
    sh.getRange(r, cT + 1).setNumberFormat('@').setValue(text); sh.getRange(r, cM + 1).setNumberFormat('@').setValue(media); sh.getRange(r, cB + 1).setNumberFormat('@').setValue(btnCell);
    const cl = (re, title, v) => { if (v == null) return; const ci = colOf(re, title); sh.getRange(r, ci + 1).setNumberFormat('@').setValue(zopNoteClean_(String(v).replace(/\n/g, ' ')).slice(0, 60)); };   // 16.10 иесі: «по классам и по группам кз / ру»
    cl(/^сынып|^класс/, 'Сынып', b.cls); cl(/тіл тоб|^тіл$|язык/, 'Тіл тобы', b.lang);
    if (cC >= 0) { const cv = +String(b.cost).replace(/[^\d.,]/g, '').replace(',', '.'); if (isFinite(cv) && cv >= 0 && cv < 1e9) sh.getRange(r, cC + 1).setValue(Math.round(cv)); }   // осы рассылканың қосымша шығыны
  } finally { lock.releaseLock(); }
  try { cache.remove('zop-n'); } catch (_) { /* кэш обновится сам */ }
  return {ok: true, text, media: media ? media.split('\n') : [], buttons: btns};
}
/** Сайттан рассылка атауы → «Рассылки» парағы, «Тема» бағаны (иесі 16.10: «возможность менять названия»). */
function zopSaveMailName_(b) {
  const cache = CacheService.getScriptCache(), fails = +(cache.get('zop-note-fail') || 0);
  if (fails >= 20) return {ok: false, error: 'locked'};
  if (!zopCodeOk_(b.code, PropertiesService.getScriptProperties().getProperty('ZOP_NOTE_CODE'))) { cache.put('zop-note-fail', String(fails + 1), 600); Utilities.sleep(700); return {ok: false, error: 'code'}; }
  const ymd = String(b.day || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return {ok: false, error: 'day'};
  const funnel = ZOP_NOTE_FUNNELS.indexOf(b.funnel) >= 0 ? b.funnel : null; if (!funnel) return {ok: false, error: 'funnel'};
  const name = zopNoteClean_(String(b.name == null ? '' : b.name).replace(/\n/g, ' ')).slice(0, 80), lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss = zopHome_(), tz = ss.getSpreadsheetTimeZone();
    let sh = ss.getSheetByName(ZOP.MAIL);
    if (!sh) { sh = ss.insertSheet(ZOP.MAIL); sh.getRange(1, 1, 1, ZOP_MAIL_HEAD.length).setValues([ZOP_MAIL_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); }
    const r = zopMailRowFor_(sh, tz, ymd, funnel, zopHm_(b.time));
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(x => String(x || '').trim().toLowerCase());
    let cT = head.findIndex(h => /^тема/.test(h)); if (cT < 0) { cT = head.length; sh.getRange(1, cT + 1).setValue('Тема').setFontWeight('bold'); }
    sh.getRange(r, cT + 1).setNumberFormat('@').setValue(name);
  } finally { lock.releaseLock(); }
  try { cache.remove('zop-n'); } catch (_) { /* кэш обновится сам */ }
  return {ok: true, name};
}
function zopSaveMailNote_(b) {
  const cache = CacheService.getScriptCache(), fails = +(cache.get('zop-note-fail') || 0);
  if (fails >= 20) return {ok: false, error: 'locked'};
  if (!zopCodeOk_(b.code, PropertiesService.getScriptProperties().getProperty('ZOP_NOTE_CODE'))) { cache.put('zop-note-fail', String(fails + 1), 600); Utilities.sleep(700); return {ok: false, error: 'code'}; }
  const ymd = String(b.day || ''), m = ymd.match(/^(\d{4})-(\d{2})-(\d{2})$/); if (!m) return {ok: false, error: 'day'};
  const funnel = ZOP_NOTE_FUNNELS.indexOf(b.funnel) >= 0 ? b.funnel : null; if (!funnel) return {ok: false, error: 'funnel'};
  const note = zopNoteClean_(b.note), lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss = zopHome_(), tz = ss.getSpreadsheetTimeZone();
    let sh = ss.getSheetByName(ZOP.MAIL);
    if (!sh) { sh = ss.insertSheet(ZOP.MAIL); sh.getRange(1, 1, 1, ZOP_MAIL_HEAD.length).setValues([ZOP_MAIL_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); }
    const r = zopMailRowFor_(sh, tz, ymd, funnel, zopHm_(b.time));   // 16.10: рассылка уақытымен
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(x => String(x || '').trim().toLowerCase());
    let cN = head.findIndex(h => /коммент/.test(h)); if (cN < 0) { cN = head.length; sh.getRange(1, cN + 1).setValue('Комментарий').setFontWeight('bold'); }
    sh.getRange(r, cN + 1).setNumberFormat('@').setValue(note);
  } finally { lock.releaseLock(); }
  try { cache.remove('zop-n'); } catch (_) { /* кэш обновится сам */ }
  return {ok: true, note};
}
/** Меню таблицы: задать / сменить код, которым сайт сохраняет комментарии (пусто — запись с сайта выключена). */
function zopSetNoteCode() {
  zopRequireUi_();
  const ui = SpreadsheetApp.getUi(), r = ui.prompt('Сайттағы пікір коды', 'Кемінде 4 таңба. Бос қалдырсаңыз — сайттан жазу өшіріледі.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const v = String(r.getResponseText() || '').trim(), P = PropertiesService.getScriptProperties();
  if (!v) { P.deleteProperty('ZOP_NOTE_CODE'); ui.alert('Сайттан жазу өшірілді.'); return; }
  if (v.length < 4) { ui.alert('Код кемінде 4 таңба болсын.'); return; }
  P.setProperty('ZOP_NOTE_CODE', v); ui.alert('Код сақталды. Оны тек қызметкерлерге беріңіз.');
}

/* =================== Защита вызовов ===================
   Страницы веб-приложения (карточка) построены на HtmlService, а с них google.script.run может вызвать ЛЮБУЮ публичную функцию,
   и выполнится она от имени владельца. Поэтому:
   · функции, меняющие таблицу или раскрывающие её структуру, проверяют, что вызывающий — сам владелец (zopRequireOwner_);
   · функции меню требуют интерфейса таблицы (в веб-приложении его нет);
   · сохранение расхода принимает только одноразовый токен формы, открытой из меню таблицы. */
function zopRequireOwner_(who) {
  let a = '', e = ''; try { a = (who || Session).getActiveUser().getEmail(); e = (who || Session).getEffectiveUser().getEmail(); } catch (_) { /* no session */ }
  if (!a || a !== e) throw new Error('Бұл функцияны тек кесте иесі Apps Script редакторынан іске қоса алады');
}
function zopRequireUi_() { try { SpreadsheetApp.getUi(); } catch (_) { throw new Error('Тек Google-кесте мәзірінен'); } }
const ZOP_TOK = 'zop-cost-form';
function zopCheckToken_(tok, cache) { if (!tok) return false; const v = (cache || CacheService.getUserCache()).get(ZOP_TOK); return !!v && tok === v; }

/** Запустите один раз (владелец): создаёт листы «Тарифтер», «Гипотезалар», «Расходы журналы», добавляет недостающие колонки в «Расходы»
    (справа, без изменения данных), список категорий и меню «EduApp» (форма расхода). Ничего не удаляет. */
function zopSetupCostModule() {
  zopRequireOwner_();
  const ss = zopHome_(), made = [];
  const ensure = (name, head) => { let sh = ss.getSheetByName(name); if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1); made.push(name); } return sh; };
  ensure(ZOP_TARIFF, ZOP_TARIFF_HEAD); const hyp = ensure(ZOP_HYP, ZOP_HYP_HEAD); ensure(ZOP_COST_LOG, ['Уақыты', 'Кім', 'Әрекет', 'ID', 'Бұрынғы', 'Жаңа']);
  hyp.getRange('A1').setNote('Мұнда тек жалпы сегменттер мен идеялар жазылады — баланың аты, телефон сияқты жеке деректерді жазбаңыз.');
  const cost = ss.getSheetByName(ZOP.COST);
  if (cost) { const head = cost.getRange(1, 1, 1, Math.max(1, cost.getLastColumn())).getValues()[0].map(x => String(x || '').trim()), miss = ZOP_COST_HEAD_V2.filter(h => head.indexOf(h) < 0);
    if (miss.length) { cost.getRange(1, head.length + 1, 1, miss.length).setValues([miss]).setFontWeight('bold'); made.push('«Расходы»: +' + miss.join(', ')); }
    const ci = cost.getRange(1, 1, 1, cost.getLastColumn()).getValues()[0].map(String).indexOf('Категория') + 1;
    if (ci > 0) cost.getRange(2, ci, Math.max(1, cost.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ZOP_COST_CATS, true).setAllowInvalid(true).build()); }
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'zopMenu_')) ScriptApp.newTrigger('zopMenu_').forSpreadsheet(ss).onOpen().create();
  zopMenu_();
  Logger.log('Cost Control: ' + (made.length ? made.join('; ') : 'бәрі бар'));
}
function zopMenu_() { try { SpreadsheetApp.getUi().createMenu('EduApp').addItem('Шығын қосу', 'zopCostFormNew').addItem('Шығынды өзгерту (ID бойынша)', 'zopCostFormEdit').addItem('Сайттағы пікір коды', 'zopSetNoteCode').addSeparator().addItem('Дереккөздерді қосу (ОП миникурс, All Trainings)', 'zopSetupSourcesMenu').addToUi(); } catch (e) { /* no UI (trigger without a session) */ } }
function zopCostFormNew() { zopRequireUi_(); zopCostForm_(null); }
function zopCostFormEdit() { zopRequireUi_(); const ui = SpreadsheetApp.getUi(), r = ui.prompt('Шығынның ID-ін жазыңыз (соңғы баған)'); if (r.getSelectedButton() === ui.Button.OK) zopCostForm_(String(r.getResponseText()).trim()); }
function zopCostRows_() {
  const sh = zopHome_().getSheetByName(ZOP.COST); if (!sh) throw new Error('«Расходы» парағы жоқ — zopSetupCostModule іске қосыңыз');
  const v = sh.getDataRange().getValues(), head = v[0].map(x => String(x || '').trim()), ix = {}; head.forEach((h, i) => { ix[h] = i; });
  return {sh, v, head, ix};
}
function zopCostForm_(id) {
  const {v, ix} = zopCostRows_(); let cur = null;
  if (id) { const i = v.findIndex((r, k) => k && String(r[ix['ID']] || '').trim() === id); if (i < 0) { SpreadsheetApp.getUi().alert('ID табылмады: ' + id); return; } cur = Object.fromEntries(ZOP_COST_HEAD_V2.map(h => [h, ix[h] === undefined ? '' : (v[i][ix[h]] instanceof Date ? Utilities.formatDate(v[i][ix[h]], zopHome_().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : v[i][ix[h]])])); }
  const tok = Utilities.getUuid(); CacheService.getUserCache().put(ZOP_TOK, tok, 1800);   // 30 minutes, only for this user
  const html = HtmlService.createHtmlOutput(zopCostFormHtml_(cur, tok)).setWidth(460).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, id ? 'Шығынды өзгерту' : 'Шығын қосу');
}
function zopCostFormHtml_(cur, tok) {
  const f = (k, d) => cur && cur[k] !== '' && cur[k] != null ? String(cur[k]) : d;
  const esc = x => String(x).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'})[c]);
  const inp = (k, lb, type, val, extra) => `<label>${lb}<input name="${k}" type="${type}" value="${esc(val)}" ${extra || ''}></label>`;
  return `<style>body{font:14px system-ui;margin:0;padding:12px}label{display:block;margin:6px 0;font-weight:600;color:#172033}input,select,textarea{display:block;width:100%;box-sizing:border-box;margin-top:3px;padding:7px;border:1px solid #DDE3ED;border-radius:8px;font:inherit}
    .row{display:grid;grid-template-columns:1fr 1fr;gap:8px}button{margin-top:12px;padding:9px 14px;border:0;border-radius:8px;background:#3659D9;color:#fff;font-weight:700}#msg{margin-top:8px;color:#B42318;white-space:pre-wrap}small{color:#526077;font-weight:400}</style>
    <form id="f"><input type="hidden" name="id" value="${esc(f('ID', ''))}"><input type="hidden" name="tok" value="${esc(tok || '')}">
    <div class="row">${inp('date', 'Күні', 'date', f('Дата', new Date().toISOString().slice(0, 10)), 'required')}<label>Санат<select name="cat">${ZOP_COST_CATS.map(c => `<option${f('Категория', '') === c ? ' selected' : ''}>${c}</option>`).join('')}</select></label></div>
    <div class="row">${inp('prov', 'Поставщик', 'text', f('Поставщик', 'Wazzup'))}${inp('sub', 'Подписка / нөмір', 'text', f('Подписка', 'Рассылка'))}</div>
    <div class="row">${inp('sum', 'Сома <small>(бос — тариф бойынша)</small>', 'number', f('Сума, ₸', f('Сумма, ₸', '')), 'step="any"')}${inp('msgs', 'Хабарлама саны', 'number', f('Хабарлама саны', ''), 'step="1" min="0"')}</div>
    <div class="row">${inp('cur', 'Валюта', 'text', f('Валюта', 'KZT'))}${inp('rate', 'Курс <small>(KZT емес болса)</small>', 'number', f('Курс', ''), 'step="any"')}</div>
    <div class="row">${inp('tariff', 'Тариф / хабарлама түрі', 'text', f('Тариф', ''))}${inp('pay', 'Төлем тәсілі', 'text', f('Оплата', 'Stripe'))}</div>
    <div class="row">${inp('from', 'Кезең басы', 'date', f('Кезең басы', ''))}${inp('to', 'Кезең соңы', 'date', f('Кезең соңы', ''))}</div>
    ${inp('camp', 'Кампания <small>(А=30000; Б=20000 немесе А:60%; Б:40%)</small>', 'text', f('Кампания', ''))}
    ${inp('doc', 'Құжат (шот нөмірі / сілтеме)', 'text', f('Құжат', ''))}
    <label><input type="checkbox" name="ver" style="display:inline;width:auto"${/иә|да|yes|true/i.test(f('Тексерілді', '')) ? ' checked' : ''}> Тексерілді (шотпен расталды)</label>
    ${inp('det', 'Түсініктеме', 'text', f('Детали', ''))}
    <button type="submit">Сақтау</button><div id="msg"></div></form>
    <script>const F=document.getElementById('f'),M=document.getElementById('msg');let force=false;
    F.onsubmit=e=>{e.preventDefault();const o=Object.fromEntries(new FormData(F));o.ver=F.ver.checked;o.force=force;M.style.color='#526077';M.textContent='Сақталуда…';
    google.script.run.withSuccessHandler(r=>{if(r.ok){M.style.color='#146C2E';M.textContent='Сақталды · ID '+r.id;setTimeout(()=>google.script.host.close(),900);}else if(r.dup&&!force){M.style.color='#8A5300';M.textContent='Осындай жол бар (дубль болуы мүмкін). Бәрібір сақтау үшін тағы басыңыз.';force=true;}else{M.style.color='#B42318';M.textContent=r.errors.join('\\n');}})
    .withFailureHandler(err=>{M.style.color='#B42318';M.textContent=err.message;}).zopSaveCost(o);};</script>`;
}
/** Форма → строка «Расходы» (+ запись в «Расходы журналы»). Только для редакторов таблицы: запускается из меню внутри таблицы, не через веб-приложение. */
function zopSaveCost(o) {
  if (!o || !zopCheckToken_(o.tok)) throw new Error('Форманы Google-кестедегі «EduApp → Шығын қосу» мәзірінен ашыңыз');
  const lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    const {sh, v, head, ix} = zopCostRows_(), ss = zopHome_(), tz = ss.getSpreadsheetTimeZone();
    const miss = ZOP_COST_HEAD_V2.filter(h => ix[h] === undefined); if (miss.length) return {ok: false, errors: ['«Расходы» парағында бағандар жоқ: ' + miss.join(', ') + ' — zopSetupCostModule іске қосыңыз']};
    const iso = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x || '');
    const existing = v.slice(1).map(r => ({id: String(r[ix['ID']] || '').trim(), key: [iso(r[ix['Дата']]), r[ix['Категория']], r[ix['Поставщик']], r[ix['Подписка']] || '', r[ix['Сумма, ₸']] === '' ? null : +r[ix['Сумма, ₸']], String(r[ix['Валюта']] || 'KZT').toUpperCase(), r[ix['Оплата']] || '', r[ix['Құжат']] || '', r[ix['Кампания']] || ''].join('|')}));
    const tz2 = (() => { const z = Utilities.formatDate(new Date(), tz, 'Z'); return (z[0] === '-' ? -1 : 1) * (+z.slice(1, 3) * 60 + +z.slice(3, 5)) * 60000; })();
    const parts = d => { const x = new Date(d.getTime() + tz2); return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate()}; }, B = Date.UTC(2026, 7, 10), dayIdx = p => Math.round((Date.UTC(p.y, p.m - 1, p.d) - B) / 864e5);
    const t = ss.getSheetByName(ZOP_TARIFF), tariffs = zopTariffs_(t ? t.getDataRange().getValues() : [], parts, dayIdx).list;
    const chk = zopCostValidate_(o, tariffs, existing); if (!chk.ok) return {ok: false, errors: chk.errors}; if (chk.dup && !o.force) return {ok: false, dup: true, errors: ['дубль']};
    // amount from the tariff is computed now and stored: a later tariff change does not rewrite closed periods
    let sum = chk.row.sum;
    if (sum == null) { const d = dayIdx(parts(new Date(o.date + 'T12:00:00'))), tt = zopTariffAt_(tariffs, o.prov, o.tariff, d); if (!tt) return {ok: false, errors: ['Сома жоқ және бұл күнге тариф табылмады («Тарифтер» парағы)']}; sum = Math.round(chk.row.msgs * tt.price * (tt.cur === 'KZT' ? 1 : (+o.rate || NaN)) * 100) / 100; if (!isFinite(sum)) return {ok: false, errors: ['Тариф валютасы ' + tt.cur + ' — курс керек']}; }
    const id = o.id || Utilities.getUuid().slice(0, 8);
    const rec = {'Дата': new Date(o.date + 'T12:00:00'), 'Подписка': o.sub || '', 'Детали': o.det || '', 'Сумма, ₸': sum, 'Оплата': o.pay || '', 'Рассылка (если знаете)': '', 'Категория': o.cat, 'Поставщик': o.prov || '', 'Валюта': String(o.cur || 'KZT').toUpperCase(),
      'Курс': o.rate || '', 'Кезең басы': o.from ? new Date(o.from + 'T12:00:00') : '', 'Кезең соңы': o.to ? new Date(o.to + 'T12:00:00') : '', 'Кампания': o.camp || '', 'Хабарлама саны': o.msgs || '', 'Тариф': o.tariff || '', 'Құжат': o.doc || '', 'Тексерілді': o.ver ? 'Иә' : '', 'ID': id};
    const row = head.map(h => rec[h] !== undefined ? rec[h] : '');
    let before = '';
    if (o.id) { const i = v.findIndex((r, k) => k && String(r[ix['ID']] || '').trim() === o.id); if (i < 0) return {ok: false, errors: ['ID табылмады']}; before = JSON.stringify(v[i]); sh.getRange(i + 1, 1, 1, row.length).setValues([row]); }
    else sh.appendRow(row);
    const log = ss.getSheetByName(ZOP_COST_LOG) || ss.insertSheet(ZOP_COST_LOG);
    log.appendRow([new Date(), Session.getActiveUser().getEmail() || '—', o.id ? 'өзгертілді' : 'қосылды', id, before, JSON.stringify(row)]);
    try { zopRefresh_(); } catch (e) { /* the next 10-minute tick refreshes */ }
    return {ok: true, id};
  } finally { lock.releaseLock(); }
}

/* =================== Discovery: структура источников без данных ===================
   Запустите вручную (zopDescribeSources) и пришлите журнал. Выводятся только: названия колонок, заполненность, число разных значений.
   Значения показываются лишь для колонок-справочников (статус, тип, оплата, источник, воронка, этап) и только если их ≤ 20.
   Имена, телефоны, суммы, комментарии — не выводятся. */
const ZOP_SAFE_COL = /статус|status|тип|type|оплат|payment|способ|method|қайдан|источник|source|воронк|pipeline|этап|stage|итог|канал|channel|валют|currency|tariff_name|жариялану/i;
const ZOP_PII_COL = /имя|фио|(^|[\s(])аты([\s)-]|$)|name|телефон|phone|номер|mail|почт|коммент|comment|чек|адрес|iin|иин|ссылк|link|контакт|contact|тег|tag/i;
/** Pure: описание таблицы rows (первая строка — заголовки). */
function zopDescribeRows_(rows) {
  const head = (rows[0] || []).map(h => String(h || '').trim()), body = rows.slice(1).filter(r => r && r.some(v => v !== '' && v !== null));
  return {rows: body.length, cols: head.map((h, i) => {
    const vals = body.map(r => r[i]).filter(v => v !== '' && v !== null && v !== undefined), uniq = {};
    vals.forEach(v => { uniq[v instanceof Date ? 'date' : String(v).trim()] = 1; });
    const n = Object.keys(uniq).length, kind = vals.length ? (vals.every(v => v instanceof Date) ? 'date' : vals.every(v => typeof v === 'number') ? 'number' : 'text') : 'empty';
    const show = ZOP_SAFE_COL.test(h) && !ZOP_PII_COL.test(h) && n <= 20 && kind !== 'number';
    return {name: h || '(без названия)', kind, fill: body.length ? Math.round(vals.length / body.length * 100) : 0, distinct: n, values: show ? Object.keys(uniq).slice(0, 20) : null};
  })};
}
function zopDescribeSources() {
  zopRequireOwner_();
  const own = zopHome_(), out = [];
  const one = (label, getRows) => { try { const d = zopDescribeRows_(getRows()); out.push('== ' + label + ': ' + d.rows + ' строк');
      d.cols.forEach(c => out.push('  · ' + c.name + ' — ' + c.kind + ', заполнено ' + c.fill + '%, разных ' + c.distinct + (c.values ? ' → ' + c.values.join(' | ') : ''))); }
    catch (e) { out.push('== ' + label + ': нет доступа — ' + e.message); } };
  const sheetRows = (id, name) => () => { const sh = (id ? SpreadsheetApp.openById(id) : own).getSheetByName(name); if (!sh) throw new Error('нет листа «' + name + '»'); return sh.getDataRange().getValues(); };
  one('amoCRM · Сделки', () => zopCrmRead_('LEADS').rows);
  if (ZOP_USE_OP_SHEET) one('Оқушы тізімі (продажи через менеджера)', sheetRows(ZOP_SALES.ID, ZOP_SALES.LIST));
  one('Платформа · ' + ZOP_PLATFORM.LIST, sheetRows(ZOP_PLATFORM.ID, ZOP_PLATFORM.LIST));
  one('Расходы', sheetRows(null, ZOP.COST)); one('Рассылки', sheetRows(null, ZOP.MAIL));
  Logger.log(out.join('\n'));
  return out.join('\n');
}
/* =================== «Өнімдер»: сначала раздел (A), потом продукт из этого раздела (B) ===================
   Каждому настоящему названию в «Өнім тізімі» назначается направление (столбец G, автоматически — можно поправить руками).
   Скрытый лист «Өнім бөлімдері» держит по столбцу на раздел; выпадающий список в B каждой строки берёт столбец своего раздела. */
const ZOP_PSEC = 'Өнім бөлімдері', ZOP_ALL = 'Барлық өнімдер';
const ZOP_FAMILIES = ['Алтын сақа', 'Алтын сақа · пробный тест', 'Алтын түлек', 'NIS', 'Математика · Мега қайталау', 'Математика · Тоқсандық', 'Ағылшын', 'Басқа өнімдер'];
/** Pure: какие направления продуктов показывать для раздела A. null — раздел не узнан: показываются все продукты. */
function zopSecFamilies_(sec) {
  const s = zopLc_(sec).trim(); if (!s) return null;
  if (/пробн/.test(s)) return ['Алтын сақа · пробный тест'];
  if (/түлек/.test(s)) return ['Алтын түлек'];
  if (/алтын/.test(s)) return ['Алтын сақа'];
  if (/(^|[^a-z])nis([^a-z]|$)/.test(s)) return ['NIS'];
  if (/негізгі/.test(s)) return ['Математика · Мега қайталау', 'Математика · Тоқсандық', 'Ағылшын'];
  if (/мега/.test(s)) return ['Математика · Мега қайталау'];
  if (/тоқсан/.test(s)) return ['Математика · Тоқсандық'];
  if (/математика/.test(s)) return ['Математика · Мега қайталау', 'Математика · Тоқсандық'];
  if (/ағылшын|english/.test(s)) return ['Ағылшын'];
  if (/басқа өнім/.test(s)) return ['Басқа өнімдер'];
  return null;
}
/** Pure: plist = [[название, направление]], sections = значения A. → {secs: [разделы…, «Барлық өнімдер»], cols: [[названия] на каждый раздел]} */
function zopSectionLists_(plist, sections) {
  const secs = [];
  sections.concat(ZOP_FAMILIES).forEach(x => { x = String(x || '').trim(); if (x && x !== ZOP_ALL && secs.indexOf(x) < 0) secs.push(x); });
  const names = plist.map(p => String(p[0] || '').trim()).filter(Boolean);
  const cols = secs.map(sec => { const fam = zopSecFamilies_(sec), hit = fam ? plist.filter(p => fam.indexOf(String(p[1] || '').trim()) >= 0).map(p => String(p[0]).trim()) : names; return hit.length ? hit : names; });
  secs.push(ZOP_ALL); cols.push(names);
  return {secs, cols};
}
/** Пересобирает списки и проверку данных в A и B листа «Өнімдер». Можно запустить вручную: zopProductSections. */
function zopProductSections() { zopRequireOwner_(); return zopApplyProductSections_(); }
function zopApplyProductSections_() {
  const ss = zopHome_(), prod = ss.getSheetByName(ZOP_SALES.PROD), pl = ss.getSheetByName(ZOP_PLIST); if (!prod || !pl) return 0;
  const plist = pl.getLastRow() > 1 ? pl.getRange(2, 1, pl.getLastRow() - 1, 7).getValues().map(r => [String(r[0]).trim(), String(r[6]).trim() || zopSecGuess_(r[0], 'p')]).filter(r => r[0]) : [];
  const n = Math.max(0, prod.getLastRow() - 1), A = n ? prod.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim()) : [];
  const L = zopSectionLists_(plist, A);
  let h = ss.getSheetByName(ZOP_PSEC); if (!h) { h = ss.insertSheet(ZOP_PSEC); h.hideSheet(); }
  const rows = Math.max(1, ...L.cols.map(c => c.length)) + 1;
  h.clearContents(); if (h.getMaxRows() < rows) h.insertRowsAfter(h.getMaxRows(), rows - h.getMaxRows()); if (h.getMaxColumns() < L.secs.length) h.insertColumnsAfter(h.getMaxColumns(), L.secs.length - h.getMaxColumns());
  h.getRange(1, 1, rows, L.secs.length).setValues(Array.from({length: rows}, (_, i) => L.secs.map((sec, j) => i === 0 ? sec : L.cols[j][i - 1] || '')));
  prod.getRange('A2:A').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(L.secs.slice(0, -1), true).setAllowInvalid(true).build());
  prod.getRange('B1').setNote('Алдымен A бағанында бөлімді таңдаңыз — B тізімінде сол бөлімнің өнімдері ғана шығады. Өнімнің бағыты — «Өнім тізімі», G бағаны (қате болса, сонда түзетіңіз).');
  if (!n) return 0;
  const tpl = prod.getRange('B2').getDataValidation(), cache = {};   // B2's rule keeps your «Чип» / multiple choice; only the list changes
  const rule = col => cache[col] || (cache[col] = (tpl ? tpl.copy() : SpreadsheetApp.newDataValidation().setAllowInvalid(true)).requireValueInRange(h.getRange(2, col, h.getMaxRows() - 1, 1), true).build());
  prod.getRange(2, 2, n, 1).setDataValidations(A.map(sec => { const i = L.secs.indexOf(sec); return [rule((sec && i >= 0 ? i : L.secs.length - 1) + 1)]; }));
  return n;
}
/** Простой триггер: изменили раздел в A листа «Өнімдер» → список в B этой строки сразу показывает продукты раздела. */
function onEdit(e) {
  try { if (!e || !e.source) return; const r = e.range; if (!r || r.getSheet().getName() !== ZOP_SALES.PROD || r.getColumn() !== 1 || r.getRow() + r.getNumRows() - 1 < 2) return; zopApplyProductSections_(); }
  catch (err) { console.warn('Өнім бөлімдері: ' + err.message); }
}
function zopRebuildProducts() {
  zopRequireOwner_();
  const ss = zopHome_(), sh = ss.getSheetByName(ZOP_SALES.PROD);
  zopNeedPlatform_(JSON.parse(zopSalesRefresh_()));
  if (sh) { let nm = ZOP_SALES.PROD + ' (ескі)'; for (let i = 2; ss.getSheetByName(nm); i++) nm = ZOP_SALES.PROD + ' (ескі ' + i + ')'; sh.setName(nm); }
  const n = zopCreateProductSheet_();
  zopWriteProductList_(JSON.parse(zopSalesRefresh_()));   // список и выпадающие списки — уже к новому листу
  Logger.log('«Өнімдер» пересобран: ' + n + ' строк' + (sh ? ' · прежний лист: «' + sh.getName() + '»' : ''));
}
/** Лист «Өнімдер» в этой таблице. Если там лист старого вида (без «Нақты атауы») — он переименовывается в «Өнімдер (ескі)» и создаётся новый. */
function zopCreateProductSheet_() {
  const ss = zopHome_(), data = JSON.parse(zopSalesRefresh_());
  if (!ss.getSheetByName(ZOP_SALES.PROD)) zopNeedPlatform_(data);   // новый лист без данных платформы не собираем
  zopWriteProductList_(data);
  let sh = ss.getSheetByName(ZOP_SALES.PROD);
  if (sh) { const h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].join(' ').toLowerCase();
    if (/нақты/.test(h)) {
      if (sh.getLastRow() > 1) { const rng = sh.getRange(2, 2, sh.getLastRow() - 1, 1), v = rng.getValues(); if (v.some(r => /,/.test(String(r[0])))) rng.setValues(v.map(r => [zopNoComma_(r[0])])); }
      if (!sh.getRange('B2').getDataValidation()) sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(ZOP_PLIST).getRange('A2:A'), true).setAllowInvalid(true).build());   // ваш «Чип» / выбор нескольких не сбрасывается
      return 0;
    }
    let nm = ZOP_SALES.PROD + ' (ескі)'; for (let i = 2; ss.getSheetByName(nm); i++) nm = ZOP_SALES.PROD + ' (ескі ' + i + ')'; sh.setName(nm); }
  const rows = zopProductSeed_(data);
  sh = ss.insertSheet(ZOP_SALES.PROD);
  sh.getRange(1, 1, 1, ZOP_PROD_HEAD.length).setValues([ZOP_PROD_HEAD]).setFontWeight('bold').setBackground('#eef1f7').setWrap(true);
  if (rows.length) sh.getRange(2, 1, rows.length, ZOP_PROD_HEAD.length).setValues(rows);
  sh.setFrozenRows(1); sh.getRange('D:D').setNumberFormat('#,##0');
  [150, 440, 220, 90, 100, 100, 100, 200, 330, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const plist = ss.getSheetByName(ZOP_PLIST).getRange('A2:A');
  const prevB = (ss.getSheets().filter(x => x.getName().indexOf(ZOP_SALES.PROD + ' (ескі') === 0).pop() || {getRange: () => ({getDataValidation: () => null})}).getRange('B2').getDataValidation();
  sh.getRange('B2:B').setDataValidation(prevB ? prevB.copy().requireValueInRange(plist, true).build() : SpreadsheetApp.newDataValidation().requireValueInRange(plist, true).setAllowInvalid(true).build()).setBackground('#fff8e1');
  sh.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Иә', 'Жоқ'], true).setAllowInvalid(true).build()).setBackground('#fff8e1');
  sh.getRange('I2:I').setFontColor('#94a3b8');
  sh.getRange(1, 12, 8, 1).setValues([
    ['Как это работает'],
    ['• Сначала A «Бөлім», потом B: в списке B только продукты этого раздела (направление продукта — «Өнім тізімі», столбец G; если ошибка — поправьте там). B «Нақты атауы» — настоящее название из платформы, выбирается из списка (лист «Өнім тізімі», обновляется сам). Чтобы считать НЕСКОЛЬКО продуктов в одной строке: в «Өнім тізімі» у каждого из них в колонке F выберите эту строку (её название из C).'],
    ['• Продажа на платформе попадает в строку с точно таким же «Нақты атауы» (и ценой, если цена указана).'],
    ['• В «Оқушы тізімі» название почти всегда «Алтын Сақа Тренинг» — различает только цена. Для таких строк выберите в B, какой это продукт; слово в H и цена в D находят эти продажи.'],
    ['• C «Есептегі атауы» — как назвать в отчёте; пусто — короткое настоящее название. Строки с одинаковым C в одном разделе — одна строка отчёта.'],
    ['• A «Бөлім» — раздел отчёта. Строка «Барлығы» — общий план раздела.  E: Иә — показывать всегда (даже 0), Жоқ — только когда купили.'],
    ['• Строки читаются сверху вниз, первая подходящая побеждает. Продажа без строки — «Басқа өнімдер» под своим названием.'],
    ['• Таблицу продаж и платформу скрипт только читает.']]).setFontColor('#64748b').setWrap(true);
  sh.setColumnWidth(12, 460);
  zopApplyProductSections_();
  return rows.length;
}

/* =================== карточка отчёта (отдельная страница) =================== */
/** Оплаты из «Оқушы тізімі» по дням: {день: [штук, сумма]} — для сравнения с успешно закрытыми сделками. */
function zopSheetDays_(sd) {
  const out = {}; (sd && sd.cube || []).forEach(r => { if (r[0] !== 's') return; const e = out[r[1]] || (out[r[1]] = [0, 0]); e[0] += r[r.length - 2]; e[1] += r[r.length - 1]; });
  return out;
}
function zopCardPage_(p) {
  const kind = p.card === 'sales' ? 'sales' : 'leads';
  const data = JSON.parse(kind === 'sales' ? (zopCacheGet_('zops') || zopSalesRefresh_()) : (zopCacheGet_() || zopRefresh_()));
  if (kind === 'leads' && ZOP_USE_OP_SHEET) { try { const sd = JSON.parse(zopCacheGet_('zops') || zopSalesRefresh_()); data.sheetDays = zopSheetDays_(sd); if (Array.isArray(sd.sheetMan)) data.sheetMan = {emp: (sd.dict && sd.dict.emp) || [], rows: sd.sheetMan}; } catch (err) { /* без сравнения */ } }
  const html = zopCardHtml_(kind, p.d || '', data, ScriptApp.getService().getUrl());
  return HtmlService.createHtmlOutput(html).setTitle(kind === 'sales' ? 'Сатылым · күнделікті есеп — EduApp' : 'Лидтер · күнделікті есеп — EduApp')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function zopCardHtml_(kind, date, data, url) {
  const cfg = JSON.stringify({kind, date, url, data}).replace(/</g, '\\u003c').replace(/\//g, '\\/').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return '<style>' + ZOP_CARD_CSS + '</style><div id="app"></div><script>(' + zopCardClient_.toString() + ')(' + cfg + ');</script>';
}
const ZOP_CARD_CSS = `
.snapnote.ok{background:#e9f8ef;border-color:#9fdcb5;color:#1d6b3a}
.snapnote{margin:0 0 12px;padding:9px 12px;border-radius:12px;background:#fff7e6;border:1px solid #f5d38a;color:#8a5300;font-size:13px;line-height:1.45}
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root{--bg:#eef2fb;--card:#fff;--ink:#0b1224;--ink2:#4a5672;--ink3:#5B677C;--line:#e6ebf4;--track:#edf1f8;--ind:#3b5bdb;--teal:#0d9488;--amb:#d97706;--vio:#7048e8;--red:#e03131;--grn:#2f9e44;--sl:#94a3b8;--ease:cubic-bezier(.2,.8,.2,1)}
*{box-sizing:border-box}html{-webkit-text-size-adjust:100%}
html,body{margin:0;min-height:100%}
body{font:15px/1.45 Inter,"Segoe UI",Roboto,system-ui,-apple-system,sans-serif;color:var(--ink);-webkit-font-smoothing:antialiased;font-variant-numeric:tabular-nums;
  background:radial-gradient(900px 520px at 105% -8%,#dbe4ff 0%,rgba(219,228,255,0) 60%),radial-gradient(760px 480px at -12% 12%,#dcf7ee 0%,rgba(220,247,238,0) 55%),var(--bg)}
.wrap{max-width:780px;margin:0 auto;padding:14px 14px 96px}
/* ---- hero ---- */
.hero{position:relative;overflow:hidden;isolation:isolate;border-radius:28px;padding:20px 20px 16px;color:#fff;
  background:linear-gradient(155deg,#050816 0%,#0b1338 48%,#1a2878 100%);box-shadow:0 30px 60px -30px rgba(16,24,80,.75),inset 0 1px 0 rgba(255,255,255,.09)}
.orb{position:absolute;border-radius:50%;pointer-events:none;z-index:-1;will-change:transform}
.orb.o1{width:340px;height:340px;right:-90px;top:-150px;background:radial-gradient(circle,rgba(92,124,250,.75) 0%,rgba(92,124,250,0) 68%);animation:float1 16s ease-in-out infinite}
.orb.o2{width:300px;height:300px;left:-120px;bottom:-170px;background:radial-gradient(circle,rgba(32,201,151,.5) 0%,rgba(32,201,151,0) 68%);animation:float2 19s ease-in-out infinite}
.orb.o3{width:220px;height:220px;left:38%;top:-120px;background:radial-gradient(circle,rgba(190,75,219,.42) 0%,rgba(190,75,219,0) 70%);animation:float3 23s ease-in-out infinite}
.gridbg{position:absolute;inset:0;z-index:-1;pointer-events:none;background-image:linear-gradient(rgba(255,255,255,.055) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.055) 1px,transparent 1px);background-size:24px 24px;
  -webkit-mask-image:radial-gradient(ellipse 80% 70% at 75% 0%,#000 0%,transparent 75%);mask-image:radial-gradient(ellipse 80% 70% at 75% 0%,#000 0%,transparent 75%)}
@keyframes float1{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(-50px,40px) scale(1.15)}}
@keyframes float2{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(70px,-30px) scale(1.1)}}
@keyframes float3{0%,100%{transform:translate(0,0)}50%{transform:translate(-60px,50px)}}
.toprow{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap}
.switch{display:inline-flex;padding:3px;border-radius:14px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.16);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px)}
.switch a{color:rgba(255,255,255,.78);text-decoration:none;font-weight:700;font-size:13px;padding:7px 12px;border-radius:11px;transition:background .2s,color .2s;white-space:nowrap}
.switch a:hover{color:#fff;background:rgba(255,255,255,.1)}.switch a.on{color:#0b1224;background:linear-gradient(135deg,#fff,#dbe4ff);box-shadow:0 6px 16px -8px rgba(0,0,0,.6)}
.xlink{display:inline-flex;align-items:center;gap:6px;margin-top:10px;font-weight:800;font-size:13px;color:var(--ind);text-decoration:none;padding:7px 12px;border-radius:11px;background:#eef2ff;transition:background .2s}.xlink:hover{background:#dbe4ff}
.brand{display:flex;align-items:center;gap:9px;font-size:11.5px;letter-spacing:.16em;text-transform:uppercase;font-weight:700;color:rgba(255,255,255,.78)}
.brand svg{width:26px;height:26px;flex:none;filter:drop-shadow(0 4px 10px rgba(76,110,245,.6))}
.hero h1{margin:8px 0 12px;font-size:27px;line-height:1.12;letter-spacing:-.025em;font-weight:800;background:linear-gradient(92deg,#fff 0%,#dbe4ff 55%,#96f2d7 100%);-webkit-background-clip:text;background-clip:text;color:transparent}
.daterow{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.daterow button,.daterow label{height:38px;border-radius:13px;border:1px solid rgba(255,255,255,.18);background:rgba(255,255,255,.08);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);color:#fff;font:inherit;font-weight:600;display:inline-flex;align-items:center;gap:8px;padding:0 13px;cursor:pointer;transition:background .2s,transform .2s}
.daterow button:hover:not(:disabled),.daterow label:hover{background:rgba(255,255,255,.16)}.daterow button:active,.daterow label:active{transform:scale(.97)}
.daterow button{width:38px;justify-content:center;padding:0;font-size:19px}.daterow button:disabled{opacity:.3;cursor:default}
.tiles{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-top:16px}
.tile{position:relative;overflow:hidden;background:linear-gradient(160deg,rgba(255,255,255,.12),rgba(255,255,255,.04));border:1px solid rgba(255,255,255,.14);border-radius:18px;padding:11px 12px 10px;min-width:0;
  animation:rise .7s var(--ease) both;animation-delay:var(--d,0ms)}
.tile:after{content:"";position:absolute;inset:0;background:linear-gradient(110deg,transparent 30%,rgba(255,255,255,.18) 50%,transparent 70%);transform:translateX(-120%);animation:sheen 2.6s var(--ease) 1s both}
.tile small{display:block;font-size:11.5px;color:rgba(255,255,255,.72);line-height:1.25;min-height:29px}.tile b{display:block;font-size:30px;line-height:1.05;margin-top:4px;letter-spacing:-.03em;font-weight:800}
.tile b small{display:inline;font-size:16px;color:rgba(255,255,255,.75);font-weight:700;min-height:0}.tile i{font-style:normal;font-size:12px;color:rgba(255,255,255,.7)}.tile b.m{font-size:24px;white-space:nowrap}
@keyframes rise{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}
@keyframes sheen{to{transform:translateX(120%)}}
.spark{margin-top:12px;border-radius:18px;padding:10px 12px 6px;background:rgba(5,9,30,.35);border:1px solid rgba(255,255,255,.09);animation:rise .8s var(--ease) .25s both}
.spark-h{display:flex;justify-content:space-between;align-items:baseline;font-size:12px;color:rgba(255,255,255,.66)}.spark-h b{color:#fff;font-size:13px}
.spark svg{display:block;max-width:100%;height:auto;overflow:visible;margin-top:4px}
.sp-line{fill:none;stroke:url(#spStroke);stroke-width:2.4;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:1;stroke-dashoffset:1;animation:draw 1.6s var(--ease) .35s forwards;filter:drop-shadow(0 3px 8px rgba(116,143,252,.55))}
.sp-area{fill:url(#spFill);opacity:0;animation:fade 1s ease 1s forwards}
.sp-dot{fill:#0b1338;stroke:#a5b4fc;stroke-width:1.6;pointer-events:none}.sp-dot.on{fill:#96f2d7;stroke:#fff;stroke-width:2}
.sp-grid{stroke:rgba(255,255,255,.09);stroke-dasharray:3 4}.sp-y{font-size:9.5px;fill:rgba(255,255,255,.45)}.sp-x{font-size:9.5px;fill:rgba(255,255,255,.5);font-weight:600}.sp-x.we{fill:rgba(255,168,168,.65)}.sp-x.on{fill:#96f2d7;font-weight:800}
.sp-sel{stroke:rgba(150,242,215,.35);stroke-dasharray:2 3}.sp-tip{font-size:11px;font-weight:800;fill:#fff;opacity:0;transition:opacity .15s;pointer-events:none}.sp-pt:hover .sp-tip,.sp-pt:focus-within .sp-tip,.sp-pt.on .sp-tip{opacity:1}.sp-pt:hover .sp-dot{fill:#fff}
.sp-pulse{fill:none;stroke:#96f2d7;stroke-width:2;transform-box:fill-box;transform-origin:center;animation:pulse 2s ease-out infinite}
.sp-hit{fill:transparent;cursor:pointer;outline:none}
@keyframes draw{to{stroke-dashoffset:0}}@keyframes fade{to{opacity:1}}
@keyframes pulse{0%{transform:scale(1);opacity:.9}100%{transform:scale(3.2);opacity:0}}
/* ---- sections ---- */
.sec{position:relative;background:var(--card);border-radius:24px;padding:18px 16px 16px;margin-top:14px;border:1px solid rgba(230,235,244,.9);
  box-shadow:0 1px 2px rgba(15,23,42,.04),0 22px 44px -28px rgba(30,41,110,.38);opacity:0;transform:translateY(22px) scale(.985);transition:opacity .75s var(--ease),transform .75s var(--ease)}
.sec.in{opacity:1;transform:none}
.sec h2{margin:0 0 14px;font-size:13px;letter-spacing:.07em;text-transform:uppercase;color:var(--ink2);display:flex;align-items:center;gap:10px;font-weight:800}
.sec h2 span.n{display:inline-grid;place-items:center;min-width:30px;height:30px;padding:0 6px;border-radius:10px;background:linear-gradient(135deg,#4c6ef5,#7048e8);color:#fff;font-size:13px;letter-spacing:0;box-shadow:0 8px 18px -8px rgba(76,110,245,.9)}
.big{display:flex;align-items:baseline;gap:8px;flex-wrap:wrap;margin:-4px 0 12px}.big b{font-size:34px;letter-spacing:-.03em;font-weight:800}.big>span{color:var(--ink2);font-size:13.5px}
.rows{display:grid;gap:10px}
.row{display:grid;grid-template-columns:minmax(0,42%) 1fr auto;align-items:center;gap:10px}
.row .lab{font-size:14px;font-weight:600;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.row .lab em{font-style:normal;color:var(--ink3);font-weight:500;font-size:12.5px;margin-left:5px}
.row.sub .lab{font-weight:500;color:var(--ink2);padding-left:14px;font-size:13px;white-space:normal;line-height:1.25}
.track{height:12px;border-radius:99px;background:var(--track);overflow:hidden;box-shadow:inset 0 1px 2px rgba(15,23,42,.06)}.row.sub .track{height:7px}
.fill{height:100%;width:var(--w);min-width:3px;border-radius:99px;background:linear-gradient(90deg,color-mix(in srgb,var(--c) 55%,#fff),var(--c));box-shadow:0 2px 8px -2px color-mix(in srgb,var(--c) 70%,transparent);transition:width 1.15s var(--ease) var(--d,0ms)}
.sec:not(.in) .fill{width:0}
.val{font-weight:800;font-size:14.5px;text-align:right;min-width:48px}.row.sub .val{font-weight:600;color:var(--ink2);font-size:13px}
.zero .fill{opacity:0}.zero .lab,.zero .val{color:var(--ink3)}
/* gauges */
.gauges{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px}
.gauge{position:relative;border:1px solid var(--line);border-radius:22px;padding:14px 10px 12px;text-align:center;background:linear-gradient(180deg,#fff,#f8faff)}
.gwrap{position:relative;width:132px;height:132px;margin:0 auto}
.gwrap svg{width:132px;height:132px;display:block}
.gtrack{fill:none;stroke:var(--track);stroke-width:12}
.arc{fill:none;stroke-width:12;stroke-linecap:round;stroke-dashoffset:var(--off);transition:stroke-dashoffset 1.5s var(--ease) var(--d,0ms);filter:drop-shadow(0 4px 8px var(--glow))}
.sec:not(.in) .arc{stroke-dashoffset:var(--C)}
.gc{position:absolute;inset:0;display:grid;place-content:center;line-height:1}.gc b{font-size:30px;font-weight:800;letter-spacing:-.03em}.gc b small{font-size:15px;color:var(--ink3);font-weight:700}
.gc i{font-style:normal;font-size:11.5px;color:var(--ink3);margin-top:5px;display:block}
.gt{font-weight:800;font-size:14.5px;margin-top:8px}.gm{font-size:12.5px;color:var(--ink2);margin-top:2px}
.minis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-top:12px}
.mini{border:1px solid var(--line);border-radius:16px;padding:10px 12px;display:flex;align-items:center;gap:10px;background:#fbfcff}
.mini i{width:34px;height:34px;border-radius:11px;display:grid;place-items:center;font-style:normal;font-size:17px;background:linear-gradient(135deg,#eef2ff,#e6fcf5);flex:none}
.mini b{display:block;font-size:18px;font-weight:800;letter-spacing:-.02em}.mini small{display:block;font-size:12px;color:var(--ink2);line-height:1.25}
/* donut */
.donut{display:grid;grid-template-columns:auto 1fr;gap:18px;align-items:center}
.dwrap{position:relative;width:150px;height:150px}.dwrap svg{width:150px;height:150px;display:block;transform:rotate(-90deg)}
.dseg{fill:none;stroke-width:20;stroke-dashoffset:0;transition:stroke-dashoffset 1.1s var(--ease) var(--d,0ms)}
.sec:not(.in) .dseg{stroke-dashoffset:var(--C)}
.dc{position:absolute;inset:0;display:grid;place-content:center;text-align:center;line-height:1.05}.dc b{font-size:28px;font-weight:800;letter-spacing:-.03em}.dc small{font-size:11.5px;color:var(--ink3)}
.legend{display:grid;gap:8px}.lg{display:flex;align-items:center;gap:9px;font-size:14px}.lg i{width:11px;height:11px;border-radius:4px;flex:none}.lg b{margin-left:auto;font-weight:800}.lg span.p{color:var(--ink3);font-weight:600;font-size:12.5px;margin-left:6px}
.lg.zero{color:var(--ink3)}.lg.zero i{opacity:.35}
.split{display:grid;grid-template-columns:auto 1fr;gap:18px;align-items:center}
.stack{height:16px;border-radius:99px;overflow:hidden;display:flex;gap:2px;background:var(--track);margin:6px 0 10px}.stack>div{height:100%;width:var(--w);transition:width 1.1s var(--ease) var(--d,0ms)}.sec:not(.in) .stack>div{width:0}
.duo{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px}
.kpi{border:1px solid var(--line);border-radius:16px;padding:11px 13px;background:linear-gradient(180deg,#fff,#f8faff)}.kpi small{color:var(--ink2);font-size:12.5px;display:flex;align-items:center;gap:6px}.kpi b{display:block;font-size:26px;letter-spacing:-.03em;margin-top:2px;font-weight:800}
.dot{width:8px;height:8px;border-radius:50%;display:inline-block}
.subh{font-size:12.5px;color:var(--ink2);font-weight:800;margin:16px 0 9px;letter-spacing:.02em}
.prog{display:grid;gap:12px}.pg{border:1px solid var(--line);border-radius:16px;padding:12px}
.pg-top{display:flex;align-items:baseline;gap:8px;justify-content:space-between}.pg-top b{font-size:15px}.pg-top .pct{font-size:22px;font-weight:800;letter-spacing:-.02em}
.pg-meta{color:var(--ink2);font-size:12.5px;margin:2px 0 8px}
.pbar{position:relative;height:12px;border-radius:99px;background:var(--track);overflow:hidden}.pbar>div{height:100%;border-radius:99px;width:var(--w);transition:width 1.2s var(--ease)}.sec:not(.in) .pbar>div{width:0}
.tbl{width:100%;border-collapse:collapse;font-size:13.5px}.tbl th{font-size:11px;color:var(--ink3);text-transform:uppercase;letter-spacing:.06em;text-align:right;font-weight:800;padding:0 0 7px}.tbl th:first-child{text-align:left}
.tbl th+th,.tbl td+td{padding-left:10px}.tbl td{padding:9px 0;border-top:1px solid var(--line);text-align:right;white-space:nowrap}.tbl td:first-child{text-align:left;white-space:normal;font-weight:700}.tbl td small{display:block;color:var(--ink3);font-weight:500;font-size:12px}
.tbl td small.rawp{color:var(--ink2);font-weight:600;white-space:normal}.tbl tr{transition:background .2s}.tbl tbody tr:hover{background:#f6f8fe}
.chip{display:inline-flex;align-items:center;border-radius:999px;padding:3px 9px;font-size:12px;font-weight:800}
.note{color:var(--ink3);font-size:12.5px;margin-top:10px;line-height:1.45}
.grp{position:relative;border:1px solid color-mix(in srgb,var(--gc) 22%,var(--line));border-radius:20px;padding:14px 14px 14px 18px;margin-top:12px;overflow:hidden;
  background:linear-gradient(135deg,color-mix(in srgb,var(--gc) 7%,#fff) 0%,#fff 55%)}
.grp:first-child{margin-top:0}.grp:before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:linear-gradient(180deg,var(--gc),color-mix(in srgb,var(--gc) 45%,#fff))}
.grp-h{display:flex;align-items:center;gap:11px;margin-bottom:10px}.grp-h i{width:38px;height:38px;border-radius:12px;display:grid;place-items:center;font-style:normal;font-size:19px;flex:none;
  background:color-mix(in srgb,var(--gc) 14%,#fff);box-shadow:inset 0 0 0 1px color-mix(in srgb,var(--gc) 25%,transparent)}
.grp-t b{display:block;font-size:16px;font-weight:800;letter-spacing:-.01em;color:color-mix(in srgb,var(--gc) 70%,var(--ink))}.grp-t small{display:block;font-size:12.5px;color:var(--ink2);font-weight:600}
.grp-p{margin-left:auto;font-size:24px;font-weight:800;letter-spacing:-.02em}.grp-bar{height:8px;margin:-2px 0 12px}.grp .tbl tr:first-child th{border-top:0}
.tm{position:relative;width:100%;margin-top:4px}
.tmc{position:absolute}.tmc>div{position:absolute;inset:3px;border-radius:14px;padding:9px 11px;overflow:hidden;display:flex;flex-direction:column;justify-content:flex-end;color:#fff;
  background:linear-gradient(140deg,color-mix(in srgb,var(--c) 70%,#fff) 0%,var(--c) 70%);box-shadow:inset 0 1px 0 rgba(255,255,255,.25),0 8px 18px -12px color-mix(in srgb,var(--c) 80%,transparent);
  transition:transform .7s var(--ease) var(--d,0ms),opacity .7s var(--ease) var(--d,0ms);cursor:default}
.tmc>div:hover{filter:brightness(1.06)}.sec:not(.in) .tmc>div{transform:scale(.7);opacity:0}
.tmc b{font-size:26px;font-weight:800;letter-spacing:-.03em;line-height:1}.tmc>div>span{font-size:12.5px;font-weight:700;line-height:1.2;margin-top:4px;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.tmc>div>small{font-size:11px;opacity:.85;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tmc.sm b{font-size:18px}.tmc.sm>div>span{font-size:11px;-webkit-line-clamp:2}.tmc.sm>div>small{display:none}
.tmc.xs>div{padding:4px;justify-content:center;align-items:center}.tmc.xs b{font-size:13px}.tmc.xs>div>span,.tmc.xs>div>small{display:none}
.pcols .cbar{width:64%;max-width:64px;margin:0 auto;background:var(--track);border-radius:12px 12px 4px 4px;border-bottom:0;overflow:hidden}.pcols .cbar i{width:100%;max-width:none;border-radius:0}
.colm b .of{font-size:12px;color:var(--ink3);font-weight:700}.colm em{display:block;font-style:normal;font-weight:800;font-size:13px;margin-top:2px}
.funnel{display:grid;gap:8px;margin:4px 0 12px}.fstep{display:grid;grid-template-columns:150px 1fr 64px;align-items:center;gap:12px}.fl{font-weight:700;font-size:13.5px;color:var(--ink2);line-height:1.25}.fwrap{min-width:0}
.fbar{width:var(--w);margin:0 auto;min-height:58px;display:flex;align-items:center;justify-content:center;gap:10px;color:#fff;border-radius:10px;padding:8px 14px;
  background:linear-gradient(135deg,color-mix(in srgb,var(--c) 70%,#fff),var(--c));clip-path:polygon(0 0,100% 0,calc(100% - 18px) 100%,18px 100%);transition:width 1.1s var(--ease) var(--d,0ms)}
.sec:not(.in) .fbar{width:20%}.fbar b{font-size:24px;font-weight:800;letter-spacing:-.02em}.fbar span{font-size:13px;font-weight:700;opacity:.95}
.fstep em{font-style:normal;font-weight:800;color:var(--ink2);min-width:58px;text-align:right}
.btn.pdf{background:linear-gradient(135deg,#fff,#fff4e6);border-color:#ffd8a8;color:#d9480f}
.pdfmode .acts,.pdfmode .toast,.pdfmode .xlink,.pdfmode .spark-n{display:none!important}.pdfmode *,.pdfmode *:before,.pdfmode *:after{animation:none!important;transition:none!important}
.pdfmode .sp-line{stroke-dashoffset:0}.pdfmode .sp-area{opacity:1}.pdfmode .tile:after,.pdfmode .btn.pri:after{display:none}
.colm small{overflow-wrap:anywhere}
.mdl{position:fixed;inset:0;z-index:80;background:rgba(8,12,30,.45);-webkit-backdrop-filter:blur(4px);backdrop-filter:blur(4px);display:flex;align-items:center;justify-content:center;padding:16px;opacity:0;transition:opacity .25s}
.mdl.on{opacity:1}.mdl-c{width:min(460px,100%);max-height:calc(100vh - 32px);overflow:auto;background:#fff;border-radius:22px;padding:18px;box-shadow:0 30px 70px -20px rgba(8,12,40,.55);transform:translateY(16px);transition:transform .3s var(--ease)}
.mdl.on .mdl-c{transform:none}.mdl-h{display:flex;justify-content:space-between;align-items:baseline;gap:10px;margin-bottom:6px}.mdl-h b{font-size:18px}.mdl-h span{color:var(--ink3);font-size:13px}
.mdl-s{font-size:12px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:var(--ink2);margin:14px 0 8px;display:flex;justify-content:space-between}.mdl-s a{color:var(--ind);text-transform:none;letter-spacing:0;font-weight:700}
.opt{display:flex;gap:10px;align-items:flex-start;border:1px solid var(--line);border-radius:14px;padding:10px 12px;margin-bottom:8px;cursor:pointer;transition:border-color .2s,background .2s}
.opt:has(input:checked){border-color:#91a7ff;background:#f3f6ff}.opt input{margin-top:3px;accent-color:#4c6ef5;width:17px;height:17px;flex:none}.opt b{display:block;font-size:14.5px}.opt small{display:block;color:var(--ink3);font-size:12px}
.opt.sm{padding:8px 12px;font-size:14px;font-weight:600}.chkl{display:grid;gap:0}.mdl-n{color:var(--ink3);font-size:12.5px;margin:4px 0 0}
.mdl-f{display:flex;gap:8px;margin-top:14px}.mdl-f .btn{height:46px}
@media (max-width:480px){.mdl{align-items:flex-end;padding:0}.mdl-c{border-radius:22px 22px 0 0;width:100%}.colm small{font-size:10.5px}}
.pcards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px}
.pcard{position:relative;border-radius:16px;padding:12px 14px 12px 18px;background:linear-gradient(135deg,color-mix(in srgb,var(--c) 12%,#fff),#fff);border:1px solid color-mix(in srgb,var(--c) 25%,var(--line));overflow:hidden}
.pcard:before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:var(--c)}.pcard b{display:block;font-size:28px;font-weight:800;letter-spacing:-.03em;color:color-mix(in srgb,var(--c) 75%,var(--ink))}
.pcard>span{display:block;font-weight:700;font-size:13.5px;margin-top:2px}.pcard>small{display:block;color:var(--ink3);font-size:12px;margin-top:2px}
.tmleg{display:flex;flex-wrap:wrap;gap:6px 14px;margin-top:10px;font-size:12.5px;color:var(--ink2)}.tmleg span{display:inline-flex;align-items:center;gap:6px}.tmleg i{width:10px;height:10px;border-radius:3px;flex:none}
.zero-chips{margin-top:10px;font-size:12.5px;color:var(--ink3);display:flex;flex-wrap:wrap;gap:6px;align-items:center}.zero-chips span{background:var(--track);color:var(--ink2);border-radius:99px;padding:3px 10px;font-weight:600}
.mini-g{grid-template-columns:repeat(auto-fit,minmax(112px,1fr));gap:8px}.mini-g .gauge{padding:10px 4px 10px}.mini-g .gwrap,.mini-g .gwrap svg{width:92px;height:92px}.mini-g .gc b{font-size:21px}.mini-g .gt{font-size:12.5px;line-height:1.25}.mini-g .gm{font-size:11.5px}
.cols{display:grid;grid-auto-flow:column;grid-auto-columns:minmax(0,1fr);gap:10px;align-items:start;padding-top:4px}
.colm{text-align:center;min-width:0}.colm b{display:block;font-size:17px;font-weight:800;margin-bottom:6px}.cbar{height:130px;display:flex;align-items:flex-end;justify-content:center;border-bottom:2px solid var(--line)}
.cbar i{display:block;width:72%;max-width:60px;height:var(--h);border-radius:12px 12px 4px 4px;background:linear-gradient(180deg,color-mix(in srgb,var(--c) 60%,#fff),var(--c));box-shadow:0 8px 16px -10px var(--c);transition:height 1.1s var(--ease) var(--d,0ms)}
.sec:not(.in) .cbar i{height:0}.colm small{display:block;font-size:11.5px;color:var(--ink2);margin-top:7px;line-height:1.2;font-weight:600;min-height:2.4em}
.wlegend{display:flex;flex-wrap:wrap;gap:8px 14px;margin:6px 0 12px;font-size:12.5px;color:var(--ink2)}.wlegend span{display:inline-flex;align-items:center;gap:6px}.wlegend i{width:10px;height:10px;border-radius:3px;display:inline-block}
.wrows{display:grid;gap:10px}.wrow{display:grid;grid-template-columns:minmax(110px,170px) 1fr 44px;align-items:center;gap:10px}
.wname b{display:block;font-size:14px}.wname small{display:block;font-size:11.5px;color:var(--ink3)}
#ctip{position:fixed;left:0;top:0;z-index:90;pointer-events:none;min-width:190px;max-width:min(300px,calc(100vw - 16px));background:#0f172a;color:#fff;border-radius:14px;padding:10px 12px 11px;font-size:13px;line-height:1.35;box-shadow:0 18px 40px -12px rgba(15,23,42,.55);opacity:0;transition:opacity .12s,transform .08s}
#ctip.on{opacity:1}#ctip .h{display:flex;align-items:center;gap:7px;font-weight:700}#ctip .b{font-size:20px;font-weight:800;letter-spacing:-.02em;margin-top:2px}
#ctip .r{display:grid;gap:3px;margin-top:7px}#ctip .r div{display:flex;align-items:center;gap:8px;color:#cbd5e1}#ctip .r span{flex:1}#ctip .r b{color:#fff}
#ctip i{width:9px;height:9px;border-radius:3px;flex:none}#ctip .n{margin-top:7px;padding-top:6px;border-top:1px solid rgba(255,255,255,.14);color:#94a3b8;font-size:12px}
[data-ct]{cursor:help}[data-ct]:focus{outline:none}.wbar i[data-ct]:hover,.wbar i[data-ct]:focus{filter:brightness(1.15);box-shadow:inset 0 0 0 2px rgba(255,255,255,.7)}.mini[data-ct]:hover,.wlegend span[data-ct]:hover{border-color:#94a3b8}
.wtrack{height:16px;background:var(--track);border-radius:8px;overflow:hidden}.wbar{display:flex;height:100%;min-width:4px}.wbar i{display:block;height:100%}.wn{text-align:right;font-size:15px}
@media (max-width:520px){.wrow{grid-template-columns:1fr 40px}.wtrack{grid-column:1/-1;grid-row:2}}
.tasks{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:10px}
.task{border:1px solid #ffd6d6;background:#fffafa;border-radius:16px;padding:12px 14px}.task-h{display:flex;align-items:center;gap:8px}.task-h i{font-style:normal}.task-h b{font-size:15px}
.task-h span{margin-left:auto;background:#ffe3e3;color:#c92a2a;border-radius:99px;padding:2px 9px;font-weight:800;font-size:12.5px}
.task ul{margin:8px 0 0;padding-left:18px;display:grid;gap:4px;font-size:13.5px;line-height:1.4}.howto{margin-top:12px;background:#f1f3f9;border-radius:14px;padding:10px 12px;font-size:13px;color:var(--ink2);line-height:1.45}
.fintro{margin:0 0 12px;font-size:13.5px;line-height:1.45;color:var(--ink2)}
.fgrid{display:grid;gap:10px}.fcard{border:1px solid var(--line);border-radius:16px;padding:12px 14px;background:#fff}.fcard.done{background:#f6fcf7;border-color:#c3ecd0}
.fcard-h{display:flex;align-items:baseline;gap:8px;flex-wrap:wrap;margin-bottom:10px}.fcard-h b{font-size:15px}.fcard-h small{color:var(--ink3);font-size:12.5px}.fok{margin-left:auto;color:#2b8a3e;font-weight:700;font-size:12.5px}
.fcells{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}@media (max-width:560px){.fcells{grid-template-columns:repeat(2,minmax(0,1fr))}}
.fcell{border-radius:12px;padding:10px 12px;background:#f4f6fb;display:flex;flex-direction:column;min-height:84px;min-width:0}.fcell small{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.fcell b{white-space:nowrap;font-variant-numeric:tabular-nums}.fcell em{margin-top:auto}.fcell small{display:block;font-size:12px;color:var(--ink2);line-height:1.3}.fcell b{display:block;font-size:20px;font-weight:800;margin-top:3px;letter-spacing:-.02em}.fcell em{display:block;font-style:normal;font-size:11.5px;color:var(--ink3)}
.fcell.bad{background:#fff0f0}.fcell.bad b{color:#c92a2a}.fcell.ok{background:#ebfbee}.fcell.ok b{color:#2b8a3e}.fcell.na b{color:var(--ink3)}
.fpay{margin-top:14px;display:grid;gap:10px}.fpay-g{display:grid;grid-template-columns:1fr 1fr;gap:10px}.fsrc{border:1px solid var(--line);border-radius:16px;padding:12px 14px;background:#fff;min-width:0}.fsrc small{display:flex;align-items:flex-start;gap:7px;font-size:13px;line-height:1.3;color:var(--ink2);min-height:2.6em}.fsrc small i{margin-top:4px}.fsrc small i{width:8px;height:8px;border-radius:50%;flex:none}.fsrc b{display:block;font-size:30px;font-weight:800;letter-spacing:-.03em;line-height:1.1;margin-top:6px;font-variant-numeric:tabular-nums}.fsrc span{display:block;font-size:13px;font-weight:650;color:var(--ink2);margin-top:2px}
.fpay-w{display:flex;align-items:flex-start;gap:10px;border-radius:14px;padding:12px 14px;background:#fff0f0;color:#c92a2a}.fpay-w span{flex:none;line-height:1.45}.fpay-w p{margin:0;font-size:14px;line-height:1.45}.fpay-w b{font-weight:800}
.fillhead{display:flex;align-items:center;gap:14px;border-radius:16px;padding:12px 16px;margin-bottom:12px}.fillhead b{font-size:34px;font-weight:800;letter-spacing:-.03em}.fillhead span{font-size:14px;line-height:1.4;font-weight:600}
.fillhead.bad{background:#fff5f5;color:#c92a2a}.fillhead.ok{background:#ebfbee;color:#2b8a3e}
.filltop{display:grid;grid-template-columns:200px 1fr;gap:14px;align-items:stretch}.fillk{display:grid;gap:10px;align-content:start}.kpi i.sum{font-style:normal;display:block;color:var(--ink2);font-size:13px;font-weight:700;margin-top:2px}
.callout{border-radius:14px;padding:10px 12px;font-size:13.5px;background:#eef2ff;color:#364fc7;line-height:1.4}.callout.bad{background:#fff0f0;color:#c92a2a}.callout.ok{background:#ebfbee;color:#2b8a3e}
.chks{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px;margin-top:12px}
.chk{display:flex;align-items:center;gap:10px;border:1px solid var(--line);border-radius:14px;padding:10px 12px;background:#fbfffc}.chk i{width:28px;height:28px;border-radius:9px;display:grid;place-items:center;font-style:normal;font-weight:800;background:#d3f9d8;color:#2b8a3e;flex:none}
.chk b{display:block;font-size:14px}.chk small{display:block;font-size:11.5px;color:var(--ink3)}.chk span{margin-left:auto;font-weight:800;color:#2b8a3e;font-size:13px}
.chk.bad{background:#fff8f8;border-color:#ffd8d8}.chk.bad i{background:#ffe3e3;color:#e03131}.chk.bad span{color:#e03131;font-size:18px}
.chip.ok{background:#ebfbee;color:#2b8a3e}.chip.bad{background:#ffe3e3;color:#e03131}.tblw{overflow-x:auto;-webkit-overflow-scrolling:touch}.tblw .tbl td,.tblw .tbl th{text-align:center}.tblw .tbl td:first-child,.tblw .tbl th:first-child{text-align:left}
.empty{border:1px dashed #cfd8e8;background:#f7f9fd;border-radius:14px;padding:12px 14px;color:var(--ink2);font-size:14px}
/* ---- actions ---- */
.acts{position:sticky;bottom:12px;z-index:5;display:flex;gap:8px;flex-wrap:wrap;margin-top:16px;padding:8px;border-radius:20px;background:rgba(255,255,255,.72);-webkit-backdrop-filter:blur(14px);backdrop-filter:blur(14px);border:1px solid rgba(230,235,244,.9);box-shadow:0 16px 36px -20px rgba(30,41,110,.45)}
.acts{transition:transform .35s var(--ease),opacity .3s}.acts.hid{transform:translateY(calc(100% + 24px));opacity:0;pointer-events:none}
.btn{flex:1 1 auto;height:46px;border-radius:14px;border:1px solid var(--line);background:#fff;color:var(--ink);font:inherit;font-weight:800;padding:0 16px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;transition:transform .15s,box-shadow .2s}
.btn:active{transform:scale(.97)}
.btn.pri{position:relative;overflow:hidden;background:linear-gradient(135deg,#4c6ef5,#7048e8);color:#fff;border:0;box-shadow:0 10px 24px -10px rgba(76,110,245,.9)}
.btn.pri:after{content:"";position:absolute;inset:0;background:linear-gradient(110deg,transparent 35%,rgba(255,255,255,.35) 50%,transparent 65%);transform:translateX(-120%);animation:sheen 3.2s ease-in-out 1.5s infinite}
.foot{color:var(--ink3);font-size:12px;margin-top:16px;text-align:center;line-height:1.55}
.toast{position:fixed;left:50%;bottom:86px;transform:translate(-50%,10px);background:#0b1224;color:#fff;padding:11px 16px;border-radius:13px;font-weight:700;font-size:14px;opacity:0;transition:opacity .25s,transform .25s;pointer-events:none;z-index:60;box-shadow:0 14px 30px -12px rgba(0,0,0,.5)}.toast.on{opacity:1;transform:translate(-50%,0)}
textarea.copy{width:100%;min-height:200px;margin-top:10px;border-radius:12px;border:1px solid var(--line);padding:10px;font:13px/1.4 ui-monospace,Consolas,monospace}
.cal{position:fixed;z-index:50;width:min(316px,calc(100vw - 20px));background:#fff;color:var(--ink);border-radius:20px;padding:14px 14px 12px;box-shadow:0 24px 60px -18px rgba(15,23,42,.45),0 2px 6px rgba(15,23,42,.08);border:1px solid var(--line)}
.cal-h{display:flex;align-items:center;gap:6px;margin-bottom:10px}.cal-h b{flex:1;text-align:center;font-size:15px}
.cal-h button{width:34px;height:34px;border-radius:11px;border:1px solid var(--line);background:#fff;color:var(--ink);font-size:17px;cursor:pointer}.cal-h button:disabled{opacity:.3;cursor:default}
.cal-w,.cal-g{display:grid;grid-template-columns:repeat(7,1fr);gap:3px}.cal-w span{font-size:11px;font-weight:700;color:var(--ink3);text-align:center;padding:2px 0 4px;text-transform:uppercase;letter-spacing:.04em}
.cal-g button{height:42px;border-radius:11px;border:0;background:transparent;color:var(--ink);font:inherit;font-weight:600;font-size:14px;cursor:pointer;display:flex;flex-direction:column;align-items:center;justify-content:center;line-height:1.05;padding:0}
.cal-g button small{font-size:9.5px;font-weight:600;color:var(--ink3);margin-top:2px}
.cal-g button:hover:not(:disabled){background:#eef2fb}.cal-g button:disabled{color:#c7cfdc;cursor:default}.cal-g button:disabled small{color:transparent}
.cal-g button.we{color:#b4235a}.cal-g button.we:disabled{color:#e6c3d1}
.cal-g button.now{box-shadow:inset 0 0 0 1.5px #c5cfe4}
.cal-g button.on{background:linear-gradient(135deg,#3b5bdb,#5f3dc4);color:#fff;box-shadow:0 6px 16px -8px rgba(59,91,219,.9)}.cal-g button.on small{color:rgba(255,255,255,.8)}
.cal-f{display:flex;gap:8px;margin-top:10px;padding-top:10px;border-top:1px solid var(--line)}.cal-f button{flex:1;height:36px;border-radius:11px;border:1px solid var(--line);background:#f7f9fd;color:var(--ink);font:inherit;font-weight:700;font-size:13px;cursor:pointer}
.cal-f button:disabled{opacity:.4;cursor:default}
@media (max-width:480px){.wrap{padding:10px 10px 92px}.hero{padding:16px 14px 14px;border-radius:24px}.hero h1{font-size:22px}.tiles{gap:7px}.tile{padding:9px 9px 8px}.tile b{font-size:23px}.tile b.m{font-size:17px}.tile small{font-size:10.5px;min-height:26px}
  .row{grid-template-columns:minmax(0,46%) 1fr auto;gap:8px}.tbl{font-size:12.5px}.tbl th+th,.tbl td+td{padding-left:6px}.chip{padding:2px 6px}.split,.donut{grid-template-columns:1fr;justify-items:center}.legend{width:100%}.duo{gap:8px}.kpi b{font-size:22px}.sec{padding:15px 12px 13px;border-radius:20px}
  .filltop{grid-template-columns:1fr}.gauges{grid-template-columns:1fr 1fr;gap:8px}.fstep{grid-template-columns:96px 1fr 52px;gap:8px}.fl{font-size:12px}.tmc b{font-size:20px}.cbar{height:110px}.colm b{font-size:15px}.gwrap,.gwrap svg{width:112px;height:112px}.gc b{font-size:25px}}
@media (prefers-reduced-motion:reduce){*,*:before,*:after{animation:none!important;transition:none!important}.sec{opacity:1;transform:none}.sp-line{stroke-dashoffset:0}.sp-area{opacity:1}}
`;

/* Клиентская часть карточки: считает отчёт за выбранный день из агрегатов и рисует его (HTML + SVG, без внешних библиотек). */
function zopCardClient_(cfg) {
  const D = cfg.data, kind = cfg.kind, $ = id => document.getElementById(id);
  if (kind === 'sales' && D && D.crm && Array.isArray(D.crm.main) && Array.isArray(D.cube)) (function () {   /* иесі 15.10.2026: ОП кестесі алынды — менеджер сатылымы = CRM «Успешно» бюджеті (жабылған күні); платформа — бұрынғыдай */
    const C = D.crm, CD = C.dict || {}, OK = (CD.outcome || []).indexOf('Успешно'), SUP = (CD.funnel || []).indexOf('Техподдержка'), B0 = Date.UTC(+C.base.slice(0, 4), +C.base.slice(5, 7) - 1, +C.base.slice(8, 10));
    const RAW = {'WhatsApp': 'Органика WhatsApp', 'Таргет': 'Таргетированная реклама в Инстаграм или Фейсбук', 'Instagram': 'Органика Instagram', 'Сарафан': 'Рекомендация (Сарафан)'};
    const code = (k, v) => { const a = D.dict[k] || (D.dict[k] = []); let i = a.indexOf(v); if (i < 0) { a.push(v); i = a.length - 1; } return i; };
    const L = C.main[0] ? C.main[0].length : 0, cube = D.cube.filter(r => r[0] !== 's');
    C.main.forEach(r => { if (r[1] === SUP || r[3] !== OK || !(r[6] >= 0)) return;
      const d = new Date(B0 + r[6] * 864e5).toISOString().slice(0, 10), n = r[L - 2], b = Math.round(r[L - 1] || 0), src = (CD.source || [])[r[5]] || '', pr = L >= 12 ? String((CD.product || [])[r[9]] || '').split(/[,|]/)[0].trim() : '';
      cube.push(['s', d, code('src', RAW[src] || src), code('name', pr && pr !== 'Не указан' ? pr : 'Көрсетілмеген'), n ? Math.round(b / n) : 0, 0, 0, 0, 0, 0, b > 0 ? 0 : 3, n, b]); });
    D.cube = cube; D.crmBased = true;
  })();
  const WD = ['дүйсенбі', 'сейсенбі', 'сәрсенбі', 'бейсенбі', 'жұма', 'сенбі', 'жексенбі'];
  const pad = n => String(n).padStart(2, '0');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'})[c]);
  const nf = (v, sp) => { const s = String(Math.round(Math.abs(v))).replace(/\B(?=(\d{3})+(?!\d))/g, sp || ' '); return (v < 0 ? '−' : '') + s; };
  const tn = v => nf(v, ' ');                                            // text for WhatsApp: ordinary spaces
  const pc1 = (a, b) => b ? (Math.round(a / b * 1000) / 10).toFixed(1).replace('.', ',') + ' %' : '0 %';
  const pc0 = (a, b) => (b ? Math.round(a / b * 100) : 0) + ' %';
  const isoToUTC = s => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));
  const utcToIso = t => { const d = new Date(t); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; };
  const dmy = s => `${s.slice(8, 10)}.${s.slice(5, 7)}.${s.slice(0, 4)}`;
  const wd = s => WD[(new Date(isoToUTC(s)).getUTCDay() + 6) % 7];
  const snapIso = String(D.snapshot || '').slice(0, 10), snapTxt = D.snapshot ? `${D.snapshot.slice(8, 10)}.${D.snapshot.slice(5, 7)} ${D.snapshot.slice(11, 16)}` : '—';
  // available days
  let minD, maxD;
  if (kind === 'leads') { minD = D.base; maxD = snapIso || D.base; }
  else { const ds = D.cube.map(r => r[1]).filter(d => !snapIso || d <= snapIso).sort(); minD = ds[0] || snapIso; maxD = snapIso || ds[ds.length - 1]; }   // future-dated platform rows are not offered
  const yesterday = snapIso ? utcToIso(isoToUTC(snapIso) - 864e5) : maxD;
  let day = /^\d{4}-\d{2}-\d{2}$/.test(cfg.date) ? cfg.date : (yesterday >= minD ? yesterday : maxD);
  if (day < minD) day = minD; if (day > maxD) day = maxD;

  /* ---------- report 1: leads ---------- */
  function leads(iso) {
    const di = Math.round((isoToUTC(iso) - isoToUTC(D.base)) / 864e5), F = D.dict.funnel, O = D.dict.outcome, S = D.dict.source;
    const SUP = F.indexOf('Техподдержка'), OP = F.indexOf('ОП'), MAIL = [F.indexOf('Рассылка'), F.indexOf('Рассылка 5-6 класс Алтын сақа')];
    const WON = O.indexOf('Успешно'), LOST = O.indexOf('Не реализовано');
    const z = () => ({}), inc = (o, k, n) => { o[k] = (o[k] || 0) + n; };
    const STG = D.dict.stage || [];
    const R = {newL: 0, newIn: 0, newS: 0, src: z(), mail: 0, mailOp: 0, won: 0, wonS: 0, lost: 0, lostSrc: z(), mailLost: 0, mailLostSrc: z(), sNewSrc: z(), sWon: 0, sLost: 0, sClosed: 0};
    D.main.forEach(r => {
      const c = r[0], fu = r[1], oc = r[3], so = S[r[5]], cl = r[6], n = r[r.length - 2];
      if (c === di) {
        if (fu === SUP) { R.newS += n; inc(R.sNewSrc, so, n); }
        else { R.newL += n; inc(R.src, so, n); if (so === 'Рассылка') { R.mail += n; if (fu === OP) R.mailOp += n; }
          if (so !== 'Рассылка' && (fu === OP || (so && so !== 'Не указан'))) R.newIn += n; }   // кіріс — «Қайдан келді» бойынша (Атласпен бірдей, 14.10.2026)
      }
      if (cl === di) {
        if (fu === SUP) { R.sClosed += n; if (oc === WON) R.sWon += n; if (oc === LOST) R.sLost += n; }
        else { if (oc === WON) { R.won += n; R.wonS += r[r.length - 1]; } if (oc === LOST) { R.lost += n; inc(R.lostSrc, so, n); if (MAIL.indexOf(fu) >= 0) { R.mailLost += n; inc(R.mailLostSrc, so, n); } } }
      }
    });
    R.g = k => R.src[k] || 0;
    // CRM толтыру: закрытые в этот день — что не заполнено, по менеджерам; оплаты из ОП кестесі за этот день
    if (D.fill) {
      const MN = D.dict.manager || [], FL = {won: 0, sum: 0, noB: 0, noP: 0, noSrc: 0, lost: 0, noR: 0, man: {}};
      D.fill.forEach(r => { if (r[0] !== di || r[1] === SUP) return; const n = r[r.length - 2], b = r[r.length - 1], nm = MN[r[2]] || '—';
        const e = FL.man[nm] || (FL.man[nm] = {won: 0, sum: 0, noB: 0, noP: 0, noSrc: 0, lost: 0, noR: 0});
        if (r[3] === WON) { FL.won += n; FL.sum += b; e.won += n; e.sum += b; if (!r[4]) { FL.noB += n; e.noB += n; } if (!r[5]) { FL.noP += n; e.noP += n; } if (!r[6]) { FL.noSrc += n; e.noSrc += n; } }
        else if (r[3] === LOST && r[1] === OP) { FL.lost += n; e.lost += n; if (!r[7]) { FL.noR += n; e.noR += n; } } });   // причина отказа — только воронка ОП (рассылки закрываются массово)
      const total = FL.won * 3 + FL.lost, miss = FL.noB + FL.noP + FL.noSrc + FL.noR;
      FL.score = total ? (total - miss) / total : null; FL.miss = miss; R.fill = FL;
    }
    R.sheet = D.sheetDays ? (D.sheetDays[iso] || [0, 0]) : null;
    return R;
  }

  // «Жауап күтіп тұр»: open leads in the stages where the client waits for the manager; days without any change in amoCRM («Изменена»), as of the snapshot.
  // ОП work stages (недозвон, консультация, счёт) are counted apart when nothing moved for 3+ days. The state is the current one, not the report day's.
  const WAIT_ST = {'ОП': ['НОВАЯ ЗАЯВКА', 'Неразобранное'], 'Рассылка': ['Неразобранное', 'Жауап берді'], 'Рассылка 5-6 класс Алтын сақа': ['Неразобранное', 'жауап берген']};
  const WORK_ST = {'ОП': ['НЕДОЗВОН', 'ПРОВЕДЕННЫЕ КОНСУЛЬТАЦИИ', 'СЧЕТ ВЫСТАВЛЕН']};
  const DBK = ['1 күнге дейін', '1–2 күн', '3–6 күн', '7+ күн'], DBC = ['#0ca678', '#f08c00', '#e8590c', '#e03131'];
  const dBucket = d => d < 0 ? 3 : d < 1 ? 0 : d < 3 ? 1 : d < 7 ? 2 : 3;
  // hover / tap / focus tooltip: data-ct = {h, big, c, rows: [{n, v, c}], note}
  const ct = o => ` data-ct="${esc(JSON.stringify(o))}" tabindex="0"`;
  if (typeof document !== 'undefined' && document.addEventListener && !document.__ct) {
    document.__ct = 1; let tipEl = null, cur = null;
    const E = x => String(x == null ? '' : x).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'})[c]);
    const show = (el, ev) => {
      if (!tipEl) { tipEl = document.createElement('div'); tipEl.id = 'ctip'; tipEl.setAttribute('role', 'tooltip'); document.body.appendChild(tipEl); }
      if (el !== cur) { cur = el; const o = JSON.parse(el.getAttribute('data-ct'));
        tipEl.innerHTML = `<div class="h">${o.c ? `<i style="background:${E(o.c)}"></i>` : ''}${E(o.h)}</div>${o.big ? `<div class="b">${E(o.big)}</div>` : ''}${(o.rows || []).length ? `<div class="r">${o.rows.map(r => `<div>${r.c ? `<i style="background:${E(r.c)}"></i>` : '<i></i>'}<span>${E(r.n)}</span><b>${E(r.v)}</b></div>`).join('')}</div>` : ''}${o.note ? `<div class="n">${E(o.note)}</div>` : ''}`; }
      const r = el.getBoundingClientRect(), t = tipEl.getBoundingClientRect(), M = 8, x0 = ev && ev.clientX != null && r.width > 160 ? ev.clientX : r.left + r.width / 2;
      let x = x0 - t.width / 2, y = r.top - t.height - 10; if (y < M) y = r.bottom + 10;
      x = Math.max(M, Math.min(x, innerWidth - t.width - M)); tipEl.style.transform = `translate(${Math.round(x)}px,${Math.round(y)}px)`; tipEl.classList.add('on');
    };
    const hide = () => { cur = null; if (tipEl) tipEl.classList.remove('on'); };
    document.addEventListener('pointermove', e => { if (e.pointerType === 'touch') return; const el = e.target.closest && e.target.closest('[data-ct]'); if (el) show(el, e); else hide(); });
    document.addEventListener('pointerdown', e => { const el = e.target.closest && e.target.closest('[data-ct]'); if (el && e.pointerType === 'touch') show(el, e); else if (!el) hide(); });
    document.addEventListener('focusin', e => { const el = e.target.closest && e.target.closest('[data-ct]'); if (el) show(el); });
    document.addEventListener('focusout', hide); addEventListener('scroll', hide, {passive: true});
  }
  const FSHORT = {'ОП': 'ОП', 'Рассылка': 'Рассылка', 'Рассылка 5-6 класс Алтын сақа': 'Рассылка 5–6'};
  const SNAPS = {}, SNAP_API = cfg.api || cfg.url || '';   /* өткен күн — сол күннің срезі «Срездер» парағынан (иесі 15.10.2026) */
  function snapFor(dd) { if (!dd || dd >= snapIso || !SNAP_API) return null; if (!(dd in SNAPS)) { SNAPS[dd] = 'loading'; fetch(SNAP_API + '?snap=' + dd, {redirect: 'follow'}).then(r => r.json()).then(j => { SNAPS[dd] = j && j.idle && j.idle.length ? j : 'none'; render(); }).catch(() => { SNAPS[dd] = 'none'; render(); }); } return SNAPS[dd]; }
  function idleNamed(dd) {
    const s = snapFor(dd);
    if (s && typeof s === 'object') return {rows: s.idle, at: `${s.d.slice(8, 10)}.${s.d.slice(5, 7)} ${s.t}`, past: true};
    if (!D.idle) return null;
    const FN = D.dict.funnel, ST = D.dict.stage, MN = D.dict.manager || [];
    return {rows: D.idle.map(r => [MN[r[0]] || '—', FN[r[1]], ST[r[2]], r[3], r[4], r[r.length - 1]]), at: snapTxt, past: false, loading: s === 'loading'};
  }
  function waiting(dd) {
    const I = idleNamed(dd); if (!I) return null;
    const W = {n: 0, b: [0, 0, 0, 0], max: 0, stuck: 0, noTask: 0, man: {}, st: {}, stm: {}, bst: [{}, {}, {}, {}], at: I.at, past: I.past, loading: I.loading, late: {}};
    I.rows.forEach(r => {
      const [nm0, fn, stFull, d, t] = r, n = r[r.length - 1], sn = String(stFull || '').split('|').slice(1).join('|');
      if (t === 2) W.late[nm0] = (W.late[nm0] || 0) + n;
      const wait = (WAIT_ST[fn] || []).indexOf(sn) >= 0, stuck = !wait && (WORK_ST[fn] || []).indexOf(sn) >= 0 && d >= 3;
      if (!wait && !stuck) return;
      const nm = nm0 || '—', e = W.man[nm] || (W.man[nm] = {n: 0, b: [0, 0, 0, 0], max: 0, stuck: 0, st: {}, bs: [{}, {}, {}, {}], stuckSt: {}});
      if (stuck) { W.stuck += n; e.stuck += n; e.stuckSt[sn] = (e.stuckSt[sn] || 0) + n; return; }
      const bi = dBucket(d), sk = (FSHORT[fn] || fn) + ' · ' + sn;
      W.n += n; W.b[bi] += n; e.n += n; e.b[bi] += n; e.st[sk] = (e.st[sk] || 0) + n; W.st[sk] = (W.st[sk] || 0) + n; e.bs[bi][sk] = (e.bs[bi][sk] || 0) + n;
      const sm = W.stm[sk] || (W.stm[sk] = {}); (sm[nm] || (sm[nm] = [0, 0, 0, 0]))[bi] += n; W.bst[bi][sk] = (W.bst[bi][sk] || 0) + n;
      if (d > e.max) e.max = d; if (d > W.max) W.max = d; if (t === 0) W.noTask += n;
    });
    W.list = Object.entries(W.man).filter(([, e]) => e.n || e.stuck).sort((a, b) => (b[1].b[3] - a[1].b[3]) || (b[1].b[2] - a[1].b[2]) || (b[1].n - a[1].n));
    return W;
  }
  const dTxt = d => d >= 60 ? '60+ күн' : d < 1 ? '1 күнге жетпейді' : d + ' күн';
  const SRC_TXT = {'Рассылка': ['🔗', 'Рассылка'], 'Instagram': ['🌱', 'Instagram органикасы'], 'WhatsApp': ['🌱', 'WhatsApp органикасы'], 'Таргет': ['🎯', 'Instagram / Facebook таргеті'],
    'Сарафан': ['🤝', 'Ұсыным арқылы'], 'Самост. обращение': ['📱', 'Өздігінен хабарласқан'], 'Не указан': ['▫️', 'Дереккөзі көрсетілмеген']};
  const SRC_ORDER = ['Рассылка', 'Instagram', 'WhatsApp', 'Таргет', 'Сарафан', 'Самост. обращение', 'Не указан'];
  const SRC_COL = {'Рассылка': 'var(--ind)', 'Таргет': 'var(--amb)', 'WhatsApp': 'var(--teal)', 'Instagram': 'var(--vio)', 'Сарафан': '#1098ad', 'Самост. обращение': '#ae3ec9', 'Не указан': 'var(--sl)'};
  const nzSrc = o => SRC_ORDER.concat(Object.keys(o).filter(k => SRC_ORDER.indexOf(k) < 0)).filter(k => o[k]);
  // WhatsApp text for leads: blocks with separators, text progress bars, yesterday for comparison, conclusion at the end
  const WLINE = '━━━━━━━━━━━━━━━━';
  const opTasks = e => [e.noB && `${e.noB} сатылған сделкаға соманы жазыңыз («Бюджет»)`, e.noP && `${e.noP} сатылған сделкаға өнімді таңдаңыз («Продукт»)`,
    e.noSrc && `${e.noSrc} сатылған сделкаға «Қайдан келді» өрісін толтырыңыз`, e.noR && `${e.noR} бас тартқан сделкаға себебін көрсетіңіз`].filter(Boolean);
  const manShort = nm => { const w = String(nm).trim().split(/\s+/); return w.length > 1 ? w[0] + ' ' + w[1].charAt(0) + '.' : w[0]; };   // «Сапарова Айым Ерланқызы» → «Сапарова А.»
  const tbar = (f, n) => { n = n || 10; const k = Math.max(0, Math.min(n, Math.round(f * n))); return '▰'.repeat(k) + '▱'.repeat(n - k); };
  const arrow = () => '';   // a report shows numbers, not up/down signals
  // иесінің шаблоны (14.10.2026): 1 жаңа өтінімдер · 2 жабылғандар · 3 тексеру керек · 4 әр менеджерге «тексеру / түзету» · 5 төлемдерді салыстыру (толық) · 6 техқолдау.
  // «бүгін/кеше» орнына күндер — мәтін келесі күні жіберіледі. Атластағы leadsDayText-пен бірдей ереже.
  const lcw = s => String(s || '').toLowerCase().replace(/ё/g, 'е').trim();
  const isoAdd = (s, n) => utcToIso(isoToUTC(s) + n * 864e5);
  function payVsCrm(a, b) {   // ОП кестесі («Енгізген қызметкер», төлем күні) мен CRM «Успешно» (жабылған күні) — менеджер бойынша; кестедегі қысқа аты CRM толық атындағы сөзбен сәйкестендіріледі
    const SM = D.sheetMan; if (!SM || !Array.isArray(SM.rows)) return null;
    const MN = D.dict.manager || [], F = D.dict.funnel, WON = D.dict.outcome.indexOf('Успешно'), SUP = F.indexOf('Техподдержка');
    let first = Infinity; D.main.forEach(r => { if (r[0] >= 0 && r[0] < first) first = r[0]; });
    if (isFinite(first) && a < isoAdd(D.base, first)) a = isoAdd(D.base, first);   // CRM деректері басталғаннан бері ғана
    const out = {}, row = (k, nm) => out[k] || (out[k] = {key: k, name: nm, pn: 0, ps: 0, cn: 0, cb: 0, other: false});
    const toCrm = e => { const w = lcw(e); return w ? MN.find(m => lcw(m).split(/\s+/).indexOf(w) >= 0) || null : null; };
    SM.rows.forEach(([d, e, n, sm]) => { if (d < a || d > b) return; const en = String((SM.emp || [])[e] || '').trim() || 'Көрсетілмеген', k = toCrm(en), r = row(k || '~' + en, en); r.pn += n; r.ps += sm; if (!k) r.other = true; });
    const ai = Math.round((isoToUTC(a) - isoToUTC(D.base)) / 864e5), bi = Math.round((isoToUTC(b) - isoToUTC(D.base)) / 864e5);
    D.main.forEach(r => { if (r[6] < ai || r[6] > bi || r[1] === SUP || r[3] !== WON) return; const k = MN[r[4]] || '—', x = row(k, out[k] ? out[k].name : manShort(k)); x.cn += r[r.length - 2]; x.cb += r[r.length - 1]; });
    return out;
  }
  function leadsText(iso, R, link, short) {   /* иесі 17.10: «отчёты нужно сократить» — short: тек «ҚЫСҚАША» + «ЛИДТЕР ҚАЙДАН КЕЛДІ?» */   /* иесінің сатылым шаблонымен бір үлгіде (15.10.2026): қысқаша · қайдан келді · жабылғандар · тексеру · менеджерлер · төлем ↔ CRM · техқолдау · ескертпе және күн қорытындысы */
    const SEP = '━━━━━━━━━━━━━━', L = [], d5 = dmy(iso).slice(0, 5), pi = isoAdd(iso, -1), P = pi >= minD ? leads(pi) : null, p5 = dmy(pi).slice(0, 5), W = waiting(iso), UNK = 'Не указан';
    const nm2 = n => { const w = String(n || '').trim().split(/\s+/); return w.length > 2 ? w[1] + ' ' + w[0] : String(n || ''); };   /* «Аты Тегі» */
    const sg = v => (v > 0 ? '+' : v < 0 ? '−' : '±') + tn(Math.abs(v));
    const SE = k => /таргет/i.test(k) ? '🎯' : /whatsapp/i.test(k) ? '💬' : /instagram/i.test(k) ? '📸' : /telegram/i.test(k) ? '✈️' : /рассылка/i.test(k) ? '🔗' : /сарафан|рекоменд/i.test(k) ? '🗣' : /вебинар/i.test(k) ? '🎥' : '▫️';
    const wait3 = W ? W.b[2] + W.b[3] : null;
    const head = () => [`📥 Жаңа өтінім — *${tn(R.newIn)}*${P ? ` (${p5}: ${tn(P.newIn)} · ${R.newIn === P.newIn ? 'өзгеріс жоқ' : sg(R.newIn - P.newIn)})` : ''}`, `✅ Сәтті жабылды — *${tn(R.won)}* · ${tn(R.wonS)} ₸ (CRM бюджеті)`, `❌ Бас тартты — *${tn(R.lost)}*`].concat(W ? [`⏳ Жауап күтіп тұр — *${tn(W.n)}* (3+ күн — ${tn(wait3)})`] : []);
    L.push('📊 *ЗЕРДЕЛІ APP — ЛИДТЕР ЕСЕБІ*', '', `📅 *Күні:* ${dmy(iso)}`, ...(link ? [`🔗 *Есеп:* ${link}`] : []), '', '🎯 *ҚЫСҚАША*', ...head());
    /* 1 */
    const MAIN = ['Таргет', 'WhatsApp', 'Instagram', 'Рассылка'], oth = Object.entries(R.src).filter(([k]) => MAIN.indexOf(k) < 0 && k !== UNK).reduce((t, [, v]) => t + v, 0);
    L.push('', SEP, '', short ? '📥 *ЛИДТЕР ҚАЙДАН КЕЛДІ?*' : '📥 *1. ЛИДТЕР ҚАЙДАН КЕЛДІ?*', `_CRM-дағы барлық жаңа сделка — ${tn(R.newL)} (рассылканы қоса)_`, '');
    MAIN.forEach(k => L.push(`${SE(k)} ${k} — ${tn(R.g(k))}${R.newL ? ' · ' + pc0(R.g(k), R.newL) : ''}`));
    L.push(`▫️ Басқа арналар — ${tn(oth)}`, `▫️ Көрсетілмеген — ${tn(R.g(UNK))}`, '', `📊 *Барлығы: ${tn(R.newL)}*`, `🔗 Рассылкадан ОП-ға өтті — *${tn(R.mailOp)}*`);
    if (short) return L.join('\n');
    /* 2 */
    const lb = R.lostSrc, lo = Object.entries(lb).filter(([k]) => ['Таргет', 'WhatsApp', 'Рассылка', 'Instagram', UNK].indexOf(k) < 0).reduce((t, [, v]) => t + v, 0);
    L.push('', SEP, '', `✅ *2. ${d5} ЖАБЫЛҒАН СДЕЛКАЛАР*`, '', `✅ Сәтті — *${tn(R.won)}* · ${tn(R.wonS)} ₸ (CRM бюджеті, төлем емес)`, `❌ Сәтсіз — *${tn(R.lost)}*`, `📊 Сәтті үлесі — *${R.won + R.lost ? pc0(R.won, R.won + R.lost) : '—'}*`);
    if (R.lost) L.push(`▫️ Сәтсіз қайдан: ${['Таргет', 'WhatsApp', 'Рассылка', 'Instagram'].filter(k => lb[k]).map(k => k + ' ' + tn(lb[k])).concat(lo ? ['басқа ' + tn(lo)] : [], lb[UNK] ? ['көрсетілмеген ' + tn(lb[UNK])] : []).join(' · ')}`);
    L.push('_Жабылғандар бұрын түскен болуы мүмкін — бұл жаңа өтінімнің конверсиясы емес._');
    /* 3 */
    L.push('', SEP, '', '⏳ *3. ТЕКСЕРУ КЕРЕК*', W && W.past ? `_${W.at} срезі_` : iso < snapIso ? `_${snapTxt} жағдайы (қазіргі күй — ${d5} күнгі емес)_` : `_${snapTxt} жағдайы_`, '');
    if (!W) L.push('▫️ Дерек жоқ — Apps Script жаңа нұсқасы керек.');
    else L.push(`📋 «НОВАЯ ЗАЯВКА» / «Жауап берді» — *${tn(W.n)}*`, `⚠️ 3+ күн жаңартылмаған — *${tn(wait3)}*`, `🔴 7+ күн жаңартылмаған — *${tn(W.b[3])}*`);
    /* 4 — бір менеджер, бір жол: тек әрекет керек болса */
    const late = W ? W.late : {};
    const pv = payVsCrm(iso, iso), pv7 = payVsCrm(isoAdd(iso, -6), iso), FM = R.fill ? R.fill.man : {};
    const names = new Set([].concat(W ? Object.keys(W.man) : [], Object.keys(FM), Object.keys(late), pv ? Object.values(pv).filter(r => !r.other).map(r => r.key) : []));
    const rows4 = [...names].filter(nm => nm !== '—').map(nm => { const e = W && W.man[nm], f = FM[nm] || {}, p = pv && pv[nm], parts = [], old = e ? e.b[2] + e.b[3] : 0;
      if (old) parts.push(`⏳ 3+ күн — ${tn(old)}`); if (late[nm]) parts.push(`⏰ мерзімі өткен — ${tn(late[nm])}`);
      const empty = [['Бюджет', f.noB], ['Продукт', f.noP], ['Қайдан келді', f.noSrc]].filter(([, n]) => n); if (empty.length) parts.push(`✏️ бос: ${empty.map(([k, n]) => k + ' ' + tn(n)).join(', ')}`);
      if (p && R.sheet && p.pn !== p.cn) parts.push(`💳 тізім ${tn(p.pn)} / CRM ${tn(p.cn)}`);
      return {nm, parts, w: old + (late[nm] || 0) * 2 + empty.length}; });
    const todo = rows4.filter(x => x.parts.length).sort((a, b) => b.w - a.w), okN = rows4.filter(x => !x.parts.length);
    L.push('', SEP, '', '👤 *4. МЕНЕДЖЕРЛЕР*', '');
    todo.forEach(x => L.push(`👤 *${nm2(x.nm)}* — ${x.parts.join(' · ')}`));
    if (okN.length) L.push(`✅ Тәртіпте: ${okN.map(x => nm2(x.nm)).join(', ')}`);
    if (!rows4.length) L.push('▫️ Менеджерлер бойынша дерек жоқ.');
    if (todo.some(x => x.parts.some(t => t.startsWith('✏️')))) L.push('', `🔎 amoCRM → Сделки → «Закрыта: ${d5}» + «Ответственный: өзіңіз» → бос өрісті толтырыңыз.`);
    /* 5 */
    if (!R.sheet) L.push('', SEP, '', '💳 *5. САТЫЛЫМ (CRM)*', '', `🗂 CRM «Успешно» — *${tn(R.won)} · ${tn(R.wonS)} ₸* (бюджет)`, '_платформа сатылымы — сатылым карточкасында_');
    else L.push('', SEP, '', '💳 *5. ТӨЛЕМ ↔ CRM*', '', `📒 Оқушы тізімі — *${tn(R.sheet[0]) + ' төлем · ' + tn(R.sheet[1]) + ' ₸'}*`, `🗂 CRM «Успешно» — *${tn(R.won)} · ${tn(R.wonS)} ₸* (бюджет)`);
    if (pv && R.sheet) { const rows = Object.values(pv).filter(r => !r.other && (r.pn || r.cn)).sort((x, y) => (y.pn + y.cn) - (x.pn + x.cn));
      if (rows.length) L.push('', ...rows.map(r => { const w = (pv7 && pv7[r.key]) || {pn: 0, cn: 0}; return `▫️ ${nm2(r.name)} — ${tn(r.pn)} / ${tn(r.cn)} (${sg(r.pn - r.cn)}) · 7 күн ${tn(w.pn)} / ${tn(w.cn)}`; }), '_тізім / CRM — айырма сделка жабылмағанын не төлем жазылмағанын білдіруі мүмкін_'); }
    /* 6 */
    L.push('', SEP, '', '🛠 *6. ТЕХҚОЛДАУ*', '', `📥 Жаңа — ${tn(R.newS)} · ✅ Сәтті — ${tn(R.sWon)} · ❌ Сәтсіз — ${tn(R.sLost)}`);
    /* 7 */
    const warn = []; if (!W) warn.push('Жауап күтіп тұрған лидтер дерегі жоқ — Apps Script жаңа нұсқасы керек.'); /* 16.10: Оқушы тізімі қолданылмайды — ескерту жоқ */
    L.push('', SEP, '', '📝 *7. ЕСКЕРТПЕ*', '');
    if (warn.length) L.push(...warn.map(w => '⚠️ ' + w), '');
    L.push('🎯 *КҮН ҚОРЫТЫНДЫСЫ:*', ...head());
    return L.join('\n');
  }


  /* ---------- report 2: sales ---------- */
  const lc = s => String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[«»"„“”]/g, '');
  const keyMatch = (keys, name) => { const n = lc(name); return String(keys || '').split(';').map(x => x.trim()).filter(Boolean).some(alt => alt.split('+').map(w => lc(w.trim())).filter(Boolean).every(w => n.indexOf(w) >= 0)); };
  const shortName = s => {
    const parts = String(s || '').split(' · '), tariff = (parts[1] || '').trim();
    let t = parts[0].split(' / ')[0].replace(/\[[^\]]*\]/g, '').replace(/\d+\s*-?\s*(сынып|класс)\.?/gi, '')
      .replace(/\d+\s*-?\s*вариант[^\s)]*/gi, '').replace(/[«»"„“”]/g, '').replace(/\s+([,.;])/g, '$1').replace(/([,.;])(\s*[,.;])+/g, '$1').replace(/^[\s,.;|:-]+|[\s,.;|:-]+$/g, '').replace(/\s{2,}/g, ' ').trim();
    return (t || '—') + (tariff ? ' · ' + tariff : '');
  };
  // same rule as the server's zopPayStatus_ (copied: this client is shipped as a separate page). tests/ check that both agree.
  const PAY_L = {paid: ['2', 'paid', 'success', 'succeeded', 'successful', 'completed', 'complete', 'approved', 'оплачен', 'оплачено', 'оплачена', 'төленді', 'төленген', 'сәтті'],
    unpaid: ['unpaid', 'not paid', 'notpaid', 'pending', 'awaiting', 'waiting', 'new', 'created', 'issued', 'не оплачен', 'не оплачено', 'ожидает оплаты', 'төленбеген', 'күтілуде'],
    rejected: ['failed', 'fail', 'error', 'declined', 'cancelled', 'canceled', 'refunded', 'refund', 'expired', 'отменен', 'отменён', 'отклонен', 'отклонён', 'возврат', 'қайтарылды', 'бас тартылды']};
  const payStatus = v => { let t = String(v == null ? '' : v).trim().toLowerCase().replace(/ё/g, 'е').replace(/[_\-]+/g, ' ').replace(/\s+/g, ' '); if (!t) return 'unknown';
    if (/^mon jan 01 1900/.test(t) || /^1900 01 01/.test(t) || t === '01.01.1900') t = '2';
    for (const k of ['paid', 'unpaid', 'rejected']) if (PAY_L[k].indexOf(t) >= 0) return k; return 'unknown'; };
  const PAID = st => payStatus(st) === 'paid';
  const skipReason = r => { if (r[0] !== 'p') return ''; if (r[7] !== 1) return 'invoice жоқ'; const ps = payStatus(D.dict.pstatus[r[6]] || '');
    if (ps !== 'paid') return {unpaid: 'төленбеген', rejected: 'бас тартылған / қайтарылған', unknown: 'статус белгісіз'}[ps];
    return (r.length >= 13 ? r[9] : 0) !== 0 ? 'төлем күні жоқ (шот күні)' : ''; };
  const TOT = /^(барлығы|итого|всего)$/i, OTHER = 'Басқа өнімдер';
  const fromCatalog = !!(D.plan && D.plan.catalog && D.plan.catalog.length);   // the leads card has no D.plan
  // каталог: лист «Өнімдер»; пока его нет — продукты из «Жоспар» в том же виде
  const catalog = () => fromCatalog ? D.plan.catalog : D.plan.products.map(p => ({sec: p.group, name: p.name, keys: p.keys || (/алтын\s*са[қк]а/i.test(p.group) && !TOT.test(p.name) ? 'алтын сақа;алтын сака' : ''), price: p.price, always: true, plan: p.plan, coef: p.coef || 1}));
  function sales(iso) {
    const N = D.dict.name, SR = D.dict.src, ST = D.dict.pstatus, TY = D.dict.ptype;
    const secs = [], secOf = name => secs.find(x => x.name === name) || (secs.push({name, items: [], plan: null}), secs[secs.length - 1]), items = [];
    catalog().forEach(c => { const s = secOf(String(c.sec || OTHER).trim()), nm = String(c.name || '').trim();
      if (TOT.test(nm)) { if (c.plan != null && c.plan !== '') s.plan = +c.plan; return; }
      if (!c.keys && !c.real && (c.price == null || c.price === '')) return;   // a rule needs a name, a word or a price
      let x = s.items.find(y => y.p.name === nm);   // same report name in one section = one line (plans add up)
      const hasPlan = c.plan != null && c.plan !== '';
      if (!x) { x = {p: {name: nm, price: c.price, pprice: c.pprice, plan: hasPlan ? +c.plan : null, coef: c.coef || 1, always: !!c.always}, n: 0, s: 0, pays: 0, prices: {}}; s.items.push(x); }
      else { if (c.always) x.p.always = true; if (hasPlan) x.p.plan = (x.p.plan || 0) + +c.plan; if (x.p.price != null && x.p.price !== '' && +x.p.price !== +c.price) x.p.price = x.p.plan ? x.p.price : null; }
      items.push({rule: c, x}); });
    const R = {rows: 0, sum: 0, secs, src: {}, plat: {lo: 0, loS: 0, hi: 0, hiS: 0}, skipped: {}, skippedS: {}, by: {s: {n: 0, s: 0}, p: {n: 0, s: 0}}, refPrice: 0};
    const srcMap = {}; D.plan.sources.forEach(([v, l]) => { srcMap[lc(v)] = l; });
    const platLine = srcMap['(платформа)'] || 'Платформадан өзі сатып алды', emptyLine = srcMap['(бос)'] || 'Басқа (Не указан)';
    D.cube.forEach(r => {
      if (r[1] !== iso) return;
      const from = r[0], src = SR[r[2]] || '', name = N[r[3]] || '', price = r[4], n = r[r.length - 2], sum = r[r.length - 1];
      const why = skipReason(r); if (why) { R.skipped[why] = (R.skipped[why] || 0) + n; R.skippedS[why] = (R.skippedS[why] || 0) + sum; return; }
      R.rows += n; R.sum += sum; R.by[from].n += n; R.by[from].s += sum; if (r.length >= 13 && r[10] === 2) R.refPrice += n;   // amount is a list price, not a paid sum
      const line = from === 'p' ? platLine : (src ? (srcMap[lc(src)] || src) : emptyLine);
      R.src[line] = (R.src[line] || 0) + n;
      if (from === 'p') { if (price <= 1000) { R.plat.lo += n; R.plat.loS += sum; } else { R.plat.hi += n; R.plat.hiS += sum; } }
      // product: first catalog row whose words are in the name and whose price fits; otherwise «Басқа өнімдер» under its own name
      const hay = name + ' ' + (TY[r[5]] || ''), lname = lc(name).trim();
      const nname = lc(String(name).replace(/\s*,\s*/g, '; ')).trim();   // the list holds names without commas («, » separates several choices)
      const realHas = cell => lc(cell).trim() === lname || String(cell).split(/,\s+/).some(it => lc(it).trim() === nname);
      const hit = items.find(({rule: c}) => (c.price == null || c.price === '' || +c.price === +price) && ((c.keys || c.real ? (c.real && realHas(c.real)) || (c.keys && keyMatch(c.keys, hay)) : true)));
      let x = hit && hit.x;
      if (!x) { const s = secOf(OTHER), label = shortName(name || '—'); x = s.items.find(y => y.auto && y.p.name === label);
        if (!x) { x = {p: {name: label, coef: 1}, n: 0, s: 0, pays: 0, prices: {}, auto: true}; s.items.push(x); } }
      x.n += n * (x.p.coef || 1); x.s += sum; x.pays += n; x.prices[price] = (x.prices[price] || 0) + n;
      const rl = shortName(name || '—'), rw = (x.raw || (x.raw = {}))[rl] || (x.raw[rl] = {n: 0, s: 0}); rw.n += n; rw.s += sum;
    });
    secs.forEach(s => {
      const grp = x => /^математика/i.test(x.p.name) ? 0 : /^ағылшын/i.test(x.p.name) ? 1 : 0;
      s.vis = s.items.filter(x => x.p.always || x.n).sort((a, b) => (grp(a) - grp(b)) || (a.auto && b.auto ? b.n - a.n : 0));
      s.n = s.items.reduce((a, x) => a + x.n, 0); s.s = s.items.reduce((a, x) => a + x.s, 0);
      const ip = s.items.filter(x => x.p.plan != null && x.p.plan !== '');
      s.itemPlans = ip.length > 0; s.planN = s.plan != null ? s.plan : ip.length ? ip.reduce((a, x) => a + +x.p.plan, 0) : null;
      s.planS = ip.length ? ip.reduce((a, x) => a + x.p.plan * (x.p.price || x.p.pprice || 0), 0) : null;
    });
    const rank = s => s.name === OTHER ? 2 : s.planN != null ? 0 : 1; secs.sort((a, b) => rank(a) - rank(b));   // stable: planned sections first, «Басқа өнімдер» last
    R.secsVis = secs.filter(s => s.vis.length || s.planN);
    R.main = secs.find(s => s.itemPlans) || null;   // the section whose products have their own plans («Негізгі өнімдер»)
    const order = []; D.plan.sources.forEach(([, l]) => { if (order.indexOf(l) < 0) order.push(l); });
    Object.keys(R.src).forEach(l => { if (order.indexOf(l) < 0) order.splice(Math.max(0, order.indexOf(emptyLine)), 0, l); });
    const ALWAYS = ['Мектепке бару', 'Маркетинг (органика)', 'Маркетинг (таргет)', 'Рассылка', platLine, 'EduApp орталықтары', emptyLine];
    R.srcOrder = order.filter(l => R.src[l] || ALWAYS.indexOf(l) >= 0); R.platLine = platLine;
    return R;
  }
  const EMO = [['мега', '📚'], ['математика.*тоқсан', '📘'], ['математика.*күнд', '⚡'], ['ағылшын.*тоқсан', '🇬🇧'], ['ағылшын.*күнд', '⚡']];
  const emo = name => (EMO.find(([re]) => new RegExp(re, 'i').test(name)) || [, '📗'])[1];
  const AS_EMO = ['🎓', '💥', '📚', '🚀', '🎁', '📨', '🎯', '🧩'];
  const secEmo = s => /негізгі/i.test(s.name) ? '📚' : /пробный/i.test(s.name) ? '🧪' : /түлек/i.test(s.name) ? '🎓' : /алтын/i.test(s.name) ? '🏆' : s.name === OTHER ? '📦' : '🗂';
  const priceTxt = x => /₸/.test(x.p.name) ? '' : x.p.price != null && x.p.price !== '' ? nf(x.p.price) + ' ₸' : Object.keys(x.prices).length ? Object.keys(x.prices).sort((a, b) => a - b).map(p => nf(+p) + ' ₸').join(', ') : x.p.pprice ? nf(x.p.pprice) + ' ₸' : '';
  const withPrice = x => x.p.name + (priceTxt(x) ? ' · ' + priceTxt(x) : '');
  const secHead = (s, T) => `${secEmo(s)} ${s.name} — ` + (s.planN ? `${T(pc0(s.n, s.planN))} (${tn(s.n)} / ${tn(s.planN)} дана · ${tn(s.s)}${s.planS ? ' / ' + tn(s.planS) : ''} ₸)` : T(`${tn(s.n)} дана · ${tn(s.s)} ₸`));
  const SRC2_EMO = l => /мектеп/i.test(l) ? '🏫' : /органика/i.test(l) ? '🌱' : /таргет/i.test(l) ? '🎯' : /рассылка/i.test(l) ? '🔗' : /платформа/i.test(l) ? '📱' : /орталық/i.test(l) ? '🏢' : /вебинар/i.test(l) ? '🎥' : '▫️';
  const itemLbl = (s, x) => { const sn = s.name.split(' · ')[0], n = x.p.name;
    if (n.length > sn.length + 3 && lc(n).startsWith(lc(sn)) && !/^\s*\(/.test(n.slice(sn.length))) { const t = n.slice(sn.length).replace(/^[\s·—:,.-]+/, ''); return t.charAt(0).toUpperCase() + t.slice(1); }
    return n; };
  const subsOf = (x, s) => { if (!x.raw || (x.p.price != null && x.p.price !== '')) return [];
    const e = Object.entries(x.raw).filter(([, r]) => r.n).sort((a, b) => b[1].n - a[1].n), generic = lc(x.p.name) === lc(s.name) || /басқа/i.test(x.p.name);
    return e.length > 1 || generic && !(e.length === 1 && lc(e[0][0]) === lc(x.p.name)) ? e : []; };   // a «basket» line lists what was really bought
  // WhatsApp text for sales — лидтер есебімен бірдей пішім (иесі, 14.10.2026): нөмірленген бөлімдер, «Атауы — *сан*», айырма — сөзбен, жолақ-белгілер жоқ
  function platGap(iso, R) {   /* платформа бұл күні әдеттегіден күрт аз (выгрузка кешігуі): алдыңғы 7 күннің медианасының 25 %-ынан аз */
    const prev = []; for (let k = 1; k <= 7; k++) { const d = utcToIso(isoToUTC(iso) - k * 864e5); if (d >= minD) prev.push(sales(d).by.p.n); }
    prev.sort((a, b) => a - b); const med = prev.length >= 4 ? prev[Math.floor(prev.length / 2)] : 0;
    return med >= 10 && R.by.p.n < med * .25 ? {n: R.by.p.n, med} : null;
  }
  function salesText(iso, R, link, short) {   /* short (иесі 17.10): жалпы сома + «ОРЫНДАЛУЫ» */   /* иесінің шаблоны 15.10.2026 (екінші нұсқа): орындалуы · жоспар/факт · жалпы нәтиже · клиенттер қайдан келді (ОП кестесі «Қайдан келді» + платформа) · ескертпе және күн қорытындысы. Лидтер мұнда жоқ. */
    const SEP = '━━━━━━━━━━━━━━', L = [], gap = platGap(iso, R);
    const m = v => { const r = Math.round(v * 100) / 100, t = Math.round(Math.abs(r) * 100) % 100; return tn(Math.trunc(r)) + (t ? ',' + String(t).padStart(2, '0') : '') + ' ₸'; };   /* тиын болса — сақталады */
    const main = R.main, AS = R.secs.find(x => x !== main && /^алтын сақа$/i.test(String(x.name).trim())) || null, others = R.secs.filter(x => x !== main && x !== AS);
    const oth = others.flatMap(sx => sx.items.filter(x => x.n).map(x => ({sx, x}))).sort((a, b) => b.x.n - a.x.n || b.x.s - a.x.s), othN = oth.reduce((t, o) => t + o.x.n, 0), othS = oth.reduce((t, o) => t + o.x.s, 0);
    const priceOf = x => x.p.price != null && x.p.price !== '' ? +x.p.price : (x.p.pprice ? +x.p.pprice : null);
    const sumLines = () => {
      const out = [];
      if (main) out.push(`📚 Негізгі өнімдер — *${main.planN ? pc0(main.n, main.planN) : 'жоспар жоқ'}* (${tn(main.n)}${main.planN ? ' / ' + tn(main.planN) : ''} дана · ${tn(main.s)}${main.planS ? ' / ' + tn(main.planS) : ''} ₸)`);
      if (AS) out.push(`🏆 Алтын сақа — *${AS.planN ? pc0(AS.n, AS.planN) : 'жоспар жоқ'}* (${tn(AS.n)}${AS.planN ? ' / ' + tn(AS.planN) : ''} дана · ${m(AS.s)})`);
      out.push(`📦 Басқа өнімдер — *${tn(othN)} дана · ${m(othS)}*`);
      return out;
    };
    L.push('📊 *ЗЕРДЕЛІ APP — КҮНДЕЛІКТІ ЕСЕП*', '', `📅 *Күні:* ${dmy(iso)}`, ...(link ? [`🔗 *Есеп:* ${link}`] : []), '', `💰 *${gap || R.refPrice ? 'Есептегі жалпы сома' : 'Жалпы сома'}: ${m(R.sum)}* (${tn(R.rows)} дана)`, '', '🎯 *ОРЫНДАЛУЫ*', ...sumLines());
    if (short) return L.join('\n');
    /* 1. Жоспар / факт */
    L.push('', SEP, '', '📈 *1. САТЫЛЫМ — ЖОСПАР / ФАКТ*');
    if (main) {
      const ORD = [/мега/i, /математика.*тоқсан/i, /математика.*күнд/i, /ағылшын.*тоқсан/i, /ағылшын.*күнд/i], oi = x => { const k = ORD.findIndex(re => re.test(x.p.name)); return k < 0 ? 9 : k; };
      main.items.filter(x => x.p.plan != null || x.n).sort((a, b) => oi(a) - oi(b)).forEach(x => {
        const pr = priceOf(x), plan = x.p.plan;
        L.push('', `${emo(x.p.name)} *${x.p.name}${plan ? ' — ' + pc0(x.n, plan) : ''}*`);
        if (pr != null) L.push(`💰 ${m(pr)}`); else if (priceTxt(x)) L.push(`💰 ${priceTxt(x)}`);
        L.push(plan != null ? `🎯 Жоспар: ${tn(plan)} дана${pr != null ? ' — ' + m(plan * pr) : ''}` : '🎯 Жоспар: берілмеген', `✅ Факт: ${tn(x.n)} дана — ${m(x.s)}`);
      });
    }
    if (AS) {
      const aE = n => /рассылка/i.test(n) ? '📨' : /алдын ала|предпрод/i.test(n) ? '▫️' : /толық|комбо|пакет|акция/i.test(n) ? '🎁' : /жеңілдік|скидк/i.test(n) ? '🏷️' : /2\s*тренинг/i.test(n) ? '📗' : /тренажер/i.test(n) ? '💥' : /тест/i.test(n) ? '📝' : /миникурс|курс/i.test(n) ? '📚' : '🎓';
      L.push('', `🏆 *Алтын сақа${AS.planN ? ' — ' + pc0(AS.n, AS.planN) : ''}*`, `🎯 Жоспар: ${AS.planN ? tn(AS.planN) + ' дана' : 'берілмеген'} · ✅ Факт: ${tn(AS.n)} дана — ${m(AS.s)}`);
      AS.vis.forEach(x => { const pr = priceOf(x), c = x.p.coef || 1, nm = itemLbl(AS, x);
        L.push(`${aE(x.p.name)} ${nm}${pr != null && !/₸/.test(nm) ? ' · ' + m(pr) : ''} — ${AS.planN ? pc0(x.n, AS.planN) + ' ' : ''}(${tn(x.n)}${c > 1 && x.n ? ` = ${tn(x.pays)} × ${c}` : ''})`); });
    }
    L.push('', `📦 *Басқа өнімдер — ${tn(othN)} дана · ${m(othS)}*`);
    oth.forEach(({sx, x}) => { const nm = sx.name === OTHER ? x.p.name : (x.p.name === sx.name ? sx.name : sx.name.split(' · ')[0] + ' — ' + itemLbl(sx, x)); L.push(`▫️ ${nm} — ${tn(x.n)} дана — ${m(x.s)}`); });
    L.push('', `🧾 Барлық сатылым: *${tn(R.rows)} дана — ${m(R.sum)}*`);
    /* 2. Жалпы нәтиже (негізгі өнімдер) */
    if (main && main.planN) {
      L.push('', SEP, '', '💰 *2. ЖАЛПЫ НӘТИЖЕ*', '', `📊 Орындалуы: *${pc0(main.n, main.planN)}*`, '', '📦 Дана бойынша:', `🎯 Жоспар: *${tn(main.planN)} дана*`, `✅ Факт: *${tn(main.n)} дана*`);
      if (main.planS) L.push('', '💵 Сома бойынша:', `🎯 Жоспар: *${m(main.planS)}*`, `✅ Факт: *${m(main.s)}*`);
    }
    /* 3. Клиенттер қайдан келді: ОП кестесі — «Қайдан келді» бойынша, платформа — өзі сатып алды (бір тізім, қосындысы = барлық сатылым) */
    const n3 = main && main.planN ? 3 : 2;
    L.push('', SEP, '', `📥 *${n3}. КЛИЕНТТЕР ҚАЙДАН КЕЛДІ?*`, '');
    R.srcOrder.forEach(l => { const v = R.src[l] || 0; L.push(`${SRC2_EMO(l)} ${l} — ${tn(v)}`);
      if (l === R.platLine && v) L.push(`   ▫️ 1 000 ₸ және одан арзан — ${tn(R.plat.lo)} дана — ${m(R.plat.loS)}`, `   ▫️ 1 000 ₸-ден қымбат — ${tn(R.plat.hi)} дана — ${m(R.plat.hiS)}`); });
    L.push('', `📊 *Барлығы: ${tn(R.rows)} дана*`);
    /* 4. Ескертпе: нақты мәселелер (болса) + күн қорытындысы */
    const warn = [];
    if (gap) warn.push(`Платформа: ${tn(gap.n)} төлем (әдетте ~${tn(gap.med)}) — выгрузка толық емес болуы мүмкін, сома нақтыланады.`);
    if (R.refPrice) warn.push(`${tn(R.refPrice)} жазбада төленген сома жоқ — бағасы алынды.`);
    Object.entries(R.skipped || {}).forEach(([why, n]) => { if (n && /белгісіз|күні жоқ/.test(why)) warn.push(`Есепке кірмеді: ${why} — ${tn(n)} жазба · ${m(R.skippedS[why] || 0)}.`); });
    const auto = R.secs.filter(x => x.name === OTHER).flatMap(x => x.items.filter(y => y.auto && y.n));
    if (auto.length) warn.push(`«Өнімдер» тізімінде жоқ атаулар (${auto.length}) «Басқа өнімдер»-ге жазылды — бөлімін тексеру керек.`);
    L.push('', SEP, '', `📝 *${n3 + 1}. ЕСКЕРТПЕ*`, '');
    if (warn.length) L.push(...warn.map(w => '⚠️ ' + w), '');
    L.push('🎯 *КҮН ҚОРЫТЫНДЫСЫ:*', ...sumLines(), `💰 Жалпы сома — *${m(R.sum)}* (${tn(R.rows)} дана)`);
    return L.join('\n');
  }

  /* ---------- drawing (vector, animated on reveal) ---------- */
  const cu = (v, f) => `<span class="cu" data-n="${Math.round(v)}">${f ? f(v) : nf(v)}</span>`;   // counts up when its section appears
  const GR = {good: ['#69db7c', '#2f9e44', 'rgba(47,158,68,.45)'], ind: ['#91a7ff', '#3b5bdb', 'rgba(59,91,219,.45)'], amb: ['#ffc078', '#e8590c', 'rgba(232,89,12,.4)'], red: ['#ffa8a8', '#e03131', 'rgba(224,49,49,.4)'], teal: ['#63e6be', '#0d9488', 'rgba(13,148,136,.45)']};
  const gkey = f => f >= 1 ? 'good' : f >= .6 ? 'ind' : f >= .3 ? 'amb' : 'red';
  const pcol = f => ({good: 'var(--grn)', ind: 'var(--ind)', amb: 'var(--amb)', red: 'var(--red)'})[gkey(f)];
  let gid = 0;
  const bars = (items, opt) => {   // items: {label, v, col, sub, em}
    const max = Math.max(1, ...items.filter(x => !x.sub || opt && opt.subScale).map(x => x.v));
    return '<div class="rows">' + items.map((x, i) => `<div class="row${x.sub ? ' sub' : ''}${x.v ? '' : ' zero'}"><div class="lab" title="${esc(x.label)}">${esc(x.label)}${x.em ? `<em>${esc(x.em)}</em>` : ''}</div>
      <div class="track"><div class="fill" style="--w:${Math.max(0, Math.min(100, x.v / max * 100)).toFixed(2)}%;--c:${x.col || 'var(--ind)'};--d:${i * 70}ms"></div></div><div class="val">${cu(x.v)}</div></div>`).join('') + '</div>';
  };
  // radial gauge: share of plan (or of a total), gradient arc, number in the middle
  const gauge = (title, a, b, meta, opt) => {
    opt = opt || {}; const r = 52, C = 2 * Math.PI * r, f = b ? a / b : 0, k = opt.key || gkey(f), g = GR[k], id = 'gg' + (++gid), pct = b ? Math.round(f * 100) : 0;
    return `<div class="gauge"><div class="gwrap"><svg viewBox="0 0 132 132" aria-hidden="true"><defs><linearGradient id="${id}" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="${g[0]}"/><stop offset="1" stop-color="${g[1]}"/></linearGradient></defs>
      <circle cx="66" cy="66" r="${r}" class="gtrack"/><circle cx="66" cy="66" r="${r}" class="arc" stroke="url(#${id})" stroke-dasharray="${C.toFixed(1)}" transform="rotate(-90 66 66)" style="--C:${C.toFixed(1)};--off:${(C * (1 - Math.min(1, f))).toFixed(1)};--glow:${g[2]};--d:${opt.d || 0}ms"/></svg>
      <div class="gc"><b>${opt.center != null ? opt.center : cu(pct) + '<small>%</small>'}</b>${opt.under ? `<i>${opt.under}</i>` : ''}</div></div><div class="gt">${esc(title)}</div><div class="gm">${meta}</div></div>`;
  };
  const ring = (a, b, colA, colB, center) => `<div class="gauge" style="border:0;background:none;padding:0">${gauge('', a, b, '', {key: 'teal', center}).replace(/^<div class="gauge">/, '').replace(/<div class="gt"><\/div><div class="gm"><\/div><\/div>$/, '')}</div>`;
  // donut of shares with a 2px gap between segments
  const donut = (segs, center, under) => {
    const r = 55, C = 2 * Math.PI * r, tot = segs.reduce((t, x) => t + x.v, 0) || 1, live = segs.filter(x => x.v > 0), gap = live.length > 1 ? 2.5 : 0; let acc = 0;
    const arcs = live.map((x, i) => { const len = Math.max(.5, x.v / tot * C - gap), rot = acc / tot * 360; acc += x.v;
      return `<circle cx="75" cy="75" r="${r}" class="dseg" style="stroke:${x.col};--C:${len.toFixed(2)};--d:${i * 110}ms" stroke-dasharray="${len.toFixed(2)} ${C.toFixed(2)}" transform="rotate(${rot.toFixed(2)} 75 75)"><title>${esc(x.label)}: ${nf(x.v)}</title></circle>`; }).join('');
    return `<div class="dwrap"><svg viewBox="0 0 150 150" aria-hidden="true"><circle cx="75" cy="75" r="${r}" fill="none" stroke="var(--track)" stroke-width="20"/>${arcs}</svg><div class="dc"><b>${center}</b><small>${under || ''}</small></div></div>`;
  };
  const stack = segs => { const t = segs.reduce((s, x) => s + x.v, 0) || 1; return `<div class="stack">${segs.filter(x => x.v).map((x, i) => `<div style="--w:${(x.v / t * 100).toFixed(2)}%;--d:${i * 120}ms;background:linear-gradient(90deg,color-mix(in srgb,${x.col} 60%,#fff),${x.col})" title="${esc(x.label)}: ${nf(x.v)}"></div>`).join('')}</div>`; };
  const legend = (segs, tot) => `<div class="legend">${segs.map(x => `<div class="lg${x.v ? '' : ' zero'}"><i style="background:${x.col}"></i>${esc(x.label)}<b>${cu(x.v)}${x.pct ? `<span class="p">${x.pct}</span>` : tot ? `<span class="p">${tot ? Math.round(x.v / tot * 100) : 0}%</span>` : ''}</b></div>`).join('')}</div>`;
  // treemap: tile area = share; squarified layout in a box with the given width:height ratio
  function treemap(items, ar) {
    const tot = items.reduce((t, x) => t + x.v, 0); if (!tot) return '';
    const W = ar, out = []; let x0 = 0, y0 = 0, w = W, h = 1, row = [];
    const nodes = items.map(x => Object.assign({}, x, {a: x.v / tot * W})).sort((a, b) => b.a - a.a);
    const worst = (r, side) => { const sm = r.reduce((t, n) => t + n.a, 0), mx = Math.max(...r.map(n => n.a)), mn = Math.min(...r.map(n => n.a)); return Math.max(side * side * mx / (sm * sm), sm * sm / (side * side * mn)); };
    const lay = r => { const sm = r.reduce((t, n) => t + n.a, 0);
      if (w >= h) { const cw = sm / h; let yy = y0; r.forEach(n => { const nh = n.a / cw; out.push({n, x: x0, y: yy, w: cw, h: nh}); yy += nh; }); x0 += cw; w -= cw; }
      else { const rh = sm / w; let xx = x0; r.forEach(n => { const nw = n.a / rh; out.push({n, x: xx, y: y0, w: nw, h: rh}); xx += nw; }); y0 += rh; h -= rh; } };
    let i = 0; while (i < nodes.length) { const side = Math.min(w, h), n = nodes[i]; if (!row.length || worst(row.concat([n]), side) <= worst(row, side)) { row.push(n); i++; } else { lay(row); row = []; } }
    if (row.length) lay(row);
    return `<div class="tm" style="aspect-ratio:${W.toFixed(2)}">${out.map((r, k) => { const sh = r.w * r.h / W, cls = sh < .035 ? ' xs' : sh < .09 ? ' sm' : '';
      return `<div class="tmc${cls}" style="left:${(r.x / W * 100).toFixed(3)}%;top:${(r.y * 100).toFixed(3)}%;width:${(r.w / W * 100).toFixed(3)}%;height:${(r.h * 100).toFixed(3)}%;--c:${r.n.col};--d:${k * 70}ms" title="${esc(r.n.label)} — ${nf(r.n.v)} дана${r.n.em ? ' · ' + esc(r.n.em) : ''}">
        <div><b>${cu(r.n.v)}</b><span>${esc(r.n.label)}</span>${r.n.em ? `<small>${esc(r.n.em)}</small>` : ''}</div></div>`; }).join('')}</div>`
      + (() => { const tiny = out.filter(r => r.w * r.h / W < .09); return tiny.length ? `<div class="tmleg">${tiny.map(r => `<span><i style="background:${r.n.col}"></i>${esc(r.n.label)} — <b>${nf(r.n.v)}</b></span>`).join('')}</div>` : ''; })();   // names of tiles too small to hold text
  }
  const tmAr = n => (window.innerWidth || 800) < 520 ? (n > 6 ? 1 : 1.3) : (n > 6 ? 2.2 : 2.8);
  // vertical columns with the value on top and the name below
  const columns = items => { const max = Math.max(1, ...items.map(x => x.v));
    return `<div class="cols">${items.map((x, i) => `<div class="colm" title="${esc(x.label)}: ${nf(x.v)}"><b>${cu(x.v)}</b><div class="cbar"><i style="--h:${Math.max(3, x.v / max * 100).toFixed(1)}%;--c:${x.col};--d:${i * 90}ms"></i></div><small>${esc(x.label)}</small></div>`).join('')}</div>`; };
  // vertical columns against the plan: grey track = plan, coloured part = sold
  const planCols = items => `<div class="cols pcols">${items.map((x, i) => { const pl = x.p.plan || 0, f = pl ? x.n / pl : 0;
    return `<div class="colm" title="${esc(x.p.name)}: ${nf(x.n)} / ${nf(pl)}"><b>${cu(x.n)}<span class="of"> / ${nf(pl)}</span></b><div class="cbar"><i style="--h:${Math.max(2, Math.min(1, f) * 100).toFixed(1)}%;--c:${pcol(f)};--d:${i * 90}ms"></i></div><small>${esc(x.p.name.replace(/^Математика\s*—\s*/, 'Мат. · ').replace(/^Ағылшын\s*—\s*/, 'Ағыл. · '))}</small><em style="color:${pcol(f)}">${pc0(x.n, pl)}</em></div>`; }).join('')}</div>`;
  // funnel: each step narrower by its share of the first one
  const funnel = steps => { const top = steps[0].v || 1;
    return `<div class="funnel">${steps.map((x, i) => { const w = Math.max(28, x.v / top * 100);
      return `<div class="fstep" style="--w:${w.toFixed(1)}%;--c:${x.col};--d:${i * 150}ms"><span class="fl">${esc(x.label)}</span><div class="fwrap"><div class="fbar"><b>${cu(x.v)}</b></div></div>${i ? `<em>${pc1(x.v, top)}</em>` : '<em>100 %</em>'}</div>`; }).join('')}</div>`; };
  const sec = (n, title, body) => `<section class="sec"><h2><span class="n">${n}</span>${esc(title)}</h2>${body}</section>`;
  const prog = (title, pct, frac, meta, col) => `<div class="pg"><div class="pg-top"><b>${esc(title)}</b><span class="pct" style="color:${col}">${pct}</span></div><div class="pg-meta">${meta}</div>
    <div class="pbar"><div style="--w:${Math.min(100, frac * 100).toFixed(1)}%;background:linear-gradient(90deg,color-mix(in srgb,${col} 55%,#fff),${col})"></div></div></div>`;
  // 14-day vector trend in the hero; click a point to open that day. Drawn at the real width (no stretching), monotone curve (no false dips)
  function monoPath(P) {
    const n = P.length, dx = [], m = [], t = []; if (n < 2) return '';
    for (let i = 0; i < n - 1; i++) { dx[i] = P[i + 1][0] - P[i][0]; m[i] = (P[i + 1][1] - P[i][1]) / dx[i]; }
    t[0] = m[0]; t[n - 1] = m[n - 2]; for (let i = 1; i < n - 1; i++) t[i] = m[i - 1] * m[i] <= 0 ? 0 : (m[i - 1] + m[i]) / 2;
    for (let i = 0; i < n - 1; i++) { if (!m[i]) { t[i] = 0; t[i + 1] = 0; continue; } const a = t[i] / m[i], b = t[i + 1] / m[i], q = a * a + b * b; if (q > 9) { const k = 3 / Math.sqrt(q); t[i] = k * a * m[i]; t[i + 1] = k * b * m[i]; } }
    let d = `M${P[0][0].toFixed(1)},${P[0][1].toFixed(1)}`;
    for (let i = 0; i < n - 1; i++) { const h = dx[i] / 3; d += ` C${(P[i][0] + h).toFixed(1)},${(P[i][1] + t[i] * h).toFixed(1)} ${(P[i + 1][0] - h).toFixed(1)},${(P[i + 1][1] - t[i + 1] * h).toFixed(1)} ${P[i + 1][0].toFixed(1)},${P[i + 1][1].toFixed(1)}`; }
    return d;
  }
  function spark(cur, boxW) {
    const C = counts(), N = 14, days = []; for (let i = N - 1; i >= 0; i--) days.push(utcToIso(isoToUTC(cur) - i * 864e5));
    const pts = days.map((d, i) => ({d, i, v: d < minD ? null : C[d] || 0})).filter(p => p.v != null); if (pts.length < 2) return '';
    const W = Math.max(240, Math.round((boxW || 300) - 24)), small = W < 420;   // width of the hero box, measured after it is on the page
    const H0 = 0, H = small ? 128 : 150, L = 4, Rp = 34, T = 22, B = 24;
    const top = Math.max(1, ...pts.map(p => p.v)), step = Math.pow(10, Math.floor(Math.log10(top))), max = Math.ceil(top / step * 2) / 2 * step;   // round scale: 119 → 150, 998 → 1 000
    const X = i => L + i * (W - L - Rp) / (N - 1), Y = v => T + (H - T - B) * (1 - v / max), base = H - B;
    const P = pts.map(p => [X(p.i), Y(p.v)]), dl = monoPath(P), area = dl + ` L${P[P.length - 1][0].toFixed(1)},${base} L${P[0][0].toFixed(1)},${base} Z`;
    const sel = pts.find(p => p.d === cur) || pts[pts.length - 1], sx = X(sel.i), sy = Y(sel.v);
    const grid = [0, .5, 1].map(f => { const y = Y(max * f).toFixed(1); return `<line class="sp-grid" x1="${L}" x2="${W - Rp + 6}" y1="${y}" y2="${y}"/><text class="sp-y" x="${W}" y="${(+y + 3.5).toFixed(1)}" text-anchor="end">${nf(max * f)}</text>`; }).join('');
    const xl = pts.map(p => { const wkd = (new Date(isoToUTC(p.d)).getUTCDay() + 6) % 7 >= 5, show = !small || p.i % 2 === (N - 1) % 2 || p.d === cur;
      return show ? `<text class="sp-x${p.d === cur ? ' on' : ''}${wkd ? ' we' : ''}" x="${X(p.i).toFixed(1)}" y="${H - 6}" text-anchor="middle">${p.d.slice(8, 10)}</text>` : ''; }).join('');
    const dots = pts.map(p => `<g class="sp-pt${p.d === cur ? ' on' : ''}"><circle class="sp-hit" cx="${X(p.i).toFixed(1)}" cy="${Y(p.v).toFixed(1)}" r="13" data-go="${p.d}" tabindex="0" role="button" aria-label="${dmy(p.d)}: ${nf(p.v)}"><title>${dmy(p.d)}, ${wd(p.d)} · ${nf(p.v)}</title></circle><circle class="sp-dot${p.d === cur ? ' on' : ''}" cx="${X(p.i).toFixed(1)}" cy="${Y(p.v).toFixed(1)}" r="${p.d === cur ? 5 : 3}"/>
      <text class="sp-tip" x="${Math.min(W - Rp - 8, Math.max(14, X(p.i))).toFixed(1)}" y="${Math.max(12, Y(p.v) - 12).toFixed(1)}" text-anchor="middle">${nf(p.v)}</text></g>`).join('');
    const unit = kind === 'leads' ? 'жаңа лид' : 'сатылым', avg = pts.reduce((t, p) => t + p.v, 0) / pts.length;
    return `<div class="spark"><div class="spark-h"><span>Соңғы 14 күн · ${unit} · орташа ${nf(avg)}</span><span>${dmy(sel.d).slice(0, 5)}: <b>${nf(sel.v)}</b></span></div>
      <svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="Соңғы 14 күн, ${unit}"><defs><linearGradient id="spStroke" x1="0" x2="1"><stop offset="0" stop-color="#748ffc"/><stop offset="1" stop-color="#96f2d7"/></linearGradient>
      <linearGradient id="spFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#748ffc" stop-opacity=".42"/><stop offset="1" stop-color="#748ffc" stop-opacity="0"/></linearGradient></defs>
      ${grid}<line class="sp-sel" x1="${sx.toFixed(1)}" x2="${sx.toFixed(1)}" y1="${T - 6}" y2="${base}"/><path class="sp-area" d="${area}"/><path class="sp-line" d="${dl}" pathLength="1"/>
      <circle class="sp-pulse" cx="${sx.toFixed(1)}" cy="${sy.toFixed(1)}" r="5"/>${dots}${xl}</svg></div>`;
  }
  const LOGO = `<svg viewBox="0 0 32 32" aria-hidden="true"><defs><linearGradient id="lgG" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#748ffc"/><stop offset="1" stop-color="#7048e8"/></linearGradient></defs><rect width="32" height="32" rx="10" fill="url(#lgG)"/><path d="M8 21l5-6 4 3 6-8" stroke="#fff" stroke-width="2.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/><circle cx="23" cy="10" r="2.2" fill="#96f2d7"/></svg>`;
  // reveal sections on scroll; numbers count up as their section appears
  function animateIn(go) {
    const reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches, secs = [...document.querySelectorAll('.sec')];
    const count = root => { if (reduce) return; root.querySelectorAll('.cu').forEach(el => { const to = +el.dataset.n; if (!to || el.dataset.done) return; el.dataset.done = 1; const t0 = performance.now(), dur = 1200;
      const step = t => { const k = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - k, 3); el.textContent = nf(to * e); if (k < 1) requestAnimationFrame(step); }; el.textContent = '0'; requestAnimationFrame(step); setTimeout(() => { el.textContent = nf(to); }, dur + 400); }); };   // final value even if the tab is in the background (no animation frames)
    const show = el => { el.classList.add('in'); count(el); };
    count(document.querySelector('.hero'));
    if (reduce || !('IntersectionObserver' in window)) secs.forEach(show);
    else { const io = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) { show(e.target); io.unobserve(e.target); } }), {threshold: .12}); secs.forEach(x => io.observe(x)); setTimeout(() => secs.forEach(x => x.classList.contains('in') || show(x)), 4000); }
    document.querySelectorAll('[data-go]').forEach(el => { const f = () => { if (el.dataset.go !== day) go(el.dataset.go); }; el.onclick = f; el.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); f(); } }; });
  }

  function render() {
    const T = kind === 'leads' ? 'Лидтер бойынша күнделікті есеп' : 'Сатылым · күнделікті есеп';
    let tiles = '', body = '', text = '', textFull = '';
    if (kind === 'leads') {
      const R = leads(day), lnk = cfg.url ? cfg.url + '?card=leads&d=' + day : canSwitch() ? location.origin + otherHref('leads', day) : ''; text = leadsText(day, R, lnk, true); textFull = leadsText(day, R, lnk);
      tiles = `<div class="tile"><small>Кіріс өтінімдер</small><b>${cu(R.newIn)}</b><i>CRM-да барлығы ${nf(R.newL)} сделка</i></div><div class="tile" style="--d:80ms"><small>Рассылкадан «ОП»-ға</small><b>${R.mail ? pc1(R.mailOp, R.mail).replace(' %', '') + '<small>%</small>' : '—'}</b><i>${nf(R.mailOp)} / ${nf(R.mail)}</i></div><div class="tile" style="--d:160ms"><small>Сәтті жабылған</small><b>${cu(R.won)}</b><i>${R.fill ? nf(R.fill.sum) + ' ₸' : 'сәтсіз ' + nf(R.lost)}</i></div>`;
      const org = R.g('WhatsApp') + R.g('Instagram');
      const srcRows = [{label: 'Рассылка', v: R.g('Рассылка'), col: SRC_COL['Рассылка']}, {label: 'Instagram / Facebook таргеті', v: R.g('Таргет'), col: SRC_COL['Таргет']},
        {label: 'Органика', v: org, col: 'var(--teal)', subs: [{label: 'WhatsApp', v: R.g('WhatsApp'), col: SRC_COL['WhatsApp'], sub: 1}, {label: 'Instagram', v: R.g('Instagram'), col: SRC_COL['Instagram'], sub: 1}].filter(x => x.v)},
        {label: 'Ұсыным арқылы', v: R.g('Сарафан'), col: SRC_COL['Сарафан']}, {label: 'Өздігінен хабарласқан', v: R.g('Самост. обращение'), col: SRC_COL['Самост. обращение']},
        {label: 'Дереккөзі көрсетілмеген', v: R.g('Не указан'), col: SRC_COL['Не указан']}].concat(Object.keys(R.src).filter(k => !SRC_TXT[k]).map(k => ({label: k, v: R.src[k], col: 'var(--sl)'})))
        .filter(x => x.v).sort((a, b) => b.v - a.v).flatMap(x => [x].concat(x.subs || [])).map(x => Object.assign({}, x, {em: x.sub ? '' : R.newL ? pc0(x.v, R.newL) : ''}));
      const srcSeg = Object.entries(R.src).filter(e => e[1]).sort((a, b) => b[1] - a[1]).map(([k, v]) => ({label: (SRC_TXT[k] || [, k])[1], v, col: SRC_COL[k] || 'var(--sl)'}));
      body += sec(1, 'Жаңа лидтер қайдан келді?', srcSeg.length ? `<div class="big"><b>${cu(R.newL)}</b><span>жаңа лид</span></div>` + treemap(srcSeg.map(x => Object.assign({}, x, {col: ({'var(--ind)': '#4263eb', 'var(--amb)': '#f08c00', 'var(--teal)': '#0ca678', 'var(--vio)': '#7048e8', 'var(--sl)': '#868e96'})[x.col] || x.col, em: pc0(x.v, R.newL)})), tmAr(srcSeg.length)) : '<div class="empty">Бұл күні жаңа лид жоқ</div>');
      const segM = [{label: 'Жауап беріп, «ОП»-ға ауысқан', v: R.mailOp, col: 'var(--teal)', pct: pc1(R.mailOp, R.mail)}, {label: '«Рассылка» воронкасында қалған', v: R.mail - R.mailOp, col: '#c5d0e6'}];
      body += sec(2, 'Жаңа рассылка лидтері', !R.mail ? `<div class="big"><b>0</b><span>лид рассылкадан келді</span></div><div class="empty">Бұл күні рассылкадан жаңа лид түспеді — «ОП»-ға ауысу үлесі есептелмейді.</div>` : funnel([{label: 'Рассылкадан келді', v: R.mail, col: '#4263eb'}, {label: 'Жауап беріп, «ОП»-ға өтті', v: R.mailOp, col: '#0ca678'}])
        + `<div class="minis"><div class="mini"><i>⏳</i><div><b>${cu(R.mail - R.mailOp)}</b><small>әлі «Рассылка» воронкасында (жауап жоқ)</small></div></div></div>
        <p class="note">«ОП»-ға ауысқандар — сделканың қазіргі воронкасы бойынша (деректер ${snapTxt} жаңартылған).</p>`);
      const lostItems = nzSrc(R.lostSrc).map(k => ({label: (SRC_TXT[k] || [, k])[1], v: R.lostSrc[k], col: SRC_COL[k] || 'var(--sl)'}));
      const mlItems = nzSrc(R.mailLostSrc).map(k => ({label: k === 'Не указан' ? 'Дереккөзі көрсетілмеген' : 'Дереккөзі «' + k + '»', v: R.mailLostSrc[k], col: SRC_COL[k] || 'var(--sl)', sub: 1}));
      const wl = [{label: 'Сәтті жабылған', v: R.won, col: 'var(--grn)'}, {label: 'Сәтсіз жабылған', v: R.lost, col: 'var(--red)'}];
      body += sec(3, 'Күн ішінде жабылған лидтер', `<div class="donut">${donut(wl, cu(R.won + R.lost), 'жабылды')}<div>${legend(wl, R.won + R.lost)}${R.wonS ? `<div class="kpi good" style="margin-top:12px"><small>Сәтті сделкалар сомасы</small><b>${cu(R.wonS)} <span style="font-size:15px;color:var(--ink3)">₸</span></b></div>` : ''}</div></div>`
        + (lostItems.length ? `<div class="subh">Сәтсіз жабылғандар — қай дереккөзден</div>${columns(lostItems.sort((a, b) => b.v - a.v))}` : '')
        + (R.mailLost ? `<p class="note">Оның ішінде рассылка воронкаларында — ${nf(R.mailLost)}.</p>` : ''));
      const W = waiting(day);
      if (W) {
        const byN = o => Object.entries(o).sort((a, b) => b[1] - a[1]).map(([k, v]) => ({n: k, v: nf(v)})), sumB = b => b.reduce((x, y) => x + y, 0);
        const seg = (b, nm, e) => b.map((v, i) => v ? `<i style="flex:${v} 0 0;background:${DBC[i]}"${ct({h: manShort(nm) + ' · ' + DBK[i], big: nf(v) + ' лид', c: DBC[i], rows: byN(e.bs[i]), note: Math.round(v / (e.n || 1) * 100) + '% — оның барлық күтіп тұрғандарынан'})}></i>` : '').join('');
        const manTip = (nm, e) => ct({h: manShort(nm), big: nf(e.n) + ' лид күтіп тұр', rows: DBK.map((k, i) => ({n: k, v: nf(e.b[i]), c: DBC[i]})).filter((x, i) => e.b[i]).concat(byN(e.st)), note: e.stuck ? 'ОП-да 3+ күн қозғалыссыз: ' + nf(e.stuck) + (Object.keys(e.stuckSt).length ? ' (' + Object.entries(e.stuckSt).map(([k, v]) => k + ' ' + nf(v)).join(', ') + ')' : '') : (e.n ? 'ең ұзағы ' + dTxt(e.max) : '')});
        const stTip = k => ct({h: k, big: nf(W.st[k]) + ' лид', rows: Object.entries(W.stm[k] || {}).sort((a, b) => sumB(b[1]) - sumB(a[1])).map(([nm, b]) => ({n: manShort(nm), v: nf(sumB(b)) + (b[2] + b[3] ? ' · 3+ күн ' + nf(b[2] + b[3]) : '')})), note: 'менеджерлер бойынша'});
        const mx = Math.max(1, ...W.list.map(([, e]) => e.n));
        const pastNote = day >= snapIso ? '' : W.past ? `<div class="snapnote ok">📸 ${dmy(day)} күнгі срез · ${W.at} жағдайы</div>` : W.loading ? `<div class="snapnote">⏳ ${dmy(day)} күнгі срез жүктелуде…</div>` : `<div class="snapnote">⏱ ${dmy(day)} күнгі срез жоқ — ${snapTxt} жағдайы (қазіргі күй). Срездер «Срездер» парағына күн сайын жазылады.</div>`;   /* иесі 15.10 */   /* иесі 15.10: «статичная» — түсіндіру */
        body += sec(4, 'Жауап күтіп тұрған лидтер', pastNote + (!W.n && !W.stuck ? `<div class="fillhead ok"><b>✓</b><span>Жауапсыз лид жоқ</span></div>` :
          `<div class="big"><b>${cu(W.n)}</b><span>лид жауап күтіп тұр · ${W.at} жағдайы</span></div>`
          + `<div class="wlegend">${DBK.map((k, i) => `<span${W.b[i] ? ct({h: k, big: nf(W.b[i]) + ' лид', c: DBC[i], rows: byN(W.bst[i])}) : ''}><i style="background:${DBC[i]}"></i>${esc(k)} <b>${nf(W.b[i])}</b></span>`).join('')}</div>`
          + `<div class="wrows">${W.list.map(([nm, e]) => `<div class="wrow"><div class="wname"${manTip(nm, e)}><b>${esc(manShort(nm))}</b><small>${e.n ? 'ең ұзағы ' + esc(dTxt(e.max)) : ''}${e.stuck ? (e.n ? ' · ' : '') + 'ОП-да 3+ күн қозғалыссыз: ' + nf(e.stuck) : ''}</small></div>`
            + `<div class="wtrack"><div class="wbar" style="width:${(e.n / mx * 100).toFixed(1)}%">${seg(e.b, nm, e)}</div></div><b class="wn">${nf(e.n)}</b></div>`).join('')}</div>`
          + `<div class="minis">${Object.entries(W.st).sort((a, b) => b[1] - a[1]).map(([k, v]) => `<div class="mini"${stTip(k)}><i>📥</i><div><b>${cu(v)}</b><small>${esc(k)}</small></div></div>`).join('')}</div>`
          + `<div class="howto">🔎 <b>Қалай табу:</b> amoCRM → Сделки → этап «НОВАЯ ЗАЯВКА» (рассылкада — «Жауап берді») + «Ответственный: өзіңіз» → «Изменена» бойынша сұрыптаңыз, ең ескісінен бастаңыз.</div>`
          + `<p class="note">Күн — amoCRM-дағы соңғы өзгерістен бері («Изменена»). Бұл — қазіргі жағдай, таңдалған күнге байланысты емес.</p>`));
      }
      const sItems = nzSrc(R.sNewSrc).map(k => ({label: (SRC_TXT[k] || [, k])[1], v: R.sNewSrc[k], col: SRC_COL[k] || 'var(--sl)'}));
      body += sec(W ? 5 : 4, 'Техникалық қолдау', `<div class="minis"><div class="mini"><i>📥</i><div><b>${cu(R.newS)}</b><small>ашылған өтініш</small></div></div><div class="mini"><i>✅</i><div><b>${cu(R.sWon)}</b><small>сәтті жабылған</small></div></div><div class="mini"><i>❌</i><div><b>${cu(R.sLost)}</b><small>сәтсіз жабылған</small></div></div></div>`);
      if (R.fill) { const F = R.fill, gap = R.sheet ? R.sheet[0] - F.won : null, miss = F.noB + F.noP + F.noSrc + F.noR, d5 = dmy(day).slice(0, 5);
        // one fixed layout per manager: the 4 fields a closed deal must have — ✓ filled, «3 / 5» empty in 3 of 5 deals, «—» no such deals that day
        const FIELDS = [['💰', 'Сома («Бюджет»)', 'noB', 'won', 'сәтті сделкада бос', 'сәтті сделка жоқ'], ['📦', 'Өнім («Продукт»)', 'noP', 'won', 'сәтті сделкада бос', 'сәтті сделка жоқ'],
          ['📍', '«Қайдан келді»', 'noSrc', 'won', 'сәтті сделкада бос', 'сәтті сделка жоқ'], ['❓', 'Бас тарту себебі', 'noR', 'lost', 'бас тартуда бос', 'бас тарту жоқ']];
        const cell = (e, [ic, lb, k, base, badTxt, naTxt]) => !e[base] ? `<div class="fcell na"><small>${ic} ${esc(lb)}</small><b>—</b><em>${naTxt}</em></div>`
          : e[k] ? `<div class="fcell bad"><small>${ic} ${esc(lb)}</small><b>${nf(e[k])} / ${nf(e[base])}</b><em>${badTxt}</em></div>` : `<div class="fcell ok"><small>${ic} ${esc(lb)}</small><b>✓</b><em>толық</em></div>`;
        const gaps = e => e.noB + e.noP + e.noSrc + e.noR;
        const men = Object.entries(F.man).filter(([, e]) => e.won || e.lost).sort((a, b) => gaps(b[1]) - gaps(a[1]) || (b[1].won + b[1].lost) - (a[1].won + a[1].lost));
        body += sec(W ? 6 : 5, 'ОП: CRM-да толтыру керек', `<p class="fintro">Сделка жабылғанда толтырылуы керек: <b>сәтті</b> сделкада — сома, өнім, «Қайдан келді»; <b>бас тартқанда</b> — себебі. Төменде — ${d5} жабылған сделкалар.</p>`
          + (men.length ? `<div class="fgrid">${men.map(([nm, e]) => `<div class="fcard${gaps(e) ? '' : ' done'}"><div class="fcard-h"><b>👤 ${esc(manShort(nm))}</b><small>${nf(e.won)} сәтті · ${nf(e.lost)} бас тарту</small>${gaps(e) ? '' : '<span class="fok">✓ бәрі толық</span>'}</div>
            <div class="fcells">${FIELDS.map(f => cell(e, f)).join('')}</div></div>`).join('')}</div>`
            + (miss ? `<div class="howto">🔎 <b>Қалай табу:</b> amoCRM → Сделки → фильтр «Закрыта: ${d5}» + «Ответственный: өзіңіз» → өрісі бос сделканы ашып, толтырыңыз.</div>` : '')
            : `<div class="empty">${d5} ОП-да жабылған сделка жоқ.</div>`)
          + (gap > 0 ? `<div class="fpay"><div class="fpay-g"><div class="fsrc"><small><i style="background:#7048e8"></i>«ОП миникурс» кестесі</small><b>${nf(R.sheet[0])}</b><span>төлем · ${nf(R.sheet[1])} ₸</span></div>`
            + `<div class="fsrc"><small><i style="background:#2b8a3e"></i>amoCRM деректері</small><b>${nf(F.won)}</b><span>«Успешно» сделка · ${nf(F.sum)} ₸</span></div></div>`
            + `<div class="fpay-w"><span>⚠️</span><p><b>${nf(gap)} төлем</b> CRM-да әлі жабылмаған — ОП ${d5} төлегендердің сделкасын тауып, «Успешно» деп жабуы керек.</p></div></div>` : '')); }
    } else {
      const R = sales(day), M = R.main, lnk = cfg.url ? cfg.url + '?card=sales&d=' + day : canSwitch() ? location.origin + otherHref('sales', day) : ''; text = salesText(day, R, lnk, true); textFull = salesText(day, R, lnk);
      const COLS = ['#4263eb', '#7048e8', '#0ca678', '#f08c00', '#1098ad', '#ae3ec9', '#e8590c', '#5c7cfa'];
      const gapW = platGap(day, R);
      tiles = `<div class="tile"><small>Тіркелген төлемдер</small><b class="m">${cu(R.sum)}&nbsp;<small>₸</small></b><i>${nf(R.rows)} дана${gapW ? ` · ⚠️ платформа ${nf(gapW.n)} (әдетте ~${nf(gapW.med)}) — толық емес болуы мүмкін` : ''}</i></div>` + R.secsVis.filter(s => s.planN).slice(0, 2).map((s, i) => `<div class="tile" style="--d:${(i + 1) * 80}ms"><small>${esc(s.name)}</small><b>${cu(Math.round(s.n / s.planN * 100))}<small>%</small></b><i>${nf(s.n)} / ${nf(s.planN)} дана</i></div>`).join('');
      const planned = R.secsVis.filter(s => s.planN), free = R.secsVis.filter(s => !s.planN);
      body += sec('🎯', 'Орындалуы', `<div class="gauges">${planned.map((s, i) => gauge(s.name, s.n, s.planN, `${nf(s.n)} / ${nf(s.planN)} дана · ${nf(s.s)} ₸`, {d: i * 150, under: 'жоспардан'})).join('')}</div>`
        + (free.length ? `<div class="minis">${free.map(s => `<div class="mini"><i>${secEmo(s)}</i><div><b>${cu(s.n)} <span style="font-size:13px;color:var(--ink3);font-weight:700">дана</span></b><small>${esc(s.name)} · ${nf(s.s)} ₸</small></div></div>`).join('')}</div>` : ''));
      body += sec(1, 'Сатылым — жоспар / факт', R.secsVis.map((s, si) => {
        const ac = /негізгі/i.test(s.name) ? '#3b5bdb' : /пробный/i.test(s.name) ? '#0d9488' : /түлек/i.test(s.name) ? '#7048e8' : /алтын/i.test(s.name) ? '#e8590c' : s.name === OTHER ? '#64748b' : '#1098ad', f = s.planN ? s.n / s.planN : null;
        const head = `<div class="grp-h"><i>${secEmo(s)}</i><div class="grp-t"><b>${esc(s.name)}</b><small>${s.planN ? `${nf(s.n)} / ${nf(s.planN)} дана` : `${nf(s.n)} дана`} · ${nf(s.s)} ₸</small></div>${f != null ? `<span class="grp-p" style="color:${pcol(f)}">${Math.round(f * 100)}%</span>` : ''}</div>`
          + (f != null ? `<div class="pbar grp-bar"><div style="--w:${Math.min(100, f * 100).toFixed(1)}%;background:linear-gradient(90deg,color-mix(in srgb,${ac} 55%,#fff),${ac})"></div></div>` : '');
        const open = `<div class="grp" style="--gc:${ac};--i:${si}">`;
        if (s.itemPlans) return open + head + planCols(s.vis) + '</div>';
        if (false) return open + head + `<table class="tbl"><thead><tr><th>Өнім</th><th>Жоспар</th><th>Факт</th><th>%</th></tr></thead><tbody>${s.vis.map(x => { const pl = x.p.plan || 0, f = pl ? x.n / pl : 0;
          const sub = Object.entries(x.raw || {}).filter(([, r]) => r.n).sort((a, b) => b[1].n - a[1].n).filter(e => lc(e[0]) !== lc(x.p.name)).slice(0, 4);
          return `<tr><td>${esc(x.p.name)}<small>${esc(priceTxt(x))}</small>${sub.map(([nm, r]) => `<small class="rawp">◦ ${esc(nm)} — ${nf(r.n)}</small>`).join('')}</td><td>${nf(pl)}<small>${nf(pl * (x.p.price || x.p.pprice || 0))} ₸</small></td><td>${nf(x.n)}<small>${nf(x.s)} ₸</small></td><td><span class="chip" style="background:color-mix(in srgb,${pcol(f)} 14%,transparent);color:${pcol(f)}">${pc0(x.n, pl)}</span></td></tr>`; }).join('')}</tbody></table></div>`;
        // treemap of what was sold: a «basket» line is opened into its real products
        const tiles = s.vis.filter(x => x.n).flatMap((x, i) => { const col = s.name === OTHER ? '#64748b' : COLS[i % COLS.length], subs = subsOf(x, s);
          return subs.length ? subs.map(([nm, r]) => ({label: nm, v: r.n, em: nf(r.s) + ' ₸', col})) : [{label: itemLbl(s, x), v: x.n, em: nf(x.s) + ' ₸' + (priceTxt(x) && !/,/.test(priceTxt(x)) ? ' · ' + priceTxt(x) : ''), col}]; });
        const none = s.vis.filter(x => !x.n);
        const cards = t => `<div class="pcards">${t.map(x => `<div class="pcard" style="--c:${x.col}"><b>${cu(x.v)}</b><span>${esc(x.label)}</span><small>${esc(x.em || '')}</small></div>`).join('')}</div>`;
        return open + head + (!tiles.length ? '<div class="empty">Сатылым болмады</div>' : tiles.length < 3 ? cards(tiles) : treemap(tiles, tmAr(tiles.length)))
          + (none.length ? `<div class="zero-chips">Сатылмады: ${none.map(x => `<span>${esc(itemLbl(s, x))}</span>`).join('')}</div>` : '') + '</div>';
      }).join('') + `<div class="note">Тіркелген төлемдер: <b>${nf(R.rows)} жазба — ${nf(R.sum)} ₸</b> (CRM «Успешно», бюджет: ${nf(R.by.s.n)} · ${nf(R.by.s.s)} ₸; платформа: ${nf(R.by.p.n)} · ${nf(R.by.p.s)} ₸).${R.refPrice ? ' Соманың орнына прайс бағасы алынған жазбалар: ' + nf(R.refPrice) + '.' : ''}${Object.keys(R.skipped).length ? ' Есепке кірмеген платформа жазбалары: ' + Object.entries(R.skipped).map(([k, v]) => esc(k) + ' — ' + v + ' (' + nf(R.skippedS[k] || 0) + ' ₸)').join(', ') + '.' : ''}${R.skipped['invoice жоқ'] === R.by.s.n && Math.round(R.skippedS['invoice жоқ'] || 0) === Math.round(R.by.s.s) && R.by.s.n ? ' «Invoice жоқ» жазбалардың саны мен сомасы ОП кестесі-мен бірдей — бұл сол төлемдердің платформадағы көшірмесі болуы мүмкін, сондықтан қосылмайды.' : ''}</div>`);
      if (M) { const fm = M.planN ? M.n / M.planN : 0;
        body += sec(2, 'Жалпы нәтиже · ' + M.name.toLowerCase(), `<div class="duo"><div class="kpi"><small>Дана · жоспар / факт</small><b>${nf(M.n)} <span style="color:var(--ink3);font-size:16px">/ ${nf(M.planN)}</span></b></div>
        <div class="kpi"><small>Сома · жоспар / факт</small><b>${nf(M.s)} <span style="color:var(--ink3);font-size:16px">/ ${nf(M.planS)} ₸</span></b></div></div>${prog('Орындалуы', pc0(M.n, M.planN), fm, 'дана бойынша', pcol(fm))}`); }
      const srcItems = []; R.srcOrder.forEach(l => { srcItems.push({label: l, v: R.src[l] || 0, col: /платформа/i.test(l) ? 'var(--vio)' : /таргет/i.test(l) ? 'var(--amb)' : /органика/i.test(l) ? 'var(--teal)' : /рассылка/i.test(l) ? 'var(--ind)' : 'var(--sl)'});
        if (l === R.platLine) srcItems.push({label: '1 000 ₸ және одан арзан', em: nf(R.plat.loS) + ' ₸', v: R.plat.lo, col: 'var(--vio)', sub: 1}, {label: '1 000 ₸-ден қымбат', em: nf(R.plat.hiS) + ' ₸', v: R.plat.hi, col: 'var(--vio)', sub: 1}); });
      const mainSrc = srcItems.filter(x => !x.sub);
      body += sec(3, 'Клиенттер қайдан келді?', `<div class="big"><b>${cu(R.rows)}</b><span>сатылым</span></div>` + columns(mainSrc.filter(x => x.v).sort((a, b) => b.v - a.v).map(x => Object.assign({}, x, {label: x.label + ' · ' + pc0(x.v, R.rows)})))
        + (canSwitch() ? `<a class="xlink" href="${otherHref('leads', day)}" target="_top">Осы күнгі лидтер мен CRM толтыру есебі →</a>` : '')
        + `<div class="minis"><div class="mini"><i>📱</i><div><b>${cu(R.plat.lo)}</b><small>платформа · 1 000 ₸ және арзан · ${nf(R.plat.loS)} ₸</small></div></div><div class="mini"><i>💎</i><div><b>${cu(R.plat.hi)}</b><small>платформа · 1 000 ₸-ден қымбат · ${nf(R.plat.hiS)} ₸</small></div></div></div>`);
      if (!fromCatalog) body += `<p class="note">Өнімдер әзірге «Жоспар» парағы бойынша — таблицада «Өнімдер» парағы әлі жоқ.</p>`;
      if ((D.errors || []).length) body += `<p class="note">⚠ ${esc(D.errors.join('; '))}</p>`;
      if (D.stale) body += `<p class="note">⚠ Платформа соңғы рет ${esc(String(D.stale.attemptAt || '').replace('T', ' '))} оқылмады — сатылым ${esc(String(D.stale.lastGoodAt || '').replace('T', ' '))} күйі бойынша көрсетілген.</p>`;
    }
    const prev = utcToIso(isoToUTC(day) - 864e5), next = utcToIso(isoToUTC(day) + 864e5);
    $('app').innerHTML = `<div class="wrap"><header class="hero"><i class="orb o1"></i><i class="orb o2"></i><i class="orb o3"></i><div class="gridbg"></div><div class="toprow"><div class="brand">${LOGO}EduApp</div>${canSwitch() ? `<nav class="switch" aria-label="Есептер"><a href="${otherHref('leads', day)}" target="_top"${kind === 'leads' ? ' class="on" aria-current="page"' : ''}>📥 Лидтер</a><a href="${otherHref('sales', day)}" target="_top"${kind === 'sales' ? ' class="on" aria-current="page"' : ''}>💰 Сатылым</a></nav>` : ''}</div><h1>${T}</h1>
      <div class="daterow"><button id="prev" ${prev < minD ? 'disabled' : ''} aria-label="Алдыңғы күн">‹</button><label id="dt-btn" role="button" tabindex="0" aria-haspopup="dialog">📅 ${dmy(day)}, ${wd(day)} <span style="opacity:.6;font-size:11px">▾</span></label><button id="next" ${next > maxD ? 'disabled' : ''} aria-label="Келесі күн">›</button></div>
      <div class="tiles">${tiles}</div><div class="spark" id="spark"></div></header>${body}
      <div class="acts"><button class="btn pri" id="cp-link">🔗 Сілтемені көшіру</button><button class="btn pdf" id="cp-pdf">📄 PDF жүктеу</button><button class="btn" id="cp-text">💬 WhatsApp мәтінін көшіру</button><button class="btn" id="cp-full" title="Менеджерлер, техқолдау, ескертпе — бәрі">📋 Толық мәтін</button></div><div id="cp-box"></div>
      <div class="foot">${kind === 'leads' ? 'Дереккөз: amoCRM → Google Sheets' : 'Дереккөз: amoCRM «Успешно» (бюджет), платформа, «Өнімдер» парағы'} · жаңартылды ${snapTxt}<br>Тек жиынтық сандар — аты-жөні, телефон, чек жоқ.</div></div><div class="toast" id="toast"></div>`;
    const go = d => { day = d; render(); try { google.script.history.replace(null, {card: kind, d: day}); } catch (e) { try { history.replaceState(null, '', location.pathname + qs(day)); } catch (e2) { /* local file */ } } };
    $('prev').onclick = () => go(prev); $('next').onclick = () => go(next);
    $('dt-btn').onclick = () => openCal($('dt-btn'), go);
    $('dt-btn').onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openCal($('dt-btn'), go); } };
    const link = cfg.url ? cfg.url + '?card=' + kind + '&d=' + day : location.href.split(/[?#]/)[0] + qs(day);
    $('cp-link').onclick = () => copy(link, 'Сілтеме көшірілді');
    $('cp-text').onclick = () => copy(text, 'Қысқа мәтін көшірілді — WhatsApp-қа қойыңыз');
    $('cp-full').onclick = () => copy(textFull || text, 'Толық мәтін көшірілді');
    $('cp-pdf').onclick = () => pdfDialog();
    // bottom bar: hides while scrolling down, comes back on scroll up and at the end of the page
    if (!render.sc) { render.sc = 1; let y0 = scrollY;
      addEventListener('scroll', () => { const bar = document.querySelector('.acts'); if (!bar) return;
        const y = scrollY, end = innerHeight + y >= document.documentElement.scrollHeight - 40;
        if (end || y < 60 || y < y0 - 4) bar.classList.remove('hid'); else if (y > y0 + 4) bar.classList.add('hid');
        y0 = y; }, {passive: true}); }
    const sp = $('spark'), spw = sp.parentElement.clientWidth - (innerWidth <= 480 ? 28 : 40); sp.outerHTML = spark(day, spw) || '';
    animateIn(go);
    if (!render.rs) { render.rs = 1; let w0 = innerWidth, tm; addEventListener('resize', () => { clearTimeout(tm); tm = setTimeout(() => { if (Math.abs(innerWidth - w0) < 30) return; w0 = innerWidth; const el = document.querySelector('.spark'); if (el) { el.outerHTML = spark(day, el.parentElement.clientWidth - (innerWidth <= 480 ? 28 : 40)); animateIn(go); } }, 200); }); }
  }
  function otherHref(k, d) {
    if (cfg.url) return cfg.url + '?card=' + k + '&d=' + d;
    const path = location.pathname;
    if (/\/(leads|sales)\/?$/i.test(path)) return path.replace(/(leads|sales)(\/?)$/i, k + '$2') + '?d=' + d;
    return path + '?card=' + k + '&d=' + d;
  }
  const canSwitch = () => !!cfg.url || /(\/(leads|sales)\/?|card\.html)$/i.test(location.pathname);   // a saved one-day file has only its own data
  // colour PDF: blocks are drawn as images exactly as on screen and laid out on A4 pages; a dialog chooses what to save
  const loadJs = src => new Promise((res, rej) => { if (document.querySelector('script[src="' + src + '"]')) return res(); const el = document.createElement('script'); el.src = src; el.onload = res; el.onerror = () => rej(new Error('кітапхана жүктелмеді')); document.head.appendChild(el); });
  const H2I = 'https://cdn.jsdelivr.net/npm/html-to-image@1.11.11/dist/html-to-image.js', JSPDF = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const blockEls = () => [document.querySelector('.hero'), ...document.querySelectorAll('.sec')];
  const blockName = (el, i) => i === 0 ? 'Басты көрсеткіштер және 14 күн графигі' : [...(el.querySelector('h2') || {childNodes: []}).childNodes].filter(n => !(n.classList && n.classList.contains('n'))).map(n => n.textContent).join('').trim();
  // images of the selected blocks of THIS page (also called from the other report opened in a hidden frame)
  window.__zopBlocks = async sel => {
    await loadJs(H2I);
    document.body.classList.add('pdfmode'); document.querySelectorAll('.sec').forEach(x => x.classList.add('in'));
    document.querySelectorAll('.cu').forEach(el => { el.dataset.done = 1; el.textContent = nf(+el.dataset.n); });
    await sleep(250);
    try { const out = [], els = blockEls();
      for (let i = 0; i < els.length; i++) { if (sel && sel.indexOf(i) < 0) continue; const el = els[i], bw = el.offsetWidth, bh = el.offsetHeight;
        const c = await htmlToImage.toCanvas(el, {width: bw, height: bh, pixelRatio: 2, skipFonts: true, backgroundColor: '#eef2fb', cacheBust: true});
        if (!bw || c.width < bw * 1.5) throw new Error('бетті алдыңғы терезеде ашып, қайта басыңыз');   // a hidden tab can measure blocks wrongly — never build a broken PDF
        out.push(c); }
      return out;
    } finally { document.body.classList.remove('pdfmode'); }
  };
  // the other report for the same day, opened invisibly on this site
  async function otherBlocks(k) {
    const f = document.createElement('iframe'); f.setAttribute('aria-hidden', 'true'); f.style.cssText = 'position:fixed;left:-12000px;top:0;width:' + Math.max(760, document.querySelector('.wrap').offsetWidth) + 'px;height:1600px;border:0';
    f.src = otherHref(k, day); document.body.appendChild(f);
    try { let w; for (let i = 0; i < 160; i++) { await sleep(250); try { w = f.contentWindow; if (w && w.__zopBlocks && w.document.querySelector('.sec')) break; } catch (e) { throw new Error('екінші есеп ашылмады'); } }
      if (!w || !w.__zopBlocks) throw new Error('екінші есеп жүктелмеді'); await sleep(700); return await w.__zopBlocks(null);
    } finally { f.remove(); }
  }
  function pdfFrom(groups) {   // groups: [{label, canvases}] — each group starts on a new page
    const {jsPDF} = window.jspdf, pdf = new jsPDF({unit: 'pt', format: 'a4', compress: true}), PW = 595.28, PH = 841.89, M = 22, CW = PW - 2 * M, TOP = M, BOT = PH - M - 16, labels = [];
    const bg = () => { pdf.setFillColor(238, 242, 251); pdf.rect(0, 0, PW, PH, 'F'); };
    let y = TOP; bg();
    groups.forEach((g, gi) => {
      if (gi) { pdf.addPage(); bg(); y = TOP; }
      g.canvases.forEach(c => { const k = CW / c.width; let sy = 0;
        while (sy < c.height) {
          const room = BOT - y, need = (c.height - sy) * k;
          if (need > room && y > TOP && (sy === 0 && need <= BOT - TOP || room < 140)) { pdf.addPage(); bg(); y = TOP; continue; }   // a block that fits on one page is not split
          const part = Math.min(c.height - sy, Math.floor(room / k)), cut = document.createElement('canvas'); cut.width = c.width; cut.height = part;
          cut.getContext('2d').drawImage(c, 0, sy, c.width, part, 0, 0, c.width, part);
          pdf.addImage(cut.toDataURL('image/jpeg', .92), 'JPEG', M, y, CW, part * k, undefined, 'FAST');
          labels[pdf.getNumberOfPages()] = g.label; sy += part; y += part * k + 10;
          if (sy < c.height) { pdf.addPage(); bg(); y = TOP; }
        } });
    });
    const n = pdf.getNumberOfPages();
    for (let i = 1; i <= n; i++) { pdf.setPage(i); pdf.setFontSize(8); pdf.setTextColor(138, 150, 173); pdf.text(`EduApp  ·  ${labels[i] || ''}  ·  ${dmy(day)}  ·  ${i} / ${n}`, PW / 2, PH - 12, {align: 'center'}); }
    return pdf;
  }
  const LBL = {leads: 'Leads', sales: 'Sales'}, FN = {leads: 'Lidter', sales: 'Satylym'}, KZ = {leads: 'Лидтер', sales: 'Сатылым'};
  async function makePdf(opt) {
    const btn = $('cp-pdf'), old = btn.innerHTML; btn.disabled = true; btn.innerHTML = '⏳ PDF дайындалуда…';
    try {
      await loadJs(JSPDF);
      const other = kind === 'leads' ? 'sales' : 'leads', mine = {label: LBL[kind], canvases: await window.__zopBlocks(opt.sel)};
      if (opt.mode === 'this') { pdfFrom([mine]).save(`EduApp_${FN[kind]}_${day}.pdf`); }
      else {
        btn.innerHTML = `⏳ ${KZ[other]} есебі жүктелуде…`;
        const theirs = {label: LBL[other], canvases: await otherBlocks(other)}, ordered = kind === 'leads' ? [mine, theirs] : [theirs, mine];
        if (opt.mode === 'both') pdfFrom(ordered).save(`EduApp_Lidter-Satylym_${day}.pdf`);
        else { pdfFrom([mine]).save(`EduApp_${FN[kind]}_${day}.pdf`); await sleep(600); pdfFrom([theirs]).save(`EduApp_${FN[other]}_${day}.pdf`); }
      }
      toast('PDF жүктелді');
    } catch (e) { toast('PDF жасалмады: ' + (e && e.message || e)); }
    finally { btn.disabled = false; btn.innerHTML = old; }
  }
  // dialog: which report(s) and which blocks
  function pdfDialog() {
    const old = $('pdfdlg'); if (old) old.remove();
    const els = blockEls(), two = canSwitch() && !cfg.url, other = kind === 'leads' ? 'sales' : 'leads';
    const box = document.createElement('div'); box.className = 'mdl'; box.id = 'pdfdlg'; box.setAttribute('role', 'dialog'); box.setAttribute('aria-label', 'PDF жүктеу');
    box.innerHTML = `<div class="mdl-c"><div class="mdl-h"><b>📄 PDF жүктеу</b><span>${dmy(day)}, ${wd(day)}</span></div>
      <div class="mdl-s">Қандай есеп?</div>
      <label class="opt"><input type="radio" name="pk" value="this" checked><span><b>Тек осы есеп — ${KZ[kind]}</b><small>бөлімдерін төменде таңдаңыз</small></span></label>
      ${two ? `<label class="opt"><input type="radio" name="pk" value="both"><span><b>Лидтер + Сатылым — бір PDF</b><small>екі есеп бір файлда, әрқайсысы жаңа беттен</small></span></label>
      <label class="opt"><input type="radio" name="pk" value="sep"><span><b>Лидтер мен Сатылым — бөлек</b><small>екі PDF файл</small></span></label>` : ''}
      <div class="mdl-s">${KZ[kind]} есебінен не сақтау керек? <a href="#" id="pdf-all">бәрі</a></div>
      <div class="chkl">${els.map((el, i) => `<label class="opt sm"><input type="checkbox" value="${i}" checked><span>${esc(blockName(el, i))}</span></label>`).join('')}</div>
      ${two ? `<p class="mdl-n">${KZ[other]} есебі толық сақталады.</p>` : ''}
      <div class="mdl-f"><button class="btn" id="pdf-x">Болдырмау</button><button class="btn pri" id="pdf-ok">📄 Жүктеу</button></div></div>`;
    document.body.appendChild(box); requestAnimationFrame(() => box.classList.add('on'));
    const close = () => { box.classList.remove('on'); setTimeout(() => box.remove(), 250); };
    box.onclick = e => { if (e.target === box) close(); };
    box.querySelector('#pdf-x').onclick = close;
    box.querySelector('#pdf-all').onclick = e => { e.preventDefault(); box.querySelectorAll('.chkl input').forEach(i => { i.checked = true; }); };
    box.querySelector('#pdf-ok').onclick = () => { const sel = [...box.querySelectorAll('.chkl input:checked')].map(i => +i.value), mode = (box.querySelector('input[name=pk]:checked') || {}).value || 'this';
      if (!sel.length) { toast('Кемінде бір бөлім таңдаңыз'); return; } close(); makePdf({mode, sel}); };
  }
  function qs(d) { return '?' + (/(leads|sales)\b/i.test(location.pathname) ? '' : 'card=' + kind + '&') + 'd=' + d; }
  const KK_MON = ['қаңтар', 'ақпан', 'наурыз', 'сәуір', 'мамыр', 'маусым', 'шілде', 'тамыз', 'қыркүйек', 'қазан', 'қараша', 'желтоқсан'];
  let dayCount = null;
  function counts() {   // leads created / sales made per day — shown under the dates
    if (dayCount) return dayCount; dayCount = {};
    if (kind === 'leads') { const SUP = D.dict.funnel.indexOf('Техподдержка'), b = isoToUTC(D.base); D.main.forEach(r => { if (r[1] === SUP || r[0] < 0) return; const k = utcToIso(b + r[0] * 864e5); dayCount[k] = (dayCount[k] || 0) + r[r.length - 2]; }); }
    else D.cube.forEach(r => { if (skipReason(r)) return; dayCount[r[1]] = (dayCount[r[1]] || 0) + r[r.length - 2]; });
    return dayCount;
  }
  function openCal(anchor, go) {
    closeCal(); let ym = day.slice(0, 7);
    const box = document.createElement('div'); box.className = 'cal'; box.id = 'cal'; box.setAttribute('role', 'dialog'); box.setAttribute('aria-label', 'Күнді таңдау');
    document.body.appendChild(box);
    const place = () => { const r = anchor.getBoundingClientRect(), w = box.offsetWidth; box.style.top = Math.min(r.bottom + 8, innerHeight - box.offsetHeight - 10) + 'px'; box.style.left = Math.max(10, Math.min(r.left, innerWidth - w - 10)) + 'px'; };
    const draw = () => {
      const y = +ym.slice(0, 4), m = +ym.slice(5, 7), first = Date.UTC(y, m - 1, 1), n = new Date(Date.UTC(y, m, 0)).getUTCDate(), lead = (new Date(first).getUTCDay() + 6) % 7, C = counts();
      const prevYm = utcToIso(Date.UTC(y, m - 2, 1)).slice(0, 7), nextYm = utcToIso(Date.UTC(y, m, 1)).slice(0, 7);
      let g = ''; for (let i = 0; i < lead; i++) g += '<span></span>';
      for (let d = 1; d <= n; d++) { const iso = `${ym}-${pad(d)}`, off = iso < minD || iso > maxD, wdI = (lead + d - 1) % 7, c = C[iso];
        g += `<button data-d="${iso}" class="${iso === day ? 'on' : ''}${iso === maxD ? ' now' : ''}${wdI >= 5 ? ' we' : ''}" ${off ? 'disabled' : ''} aria-label="${dmy(iso)}">${d}<small>${c ? nf(c) : '·'}</small></button>`; }
      box.innerHTML = `<div class="cal-h"><button data-m="${prevYm}" ${prevYm < minD.slice(0, 7) ? 'disabled' : ''} aria-label="Алдыңғы ай">‹</button><b>${KK_MON[m - 1].charAt(0).toUpperCase() + KK_MON[m - 1].slice(1)} ${y}</b><button data-m="${nextYm}" ${nextYm > maxD.slice(0, 7) ? 'disabled' : ''} aria-label="Келесі ай">›</button></div>
        <div class="cal-w">${['Дс', 'Сс', 'Ср', 'Бс', 'Жм', 'Сб', 'Жс'].map(x => `<span>${x}</span>`).join('')}</div><div class="cal-g">${g}</div>
        <div class="cal-f"><button data-d="${yesterday}" ${yesterday < minD || yesterday > maxD ? 'disabled' : ''}>Кеше</button><button data-d="${maxD}">Соңғы күн</button></div>`;
      place();
    };
    box.onclick = e => { const b = e.target.closest('button'); if (!b || b.disabled) return; if (b.dataset.m) { ym = b.dataset.m; draw(); } else if (b.dataset.d) { closeCal(); go(b.dataset.d); } };
    draw(); setTimeout(() => { const on = box.querySelector('.on') || box.querySelector('.cal-g button:not(:disabled)'); if (on) on.focus({preventScroll: true}); }, 0);
    closeCal.off = e => { if (e.type === 'keydown' ? e.key === 'Escape' : !box.contains(e.target) && !anchor.contains(e.target)) closeCal(); };
    setTimeout(() => { document.addEventListener('mousedown', closeCal.off); document.addEventListener('keydown', closeCal.off); addEventListener('resize', closeCal); addEventListener('scroll', closeCal, {passive: true}); }, 0);
  }
  function closeCal() { const b = document.getElementById('cal'); if (b) b.remove(); if (closeCal.off) { document.removeEventListener('mousedown', closeCal.off); document.removeEventListener('keydown', closeCal.off); removeEventListener('resize', closeCal); removeEventListener('scroll', closeCal); closeCal.off = null; } }
  function toast(m) { const t = $('toast'); t.textContent = m; t.classList.add('on'); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('on'), 2200); }
  function copy(s, msg) {
    const fallback = () => { const ta = document.createElement('textarea'); ta.value = s; ta.className = 'copy'; ta.readOnly = true; $('cp-box').innerHTML = ''; $('cp-box').appendChild(ta); ta.focus(); ta.select();
      let ok = false; try { ok = document.execCommand('copy'); } catch (e) { /* blocked */ } toast(ok ? msg : 'Көшіру үшін мәтінді белгілеп, Ctrl+C басыңыз'); };
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(s).then(() => toast(msg), fallback); else fallback();
  }
  render();
}
if (typeof module !== 'undefined') module.exports = {zopHm_, zopReplyCounts_, zopReplyDelta_, zopReplyCube_, zopRegion_, zopSnapOf_, zopSnapCell_, zopPayMatch_, zopMissRows_, zopBtnList_, zopMediaList_, zopTextClean_, zopImportTags_, zopCodeOk_, zopNoteClean_, zopNoteTarget_, zopColLetter_, zopMirrorFormula_, zopMirrorOk_, zopRequireOwner_, zopCheckToken_, zopSaveCost, zopSetup, zopDescribeSources, zopRebuildProducts, zopCosts_, zopCostAlloc_, zopTariffs_, zopTariffAt_, zopCostValidate_, zopHypotheses_, zopReconcileRows_, zopDescribeRows_, zopPickSheet_, zopSecFamilies_, zopSectionLists_, zopSecGuess_, zopPayStatus_, zopBudget_, ZOP_PAY, zopSkipReason_, zopAggregate_, zopMailings_, zopCosts_, zopSalesAggregate_, zopProductSeed_, zopProductList_, zopCardHtml_, zopCardClient_};
