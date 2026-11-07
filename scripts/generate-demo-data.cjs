// Demo data generator for EduApp Pulse.
//
// Produces RANDOM raw "Google Sheets" rows (amoCRM deals export, pipelines, sync log, mailings, costs, tariffs, KPI,
// hypotheses, platform payments) and runs them through the REAL aggregation code of the Apps Script
// (zopAggregate_ / zopSalesAggregate_ / zopProductSeed_ in инструменты/EduApp_OP_API.gs) — exactly what the live
// web app would return. Results:
//   • Атлас_отчёты.html      — embedded snapshots <script id="crm"> and <script id="sales"> (the dashboard works offline)
//   • demo/crm.json, demo/sales.json — the same snapshots for card.html in demo mode
//   • the regression tests read the same two files (tests/lib/load.cjs → fixture())
//
// Run:  node scripts/generate-demo-data.cjs [seed]
// Nothing here is real: names are random fake names, numbers come from a seeded PRNG.
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const ROOT = path.join(__dirname, '..');
const SEED = Number(process.argv[2] || 20261030);

// ---------- deterministic PRNG ----------
function mulberry32(a) { return () => { a = (a + 0x6d2b79f5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }
const R = mulberry32(SEED);
const rnd = (a, b) => a + R() * (b - a);
const int = (a, b) => Math.floor(rnd(a, b + 1));
const chance = p => R() < p;
const pick = arr => arr[Math.floor(R() * arr.length)];
const wpick = pairs => { const tot = pairs.reduce((s, p) => s + p[1], 0); let x = R() * tot; for (const [v, w] of pairs) { if ((x -= w) < 0) return v; } return pairs[pairs.length - 1][0]; };
const poisson = lam => { if (lam > 30) return Math.max(0, Math.round(lam + Math.sqrt(lam) * gauss())); let L = Math.exp(-lam), k = 0, p = 1; do { k++; p *= R(); } while (p > L); return k - 1; };
function gauss() { let u = 0, v = 0; while (!u) u = R(); while (!v) v = R(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); }

// ---------- time: local = UTC+5 (Asia/Qyzylorda), day index from BASE (2026-08-10) ----------
const OFF = 5 * 3600e3, BASE = Date.UTC(2026, 7, 10);
const SNAP_DAY = 81, SNAP_H = 18, SNAP_MIN = 40;                 // snapshot: 2026-10-30 18:40 local
const at = (day, h = 0, m = 0) => new Date(BASE + day * 864e5 + h * 3600e3 + m * 60e3 - OFF);   // local time → Date
const SNAP = at(SNAP_DAY, SNAP_H, SNAP_MIN);
const parts = d => { if (!(d instanceof Date) || isNaN(d.getTime())) return null; const x = new Date(d.getTime() + OFF);
  return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours(), min: x.getUTCMinutes(), wd: (x.getUTCDay() + 6) % 7}; };
const wdOf = day => (new Date(BASE + day * 864e5).getUTCDay() + 6) % 7;   // 0 = Monday
const pad = n => String(n).padStart(2, '0');
const ddmmyyyy = day => { const p = parts(at(day, 12)); return `${pad(p.d)}.${pad(p.m)}.${p.y}`; };
const clampSnap = d => (d > SNAP ? null : d);
// an hour of the day with a realistic profile: late morning and evening peaks
const hourOf = () => wpick([[8, 2], [9, 5], [10, 8], [11, 10], [12, 10], [13, 8], [14, 7], [15, 7], [16, 7], [17, 7], [18, 8], [19, 9], [20, 9], [21, 7], [22, 3], [23, 1]]);
const WEEKDAY = [1.12, 1.06, 1.0, 1.0, 0.94, 0.72, 0.62];
const trend = day => 0.75 + 0.55 * Math.min(1, day / SNAP_DAY) + 0.15 * Math.sin(day / 9);   // autumn olympiad season grows

// ---------- dictionaries (fake people, generic products) ----------
const MANAGERS = [['Мадина Асқарова', 1.15], ['Сапарова Айым Ерланқызы', 1.0], ['Ахметова Томирис Болатқызы', 0.9], ['Әбдірова Жансая Маратқызы', 0.8], ['Dinara', 1.05]];   // «Dinara» = amoCRM account of «Айгерім» (ZOP_MAN_ALIAS)
const SRC_IN = [['Органика WhatsApp', 30], ['Таргетированная реклама в Инстаграм или Фейсбук', 30], ['Органика Instagram', 14], ['Рекомендация (Сарафан)', 10], ['Самост. обращение', 5], ['', 11]];
const PRODUCTS = [['Алтын сақа 1 мектепшілік', 5600, 22], ['Алтын сақа 2 мектепшілік', 5600, 16], ['Алтын сақа 3 мектепшілік', 5600, 10], ['Алтын сақа аудандық Комбо пакет', 24900, 12],
  ['Мега қайталау', 13500, 14], ['Математика тоқсандық', 8400, 12], ['Ағылшын тоқсандық', 8400, 8], ['Алтын түлек', 9900, 6]];
const REGIONS = [['Астана', 16], ['Алматы', 15], ['Шымкент', 10], ['Қарағанды облысы', 7], ['Түркістан облысы', 7], ['Ақтөбе', 6], ['Атырау', 5], ['Павлодар облысы', 4], ['Маңғыстау облысы', 4],
  ['Жамбыл облысы', 4], ['Абай облысы', 3], ['Қостанай облысы', 3], ['Батыс Қазақстан облысы', 3], ['Қызылорда облысы', 3], ['Ұлытау облысы', 1], ['', 12]];
const REASON_IDS = [[1000101, 40], [1000102, 25], [1000103, 20], [1000104, 15]];
const SEGMENTS = ['3 сынып', '4 сынып', '5 сынып', '6 сынып', 'ата-ана', 'Алтын сақа базасы', 'қаңтар оқушылары'];
const STAGES = {
  'ОП': ['НОВАЯ ЗАЯВКА', 'Неразобранное', 'НЕДОЗВОН', 'ПРОВЕДЕННЫЕ КОНСУЛЬТАЦИИ', 'СЧЕТ ВЫСТАВЛЕН', 'рассылка', 'Успешно реализовано', 'Закрыто и не реализовано'],
  'Рассылка': ['Неразобранное', 'Первичный контакт', 'Жауап берді', 'Рассылка 3', 'Принимают решение', 'Успешно реализовано', 'Закрыто и не реализовано'],
  'Рассылка 5-6 класс Алтын сақа': ['Неразобранное', 'Первичный контакт', 'жауап берген', 'Принимают решение', 'Успешно реализовано', 'Закрыто и не реализовано'],
  'Техподдержка': ['НОВОЕ ОБРАЩЕНИЕ', 'Неразобранное', 'ПРИНЯТО В РАБОТУ', 'ОЖИДАЕМ ДАННЫЕ ОТ КЛИЕНТА', 'В РАБОТЕ', 'Начали обучение', 'возврат', 'Успешно реализовано', 'Закрыто и не реализовано']};
const P_AREA = [['Тренинги', 25], ['Челленджи', 18], ['Вход', 15], ['Покупка подписки', 14], ['Техническая проблема', 12], ['Профиль', 6], ['Другое', 5], ['', 5]];
const P_TYPE = [['Не открывается', 18], ['Ошибка оплаты', 14], ['Доступ к тренингу', 14], ['Не загружается', 12], ['Не начислилась награда', 8], ['Непонятный интерфейс', 8], ['Неверный формат ответа', 6], ['Вылетает приложение', 6], ['Другое', 8], ['', 6]];
const UTM_CAMP = ['AS_autumn_reach', 'AS_retarget', 'Mega_revision_leads'], UTM_CREATIVE = ['video_teacher_1', 'video_parent_review', 'carousel_prices', 'static_offer_2plus1', 'reels_olympiad_tips'];

// ---------- amoCRM «Сделки» ----------
const HEAD = ['ID', 'Воронка', 'Этап', 'Итог', 'Создана', 'Закрыта', 'Изменена', 'Ответственный', 'Қайдан келді', 'Причина отказа ID', 'Бюджет', 'Жеңілдік', 'Продукт', 'Теги',
  'Ближайшая задача', 'Язык обучения', 'Чек', 'Класс', 'Пакет', 'Клиент', 'Облыс', 'problem_area', 'problem_type', 'utm_source', 'utm_campaign', 'utm_content'];
const leads = [HEAD];
let nextId = 31000000;
const manPick = () => wpick(MANAGERS.map(m => [m, 1]));
function deal(o) {
  const r = Object.assign({ID: nextId += int(1, 9)}, o), row = HEAD.map(h => r[h] === undefined ? '' : r[h]);
  leads.push(row); return r;
}
// outcome of a commercial deal, depending on its age and the manager
function close(created, man, opts = {}) {
  const ageDays = (SNAP - created) / 864e5, mgr = man[1];
  const win = (opts.win || 0.17) * mgr, lose = opts.lose || 0.55;
  const settled = Math.min(1, ageDays / 6);   // young deals are mostly still open
  const x = R();
  if (x < win * settled) return 'Успешно';
  if (x < (win + lose) * settled) return 'Не реализовано';
  return 'В работе';
}
function closeAt(created, outcome) {
  if (outcome === 'В работе') return null;
  const d = new Date(created.getTime() + Math.min(9, -Math.log(1 - R()) * 1.6) * 864e5 + int(0, 5) * 3600e3);
  const p = parts(d); if (p.h < 9 || p.h > 21) d.setTime(d.getTime() + (p.h < 9 ? 9 - p.h : 24 - p.h + 10) * 3600e3);
  return clampSnap(d) || new Date(SNAP.getTime() - int(10, 300) * 60e3);
}
function commercialFields(created, outcome, closed, opts = {}) {
  const prod = wpick(PRODUCTS.map(p => [p, p[2]])), offer = wpick([['', 72], ['Предпродажа', 9], ['2+1 Акция', 8], [0.15, 6], [0.2, 5]]);
  let budget = '';
  if (outcome === 'Успешно') {
    budget = prod[1] * (chance(0.12) ? 2 : 1);
    if (typeof offer === 'number') budget = Math.round(budget * (1 - offer) / 10) * 10;
    if (offer === '2+1 Акция') budget = prod[1] * 2;
    if (chance(0.04)) budget = '';   // a won deal without a budget — shows up in «data quality»
  } else if (chance(0.25)) budget = prod[1];
  const changed = outcome === 'В работе' ? clampSnap(new Date(created.getTime() + rnd(0, Math.max(0.1, (SNAP - created) / 864e5)) * 864e5)) || SNAP : closed;
  const task = outcome === 'В работе' && chance(0.65) ? new Date(SNAP.getTime() + rnd(-3, 3) * 864e5) : '';
  return {
    'Изменена': changed, 'Ближайшая задача': task, 'Бюджет': budget, 'Жеңілдік': offer,
    'Продукт': outcome === 'Успешно' ? (chance(0.93) ? prod[0] : '') : (chance(0.35) ? prod[0] : ''),
    'Причина отказа ID': outcome === 'Не реализовано' && chance(opts.reasonRate || 0.8) ? wpick(REASON_IDS) : '',
    'Чек': outcome === 'Успешно' && chance(0.88) ? 'чек.pdf' : '',
    'Язык обучения': wpick([['Қазақша', 68], ['Русский', 22], ['', 10]]), 'Класс': wpick([[2, 5], [3, 22], [4, 24], [5, 18], [6, 14], [7, 6], [8, 4], ['', 7]]),
    'Пакет': wpick([['AltynSaqa', 40], ['Basic', 10], ['', 50]]), 'Клиент': wpick([['Жаңа клиент', 60], ['Ескі клиент', 25], ['', 15]]), 'Облыс': wpick(REGIONS)};
}
function stageFor(funnel, outcome) {
  if (outcome === 'Успешно') return 'Успешно реализовано';
  if (outcome === 'Не реализовано') return 'Закрыто и не реализовано';
  return funnel === 'ОП' ? wpick([['НОВАЯ ЗАЯВКА', 26], ['Неразобранное', 6], ['НЕДОЗВОН', 20], ['ПРОВЕДЕННЫЕ КОНСУЛЬТАЦИИ', 30], ['СЧЕТ ВЫСТАВЛЕН', 18]]) : 'Неразобранное';
}

// 1) inbound ОП leads (WhatsApp, target ads, Instagram, referrals)
for (let day = 0; day <= SNAP_DAY; day++) {
  const n = poisson(17 * trend(day) * WEEKDAY[wdOf(day)]);
  for (let i = 0; i < n; i++) {
    const h = hourOf(), created = at(day, h, int(0, 59)); if (created > SNAP) continue;
    const man = manPick(), src = wpick(SRC_IN), outcome = close(created, man), closed = closeAt(created, outcome);
    const target = src.startsWith('Таргет');
    deal(Object.assign({'Воронка': 'ОП', 'Этап': stageFor('ОП', outcome), 'Итог': outcome === 'В работе' ? '' : outcome, 'Создана': created, 'Закрыта': closed || '',
      'Ответственный': man[0], 'Қайдан келді': src, 'Теги': chance(0.05) ? pick(SEGMENTS) : '',
      'utm_source': target ? wpick([['instagram', 70], ['facebook', 30]]) : '', 'utm_campaign': target && chance(0.85) ? pick(UTM_CAMP) : '', 'utm_content': target && chance(0.8) ? pick(UTM_CREATIVE) : ''},
    commercialFields(created, outcome, closed)));
  }
}

// 2) WhatsApp mailing waves: deals are imported with a tag «импорт_ДДММГГГГ_ЧЧММ»; replies move to ОП
const WAVES = [[19, 'Рассылка', 'Алтын сақа: күзгі дайындық', 900], [26, 'Рассылка', 'Мега қайталау: тоқсан қорытындысы', 1100], [33, 'Рассылка 5-6 класс Алтын сақа', '5–6 сынып: мектепішілік кезең', 520],
  [40, 'Рассылка', 'Алтын сақа 2-жинақ', 1300], [47, 'Рассылка', '2+1 акциясы', 1450], [52, 'Рассылка 5-6 класс Алтын сақа', '5–6 сынып: аудандық кезең', 610],
  [58, 'Рассылка', 'Математика тоқсандық', 1200], [64, 'Рассылка', 'Ағылшын: Get Ready', 980], [70, 'Рассылка', 'Комбо пакет — соңғы 3 күн', 1500], [76, 'Рассылка 5-6 класс Алтын сақа', 'Олимпиада алдындағы апта', 700]];
const mailRows = [['Рассылка', 'Тема', 'Уақыты', 'Дата начала', 'Дата окончания', 'Воронка', 'Кому (сегмент)', 'Доп. расход, ₸ (кроме WhatsApp)', 'Отправлено, шт', 'Комментарий', 'Мәтін', 'Батырмалар', 'Формат', 'Сынып', 'Тіл тобы', 'Өнім', 'Мақсат']];
const costRows = [['Дата', 'Подписка', 'Детали', 'Сумма, ₸', 'Оплата', 'Рассылка (если знаете)', 'Категория', 'Поставщик', 'Кампания', 'Хабарлама саны', 'Тариф']];
WAVES.forEach(([day, funnel, topic, sentFull], wi) => {
  const sent = Math.round(sentFull * 0.55);
  const h = pick([10, 11, 12, 15, 16]), mi = pick([0, 15, 30, 45]), t0 = at(day, h, mi), p = parts(t0);
  const tag = `импорт_${pad(p.d)}${pad(p.m)}${p.y}_${pad(h)}${pad(mi)}`, seg = pick(SEGMENTS), replyRate = rnd(0.05, 0.13), toOp = rnd(0.45, 0.7);
  for (let i = 0; i < sent; i++) {
    const created = new Date(t0.getTime() + int(0, 40) * 60e3), man = manPick(), tags = [tag, chance(0.7) ? seg : pick(SEGMENTS)].join(', ');
    if (chance(replyRate)) {
      const replied = new Date(created.getTime() + Math.min(72, -Math.log(1 - R()) * 6) * 3600e3); if (replied > SNAP) continue;
      if (chance(toOp)) {   // a manager moved the reply to ОП
        const outcome = close(replied, man, {win: 0.24, lose: 0.5}), closed = closeAt(replied, outcome);
        deal(Object.assign({'Воронка': 'ОП', 'Этап': outcome === 'В работе' ? wpick([['рассылка', 45], ['ПРОВЕДЕННЫЕ КОНСУЛЬТАЦИИ', 30], ['СЧЕТ ВЫСТАВЛЕН', 25]]) : stageFor('ОП', outcome),
          'Итог': outcome === 'В работе' ? '' : outcome, 'Создана': created, 'Закрыта': closed || '', 'Ответственный': man[0], 'Қайдан келді': 'Рассылка', 'Теги': tags}, commercialFields(created, outcome, closed)));
      } else {
        const st = funnel === 'Рассылка' ? wpick([['Жауап берді', 50], ['Первичный контакт', 25], ['Принимают решение', 25]]) : wpick([['жауап берген', 55], ['Первичный контакт', 20], ['Принимают решение', 25]]);
        const outcome = chance(0.12) && (SNAP - replied) > 2 * 864e5 ? 'Успешно' : chance(0.3) ? 'Не реализовано' : 'В работе', closed = closeAt(replied, outcome);
        deal(Object.assign({'Воронка': funnel, 'Этап': outcome === 'В работе' ? st : stageFor(funnel, outcome), 'Итог': outcome === 'В работе' ? '' : outcome, 'Создана': created, 'Закрыта': closed || '',
          'Ответственный': man[0], 'Қайдан келді': chance(0.8) ? 'Рассылка' : '', 'Теги': tags}, commercialFields(created, outcome, closed, {reasonRate: 0.3})));
      }
    } else {   // no reply: stays in the mailing funnel, part of it is closed later in bulk
      const bulk = (SNAP - created) > 10 * 864e5 && chance(0.55), closed = bulk ? clampSnap(new Date(created.getTime() + int(7, 12) * 864e5)) : null;
      deal({'Воронка': funnel, 'Этап': closed ? 'Закрыто и не реализовано' : funnel === 'Рассылка' ? wpick([['Неразобранное', 70], ['Рассылка 3', 30]]) : 'Неразобранное', 'Итог': closed ? 'Не реализовано' : '',
        'Создана': created, 'Закрыта': closed || '', 'Изменена': closed || created, 'Ответственный': man[0], 'Қайдан келді': chance(0.6) ? 'Рассылка' : '', 'Теги': tags, 'Облыс': wpick(REGIONS),
        'Язык обучения': wpick([['Қазақша', 70], ['Русский', 20], ['', 10]])});
    }
  }
  mailRows.push([`Рассылка ${wi + 1}`, topic, `${pad(h)}:${pad(mi)}`, ddmmyyyy(day), ddmmyyyy(day + 2), funnel, seg, chance(0.3) ? int(5, 20) * 1000 : '', sent, chance(0.4) ? 'демо: тест мәтіні' : '',
    `Сәлеметсіз бе! ${topic} — EduApp платформасында. Толығырақ: төмендегі батырма.`, 'Толығырақ | https://eduapp.example.com/course\nМенеджерге жазу', pick(['мәтін + сурет', 'мәтін + видео', 'мәтін']),
    funnel.includes('5-6') ? '5–6' : '3–4', 'қазақ', pick(['Алтын сақа', 'Мега қайталау', 'Математика тоқсандық']), pick(['сатылым', 'жауап', 'вебинарға шақыру'])]);
  costRows.push([ddmmyyyy(day - 1), 'Рассылка', 'Subscription balance', int(3, 9) * 5000, pick(['Stripe', 'партнёр']), `Рассылка ${wi + 1}`, 'Wazzup top-up', 'Wazzup', '', '', '']);
  costRows.push([ddmmyyyy(day), 'Рассылка', 'Marketing template messages', '', '', `Рассылка ${wi + 1}`, 'WhatsApp template charges', 'Wazzup', '', sent, 'marketing']);
});

// 3) technical support (Техподдержка)
for (let day = 0; day <= SNAP_DAY; day++) {
  const n = poisson(4 * WEEKDAY[wdOf(day)] * (0.8 + 0.4 * day / SNAP_DAY));
  for (let i = 0; i < n; i++) {
    const created = at(day, hourOf(), int(0, 59)); if (created > SNAP) continue;
    const age = (SNAP - created) / 864e5, outcome = age > 2 ? wpick([['Успешно', 70], ['Не реализовано', 18], ['В работе', 12]]) : wpick([['Успешно', 30], ['В работе', 70]]);
    const closed = outcome === 'В работе' ? null : clampSnap(new Date(created.getTime() + rnd(0.1, 2) * 864e5)) || SNAP;
    deal({'Воронка': 'Техподдержка', 'Этап': outcome === 'В работе' ? wpick([['НОВОЕ ОБРАЩЕНИЕ', 30], ['ПРИНЯТО В РАБОТУ', 30], ['ОЖИДАЕМ ДАННЫЕ ОТ КЛИЕНТА', 20], ['В РАБОТЕ', 20]]) : stageFor('Техподдержка', outcome),
      'Итог': outcome === 'В работе' ? '' : outcome, 'Создана': created, 'Закрыта': closed || '', 'Изменена': closed || created, 'Ответственный': manPick()[0],
      'problem_area': wpick(P_AREA), 'problem_type': wpick(P_TYPE), 'Облыс': wpick(REGIONS)});
  }
}

// ---------- pipelines, sync log ----------
const pipes = [['ID воронки', 'Воронка', 'ID этапа', 'Тип', 'Этап', 'Порядок', 'Тип этапа']];
Object.entries(STAGES).forEach(([f, list], fi) => list.forEach((s, i) => pipes.push([100 + fi, f, 1000 + fi * 50 + i, '', s, i * 10, /Успешно/.test(s) ? 142 : /Закрыто/.test(s) ? 143 : 0])));
const logRows = [['Время', 'Режим', 'Сделок']];
for (let t = SNAP.getTime() - 3 * 864e5; t <= SNAP.getTime(); t += 30 * 60e3) logRows.push([new Date(t), (t - SNAP.getTime()) % (12 * 3600e3) === 0 ? 'Полная' : 'Изменения', int(0, 40)]);

// ---------- costs, tariffs, KPI, hypotheses ----------
const tariffRows = [['Поставщик', 'Хабарлама түрі', 'Басталуы', 'Аяқталуы', 'Бағасы', 'Валюта', 'Дереккөз (шот / бет)'],
  ['Wazzup', 'marketing', '01.08.2026', '30.09.2026', 22, 'KZT', 'демо тариф'], ['Wazzup', 'marketing', '01.10.2026', '', 26, 'KZT', 'демо тариф'], ['Wazzup', 'utility', '01.08.2026', '', 9, 'KZT', 'демо тариф']];
costRows.push([ddmmyyyy(0), 'WhatsApp', '1 channel on the "Max" plan for 1 year', 240000, 'партнёр', '', 'Wazzup subscription', 'Wazzup', '', '', '']);
for (const day of [22, 52]) costRows.push([ddmmyyyy(day), 'Рассылка', '1 channel on the "Waba" plan for 1 month', 32000, 'партнёр', '', 'Wazzup subscription', 'Wazzup', '', '', '']);
for (const day of [0, 22, 52]) costRows.push([ddmmyyyy(day), 'Sipuni', 'Sipuni monthly plan', 18000, 'Stripe', '', 'Sipuni subscription', 'Sipuni', '', '', '']);
for (let day = 3; day <= SNAP_DAY; day += 7) {
  const sum = int(12, 26) * 5000, a = int(45, 70);
  costRows.push([ddmmyyyy(day), 'Таргет', 'Meta ads weekly budget', sum, 'карта', '', 'Marketing campaign costs', 'Meta', `AS_autumn_reach=${a}%;AS_retarget=${100 - a}%`, '', '']);
}
costRows.push([ddmmyyyy(60), 'Басқа', 'Printed flyers for school visits', 45000, 'карта', '', 'Other sales expenses', 'Typography', '', '', '']);
const kpiRows = [['Ай', 'Бағыт', 'Көрсеткіш', 'Мақсат', 'Арна (міндетті емес)', 'Акция (міндетті емес)', 'Өнім (міндетті емес)', 'Менеджер (міндетті емес)'],
  ['2026-09', 'Маркетинг', 'Лидтер', 900, '', '', '', ''], ['2026-09', 'ОП', 'Сатылым', 170, '', '', '', ''], ['2026-09', 'ОП', 'Түсім', 1600000, '', '', '', ''],
  ['2026-10', 'Маркетинг', 'Лидтер', 1150, '', '', '', ''], ['2026-10', 'ОП', 'Сатылым', 230, '', '', '', ''], ['2026-10', 'ОП', 'Түсім', 2300000, '', '', '', ''], ['2026-10', 'ОП', 'Конверсия', 20, '', '', '', ''],
  ['2026-10', 'Рассылка', 'Сатылым', 60, 'Рассылка', '', '', '']];
const hypHead = ['ID', 'Сегмент', 'Өнім', 'Ұсыныс', 'Неге бұл сегмент', 'Дәлел (дерек)', 'Эксперимент', 'Табыс критерийі', 'Тәуекел', 'Статус', 'Нәтиже', 'Шешім', 'Күні', 'Гипотеза', 'Рассылкалар'];
const hypRows = [hypHead,
  ['H1', '3–4 сынып ата-аналары', 'Алтын сақа 1 мектепшілік', '2+1 акциясы', 'өткен жылы ең көп сатып алған', 'жауап беру 9%', '2 рассылка, A/B мәтін', 'жауап ≥ 10%', 'база шаршауы', 'тексерілуде', '', '', '2026-09-28', 'Акция жауап беруді арттырады', 'Рассылка 5'],
  ['H2', '5–6 сынып', 'Мега қайталау', 'тегін вебинар', 'олимпиадаға дейін 3 апта', '—', 'вебинарға шақыру', 'тіркелу ≥ 150', 'уақыт аз', 'жоспарда', '', '', '2026-10-20', 'Вебинар сатуды жылдамдатады', ''],
  ['H3', 'Таргет: Instagram', 'Математика тоқсандық', 'видео-отзыв', 'CPL төмен', 'CPL −18%', 'креатив A/B', 'CPL < 900 ₸', 'аудитория тар', 'қабылданды', 'CPL −12%', 'масштабтау', '2026-10-05', 'Ата-ана отзыві жақсы жұмыс істейді', '']];

// ---------- platform payments (Тренинги / челлендждер) ----------
const PLAT = [['Алтын Сақа Тренинг 1-жинақ', '', 5600, 'Тренинг', 20], ['Алтын Сақа Тренинг 2-жинақ', '', 5600, 'Тренинг', 14], ['Алтын Сақа Тренинг Полный пакет', '', 13440, 'Тренинг', 6],
  ['Алтын сақа пробный тест мектепішілік', '', 2500, 'Челлендж', 9], ['Алтын сақа пробный тест аудандық', '', 2500, 'Челлендж', 5], ['Мегақайталау', '', 13500, 'Тренинг', 7],
  ['Математика I-IV тоқсан', 'Күндік', 1190, 'Тренинг', 15], ['Математика I-IV тоқсан', 'Тоқсандық', 8400, 'Тренинг', 8], ['English Get Ready', 'Күндік', 1190, 'Тренинг', 8], ['English Get Ready', 'Тоқсандық', 8400, 'Тренинг', 4],
  ['Алтын түлек', '', 9900, 'Тренинг', 3], ['Altyn Asyq 4 сынып', '', 4500, 'Челлендж', 5], ['NIS дайындық', 'Айлық', 6900, 'Тренинг', 3]];
const platRows = [['paid_at', 'date_issued', 'invoice_id', 'training_name', 'tariff_name', 'item_total', 'Тип продукта', 'invoice_status', 'Жариялану статусы', 'Қайдан келді', 'Сынып', 'Облыс']];
let inv = 500000;
for (let day = -200; day <= SNAP_DAY; day++) {   // since mid-January: «all time» view has history
  const season = day < -120 ? 0.55 : day < 0 ? 0.75 : 1;
  const n = poisson(20 * season * trend(Math.max(0, day)) * WEEKDAY[wdOf(day)]);
  for (let i = 0; i < n; i++) {
    const t = at(day, hourOf(), int(0, 59)); if (t > SNAP) continue;
    const p = wpick(PLAT.map(x => [x, x[4]])), status = wpick([['2', 90], ['pending', 5], ['failed', 3], ['', 2]]), noInv = chance(0.015), invOnly = chance(0.01);
    platRows.push([invOnly || status !== '2' ? '' : t, invOnly || status !== '2' ? t : '', noInv ? '' : 'INV-' + (inv += int(1, 3)), p[0], p[1], chance(0.01) ? '' : p[2], p[3], status,
      wpick([['Жарияланған', 85], ['Жарияланбаған', 10], ['Белгісіз', 5]]), wpick([['', 70], ['Органика Instagram', 8], ['Таргетированная реклама в Инстаграм или Фейсбук', 10], ['Рассылка', 8], ['Органика WhatsApp', 4]]),
      wpick([[3, 20], [4, 22], [5, 18], [6, 14], [7, 8], [8, 6], ['', 12]]), wpick(REGIONS)]);
  }
}

// ---------- run the real Apps Script aggregators ----------
const gsFile = path.join(ROOT, 'инструменты', 'EduApp_OP_API.gs');
const ctx = {module: {exports: {}}, Date, console, Math, JSON};
vm.createContext(ctx); vm.runInContext(fs.readFileSync(gsFile, 'utf8'), ctx, {filename: 'EduApp_OP_API.gs'});
const gs = name => vm.runInContext(name, ctx);
const crm = gs('zopAggregate_')(leads, pipes, logRows, parts, mailRows, costRows, kpiRows, tariffRows, hypRows, []);
crm.crmSource = {leads: 'Демо · Сделки (кездейсоқ деректер)', pipes: 'Демо · Воронки и этапы', log: 'Демо · Журнал', error: null};
const snapIso = `${parts(SNAP).y}-${pad(parts(SNAP).m)}-${pad(parts(SNAP).d)}T${pad(SNAP_H)}:${pad(SNAP_MIN)}`;
const salesOnce = prodRows => { const s = gs('zopSalesAggregate_')([], [], platRows, parts, OFF, prodRows, []); s.snapshot = snapIso; s.errors = []; s.via = {sheet: 'қолданылмайды (CRM + платформа)', plat: 'Демо · платформа (кездейсоқ деректер)'}; return s; };
const first = salesOnce([]);
const sales = salesOnce([gs('ZOP_PROD_HEAD')].concat(gs('zopProductSeed_')(first)));
const plain = x => JSON.parse(JSON.stringify(x));   // vm objects → plain JSON

// ---------- write outputs ----------
const json = o => JSON.stringify(plain(o));
const atlasFile = path.join(ROOT, 'Атлас_отчёты.html');
let html = fs.readFileSync(atlasFile, 'utf8');
const put = (id, obj) => { const a = `<script id="${id}" type="application/json">`, i = html.indexOf(a) + a.length, j = html.indexOf('</script>', i);
  if (i < a.length) throw new Error('no block ' + id); html = html.slice(0, i) + json(obj).replace(/<\//g, '<\\/') + html.slice(j); };
put('sales', sales); put('crm', crm);
fs.writeFileSync(atlasFile, html);
fs.mkdirSync(path.join(ROOT, 'demo'), {recursive: true});
fs.writeFileSync(path.join(ROOT, 'demo', 'crm.json'), json(crm));
fs.writeFileSync(path.join(ROOT, 'demo', 'sales.json'), json(sales));
// tests/fixtures/*_demo*.json are produced on the fly by tests/lib/load.cjs from demo/crm.json + demo/sales.json (no duplicate copies)
const won = leads.filter(r => r[3] === 'Успешно').length;
console.log(`seed ${SEED} · snapshot ${snapIso}`);
console.log(`deals ${leads.length - 1} (won ${won}) · mailings ${mailRows.length - 1} · costs ${costRows.length - 1} · platform rows ${platRows.length - 1}`);
console.log(`crm: main ${crm.main.length} cells · sales: cube ${sales.cube.length} rows · catalog ${sales.plan.catalog.length}`);
console.log('→ Атлас_отчёты.html (embedded crm + sales), demo/crm.json, demo/sales.json');
