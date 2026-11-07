// Cost Control (P1.4): расходы v2, тарифы, распределение, дубли, валюты, неизвестная сумма ≠ 0.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const g = loadGs().exp;
const B = Date.UTC(2026, 7, 10), parts = d => { const x = new Date(d.getTime() + 5 * 3600e3); return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate()}; };
const dayIdx = p => p ? Math.round((Date.UTC(p.y, p.m - 1, p.d) - B) / 864e5) : -1, D = s => new Date(Date.parse(s + 'T07:00:00Z')), di = s => dayIdx(parts(D(s)));
const HEAD = ['Дата', 'Подписка', 'Детали', 'Сумма, ₸', 'Оплата', 'Рассылка (если знаете)', 'Категория', 'Поставщик', 'Валюта', 'Курс', 'Кезең басы', 'Кезең соңы', 'Кампания', 'Хабарлама саны', 'Тариф', 'Құжат', 'Тексерілді', 'ID'];
const row = o => HEAD.map(h => o[h] === undefined ? '' : o[h]);
const costs = (rows, tariffs) => g.zopCosts_([HEAD].concat(rows.map(row)), parts, dayIdx, tariffs || []);

test('Старый формат листа читается как раньше: пополнение / тариф по «Детали»', () => {
  const C = g.zopCosts_([['Дата', 'Подписка', 'Детали', 'Сумма, ₸', 'Оплата'], ['07.10.2026', 'Рассылка', 'Subscription balance', 50000, 'партнёр'], ['04.10.2026', '7000', '1 channel on the "Waba" plan for 1 year', 346600, 'партнёр']], parts, dayIdx, []);
  assert.equal(C[0].cat, 'Wazzup top-up'); assert.equal(C[0].kind, 'balance'); assert.equal(C[0].sum, 50000);
  assert.equal(C[1].cat, 'Wazzup subscription'); assert.equal(C[1].kind, 'plan'); assert.equal(C[1].months, 12);
});

test('Пополнение не удваивается; повторный импорт тех же строк — дубли, в итог не идут', () => {
  const r = {'Дата': D('2026-10-09'), 'Подписка': 'Рассылка', 'Детали': 'Subscription balance', 'Сумма, ₸': 50000, 'Категория': 'Wazzup top-up', 'Поставщик': 'Wazzup'};
  const C = costs([r, r]); assert.equal(C[0].dup, false); assert.equal(C[1].dup, true);
  const total = C.filter(c => !c.dup).reduce((a, c) => a + c.sum, 0); assert.equal(total, 50000);
  const W = costs([Object.assign({}, r, {ID: 'a1'}), Object.assign({}, r, {ID: 'a1'}), Object.assign({}, r, {ID: 'a2'})]);
  assert.deepEqual([...W.map(c => c.dup)], [false, true, false], 'same ID twice = re-import; another ID = a real second top-up');
});

test('Распределение по кампаниям: суммы, проценты, ошибка при несходящейся сумме', () => {
  assert.deepEqual([...g.zopCostAlloc_('Рассылка 5=30000; Рассылка 6=20000', 50000).items.map(x => x.sum)], [30000, 20000]);
  assert.equal(g.zopCostAlloc_('Рассылка 5=30000; Рассылка 6=20000', 50000).err, null);
  assert.deepEqual([...g.zopCostAlloc_('А:60%; Б:40%', 50000).items.map(x => x.sum)], [30000, 20000]);
  assert.match(g.zopCostAlloc_('А=10000; Б=10000', 50000).err, /≠/);
  assert.deepEqual([...g.zopCostAlloc_('Рассылка 7', 12000).items.map(x => x.sum)], [12000]);
});

test('Неизвестная сумма не становится нулём; валюта без курса — неизвестно; с курсом — в ₸', () => {
  const C = costs([{'Дата': D('2026-10-09'), 'Категория': 'WhatsApp template charges', 'Хабарлама саны': 1000}, {'Дата': D('2026-10-09'), 'Категория': 'Other sales expenses', 'Сумма, ₸': 20, 'Валюта': 'USD'},
    {'Дата': D('2026-10-09'), 'Категория': 'Other sales expenses', 'Сумма, ₸': 20, 'Валюта': 'USD', 'Курс': 500}]);
  assert.equal(C[0].sum, null); assert.ok(C[0].warn.length); assert.equal(C[1].sum, null); assert.match(C[1].warn[0], /курс/); assert.equal(C[2].sum, 10000);
});

test('Тариф: сумма по тарифу на дату; пересечение периодов видно; записанная сумма важнее тарифа', () => {
  const T = g.zopTariffs_([['Поставщик', 'Хабарлама түрі', 'Басталуы', 'Аяқталуы', 'Бағасы', 'Валюта', 'Дереккөз'], ['Wazzup', 'marketing', D('2026-10-01'), D('2026-10-31'), 54, 'KZT', 'шот'],
    ['Wazzup', 'marketing', D('2026-11-01'), '', 60, 'KZT', 'шот'], ['Wazzup', 'utility', D('2026-10-01'), D('2026-10-31'), 16, 'KZT', 'шот'], ['Wazzup', 'utility', D('2026-10-20'), '', 18, 'KZT', '']], parts, dayIdx);
  assert.equal(T.list.length, 4); assert.equal(T.warn.length, 1, 'utility periods overlap');
  const C = costs([{'Дата': D('2026-10-09'), 'Категория': 'WhatsApp template charges', 'Поставщик': 'Wazzup', 'Тариф': 'marketing', 'Хабарлама саны': 1000},
    {'Дата': D('2026-11-02'), 'Категория': 'WhatsApp template charges', 'Поставщик': 'Wazzup', 'Тариф': 'marketing', 'Хабарлама саны': 1000},
    {'Дата': D('2026-10-09'), 'Категория': 'WhatsApp template charges', 'Поставщик': 'Wazzup', 'Тариф': 'marketing', 'Хабарлама саны': 1000, 'Сумма, ₸': 51000}], T.list);
  assert.equal(C[0].sum, 54000); assert.equal(C[0].src, 'tariff'); assert.equal(C[1].sum, 60000); assert.equal(C[2].sum, 51000); assert.equal(C[2].src, 'row');
});

test('Форма: проверка ввода и дублей', () => {
  const ok = g.zopCostValidate_({date: '2026-10-11', cat: 'Wazzup top-up', prov: 'Wazzup', sum: 10000, cur: 'KZT'}, [], []);
  assert.equal(ok.ok, true); assert.equal(ok.dup, false);
  assert.equal(g.zopCostValidate_({date: '2026-10-11', cat: 'Wazzup top-up', prov: 'Wazzup', sum: 10000, cur: 'KZT'}, [], [{id: 'x', key: ok.key}]).dup, true);
  const bad = g.zopCostValidate_({date: '', cat: 'Кофе', sum: '', msgs: '', cur: 'USD'}, [], []);
  assert.equal(bad.ok, false); assert.ok(bad.errors.length >= 4);
  assert.equal(g.zopCostValidate_({date: '2026-10-11', cat: 'Other sales expenses', sum: 50000, cur: 'KZT', camp: 'А=1; Б=2'}, [], []).ok, false);
});
