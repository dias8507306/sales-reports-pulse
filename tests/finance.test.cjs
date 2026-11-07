// Қаржы: доход = тіркелген төлемдер, расход = все строки «Расходы» (баланс — в день оплаты, тариф — по дням).
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadAtlas, fixture} = require('./lib/load.cjs');
const SNAP = fixture('atlas_demo.json'), CARD = fixture('card_sales_demo.json');
const day = iso => Math.round((Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) - Date.UTC(2026, 7, 10)) / 864e5);

test('Пайда = төлемдер − шығын; тариф делится по дням; категории — из «Подписка»', () => {
  const d = JSON.parse(JSON.stringify(SNAP)), a = day('2026-10-08');
  d.costs = [{d: 0, sub: 'WhatsApp', det: 'plan for 1 year', sum: 365000, kind: 'plan', months: 12},
    {d: a, sub: 'Рассылка', det: 'Subscription balance', sum: 5000, kind: 'balance', months: 0}, {d: a, sub: 'Таргет', det: 'Instagram', sum: 20000, kind: 'balance', months: 0}];
  const A = loadAtlas(d); A.initSales(CARD.data);
  const f = A.finStats(a, a), share = A.planShare(d.costs[0], a, a);
  assert.ok(share > 0 && share < 2000, 'one day of a yearly tariff');
  assert.equal(f.cost, 25000 + share); assert.equal(f.paid, 25000);
  assert.deepEqual([...f.cats.map(c => c.name)].sort(), ['Wazzup subscription', 'Wazzup top-up'].sort(), 'old rows without «Категория» are mapped by kind');
  assert.equal(f.topup, 25000); assert.equal(f.subs, share); assert.equal(f.cashOut, 25000);
  const ps = A.payStats(a, a); assert.equal(f.income, ps.sheet.s + ps.plat.s); assert.equal(f.profit, f.income - f.cost);
  assert.equal(f.costPartial, false);
});

test('Период раньше первой строки «Расходы» — неполный; нет продаж — нет дохода, а не ноль', () => {
  const d = JSON.parse(JSON.stringify(SNAP)), a = day('2026-10-05');
  d.costs = [{d: a, sub: 'Рассылка', det: 'Subscription balance', sum: 1000, kind: 'balance', months: 0}];
  const A = loadAtlas(d);
  assert.equal(A.finStats(a - 3, a).costPartial, true);
  assert.equal(A.finStats(a, a).income, null); assert.equal(A.finStats(a, a).profit, null);
});

test('Калькулятор рассылки: порог окупаемости и пример из аудита (54 ₸, 4 800 ₸, 1 000 сообщений)', () => {
  const A = loadAtlas(SNAP);
  const r1 = A.mailCalc({n: 1000, unit: 54, price: 4800, conv: 0.01}), r2 = A.mailCalc({n: 1000, unit: 54, price: 4800, conv: 0.02});
  assert.equal(r1.cost, 54000); assert.equal(r1.rev, 48000); assert.equal(r1.diff, -6000); assert.equal(r1.ok, false);
  assert.equal(r2.diff, 42000); assert.equal(r2.ok, true);
  assert.equal(r1.beConv, 54000 / 4800000); assert.equal(r1.beBuys, 12);
  assert.equal(A.mailCalc({n: 1000, unit: 54, price: 0, conv: 0.01}).beConv, null, 'no price → no threshold, not 0');
  const r3 = A.mailCalc({n: 1000, unit: 54, fixed: 6000, price: 4800, conv: 0.0125}); assert.equal(r3.cost, 60000); assert.equal(r3.diff, 0); assert.equal(r3.ok, true);
});

test('Сверка потоков (до 15.10): платформа без invoice = копия «Оқушы тізімі»; с CRM-продажами сверка выключена', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A0 = loadAtlas(SNAP); A0.initSales(JSON.parse(JSON.stringify(CARD.data))); assert.equal(A0.payReconcile(), null, '15.10: ОП кестесі алынды');
  const A = loadAtlas(null), d = JSON.parse(JSON.stringify(CARD.data)); A.initSales(d);
  const rc = A.payReconcile(); assert.ok(rc.days > 0); assert.equal(rc.matched, rc.days, 'fixture: every day matches');
  const r = d.cube.find(x => x[0] === 's'); r[r.length - 1] += 1000; A.initSales(d);   // break one day
  const rc2 = A.payReconcile(); assert.equal(rc2.bad.length, 1); assert.equal(rc2.bad[0].day, r[1]);
  A.initSales(null); assert.equal(A.payReconcile(), null);
});
