// EduApp Sales Intelligence: хранилище проекта, расходы, кошелёк, тарифы, планы, направления, кампании, A/B, инсайты, отчёты.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadAtlas, loadGs, fixture} = require('./lib/load.cjs');
const SNAP = fixture('atlas_demo.json'), CARD = fixture('card_sales_demo.json');
const atlas = () => { const A = loadAtlas(SNAP); A.initSales(JSON.parse(JSON.stringify(CARD.data))); return A; };
const arr = x => JSON.parse(JSON.stringify(x));

test('Хранилище: пустое/битое → пустая схема; миграция v1; журнал изменений', () => {
  const A = atlas();
  assert.deepEqual(arr(A.zsiLoad('не json')).costs, []); assert.equal(A.zsiLoad(null).v, 2);
  const v1 = A.zsiLoad({costs: [{id: 'c1', amount: 500, date: '2026-10-01'}]}); assert.equal(v1.costs[0].sum, 500);
  let s = A.zsiApply(A.zsiEmpty(), {type: 'upsert', coll: 'costs', item: {id: 'x', date: '2026-10-09', cat: 'Wazzup top-up', sum: 10000}});
  s = A.zsiApply(s, {type: 'upsert', coll: 'costs', item: {id: 'x', date: '2026-10-09', cat: 'Wazzup top-up', sum: 12000}});
  s = A.zsiApply(s, {type: 'delete', coll: 'costs', id: 'x'});
  assert.deepEqual(s.log.map(l => l.action), ['қосылды', 'өзгертілді', 'жойылды']); assert.equal(s.log[1].before.sum, 10000); assert.equal(s.costs.length, 0);
  assert.throws(() => A.zsiApply(s, {type: 'upsert', coll: 'nope', item: {}}));
});

test('Расходы: пополнение — не списание; строка проекта, совпадающая со строкой таблицы, — дубль; неизвестная сумма не 0', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = atlas(); A.D.costs = []; A.D.costsAll = [{d: 59, sub: 'Рассылка', det: 'Subscription balance', sum: 50000, pay: '', kind: 'balance', cat: 'Wazzup top-up', prov: 'Wazzup', doc: '', camp: '', warn: []}];
  const st = A.zsiApply(A.zsiApply(A.zsiEmpty(), {type: 'upsert', coll: 'costs', item: {date: '2026-10-08', cat: 'Wazzup top-up', prov: 'Wazzup', sub: 'Рассылка', sum: 50000}}),
    {type: 'upsert', coll: 'costs', item: {date: '2026-10-08', cat: 'WhatsApp template charges', prov: 'Wazzup', msgs: 500}});
  const f = A.finStats(59, 59, st);
  assert.equal(f.topup, 50000, 'the same top-up entered twice counts once'); assert.equal(f.dup, 1); assert.equal(f.usage, 0); assert.equal(f.usageKnown, false);
  assert.equal(f.unknown, 1, 'messages without a tariff = unknown, not zero');
  assert.equal(f.cashOut, 50000);
});

test('Тарифы зависят от даты; сумма по тарифу; изменение тарифа не меняет сохранённую сумму', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = atlas(); A.D.costsAll = [];
  let st = A.zsiApply(A.zsiEmpty(), {type: 'upsert', coll: 'tariffs', item: {prov: 'Wazzup', type: 'marketing', from: '2026-10-01', to: '2026-10-31', price: 54, cur: 'KZT'}});
  st = A.zsiApply(st, {type: 'upsert', coll: 'tariffs', item: {prov: 'Wazzup', type: 'marketing', from: '2026-11-01', price: 60, cur: 'KZT'}});
  const T = A.allTariffs(st).list; assert.equal(A.tariffAt(T, 'Wazzup', 'marketing', 59).price, 54); assert.equal(A.tariffAt(T, 'Wazzup', 'marketing', 83).price, 60);
  const c1 = A.costFromItem({id: 'a', date: '2026-10-08', cat: 'WhatsApp template charges', prov: 'Wazzup', tariff: 'marketing', msgs: 1000}, T);
  assert.equal(c1.sum, 54000); assert.equal(c1.src, 'tariff');
  const c2 = A.costFromItem({id: 'b', date: '2026-10-08', cat: 'WhatsApp template charges', prov: 'Wazzup', tariff: 'marketing', msgs: 1000, sum: 51000}, T);
  assert.equal(c2.sum, 51000, 'a stored amount wins over the tariff');
  st = A.zsiApply(st, {type: 'upsert', coll: 'tariffs', item: {prov: 'Wazzup', type: 'marketing', from: '2026-10-15', price: 70}});
  assert.equal(A.allTariffs(st).warn.length >= 1, true, 'overlapping periods are flagged');
});

test('Кошелёк: начальный остаток + пополнения − списания ± корректировки', () => {
  const A = atlas(); A.D.costsAll = [];
  let st = A.zsiEmpty(); const put = item => { st = A.zsiApply(st, {type: 'upsert', coll: 'costs', item}); };
  put({date: '2026-10-01', cat: 'Wallet opening balance', prov: 'Wazzup', sum: 5000}); put({date: '2026-10-02', cat: 'Wazzup top-up', prov: 'Wazzup', sum: 50000});
  put({date: '2026-10-05', cat: 'WhatsApp template charges', prov: 'Wazzup', sum: 30000}); put({date: '2026-10-06', cat: 'Adjustment', prov: 'Wazzup', sum: -1000});
  const W = A.walletCalc(60, st); assert.equal(W.balance, 5000 + 50000 - 30000 - 1000); assert.equal(W.chargesKnown, true);
  assert.equal(A.walletCalc(53, st).balance, 55000, 'before the charge');
});

test('Распределение по кампаниям совпадает с Apps Script', () => {
  const A = atlas(), g = loadGs().exp;
  for (const [t, s] of [['А=30000; Б=20000', 50000], ['А:60%; Б:40%', 50000], ['А=1; Б=2', 50000], ['Рассылка 7', 12000], ['', 1]]) assert.deepEqual(arr(A.costAlloc(t, s)), arr(g.zopCostAlloc_(t, s)), t);
});

test('Направления: то же правило, что в Apps Script; сумма направлений = сумма оплат', () => {
  const A = atlas(), g = loadGs().exp;
  for (const n of ['Алтын Сақа Тренинг', '"Алтын сақа" пробный тест; 3 сынып', '«Алтын түлек» дайындық', '6-сынып; NIS ағылшын', 'English A0 · Күндік', '3-сынып. I-IV тоқсан. · Жылдық', 'Челлендж'])
    assert.equal(A.prodFamily(n, 'p'), g.zopSecGuess_(n, 'p'), n);
  const day = Math.round((Date.UTC(2026, 9, 8) - Date.UTC(2026, 7, 10)) / 864e5), fs = A.famStats(day - 6, day), ps = A.payStats(day - 6, day);
  assert.equal(fs.tot.s, ps.sheet.s + ps.plat.s); assert.equal(fs.tot.n, ps.sheet.n + ps.plat.n);
  const ch = A.famChange(day - 6, day); assert.equal(Math.round(ch.rows.reduce((a, r) => a + r.delta, 0)), Math.round(ch.delta));
});

test('Направление из «Өнім тізімі» G важнее правила; новое направление появляется само', () => {
  const A = atlas(), S = A.SALES, N = S.dict.name, i = N.findIndex(n => /алтын\s*са[қк]а/i.test(n)), old = S.dict.fam;
  try {
    S.dict.fam = N.map((_, j) => j === i ? 'Тест бағыты' : '');
    assert.equal(A.prodFamily(N[i], 'p', i), 'Тест бағыты'); assert.equal(A.prodFamily(N[i], 'p'), A.prodFamily(N[i], 'p', N.length + 5));
    const day = Math.round((Date.UTC(2026, 9, 8) - Date.UTC(2026, 7, 10)) / 864e5), fs = A.famStats(0, day + 400), ps = A.payStats(0, day + 400);
    assert.equal(fs.tot.s, ps.sheet.s + ps.plat.s);
    assert.ok(fs.list.some(e => e.name === 'Тест бағыты'));
  } finally { S.dict.fam = old; }
});

test('План-факт: факт, ожидаемый темп и нужный темп на остаток месяца', () => {
  const A = atlas(), month = A.MONTHS.find(m => m.k === '2026-10');
  const p = A.planFact({month: '2026-10', metric: 'paySum', scope: 'all', target: 10000000});
  const last = Math.min(A.END, month.b), fact = A.planActual('paySum', 'all', Math.max(0, month.a), last);
  assert.equal(p.fact, fact); assert.equal(p.elapsed, last - month.a + 1); assert.equal(p.left, month.b - last);
  assert.ok(Math.abs(p.need - (10000000 - fact) / p.left) < 1e-6);
  assert.equal(A.planFact({month: '1999-01', metric: 'paySum', target: 1}).status, 'no-month');
});

test('Инсайты и отчёты: структура, пороги, без «кеше/бүгін»', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = atlas(), I = A.ruleInsights(A.END, A.zsiEmpty());
  assert.ok(I.length > 0); I.forEach(x => { for (const k of ['severity', 'area', 'observation', 'evidence', 'action', 'confidence']) assert.ok(k in x, k); });
  // 15.10: ОП кестесі алынды — «математика / ағылшын менеджерде жоқ» ескертуі CRM-сатылымда шықпайды (тексеру алынды)
  for (const k of ['day', 'week', 'month', 'product', 'source', 'dq']) { const R = A.buildReport(k, A.END, A.zsiEmpty()); assert.ok(R.sections.length >= 2, k); const t = A.reportText(R); assert.ok(!/кеше|бүгін|[↑↓▲▼]/i.test(t), k); }
});

test('Регрессия: два платежа одной суммы в один день для разных номеров — не дубль', () => {
  const A = atlas(); A.D.costsAll = [{d: 56, sub: '7000', det: 'Subscription balance', sum: 20000, pay: 'партнёр', kind: 'balance'}, {d: 56, sub: 'Рассылка', det: 'Subscription balance', sum: 20000, pay: 'Stripe', kind: 'balance'},
    {d: 56, sub: '7000', det: 'Subscription balance', sum: 0, pay: 'бонусы', kind: 'balance'}, {d: 56, sub: 'Рассылка', det: 'Subscription balance', sum: 0, pay: 'бонусы', kind: 'balance'}];
  const C = A.allCosts(A.zsiEmpty()); assert.equal(C.filter(c => c.dup).length, 0);
  assert.equal(A.finStats(56, 56, A.zsiEmpty()).topup, 40000);
});

test('Отчёт по умолчанию — последний полный день, а не текущий незавершённый', () => {
  const A = atlas(), sd = Math.round((Date.UTC(+A.D.snapshot.slice(0, 4), +A.D.snapshot.slice(5, 7) - 1, +A.D.snapshot.slice(8, 10)) - Date.UTC(2026, 7, 10)) / 864e5);
  assert.equal(A.lastFullDay(), Math.min(sd - 1, A.END));
});
