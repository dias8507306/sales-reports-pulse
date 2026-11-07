// F01, F05, F06, F16: статусы оплат, происхождение даты/суммы платформы, дубли, бюджеты-строки.
// Данные синтетические: только поля, которые читает агрегатор, без персональных данных.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs, cardCalc, fixture} = require('./lib/load.cjs');
const gs = loadGs(), X = gs.exp;

// Asia/Qyzylorda = UTC+5 круглый год (как parts() в Apps Script при таймзоне таблицы)
const parts = d => { if (!(d instanceof Date) || isNaN(d)) return null; const x = new Date(d.getTime() + 5 * 3600e3); return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours()}; };
const PH = ['invoice_id', 'invoice_status', 'paid_at', 'date_issued', 'item_total', 'invoice_total', 'tariff_price', 'training_name', 'tariff_name', 'Тип продукта', 'Қайдан келді'];
const P = o => PH.map(h => o[h] ?? '');
const agg = (plat, sheet) => X.zopSalesAggregate_(sheet || [['Дата', 'Қайдан келді', 'Миникурс Атауы', 'Төлем сомасы', 'Бағасы']], [], [PH].concat(plat.map(P)), parts, 0, [], []);
const sum = (out, pred) => out.cube.filter(pred).reduce((a, r) => a + r[r.length - 1], 0);
const cnt = (out, pred) => out.cube.filter(pred).reduce((a, r) => a + r[r.length - 2], 0);
const ok = out => r => !X.zopSkipReason_(r, out.dict.pstatus);

test('F01: «unpaid» не считается оплатой; пустой и незнакомый статус — unknown', () => {
  const S = X.zopPayStatus_;
  assert.equal(S('unpaid'), 'unpaid', 'substring «paid» must not make it paid');
  assert.equal(S('UNPAID'), 'unpaid'); assert.equal(S('not_paid'), 'unpaid');
  assert.equal(S('2'), 'paid'); assert.equal(S('paid'), 'paid'); assert.equal(S('Оплачено'), 'paid'); assert.equal(S('Mon Jan 01 1900 00:00:00 GMT+0500'), 'paid');
  assert.equal(S(''), 'unknown'); assert.equal(S(null), 'unknown'); assert.equal(S('3'), 'unknown'); assert.equal(S('paid?'), 'unknown');
  assert.equal(S('refunded'), 'rejected'); assert.equal(S('canceled'), 'rejected');
});

test('F01: клиент карточки и сервер классифицируют статусы одинаково', () => {
  const C = cardCalc(gs, {data: {dict: {pstatus: []}, cube: [], snapshot: '2026-10-09T11:00', plan: {sources: [], products: []}}, kind: 'sales'});
  assert.ok(C.payStatus, 'client exposes payStatus');
  const samples = ['', ' ', '2', 'paid', 'PAID', 'unpaid', 'Unpaid', 'not paid', 'pending', 'failed', 'refund', 'success', 'сәтті', 'төленбеген', 'Mon Jan 01 1900', '01.01.1900', '3', 'x', null, undefined, 2];
  for (const v of samples) assert.equal(C.payStatus(v), X.zopPayStatus_(v), 'status ' + JSON.stringify(v));
  assert.deepEqual(JSON.parse(JSON.stringify(X.ZOP_PAY)), JSON.parse(JSON.stringify(gs.value('ZOP_PAY'))));
});

test('F01: суммы unpaid / unknown остаются отдельно, не в оплатах', () => {
  const out = agg([
    {invoice_id: 'A', invoice_status: '2', paid_at: '2026-10-08 10:00', item_total: 12000, training_name: 'Мега қайталау'},
    {invoice_id: 'B', invoice_status: 'unpaid', paid_at: '2026-10-08 11:00', item_total: 990, training_name: 'Күнделікті'},
    {invoice_id: 'C', invoice_status: '', paid_at: '2026-10-08 12:00', item_total: 500, training_name: 'Күнделікті'}]);
  assert.equal(sum(out, ok(out)), 12000);
  const C = cardCalc(gs, {data: Object.assign(out, {snapshot: '2026-10-09T11:00', plan: {sources: [], products: [], catalog: []}}), kind: 'sales'});
  const R = C.sales('2026-10-08');
  assert.equal(R.sum, 12000); assert.equal(R.rows, 1);
  assert.equal(R.skipped['төленбеген'], 1); assert.equal(R.skippedS['төленбеген'], 990);
  assert.equal(R.skipped['статус белгісіз'], 1); assert.equal(R.skippedS['статус белгісіз'], 500);
});

test('F06: строка без paid_at не становится оплатой дня счёта', () => {
  const out = agg([
    {invoice_id: 'A', invoice_status: '2', date_issued: '2026-10-07 09:00', item_total: 990, training_name: 'Күнделікті'},
    {invoice_id: 'B', invoice_status: '2', item_total: 990, training_name: 'Күнделікті'}]);
  assert.equal(out.quality.platInvoiceDateOnly, 1); assert.equal(out.quality.platNoDate, 1);
  assert.equal(sum(out, ok(out)), 0, 'neither row is a confirmed payment of a day');
  const r = out.cube.find(x => x[1] === '2026-10-07'); assert.equal(X.zopSkipReason_(r, out.dict.pstatus), 'төлем күні жоқ (шот күні)');
});

test('F06: граница суток по Asia/Qyzylorda (UTC+5)', () => {
  const out = agg([
    {invoice_id: 'A', invoice_status: '2', paid_at: '2026-10-08T19:30:00Z', item_total: 100, training_name: 'X'},   // 00:30 09.10 в Кызылорде
    {invoice_id: 'B', invoice_status: '2', paid_at: '2026-10-08T18:59:00Z', item_total: 200, training_name: 'X'},   // 23:59 08.10
    {invoice_id: 'C', invoice_status: '2', paid_at: '2026-10-08 23:59', item_total: 400, training_name: 'X'}]);      // без зоны = уже местное время
  assert.equal(sum(out, r => r[1] === '2026-10-09'), 100);
  assert.equal(sum(out, r => r[1] === '2026-10-08'), 600);
});

test('F05: invoice_total по нескольким позициям считается один раз; точный дубль позиции — один раз', () => {
  const out = agg([
    {invoice_id: 'A', invoice_status: '2', paid_at: '2026-10-08 10:00', invoice_total: 11520, training_name: 'Т1'},
    {invoice_id: 'A', invoice_status: '2', paid_at: '2026-10-08 10:00', invoice_total: 11520, training_name: 'Т2'},
    {invoice_id: 'B', invoice_status: '2', paid_at: '2026-10-08 11:00', item_total: 990, training_name: 'Т3'},
    {invoice_id: 'B', invoice_status: '2', paid_at: '2026-10-08 11:00', item_total: 990, training_name: 'Т3'},
    {invoice_id: 'C', invoice_status: '2', paid_at: '2026-10-08 12:00', item_total: 990, training_name: 'Т3'}]);
  assert.equal(sum(out, ok(out)), 11520 + 990 + 990);
  assert.equal(out.quality.platInvoiceTotalRepeat, 1); assert.equal(out.quality.duplicateRows, 1);
  assert.equal(out.rows.platform, 4);
});

test('F06: прайс-цена вместо суммы помечается; отсутствие суммы ≠ ноль оплаты (видно в quality)', () => {
  const out = agg([
    {invoice_id: 'A', invoice_status: '2', paid_at: '2026-10-08 10:00', tariff_price: 12000, training_name: 'Т'},
    {invoice_id: 'B', invoice_status: '2', paid_at: '2026-10-08 10:00', training_name: 'Т'}]);
  assert.equal(out.quality.platReferencePrice, 1); assert.equal(out.quality.platNoAmount, 1);
  assert.equal(out.cube.filter(r => r[10] === 2).length, 1); assert.equal(out.cube.filter(r => r[10] === 3).length, 1);
});

test('F16: бюджет «12 000» читается как 12000; нечитаемый — null, не 0', () => {
  const B = X.zopBudget_;
  assert.equal(B(12000), 12000); assert.equal(B('12 000'), 12000); assert.equal(B('12 000 ₸'), 12000); assert.equal(B('990,50'), 990.5); assert.equal(B('1,200,000'), 1200000);
  assert.equal(B(''), 0); assert.equal(B(null), 0);
  assert.equal(B('около 10к'), null); assert.equal(B(NaN), null);
});

test('Старые кубы (без колонок происхождения) считаются как раньше', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const cfg = fixture('card_sales_demo.json'), old = cfg.data;   // fixture = the card's cfg as embedded in the page
  assert.ok(old.cube.every(r => r.length === 11), 'fixture is a schema-1 cube');
  const C = cardCalc(gs, cfg), iso = cfg.date;
  const R = C.sales(iso);
  // контроль: сумма = все строки листа + платформа с invoice и статусом «2»
  const ST = old.dict.pstatus, exp = old.cube.filter(r => r[1] === iso && (r[0] === 's' || (r[7] === 1 && X.zopPayStatus_(ST[r[6]] || '') === 'paid')));
  assert.equal(R.sum, exp.reduce((a, r) => a + r[r.length - 1], 0));
  assert.equal(R.rows, exp.reduce((a, r) => a + r[r.length - 2], 0));
  assert.ok(R.rows > 0);
});

test('ОП: «Оқушы тізімі» × «Енгізген қызметкер» — количество и сумма по дням, без данных учеников', () => {
  const H = ['Дата', 'Қайдан келді', 'Миникурс Атауы', 'Төлем сомасы', 'Бағасы', 'Енгізген қызметкер', 'Оқушы аты-жөні'];
  const out = agg([], [H, ['10.10.2026', 'WhatsApp', 'Алтын сақа', 12000, '', 'Менеджер А', 'Ученик 1'], ['10.10.2026', 'Таргет', 'Алтын сақа', 8000, '', 'Менеджер А', 'Ученик 2'],
    ['10.10.2026', 'WhatsApp', 'Алтын сақа', 5000, '', 'Менеджер Б', 'Ученик 3'], ['11.10.2026', 'WhatsApp', 'Алтын сақа', '', 3000, '', 'Ученик 4']]);
  const E = out.dict.emp, row = (d, e) => out.sheetMan.find(r => r[0] === d && E[r[1]] === e);
  assert.deepEqual([...row('2026-10-10', 'Менеджер А').slice(2)], [2, 20000]);
  assert.deepEqual([...row('2026-10-10', 'Менеджер Б').slice(2)], [1, 5000]);
  assert.deepEqual([...row('2026-10-11', '').slice(2)], [1, 3000]);   // пусто → в Атласе «Көрсетілмеген»
  assert.equal(out.sheetMan.reduce((a, r) => a + r[2], 0), cnt(out, r => r[0] === 's'));   // столько же, сколько в кубе
  assert.ok(!JSON.stringify(out).includes('Ученик'), 'имена учеников не выходят');
});
