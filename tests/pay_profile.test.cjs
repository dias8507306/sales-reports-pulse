// Төлем × сынып / өңір (иесі 15.10: «покупки по регионам и по классам»): ОП кестесі «Сыныбы» + платформадағы ұқсас баған; 1–2 төлем жасырылады.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const X = loadGs().exp, J = x => JSON.parse(JSON.stringify(x));
const parts = d => { if (!(d instanceof Date) || isNaN(d)) return null; const x = new Date(d.getTime() + 5 * 3600e3); return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours()}; };

test('payProf: сынып пен өңір күн бойынша, платформадан тек төленген, 3-тен аз — жасырылған', () => {
  const SH = ['Дата', 'Қайдан келді', 'Миникурс Атауы', 'Төлем сомасы', 'Сыныбы', 'Облыс'];
  const s = (cls, obl) => ['16.09.2026', 'WhatsApp', 'Алтын сақа', 5000, cls, obl];
  const sheet = [SH, s('5 сынып', 'Павлодар облысы'), s(5, 'Павлодар'), s('5', 'ПАВЛОДАР'), s(6, 'Қызылорда облысы')];
  const PH = ['invoice_id', 'invoice_status', 'paid_at', 'item_total', 'training_name', 'grade', 'region'];
  const plat = [PH, ['i1', 'paid', '2026-09-16', 3000, 'Тренинг', 7, 'Астана'], ['i2', 'paid', '2026-09-16', 3000, 'Тренинг', 7, 'Астана'], ['i3', 'paid', '2026-09-16', 3000, 'Тренинг', 7, 'Астана'], ['i4', 'unpaid', '2026-09-16', 3000, 'Тренинг', 7, 'Астана']];
  const out = X.zopSalesAggregate_(sheet, [], plat, parts, 0, [], []), R = out.dict.region, P = out.payProf;
  assert.deepEqual(J(P.cls), [['p', '2026-09-16', 7, 3, 9000], ['s', '2026-09-16', -1, 1, 5000], ['s', '2026-09-16', 5, 3, 15000]]);
  assert.deepEqual(J(P.reg).map(r => [r[0], R[r[2]], r[3]]), [['p', 'Астана', 3], ['s', 'Павлодар', 3], ['s', 'Жасырылған (< 3)', 1]]);
  assert.deepEqual(J(P.cols), {sheet: {cls: 'Сыныбы', reg: 'Облыс'}, plat: {cls: 'grade', reg: 'region'}});
});
