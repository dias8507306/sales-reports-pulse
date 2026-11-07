// P0.1: построчная сверка «Оқушы тізімі» ↔ строки платформы без invoice; уникальность счетов. Наружу — только счётчики.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const g = loadGs().exp;
const d = s => new Date(Date.parse(s + 'T05:00:00Z'));
const SH = ['Дата', 'Аты-жөні', 'Телефон', 'Миникурс Атауы', 'Төлем сомасы', 'Қайдан келді'];
const PH = ['invoice_id', 'invoice_status', 'paid_at', 'item_total', 'training_name', 'Дата', 'Аты-жөні', 'Телефон', 'Миникурс Атауы', 'Төлем сомасы', 'Қайдан келді'];
const sheet = rows => [SH].concat(rows);
const copy = r => ['', '', '', '', ''].concat(r);   // platform row without invoice = the same sheet columns

test('Полная копия: все строки совпали, ничего лишнего; два одинаковых ребёнка в один день — не дубль', () => {
  const a = [d('2026-10-08'), 'А', '+7701', 'Алтын Сақа Тренинг', 4800, 'Рассылка'], b = [d('2026-10-08'), 'Б', '+7702', 'Алтын Сақа Тренинг', 9600, 'Рассылка'];
  const R = g.zopReconcileRows_(sheet([a, b, a]), [PH, copy(a), copy(b), copy(a), ['INV1', '2', d('2026-10-08'), 990, 'X', '', '', '', '', '', '']]);
  assert.equal(R.matched, 3); assert.equal(R.sheetOnly, 0); assert.equal(R.copyOnly, 0);
  assert.equal(R.sheet.duplicateKeys, 1, 'identical rows are visible as a duplicate key (to check), still matched one-to-one');
  assert.equal(R.keyCols, 6); assert.equal(R.invoices.unique, 1);
  assert.ok(!JSON.stringify(R).includes('+7701') && !JSON.stringify(R).includes('Алтын Сақа'), 'no values leave');
});

test('Строка только в листе / только в платформе видна с суммой; расхождение в сумме — это два несовпадения', () => {
  const a = [d('2026-10-08'), 'А', '+7701', 'Т', 4800, 'Рассылка'], b = [d('2026-10-09'), 'Б', '+7702', 'Т', 9600, 'WhatsApp'];
  const R = g.zopReconcileRows_(sheet([a, b]), [PH, copy(a), copy(b.slice(0, 4).concat([9000, 'WhatsApp']))]);
  assert.equal(R.matched, 1); assert.equal(R.sheetOnly, 1); assert.equal(R.copyOnly, 1);
  assert.equal(R.sheetOnlySum, 9600); assert.equal(R.copyOnlySum, 9000);
});

test('Счета: несколько позиций, смешанные статусы, возвраты считаются', () => {
  const R = g.zopReconcileRows_(sheet([]), [PH, ['I1', '2', d('2026-10-01'), 100, 'X', '', '', '', '', '', ''], ['I1', '2', d('2026-10-01'), 200, 'Y', '', '', '', '', '', ''],
    ['I2', '2', d('2026-10-02'), 100, 'X', '', '', '', '', '', ''], ['I2', 'refunded', d('2026-10-03'), 100, 'X', '', '', '', '', '', '']]);
  assert.equal(R.invoices.rows, 4); assert.equal(R.invoices.unique, 2); assert.equal(R.invoices.multiItem, 2); assert.equal(R.invoices.mixedStatus, 1); assert.equal(R.invoices.refundLike, 1);
});
