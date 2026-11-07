// ОП кестесі ↔ amoCRM: төлем сделкаға телефон («Доступ номер») не «Чек» арқылы байланады; сыртқа тек белгі мен сандар шығады.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const g = loadGs().exp;
const D = s => new Date(s + 'T10:00:00Z');

test('zopPayMatch_: чек бірінші, сосын телефон; төлемге дейінгі ең соңғы сделка', () => {
  const leads = [['ID', 'Создана', 'Доступ номер', 'Чек'],
    [1, D('2026-09-01'), '+7 700 000 11 22', ''],
    [2, D('2026-10-01'), '87000001122', ''],          // сол телефон, кейінірек — төлем осыған
    [3, D('2026-10-02'), '', 'A-1234'],
    [4, D('2026-10-20'), '87000009988', '']];         // төлемнен кейін құрылған — байланбайды
  const pays = [['Дата', 'Телефон нөмірі', 'Чек нөмірі', 'Төлем сомасы'],
    ['05.10.2026', '7000001122', '', 15000],
    [D('2026-10-03'), '', 'a-1234 ', 15000],
    ['10.10.2026', '+7 (700) 000-99-88', '', 15000],
    ['11.10.2026', '', '', 15000],
    ['', '7000001122', '', 15000]];
  const r = g.zopPayMatch_(leads, pays);
  assert.deepEqual([...r.paid].map(x => x[0]).sort(), [2, 3]);
  const {cols, crmPhones, crmChecks, payPhones, payChecks, paysCov, matchedCov, ...st} = r.stats;
  assert.equal(paysCov, 4); assert.equal(matchedCov, 2);
  assert.deepEqual({...st}, {pays: 4, matched: 2, byPhone: 1, byCheck: 1, noKey: 1});
  assert.equal(cols.phone, 'телефон нөмірі'); assert.equal(cols.check, 'чек нөмірі'); assert.equal(crmPhones, 3); assert.equal(payPhones, 2);
  const sums = Object.fromEntries([...r.sum].map(([row, v]) => [row[0], v]));
  assert.deepEqual(sums, {2: 15000, 3: 15000}, 'сома сделкаға жазылады (рассылка тегі бойынша «1 000-ға» үшін)');
});

test('zopPayMatch_: ОП кестесі жоқ — бос нәтиже, қате жоқ', () => {
  const r = g.zopPayMatch_([['ID', 'Создана'], [1, D('2026-10-01')]], []);
  assert.equal(r.paid.size, 0); assert.equal(r.stats.pays, 0);
});

test('zopPayMatch_: чек жазылуы әр түрлі («№ 123456», «QR-123456») және сделка төлемнен кейін ашылса (7 күнге дейін)', () => {
  const leads = [['ID', 'Создана', 'Доступ номер', 'Чек'], [1, D('2026-10-05'), '', '№ 4455667'], [2, D('2026-10-12'), '', 'QR-998877']];
  const pays = [['Дата', 'Чек номер', 'Төлем сомасы'], ['05.10.2026', '4455667', 15000], ['09.10.2026', 'чек 998877', 990]];
  const r = g.zopPayMatch_(leads, pays);
  assert.deepEqual([...r.paid].map(x => x[0]).sort(), [1, 2]);
});
