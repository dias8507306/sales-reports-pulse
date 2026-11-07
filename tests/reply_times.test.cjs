// Жауап уақыты (иесі 15.10: «точное время»): тег бойынша жауап санының өсуі → уақыт; ескі жауаптар уақытсыз қалады. Рассылка × өңір кубы impreg.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const J = x => JSON.parse(JSON.stringify(x));

test('zopReplyDelta_: бірінші іске қосу — тек бастапқы мән; кейін — өскені; жаңа тег — бүкіл саны', () => {
  const g = loadGs().exp;
  const a = g.zopReplyDelta_(null, {'2026-10-14 18:25': 40}, '2026-10-15 09:00');
  assert.deepEqual(J(a.add), [], 'old replies have no known time');
  const b = g.zopReplyDelta_(a.state, {'2026-10-14 18:25': 43, '2026-10-15 10:00': 5}, '2026-10-15 10:10');
  assert.deepEqual(J(b.add), [['2026-10-14 18:25', '2026-10-15 10:10', 3], ['2026-10-15 10:00', '2026-10-15 10:10', 5]]);
  const c = g.zopReplyDelta_(b.state, {'2026-10-14 18:25': 42, '2026-10-15 10:00': 5}, '2026-10-15 10:20');
  assert.deepEqual(J(c.add), [], 'a decrease is not a reply');
  const d = g.zopReplyDelta_(c.state, {'2026-10-14 18:25': 44}, '2026-10-15 10:30');
  assert.deepEqual(J(d.add), [['2026-10-14 18:25', '2026-10-15 10:30', 1]], 'after a drop only growth above the max counts');
});

test('zopReplyCounts_ + zopReplyCube_: impgeo → тег бойынша жауап; парақ → [импорт, күн, сағат, минут, n]', () => {
  const g = loadGs().exp, out = {dict: {imp: ['2026-10-14 18:25', '2026-10-15 10:00']}, impgeo: [[0, 0, 0, 0, 100], [0, 0, 0, 1, 7], [0, 3, 0, 2, 2], [1, 0, 0, 1, 5]]};
  assert.deepEqual(J(g.zopReplyCounts_(out)), {'2026-10-14 18:25': 9, '2026-10-15 10:00': 5});
  const cube = g.zopReplyCube_([['2026-10-14 18:25', '2026-10-15 10:10', '3'], ['2026-10-14 18:25', '2026-10-15 10:10', '1'], ['белгісіз', '2026-10-15 10:10', '2'], ['2026-10-15 10:00', 'жоқ', '1']], out.dict.imp, '2026-08-10');
  assert.deepEqual(J(cube), [[0, 66, 10, 10, 4]]);
});

test('impreg: өңір × нәтиже × тікелей импорт, 3-тен аз ұяшықтың өңірі жасырылады', () => {
  const g = loadGs().exp, parts = d => { if (!(d instanceof Date)) return null; const x = new Date(d.getTime() + 5 * 3600e3); return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours(), mi: x.getUTCMinutes()}; };
  const H = ['Воронка', 'Этап', 'Итог', 'Ответственный', 'Создана', 'Изменена', 'Закрыта', 'Облыс', 'Класс', 'Язык обучения', 'Қайдан келді', 'Теги'];
  const row = (fu, obl, tag) => [fu, 'НОВАЯ ЗАЯВКА', '', 'М', new Date('2026-10-01T05:00:00Z'), new Date('2026-10-01T05:00:00Z'), '', obl, 5, 'Қазақша', 'Рассылка', tag];
  const T = 'импорт_14102026_1825';
  const leads = [H].concat(Array(5).fill(row('Рассылка', 'Павлодар облысы', T)), Array(3).fill(row('ОП', 'Павлодар', T)), [row('ОП', 'Қызылорда', T)]);
  const pipes = [['', 'Воронка', '', '', 'Этап', 'Порядок'], ['', 'ОП', '', '', 'НОВАЯ ЗАЯВКА', 1], ['', 'Рассылка', '', '', 'НОВАЯ ЗАЯВКА', 1]];
  const out = g.zopAggregate_(leads, pipes, [['Время'], [new Date('2026-10-02T05:00:00Z')]], parts, [], [], []);
  const R = out.dict.region, rows = J(out.impreg).map(r => [R[r[0]], r[1], r[2], r[3]]);
  assert.deepEqual(rows.sort(), [['Жасырылған (< 3)', 1, 0, 1], ['Павлодар', 0, 0, 5], ['Павлодар', 1, 0, 3]].sort());
});
