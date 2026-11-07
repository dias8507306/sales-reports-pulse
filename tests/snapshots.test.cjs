// Срездер (иесі 15.10.2026): amoCRM тек қазіргі күйді береді — күн сайын лист «Срездер»-ге атаулармен жазамыз.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const g = loadGs().exp;

test('zopSnapOf_: атаулармен, ашық сделкалар этап бойынша, қысқаша сандар', () => {
  const out = {snapshot: '2026-10-14T23:50', paymatch: {pays: 10, matched: 1},
    dict: {manager: ['Айгерім', 'Мадина'], funnel: ['ОП'], stage: ['ОП|НОВАЯ ЗАЯВКА', 'ОП|ПРОВЕДЕННЫЕ КОНСУЛЬТАЦИИ'], outcome: ['В работе', 'Успешно', 'Не реализовано']},
    idle: [[0, 0, 0, 5, 2, 3], [1, 0, 1, 1, 0, 4]],
    main: [[10, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 3, 0], [10, 0, 1, 0, 1, 0, -1, 0, 0, 0, 0, 0, 4, 0], [10, 0, 1, 1, 1, 0, 12, 0, 0, 0, 0, 0, 9, 0]]};
  const s = g.zopSnapOf_(out);
  assert.equal(s.d, '2026-10-14'); assert.equal(s.t, '23:50');
  assert.deepEqual([...s.idle[0]], ['Айгерім', 'ОП', 'ОП|НОВАЯ ЗАЯВКА', 5, 2, 3]);
  assert.equal(s.sum.open, 7, 'тек «В работе»'); assert.equal(s.sum.late, 3); assert.equal(s.sum.idle3, 3);
  assert.ok(!JSON.stringify(s).match(/\d{10}/), 'телефон/ID жоқ');
});

test('zopSnapCell_: ұяшыққа сыяды (50 000 таңбадан аз), үлкен жолдар қалады', () => {
  const rows = Array.from({length: 4000}, (_, i) => ['Менеджер ' + i, 'ОП', 'ОП|ЭТАП ' + i, i % 30, 0, i]);
  const j = g.zopSnapCell_(rows, 5); assert.ok(j.length <= 49000);
  const kept = JSON.parse(j); assert.equal(kept[0][5], 3999, 'ең үлкені сақталды');
});
