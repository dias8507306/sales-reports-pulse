// Владелец 15.10.2026: «Динара — это Айгерім» — два аккаунта amoCRM одного человека считаются одним менеджером.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadAtlas, loadGs} = require('./lib/load.cjs');
const A = loadAtlas();

test('mergeManagers: индексы во всех кубах переписаны, имя одно', () => {
  const d = {dict: {manager: ['Dinara', 'Мадина', 'Айгерім']}, main: [[0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 2, 0], [0, 0, 0, 0, 1, 0]], idle: [[0, 0, 0, 1, 0, 5]], fill: [[0, 0, 0, 0]], miss: [[123, 5, 0, 0, 1]], imp: [[0, 0, 0, 0, 0, 0]]};
  A.mergeManagers(d, {Dinara: 'Айгерім'});
  assert.deepEqual([...d.dict.manager], ['Айгерім', 'Мадина']);
  assert.deepEqual(d.main.map(r => r[4]), [0, 0, 1]);
  assert.equal(d.idle[0][0], 0); assert.equal(d.miss[0][2], 0); assert.equal(d.imp[0][5], 0);
});

test('Apps Script и Python знают тот же псевдоним', () => {
  const g = loadGs(); assert.equal(g.value('ZOP_MAN_ALIAS').Dinara, 'Айгерім');
  const py = require('fs').readFileSync(require('path').join(__dirname, '..', 'инструменты', 'обновить_данные.py'), 'utf8');
  assert.match(py, /MAN_ALIAS = \{'Dinara': 'Айгерім'\}/);
});
