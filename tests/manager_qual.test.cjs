// Квал по менеджерам ОП: квал ⊆ өтінімдер, сатылым ⊆ квал, сумма по менеджерам = итог; правило этапов.
'use strict';
const test = require('node:test'), assert = require('node:assert');
const {fixture, loadAtlas} = require('./lib/load.cjs');
const SNAP = fixture('atlas_demo.json');

test('Квал: подмножества и сумма по менеджерам совпадает с итогом', () => {
  const A = loadAtlas(SNAP), all = A.qualStats(0, 1e6, {f: 0, m: -1});
  assert.ok(all.leads > 0, 'есть сделки ОП');
  assert.ok(all.qual <= all.leads && all.won <= all.qual, 'won ≤ qual ≤ leads');
  let L = 0, Q = 0, W = 0;
  for (let m = 0; m < 50; m++) { const t = A.qualStats(0, 1e6, {f: 0, m}); L += t.leads; Q += t.qual; W += t.won; }
  assert.deepEqual([L, Q, W], [all.leads, all.qual, all.won]);
  assert.equal(all.wonOfQual, all.won / all.qual);
});

test('Квал: пустой период — проценты «нет данных», не 0', () => {
  const A = loadAtlas(SNAP), t = A.qualStats(-10, -5, {f: 0, m: -1});
  assert.deepEqual([t.leads, t.qual, t.won, t.qualPct, t.wonOfQual], [0, 0, 0, null, null]);
});
