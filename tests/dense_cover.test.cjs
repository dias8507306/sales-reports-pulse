// ТЗ 15.10: алдыңғы кезеңмен салыстыру тек деректер толық болған күннен («+302%» — синхрондау басталған апта).
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadAtlas} = require('./lib/load.cjs');
const A = loadAtlas();

test('denseStart: сирек ескі сделкалар салыстыру базасы емес', () => {
  const rows = [];
  for (let d = 0; d < 20; d++) if (d % 5 === 0) rows.push({c: d, n: 1});   // синхрондауға дейін — анда-санда бір сделка
  for (let d = 20; d <= 60; d++) rows.push({c: d, n: 100});
  assert.equal(A.denseStart(rows, 60, 0), 20);
});

test('denseStart: деректер басынан толық — бірінші күн', () => {
  const rows = []; for (let d = 0; d <= 60; d++) rows.push({c: d, n: 50 + (d % 7) * 5});
  assert.equal(A.denseStart(rows, 60, 0), 0);
});
