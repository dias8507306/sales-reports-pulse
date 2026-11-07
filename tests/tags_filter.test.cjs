// Фильтр по тегам amoCRM: 11-й ключ main — набор тегов; сделка проходит, если у неё есть хотя бы один выбранный тег.
'use strict';
const test = require('node:test'), assert = require('node:assert');
const {fixture, loadAtlas} = require('./lib/load.cjs');

test('Тег: TAGS из набора, фильтр ИЛИ внутри тегов, без тегов — все', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const d = JSON.parse(JSON.stringify(fixture('atlas_demo.json')));
  d.dict.tagset = ['', 'Алматы облысы|5 сынып', '5 сынып'];
  d.main = d.main.map((r, i) => r.length >= 12 ? r.slice(0, 10).concat([i % 3], r.slice(10)) : r);
  const A = loadAtlas(d), q = {f: -1, src: -1, of: -1, pr: null, m: -1};
  assert.deepEqual([...A.TAGS].sort(), ['5 сынып', 'Алматы облысы']);
  const all = A.stats(0, A.END, q).new, t5 = A.stats(0, A.END, Object.assign({}, q, {tg: '5 сынып'})).new, ta = A.stats(0, A.END, Object.assign({}, q, {tg: 'Алматы облысы'})).new;
  const both = A.stats(0, A.END, Object.assign({}, q, {tg: ['5 сынып', 'Алматы облысы']})).new;
  assert.ok(t5 > 0 && ta > 0 && ta < t5 && t5 < all);
  assert.equal(both, t5, 'ИЛИ: «Алматы облысы» всегда вместе с «5 сынып»');
});
