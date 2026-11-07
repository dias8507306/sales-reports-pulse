// «Өнімдер»: сначала раздел (A), потом продукт (B) — список B содержит только продукты раздела.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs, fixture} = require('./lib/load.cjs');
const g = loadGs().exp;

test('Раздел → направления продуктов', () => {
  assert.deepEqual([...g.zopSecFamilies_('Алтын сақа')], ['Алтын сақа']);
  assert.deepEqual([...g.zopSecFamilies_('Алтын сақа · пробный тест')], ['Алтын сақа · пробный тест']);
  assert.deepEqual([...g.zopSecFamilies_('Алтын түлек')], ['Алтын түлек']);
  assert.deepEqual([...g.zopSecFamilies_('Негізгі өнімдер')], ['Математика · Мега қайталау', 'Математика · Тоқсандық', 'Ағылшын']);
  assert.equal(g.zopSecGuess_('Мега Қайталау', 'p'), 'Математика · Мега қайталау'); assert.equal(g.zopSecGuess_('3-сынып. I-IV тоқсан. · Күндік', 'p'), 'Математика · Мега қайталау');
  assert.equal(g.zopSecGuess_('4-сынып ІІІ тоқсан · Basic', 'p'), 'Математика · Тоқсандық'); assert.equal(g.zopSecGuess_('3-сынып (I-II тоқсан) тренинг математика', 'p'), 'Математика · Тоқсандық');
  assert.equal(g.zopSecFamilies_('Жаңа бөлім'), null, 'unknown section → all products');
});

test('«1 тренинг» не перепутать: Алтын сақа и Алтын түлек — разные списки', () => {
  const plist = [['Алтын Сақа тренинг 1', 'Алтын сақа'], ['«Алтын түлек» дайындық нұсқасы 1', 'Алтын түлек'], ['3-сынып. I-IV тоқсан. · Күндік', 'Математика · Мега қайталау'], ['English A0 · Күндік', 'Ағылшын']];
  const L = g.zopSectionLists_(plist, ['Негізгі өнімдер', 'Алтын сақа', 'Алтын түлек', 'Жаңа бөлім']), col = s => [...L.cols[L.secs.indexOf(s)]];
  assert.deepEqual(col('Алтын сақа'), ['Алтын Сақа тренинг 1']);
  assert.deepEqual(col('Алтын түлек'), ['«Алтын түлек» дайындық нұсқасы 1']);
  assert.deepEqual(col('Негізгі өнімдер'), ['3-сынып. I-IV тоқсан. · Күндік', 'English A0 · Күндік']);
  assert.equal(col('Жаңа бөлім').length, 4); assert.equal(L.secs[L.secs.length - 1], 'Барлық өнімдер');
  assert.equal(col('NIS').length, 4, 'a section with no products of its kind shows all, never an empty list');
});

test('Реальные названия платформы: все попадают хотя бы в один раздел', () => {
  const list = g.zopProductList_(fixture('card_sales_demo.json').data), plist = list.map(r => [r[0], g.zopSecGuess_(r[0], 'p')]);
  const L = g.zopSectionLists_(plist, ['Алтын сақа', 'Алтын түлек']);
  assert.equal(L.cols[L.secs.length - 1].length, list.length);
  assert.ok(L.cols[L.secs.indexOf('Алтын түлек')].every(n => /түлек/i.test(n)));
  assert.ok(L.cols[L.secs.indexOf('Алтын сақа')].every(n => /сақа|сака/i.test(n) && !/пробный тест/i.test(n)));
});
