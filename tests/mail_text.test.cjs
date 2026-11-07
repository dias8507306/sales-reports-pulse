// Рассылка паспорты: медиа сілтемелері (тек https, ең көбі 4) және мәтін тазалау (2000 таңба, формула емес).
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const g = loadGs().exp;

test('zopMediaList_: тек https, бөлгіштер, 4-тен аспайды', () => {
  assert.deepEqual([...g.zopMediaList_('https://a.kz/1.jpg, http://b.kz/x.png\nhttps://drive.google.com/file/d/abcdefghijk/view javascript:alert(1)')], ['https://a.kz/1.jpg', 'https://drive.google.com/file/d/abcdefghijk/view']);
  assert.equal(g.zopMediaList_('https://a.kz/1 https://a.kz/2 https://a.kz/3 https://a.kz/4 https://a.kz/5').length, 4);
  assert.deepEqual([...g.zopMediaList_('')], []);
});

test('zopTextClean_: басқару таңбалары жоқ, формула болмайды, 2000 таңба', () => {
  assert.equal(g.zopTextClean_('=SUM(A1)'), "'=SUM(A1)");
  assert.equal(g.zopTextClean_('Сәлем\u0007 *әлем*'), 'Сәлем *әлем*');
  assert.equal(g.zopTextClean_('a'.repeat(2500)).length, 2000);
  assert.equal(g.zopTextClean_('жол 1\nжол 2'), 'жол 1\nжол 2');
});

test('zopBtnList_: «Мәтін» немесе «Мәтін | https://…», ең көбі 3', () => {
  assert.deepEqual(JSON.parse(JSON.stringify(g.zopBtnList_('Иә, қызықтырады\nСайтқа өту | https://eduapp.example.com\nhttp-сілтеме | http://x.kz\nТөртінші'))),
    [{t: 'Иә, қызықтырады'}, {t: 'Сайтқа өту', u: 'https://eduapp.example.com'}, {t: 'http-сілтеме'}]);
  assert.deepEqual([...g.zopBtnList_('')], []);
});
