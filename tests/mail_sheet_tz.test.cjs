// ТЗ 15.10: «Рассылки» парағының жаңа бағандары (Формат, Нөмір, Гипотеза) және «Гипотезалар» журналы.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const g = loadGs().exp;
const parts = d => ({y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate()}), dayIdx = p => Math.round((Date.UTC(p.y, p.m - 1, p.d) - Date.UTC(2026, 7, 10)) / 864e5);

test('zopMailings_: формат, нөмір (5+ цифр жасырылады), гипотеза', () => {
  const rows = [['Рассылка', 'Дата', 'Воронка', 'Формат', 'Нөмір', 'Гипотеза'], ['11.10 · Алтын сақа', '11.10.2026', 'Рассылка', 'Видео', '7000 / 77000001234', 'Күндіз кешке қарағанда нашар емес']];
  const [m] = g.zopMailings_(rows, parts, dayIdx, ['ОП', 'Рассылка', 'Рассылка 5-6 класс Алтын сақа']);
  assert.equal(m.fmt, 'Видео'); assert.equal(m.hyp, 'Күндіз кешке қарағанда нашар емес');
  assert.ok(!/\d{5,}/.test(m.num), 'телефон нөмірі шықпайды'); assert.ok(m.num.startsWith('7000'));
});

test('zopHypotheses_: «Гипотеза» мен «Рассылкалар» бағандары оқылады', () => {
  const rows = [['Гипотеза', 'Эксперимент', 'Рассылкалар', 'Нәтиже', 'Шешім'], ['Днём продаёт не хуже вечера', 'бір мәтін, әр уақыт', '11.10 12:02', '', '']];
  const [h] = g.zopHypotheses_(rows);
  assert.equal(h['Гипотеза'], 'Днём продаёт не хуже вечера'); assert.equal(h['Рассылкалар'], '11.10 12:02');
});
