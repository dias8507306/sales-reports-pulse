// Дневной WhatsApp-отчёт по продажам — шаблон владельца 15.10.2026 (вторая версия): орындалуы · жоспар/факт · жалпы нәтиже · клиенттер қайдан келді · ескертпе.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs, cardCalc, fixture} = require('./lib/load.cjs');
const CARD = fixture('card_sales_demo.json'), gs = loadGs(), C = cardCalc(gs, CARD), R = C.sales(CARD.date), T = C.salesText(CARD.date, R);
const num = s => +String(s).replace(/[^\d]/g, '');

test('разделы в порядке шаблона', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const order = ['ЗЕРДЕЛІ APP — КҮНДЕЛІКТІ ЕСЕП', 'Күні:', 'Жалпы сома', 'ОРЫНДАЛУЫ', '1. САТЫЛЫМ — ЖОСПАР / ФАКТ', 'Барлық сатылым', '2. ЖАЛПЫ НӘТИЖЕ', '3. КЛИЕНТТЕР ҚАЙДАН КЕЛДІ?', '4. ЕСКЕРТПЕ', 'КҮН ҚОРЫТЫНДЫСЫ'];
  let at = -1; for (const k of order) { const i = T.indexOf(k); assert.ok(i > at, k); at = i; }
  assert.match(T, /📅 \*Күні:\* 08\.10\.2026/);
});

test('«Клиенттер қайдан келді» = все продажи: ОП кестесі по «Қайдан келді» + платформа', () => {
  const sec = T.slice(T.indexOf('КЛИЕНТТЕР ҚАЙДАН КЕЛДІ'), T.indexOf('📊 *Барлығы'));
  const n = sec.split('\n').filter(l => /^\S+ .+ — \d/.test(l) && !/^\s/.test(l)).reduce((t, l) => t + num(l.split(' — ').pop()), 0);
  assert.equal(n, R.rows);
  assert.match(sec, /Маркетинг \(органика\) — \d/); assert.match(sec, /Маркетинг \(таргет\) — \d/);
});

test('суммы: Барлық сатылым = Жалпы сома; группы складываются', () => {
  assert.equal(num((T.match(/Барлық сатылым: \*([\d  ]+) дана/) || [])[1]), R.rows);
  assert.equal(num((T.match(/Жалпы сома: ([\d  ]+) ₸/) || [])[1]), R.sum);
});

test('нет лидов; код без «//» (Apps Script HtmlService)', () => {
  assert.doesNotMatch(T, /лид/i);
  const src = gs.exp.zopCardClient_.toString(), i = src.indexOf('function salesText'), j = src.indexOf('/* ---------- drawing', i);
  assert.doesNotMatch(src.slice(i, j), /\/\//);
});
