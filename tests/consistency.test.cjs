// IA 15.10: одно слово — одна цифра. Один и тот же показатель в разных разделах обязан совпадать.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadAtlas, fixture} = require('./lib/load.cjs');
const SNAP = fixture('atlas_demo.json');
const A = loadAtlas(SNAP), OPQ = {f: 0, src: -1, of: -1, pr: null, m: -1};
const P = {a: Math.max(0, A.END - 29), b: A.END};

test('кіріс өтінім: Маркетинг (mkChannels) = Лидтер (leadKinds, ОП)', () => {
  assert.equal(A.mkChannels(P.a, P.b).tot.leads, A.leadKinds(P.a, P.b, OPQ).in);
});

test('Таргет: лиды в Маркетинге = строка «Таргет» в таблице каналов Лидтер', () => {
  const mk = A.mkChannels(P.a, P.b).list.find(e => e.name === 'Таргет'), ch = A.channels(P, OPQ).find(r => r.name === 'Таргет');
  if (!mk && !ch) return; assert.equal(mk ? mk.leads : 0, ch ? ch.s.new : 0);
});

test('CRM Успешно: сумма по каналам = итог раздела (ОП, по дате закрытия)', () => {
  const tot = A.stats(P.a, P.b, OPQ).won, byCh = A.channels(P, OPQ).reduce((t, r) => t + r.s.won, 0);
  assert.equal(byCh, tot);
});

test('Рассылки: «Жіберілді» в плитке = сумма строк таблицы', () => {
  const ms = A.mailings().filter(m => m.a >= P.a && m.a <= P.b), sent = m => (m.leads || 0) + (m.toOp || 0);
  const rowsSum = ms.reduce((t, m) => t + m.rows.reduce((x, r) => x + r.n, 0), 0);
  assert.equal(ms.reduce((t, m) => t + sent(m), 0), rowsSum, 'сделки рассылки = отправлено');
});
