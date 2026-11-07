// F02, F15, F14, F09: регрессионные тесты расчётного слоя Atlas на фиксированном срезе 09.10.2026 11:00.
// Контрольные числа относятся к этому срезу, а не к live-данным.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadAtlas, fixture} = require('./lib/load.cjs');
const SNAP = fixture('atlas_demo.json');

test('F02: «Рассылка (екеуі)» = Рассылка + Рассылка 5–6 во всех блоках', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = loadAtlas(SNAP);
  const b = f => A.buyGrid({f, src: -1}, 'buy').tot, h = f => A.buyGrid({f, src: -1}, 'lead').tot;
  assert.equal(b(1), 25); assert.equal(b(2), 7);
  assert.equal(b(-2), b(1) + b(2), 'buy heatmap: -2 must be 1+2 (was 102 = all funnels)');
  assert.equal(h(-2), h(1) + h(2), 'lead heatmap: -2 must be 1+2');
  const p = f => A.profile({f, of: -1, src: -1}, 2, k => A.D.dict.region[k]).reduce((s, r) => s + r.n, 0);   // translated dictionary, as on the page
  assert.equal(p(-2), p(1) + p(2)); assert.equal(p(-2), 2349);
  const st = f => A.stats(0, A.END, {f});
  for (const k of ['new', 'won', 'rev', 'lost']) assert.equal(st(-2)[k], st(1)[k] + st(2)[k], 'stats.' + k);
});

test('F02: «Барлық маркетинг» = ОП + обе рассылки; техподдержка не попадает', () => {
  const A = loadAtlas(SNAP);
  const b = f => A.buyGrid({f, src: -1}, 'buy').tot;
  assert.equal(b(-1), b(0) + b(1) + b(2));
  assert.equal(A.inFunnel(A.SUP, -1), false); assert.equal(A.inFunnel(A.SUP, A.SUP), false, 'support is never commercial, even if asked directly');
  const st = f => A.stats(0, A.END, {f});
  assert.equal(st(-1).won, st(0).won + st(1).won + st(2).won);
});

test('F15: повторный initData того же среза даёт тот же порядок продуктов', () => {
  const A = loadAtlas(SNAP), first = A.PRODUCTS.slice();
  A.initData(JSON.parse(JSON.stringify(SNAP)));
  assert.deepEqual([...A.PRODUCTS], [...first]);
  // the order is by successful deals: the first product must really have the most wins
  const won = {}; A.ROWS.forEach(r => { if (r.f !== A.SUP && r.o === A.OK && r.pr >= 0) String(SNAP.dict.product[r.pr]).split(/[,|]/).map(x => x.trim()).filter(Boolean).forEach(x => { won[x] = (won[x] || 0) + r.n; }); });
  const max = Math.max(...Object.values(won)); assert.equal(won[first[0]], max);
});

test('F14: KPI находится по области и всем измерениям, а не по устаревшему F.f', () => {   // 14.10.2026: Маркетинг — тек ОП воронкасы (f: 0); 15.10: Маркетинг → «Лидтер және маркетинг» (leads)
  const A = loadAtlas(SNAP), P = {a: 0, b: A.END};
  const month = '2026-10', base = {area: 'Рассылка', metric: 'Сатылым', month, target: 30};
  A.KPIS = A.kpisFrom([base, {area: 'Маркетинг', metric: 'Лидтер', month, target: 5000, src: 'Рассылка'}]);
  const per = A.KPIS[0].per;   // m:2026-10
  const qMail = {f: -2, src: -1, of: -1, pr: null, m: -1};
  const P10 = (() => { A.F.p = per; return A.period(); })();
  assert.ok(A.kpiFor('won', 'mail', qMail, P10), 'mail KPI must match the mail area query');
  assert.equal(A.kpiFor('won', 'mkt', {f: 0, src: -1, of: -1, pr: null, m: -1}, P10), null, 'not in another area');
  assert.equal(A.kpiFor('won', 'mail', Object.assign({}, qMail, {src: 2}), P10), null, 'not when an extra filter is active');
  const srcI = A.SRC.indexOf('Рассылка');
  assert.ok(A.kpiFor('new', 'leads', {f: 0, src: srcI, of: -1, pr: null, m: -1}, P10), 'KPI with a source filter matches the same source');
  assert.equal(A.kpiFor('new', 'leads', {f: 0, src: -1, of: -1, pr: null, m: -1}, P10), null, 'and not the unfiltered tile');
  void P;
});

test('F09/F10: предыдущий период вне покрытия → нет базы, без пропорционального пересчёта', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = loadAtlas(SNAP);
  assert.equal(A.COVER.a, 48, 'first commercial deal in this snapshot: day 48 = 27.09.2026');
  A.F.p = '30'; A.F.cmp = true; const P = A.period();
  assert.equal(P.prev, null); assert.match(P.noBase, /салыстыру жоқ/);
  const t = A.tileData('won', P, {f: -1, src: -1, of: -1, pr: null, m: -1});
  assert.equal(t.prev, null, 'no fake zero / scaled previous value');
  A.F.p = '7'; assert.equal(A.period().prev, null, 'only 13 days of history: a 7-day window has no full previous window');
  A.F.p = 'd:' + A.END; const P1 = A.period();   // one day: the day before is inside coverage → a real equal window
  assert.ok(P1.prev && P1.prev.a === A.END - 1 && P1.prev.b === A.END - 1); assert.equal(P1.prev.scale, undefined);
});

test('F07/F08: блоки называют фильтры, которые не могут применить', () => {
  const A = loadAtlas(SNAP);
  assert.deepEqual([...A.notApplied('buy', {src: 'Рассылка', of: 'Предпродажа', pr: 'X', m: null})], ['акция', 'өнім']);
  assert.deepEqual([...A.notApplied('aud', {src: 'Рассылка', of: 'Предпродажа', pr: null, m: 'Y'})], ['менеджер']);
  assert.deepEqual([...A.notApplied('reasons', {src: 'Рассылка', of: null, pr: null, m: null})], ['арна']);
});

test('Пустой срез не ломает initData', () => {
  const empty = JSON.parse(JSON.stringify(SNAP)); empty.main = []; empty.buy = []; empty.hours = []; empty.prof = [];
  const A = loadAtlas(empty);
  assert.equal(A.END, 0); assert.equal(A.stats(0, 0, {f: -1}).new, 0);
});

test('F13: статус источников — срез, ескірген live, пустые листы, нечитаемый бюджет', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = loadAtlas(SNAP), d = A.D, t = A.localMs(d.snapshot);
  assert.equal(t, Date.UTC(2026, 9, 9, 6, 0), '11:00 in Qyzylorda = 06:00 UTC');
  const st = (mode, h) => A.sourceStatus(d, {mode, error: mode === 'error' ? 'Google Sheets: HTTP 500' : null}, t + h * 36e5).find(x => x.id === 'crm');
  assert.equal(st('live', 1).state, 'ok'); assert.equal(st('live', 5).state, 'stale');
  assert.equal(st('embedded', 0).state, 'embedded'); assert.equal(st('error', 0).state, 'error');
  const empty = Object.assign({}, d, {mailings: [], costs: [], quality: {badBudget: 3}});
  const L = A.sourceStatus(empty, {mode: 'live'}, t);
  assert.equal(L.find(x => x.id === 'mail').state, 'missing'); assert.equal(L.find(x => x.id === 'cost').state, 'missing');
  assert.match(L.find(x => x.id === 'budget').note, /3 мәміледе/);
});
