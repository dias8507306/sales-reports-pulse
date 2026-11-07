// Этап 2–3: реестр метрик, платежи, качество заполнения, состояние фильтров в URL, снимки.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadAtlas, loadGs, cardCalc, fixture} = require('./lib/load.cjs');
const SNAP = fixture('atlas_demo.json'), CARD = fixture('card_sales_demo.json');

test('Реестр: у каждой метрики страницы есть полное описание', () => {
  const A = loadAtlas(SNAP), need = ['id', 'labelKK', 'labelRU', 'unit', 'source', 'grain', 'dateBasis', 'formula', 'supportedDimensions', 'nullPolicy', 'polarity', 'freshness', 'methodVersion'];
  for (const k of Object.keys(A.METRICS)) assert.ok(A.metricInfo(k), 'registry entry for ' + k);
  for (const m of A.METRIC_REG) for (const f of need) assert.ok(m[f] != null && m[f] !== '', m.id + '.' + f);
  assert.equal(new Set(A.METRIC_REG.map(m => m.id)).size, A.METRIC_REG.length, 'ids are unique');
  assert.match(A.metricInfo('rev').formula, /не подтверждённая оплата/);
});

test('Статус оплаты одинаков в Атласе, карточке и Apps Script', () => {
  const A = loadAtlas(SNAP), gs = loadGs(), C = cardCalc(gs, CARD);
  for (const v of ['', '2', 'paid', 'unpaid', 'Unpaid', 'pending', 'refunded', 'сәтті', 'Mon Jan 01 1900', '3', null, 2]) {
    assert.equal(A.payStatus(v), gs.exp.zopPayStatus_(v), 'atlas vs server ' + v);
    assert.equal(C.payStatus(v), gs.exp.zopPayStatus_(v), 'card vs server ' + v);
  }
});

test('Сатылым (15.10): ОП кестесі алынды — «s» = CRM «Успешно» бюджеті; Атлас пен карточка бір күнге бірдей', () => {
  const J = x => JSON.parse(JSON.stringify(x)), A = loadAtlas(SNAP); A.initSales(J(CARD.data));
  const day = '2026-10-08', i = Math.round((Date.UTC(2026, 9, 8) - Date.UTC(2026, 7, 10)) / 864e5);
  const crm = A.stats(i, i, {f: -1}), P = A.payStats(i, i);
  let won = 0, bud = 0; SNAP.main.forEach(r => { const L = r.length; if (r[1] !== SNAP.dict.funnel.indexOf('Техподдержка') && r[3] === SNAP.dict.outcome.indexOf('Успешно') && r[6] === i) { won += r[L - 2]; bud += Math.round(r[L - 1]); } });
  assert.equal(P.sheet.n, won, 'manager sales = CRM won deals closed that day'); assert.equal(P.sheet.s, bud, 'their CRM budget');
  const R = cardCalc(loadGs(), {kind: 'sales', date: day, data: Object.assign(J(CARD.data), {crm: J(SNAP)})}).sales(day);
  assert.equal(R.by.s.n, P.sheet.n); assert.equal(R.by.s.s, P.sheet.s); assert.equal(R.by.p.s, P.plat.s, 'platform unchanged');
  assert.ok(crm.won >= 0);
  A.initSales(null); assert.equal(A.payStats(i, i), null, 'no sales data → no numbers, not zeros');
});

test('Качество заполнения: старый срез без fill — бюджет «нет данных», остальные доли считаются', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = loadAtlas(SNAP), Q = A.fillQuality(0, A.END), S = A.stats(0, A.END, {f: -1});
  assert.equal(Q.hasFill, false); assert.equal(Q.total.fillBudget, null);
  assert.equal(Q.total.won, S.won); assert.equal(Q.total.lost, S.lost);
  assert.ok(Q.total.fillProduct >= 0 && Q.total.fillProduct <= 1);
  // с кубом fill бюджет появляется
  const withFill = JSON.parse(JSON.stringify(SNAP)), OK = SNAP.dict.outcome.indexOf('Успешно');
  withFill.fill = [[A.END, 0, 0, OK, 1, 1, 1, 0, 3, 30000], [A.END, 0, 0, OK, 0, 1, 1, 0, 1, 0]];
  const B = loadAtlas(withFill), Q2 = B.fillQuality(B.END, B.END);
  assert.equal(Q2.hasFill, true); assert.equal(Q2.total.won, 4); assert.equal(Q2.total.fillBudget, 0.75);
});

test('URL-состояние: туда-обратно без имён в адресе', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = loadAtlas(SNAP), man = SNAP.dict.manager[0], src = 'Рассылка';
  const F = {p: 'd:' + A.END, cmp: false, fBy: {op: 0}, src, of: null, pr: null, m: man};
  const s = A.encodeState(F, 'op');
  assert.ok(!s.includes(encodeURIComponent(man)) && !s.includes(man), 'no manager name in URL');
  assert.ok(!/Рассылка/.test(decodeURIComponent(s)));
  assert.match(s, /a=2026-10-09/);
  const back = A.decodeState(s, 'op', {src: A.SRC, of: [], pr: [], m: SNAP.dict.manager});
  assert.equal(back.p, F.p); assert.equal(back.cmp, false); assert.equal(back.src, src); assert.equal(back.m, man); assert.deepEqual({...back.fBy}, {op: 0});
  assert.deepEqual({...A.decodeState('p=zzz&x=1&m=unknown', 'op', {m: SNAP.dict.manager})}, {m: null}, 'garbage is ignored');
});

test('Множественный выбор: ИЛИ внутри справочника, И между справочниками; ссылка сохраняет выбор', () => {
  const A = loadAtlas(SNAP), src = A.SRC.filter(Boolean).slice(0, 2), products = A.PRODUCTS.slice(0, 2), man = SNAP.dict.manager.filter(Boolean).slice(0, 2);
  assert.equal(src.length, 2); assert.equal(products.length, 2); assert.equal(man.length, 2);
  const q = {f: -1, src: src.map(x => A.SRC.indexOf(x)), of: -1, pr: products, m: -1};
  for (const r of A.ROWS) assert.equal(A.pass(r, q), src.some(s => products.some(p => A.pass(r, {...q, src: A.SRC.indexOf(s), pr: p}))), 'row union');
  const total = A.stats(0, A.END, {...q, pr: null}).new;
  const parts = src.reduce((n, s) => n + A.stats(0, A.END, {...q, src: A.SRC.indexOf(s), pr: null}).new, 0);
  assert.equal(total, parts, 'each source contributes once');
  const b = A.buyGrid({...q, pr: null}, 'buy').tot;
  const bParts = src.reduce((n, s) => n + A.buyGrid({...q, src: A.SRC.indexOf(s), pr: null}, 'buy').tot, 0);
  assert.equal(b, bParts, 'the purchase-hour cube respects multiple sources');
  const F = {p: '30', cmp: true, fBy: {mkt: -1}, src, of: null, pr: products, m: man};
  const url = A.encodeState(F, 'mkt');
  for (const name of [...src, ...products, ...man]) assert.ok(!url.includes(name), 'names stay out of URL');
  const back = A.decodeState(url, 'mkt', {src: A.SRC, of: [], pr: A.PRODUCTS, m: SNAP.dict.manager});
  assert.deepEqual([...back.src], [...src]); assert.deepEqual([...back.pr], [...products]); assert.deepEqual([...back.m], [...man]);
});

test('Снимок: одинаковые данные → одинаковый snapshotId; другие данные → другой', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const A = loadAtlas(SNAP);
  const a = A.snapshotMeta(SNAP, CARD.data), b = A.snapshotMeta(JSON.parse(JSON.stringify(SNAP)), CARD.data);
  assert.equal(a.snapshotId, b.snapshotId); assert.equal(a.methodVersion, 'atlas-2026-10-10');
  assert.equal(a.sources.crm.at, '2026-10-09T11:00');
  const c = A.snapshotMeta(Object.assign({}, SNAP, {rows: SNAP.rows + 1}), CARD.data);
  assert.notEqual(a.snapshotId, c.snapshotId);
});

test('CSV: BOM, «;», кавычки, html-ячейки как текст', () => {
  const A = loadAtlas(SNAP);
  const csv = A.csvFrom([{title: 'Т; 1', head: ['A', 'B'], rows: [{c: ['x"y', {html: '<span class="st ok">90%</span>'}]}]}], [['Кезең', '30 күн']]);
  assert.ok(csv.startsWith('﻿')); assert.match(csv, /Кезең;30 күн/);
  assert.match(csv, /"Т; 1"/); assert.match(csv, /"x""y";90%/); assert.ok(!/<span/.test(csv));
});

test('WhatsApp: без «кеше/бүгін» и стрелок; CRM-бюджет помечен как не оплата', () => {
  const A = loadAtlas(SNAP);
  const t = A.waSummary({title: 'Шолу', period: 'соңғы 30 күн (10 қыркүйек – 9 қазан)', lines: [['Лидтер', '8 581'], ['CRM бюджеті', '1 900 000 ₸', 'менеджер жазған сома, төлем емес']], sources: 'amoCRM 2026-10-09 11:00', snapshotId: 'abc'});
  assert.ok(!/кеше|бүгін|[↑↓▲▼]/i.test(t)); assert.match(t, /төлем емес/); assert.match(t, /Срез: abc/);
});

test('Лиды: входящие отдельно от созданных рассылкой; сумма видов = все созданные', () => {
  const {loadAtlas: LA} = require('./lib/load.cjs'), A = LA(SNAP), q = {f: -1, src: -1, of: -1, pr: null, m: -1};
  const k = A.leadKinds(0, A.END, q), S = A.stats(0, A.END, q);
  assert.equal(k.in + k.campOp + k.camp, S.new);
  assert.ok(k.camp > 0 && k.in > 0);
  const opOnly = A.leadKinds(0, A.END, Object.assign({}, q, {f: 0}));
  assert.equal(opOnly.camp, 0, 'ОП has no mailing-funnel deals');
  const mail = A.leadKinds(0, A.END, Object.assign({}, q, {f: -2}));
  // 14.10.2026: түрі — «Қайдан келді» бойынша: рассылка воронкасында тек нақты арнасы (Таргет, WhatsApp…) бар сделка кіріс болады
  const realSrc = A.ROWS.filter(r => (r.f === 1 || r.f === 2) && !['Рассылка', 'Көрсетілмеген'].includes(A.SRC[r.src])).reduce((t, r) => t + r.n, 0);
  assert.equal(mail.campOp, 0, 'no ОП deals in mailing funnels');
  assert.equal(mail.in, realSrc, 'mailing-funnel deals count as inbound only by a real «Қайдан келді»');
});
