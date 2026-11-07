// Карточка лидов: «Жауап күтіп тұрған лидтер» — по менеджерам, дни без изменений (куб idle).
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs, cardCalc, fixture} = require('./lib/load.cjs');
const gs = loadGs();
const base = fixture('card_leads_demo.json');
const withIdle = idle => { const cfg = JSON.parse(JSON.stringify(base)); cfg.data.idle = idle; return cfg; };
const st = name => base.data.dict.stage.indexOf(name), fu = name => base.data.dict.funnel.indexOf(name);

test('Без куба idle (старый Apps Script) блок не показывается и ничего не ломает', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const C = cardCalc(gs, base);
  assert.equal(C.waiting(), null);
  const t0 = C.leadsText(base.date, C.leads(base.date), ''); assert.match(t0, /3\. ТЕКСЕРУ КЕРЕК/); assert.match(t0, /Дерек жоқ/);
});

test('Считаются только этапы ожидания; дни раскладываются по корзинам; ОП-этапы 3+ дня — отдельно', () => {
  const OP = fu('ОП'), ML = fu('Рассылка'), SUP = fu('Техподдержка');
  const cfg = withIdle([
    [0, OP, st('ОП|НОВАЯ ЗАЯВКА'), 0, 0, 5], [0, OP, st('ОП|НОВАЯ ЗАЯВКА'), 2, 0, 3], [0, OP, st('ОП|НОВАЯ ЗАЯВКА'), 9, 2, 1],
    [1, ML, st('Рассылка|Жауап берді'), 4, 0, 2],
    [1, OP, st('ОП|НЕДОЗВОН'), 5, 0, 7], [1, OP, st('ОП|НЕДОЗВОН'), 1, 0, 9],   // only 3+ days counts as «қозғалыссыз»
    [1, ML, st('Рассылка|Рассылка 3 15%'), 20, 0, 100],                          // mass mailing stage — not waiting
    [0, SUP, st('Техподдержка|НОВОЕ ОБРАЩЕНИЕ'), 30, 0, 50]]);                   // support — not here
  const W = cardCalc(gs, cfg).waiting(), M = base.data.dict.manager;
  assert.equal(W.n, 11); assert.deepEqual([...W.b], [5, 3, 2, 1]); assert.equal(W.max, 9); assert.equal(W.stuck, 7);
  assert.equal(W.man[M[0]].n, 9); assert.equal(W.man[M[0]].max, 9); assert.equal(W.man[M[1]].n, 2); assert.equal(W.man[M[1]].stuck, 7);
  assert.equal(W.list[0][0], M[0], 'the manager with the oldest leads comes first');
  const t = cardCalc(gs, cfg).leadsText(base.date, cardCalc(gs, cfg).leads(base.date), '');
  // шаблон 15.10.2026: бір менеджер — бір жол, «Аты Тегі»
  const nm2 = n => { const w = n.split(/\s+/); return w.length > 2 ? w[1] + ' ' + w[0] : n; };
  assert.match(t, /«Жауап берді» — \*11\*/); assert.match(t, /3\+ күн жаңартылмаған — \*3\*/); assert.match(t, /7\+ күн жаңартылмаған — \*1\*/);
  assert.ok(t.includes(`👤 *${nm2(M[0])}* — ⏳ 3+ күн — 1`), 'manager line: old leads');
  assert.ok(t.includes(`👤 *${nm2(M[1])}* — ⏳ 3+ күн — 2`), 'second manager: 3–6 days');
  assert.match(t, /мерзімі өткен — 1/);
  assert.ok(!/БАСТЫ МІНДЕТ/.test(t));
  assert.ok(!/кеше|бүгін|[↑↓▲▼]/i.test(t));
});

test('Кіріс өтінімдер: карточка и Атлас считают одинаково', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const {loadAtlas} = require('./lib/load.cjs'), A = loadAtlas(base.data), C = cardCalc(gs, base), iso = base.date;
  const di = Math.round((Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) - Date.UTC(2026, 7, 10)) / 864e5);
  const R = C.leads(iso), k = A.leadKinds(di, di, {f: -1, src: -1, of: -1, pr: null, m: -1});
  assert.equal(R.newIn, k.in); assert.equal(R.newL, k.n);
  assert.ok(R.newIn < R.newL);
});
