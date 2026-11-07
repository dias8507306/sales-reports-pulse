// Рассылка талдауы (иесі 17.10): воронка тізбекті, қосынды = жіберілді, салыстыру — тек ұқсас және жеткілікті болса, ауытқу — тек сенімді болса.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadAtlas, fixture} = require('./lib/load.cjs');
const SNAP = fixture('atlas_demo.json');
const J = x => JSON.parse(JSON.stringify(x));

function setup() {
  const s = J(SNAP), st = s.dict.stage, add = n => { let i = st.indexOf(n); if (i < 0) { st.push(n); i = st.length - 1; } return i; };
  const S = {opWait: add('ОП|рассылка'), opCons: add('ОП|ПРОВЕДЕННЫЕ КОНСУЛЬТАЦИИ'), opNd: add('ОП|НЕДОЗВОН'), mWait: add('Рассылка|Рассылка 4 кз'), mRep: add('Рассылка|Жауап берді'), opWon: add('ОП|Успешно реализовано'), opLost: add('ОП|Закрыто и не реализовано')};
  const A = loadAtlas(s), O = s.dict.outcome, OK = O.indexOf('Успешно'), LOST = O.indexOf('Не реализовано'), OPEN = O.indexOf('В работе');
  return {A, S, OK, LOST, OPEN};
}
const mk = (rows, extra) => Object.assign({imp: null, a: 50, time: '18:00', f: 1, rows, leads: rows.filter(r => r.f !== 0).reduce((t, r) => t + r.n, 0), toOp: rows.filter(r => r.f === 0).reduce((t, r) => t + r.n, 0),
  repIn: rows.filter(r => r.f !== 0 && r.rep).reduce((t, r) => t + r.n, 0), won: 0}, extra || {});

test('mailEval: тізбекті воронка, қазіргі күй топтарының қосындысы = жіберілді, өтпегендер дұрыс', () => {
  const {A, S, OK, LOST, OPEN} = setup();
  const rows = [{f: 1, s: S.mWait, o: OPEN, n: 60, b: 0}, {f: 1, s: S.mRep, o: OPEN, n: 5, b: 0, rep: 1}, {f: 0, s: S.opWait, o: OPEN, n: 10, b: 0}, {f: 0, s: S.opCons, o: OPEN, n: 7, b: 0},
    {f: 0, s: S.opNd, o: OPEN, n: 3, b: 0}, {f: 0, s: S.opWon, o: OK, n: 4, b: 60000}, {f: 0, s: S.opLost, o: LOST, n: 11, b: 0}];
  const m = mk(rows), E = A.mailEval(m, [m], 50 * 24 + 18 + 100);
  const [sent, rep, proc, won] = E.steps.map(x => x.v);
  assert.deepEqual([sent, rep, proc, won], [100, 40, 25, 4]);   // жауап: ОП 35 + рассылка жауап кезеңі 5; өңделмеген: ОП «рассылка» 10 + «Жауап берді» 5
  assert.ok(sent >= rep && rep >= proc && proc >= won, 'monotonic');
  const B = E.B, sum = B.noRep + B.wait + B.nd + B.cons + B.inv + B.other + B.won + B.lost; assert.equal(sum, sent, 'every deal in exactly one bucket');
  assert.equal(E.steps[2].lost, 15); assert.equal(B.wonS, 60000); assert.equal(E.stage.k, 'd7');
  assert.equal(E.cmp.ok, false, 'no peers → no comparison'); assert.equal(E.dev, null);
});

test('mailCompare: ұқсастары 3-тен аз — салыстыру жоқ; дәуір мен түр бөлек; сенімсіз айырма ауытқу деп аталмайды', {skip: 'asserts values/format of the original private snapshot; demo data is random'}, () => {
  const {A, S, OK, OPEN} = setup();
  const r = (won, n) => [{f: 1, s: S.mWait, o: OPEN, n: n - 20 - won, b: 0}, {f: 0, s: S.opCons, o: OPEN, n: 20, b: 0}].concat(won ? [{f: 0, s: S.opWon, o: OK, n: won, b: 1}] : []);
  const m = mk(r(1, 100), {imp: 9, a: 70}), peers = [1, 2, 3].map(i => mk(r(1, 100), {imp: i, a: 66})), old = mk(r(30, 100), {imp: 5, a: 40});
  const now = 70 * 24 + 18 + 30;
  const C = A.mailCompare(m, [m].concat(peers, old), now);
  assert.equal(C.n, 3, 'pre-13.10 mailing is not a peer'); assert.equal(C.ok, true);
  assert.ok(C.rows.every(x => x.k === 'won' ? x.fair : !x.fair), 'younger than 72h: only the windowed sales row is fair');
  const E = A.mailEval(m, [m].concat(peers, old), now); assert.equal(E.dev, null, 'no dhi times → sales in window are 0 vs 0, nothing to report');
  const w = A.wilson(3, 125); assert.ok(w[0] > .002 && w[1] < .08);
});
