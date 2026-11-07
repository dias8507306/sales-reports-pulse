// Өңірлер (иесі 15.10: «где новые регионы»): «Облыс» жазылуы әр түрлі болса да — ресми 17 облыс + 3 қала; Apps Script пен Python бірдей.
// Өңір графигі — жеке куб reg [воронка, нәтиже, өңір]: тек 3-тен аз өңір жасырылады (профильдің 9 өлшемді жасыруы өңірлерді «Шағын өңірлерге» жұтып қойған).
'use strict';
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('fs'), path = require('path'), {spawnSync} = require('child_process');
const {loadGs, ROOT} = require('./lib/load.cjs');

const CASES = [['Қызылорда облысы', 'Қызылорда'], ['Павлодар облысы', 'Павлодар'], ['Кызылординская область', 'Қызылорда'], ['ПАВЛОДАР', 'Павлодар'],
  ['Алматы қ.', 'Алматы'], ['Алматы қаласы', 'Алматы'], ['г. Алматы', 'Алматы'], ['Алматы облысы', 'Алматы обл.'], ['Алматы обл', 'Алматы обл.'], ['Алматинская область', 'Алматы обл.'],
  ['Астана қаласы', 'Астана'], ['АСТАНА', 'Астана'], ['Нұр-Сұлтан', 'Астана'], ['Шымкент қаласы', 'Шымкент'], ['Ақмола облысы', 'Ақмола обл.'], ['Ақмола', 'Ақмола обл.'],
  ['Батыс Қазақстан облысы', 'Батыс Қазақстан'], ['ЗКО', 'Батыс Қазақстан'], ['Солтүстік Қазақстан облысы', 'Солтүстік Қазақстан'], ['Шығыс Қазақстан облысы', 'Шығыс Қазақстан'],
  ['Восточно-Казахстанская область', 'Шығыс Қазақстан'], ['Жамбыл облысы', 'Жамбыл'], ['Жетісу облысы', 'Жетісу'], ['Ұлытау облысы', 'Ұлытау'], ['Абай облысы', 'Абай'],
  ['Қостанай облысы', 'Қостанай'], ['Маңғыстау облысы', 'Маңғыстау'], ['Мангистауская область', 'Маңғыстау'], ['Ақтөбе облысы', 'Ақтөбе'], ['Атырау облысы', 'Атырау'],
  ['Қарағанды облысы', 'Қарағанды'], ['Түркістан облысы', 'Түркістан'], ['  ', 'Не указан'], ['', 'Не указан'], [null, 'Не указан'], ['Ресей', 'Не указан'], ['123', 'Не указан']];

test('zopRegion_: барлық жазылу түрі → бір ресми атау; белгісіз → «Не указан»', () => {
  const g = loadGs().exp;
  for (const [v, want] of CASES) assert.equal(g.zopRegion_(v), want, String(v));
});

test('region_of (Python) = zopRegion_ (Apps Script)', t => {
  const py = spawnSync('python', ['--version']); if (py.status !== 0) { t.skip('python недоступен'); return; }
  const src = fs.readFileSync(path.join(ROOT, 'инструменты', 'обновить_данные.py'), 'utf8'), i = src.indexOf('\nimport re as _re'), j = src.indexOf('\nSOURCES = {', i);
  assert.ok(i > 0 && j > i, 'region_of block');
  const prog = 'import json, sys, io\nsys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding="utf-8")\nsys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")\n' + src.slice(i, j) + '\nprint(json.dumps([region_of(v) for v in json.load(sys.stdin)], ensure_ascii=False))';
  const r = spawnSync('python', ['-I', '-c', prog], {input: JSON.stringify(CASES.map(c => c[0])), env: Object.assign({}, process.env, {PYTHONIOENCODING: 'utf-8'}), encoding: 'utf8'});
  assert.equal(r.status, 0, r.stderr);
  assert.deepEqual(JSON.parse(r.stdout), CASES.map(c => c[1]));
});

test('reg кубы: өңір 3+ болса көрінеді (класс/тіл әр түрлі болса да), 1–2 болса жасырылады', () => {
  const g = loadGs().exp, parts = d => { if (!(d instanceof Date)) return null; const x = new Date(d.getTime() + 5 * 3600e3); return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours(), mi: x.getUTCMinutes()}; };
  const H = ['Воронка', 'Этап', 'Итог', 'Ответственный', 'Создана', 'Изменена', 'Закрыта', 'Облыс', 'Класс', 'Язык обучения', 'Қайдан келді'];
  const row = (obl, cls) => ['ОП', 'НОВАЯ ЗАЯВКА', '', 'М', new Date('2026-10-01T05:00:00Z'), new Date('2026-10-01T05:00:00Z'), '', obl, cls, 'Қазақша', 'Рассылка'];
  const leads = [H, row('Павлодар облысы', 3), row('Павлодар', 5), row('ПАВЛОДАР облысы', 7), row('Қызылорда облысы', 6)];
  const pipes = [['', 'Воронка', '', '', 'Этап', 'Порядок'], ['', 'ОП', '', '', 'НОВАЯ ЗАЯВКА', 1]];
  const out = g.zopAggregate_(leads, pipes, [['Время'], [new Date('2026-10-02T05:00:00Z')]], parts, [], [], []);
  const R = out.dict.region, n = name => out.reg.filter(r => R[r[2]] === name).reduce((a, r) => a + r[r.length - 1], 0);
  assert.equal(n('Павлодар'), 3, 'three children with different classes — region visible in reg');
  assert.equal(n('Қызылорда'), 0, 'a single child from Қызылорда is hidden in reg too');
  assert.equal(n('Жасырылған (< 3)'), 1);
  assert.ok(!out.prof.some(r => R[r[2]] === 'Павлодар'), 'prof still masks 1-child cells');
});

test('Атлас: өңір графигі reg кубынан — аз өңірлер «Басқалары»-ға жұтылмайды, үлес сол кубтан', () => {
  const {loadAtlas, fixture} = require('./lib/load.cjs');
  const snap = fixture('atlas_demo.json'), R = snap.dict.region, fu = snap.dict.funnel.indexOf('ОП'), O = snap.dict.outcome, add = name => { let i = R.indexOf(name); if (i < 0) { R.push(name); i = R.length - 1; } return i; };
  const ok = O.indexOf('Успешно'), lost = O.indexOf('Не реализовано');
  snap.reg = [[fu, ok, add('Павлодар'), 4], [fu, lost, add('Павлодар'), 2], [fu, ok, add('Қызылорда'), 3], [fu, ok, add('Жасырылған (< 3)'), 2], [fu, lost, add('Не указан'), 9]];
  const A = loadAtlas(snap), r = A.regionRows({f: -1, of: null, src: null});
  assert.equal(r.viaReg, true);
  assert.deepEqual(JSON.parse(JSON.stringify(r.rows.map(e => [e.name, e.won, e.n, e.few, e.hid]))), [['Павлодар', 4, 6, true, false], ['Қызылорда', 3, 3, true, false], ['Шағын өңірлер', 2, 2, true, true]]);
  assert.equal(r.all, 20); assert.equal(r.filled, 11);
  assert.equal(A.regionRows({f: -1, of: [0], src: null}).viaReg, false, 'offer filter → profile cube');
});
