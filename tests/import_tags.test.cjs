// Рассылки по тегам импорта: разбор «Теги» одинаков в Apps Script и Python; номера/телефоны не выходят наружу.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('fs'), path = require('path'), {spawnSync} = require('child_process');
const {loadGs, ROOT} = require('./lib/load.cjs');
const SAMPLES = [
  'Алтын сақа, Астана қ., 3 сынып, импорт_06102026_1617',
  'WZ (77000001234), импорт_07102026_1949',
  'импорт_28092026_1159, импорт_28092026_1640, Алтын сақа',
  '  , рассылка, импорт_123_1, тег 1234567, ' + 'х'.repeat(45),
  '', null, 'Атырау қ., 5 сынып, Жылыой ауданы, импорт_04102026_1304, Алтын сақа, Алтын сақа'];

test('Теги: импорт → дата и время, сегменты без номеров', () => {
  const f = loadGs().exp.zopImportTags_;
  assert.deepEqual(JSON.parse(JSON.stringify(f(SAMPLES[0]))), {imps: ['2026-10-06 16:17'], segs: ['Алтын сақа', 'Астана қ.', '3 сынып']});
  assert.deepEqual(JSON.parse(JSON.stringify(f(SAMPLES[1]))), {imps: ['2026-10-07 19:49'], segs: []});            // WZ (телефон) не выходит
  assert.deepEqual(JSON.parse(JSON.stringify(f(SAMPLES[2]))).imps, ['2026-09-28 11:59', '2026-09-28 16:40']);
  assert.deepEqual(JSON.parse(JSON.stringify(f(SAMPLES[3]))), {imps: [], segs: ['рассылка']});                      // кривой импорт, длинные числа, > 40 символов — нет
});

test('Теги: Apps Script и Python совпадают', t => {
  const py = spawnSync('python', ['--version']); if (py.status !== 0) { t.skip('python недоступен'); return; }
  const f = loadGs().exp.zopImportTags_, gsOut = SAMPLES.map(s => { const r = f(s); return [Array.from(r.imps), Array.from(r.segs)]; });
  const src = fs.readFileSync(path.join(ROOT, 'инструменты', 'обновить_данные.py'), 'utf8'), i = src.indexOf('def import_tags(v):'), j = src.indexOf('\n    return imps, segs\n', i) + 23;
  const prog = 'import re, json, sys, io\nsys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding="utf-8")\nsys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")\n' + src.slice(i, j) + '\nprint(json.dumps([list(import_tags(s)) for s in json.loads(sys.stdin.read())], ensure_ascii=False))\n';
  const r = spawnSync('python', ['-I', '-c', prog], {input: JSON.stringify(SAMPLES), encoding: 'utf8', env: Object.assign({}, process.env, {PYTHONIOENCODING: 'utf-8'})});
  assert.equal(r.status, 0, r.stderr);
  assert.deepEqual(JSON.parse(r.stdout), gsOut);
});
