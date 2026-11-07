// Паритет: «Расходы» v2 в Apps Script (zopCosts_) и в Python (обновить_данные.py → costs_from) дают одно и то же.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('fs'), path = require('path'), {spawnSync} = require('child_process');
const {loadGs, ROOT} = require('./lib/load.cjs');

const HEAD = ['Дата', 'Подписка', 'Детали', 'Сумма, ₸', 'Оплата', 'Рассылка (если знаете)', 'Категория', 'Поставщик', 'Валюта', 'Курс', 'Кезең басы', 'Кезең соңы', 'Кампания', 'Хабарлама саны', 'Тариф', 'Құжат', 'Тексерілді', 'ID'];
const R = o => HEAD.map(h => o[h] === undefined ? '' : o[h]);
const ROWS = [HEAD,
  R({'Дата': '07.10.2026', 'Подписка': 'Рассылка', 'Детали': 'Subscription balance', 'Сумма, ₸': 50000, 'Оплата': 'партнёр'}),
  R({'Дата': '07.10.2026', 'Подписка': 'Рассылка', 'Детали': 'Subscription balance', 'Сумма, ₸': 50000, 'Оплата': 'партнёр'}),
  R({'Дата': '05.10.2026', 'Подписка': 7000, 'Детали': 'Subscription balance', 'Сумма, ₸': 20000, 'Оплата': 'партнёр'}),
  R({'Дата': '04.10.2026', 'Подписка': 7000, 'Детали': '1 channel on the "Waba" plan for 1 year', 'Сумма, ₸': 346600, 'Оплата': 'партнёр'}),
  R({'Дата': '09.10.2026', 'Категория': 'WhatsApp template charges', 'Поставщик': 'Wazzup', 'Хабарлама саны': 1000, 'Тариф': 'marketing', 'Кампания': 'А=30000; Б=24000'}),
  R({'Дата': '09.10.2026', 'Категория': 'Other sales expenses', 'Сумма, ₸': 20, 'Валюта': 'USD'}),
  R({'Дата': '09.10.2026', 'Категория': 'Other sales expenses', 'Сумма, ₸': 20, 'Валюта': 'USD', 'Курс': 500, 'Тексерілді': 'Иә', 'ID': 'x1'}),
  R({'Дата': '10.10.2026', 'Категория': 'Marketing campaign costs', 'Сумма, ₸': 50000, 'Кампания': 'А:60%; Б:40%'}),
  R({'Дата': '10.10.2026', 'Категория': 'Marketing campaign costs', 'Сумма, ₸': 50000, 'Кампания': 'А=1; Б=2'})];
const TARIFFS = [{prov: 'Wazzup', type: 'marketing', from: 52, to: 82, price: 54, cur: 'KZT'}];
const pick = c => ({d: c.d, cat: c.cat, kind: c.kind, sum: c.sum == null ? null : Math.round(c.sum * 100) / 100, cur: c.cur, prov: c.prov, months: c.months, dup: c.dup, src: c.src, ver: c.ver, sub: c.sub,
  alloc: (c.alloc || []).map(a => [a.name, a.sum == null ? null : Math.round(a.sum)]), warn: c.warn.length});

test('Расходы v2: Apps Script и Python совпадают', t => {
  const py = spawnSync('python', ['--version']); if (py.status !== 0) { t.skip('python недоступен'); return; }
  const g = loadGs().exp, B = Date.UTC(2026, 7, 10), dayIdx = p => p ? Math.round((Date.UTC(p.y, p.m - 1, p.d) - B) / 864e5) : -1;
  const gsOut = g.zopCosts_(ROWS, d => ({y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate()}), dayIdx, TARIFFS).map(pick);
  const src = fs.readFileSync(path.join(ROOT, 'инструменты', 'обновить_данные.py'), 'utf8'), i = src.indexOf('# ---------- «Расходы» v2'), j = src.indexOf('\n    return out\n', src.indexOf('def costs_from', i)) + 16;
  const prog = 'import re, json, sys, io, datetime as dt\nsys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding="utf-8")\nsys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")\nBASE = dt.date(2026, 8, 10)\n' + src.slice(i, j) + '\nrows = json.loads(sys.stdin.read())\nprint(json.dumps(costs_from(rows["rows"], rows["tariffs"]), ensure_ascii=False))\n';
  const r = spawnSync('python', ['-I', '-c', prog], {input: JSON.stringify({rows: ROWS, tariffs: TARIFFS}), env: Object.assign({}, process.env, {PYTHONIOENCODING: 'utf-8'}), encoding: 'utf8'});
  assert.equal(r.status, 0, r.stderr);
  const pyOut = JSON.parse(r.stdout).map(pick);
  assert.equal(pyOut.length, gsOut.length);
  gsOut.forEach((x, k) => assert.deepEqual(pyOut[k], JSON.parse(JSON.stringify(x)), 'row ' + k));
  assert.equal(gsOut[1].dup, true); assert.equal(gsOut[4].sum, 54000); assert.equal(gsOut[5].sum, null); assert.equal(gsOut[6].sum, 10000);
});
