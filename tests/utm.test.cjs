// Креативтер amoCRM бойынша (иесі 15.10: «7000 — серверсіз»): utm бағандары аты бойынша табылады, ұзын сандар шықпайды.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const J = x => JSON.parse(JSON.stringify(x));
test('utm кубы: креатив × нәтиже; телефон тәрізді мән жасырылады; utm бағандары болмаса — бос', () => {
  const g = loadGs().exp, parts = d => { if (!(d instanceof Date)) return null; const x = new Date(d.getTime() + 5 * 3600e3); return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours(), mi: x.getUTCMinutes()}; };
  const H = ['Воронка', 'Этап', 'Итог', 'Ответственный', 'Создана', 'Изменена', 'Закрыта', 'Қайдан келді', 'UTM_CONTENT', 'utm campaign', 'utm_source'];
  const row = (cr, oc) => ['ОП', 'НОВАЯ ЗАЯВКА', oc, 'М', new Date('2026-10-01T05:00:00Z'), new Date('2026-10-01T05:00:00Z'), '', 'Таргет', cr, 'Алтын сақа күз', 'instagram'];
  const pipes = [['', 'Воронка', '', '', 'Этап', 'Порядок'], ['', 'ОП', '', '', 'НОВАЯ ЗАЯВКА', 1]];
  const out = g.zopAggregate_([H, row('Видео 1', ''), row('Видео 1', 'Успешно'), row('87071234567', ''), ['ОП', 'НОВАЯ ЗАЯВКА', '', 'М', new Date('2026-10-01T05:00:00Z'), new Date('2026-10-01T05:00:00Z'), '', 'WhatsApp', '', '', '']], pipes, [['Время'], [new Date('2026-10-02T05:00:00Z')]], parts, [], [], []);
  const X = out.dict.utmx, O = out.dict.outcome;
  assert.deepEqual(J(out.utmCols), {content: 'UTM_CONTENT', campaign: 'utm campaign', source: 'utm_source'});
  const by = {}; out.utm.forEach(r => { const k = X[r[2]] + '|' + O[r[4]]; by[k] = (by[k] || 0) + r[r.length - 1]; });
  assert.deepEqual(by, {'Видео 1|В работе': 1, 'Видео 1|Успешно': 1, '—|В работе': 1}, 'phone-like creative value is dropped, deal without utm is not counted');
  assert.ok(!JSON.stringify(out).includes('87071234567'));
});
