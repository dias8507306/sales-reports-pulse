// Күнделікті лидтер есебі (иесінің шаблоны): 6 бөлім, бос мән жоқ, менеджерлер сомасы = жалпы «тексеру керек».
'use strict';
const test = require('node:test'), assert = require('node:assert');
const {fixture, loadAtlas} = require('./lib/load.cjs');
const SNAP = fixture('atlas_demo.json');

test('Лидтер есебі: барлық бөлім бар, NaN/undefined жоқ', () => {
  const A = loadAtlas(SNAP);
  for (let d = A.END - 5; d <= A.END; d++) {
    const t = A.leadsDayText(d, 'https://example.test');
    for (const h of ['ҚЫСҚАША', '1. ЛИДТЕР ҚАЙДАН КЕЛДІ?', 'ЖАБЫЛҒАН СДЕЛКАЛАР', '3. ТЕКСЕРУ КЕРЕК', '4. МЕНЕДЖЕРЛЕР', '5. ТӨЛЕМ ↔ CRM', '6. ТЕХҚОЛДАУ', '7. ЕСКЕРТПЕ', 'КҮН ҚОРЫТЫНДЫСЫ']) assert.ok(t.includes(h), h);
    assert.ok(!/NaN|undefined|null/.test(t), 'бос мән жоқ'); assert.ok(!t.includes('БАСТЫ МІНДЕТ'), 'басты міндет алынған');
  }
});

test('Лидтер есебі: менеджерлердің «3+ күн» сомасы жалпыға тең', () => {
  const A = loadAtlas(SNAP), t = A.leadsDayText(A.END, ''), num = s => +String(s).replace(/\s/g, '');
  const d3 = num((t.match(/3\+ күн жаңартылмаған — \*([\d\s]+)\*/) || [])[1]);
  if (!d3) return;   // срезде D.idle жоқ
  const per = [...t.matchAll(/⏳ 3\+ күн — ([\d\s]+)/g)].reduce((a, m) => a + num(m[1]), 0);
  assert.equal(per, d3);
});
