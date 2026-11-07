// Все источники в таблице AMO: формула зеркала, проверка готовности, без персональных столбцов.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const gs = loadGs(), X = gs.exp, M = gs.value('ZOP_MIRROR');

test('Буквы столбцов', () => {
  assert.equal(X.zopColLetter_(1), 'A'); assert.equal(X.zopColLetter_(26), 'Z'); assert.equal(X.zopColLetter_(27), 'AA'); assert.equal(X.zopColLetter_(53), 'BA');
});
test('Формула: IMPORTRANGE по ID и листу, выбор столбцов по заголовку', () => {
  const f = X.zopMirrorFormula_('ID1', 'Оқушы тізімі', 30, ['Дата', 'Төлем сомасы']);
  assert.match(f, /^=LET\(s,IMPORTRANGE\("ID1","'Оқушы тізімі'!A1:AD"\)/);
  assert.match(f, /ARRAYFORMULA\(XMATCH\(SPLIT\("Дата\|Төлем сомасы","\|"\),h\)\)/);   // без ARRAYFORMULA — только первый столбец
  assert.match(f, /CHOOSECOLS\(s,FILTER\(i,ISNUMBER\(i\)\)\)\)$/);
});
test('Зеркало готово только с данными и нужными заголовками', () => {
  assert.equal(X.zopMirrorOk_([['#REF!'], ['']], ['Дата']), false);
  assert.equal(X.zopMirrorOk_([['Loading...']], ['Дата']), false);
  assert.equal(X.zopMirrorOk_([['Дата', 'Төлем сомасы']], ['Дата']), false);   // только заголовок
  assert.equal(X.zopMirrorOk_([['Қайдан келді'], ['WhatsApp']], ['Дата']), false);
  assert.equal(X.zopMirrorOk_([['Дата', 'Төлем сомасы'], ['10.10.2026', 5000]], ['Дата']), true);
});
test('Зеркала — полная копия листа (решение владельца): последняя строка ищется сама, части по 2000, без молчаливых IFERROR', () => {
  assert.equal(M.sales.cols, null); assert.equal(M.plat.cols, null);
  const f = X.zopMirrorFormula_('ID1', 'все тренинги и челленджи платформа', 43, null, 25000);
  assert.ok(f.startsWith(`=LET(u,"ID1",s,"'все тренинги и челленджи платформа'!",k,2000,b,IMPORTRANGE(u,s&"B1:B"),n,MAX(FILTER(SEQUENCE(ROWS(b)),b<>"")),`), f);
  assert.ok(f.includes(`f,IMPORTRANGE(u,s&"A1:AQ"&MIN(n,k))`) && f.includes(`IMPORTRANGE(u,s&"A"&(i*k+1)&":AQ"&MIN(n,(i+1)*k))`), f);
  assert.ok(!/IFERROR/.test(f) && !/\d{5}/.test(f), 'нет жёсткого предела строк и скрытых ошибок');
  assert.equal((f.match(/\(/g) || []).length, (f.match(/\)/g) || []).length, 'скобки');
});
test('Платформа из исходного листа (без копии ОП): дневная сверка не поднимает ложную тревогу', () => {
  const {loadAtlas} = require('./lib/load.cjs'), A = loadAtlas();
  const dict = {src: [''], name: [''], ptype: [''], pstatus: ['2'], pub: ['']};
  A.initSales({cube: [['s', '2026-10-10', 0, 0, 100, 0, 0, 0, 0, 0, 0, 1, 100], ['p', '2026-10-10', 0, 0, 50, 0, 0, 1, 0, 0, 0, 1, 50]], dict, snapshot: '2026-10-12T09:00'});
  assert.equal(A.payReconcile(), null);
  A.initSales({cube: [['s', '2026-10-10', 0, 0, 100, 0, 0, 0, 0, 0, 0, 1, 100], ['p', '2026-10-10', 0, 0, 100, 0, 0, 0, 0, 0, 0, 1, 100]], dict, snapshot: '2026-10-12T09:00'});
  assert.equal(A.payReconcile().matched, 1);   // копия есть — сверка работает как раньше
});
