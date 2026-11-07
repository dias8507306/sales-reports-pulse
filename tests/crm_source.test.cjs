// Источник CRM: сделки — только таблица amoCRM (лист «Сделки»); этапы и журнал — там же, иначе прежние листы.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const g = loadGs().exp, ZOP_CRM = loadGs().value('ZOP_CRM');

test('Берётся лист «Сделки» из новой таблицы, если он есть', () => {
  const has = (w, n) => w === 'crm' && n === 'Сделки' || w === 'own' && n === 'CRM Лиды — авто';
  assert.deepEqual({...g.zopPickSheet_(has, ZOP_CRM.LEADS, 'CRM Лиды — авто')}, {where: 'crm', name: 'Сделки'});
});
test('Нет доступа к новой таблице → прежний лист; нет ничего → null', () => {
  assert.deepEqual({...g.zopPickSheet_((w, n) => w === 'own' && n === 'CRM Лиды — авто', ZOP_CRM.LEADS, 'CRM Лиды — авто')}, {where: 'own', name: 'CRM Лиды — авто'});
  assert.equal(g.zopPickSheet_(() => false, ZOP_CRM.LEADS, 'CRM Лиды — авто'), null);
  assert.equal(ZOP_CRM.ID, 'YOUR_CRM_SHEET_ID');
});
test('Сделки: без запасного листа «CRM Лиды — авто»', () => {
  const STRICT = loadGs().value('ZOP_CRM_STRICT');
  assert.equal(STRICT.LEADS, true);
  assert.equal(g.zopPickSheet_((w, n) => w === 'own' && n === 'CRM Лиды — авто', ZOP_CRM.LEADS, null), null);
});
test('Все служебные листы — в главной таблице AMO, а не в «той, к которой привязан скрипт»', () => {
  const fs = require('fs'), path = require('path'), src = fs.readFileSync(path.join(__dirname, '..', 'инструменты', 'EduApp_OP_API.gs'), 'utf8');
  assert.equal(loadGs().value('ZOP_HOME_ID'), 'YOUR_MAIN_SHEET_ID');   // «All Info»
  assert.equal(loadGs().value('ZOP_HOME_OLD'), ZOP_CRM.ID);   // до переноса — AMO
  assert.match(src, /if \(info && info\.getSheetByName\('Расходы'\)\) return/);   // переход — по наличию «Расходы»
  assert.match(src, /'Расходы' — последним|«Расходы» — последним/);
  const body = src.slice(src.indexOf('function zopHome_()'), src.indexOf('\n}', src.indexOf('function zopHome_()')));
  assert.equal(src.split('SpreadsheetApp.getActive()').length - 1, body.split('SpreadsheetApp.getActive()').length - 1, 'getActive() только внутри zopHome_');
  const html = fs.readFileSync(path.join(__dirname, '..', 'Атлас_отчёты.html'), 'utf8');
  assert.ok(!html.includes('YOUR_OTHER_SHEET_ID'), 'Атлас не ссылается на старую «AmoCrm выгрузка»');
});
test('Основной источник — таблица «All Info»', () => {
  const I = loadGs().value('ZOP_INFO');
  assert.equal(I.ID, 'YOUR_MAIN_SHEET_ID');
  const src = require('fs').readFileSync(require('path').join(__dirname, '..', 'инструменты', 'EduApp_OP_API.gs'), 'utf8');
  assert.match(src, /if \(kind === 'LEADS'\) \{ const r = zopInfoRows_\(/);   // сделки сначала из All Info
  assert.match(src, /info = zopInfoRows_\(M\.sheet, M\.need\)/);   // продажи и платформа — тоже
});
