// Комментарий к рассылке с сайта: код доступа, очистка текста, выбор строки листа «Рассылки».
'use strict';
const test = require('node:test'), assert = require('node:assert');
const {loadGs} = require('./lib/load.cjs');
const gs = loadGs(), {zopCodeOk_, zopNoteClean_, zopNoteTarget_} = gs.exp;

test('Код доступа: совпадение, пустой/короткий код в свойствах = запись выключена', () => {
  assert.equal(zopCodeOk_('eduapp7', 'eduapp7'), true);
  assert.equal(zopCodeOk_('eduapp8', 'eduapp7'), false);
  assert.equal(zopCodeOk_('eduapp7', ''), false);
  assert.equal(zopCodeOk_('', ''), false);
  assert.equal(zopCodeOk_('abc', 'abc'), false);          // короче 4 — не принимается
  assert.equal(zopCodeOk_(null, 'eduapp7'), false);
});

test('Текст: без управляющих символов, ≤ 500, формула → текст', () => {
  assert.equal(zopNoteClean_('  Жақсы шықты\r\n15% жеңілдік  '), 'Жақсы шықты\n15% жеңілдік');
  assert.equal(zopNoteClean_('=HYPERLINK("x")'), "'=HYPERLINK(\"x\")");
  assert.equal(zopNoteClean_('-5% нашар'), "'-5% нашар");
  assert.equal(zopNoteClean_('a\u0000b\u0007c'), 'abc');
  assert.equal(zopNoteClean_('x'.repeat(700)).length, 500);
  assert.equal(zopNoteClean_(null), '');
});

test('Строка: тот же день и воронка; иначе новая; нет столбца «Комментарий» — добавляется', () => {
  const ymd = v => v instanceof Date ? v.toISOString().slice(0, 10) : String(v || '');
  const head = ['Рассылка', 'Тема', 'Дата начала', 'Дата окончания', 'Воронка', 'Кому (сегмент)', 'Доп. расход', 'Отправлено, шт', 'Комментарий'];
  const v = [head, ['Р1', '', '2026-10-08', '', 'Рассылка', '', '', '', ''], ['Р2', '', '2026-10-08', '', 'Рассылка 5-6 класс Алтын сақа', '', '', '', 'бар']];
  assert.deepEqual(zopNoteTarget_(v, ymd, '2026-10-08', 'Рассылка'), {row: 1, cA: 2, cF: 4, cN: 8, cName: 0, addNoteCol: false});
  assert.equal(zopNoteTarget_(v, ymd, '2026-10-08', 'Рассылка 5-6 класс Алтын сақа').row, 2);
  assert.equal(zopNoteTarget_(v, ymd, '2026-10-09', 'Рассылка').row, -1);
  const noNote = [head.slice(0, 8), ['Р1', '', '2026-10-08', '', 'Рассылка', '', '', '']];
  const t = zopNoteTarget_(noNote, ymd, '2026-10-08', 'Рассылка');
  assert.equal(t.addNoteCol, true); assert.equal(t.cN, 8); assert.equal(t.row, 1);
});

test('16.10: бір күнде бірнеше рассылка — жол уақыт бойынша («Уақыты»); уақытсыз жол бірінші рассылкаға беріледі', () => {
  const {zopHm_} = gs.exp, ymd = v => String(v || '');
  assert.equal(zopHm_('9:05'), '09:05'); assert.equal(zopHm_('18:34'), '18:34'); assert.equal(zopHm_('25:00'), ''); assert.equal(zopHm_(''), '');
  const head = ['Рассылка', 'Тема', 'Дата начала', 'Воронка', 'Комментарий', 'Уақыты'];
  const v = [head, ['Р', '', '2026-10-15', 'Рассылка 5-6', '', '18:29'], ['Р', '', '2026-10-15', 'Рассылка 5-6', '', '']];
  const F56 = 'Рассылка 5-6 класс Алтын сақа';
  assert.equal(zopNoteTarget_(v, ymd, '2026-10-15', F56, '18:29').row, 1);
  const t = zopNoteTarget_(v, ymd, '2026-10-15', F56, '18:34'); assert.equal(t.row, 2); assert.equal(t.setTime, true); assert.equal(t.cTm, 5);
  const v2 = [head, ['Р', '', '2026-10-15', 'Рассылка 5-6', '', '18:29']];
  assert.equal(zopNoteTarget_(v2, ymd, '2026-10-15', F56, '18:34').row, -1);   // жаңа жол
  const noCol = [head.slice(0, 5), ['Р', '', '2026-10-15', 'Рассылка 5-6', '']];
  const t3 = zopNoteTarget_(noCol, ymd, '2026-10-15', F56, '18:34'); assert.equal(t3.row, 1); assert.equal(t3.cTm, -1); assert.equal(t3.setTime, true);
});
