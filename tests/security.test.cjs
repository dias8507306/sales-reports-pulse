// Защита Apps Script: функции, меняющие таблицу, не вызываются анонимно со страницы веб-приложения; форма — только с токеном.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs, fixture} = require('./lib/load.cjs');
const S = (active, eff) => ({getActiveUser: () => ({getEmail: () => active}), getEffectiveUser: () => ({getEmail: () => eff})});

test('Только владелец: анонимный посетитель веб-приложения (active = "") и чужой аккаунт отклоняются', () => {
  const g = loadGs().exp;
  assert.throws(() => g.zopRequireOwner_(S('', 'owner@x')), /иесі/);
  assert.throws(() => g.zopRequireOwner_(S('other@x', 'owner@x')), /иесі/);
  assert.doesNotThrow(() => g.zopRequireOwner_(S('owner@x', 'owner@x')));
});
test('Публичные административные функции без сессии владельца падают до любых изменений', () => {
  const gs = loadGs(); gs.ctx.Session = S('', 'owner@x');
  for (const f of ['zopSetup', 'zopDescribeSources', 'zopRebuildProducts']) assert.throws(() => gs.exp[f](), /иесі/, f);
});
test('Сохранение расхода: без токена формы — отказ; токен сверяется с кэшем пользователя', () => {
  const g = loadGs().exp, cache = v => ({get: () => v});
  assert.throws(() => g.zopSaveCost({}), /мәзірінен/); assert.throws(() => g.zopSaveCost(null), /мәзірінен/);
  assert.equal(g.zopCheckToken_('abc', cache('abc')), true); assert.equal(g.zopCheckToken_('abc', cache('zzz')), false); assert.equal(g.zopCheckToken_('', cache('')), false); assert.equal(g.zopCheckToken_('abc', cache(null)), false);
});
test('Данные веб-приложения: ни телефонов, ни e-mail в выдаче CRM и продаж (фикстуры)', () => {
  for (const f of ['atlas_demo.json', 'card_sales_demo.json', 'card_leads_demo.json']) {
    const t = JSON.stringify(fixture(f));
    assert.ok(!/(?:\+7|\b8)[\s(-]*7\d{2}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}\b/.test(t), f + ': phone');
    assert.ok(!/[\w.+-]+@[\w-]+\.[a-z]{2,}/i.test(t), f + ': e-mail');
  }
});

test('Профиль аудитории: ячейки из 1–2 детей обобщаются (регион/класс/язык скрыты), итоги сохраняются', () => {
  const g = loadGs().exp, parts = d => { if (!(d instanceof Date)) return null; const x = new Date(d.getTime() + 5 * 3600e3); return {y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours(), min: 0, wd: (x.getUTCDay() + 6) % 7}; };
  const H = ['Воронка', 'Этап', 'Итог', 'Ответственный', 'Создана', 'Изменена', 'Закрыта', 'Облыс', 'Класс', 'Язык обучения', 'Қайдан келді'];
  const row = (obl, cls) => ['ОП', 'НОВАЯ ЗАЯВКА', '', 'М', new Date('2026-10-01T05:00:00Z'), new Date('2026-10-01T05:00:00Z'), '', obl, cls, 'Қазақша', 'Рассылка'];
  const leads = [H, row('Атырау', 5), row('Атырау', 5), row('Атырау', 5), row('Атырау', 5), row('Қызылорда', 6)];
  const pipes = [['', 'Воронка', '', '', 'Этап', 'Порядок'], ['', 'ОП', '', '', 'НОВАЯ ЗАЯВКА', 1]];
  const out = g.zopAggregate_(leads, pipes, [['Время'], [new Date('2026-10-02T05:00:00Z')]], parts, [], [], []);
  const R = out.dict.region, total = out.prof.reduce((a, r) => a + r[r.length - 1], 0);
  assert.equal(total, 5);
  assert.ok(!out.prof.some(r => R[r[2]] === 'Қызылорда'), 'a single child from Қызылорда is not visible as such');
  assert.ok(out.prof.some(r => R[r[2]] === 'Атырау' && r[3] === 5 && r[r.length - 1] === 4));
});
