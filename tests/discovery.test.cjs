// Discovery: структура источников без утечки данных.
'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const {loadGs} = require('./lib/load.cjs');
const g = loadGs().exp;
test('Показываются только колонки-справочники; имена, телефоны, суммы — никогда', () => {
  const rows = [['ФИО ребенка', 'Телефон', 'Статус оплаты', 'Сумма', 'Способ оплаты', 'Комментарий'],
    ['Айгерім Б.', '+77000001234', 'Оплачено', 12000, 'Kaspi', 'перезвонить'], ['Нұрлан С.', '+77000004321', 'Ожидает', 4800, 'Kaspi', ''], ['', '', '', '', '', '']];
  const d = g.zopDescribeRows_(rows), by = Object.fromEntries(d.cols.map(c => [c.name, c]));
  assert.equal(d.rows, 2);
  assert.equal(by['ФИО ребенка'].values, null); assert.equal(by['Телефон'].values, null); assert.equal(by['Комментарий'].values, null); assert.equal(by['Сумма'].values, null);
  assert.deepEqual([...by['Статус оплаты'].values], ['Оплачено', 'Ожидает']); assert.deepEqual([...by['Способ оплаты'].values], ['Kaspi']);
  assert.equal(by['Комментарий'].fill, 50);
  assert.equal(g.zopDescribeRows_([['Оқушының аты', 'Аты-жөні'], ['x', 'y']]).cols.every(c => c.values === null), true, 'Kazakh «аты» = name');
  assert.ok(!JSON.stringify(d).includes('+7700') && !JSON.stringify(d).includes('Айгерім'));
});
