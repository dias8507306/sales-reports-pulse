// Чектен шығын: Wazzup шотының мәтінінен күн, нөмір, сома, валюта, санат; төлеушінің аты алынбайды. Үлгі — жеке дерексіз.
'use strict';
const test = require('node:test'), assert = require('node:assert');
const {fixture, loadAtlas} = require('./lib/load.cjs');
const A = loadAtlas(fixture('atlas_demo.json'));

test('parseReceipt: Wazzup шоты', () => {
  const t = `Wazzup INVOICE № 111111/2026 Invoice Date: 13.10.2026 Date of payment: 14.10.2026 SELLER: SIA "WAZZUP" CUSTOMER: Test Customer
    ITEM DESCRIPTION QTY PRICE excl. VAT AMOUNT 1 Service wazzup24.com account 1234-5678 1 20000 20 000 Currency KZT VAT% 0% Total VAT 0 Total incl.VAT 20 000 Total: 20 000`;
  const o = A.parseReceipt(t);
  assert.equal(o.date, '2026-10-14'); assert.equal(o.doc, '111111/2026'); assert.equal(o.sum, 20000); assert.equal(o.cur, 'KZT');
  assert.equal(o.prov, 'Wazzup'); assert.equal(o.cat, 'Wazzup top-up'); assert.equal(o.note, 'Service wazzup24.com account 1234-5678');
  assert.ok(!JSON.stringify(o).includes('Test Customer'), 'төлеушінің аты алынбайды');
});

test('parseReceipt: мәтін жоқ — бос нәтиже', () => {
  assert.deepEqual(Object.keys(A.parseReceipt('')), []);
});
