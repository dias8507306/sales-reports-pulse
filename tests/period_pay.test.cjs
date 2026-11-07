// Шолу басы (periodPay): күндер сомасы = жалпы; ай режимінде — айдың барлық күні, срезден кейінгілері бос; ОП кестесі + платформа = жалпы.
'use strict';
const test = require('node:test'), assert = require('node:assert');
const {fixture, loadAtlas} = require('./lib/load.cjs');

test('periodPay: кезең және ай', () => {
  const A = loadAtlas(fixture('atlas_demo.json'));
  if (!A.SALES) return;   // срезде төлем жоқ
  const P = A.periodPay(A.END - 29, A.END, null, null);
  assert.equal(P.days.length, 30);
  assert.equal(Math.round(P.days.reduce((t, x) => t + x.s, 0)), Math.round(P.sum));
  assert.equal(Math.round(P.sheet.s + P.plat.s), Math.round(P.sum));
  const snap = String(A.SALES.snapshot).slice(0, 10), M = A.periodPay(0, 0, snap.slice(0, 7), null);
  assert.equal(M.days.length, new Date(Date.UTC(+snap.slice(0, 4), +snap.slice(5, 7), 0)).getUTCDate());
  assert.ok(M.days.every(x => x.future === (x.iso > snap)));
  assert.ok(M.days.filter(x => x.future).every(x => x.s === 0));
});
