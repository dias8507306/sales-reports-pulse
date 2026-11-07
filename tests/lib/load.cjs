// Загрузчики для тестов: расчётный слой Atlas (блок @domain) и Apps Script (сервер + клиент карточки).
// Ничего не меняют в исходниках; работают на демо-срезах из demo/ (случайные данные).
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const ROOT = path.join(__dirname, '..', '..');
const read = f => fs.readFileSync(path.join(ROOT, f), 'utf8');
// Fixtures = the randomly generated demo snapshots (scripts/generate-demo-data.cjs → demo/crm.json, demo/sales.json).
// The card fixtures wrap them the way the card page receives them: {kind, date, url, data}; date = last full day before the snapshot.
const demo = f => JSON.parse(fs.readFileSync(path.join(ROOT, 'demo', f), 'utf8'));
const dayBefore = (snap, n) => new Date(Date.parse(String(snap).slice(0, 10) + 'T00:00:00Z') - n * 864e5).toISOString().slice(0, 10);
const FIXTURES = {
  'atlas_demo.json': () => demo('crm.json'),
  'card_leads_demo.json': () => { const d = demo('crm.json'); return {kind: 'leads', date: dayBefore(d.snapshot, 1), url: '', data: d}; },
  'card_sales_demo.json': () => { const d = demo('sales.json'); return {kind: 'sales', date: dayBefore(d.snapshot, 1), url: '', data: d}; },
  'card_sales_demo_prev.json': () => { const d = demo('sales.json'); return {kind: 'sales', date: dayBefore(d.snapshot, 2), url: '', data: d}; },
};
const fixture = f => FIXTURES[f] ? FIXTURES[f]() : JSON.parse(fs.readFileSync(path.join(ROOT, 'tests', 'fixtures', f), 'utf8'));

// Atlas: only the pure layer between @domain markers + the AREAS config. Returns a vm context after initData(snapshot).
function loadAtlas(snapshot, opts = {}) {
  const html = read(opts.file || 'Атлас_отчёты.html');
  const part = (a, b) => { const i = html.indexOf(a), j = html.indexOf(b, i); if (i < 0 || j < 0) throw new Error('marker not found: ' + a); return html.slice(i + a.length, j); };
  const code = part('/* @domain:start', '/* @domain:pause */') .replace(/^[^\n]*\n/, '')
    + part('/* @domain:resume */', '/* @domain:end */')
    + html.slice(html.indexOf('const AREAS = {'), html.indexOf('\nconst BLOCK_NAV'));
  const store = {};
  const ctx = {console, Date, Intl, Math, JSON, localStorage: {getItem: k => store[k] ?? null, setItem: (k, v) => { store[k] = String(v); }}};
  vm.createContext(ctx);
  vm.runInContext(code + '\n;this.__api = {initData, regionRows, mailEval, mailBuckets, mailCompare, wilson, denseStart, mergeManagers, mkChannels, channels, period, qualStats, leadsDayText, periodPay, parseReceipt, get TAGS() { return TAGS; }, get TAGSETS() { return TAGSETS; }, stats, buyGrid, profile, products, period, tileData, kpiFor, kpisFrom, notApplied, inFunnel, sourceStatus, localMs, METRIC_REG, metricInfo, payStatus, initSales, payStats, fillQuality, encodeState, decodeState, snapshotMeta, fnv, isoDay, csvFrom, waSummary, finStats, planShare, leadKinds, leadKind, mailCalc, mailCalcDefaults, payReconcile, zsiLoad, zsiEmpty, zsiApply, allCosts, costFromItem, allTariffs, tariffAt, walletCalc, costAlloc, prodFamily, famStats, famChange, planFact, planActual, allPlans, ruleInsights, buildReport, reportText, lastFullDay, waitingSummary, get MONTHS() { return MONTHS; }, get SALES() { return SALES; }, get MAN() { return MAN; }, pass, mailings, costModel, F, get KPIS() { return KPIS; }, set KPIS(v) { KPIS = v; }, get PRODUCTS() { return PRODUCTS; }, get COVER() { return COVER; }, get END() { return END; }, get SRC() { return SRC; }, get SUP() { return SUP; }, get ROWS() { return ROWS; }, get D() { return D; }, get OK() { return OK; }, METRICS, AREAS};', ctx, {filename: 'atlas-domain.js'});
  if (snapshot) ctx.__api.initData(JSON.parse(JSON.stringify(snapshot)));
  return ctx.__api;
}

// Apps Script: whole file in a sandbox; exports + the card client's calculation part (sales/leads/texts).
function loadGs() {
  const src = read('инструменты/EduApp_OP_API.gs');
  const ctx = {module: {exports: {}}, Date, console, Math, JSON};
  vm.createContext(ctx); vm.runInContext(src, ctx, {filename: 'EduApp_OP_API.gs'});
  return {ctx, exp: ctx.module.exports, value: name => vm.runInContext(name, ctx)};
}
// the card client is one function; take its body up to the drawing helpers and return the pure calculators
function cardCalc(gs, cfg) {
  const client = gs.exp.zopCardClient_.toString();
  const body = client.slice(client.indexOf('{') + 1, client.indexOf('/* ---------- drawing'));
  const doc = {getElementById: () => null, querySelector: () => null, querySelectorAll: () => []};
  const f = vm.runInContext('(function(cfg, document, location){' + body + ';return {sales, leads, waiting: typeof waiting === "function" ? waiting : null, salesText: typeof salesText === "function" ? salesText : null, leadsText: typeof leadsText === "function" ? leadsText : null, payStatus: typeof payStatus === "function" ? payStatus : null};})', gs.ctx);
  return f(cfg, doc, {pathname: '/card.html', origin: 'http://test'});
}
module.exports = {ROOT, read, fixture, loadAtlas, loadGs, cardCalc};
