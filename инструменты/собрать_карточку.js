// Собирает card.html — отдельную страницу-карточку (без сайта и без Apps Script) из того же кода, что в EduApp_OP_API.gs.
// Карточка берёт свежие агрегаты у веб-приложения Apps Script (только суммы и количества) и рисует отчёт за выбранный день.
// Запуск:  node инструменты/собрать_карточку.js   → рядом с Атлас_отчёты.html появляется card.html
// Адреса после публикации: /leads?d=2026-10-08 и /sales?d=2026-10-08 (см. vercel.json), или card.html?card=sales&d=…
const fs = require('fs'), vm = require('vm'), path = require('path');
const dir = path.join(__dirname, '..');
const gs = fs.readFileSync(path.join(__dirname, 'EduApp_OP_API.gs'), 'utf8');
const ctx = {module: {exports: {}}, Date}; vm.createContext(ctx); vm.runInContext(gs, ctx);
const css = vm.runInContext('ZOP_CARD_CSS', ctx), client = ctx.module.exports.zopCardClient_.toString();
const html = fs.readFileSync(path.join(dir, 'Атлас_отчёты.html'), 'utf8');
const m = html.match(/const LIVE_URL = '([^']*)'/);
if (!m) throw new Error('LIVE_URL не найден в Атлас_отчёты.html');
const api = m[1];   // пусто → демо: карточка читает demo/crm.json и demo/sales.json (случайные данные)

const page = `<!doctype html>
<html lang="kk">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>EduApp · күнделікті есеп</title>
<meta property="og:title" content="EduApp · күнделікті есеп">
<meta property="og:description" content="Лидтер мен сатылым бойынша күнделікті карточка — графиктермен">
<meta name="theme-color" content="#16224f">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='9' fill='%233b5bdb'/%3E%3Cpath d='M9 22l5-6 4 3 5-8' stroke='white' stroke-width='3' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E">
<style>${css}
.load{max-width:760px;margin:0 auto;padding:14px}.load .hero{min-height:190px}.sk{height:150px;border-radius:20px;margin-top:12px;background:linear-gradient(90deg,#e6ebf4 0%,#f3f6fb 50%,#e6ebf4 100%);background-size:200% 100%;animation:sk 1.2s linear infinite}
@keyframes sk{to{background-position:-200% 0}}.err{background:#fff;border-radius:20px;padding:18px;margin-top:12px;color:#475569}.err b{color:#0f172a}
</style>
</head>
<body>
<div id="app"><div class="load"><header class="hero"><div class="eyebrow">EduApp</div><h1>Есеп жүктелуде…</h1></header><div class="sk"></div><div class="sk"></div></div></div>
<script>
const API = ${JSON.stringify(api)};
const RENDER = ${client};
(function () {
  const q = new URLSearchParams(location.search), kind = q.get('card') === 'sales' || /sales/i.test(location.pathname) ? 'sales' : 'leads';
  const date = q.get('d') || '';
  const fail = msg => { document.getElementById('app').innerHTML = '<div class="load"><header class="hero"><div class="eyebrow">EduApp</div><h1>Есеп ашылмады</h1></header><div class="err"><b>Деректер жүктелмеді.</b><br>' + msg + '<br><br><button class="btn pri" onclick="location.reload()">Қайта көру</button></div></div>'; };
  const DEMO = !API, src = s => DEMO ? '/demo/' + (s ? 'sales' : 'crm') + '.json' : API + (s ? '?sales=1' : '');   // демо: без Apps Script
  const get = (n) => fetch(src(kind === 'sales'), {redirect: 'follow'}).then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
    .catch(e => n > 0 ? new Promise(res => setTimeout(res, 1500)).then(() => get(n - 1)) : Promise.reject(e));
  const sheet = kind === 'sales' ? fetch(src(false), {redirect: 'follow'}).then(r => r.ok ? r.json() : null).catch(() => null) : Promise.resolve(null);   // 15.10: сатылым карточкасына CRM керек (ОП кестесі алынды)
  Promise.all([get(1), sheet]).then(([data, sd]) => {
    if (data.error) throw new Error(data.error);
    if (kind === 'sales' ? !data.cube : !data.main) throw new Error('Google Apps Script ескі нұсқада — жаңа нұсқаны жариялаңыз.');
    if (sd && sd.main) data.crm = sd;
    RENDER({kind, date, url: '', api: API, data});
  }).catch(e => fail(String(e && e.message || e)));
})();
</script>
</body>
</html>
`;
fs.writeFileSync(path.join(dir, 'card.html'), page);
console.log('card.html', page.length, 'символов · данные из', api ? api.replace(/\/s\/(.{8}).*\/exec/, '/s/$1…/exec') : 'demo/*.json (демо)');
