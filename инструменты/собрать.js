// Детерминированная сборка: исходники → card.html и папка vercel_deploy/. Ничего не публикует.
// Источники (правятся руками): Атлас_отчёты.html, инструменты/EduApp_OP_API.gs, atlas_v4_assets/, vercel_deploy/vercel.json.
// Генерируется (руками НЕ править): card.html, vercel_deploy/index.html, vercel_deploy/card.html, vercel_deploy/atlas_v4_assets/, vercel_deploy/BUILD.json.
//
// Запуск:   node инструменты/собрать.js            тесты → карточка → проверки → копирование → BUILD.json
//           node инструменты/собрать.js --check    только проверить, что vercel_deploy совпадает со сборкой (exit 1, если нет)
//           node инструменты/собрать.js --no-test  без node --test (например, при правке только текста)
// Один и тот же исходник даёт побайтно тот же результат: в BUILD.json нет времени сборки, только хэши.
'use strict';
const fs = require('fs'), path = require('path'), crypto = require('crypto'), {execFileSync} = require('child_process');
const ROOT = path.join(__dirname, '..'), OUT = path.join(ROOT, 'vercel_deploy');
const args = new Set(process.argv.slice(2)), CHECK = args.has('--check');
const rd = f => fs.readFileSync(path.join(ROOT, f));
const sha = b => crypto.createHash('sha256').update(b).digest('hex');
const fail = m => { console.error('✖ ' + m); process.exit(1); };
const ok = m => console.log('✔ ' + m);

// 1. tests
if (!CHECK && !args.has('--no-test')) {
  const files = fs.readdirSync(path.join(ROOT, 'tests')).filter(f => f.endsWith('.test.cjs')).map(f => path.join('tests', f));
  try { execFileSync(process.execPath, ['--test', ...files], {cwd: ROOT, stdio: 'pipe'}); ok('тесты: ' + files.length + ' файла(ов) прошли'); }
  catch (e) { process.stdout.write(String(e.stdout || '')); fail('тесты не прошли — сборка остановлена'); }
}

// 2. card.html from the Apps Script client (same code as the Apps Script card)
if (!CHECK) { execFileSync(process.execPath, [path.join(__dirname, 'собрать_карточку.js')], {cwd: ROOT, stdio: 'pipe'}); ok('card.html собран из EduApp_OP_API.gs'); }

// 3. checks on what will be published
const atlas = rd('Атлас_отчёты.html').toString('utf8'), card = rd('card.html').toString('utf8');
const scripts = h => { const re = /<script(?![^>]*application\/json)[^>]*>([\s\S]*?)<\/script>/g, out = []; let m; while ((m = re.exec(h))) if (m[1].trim()) out.push(m[1]); return out; };
for (const [n, h] of [['Атлас', atlas], ['card', card]]) { scripts(h).forEach((s, i) => { try { new Function(s); } catch (e) { fail(`${n}: скрипт #${i + 1} не разбирается: ${e.message}`); } }); }
ok('JS-синтаксис Атласа и карточки');
for (const id of ['crm', 'sales']) { const a = `<script id="${id}" type="application/json">`, i = atlas.indexOf(a); if (i < 0) fail('нет блока данных ' + id); try { JSON.parse(atlas.slice(i + a.length, atlas.indexOf('</script>', i))); } catch (e) { fail(`блок ${id}: JSON не разбирается`); } }
ok('встроенные срезы crm и sales — корректный JSON');
// personal data / secrets: aggregates only. Phones, e-mails, tokens must never reach the published files.
const PII = [[/(?:\+7|\b8)[\s(-]*7\d{2}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}\b/, 'телефон'], [/\b7\d{9}\b/, 'телефон (10 цифр)'], [/[\w.+-]+@[\w-]+\.[a-z]{2,}/i, 'e-mail'],
  [/VERCEL_OIDC|\bey[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{20,}\./, 'токен'], [/access_token|refresh_token|client_secret/i, 'секрет amoCRM']];
const scan = (n, h) => { for (const [re, what] of PII) { const m = h.match(re); if (m) fail(`${n}: найдено похожее на «${what}»: ${m[0].slice(0, 6)}…`); } };
const dataOnly = h => (h.match(/<script id="(?:crm|sales)" type="application\/json">[\s\S]*?<\/script>/g) || []).join('\n');
scan('Атлас (данные)', dataOnly(atlas)); scan('card.html', card.replace(/https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec/g, ''));
ok('в данных нет телефонов, e-mail и токенов');
const vj = JSON.parse(rd('vercel_deploy/vercel.json'));
for (const r of ['/leads', '/sales']) if (!(vj.rewrites || []).some(x => x.source === r && x.destination === '/card.html')) fail('vercel.json: нет маршрута ' + r);
ok('vercel.json: маршруты /leads и /sales на месте');

// 3b. docs/метрики.md — from the metric registry in the Atlas (one source of truth)
{
  const {loadAtlas} = require(path.join(ROOT, 'tests', 'lib', 'load.cjs'));
  const A = loadAtlas(null), esc = v => String(v).replace(/\|/g, '\\|');
  const SRC = {crm: 'amoCRM', costs: '«Расходы»', 'crm+costs': 'amoCRM + «Расходы»', sheet: '«Оқушы тізімі»', platform: 'платформа'};
  const md = ['# Справочник метрик', '', '> Генерируется `node инструменты/собрать.js` из реестра `REG` в `Атлас_отчёты.html`. Не править вручную.', '',
    `Версия метода: \`${A.METRIC_REG[0].methodVersion}\``, '',
    '| id | Подпись (каз.) | Что это | Ед. | Источник | Дата | Формула | Срезы | Пусто / нет данных | Лучше |', '|---|---|---|---|---|---|---|---|---|---|',
    ...A.METRIC_REG.map(m => `| \`${m.id}\` | ${esc(m.labelKK)} | ${esc(m.labelRU)} | ${esc(m.unit)} | ${SRC[m.source] || m.source} | ${esc(m.dateBasis)} | ${esc(m.formula)} | ${m.supportedDimensions.join(', ') || '—'} | ${esc(m.nullPolicy)} | ${{up: '↑ больше', down: '↓ меньше', neutral: '—'}[m.polarity]} |`), ''].join('\n');
  const f = path.join(ROOT, 'docs', 'метрики.md');
  if (CHECK) { if (!fs.existsSync(f) || fs.readFileSync(f, 'utf8') !== md) fail('docs/метрики.md устарел — запустите сборку'); }
  else { fs.mkdirSync(path.dirname(f), {recursive: true}); fs.writeFileSync(f, md); }
  ok('docs/метрики.md ' + (CHECK ? 'актуален' : 'сгенерирован') + ': ' + A.METRIC_REG.length + ' метрик');
}

// 4. the deploy folder: exactly these files
const plan = [['index.html', 'Атлас_отчёты.html'], ['card.html', 'card.html'], ['demo/crm.json', 'demo/crm.json'], ['demo/sales.json', 'demo/sales.json'], ...fs.readdirSync(path.join(ROOT, 'atlas_v4_assets')).sort().map(f => ['atlas_v4_assets/' + f, 'atlas_v4_assets/' + f])];
const manifest = {note: 'сгенерировано инструменты/собрать.js — не править вручную', methodVersion: (atlas.match(/const METHOD_VERSION = '([^']+)'/) || [])[1] || null,
  snapshots: {crm: (atlas.match(/<script id="crm" type="application\/json">\{[^]*?"snapshot":"([^"]+)"/) || [])[1] || null, sales: (atlas.match(/<script id="sales" type="application\/json">[^]*?"snapshot":"([^"]+)"/) || [])[1] || null},
  files: Object.fromEntries(plan.map(([to, from]) => [to, sha(rd(from))]).concat([['vercel.json', sha(rd('vercel_deploy/vercel.json'))]]))};
const mtext = JSON.stringify(manifest, null, 2) + '\n';
const envs = fs.readdirSync(OUT).filter(f => /^\.env/.test(f));
if (CHECK) {
  const bad = plan.filter(([to, from]) => !fs.existsSync(path.join(OUT, to)) || sha(fs.readFileSync(path.join(OUT, to))) !== sha(rd(from))).map(x => x[0]);
  if (envs.length) bad.push(...envs.map(f => f + ' (секреты!)'));
  if (!fs.existsSync(path.join(OUT, 'BUILD.json')) || fs.readFileSync(path.join(OUT, 'BUILD.json'), 'utf8') !== mtext) bad.push('BUILD.json');
  if (bad.length) fail('vercel_deploy не совпадает со сборкой: ' + bad.join(', ') + ' — запустите node инструменты/собрать.js');
  ok('vercel_deploy совпадает со сборкой'); process.exit(0);
}
envs.forEach(f => { fs.unlinkSync(path.join(OUT, f)); console.log('  удалён ' + f + ' (секреты не публикуются)'); });
fs.mkdirSync(path.join(OUT, 'atlas_v4_assets'), {recursive: true}); fs.mkdirSync(path.join(OUT, 'demo'), {recursive: true});
for (const f of fs.readdirSync(path.join(OUT, 'atlas_v4_assets'))) if (!plan.some(([to]) => to === 'atlas_v4_assets/' + f)) fs.unlinkSync(path.join(OUT, 'atlas_v4_assets', f));
plan.forEach(([to, from]) => fs.writeFileSync(path.join(OUT, to), rd(from)));
fs.writeFileSync(path.join(OUT, 'BUILD.json'), mtext);
if (!fs.existsSync(path.join(OUT, '.vercelignore')) || !/\.env/.test(fs.readFileSync(path.join(OUT, '.vercelignore'), 'utf8'))) fs.writeFileSync(path.join(OUT, '.vercelignore'), '.env*\nBUILD.json\n');
ok('vercel_deploy обновлён: ' + plan.length + ' файлов + BUILD.json');
console.log('  index.html ' + manifest.files['index.html'].slice(0, 12) + ' · card.html ' + manifest.files['card.html'].slice(0, 12));
console.log('Публикация — отдельный шаг: cd vercel_deploy && npx vercel deploy --prod');
