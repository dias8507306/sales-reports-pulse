r"""amoCRM export -> de-identified aggregate cubes for the Atlas dashboard.

Only business dimensions leave this script: day, funnel, stage, outcome, responsible manager (approved by the
user), source, region (whitelist), class, package, language, client type, support problem area/type.
Deal IDs, titles, children's names, contacts, phones, links, comments and tags are never read into the output.
Обновление цифр сайта EduApp OP из свежей выгрузки amoCRM.
Запуск:  python обновить_данные.py "путь\к\AMO.xlsx"   (выгрузка главной таблицы AMO: «Сделки», «Воронки и этапы», «Расходы»…)
Скрипт пересчитывает агрегаты и заменяет блок данных в Атлас_отчёты.html (рядом с папкой «инструменты»).
Нужен пакет openpyxl:  pip install openpyxl
"""
import sys, json, datetime as dt
from collections import defaultdict
import openpyxl

import os
src = sys.argv[1]
# второй аргумент (необязательный) — другой HTML, например копия для проверки; по умолчанию рабочий Атлас
html = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Атлас_отчёты.html')
wb = openpyxl.load_workbook(src, read_only=True, data_only=True)

# ---------- stage order ----------
FUNNELS = ['ОП', 'Рассылка', 'Рассылка 5-6 класс Алтын сақа', 'Техподдержка']
stages = {f: [] for f in FUNNELS}
for r in list(wb['Воронки и этапы'].iter_rows(values_only=True))[1:]:
    if not r or r[1] not in stages:
        continue
    stages[r[1]].append({'name': r[4], 'order': r[5], 'type': r[6]})
for f in FUNNELS:
    stages[f].sort(key=lambda s: (s['order'], s['name']))

# ---------- deals ----------
# сделки — только лист «Сделки» выгрузки amoCRM (как в Apps Script; «CRM Лиды — авто» не используется)
if 'Сделки' not in wb.sheetnames:
    sys.exit('Нет листа «Сделки»: нужна выгрузка таблицы amoCRM')
LEADS = 'Сделки'
ws = wb[LEADS]
it = ws.iter_rows(values_only=True)
head = [str(h) for h in next(it)]
ix = {h: i for i, h in enumerate(head)}
ID_COL = next((h for h in ('ID', 'ID сделки', 'Id', 'id', 'Сделка ID') if h in ix), None)
miss_raw = []
MAN_ALIAS = {'Dinara': 'Айгерім'}   # аккаунт amoCRM → человек (владелец 15.10.2026: «Динара — это Айгерім»)

def g(r, k):
    i = ix.get(k)
    return r[i] if i is not None and i < len(r) else None

BASE = dt.date(2026, 8, 10)
def day(v):
    return (v.date() - BASE).days if isinstance(v, dt.datetime) else -1

import re as _re
_REG = [('астана', 'Астана'), ('нурсултан', 'Астана'), ('нұрсұлтан', 'Астана'), ('шымкент', 'Шымкент'), ('алматинск', 'Алматы обл.'), ('алматы', None),
    ('абай', 'Абай'), ('семей', 'Абай'), ('ақмола', 'Ақмола обл.'), ('акмол', 'Ақмола обл.'), ('көкшетау', 'Ақмола обл.'), ('кокшетау', 'Ақмола обл.'), ('ақтөбе', 'Ақтөбе'), ('актоб', 'Ақтөбе'), ('актюб', 'Ақтөбе'), ('атырау', 'Атырау'),
    ('батысқазақстан', 'Батыс Қазақстан'), ('батысказахстан', 'Батыс Қазақстан'), ('западноказахстан', 'Батыс Қазақстан'), ('бқо', 'Батыс Қазақстан'), ('зко', 'Батыс Қазақстан'), ('орал', 'Батыс Қазақстан'), ('уральск', 'Батыс Қазақстан'),
    ('жамбыл', 'Жамбыл'), ('тараз', 'Жамбыл'), ('жетісу', 'Жетісу'), ('жетысу', 'Жетісу'), ('талдықорған', 'Жетісу'), ('талдыкорган', 'Жетісу'), ('қарағанды', 'Қарағанды'), ('караганд', 'Қарағанды'),
    ('қостанай', 'Қостанай'), ('костанай', 'Қостанай'), ('қызылорда', 'Қызылорда'), ('кызылорд', 'Қызылорда'), ('маңғыстау', 'Маңғыстау'), ('мангист', 'Маңғыстау'), ('мангыст', 'Маңғыстау'), ('ақтау', 'Маңғыстау'), ('актау', 'Маңғыстау'),
    ('павлодар', 'Павлодар'), ('солтүстікқазақстан', 'Солтүстік Қазақстан'), ('североказахстан', 'Солтүстік Қазақстан'), ('сқо', 'Солтүстік Қазақстан'), ('ско', 'Солтүстік Қазақстан'), ('петропавл', 'Солтүстік Қазақстан'),
    ('түркістан', 'Түркістан'), ('туркестан', 'Түркістан'), ('ұлытау', 'Ұлытау'), ('улытау', 'Ұлытау'), ('жезқазған', 'Ұлытау'), ('жезказган', 'Ұлытау'),
    ('шығысқазақстан', 'Шығыс Қазақстан'), ('восточноказахстан', 'Шығыс Қазақстан'), ('шқо', 'Шығыс Қазақстан'), ('вко', 'Шығыс Қазақстан'), ('өскемен', 'Шығыс Қазақстан'), ('усткаменогорск', 'Шығыс Қазақстан'), ('устькаменогорск', 'Шығыс Қазақстан')]
def region_of(v):
    """«Облыс» → официальный регион (как zopRegion_ в Apps Script); неизвестное → «Не указан»."""
    raw = str(v if v is not None else '').strip()
    if not raw: return 'Не указан'
    low = raw.lower().replace('ё', 'е'); is_obl = bool(_re.search(r'облыс|обл\.|област|(^|\s)обл(\s|$)', low))
    t = _re.sub(r'облысы|облыс|область|обл\.?|қаласы|каласы|город|г\.|қ\.', ' ', low)
    t = _re.sub(r'[^a-zа-яәғқңөұүһі]+', '', t)
    hit = next((x for x in _REG if t.startswith(x[0])), None) or next((x for x in _REG if len(x[0]) >= 5 and x[0] in t), None)
    if not hit: return 'Не указан'
    return hit[1] if hit[1] else ('Алматы обл.' if is_obl else 'Алматы')

SOURCES = {'Рассылка': 'Рассылка', 'Органика WhatsApp': 'WhatsApp', 'Таргетированная реклама в Инстаграм или Фейсбук': 'Таргет',
           'Органика Instagram': 'Instagram', 'Рекомендация (Сарафан)': 'Сарафан', 'Самост. обращение': 'Самост. обращение'}
OUTCOMES = ['В работе', 'Успешно', 'Не реализовано']
OFFERS = {'Предпродажа': 'Предпродажа', '2+1 Акция': '2+1 Акция'}   # поле «Жеңілдік»; прочее = без акции
# в «CRM Лиды — авто» причина отказа приходит только как ID; названия — как в amoCRM
REASONS = {1000101: 'Слишком дорого', 1000102: 'Пропала потребность', 1000103: 'Не устроили условия', 1000104: 'Выбрали других'}

import re
def budget_of(v):
    '''Бюджет: число или текст «12 000», «990,50 ₸». Пусто → 0; нечитаемое → None (считается в quality.badBudget). Как zopBudget_ в Apps Script.'''
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v) if v == v and abs(v) != float('inf') else None
    if v is None: return 0.0
    t = str(v).strip()
    if not t: return 0.0
    c = re.sub(r'[\s\u00a0₸тгTKZkzt]', '', t, flags=re.I)
    c = re.sub(r',(?=\d{1,2}$)', '.', c).replace(',', '')
    return float(c) if re.fullmatch(r'-?\d+(\.\d+)?', c) else None
def filled(v): return v is not None and str(v).strip() != ''

def offer_of(v):
    """«Жеңілдік» → акция (как zopOffer_ в Apps Script): названия как есть, процент → «Жеңілдік 15%», пусто → «Без акции», прочее → «Басқа акция»"""
    if v is None or str(v).strip() == '': return 'Без акции'
    if isinstance(v, str) and v in OFFERS: return OFFERS[v]
    x = float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None
    if x is None:
        m = re.match(r'^(\d+(?:[.,]\d+)?)\s*%?$', str(v).strip())
        if m: x = float(m.group(1).replace(',', '.'))
    if x is not None and x > 0:
        pct = round(x * 1000) / 10 if x < 1 else x
        if 0 < pct <= 100:
            t = str(int(pct)) if float(pct).is_integer() else str(pct).replace('.', ',')
            return 'Жеңілдік ' + t + '%'
    return 'Басқа акция'

idx = defaultdict(dict)
def code(dim, v):
    d = idx[dim]
    if v not in d:
        d[v] = len(d)
    return d[v]
for f in FUNNELS: code('funnel', f)
for o in OUTCOMES: code('outcome', o)

main = defaultdict(lambda: [0, 0.0])      # (cDay, funnel, stage, outcome, manager, source, clDay, reason, offer, product) -> [n, budget]
imp = defaultdict(lambda: [0, 0.0])       # рассылка по тегу импорта: (импорт, день импорта, воронка, этап, итог, менеджер, день закрытия, продукт) -> [n, budget]
impseg = defaultdict(int)                 # (импорт, сегмент-тег) -> n
impgeo = defaultdict(int)                 # (импорт, облыс, тіл, нәтиже 0/1/2) -> n — как impgeo в Apps Script
impstg = defaultdict(int)                 # (облыс, этап, итог) -> n
MAIL_REPLY = re.compile(r'жауап|первичный контакт|принимают решение', re.I)
idle_raw = []                            # open commercial deals: (manager, funnel, stage, changed, closest task) — aggregated after the snapshot time is known
fill = defaultdict(lambda: [0, 0.0])      # (clDay, funnel, manager, outcome, hasBudget, hasProduct, hasSource, hasReason) -> [n, budget]  (как fill в Apps Script)
hours = defaultdict(int)                  # (weekday, hour, funnel) -> n   (создание сделки)
hsrc = defaultdict(int)                   # (weekday, hour, funnel, source) -> n
buy = defaultdict(int)                    # (weekday, hour, funnel, source) -> n   (закрытие «Успешно» = покупка)
prof = defaultdict(int)                   # (funnel, outcome, region, class, package, lang, client, offer, source) -> n
reg = defaultdict(int)                    # (funnel, outcome, region, created day) -> n — «Аудитория» (15.10, как reg в Apps Script)
clsd = defaultdict(int)                   # (funnel, outcome, class, created day) -> n
supp = defaultdict(int)                   # (cDay, area, type, outcome) -> n
reasons = defaultdict(int)                # (funnel, reason) -> n
rsrc = defaultdict(int)                   # (funnel, source, reason) -> n — себеп арна бойынша
def import_tags(v):
    """«Теги» → (импорты 'YYYY-MM-DD HH:MM', сегменты) — как zopImportTags_ в Apps Script"""
    imps, segs = [], []
    for t in [x.strip() for x in str('' if v is None else v).split(',') if x.strip()]:
        m = re.match(r'^импорт_(\d{2})(\d{2})(\d{4})_(\d{2})(\d{2})$', t, re.I)
        if m:
            k = '%s-%s-%s %s:%s' % (m.group(3), m.group(2), m.group(1), m.group(4), m.group(5))
            if k not in imps: imps.append(k)
            continue
        if re.search(r'\d{5,}', t) or re.match(r'^импорт_', t, re.I) or re.match(r'^wz\b', t, re.I) or len(t) > 40: continue
        if t not in segs: segs.append(t)
    return imps, segs

n_rows, last_change, bad_budget = 0, None, 0
for r in it:
    f = g(r, 'Воронка')
    if f not in stages:
        continue
    n_rows += 1
    created, closed, changed = g(r, 'Создана'), g(r, 'Закрыта'), g(r, 'Изменена')
    if isinstance(changed, dt.datetime) and (last_change is None or changed > last_change):
        last_change = changed
    fu = code('funnel', f)
    st = code('stage', f + '|' + str(g(r, 'Этап')))
    oc = code('outcome', g(r, 'Итог') or 'В работе')
    mg = code('manager', (lambda n: MAN_ALIAS.get(n, n))(str(g(r, 'Ответственный') or 'Не назначен').strip()))   # как ZOP_MAN_ALIAS в Apps Script
    _src_raw = str(g(r, 'Қайдан келді') or '').strip()
    so = code('source', SOURCES.get(_src_raw) or _src_raw[:60] or 'Не указан')   # белгісіз мән — өз атымен, как в Apps Script
    rid = g(r, 'Причина отказа ID')
    reason = REASONS.get(int(rid), 'Причина %d' % int(rid)) if isinstance(rid, (int, float)) else g(r, 'Причина отказа')
    budget = g(r, 'Бюджет CRM') if 'Бюджет CRM' in ix else g(r, 'Бюджет')
    b0, budget = budget, budget_of(budget)
    if budget is None:
        bad_budget += 1 if filled(b0) else 0
        budget = 0.0
    of = code('offer', offer_of(g(r, 'Жеңілдік')))
    prd = g(r, 'Продукт')
    pr = code('product', str(prd).strip() if prd not in (None, '') else 'Не указан')   # amoCRM «Продукт», может быть «A, B»
    tgs = code('tagset', '|'.join(sorted(import_tags(g(r, 'Теги'))[1])))   # теги для фильтра — как в Apps Script
    key = (day(created), fu, st, oc, mg, so, day(closed), 1 if reason else 0, of, pr, tgs, 0)   # 0 — «төленді» белгісі тек Apps Script-те (ОП кестесімен сәйкестендіру)
    main[key][0] += 1
    main[key][1] += budget
    if f != 'Техподдержка':
        _imps, _segs = import_tags(g(r, 'Теги'))
        for k in _imps:
            ic = code('imp', k)
            ik = (ic, (dt.date(int(k[:4]), int(k[5:7]), int(k[8:10])) - BASE).days, fu, st, oc, mg, day(closed), pr)
            imp[ik][0] += 1
            imp[ik][1] += budget
            for sg in _segs:
                impseg[(ic, code('tagseg', sg))] += 1
    if (g(r, 'Итог') or 'В работе') == OUTCOMES[0] and f != 'Техподдержка':
        ch = changed if isinstance(changed, dt.datetime) else created
        tk = g(r, 'Ближайшая задача')
        idle_raw.append((mg, fu, st, ch if isinstance(ch, dt.datetime) else None, tk if isinstance(tk, dt.datetime) else None))
    if isinstance(closed, dt.datetime) and f != 'Техподдержка':   # miss — как zopMissRows_ в Apps Script (только ID сделки)
        _won, _lost_op = (g(r, 'Итог') or '') == OUTCOMES[1], (g(r, 'Итог') or '') == OUTCOMES[2] and f == 'ОП'
        _lid = re.sub(r'\D', '', str(g(r, ID_COL) if ID_COL else '') or '')
        _fl = ((0 if budget > 0 else 1) | (0 if filled(prd) else 2) | (0 if filled(g(r, 'Қайдан келді')) else 4) | (0 if filled(g(r, 'Язык обучения')) else 16) | (0 if filled(g(r, 'Чек')) else 32)) if _won else ((0 if reason else 8) if _lost_op else 0)
        if _fl and _lid: miss_raw.append([int(_lid), day(closed), mg, 0 if _won else 1, _fl])
    if isinstance(closed, dt.datetime):
        fk = (day(closed), fu, mg, oc, 1 if budget > 0 else 0, 1 if filled(prd) else 0, 1 if filled(g(r, 'Қайдан келді')) else 0, 1 if reason else 0)
        fill[fk][0] += 1
        fill[fk][1] += budget
    if isinstance(created, dt.datetime):
        hours[(created.weekday(), created.hour, fu)] += 1
        hsrc[(created.weekday(), created.hour, fu, so)] += 1
    if OUTCOMES[1] == (g(r, 'Итог') or '') and isinstance(closed, dt.datetime):
        buy[(closed.weekday(), closed.hour, fu, so)] += 1
    cls = g(r, 'Класс')
    cls = int(cls) if isinstance(cls, (int, float)) and 0 < cls < 13 else 0
    pk = g(r, 'Пакет')
    pk = 'AltynSaqa' if pk == 'AltynSaqa' else ('Другое' if pk else 'Не указан')
    lang = g(r, 'Язык обучения') if g(r, 'Язык обучения') in ('Қазақша', 'Русский') else 'Не указан'
    cl = {'Жаңа клиент': 'Новый', 'Ескі клиент': 'Повторный'}.get(g(r, 'Клиент'), 'Не указан')
    reg[(fu, oc, code('region', region_of(g(r, 'Облыс'))), day(created))] += 1
    clsd[(fu, oc, cls, day(created))] += 1
    prof[(fu, oc, code('region', region_of(g(r, 'Облыс'))), cls, code('package', pk), code('lang', lang), code('client', cl), of, so)] += 1
    if f != 'Техподдержка':
        _res = 2 if (g(r, 'Итог') or 'В работе') == OUTCOMES[1] else 1 if f == 'ОП' or MAIL_REPLY.search(str(g(r, 'Этап'))) else 0
        _rg, _lg = code('region', region_of(g(r, 'Облыс'))), code('lang', lang)
        _ims = import_tags(g(r, 'Теги'))[0]
        for k in _ims:
            impgeo[(code('imp', k), _rg, _lg, _res)] += 1
        if _ims: impstg[(_rg, st, oc)] += 1   # өңір × этап (как impstg в Apps Script)
    if f == 'Техподдержка':
        supp[(day(created), code('parea', g(r, 'problem_area') or 'Не указано'), code('ptype', g(r, 'problem_type') or 'Не указано'), oc)] += 1
    if reason:
        rc = code('reason', str(reason))
        reasons[(fu, rc)] += 1
        rsrc[(fu, so, rc)] += 1

# журнал синхронизации: «Время» (+ «Режим» в старом листе «Журнал»; в «CRM Журнал — авто» режима нет)
_lr = list(wb['CRM Журнал — авто' if 'CRM Журнал — авто' in wb.sheetnames else 'Журнал'].iter_rows(values_only=True))
_lh = [str(x) for x in (_lr[0] if _lr else [])]
_lt, _lm = (_lh.index('Время') if 'Время' in _lh else 0), (_lh.index('Режим') if 'Режим' in _lh else -1)
log = [(r[_lt], r[_lm] if _lm >= 0 else '') for r in _lr[1:] if r and len(r) > _lt and isinstance(r[_lt], dt.datetime)]

# ---------- лист «Рассылки» (вносится вручную: тема, даты, воронка, расход) ----------
import re
def btn_list(v):
    """Батырмалар: әр жолға «Мәтін» немесе «Мәтін | https://…»; ең көбі 3 (Apps Script zopBtnList_ сияқты)."""
    out = []
    for line in re.split(r'\r?\n', str(v or '')):
        parts = [x.strip() for x in line.split('|')]
        t = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', parts[0] if parts else '').strip()[:2000]
        t = t[:40]
        if not t:
            continue
        u = parts[1] if len(parts) > 1 else ''
        out.append({'t': t, 'u': u} if re.match(r'^https://[^\s<>"]{4,500}$', u) else {'t': t})
    return out[:3]


def mailings_from(ws):
    rows = [r for r in ws.iter_rows(values_only=True)]
    if not rows: return []
    head = [str(x or '').strip().lower() for x in rows[0]]
    def col(rx):
        for i, h in enumerate(head):
            if re.search(rx, h): return i
        return -1
    c = {'name': col(r'^рассылка|назв'), 'topic': col(r'^тема'), 'a': col(r'начал|^дата$'), 'b': col(r'оконч'), 'f': col(r'воронк'), 'seg': col(r'кому|сегмент'),
         'cost': col(r'расход|затрат|сумм'), 'sent': col(r'отправ|охват'), 'note': col(r'коммент'), 'text': col(r'мәтін|текст'), 'media': col(r'медиа|сурет|видео|media'), 'btn': col(r'батырма|кнопк|button')}
    def get(r, k): return r[c[k]] if c[k] >= 0 and c[k] < len(r) else ''
    def num(v):
        if isinstance(v, (int, float)): return v
        t = re.sub(r'[^\d.,-]', '', str(v or '')).replace(',', '.')
        try: return float(t) if t else None
        except ValueError: return None
    def dayv(v):
        if isinstance(v, dt.datetime): return (v.date() - BASE).days
        m = re.search(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', str(v or ''))
        return (dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) - BASE).days if m else None
    def fun(v):
        t = str(v or '').lower()
        if re.search(r'5\s*[-–]\s*6', t): return FUNNELS.index('Рассылка 5-6 класс Алтын сақа')
        return FUNNELS.index('Рассылка') if 'рассыл' in t else -1
    out = []
    for r in rows[1:]:
        m = {'name': str(get(r, 'name') or '').strip(), 'topic': str(get(r, 'topic') or '').strip(), 'a': dayv(get(r, 'a')), 'b': dayv(get(r, 'b')), 'f': fun(get(r, 'f')),
             'seg': str(get(r, 'seg') or '').strip(), 'cost': num(get(r, 'cost')), 'sent': num(get(r, 'sent')), 'note': str(get(r, 'note') or '').strip(),
             'text': str(get(r, 'text') or '').strip()[:2000], 'media': [x for x in re.split(r'[\s,]+', str(get(r, 'media') or '')) if re.match(r'^https://[^\s<>"]{4,500}$', x)][:4],
             'buttons': btn_list(get(r, 'btn'))}
        for k in ('cost', 'sent'):
            if isinstance(m[k], float) and m[k].is_integer(): m[k] = int(m[k])
        if m['a'] is not None and (m['name'] or m['topic']): out.append(m)
    return out
mailings = mailings_from(wb['Рассылки']) if 'Рассылки' in wb.sheetnames else []

# ---------- лист «Расходы» (или файл Расходы.csv рядом со скриптом, пока листа нет) ----------
# ---------- «Расходы» v2 — тот же разбор, что zopCosts_ в Apps Script (tests/python_parity.test.cjs сверяет) ----------
COST_CATS = ['Wazzup top-up', 'Wazzup subscription', 'WhatsApp template charges', 'Provider service fees', 'Sipuni subscription', 'Marketing campaign costs', 'Other sales expenses', 'Wallet opening balance', 'Adjustment']
COST_KIND = {'Wazzup top-up': 'balance', 'Wazzup subscription': 'plan', 'Sipuni subscription': 'plan', 'WhatsApp template charges': 'usage', 'Provider service fees': 'usage',
             'Marketing campaign costs': 'direct', 'Other sales expenses': 'direct', 'Wallet opening balance': 'opening', 'Adjustment': 'adjust'}
def cost_alloc(text, total):
    parts = [x.strip() for x in str(text or '').split(';') if x.strip()]
    if not parts: return {'items': [], 'err': None}
    if len(parts) == 1 and not re.search(r'[=:]', parts[0]): return {'items': [{'name': parts[0], 'sum': total}], 'err': None}
    items = []
    for x in parts:
        m = re.match(r'^(.+?)\s*[=:]\s*([\d\s.,]+)\s*(%?)$', x)
        if not m:
            items.append({'name': x, 'sum': None, 'bad': True}); continue
        v = float(re.sub(r'\s', '', m.group(2)).replace(',', '.'))
        items.append({'name': m.group(1).strip(), 'sum': (None if total is None else total * v / 100) if m.group(3) else v, 'pct': v if m.group(3) else None})
    if any(x.get('bad') for x in items): return {'items': items, 'err': 'кампания жолы оқылмады'}
    if total is None: return {'items': items, 'err': None}
    s = sum(x['sum'] or 0 for x in items)
    return {'items': items, 'err': ('бөлінген %d ≠ сома %d' % (round(s), round(total))) if abs(s - total) > 0.5 else None}
def costs_from(rows, tariffs=None):
    if not rows: return []
    head = [str(x or '').strip().lower() for x in rows[0]]
    def col(rx):
        for i, h in enumerate(head):
            if re.search(rx, h): return i
        return -1
    c = {'d': col(r'^дата|^date'), 'sub': col(r'подпис|subscr|канал'), 'det': col(r'детал|detail|описан'), 'sum': col(r'^сумм|^sum|сома'), 'pay': col(r'^оплат|payment method|^төлем'), 'mail': col(r'^рассылк'),
         'cat': col(r'категор|category'), 'prov': col(r'поставщ|провайдер|provider'), 'cur': col(r'валют|currency'), 'rate': col(r'^курс|rate'), 'from': col(r'кезең басы|начало периода|period from'), 'to': col(r'кезең соңы|конец периода|period to'),
         'camp': col(r'кампан|campaign'), 'msgs': col(r'хабарлама саны|сообщени|messages'), 'tariff': col(r'^тариф$|tariff'), 'doc': col(r'құжат|документ|document'), 'ver': col(r'тексерілді|провер|verified'), 'id': col(r'^id$')}
    def get(r, k): return r[c[k]] if c[k] >= 0 and c[k] < len(r) else ''
    def s_(r, k):
        v = get(r, k)
        if isinstance(v, float) and v.is_integer(): v = int(v)   # 7000 из таблицы — число, не «7000.0»
        return str('' if v is None else v).strip()
    def num(v):
        if isinstance(v, bool): return None
        if isinstance(v, (int, float)): return v
        t = re.sub(r'[^\d.,-]', '', str(v if v is not None else '')).replace(',', '.')
        try: return float(t) if t else None
        except ValueError: return None
    def dayv(v):
        if isinstance(v, dt.datetime): return (v.date() - BASE).days
        m = re.search(r'(\d{1,2})[./](\d{1,2})[./](\d{2,4})', str(v or ''))
        if not m: return None
        y = int(m.group(3)); y = 2000 + y if len(m.group(3)) == 2 else y
        return (dt.date(y, int(m.group(2)), int(m.group(1))) - BASE).days
    def tariff_at(prov, typ, d):
        L = [t for t in (tariffs or []) if t['prov'].lower() == str(prov or '').lower() and (not typ or not t['type'] or t['type'].lower() == str(typ).lower()) and t['from'] <= d and (t['to'] is None or d <= t['to'])]
        return sorted(L, key=lambda t: -t['from'])[0] if L else None
    out, seen = [], set()
    for r in rows[1:]:
        det, sub = s_(r, 'det'), s_(r, 'sub'); plan = bool(re.search(r'plan|тариф', det, re.I)); d = dayv(get(r, 'd'))
        cat = s_(r, 'cat')
        if cat not in COST_CATS:
            cat = cat or ('Sipuni subscription' if re.search('sipuni', sub + ' ' + det, re.I) else 'Wazzup subscription' if plan else 'Wazzup top-up' if re.search(r'balance|баланс', det, re.I) else 'Other sales expenses')
        kind, cur, rate, raw, msgs = COST_KIND.get(cat, 'direct'), (s_(r, 'cur') or 'KZT').upper(), num(get(r, 'rate')), num(get(r, 'sum')), num(get(r, 'msgs'))
        prov, warn, tariff = s_(r, 'prov') or ('Sipuni' if re.search('sipuni', sub, re.I) else 'Wazzup'), [], s_(r, 'tariff')
        sm, src = raw, ('row' if raw is not None else None)
        if sm is None and msgs is not None and d is not None:
            t = tariff_at(prov, tariff, d)
            if t:
                sm, src = msgs * t['price'], 'tariff'
                if t['cur'] != 'KZT':
                    sm = sm * rate if rate else None
                    if not rate: warn.append('тариф валютасы ' + t['cur'] + ', курс жоқ')
            else:
                warn.append('сома да, тариф та жоқ')
        elif sm is not None and cur != 'KZT':
            if rate: sm = sm * rate
            else:
                sm = None; warn.append('валюта ' + cur + ', курс жоқ')
        elif sm is None:
            warn.append('сома жоқ')
        if isinstance(sm, float) and sm.is_integer(): sm = int(sm)
        months = (12 if re.search(r'year|год|жыл', det, re.I) else 1) if kind == 'plan' else 0
        rid = s_(r, 'id')
        raw_s = '' if raw is None else (str(int(raw)) if float(raw).is_integer() else str(raw))
        key = rid or '|'.join(['' if d is None else str(d), cat, prov, sub, raw_s, cur, s_(r, 'pay'), s_(r, 'doc'), s_(r, 'camp')])
        dup = key in seen; seen.add(key)
        al = cost_alloc(s_(r, 'camp'), sm)
        if al['err']: warn.append(al['err'])
        x = {'d': d, 'sub': sub, 'det': det, 'sum': sm, 'pay': s_(r, 'pay'), 'mail': s_(r, 'mail'), 'kind': kind, 'months': months, 'cat': cat, 'prov': prov, 'cur': cur, 'rate': rate,
             'from': dayv(get(r, 'from')), 'to': dayv(get(r, 'to')), 'camp': s_(r, 'camp'), 'alloc': al['items'], 'msgs': msgs, 'tariff': tariff, 'doc': s_(r, 'doc'),
             'ver': bool(re.match(r'^(иә|да|yes|true|1|✓)$', s_(r, 'ver'), re.I)), 'id': rid, 'src': src, 'dup': dup, 'warn': warn}
        if d is not None and (sm is not None or warn or msgs is not None): out.append(x)
    return out
import os, csv as _csv
_csvp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Расходы.csv')
if 'Расходы' in wb.sheetnames: costs = costs_from([list(r) for r in wb['Расходы'].iter_rows(values_only=True)])
elif os.path.exists(_csvp): costs = costs_from([r for r in _csv.reader(open(_csvp, encoding='utf-8'), delimiter=';')])
else: costs = []
log.sort(key=lambda r: r[0])

# ---------- лист «KPI»: цели на месяц (Ай · Бағыт · Көрсеткіш · Мақсат · необязательные фильтры) ----------
def kpi_from(rows):
    if not rows: return []
    head = [str(x or '').strip().lower() for x in rows[0]]
    def col(rx):
        for i, h in enumerate(head):
            if re.search(rx, h): return i
        return -1
    c = {'mo': col(r'^ай'), 'a': col(r'^бағыт'), 'k': col(r'^көрсеткіш'), 't': col(r'^мақсат'), 'src': col(r'^арна'), 'of': col(r'^акция'), 'pr': col(r'^өнім'), 'm': col(r'^менеджер')}
    def get(r, k): return r[c[k]] if c[k] >= 0 and c[k] < len(r) else None
    def txt(r, k):
        v = get(r, k)
        if v is None: return ''
        if isinstance(v, float) and v.is_integer(): v = int(v)
        return str(v).strip()
    def month(v):
        if isinstance(v, (dt.datetime, dt.date)): return '%d-%02d' % (v.year, v.month)
        t = str(v if v is not None else '').strip()
        m = re.match(r'^(\d{4})[-./](\d{1,2})', t)
        if m: return '%s-%02d' % (m.group(1), int(m.group(2)))
        m = re.match(r'^(?:\d{1,2}[./])?(\d{1,2})[./](\d{4})$', t)
        return '%s-%02d' % (m.group(2), int(m.group(1))) if m else ''
    def num(v):
        if isinstance(v, (int, float)): x = v
        else:
            t = re.sub(r'[^\d.,-]', '', str(v if v is not None else '')).replace(',', '.')
            try: x = float(t) if str(v if v is not None else '').strip() != '' else None
            except ValueError: x = None
        if isinstance(x, float) and x.is_integer(): x = int(x)
        return x
    out = []
    for r in rows[1:]:
        k = {'month': month(get(r, 'mo')) if c['mo'] >= 0 else '', 'area': txt(r, 'a'), 'metric': txt(r, 'k'), 'target': num(get(r, 't')) if c['t'] >= 0 else None,
             'src': txt(r, 'src'), 'of': txt(r, 'of'), 'pr': txt(r, 'pr'), 'm': txt(r, 'm')}
        if k['month'] and k['area'] and k['metric'] and k['target'] is not None and k['target'] > 0: out.append(k)
    return out
# idle: [manager, funnel, stage, days without change at the snapshot (60 = 60+; -1 = no date), task: 0 none · 1 planned · 2 overdue, n] — как в Apps Script
# privacy: profile cells of 1–2 children → quasi-identifiers masked (same rule as Apps Script)
PROF_MIN, prof_safe = 3, defaultdict(int)
for k, v in prof.items():
    if v >= PROF_MIN: prof_safe[k] += v
    else: prof_safe[(k[0], k[1], code('region', 'Жасырылған (< 3)'), 0, code('package', 'Не указан'), code('lang', 'Не указан'), code('client', 'Не указан'), k[7], k[8])] += v
reg_safe = defaultdict(int)               # өңір графигі: 3-тен аз өңір → «Жасырылған (< 3)»
_rt, _ct = defaultdict(int), defaultdict(int)
for k, v in reg.items(): _rt[k[2]] += v
for k, v in clsd.items(): _ct[k[2]] += v
for k, v in reg.items(): reg_safe[k if _rt[k[2]] >= PROF_MIN else (k[0], k[1], code('region', 'Жасырылған (< 3)'), k[3])] += v
cls_safe = defaultdict(int)
for k, v in clsd.items(): cls_safe[k if _ct[k[2]] >= PROF_MIN else (k[0], k[1], -1, k[3])] += v
idle = defaultdict(int)
if log:
    _snap = log[-1][0]
    for mg_, fu_, st_, ch_, tk_ in idle_raw:
        d_ = -1 if ch_ is None else min(60, max(0, int((_snap - ch_).total_seconds() // 86400)))
        idle[(mg_, fu_, st_, d_, 0 if tk_ is None else (2 if tk_ < _snap else 1))] += 1
_tt, _tr = defaultdict(int), defaultdict(int)
for k, v in impgeo.items():
    _tt[k[0]] += v
    if k[3] >= 1: _tr[k[0]] += v
impreg = defaultdict(int)                # (облыс, нәтиже, тікелей импорт 0/1) -> n — как impreg в Apps Script
for k, v in impgeo.items(): impreg[(k[1], k[3], 1 if _tr[k[0]] / _tt[k[0]] >= .9 else 0)] += v
impreg_safe = defaultdict(int)
impstg_safe = defaultdict(int)
for k, v in impstg.items(): impstg_safe[k if v >= 3 else (code('region', 'Жасырылған (< 3)'), k[1], k[2])] += v
for k, v in impreg.items(): impreg_safe[k if v >= 3 else (code('region', 'Жасырылған (< 3)'), k[1], k[2])] += v
impgeo_safe = defaultdict(int)          # 1–2 сделки в ячейке → облыс скрыт, тіл «Не указан» (как в Apps Script)
for k, v in impgeo.items():
    impgeo_safe[k if v >= 3 else (k[0], code('region', 'Жасырылған (< 3)'), code('lang', 'Не указан'), k[3])] += v
impseg_safe = defaultdict(int)          # сегмент импорта меньше 3 сделок → «Басқа (< 3)»
for k, v in impseg.items():
    impseg_safe[k if v >= 3 else (k[0], code('tagseg', 'Басқа (< 3)'))] += v
kpi = kpi_from([list(r) for r in wb['KPI'].iter_rows(values_only=True)]) if 'KPI' in wb.sheetnames else []
data = {
    'base': BASE.isoformat(),
    'schemaVersion': 2, 'methodVersion': 'crm-2026-10-10',
    'rows': n_rows,
    'quality': {'badBudget': bad_budget},
    'snapshot': log[-1][0].isoformat(timespec='minutes') if log else None,
    'lastChange': last_change.isoformat(timespec='minutes') if last_change else None,
    'sync': {'runs': len(log), 'full': sum(1 for r in log if r[1] == 'Полная'), 'first': log[0][0].isoformat(timespec='minutes') if log else None,
             'gaps': [round((b[0] - a[0]).total_seconds() / 60) for a, b in zip(log, log[1:])]},
    'dict': {k: [name for name, _ in sorted(v.items(), key=lambda kv: kv[1])] for k, v in idx.items()},
    'stages': {f: [s['name'] for s in stages[f]] for f in FUNNELS},
    'main': [list(k) + [v[0], round(v[1])] for k, v in sorted(main.items())],
    'imp': [list(k) + [v[0], round(v[1])] for k, v in sorted(imp.items())],
    'impseg': [list(k) + [v] for k, v in sorted(impseg_safe.items())],
    'impgeo': [list(k) + [v] for k, v in sorted(impgeo_safe.items())],
    'impreg': [list(k) + [v] for k, v in sorted(impreg_safe.items())],
    'impstg': [list(k) + [v] for k, v in sorted(impstg_safe.items())],
    'fill': [list(k) + [v[0], round(v[1])] for k, v in sorted(fill.items())],
    'idle': [list(k) + [v] for k, v in sorted(idle.items())],
    'hours': [list(k) + [v] for k, v in sorted(hours.items())],
    'hsrc': [list(k) + [v] for k, v in sorted(hsrc.items())],
    'buy': [list(k) + [v] for k, v in sorted(buy.items())],
    'prof': [list(k) + [v] for k, v in sorted(prof_safe.items())],
    'reg': [list(k) + [v] for k, v in sorted(reg_safe.items())],
    'clsd': [list(k) + [v] for k, v in sorted(cls_safe.items())],
    'supp': [list(k) + [v] for k, v in sorted(supp.items())],
    'reasons': [list(k) + [v] for k, v in sorted(reasons.items())],
    'miss': (lambda raw: sorted([x for x in raw if x[1] >= max([y[1] for y in raw] or [0]) - 45], key=lambda x: (-x[1], -x[0]))[:3000])(miss_raw),
    'rsrc': [list(k) + [v] for k, v in sorted(rsrc.items())],
    'mailings': mailings,
    'costs': costs,
    'kpi': kpi,
}
payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
with open(html, encoding='utf-8') as fh:
    page = fh.read()
start = page.index('<script id="crm" type="application/json">') + len('<script id="crm" type="application/json">')
end = page.index('</script>', start)
with open(html, 'w', encoding='utf-8', newline='') as fh:
    fh.write(page[:start] + payload + page[end:])
print('Готово: данные обновлены в', os.path.normpath(html))
print('rows', n_rows, 'main cells', len(main), 'prof', len(prof), 'hours', len(hours), 'supp', len(supp))
print('dicts', {k: len(v) for k, v in data['dict'].items()})
print('snapshot', data['snapshot'], 'lastChange', data['lastChange'])
