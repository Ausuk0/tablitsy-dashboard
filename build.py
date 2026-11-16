# -*- coding: utf-8 -*-
"""
Сборка двух демо-файлов:

  dashboard.xlsx — дашборд продаж (открывается в Excel и импортируется в Google Таблицы)
  uchet.xlsx     — учёт прихода/расхода и остатков склада в Excel

Все показатели в файлах считаются ФОРМУЛАМИ (СУММЕСЛИМН, СЧЁТЕСЛИМН, СРЗНАЧЕСЛИМН,
ИНДЕКС+ПОИСКПОЗ, НАИБОЛЬШИЙ, ДАТАМЕС). Python только раскладывает тестовые данные
и формулы по листам. Данные генерируются с фиксированным seed — при повторном
запуске получаются те же самые файлы и те же контрольные числа.

Запуск:  python build.py
Проверка формул:  python check.py   (нужен пакет pycel)
"""
import datetime as dt
import random
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent
SEED = 2026

# ---------------------------------------------------------------- оформление
EMERALD = "1E6B57"      # шапки таблиц
EMERALD_SOFT = "E4F2EC"  # плитки KPI, итоговые строки
AUTO_HDR = "7FA99A"     # шапки служебных (формульных) колонок
INPUT = "FFF3BF"        # ячейки, куда пользователь вводит/выбирает значение
RED_SOFT = "F6CFCA"
RED_STRONG = "E8968C"
YELLOW_SOFT = "FFEFB0"
GREEN_SOFT = "CBE8D3"

RUB = '#,##0" ₽"'
DATE = "DD.MM.YYYY"
MONTH_NUM = "MM.YYYY"
PCT = "0%"
PCT1 = "0.0%"
CHG = "▲ 0.0%;▼ 0.0%;0.0%"

THIN = Side(style="thin", color="C9D6D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
             "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

DATA_ROWS = 1000  # до какой строки протянуты формулы и диапазоны (запас под свои данные)
PRODUCT_ROWS = 30  # сколько строк товаров заготовлено на листах «Остатки» и «Итоги» (uchet.xlsx)


def month_label(d):
    return f"{MONTHS_RU[d.month - 1]} {d.year}"


def month_start(d):
    return dt.date(d.year, d.month, 1)


def header(ws, row, titles, fill=EMERALD, start_col=1, height=30):
    for i, t in enumerate(titles):
        c = ws.cell(row=row, column=start_col + i, value=t)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def box(ws, rng, fill=None, bold=False, size=None, fmt=None, align=None, color=None):
    cells = ws[rng] if ":" in rng else ((ws[rng],),)
    for row in cells:
        for c in row:
            c.border = BORDER
            if fill:
                c.fill = PatternFill("solid", fgColor=fill)
            if bold or size or color:
                c.font = Font(bold=bold, size=size or 11, color=color or "000000")
            if fmt:
                c.number_format = fmt
            if align:
                c.alignment = Alignment(horizontal=align, vertical="center")


def merge(ws, rng, value=None, **kw):
    ws.merge_cells(rng)
    first = ws[rng.split(":")[0]]
    if value is not None:
        first.value = value
    box(ws, rng, **kw)
    return first


# ============================================================================
#                         ДАННЫЕ ДЛЯ ДАШБОРДА
# ============================================================================
# (товар, категория, цена, диапазон кол-ва, вес популярности)
PRODUCTS = [
    ("Перфоратор Makita HR2470", "Инструмент", 11490, (1, 2), 6),
    ("Шуруповёрт Bosch GSR 120-LI", "Инструмент", 7990, (1, 3), 8),
    ("Ламинат Tarkett 33 кл., упаковка", "Напольные покрытия", 1890, (4, 24), 10),
    ("Керамогранит 60×60, м²", "Напольные покрытия", 1290, (6, 40), 9),
    ("Краска Tikkurila Euro Power 7, 9 л", "Краски и смеси", 6450, (1, 4), 9),
    ("Грунтовка Ceresit CT 17, 10 л", "Краски и смеси", 1190, (1, 6), 8),
    ("Штукатурка Knauf Ротбанд, 30 кг", "Краски и смеси", 590, (3, 25), 10),
    ("Обои флизелиновые, рулон", "Обои", 1350, (4, 14), 8),
    ("Смеситель Grohe Eurosmart", "Сантехника", 8790, (1, 2), 6),
    ("Унитаз-компакт Cersanit", "Сантехника", 9990, (1, 1), 4),
    ("Дверь межкомнатная, полотно", "Двери и свет", 7450, (1, 4), 6),
    ("Светильник потолочный LED 36 Вт", "Двери и свет", 2390, (1, 6), 7),
]
MANAGERS = ["Иванова Ольга", "Петров Сергей", "Смирнова Анна",
            "Кузнецов Дмитрий", "Волкова Елена", "Морозов Артём"]
MANAGER_W = [1.35, 1.15, 1.0, 0.9, 0.8, 0.55]
CLIENTS = [
    ("Розничный покупатель", 14), ("ООО «СтройДом»", 4), ("ИП Гусев А. Н.", 3),
    ("ООО «Ремонт-Сервис»", 4), ("ИП Ларионова Т. В.", 2), ("ООО «Уютная квартира»", 3),
    ("ТСЖ «Северный»", 1), ("ООО «ПрофОтделка»", 3), ("ИП Карпов Д. С.", 2),
    ("ООО «Мастер на час»", 3), ("ООО «ДекорСтрой»", 2), ("ИП Егорова М. А.", 1),
    ("ООО «АльфаРемонт»", 2), ("Бригада Захарова", 2), ("ООО «Новосёл»", 2),
]
PERIOD_START = dt.date(2026, 8, 1)
PERIOD_END = dt.date(2026, 10, 31)
DEMO_TODAY = dt.date(2026, 11, 16)  # «сегодня» для демо-заявок


def gen_sales(rng):
    rows = []
    d, n = PERIOD_START, 1
    weekday_base = {0: 3.0, 1: 2.7, 2: 2.8, 3: 2.9, 4: 3.3, 5: 3.6, 6: 1.4}
    while d <= PERIOD_END:
        trend = 1 + 0.08 * (d.month - 8)
        k = max(0, round(rng.gauss(weekday_base[d.weekday()] * trend, 1.0)))
        for _ in range(k):
            name, cat, price, (q1, q2), _w = rng.choices(PRODUCTS, weights=[p[4] for p in PRODUCTS])[0]
            qty = rng.randint(q1, q2)
            if rng.random() < 0.15:  # скидка 5 %
                price = round(price * 0.95 / 10) * 10
            manager = rng.choices(MANAGERS, weights=MANAGER_W)[0]
            client = rng.choices([c[0] for c in CLIENTS], weights=[c[1] for c in CLIENTS])[0]
            status = "Отменён" if rng.random() < 0.06 else "Оплачен"
            rows.append(dict(date=d, num=f"ЗК-{n:04d}", client=client, manager=manager,
                             product=name, category=cat, qty=qty, price=price, status=status))
            n += 1
        d += dt.timedelta(days=1)
    return rows


def gen_plan(rng, sales):
    fact = defaultdict(int)
    for r in sales:
        if r["status"] != "Отменён":
            fact[(r["manager"], month_start(r["date"]))] += r["qty"] * r["price"]
    plan = []
    for m in (dt.date(2026, 8, 1), dt.date(2026, 9, 1), dt.date(2026, 10, 1)):
        for man in MANAGERS:
            ratio = rng.uniform(0.82, 1.16)  # насколько менеджер перевыполнит/недовыполнит
            value = max(50000, round(fact[(man, m)] / ratio / 10000) * 10000)
            plan.append(dict(manager=man, month=m, plan=value))
    return plan


def gen_requests(rng):
    """Заявки в работе за ноябрь — лист, с которым работает Apps Script."""
    statuses_open = ["Новая", "В работе", "В работе", "Ожидает оплаты"]
    rows = []
    for i in range(26):
        d = dt.date(2026, 11, 2) + dt.timedelta(days=rng.randint(0, 14))
        srok = d + dt.timedelta(days=rng.randint(2, 9))
        name, _c, price, (q1, q2), _w = rng.choice(PRODUCTS)
        qty = rng.randint(q1, q2)
        if srok < DEMO_TODAY:
            status = rng.choice(["Выполнено", "Выполнено", "В работе", "Ожидает оплаты"])
        else:
            status = rng.choice(statuses_open + ["Выполнено"])
        rows.append(dict(date=d, client=rng.choice([c[0] for c in CLIENTS]),
                         manager=rng.choices(MANAGERS, weights=MANAGER_W)[0],
                         what=f"{name.split(',')[0]} — {qty} шт.", total=qty * price,
                         srok=srok, status=status))
    rows.sort(key=lambda r: r["date"])
    for i, r in enumerate(rows, 1):
        r["num"] = f"Z-{1100 + i}"
    return rows


def gen_archive(rng):
    rows = []
    for i in range(5):
        d = dt.date(2026, 10, 20) + dt.timedelta(days=i * 2)
        name, _c, price, (q1, q2), _w = rng.choice(PRODUCTS)
        qty = rng.randint(q1, q2)
        rows.append(dict(date=d, num=f"Z-{1090 + i}", client=rng.choice([c[0] for c in CLIENTS]),
                         manager=rng.choice(MANAGERS), what=f"{name.split(',')[0]} — {qty} шт.",
                         total=qty * price, srok=d + dt.timedelta(days=4), status="Выполнено",
                         moved=dt.datetime(2026, 10, 31, 20, 5)))
    return rows


def dashboard_data():
    rng = random.Random(SEED)
    sales = gen_sales(rng)
    plan = gen_plan(rng, sales)
    requests = gen_requests(rng)
    archive = gen_archive(rng)
    return sales, plan, requests, archive


# --------------------------------------------- контрольные расчёты в Python
def dashboard_controls(sales, plan, month):
    """Те же показатели, что считают формулы листа «Дашборд», но в Python."""
    ok = [r for r in sales if r["status"] != "Отменён"]
    prev = dt.date(month.year - (month.month == 1), (month.month - 2) % 12 + 1, 1)

    def kpi(m):
        rs = [r for r in ok if month_start(r["date"]) == m]
        rev = sum(r["qty"] * r["price"] for r in rs)
        return rev, len(rs), (rev / len(rs) if rs else 0)

    rev, cnt, avg = kpi(month)
    prev_rev, prev_cnt, prev_avg = kpi(prev)
    plan_total = sum(p["plan"] for p in plan if p["month"] == month)

    prod = []
    for i, p in enumerate(PRODUCTS):
        rs = [r for r in ok if month_start(r["date"]) == month and r["product"] == p[0]]
        prod.append((p[0], sum(r["qty"] for r in rs), sum(r["qty"] * r["price"] for r in rs), i))
    prod.sort(key=lambda x: (-x[2], x[3]))
    top5 = [(n, q, v, v / rev if rev else 0) for n, q, v, _ in prod[:5]]

    mans = []
    for i, m in enumerate(MANAGERS):
        fact = sum(r["qty"] * r["price"] for r in ok if month_start(r["date"]) == month and r["manager"] == m)
        pl = sum(p["plan"] for p in plan if p["month"] == month and p["manager"] == m)
        mans.append((m, pl, fact, i))
    mans.sort(key=lambda x: (-x[2], x[3]))
    managers = [(m, pl, f, f / pl if pl else 0, f - pl) for m, pl, f, _ in mans]

    return dict(month=month, revenue=rev, orders=cnt, avg=avg,
                rev_chg=(rev / prev_rev - 1) if prev_rev else None,
                cnt_chg=(cnt / prev_cnt - 1) if prev_cnt else None,
                avg_chg=(avg / prev_avg - 1) if prev_avg else None,
                plan_total=plan_total, plan_pct=rev / plan_total if plan_total else 0,
                top5=top5, managers=managers,
                canceled=sum(1 for r in sales if r["status"] == "Отменён" and month_start(r["date"]) == month))


def weekly_controls(sales):
    first = PERIOD_START - dt.timedelta(days=PERIOD_START.weekday())
    weeks = []
    w = first
    while w <= PERIOD_END:
        rs = [r for r in sales if r["status"] != "Отменён"
              and w <= r["date"] < w + dt.timedelta(days=7)]
        weeks.append((w, sum(r["qty"] * r["price"] for r in rs), len(rs)))
        w += dt.timedelta(days=7)
    return weeks


# ============================================================================
#                         СБОРКА dashboard.xlsx
# ============================================================================
def build_dashboard(path):
    sales, plan, requests, archive = dashboard_data()
    wb = Workbook()
    dash = wb.active
    dash.title = "Дашборд"
    ws_s = wb.create_sheet("Продажи")
    ws_p = wb.create_sheet("План")
    ws_z = wb.create_sheet("Заявки")
    ws_a = wb.create_sheet("Архив")
    ws_r = wb.create_sheet("Справочник")
    ws_c = wb.create_sheet("Расчёты")

    # ---------------- Справочник
    header(ws_r, 1, ["Месяц", "Начало месяца"], start_col=1)
    for i, m in enumerate(range(1, 13)):
        d = dt.date(2026, m, 1)
        ws_r.cell(row=2 + i, column=1, value=month_label(d))
        ws_r.cell(row=2 + i, column=2, value=d).number_format = DATE
    header(ws_r, 1, ["Товар", "Категория", "Цена, ₽"], start_col=4)
    for i, p in enumerate(PRODUCTS):
        ws_r.cell(row=2 + i, column=4, value=p[0])
        ws_r.cell(row=2 + i, column=5, value=p[1])
        ws_r.cell(row=2 + i, column=6, value=p[2]).number_format = RUB
    header(ws_r, 1, ["Менеджер"], start_col=8)
    for i, m in enumerate(MANAGERS):
        ws_r.cell(row=2 + i, column=8, value=m)
    widths(ws_r, {"A": 18, "B": 15, "C": 3, "D": 36, "E": 20, "F": 12, "G": 3, "H": 22})
    ws_r.freeze_panes = "A2"
    ws_r["J1"] = "Справочник: месяцы для выпадающего списка, товары и менеджеры. Новые товары/менеджеры добавляйте в конец списка."
    ws_r["J1"].font = Font(italic=True, color="5B6B65")

    # ---------------- Продажи (сырые данные + 3 служебные колонки)
    cols = ["Дата", "№ заказа", "Клиент", "Менеджер", "Товар", "Категория",
            "Кол-во", "Цена, ₽", "Статус"]
    header(ws_s, 1, cols)
    header(ws_s, 1, ["Сумма, ₽ (формула)", "Месяц (формула)", "Неделя с пн (формула)"],
           fill=AUTO_HDR, start_col=10)
    for i, r in enumerate(sales):
        row = 2 + i
        vals = [r["date"], r["num"], r["client"], r["manager"], r["product"],
                r["category"], r["qty"], r["price"], r["status"]]
        for j, v in enumerate(vals, 1):
            ws_s.cell(row=row, column=j, value=v)
        ws_s.cell(row=row, column=1).number_format = DATE
        ws_s.cell(row=row, column=8).number_format = RUB
    for row in range(2, DATA_ROWS + 1):
        ws_s.cell(row=row, column=10, value=f'=IF(G{row}="","",G{row}*H{row})').number_format = RUB
        ws_s.cell(row=row, column=11, value=f'=IF(A{row}="","",DATE(YEAR(A{row}),MONTH(A{row}),1))').number_format = MONTH_NUM
        ws_s.cell(row=row, column=12, value=f'=IF(A{row}="","",A{row}-WEEKDAY(A{row},2)+1)').number_format = DATE
        for col in (10, 11, 12):
            ws_s.cell(row=row, column=col).font = Font(color="3F5A50")
    widths(ws_s, {"A": 12, "B": 11, "C": 24, "D": 19, "E": 36, "F": 19, "G": 8,
                  "H": 12, "I": 11, "J": 15, "K": 12, "L": 14})
    ws_s.freeze_panes = "A2"
    ws_s.auto_filter.ref = f"A1:L{len(sales) + 1}"

    # ---------------- План
    header(ws_p, 1, ["Менеджер", "Месяц", "План, ₽"])
    for i, p in enumerate(plan):
        ws_p.cell(row=2 + i, column=1, value=p["manager"])
        ws_p.cell(row=2 + i, column=2, value=p["month"]).number_format = MONTH_NUM
        ws_p.cell(row=2 + i, column=3, value=p["plan"]).number_format = RUB
    widths(ws_p, {"A": 22, "B": 12, "C": 15})
    ws_p.freeze_panes = "A2"
    ws_p["E1"] = "Месяц — первое число месяца (01.10.2026 = октябрь). Одна строка = план одного менеджера на месяц."
    ws_p["E1"].font = Font(italic=True, color="5B6B65")

    # ---------------- Заявки и Архив (с ними работает Apps Script)
    zcols = ["Дата", "№ заявки", "Клиент", "Менеджер", "Что заказано", "Сумма, ₽", "Срок", "Статус", "Комментарий"]
    for ws, rows, extra in ((ws_z, requests, []), (ws_a, archive, ["Перенесено в архив"])):
        header(ws, 1, zcols + extra)
        for i, r in enumerate(rows):
            row = 2 + i
            vals = [r["date"], r["num"], r["client"], r["manager"], r["what"], r["total"], r["srok"], r["status"], ""]
            if extra:
                vals.append(r["moved"])
            for j, v in enumerate(vals, 1):
                ws.cell(row=row, column=j, value=v)
            ws.cell(row=row, column=1).number_format = DATE
            ws.cell(row=row, column=6).number_format = RUB
            ws.cell(row=row, column=7).number_format = DATE
            if extra:
                ws.cell(row=row, column=10).number_format = "DD.MM.YYYY HH:MM"
        widths(ws, {"A": 12, "B": 10, "C": 24, "D": 19, "E": 40, "F": 13, "G": 12, "H": 16, "I": 24, "J": 19})
        ws.freeze_panes = "A2"
    dv = DataValidation(type="list", formula1='"Новая,В работе,Ожидает оплаты,Выполнено"', allow_blank=True)
    dv.add(f"H2:H{DATA_ROWS}")
    ws_z.add_data_validation(dv)
    dv = DataValidation(type="list", formula1="Справочник!$H$2:$H$30", allow_blank=True)
    dv.add(f"D2:D{DATA_ROWS}")
    ws_z.add_data_validation(dv)
    ws_z["K1"] = "Просроченные строки подсвечивает скрипт (меню «Отчёты»), «Выполнено» он переносит в «Архив»."
    ws_z["K1"].font = Font(italic=True, color="5B6B65")

    # ---------------- Расчёты (служебный лист для рейтингов)
    R = "Продажи!"
    rng_ = lambda col: f"{R}${col}$2:${col}${DATA_ROWS}"
    ws_c["A1"], ws_c["A2"] = "Выбранный месяц", "Предыдущий месяц"
    ws_c["B1"] = "=INDEX(Справочник!$B$2:$B$13,MATCH(Дашборд!$C$4,Справочник!$A$2:$A$13,0))"
    ws_c["B2"] = "=EDATE(B1,-1)"
    ws_c["B1"].number_format = ws_c["B2"].number_format = DATE
    ws_c["D1"] = "Служебный лист: суммы по товарам и менеджерам за выбранный месяц и номера строк для рейтингов. Руками не менять."
    ws_c["D1"].font = Font(italic=True, color="5B6B65")
    header(ws_c, 4, ["Товар", "Выручка", "Кол-во", "Ключ сортировки", "", "Место", "Строка"], fill=AUTO_HDR)
    P1, P2 = 5, 34  # до 30 товаров
    for row in range(P1, P2 + 1):
        src = row - P1 + 2
        ws_c[f"A{row}"] = f'=IF(Справочник!$D{src}="","",Справочник!$D{src})'
        ws_c[f"B{row}"] = (f'=IF(A{row}="","",SUMIFS({rng_("J")},{rng_("E")},$A{row},'
                           f'{rng_("K")},$B$1,{rng_("I")},"<>Отменён"))')
        ws_c[f"C{row}"] = (f'=IF(A{row}="","",SUMIFS({rng_("G")},{rng_("E")},$A{row},'
                           f'{rng_("K")},$B$1,{rng_("I")},"<>Отменён"))')
        ws_c[f"D{row}"] = f'=IF(A{row}="","",B{row}-ROW()/1000000)'
        ws_c[f"B{row}"].number_format = RUB
    for k in range(1, 6):
        row = P1 + k - 1
        ws_c[f"F{row}"] = k
        ws_c[f"G{row}"] = f"=IFERROR(MATCH(LARGE($D${P1}:$D${P2},F{row}),$D${P1}:$D${P2},0),\"\")"
    header(ws_c, 37, ["Менеджер", "Факт", "План", "Ключ сортировки", "", "Место", "Строка"], fill=AUTO_HDR)
    M1, M2 = 38, 47  # до 10 менеджеров
    for row in range(M1, M2 + 1):
        src = row - M1 + 2
        ws_c[f"A{row}"] = f'=IF(Справочник!$H{src}="","",Справочник!$H{src})'
        ws_c[f"B{row}"] = (f'=IF(A{row}="","",SUMIFS({rng_("J")},{rng_("D")},$A{row},'
                           f'{rng_("K")},$B$1,{rng_("I")},"<>Отменён"))')
        ws_c[f"C{row}"] = f'=IF(A{row}="","",SUMIFS(План!$C$2:$C$500,План!$A$2:$A$500,$A{row},План!$B$2:$B$500,$B$1))'
        ws_c[f"D{row}"] = f'=IF(A{row}="","",B{row}-ROW()/1000000)'
        ws_c[f"F{row}"] = row - M1 + 1
        ws_c[f"G{row}"] = f"=IFERROR(MATCH(LARGE($D${M1}:$D${M2},F{row}),$D${M1}:$D${M2},0),\"\")"
        ws_c[f"B{row}"].number_format = ws_c[f"C{row}"].number_format = RUB
    widths(ws_c, {"A": 36, "B": 15, "C": 15, "D": 17, "E": 3, "F": 8, "G": 8})
    ws_c.sheet_properties.tabColor = "9AA8A2"

    # ---------------- Дашборд
    d = dash
    d.sheet_view.showGridLines = False
    d.sheet_properties.tabColor = EMERALD
    widths(d, {"A": 2, "B": 15, "C": 36, "D": 12, "E": 16, "F": 11, "G": 3,
               "H": 24, "I": 16, "J": 15, "K": 12, "L": 15, "M": 2})
    t = merge(d, "B2:L2", "Дашборд продаж · магазин «Дом и ремонт»", fill=EMERALD, bold=True, size=18, color="FFFFFF")
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    d.row_dimensions[2].height = 36
    d["B3"] = "Все цифры считаются формулами из листа «Продажи». Выберите месяц в жёлтой ячейке C4 — показатели пересчитаются."
    d["B3"].font = Font(italic=True, color="5B6B65")
    d["B4"] = "Месяц:"
    d["B4"].font = Font(bold=True)
    d["B4"].alignment = Alignment(horizontal="right")
    d["C4"] = month_label(dt.date(2026, 10, 1))
    box(d, "C4", fill=INPUT, bold=True, align="center")
    dv = DataValidation(type="list", formula1="Справочник!$A$2:$A$13", allow_blank=False,
                        showErrorMessage=True, errorTitle="Месяц", error="Выберите месяц из списка")
    dv.add("C4")
    d.add_data_validation(dv)
    d["D4"] = "◀ выпадающий список"
    d["D4"].font = Font(color="5B6B65", italic=True)
    d["H4"] = "Обновлено скриптом:"
    d["H4"].alignment = Alignment(horizontal="right")
    d["I4"] = "—"
    d["I4"].alignment = Alignment(horizontal="center")
    d["J4"] = "Просрочено заявок:"
    d.merge_cells("J4:K4")
    d["J4"].alignment = Alignment(horizontal="right")
    d["L4"] = f'=COUNTIFS(Заявки!$G$2:$G${DATA_ROWS},"<"&TODAY(),Заявки!$H$2:$H${DATA_ROWS},"<>Выполнено")'
    box(d, "L4", bold=True, align="center", color="B03A2E")
    for c in ("H4", "J4"):
        d[c].font = Font(color="5B6B65")

    ok_crit = f'{rng_("I")},"<>Отменён"'
    cur = f'{rng_("K")},Расчёты!$B$1'
    prv = f'{rng_("K")},Расчёты!$B$2'
    plan_sum = "SUMIFS(План!$C$2:$C$500,План!$B$2:$B$500,Расчёты!$B$1)"
    tiles = [  # (label range, value range, change label, change cell, label, formula, fmt)
        ("B6:C6", "B7:C7", "B8", "C8", "Выручка за месяц",
         f"=SUMIFS({rng_('J')},{cur},{ok_crit})", RUB,
         f'=IFERROR(B7/SUMIFS({rng_("J")},{prv},{ok_crit})-1,"")'),
        ("D6:F6", "D7:F7", "D8:E8", "F8", "Заказов (без отмен)",
         f"=COUNTIFS({cur},{ok_crit})", "0",
         f'=IFERROR(D7/COUNTIFS({prv},{ok_crit})-1,"")'),
        ("H6:I6", "H7:I7", "H8", "I8", "Средний чек",
         f"=IFERROR(AVERAGEIFS({rng_('J')},{cur},{ok_crit}),0)", RUB,
         f'=IFERROR(H7/AVERAGEIFS({rng_("J")},{prv},{ok_crit})-1,"")'),
        ("J6:L6", "J7:L7", "J8:K8", "L8", "Выполнение плана",
         f"=IFERROR(B7/{plan_sum},0)", PCT, f"={plan_sum}"),
    ]
    for lab_rng, val_rng, chg_lab, chg_cell, label, formula, fmt, chg in tiles:
        c = merge(d, lab_rng, label, fill=EMERALD_SOFT, color="2F5247", bold=True)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c = merge(d, val_rng, formula, fill=EMERALD_SOFT, bold=True, size=22, fmt=fmt, color="12382E")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        is_plan = chg_cell == "L8"
        lab = "План на месяц:" if is_plan else "к пред. месяцу:"
        if ":" in chg_lab:
            c = merge(d, chg_lab, lab, fill=EMERALD_SOFT, color="5B6B65")
        else:
            d[chg_lab] = lab
            box(d, chg_lab, fill=EMERALD_SOFT, color="5B6B65")
            c = d[chg_lab]
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        d[chg_cell] = chg
        box(d, chg_cell, fill=EMERALD_SOFT, bold=True, fmt=RUB if is_plan else CHG, align="right")
    d.row_dimensions[6].height = 22
    d.row_dimensions[7].height = 42
    d.row_dimensions[8].height = 20
    for cell in ("C8", "F8", "I8"):
        d.conditional_formatting.add(cell, FormulaRule(formula=[f"AND(ISNUMBER({cell}),{cell}>0)"], font=Font(bold=True, color="1E7B4F")))
        d.conditional_formatting.add(cell, FormulaRule(formula=[f"AND(ISNUMBER({cell}),{cell}<0)"], font=Font(bold=True, color="B03A2E")))

    # Топ-5 товаров
    merge(d, "B10:F10", "Топ-5 товаров по выручке за выбранный месяц", bold=True, color="12382E", fill="D3E9DF")
    header(d, 11, ["Место", "Товар", "Кол-во", "Выручка", "Доля"], start_col=2, height=22)
    for k in range(1, 6):
        row = 11 + k
        pos = f"Расчёты!$G${P1 + k - 1}"
        d[f"B{row}"] = k
        d[f"C{row}"] = f'=IFERROR(INDEX(Расчёты!$A${P1}:$A${P2},{pos}),"")'
        d[f"D{row}"] = f'=IFERROR(INDEX(Расчёты!$C${P1}:$C${P2},{pos}),"")'
        d[f"E{row}"] = f'=IFERROR(INDEX(Расчёты!$B${P1}:$B${P2},{pos}),"")'
        d[f"F{row}"] = f'=IFERROR(E{row}/$B$7,"")'
        box(d, f"B{row}:F{row}")
        d[f"B{row}"].alignment = Alignment(horizontal="center")
        d[f"E{row}"].number_format = RUB
        d[f"F{row}"].number_format = PCT1

    # Менеджеры: рейтинг + план/факт
    merge(d, "H10:L10", "Рейтинг менеджеров: план / факт за месяц", bold=True, color="12382E", fill="D3E9DF")
    header(d, 11, ["Менеджер", "План", "Факт", "% плана", "Отклонение"], start_col=8, height=22)
    for k in range(1, 7):
        row = 11 + k
        pos = f"Расчёты!$G${M1 + k - 1}"
        d[f"H{row}"] = f'=IFERROR(INDEX(Расчёты!$A${M1}:$A${M2},{pos}),"")'
        d[f"I{row}"] = f'=IFERROR(INDEX(Расчёты!$C${M1}:$C${M2},{pos}),"")'
        d[f"J{row}"] = f'=IFERROR(INDEX(Расчёты!$B${M1}:$B${M2},{pos}),"")'
        d[f"K{row}"] = f'=IFERROR(J{row}/I{row},"")'
        d[f"L{row}"] = f'=IFERROR(J{row}-I{row},"")'
        box(d, f"H{row}:L{row}")
    d["H18"] = "Итого"
    d["I18"] = "=SUM(I12:I17)"
    d["J18"] = "=SUM(J12:J17)"
    d["K18"] = '=IFERROR(J18/I18,"")'
    d["L18"] = "=J18-I18"
    box(d, "H18:L18", fill=EMERALD_SOFT, bold=True)
    for row in range(12, 19):
        for col in "IJL":
            d[f"{col}{row}"].number_format = RUB
        d[f"K{row}"].number_format = PCT
        d[f"K{row}"].alignment = Alignment(horizontal="center")
    d.conditional_formatting.add("K12:K18", FormulaRule(formula=["AND(ISNUMBER(K12),K12>=1)"], fill=PatternFill("solid", fgColor=GREEN_SOFT), stopIfTrue=True))
    d.conditional_formatting.add("K12:K18", FormulaRule(formula=["AND(ISNUMBER(K12),K12>=0.9)"], fill=PatternFill("solid", fgColor=YELLOW_SOFT), stopIfTrue=True))
    d.conditional_formatting.add("K12:K18", FormulaRule(formula=["AND(ISNUMBER(K12),K12<0.9)"], fill=PatternFill("solid", fgColor=RED_SOFT)))
    d.conditional_formatting.add("L12:L18", FormulaRule(formula=["AND(ISNUMBER(L12),L12<0)"], font=Font(color="B03A2E")))
    d.conditional_formatting.add("L12:L18", FormulaRule(formula=["AND(ISNUMBER(L12),L12>=0)"], font=Font(color="1E7B4F")))
    d.conditional_formatting.add("J7", FormulaRule(formula=["J7>=1"], font=Font(bold=True, size=22, color="1E7B4F")))
    d.conditional_formatting.add("J7", FormulaRule(formula=["J7<0.9"], font=Font(bold=True, size=22, color="B03A2E")))

    # Недели
    merge(d, "B20:F20", "Динамика выручки по неделям (весь период)", bold=True, color="12382E", fill="D3E9DF")
    merge(d, "H20:J20", "Неделя: выручка и заказы", bold=True, color="12382E", fill="D3E9DF")
    header(d, 21, ["Неделя (с пн)", "Выручка", "Заказов"], start_col=8, height=22)
    weeks = weekly_controls(sales)
    W1 = 22
    for i, (w, _rev, _cnt) in enumerate(weeks):
        row = W1 + i
        d[f"H{row}"] = w if i == 0 else f"=H{row - 1}+7"
        d[f"I{row}"] = f"=SUMIFS({rng_('J')},{rng_('L')},H{row},{ok_crit})"
        d[f"J{row}"] = f"=COUNTIFS({rng_('L')},H{row},{ok_crit})"
        box(d, f"H{row}:J{row}")
        d[f"H{row}"].number_format = DATE
        d[f"H{row}"].alignment = Alignment(horizontal="center")
        d[f"I{row}"].number_format = RUB
        d[f"J{row}"].alignment = Alignment(horizontal="center")
    W2 = W1 + len(weeks) - 1

    ch = BarChart()
    ch.type = "col"
    ch.title = "Выручка по неделям, ₽"
    ch.add_data(Reference(d, min_col=9, min_row=21, max_row=W2), titles_from_data=True)
    ch.set_categories(Reference(d, min_col=8, min_row=W1, max_row=W2))
    ch.legend = None
    ch.y_axis.numFmt = "#,##0"
    ch.x_axis.number_format = "DD.MM"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.gapWidth = 60
    ch.series[0].graphicalProperties.solidFill = EMERALD
    ch.series[0].graphicalProperties.line.solidFill = EMERALD
    ch.width, ch.height = 17.5, 7.8
    d.add_chart(ch, "B21")

    # Итоги по месяцам
    MR = W2 + 3
    merge(d, f"H{MR}:L{MR}", "Итоги по месяцам", bold=True, color="12382E", fill="D3E9DF")
    header(d, MR + 1, ["Месяц", "Выручка", "Заказов", "Средний чек", "Отмен"], start_col=8, height=22)
    for i, m in enumerate((8, 9, 10)):
        row = MR + 2 + i
        src = f"Справочник!$B${m + 1}"
        d[f"H{row}"] = f"=Справочник!A{m + 1}"
        d[f"I{row}"] = f"=SUMIFS({rng_('J')},{rng_('K')},{src},{ok_crit})"
        d[f"J{row}"] = f"=COUNTIFS({rng_('K')},{src},{ok_crit})"
        d[f"K{row}"] = f"=IFERROR(I{row}/J{row},0)"
        d[f"L{row}"] = f'=COUNTIFS({rng_("K")},{src},{rng_("I")},"Отменён")'
        box(d, f"H{row}:L{row}")
        d[f"I{row}"].number_format = d[f"K{row}"].number_format = RUB
        for col in "JL":
            d[f"{col}{row}"].alignment = Alignment(horizontal="center")
    row = MR + 5
    d[f"H{row}"] = "Итого"
    d[f"I{row}"] = f"=SUM(I{MR + 2}:I{MR + 4})"
    d[f"J{row}"] = f"=SUM(J{MR + 2}:J{MR + 4})"
    d[f"K{row}"] = f"=IFERROR(I{row}/J{row},0)"
    d[f"L{row}"] = f"=SUM(L{MR + 2}:L{MR + 4})"
    box(d, f"H{row}:L{row}", fill=EMERALD_SOFT, bold=True)
    d[f"I{row}"].number_format = d[f"K{row}"].number_format = RUB
    for col in "JL":
        d[f"{col}{row}"].alignment = Alignment(horizontal="center")

    merge(d, f"B{MR}:F{MR}", "План и факт по менеджерам (выбранный месяц)", bold=True, color="12382E", fill="D3E9DF")
    ch2 = BarChart()
    ch2.type = "bar"
    ch2.title = "План / факт, ₽"
    ch2.add_data(Reference(d, min_col=9, max_col=10, min_row=11, max_row=17), titles_from_data=True)
    ch2.set_categories(Reference(d, min_col=8, min_row=12, max_row=17))
    ch2.x_axis.delete = False
    ch2.y_axis.delete = False
    ch2.y_axis.numFmt = "#,##0"
    ch2.x_axis.scaling.orientation = "maxMin"  # первое место — сверху
    ch2.legend.position = "b"
    ch2.series[0].graphicalProperties.solidFill = "B9CEC5"
    ch2.series[0].graphicalProperties.line.solidFill = "B9CEC5"
    ch2.series[1].graphicalProperties.solidFill = EMERALD
    ch2.series[1].graphicalProperties.line.solidFill = EMERALD
    ch2.width, ch2.height = 17.5, 8.5
    d.add_chart(ch2, f"B{MR + 1}")

    d.freeze_panes = "A5"
    wb.active = 0
    wb.save(path)
    return sales, plan, requests, weeks


# ============================================================================
#                         ДАННЫЕ И СБОРКА uchet.xlsx
# ============================================================================
# артикул, товар, ед., закупка, продажа, мин. остаток, остаток на начало, (продажа за раз), популярность
ZOO = [
    ("ZT-001", "Корм Royal Canin Indoor для кошек, 2 кг", "шт", 1650, 2290, 6, 15, (1, 2), 8),
    ("ZT-002", "Корм Purina Pro Plan для собак, 3 кг", "шт", 1980, 2690, 5, 12, (1, 2), 6),
    ("ZT-003", "Наполнитель древесный «Котофей», 10 л", "шт", 210, 349, 15, 40, (1, 4), 9),
    ("ZT-004", "Наполнитель комкующийся Ever Clean, 6 л", "шт", 890, 1290, 6, 14, (1, 2), 6),
    ("ZT-005", "Влажный корм Whiskas, пауч 75 г", "шт", 22, 39, 60, 200, (4, 24), 10),
    ("ZT-006", "Лакомство Dreamies, 60 г", "шт", 75, 129, 30, 80, (1, 6), 7),
    ("ZT-007", "Шлейка с поводком для кошек", "шт", 260, 490, 5, 12, (1, 1), 3),
    ("ZT-008", "Когтеточка-столбик 60 см", "шт", 690, 1190, 4, 8, (1, 1), 3),
    ("ZT-009", "Лоток закрытый «Домик»", "шт", 780, 1350, 3, 6, (1, 1), 2),
    ("ZT-010", "Миска керамическая 0,3 л", "шт", 160, 290, 8, 20, (1, 2), 4),
    ("ZT-011", "Капли от блох Advantage для кошек", "уп", 520, 790, 10, 25, (1, 3), 5),
    ("ZT-012", "Игрушка-мышь с кошачьей мятой", "шт", 45, 99, 20, 50, (1, 3), 5),
    ("ZT-013", "Переноска пластиковая 48 см", "шт", 1150, 1890, 3, 5, (1, 1), 2),
    ("ZT-014", "Корм для рыб Tetra Min, 100 мл", "шт", 230, 390, 8, 18, (1, 2), 4),
]
SUPPLIERS = ["ООО «ЗооОпт-Регион»", "ООО «Хвостик-Дистрибуция»", "ИП Сафонов А. В."]
BUYERS = ["Розница", "Розница", "Розница", "Розница", "Заказ с сайта", "Заказ с сайта",
          "Ветклиника «Айболит»", "Зоогостиница «Лапки»"]
# часть товаров с какой-то даты не дозаказывали — чтобы на листе «Остатки» была подсветка
NO_RESTOCK = {"ZT-004": dt.date(2026, 10, 5), "ZT-006": dt.date(2026, 8, 1),
              "ZT-011": dt.date(2026, 8, 1), "ZT-013": dt.date(2026, 8, 1)}


def uchet_data():
    rng = random.Random(SEED + 1)
    rng_sup = random.Random(SEED + 2)  # отдельный генератор для поставщиков, чтобы не сбивать продажи
    stock = {p[0]: p[6] for p in ZOO}
    by_art = {p[0]: p for p in ZOO}
    moves = []
    d = PERIOD_START
    while d <= PERIOD_END:
        # приход по понедельникам и четвергам: дозаказ того, что подходит к минимуму
        if d.weekday() in (0, 3):
            for p in ZOO:
                art, minq = p[0], p[5]
                if art in NO_RESTOCK and d >= NO_RESTOCK[art]:
                    continue
                if stock[art] < minq * 1.6:
                    qty = minq * 3 - stock[art]
                    qty = max(minq, int(round(qty / 5.0) * 5) if minq >= 10 else qty)
                    stock[art] += qty
                    moves.append(dict(date=d, op="Приход", art=art, qty=qty, who=rng_sup.choice(SUPPLIERS)))
        # продажи: 1–3 строки в день (в выходные чуть больше)
        k = rng.choice([1, 2, 2, 3] if d.weekday() < 5 else [2, 3, 3])
        for _ in range(k):
            p = rng.choices(ZOO, weights=[z[8] for z in ZOO])[0]
            art = p[0]
            qty = min(rng.randint(*p[7]), stock[art])
            if qty <= 0:
                continue
            stock[art] -= qty
            moves.append(dict(date=d, op="Расход", art=art, qty=qty, who=rng.choice(BUYERS)))
        d += dt.timedelta(days=1)
    for m in moves:
        m["product"] = by_art[m["art"]][1]
    return moves


def uchet_controls(moves):
    by_name = {p[1]: p for p in ZOO}
    res = []
    for p in ZOO:
        inc = sum(m["qty"] for m in moves if m["product"] == p[1] and m["op"] == "Приход")
        out = sum(m["qty"] for m in moves if m["product"] == p[1] and m["op"] == "Расход")
        ost = p[6] + inc - out
        status = "Нет в наличии" if ost <= 0 else ("Заказать" if ost < p[5] else "Норма")
        res.append(dict(art=p[0], name=p[1], start=p[6], inc=inc, out=out, ost=ost, min=p[5],
                        status=status, to_order=max(0, p[5] * 2 - ost) if ost < p[5] else 0,
                        value=ost * p[3]))
    months = []
    for m in (8, 9, 10):
        ms = [x for x in moves if x["date"].month == m]
        buy = sum(x["qty"] * by_name[x["product"]][3] for x in ms if x["op"] == "Приход")
        sold = sum(x["qty"] * by_name[x["product"]][4] for x in ms if x["op"] == "Расход")
        cost = sum(x["qty"] * by_name[x["product"]][3] for x in ms if x["op"] == "Расход")
        cnt = sum(1 for x in ms if x["op"] == "Расход")
        months.append(dict(month=dt.date(2026, m, 1), buy=buy, sold=sold, cost=cost,
                           profit=sold - cost, margin=(sold - cost) / sold if sold else 0, cnt=cnt))
    return res, months


def build_uchet(path):
    moves = uchet_data()
    wb = Workbook()
    ws_m = wb.active
    ws_m.title = "Приход и расход"
    ws_o = wb.create_sheet("Остатки")
    ws_i = wb.create_sheet("Итоги")
    ws_r = wb.create_sheet("Справочник")
    MOV = "'Приход и расход'!"
    mr = lambda col: f"{MOV}${col}$2:${col}${DATA_ROWS}"

    # ---------------- Справочник
    header(ws_r, 1, ["Артикул", "Товар", "Ед.", "Цена закупки, ₽", "Цена продажи, ₽",
                     "Мин. остаток", "Остаток на 01.08.2026"])
    for i, p in enumerate(ZOO):
        row = 2 + i
        for j, v in enumerate(p[:7], 1):
            ws_r.cell(row=row, column=j, value=v)
        ws_r.cell(row=row, column=4).number_format = RUB
        ws_r.cell(row=row, column=5).number_format = RUB
        box(ws_r, f"A{row}:G{row}")
    widths(ws_r, {"A": 10, "B": 42, "C": 6, "D": 16, "E": 16, "F": 13, "G": 15})
    ws_r.freeze_panes = "A2"
    ws_r["I1"] = "Новый товар — добавьте строку в конец списка; он сразу появится в выпадающем списке на листе «Приход и расход»."
    ws_r["I1"].font = Font(italic=True, color="5B6B65")

    # ---------------- Приход и расход
    header(ws_m, 1, ["Дата", "Операция", "Товар", "Кол-во", "Контрагент / комментарий"])
    header(ws_m, 1, ["Цена, ₽ (авто)", "Сумма, ₽ (авто)", "Себестоимость, ₽ (авто)", "Месяц (авто)"],
           fill=AUTO_HDR, start_col=6)
    for i, m in enumerate(moves):
        row = 2 + i
        ws_m.cell(row=row, column=1, value=m["date"]).number_format = DATE
        ws_m.cell(row=row, column=2, value=m["op"])
        ws_m.cell(row=row, column=3, value=m["product"])
        ws_m.cell(row=row, column=4, value=m["qty"])
        ws_m.cell(row=row, column=5, value=m["who"])
    ref = "Справочник!$B$2:$B$200"
    for row in range(2, DATA_ROWS + 1):
        price = (f'=IF(OR(B{row}="",C{row}=""),"",IFERROR(IF(B{row}="Приход",'
                 f'INDEX(Справочник!$D$2:$D$200,MATCH(C{row},{ref},0)),'
                 f'INDEX(Справочник!$E$2:$E$200,MATCH(C{row},{ref},0))),""))')
        ws_m.cell(row=row, column=6, value=price).number_format = RUB
        ws_m.cell(row=row, column=7, value=f'=IF(OR(F{row}="",D{row}=""),"",D{row}*F{row})').number_format = RUB
        ws_m.cell(row=row, column=8, value=(f'=IF(B{row}="Расход",IFERROR(D{row}*INDEX(Справочник!$D$2:$D$200,'
                                             f'MATCH(C{row},{ref},0)),""),"")')).number_format = RUB
        ws_m.cell(row=row, column=9, value=f'=IF(A{row}="","",DATE(YEAR(A{row}),MONTH(A{row}),1))').number_format = MONTH_NUM
        for col in (6, 7, 8, 9):
            ws_m.cell(row=row, column=col).font = Font(color="3F5A50")
    dv_op = DataValidation(type="list", formula1='"Приход,Расход"', allow_blank=True, showErrorMessage=True,
                           errorTitle="Операция", error="Выберите «Приход» или «Расход»")
    dv_op.add(f"B2:B{DATA_ROWS}")
    dv_pr = DataValidation(type="list", formula1="Справочник!$B$2:$B$200", allow_blank=True, showErrorMessage=True,
                           errorTitle="Товар", error="Такого товара нет в справочнике. Добавьте его на лист «Справочник».")
    dv_pr.add(f"C2:C{DATA_ROWS}")
    dv_q = DataValidation(type="whole", operator="greaterThan", formula1="0", allow_blank=True, showErrorMessage=True,
                          errorTitle="Количество", error="Введите целое число больше нуля")
    dv_q.add(f"D2:D{DATA_ROWS}")
    dv_d = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2100,12,31)",
                          allow_blank=True, showErrorMessage=True, errorTitle="Дата", error="Введите дату, например 05.11.2026")
    dv_d.add(f"A2:A{DATA_ROWS}")
    for v in (dv_op, dv_pr, dv_q, dv_d):
        ws_m.add_data_validation(v)
    ws_m.conditional_formatting.add(f"B2:B{DATA_ROWS}", FormulaRule(formula=['B2="Приход"'], font=Font(color="1E7B4F", bold=True)))
    widths(ws_m, {"A": 12, "B": 11, "C": 42, "D": 8, "E": 26, "F": 14, "G": 14, "H": 17, "I": 12})
    ws_m.freeze_panes = "A2"
    ws_m.auto_filter.ref = f"A1:I{len(moves) + 1}"

    # ---------------- Остатки
    header(ws_o, 1, ["Артикул", "Товар", "Ед.", "Остаток на начало", "Приход", "Расход",
                     "Остаток", "Мин. остаток", "Статус", "Рекомендуем заказать", "Стоимость остатка (закуп), ₽"])
    n = PRODUCT_ROWS  # строки протянуты с запасом: новый товар из «Справочника» появится сам
    for i in range(n):
        row, src = 2 + i, 2 + i
        empty = f'Справочник!$B{src}=""'
        ws_o[f"A{row}"] = f'=IF({empty},"",Справочник!A{src})'
        ws_o[f"B{row}"] = f'=IF({empty},"",Справочник!B{src})'
        ws_o[f"C{row}"] = f'=IF({empty},"",Справочник!C{src})'
        ws_o[f"D{row}"] = f'=IF({empty},"",Справочник!G{src})'
        ws_o[f"E{row}"] = f'=IF($B{row}="","",SUMIFS({mr("D")},{mr("C")},$B{row},{mr("B")},"Приход"))'
        ws_o[f"F{row}"] = f'=IF($B{row}="","",SUMIFS({mr("D")},{mr("C")},$B{row},{mr("B")},"Расход"))'
        ws_o[f"G{row}"] = f'=IF($B{row}="","",D{row}+E{row}-F{row})'
        ws_o[f"H{row}"] = f'=IF({empty},"",Справочник!F{src})'
        ws_o[f"I{row}"] = f'=IF($B{row}="","",IF(G{row}<=0,"Нет в наличии",IF(G{row}<H{row},"Заказать","Норма")))'
        ws_o[f"J{row}"] = f'=IF($B{row}="","",IF(G{row}<H{row},H{row}*2-MAX(G{row},0),0))'
        ws_o[f"K{row}"] = f'=IF($B{row}="","",MAX(G{row},0)*Справочник!D{src})'
        box(ws_o, f"A{row}:K{row}")
        ws_o[f"K{row}"].number_format = RUB
        for col in "CDEFGHIJ":
            ws_o[f"{col}{row}"].alignment = Alignment(horizontal="center")
        ws_o[f"G{row}"].font = Font(bold=True)
    last = 1 + n
    tot = last + 1
    ws_o[f"B{tot}"] = "Итого"
    for col in "EFG":
        ws_o[f"{col}{tot}"] = f"=SUM({col}2:{col}{last})"
    ws_o[f"I{tot}"] = f'=COUNTIF(I2:I{last},"Заказать")+COUNTIF(I2:I{last},"Нет в наличии")&" поз. к заказу"'
    ws_o[f"K{tot}"] = f"=SUM(K2:K{last})"
    box(ws_o, f"A{tot}:K{tot}", fill=EMERALD_SOFT, bold=True)
    ws_o[f"K{tot}"].number_format = RUB
    for col in "EFGI":
        ws_o[f"{col}{tot}"].alignment = Alignment(horizontal="center")
    rng_o = f"A2:K{last}"
    ws_o.conditional_formatting.add(rng_o, FormulaRule(formula=['AND($B2<>"",$G2<=0)'], fill=PatternFill("solid", fgColor=RED_STRONG),
                                                       font=Font(bold=True, color="7A1C12"), stopIfTrue=True))
    ws_o.conditional_formatting.add(rng_o, FormulaRule(formula=['AND($B2<>"",$G2<$H2)'], fill=PatternFill("solid", fgColor=RED_SOFT), stopIfTrue=True))
    ws_o.conditional_formatting.add(f"I2:I{last}", FormulaRule(formula=['I2="Норма"'], font=Font(color="1E7B4F")))
    widths(ws_o, {"A": 10, "B": 42, "C": 6, "D": 11, "E": 10, "F": 10, "G": 10, "H": 10,
                  "I": 15, "J": 13, "K": 17, "L": 3, "M": 40})
    ws_o.row_dimensions[1].height = 36
    ws_o.freeze_panes = "C2"
    ws_o["M1"] = "Строка краснеет, когда остаток ниже минимума; ярко-красная — товара нет."
    ws_o["M1"].font = Font(italic=True, color="5B6B65")
    ws_o["M1"].alignment = Alignment(wrap_text=True, vertical="center")

    # ---------------- Итоги
    t = merge(ws_i, "A1:H1", "Итоги по месяцам · зоомагазин", fill=EMERALD, bold=True, size=15, color="FFFFFF")
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_i.row_dimensions[1].height = 30
    header(ws_i, 3, ["Месяц", "Начало месяца", "Закуплено, ₽", "Продано, ₽", "Себестоимость проданного, ₽",
                     "Валовая прибыль, ₽", "Маржа", "Продаж (строк)"])
    # учётный год: август 2026 — июль 2027 (будущие месяцы пока нулевые и заполнятся сами)
    months = [dt.date(2026 + (8 + k > 12), (8 + k - 1) % 12 + 1, 1) for k in range(12)]
    for i, m in enumerate(months):
        row = 4 + i
        ws_i[f"A{row}"] = month_label(m)
        ws_i[f"B{row}"] = m
        ws_i[f"C{row}"] = f'=SUMIFS({mr("G")},{mr("I")},$B{row},{mr("B")},"Приход")'
        ws_i[f"D{row}"] = f'=SUMIFS({mr("G")},{mr("I")},$B{row},{mr("B")},"Расход")'
        ws_i[f"E{row}"] = f'=SUMIFS({mr("H")},{mr("I")},$B{row},{mr("B")},"Расход")'
        ws_i[f"F{row}"] = f"=D{row}-E{row}"
        ws_i[f"G{row}"] = f"=IFERROR(F{row}/D{row},0)"
        ws_i[f"H{row}"] = f'=COUNTIFS({mr("I")},$B{row},{mr("B")},"Расход")'
        box(ws_i, f"A{row}:H{row}")
        ws_i[f"B{row}"].number_format = DATE
        ws_i[f"B{row}"].font = Font(color="7A8A84")
    row = 4 + len(months)
    ws_i[f"A{row}"] = "Итого"
    for col in "CDEFH":
        ws_i[f"{col}{row}"] = f"=SUM({col}4:{col}{row - 1})"
    ws_i[f"G{row}"] = f"=IFERROR(F{row}/D{row},0)"
    box(ws_i, f"A{row}:H{row}", fill=EMERALD_SOFT, bold=True)
    for r in range(4, row + 1):
        for col in "CDEF":
            ws_i[f"{col}{r}"].number_format = RUB
        ws_i[f"G{r}"].number_format = PCT1
        ws_i[f"G{r}"].alignment = ws_i[f"H{r}"].alignment = Alignment(horizontal="center")
    TOT_I = row

    T2 = TOT_I + 3
    merge(ws_i, f"A{T2}:O{T2}", "Продажи по товарам, шт.", bold=True, color="12382E", fill="D3E9DF")
    MCOLS = "CDEFGHIJKLMN"  # 12 месяцев
    header(ws_i, T2 + 1, ["Товар", ""] + [month_label(m) for m in months] + ["Итого"])
    ws_i.merge_cells(f"A{T2 + 1}:B{T2 + 1}")
    for i in range(n):
        row = T2 + 2 + i
        ws_i.merge_cells(f"A{row}:B{row}")
        ws_i[f"A{row}"] = f'=IF(Справочник!$B{2 + i}="","",Справочник!B{2 + i})'
        for j, col in enumerate(MCOLS):
            ws_i[f"{col}{row}"] = (f'=IF($A{row}="","",SUMIFS({mr("D")},{mr("C")},$A{row},{mr("I")},$B${4 + j},'
                                   f'{mr("B")},"Расход"))')
        ws_i[f"O{row}"] = f'=IF($A{row}="","",SUM(C{row}:N{row}))'
        box(ws_i, f"A{row}:O{row}")
        for col in MCOLS + "O":
            ws_i[f"{col}{row}"].alignment = Alignment(horizontal="center")
        ws_i[f"O{row}"].font = Font(bold=True)
    widths(ws_i, {"A": 30, "B": 14, "C": 15, "D": 15, "E": 17, "F": 16, "G": 13, "H": 13,
                  "I": 14, "J": 13, "K": 13, "L": 13, "M": 13, "N": 13, "O": 12})
    ws_i.row_dimensions[3].height = 36
    ws_i.freeze_panes = "A4"

    ch = BarChart()
    ch.type = "col"
    ch.title = "Закупки и продажи по месяцам, ₽"
    ch.add_data(Reference(ws_i, min_col=3, max_col=4, min_row=3, max_row=TOT_I - 1), titles_from_data=True)
    ch.set_categories(Reference(ws_i, min_col=1, min_row=4, max_row=TOT_I - 1))
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.numFmt = "#,##0"
    ch.legend.position = "b"
    ch.series[0].graphicalProperties.solidFill = "B9CEC5"
    ch.series[0].graphicalProperties.line.solidFill = "B9CEC5"
    ch.series[1].graphicalProperties.solidFill = EMERALD
    ch.series[1].graphicalProperties.line.solidFill = EMERALD
    ch.width, ch.height = 17, 7.2
    ws_i.add_chart(ch, "J3")

    ws_o.sheet_properties.tabColor = "C0392B"
    ws_i.sheet_properties.tabColor = EMERALD
    wb.save(path)
    return moves


# ============================================================================
def rub(x):
    return f"{x:,.0f} ₽".replace(",", " ")


def pct(x):
    return "—" if x is None else f"{x * 100:+.1f}%".replace(".", ",")


def main():
    sales, plan, requests, weeks = build_dashboard(HERE / "dashboard.xlsx")
    moves = build_uchet(HERE / "uchet.xlsx")
    print(f"dashboard.xlsx: продаж {len(sales)}, план {len(plan)} строк, заявок {len(requests)}")
    print(f"uchet.xlsx: движений {len(moves)} "
          f"(приход {sum(m['op'] == 'Приход' for m in moves)}, расход {sum(m['op'] == 'Расход' for m in moves)})")
    print()
    for m in (8, 9, 10):
        c = dashboard_controls(sales, plan, dt.date(2026, m, 1))
        print(f"== {month_label(c['month'])}: выручка {rub(c['revenue'])}, заказов {c['orders']}, "
              f"средний чек {rub(c['avg'])}, план {rub(c['plan_total'])} ({c['plan_pct'] * 100:.1f}%), "
              f"к пред. мес.: {pct(c['rev_chg'])} / {pct(c['cnt_chg'])} / {pct(c['avg_chg'])}, отмен {c['canceled']}")
        for i, (n_, q, v, s) in enumerate(c["top5"], 1):
            print(f"   {i}. {n_}: {q} шт., {rub(v)}, {s * 100:.1f}%")
        for n_, pl, f, p, dev in c["managers"]:
            print(f"   - {n_}: план {rub(pl)}, факт {rub(f)}, {p * 100:.0f}%, {rub(dev)}")
    print("\nНедели:")
    for w, v, k in weeks:
        print(f"   {w:%d.%m.%Y}: {rub(v)}, {k}")
    ost, months = uchet_controls(moves)
    print("\nОстатки:")
    for o in ost:
        print(f"   {o['art']} {o['name']}: {o['start']}+{o['inc']}-{o['out']}={o['ost']} (мин {o['min']}) {o['status']}"
              f"{' заказать ' + str(o['to_order']) if o['to_order'] else ''}")
    print(f"   стоимость остатков: {rub(sum(o['value'] for o in ost))}")
    for m in months:
        print(f"   {month_label(m['month'])}: закуплено {rub(m['buy'])}, продано {rub(m['sold'])}, "
              f"себест. {rub(m['cost'])}, прибыль {rub(m['profit'])}, маржа {m['margin'] * 100:.1f}%, продаж {m['cnt']}")


if __name__ == "__main__":
    main()
