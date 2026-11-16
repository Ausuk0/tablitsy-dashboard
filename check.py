# -*- coding: utf-8 -*-
"""
Проверка формул: файлы пересчитываются движком pycel (вычисляет формулы Excel
в Python), а результат сравнивается с независимым расчётом тех же показателей
по «сырым» данным, прочитанным из самих xlsx.

Запуск (после python build.py):  python check.py
"""
import datetime as dt
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

try:
    from pycel import ExcelCompiler
except ImportError:
    sys.exit("Нужен pycel:  pip install pycel")

import pycel.lib.date_time as _pycel_dt


def _weekday(serial, return_type=1):
    """В pycel ДЕНЬНЕД (WEEKDAY) без второго аргумента — добавляем типы 1, 2, 3 как в Excel."""
    if isinstance(serial, (dt.date, dt.datetime)):
        d = serial if isinstance(serial, dt.date) and not isinstance(serial, dt.datetime) else serial.date()
    else:  # пустая ячейка = 0 (pycel вычисляет обе ветки ЕСЛИ, поэтому сюда попадают и пустые строки)
        d = dt.date(1899, 12, 30) + dt.timedelta(days=int(serial or 0))
    py = d.weekday()  # пн = 0
    return {1: (py + 1) % 7 + 1, 2: py + 1, 3: py}[int(return_type)]


_pycel_dt.weekday = _weekday

HERE = Path(__file__).resolve().parent
MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
             "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
errors = 0
checked = 0


def same(label, got, exp, tol=0.005):
    global errors, checked
    checked += 1
    if isinstance(exp, (int, float)) and not isinstance(exp, bool):
        ok = isinstance(got, (int, float)) and abs(got - exp) <= tol * max(1, abs(exp)) * 0.001 + 1e-6
    else:
        ok = got == exp
    if not ok:
        errors += 1
        print(f"  НЕ СОВПАЛО  {label}: формула={got!r}  python={exp!r}")


def as_date(v):
    return v.date() if isinstance(v, dt.datetime) else v


def ms(d):
    return dt.date(d.year, d.month, 1)


# ============================================================ dashboard.xlsx
def check_dashboard():
    path = HERE / "dashboard.xlsx"
    wb = load_workbook(path)
    sales = []
    for r in wb["Продажи"].iter_rows(min_row=2, max_col=9, values_only=True):
        if r[0] is None:
            break
        sales.append(dict(date=as_date(r[0]), manager=r[3], product=r[4], qty=r[6], price=r[7], status=r[8]))
    plan = [dict(manager=r[0], month=as_date(r[1]), plan=r[2])
            for r in wb["План"].iter_rows(min_row=2, max_col=3, values_only=True) if r[0]]
    products = [r[0] for r in wb["Справочник"].iter_rows(min_row=2, min_col=4, max_col=4, values_only=True) if r[0]]
    managers = [r[0] for r in wb["Справочник"].iter_rows(min_row=2, min_col=8, max_col=8, values_only=True) if r[0]]
    req = [dict(srok=as_date(r[6]), status=r[7])
           for r in wb["Заявки"].iter_rows(min_row=2, max_col=9, values_only=True) if r[0]]
    ok = [s for s in sales if s["status"] != "Отменён"]
    print(f"dashboard.xlsx: {len(sales)} продаж, {len(products)} товаров, {len(managers)} менеджеров")

    xl = ExcelCompiler(filename=str(path))
    ev = lambda a: xl.evaluate(f"Дашборд!{a}")
    ev("B7")  # загружает в pycel цепочку зависимостей, включая ячейку выбора месяца C4
    ev("C4")

    for month in (dt.date(2026, 8, 1), dt.date(2026, 9, 1), dt.date(2026, 10, 1)):
        label = f"{MONTHS_RU[month.month - 1]} {month.year}"
        xl.set_value("Дашборд!C4", label)
        prev = dt.date(month.year - (month.month == 1), (month.month - 2) % 12 + 1, 1)
        cur = [s for s in ok if ms(s["date"]) == month]
        prv = [s for s in ok if ms(s["date"]) == prev]
        rev = sum(s["qty"] * s["price"] for s in cur)
        prev_rev = sum(s["qty"] * s["price"] for s in prv)
        plan_total = sum(p["plan"] for p in plan if p["month"] == month)
        same(f"{label} выручка B7", ev("B7"), rev)
        same(f"{label} заказов D7", ev("D7"), len(cur))
        same(f"{label} средний чек H7", ev("H7"), rev / len(cur))
        same(f"{label} % плана J7", ev("J7"), rev / plan_total)
        same(f"{label} план L8", ev("L8"), plan_total)
        if prv:
            same(f"{label} выручка к пред. C8", ev("C8"), rev / prev_rev - 1)
            same(f"{label} заказы к пред. F8", ev("F8"), len(cur) / len(prv) - 1)
            same(f"{label} ср. чек к пред. I8", ev("I8"),
                 (rev / len(cur)) / (prev_rev / len(prv)) - 1)
        else:
            same(f"{label} C8 без пред. месяца", ev("C8"), "")
        # топ-5 товаров
        tops = []
        for i, p in enumerate(products):
            rs = [s for s in cur if s["product"] == p]
            tops.append((-sum(s["qty"] * s["price"] for s in rs), i, p, sum(s["qty"] for s in rs)))
        tops.sort()
        for k, (neg, _i, p, q) in enumerate(tops[:5]):
            row = 12 + k
            same(f"{label} топ-{k + 1} товар", ev(f"C{row}"), p)
            same(f"{label} топ-{k + 1} кол-во", ev(f"D{row}"), q)
            same(f"{label} топ-{k + 1} выручка", ev(f"E{row}"), -neg)
            same(f"{label} топ-{k + 1} доля", ev(f"F{row}"), -neg / rev)
        # менеджеры
        mans = []
        for i, m in enumerate(managers):
            f = sum(s["qty"] * s["price"] for s in cur if s["manager"] == m)
            pl = sum(p["plan"] for p in plan if p["month"] == month and p["manager"] == m)
            mans.append((-f, i, m, pl))
        mans.sort()
        for k, (negf, _i, m, pl) in enumerate(mans):
            row = 12 + k
            same(f"{label} менеджер #{k + 1}", ev(f"H{row}"), m)
            same(f"{label} {m} план", ev(f"I{row}"), pl)
            same(f"{label} {m} факт", ev(f"J{row}"), -negf)
            same(f"{label} {m} % плана", ev(f"K{row}"), -negf / pl)
            same(f"{label} {m} отклонение", ev(f"L{row}"), -negf - pl)
        same(f"{label} итого план I18", ev("I18"), plan_total)
        same(f"{label} итого факт J18", ev("J18"), rev)

    # недели (от выбранного месяца не зависят)
    row = 22
    while True:
        w = ev(f"H{row}")
        if not isinstance(w, (int, float)) or w == 0:
            break
        wd = dt.date(1899, 12, 30) + dt.timedelta(days=int(w))
        rs = [s for s in ok if wd <= s["date"] < wd + dt.timedelta(days=7)]
        same(f"неделя {wd:%d.%m} понедельник", wd.weekday(), 0)
        same(f"неделя {wd:%d.%m} выручка", ev(f"I{row}"), sum(s["qty"] * s["price"] for s in rs))
        same(f"неделя {wd:%d.%m} заказов", ev(f"J{row}"), len(rs))
        row += 1
    weeks_total = sum(ev(f"I{r}") for r in range(22, row))
    same("сумма по неделям = вся выручка", weeks_total, sum(s["qty"] * s["price"] for s in ok))
    # итоги по месяцам
    mr = row + 3  # строка с заголовком «Месяц» (после недель: пустая строка, заголовок блока, шапка)
    for i, m in enumerate((8, 9, 10)):
        r = mr + 1 + i
        month = dt.date(2026, m, 1)
        cur = [s for s in ok if ms(s["date"]) == month]
        same(f"итоги {m}: выручка", ev(f"I{r}"), sum(s["qty"] * s["price"] for s in cur))
        same(f"итоги {m}: заказов", ev(f"J{r}"), len(cur))
        same(f"итоги {m}: отмен", ev(f"L{r}"),
             sum(1 for s in sales if s["status"] == "Отменён" and ms(s["date"]) == month))
    same("итоги: всего", ev(f"I{mr + 4}"), sum(s["qty"] * s["price"] for s in ok))
    today = dt.date.today()
    same("просрочено заявок L4 (на сегодня)", ev("L4"),
         sum(1 for r in req if r["srok"] < today and r["status"] != "Выполнено"))


# ============================================================ uchet.xlsx
def check_uchet():
    path = HERE / "uchet.xlsx"
    wb = load_workbook(path)
    ref = {}
    for r in wb["Справочник"].iter_rows(min_row=2, max_col=7, values_only=True):
        if r[0]:
            ref[r[1]] = dict(art=r[0], buy=r[3], sell=r[4], min=r[5], start=r[6])
    moves = []
    for r in wb["Приход и расход"].iter_rows(min_row=2, max_col=5, values_only=True):
        if r[0] is None:
            break
        moves.append(dict(date=as_date(r[0]), op=r[1], product=r[2], qty=r[3]))
    print(f"uchet.xlsx: {len(moves)} движений, {len(ref)} товаров")
    xl = ExcelCompiler(filename=str(path))

    # автоподстановка цены / суммы / себестоимости — каждая строка
    for i, m in enumerate(moves):
        row = 2 + i
        p = ref[m["product"]]
        price = p["buy"] if m["op"] == "Приход" else p["sell"]
        same(f"движение стр.{row} цена", xl.evaluate(f"'Приход и расход'!F{row}"), price)
        same(f"движение стр.{row} сумма", xl.evaluate(f"'Приход и расход'!G{row}"), price * m["qty"])
        same(f"движение стр.{row} себестоимость", xl.evaluate(f"'Приход и расход'!H{row}"),
             p["buy"] * m["qty"] if m["op"] == "Расход" else "")
    same("пустая строка: цена пустая", xl.evaluate(f"'Приход и расход'!F{len(moves) + 5}"), "")

    names = list(ref)
    total_value = 0
    for i, name in enumerate(names):
        row = 2 + i
        p = ref[name]
        inc = sum(m["qty"] for m in moves if m["product"] == name and m["op"] == "Приход")
        out = sum(m["qty"] for m in moves if m["product"] == name and m["op"] == "Расход")
        ost = p["start"] + inc - out
        status = "Нет в наличии" if ost <= 0 else ("Заказать" if ost < p["min"] else "Норма")
        ev = lambda c: xl.evaluate(f"Остатки!{c}{row}")
        same(f"{name}: товар", ev("B"), name)
        same(f"{name}: приход", ev("E"), inc)
        same(f"{name}: расход", ev("F"), out)
        same(f"{name}: остаток", ev("G"), ost)
        same(f"{name}: статус", ev("I"), status)
        same(f"{name}: к заказу", ev("J"), p["min"] * 2 - max(ost, 0) if ost < p["min"] else 0)
        same(f"{name}: стоимость", ev("K"), max(ost, 0) * p["buy"])
        total_value += max(ost, 0) * p["buy"]
    tot = next(r for r in range(2, 200) if wb["Остатки"][f"B{r}"].value == "Итого")
    for row in range(2 + len(names), tot):  # заготовленные пустые строки под новые товары
        same(f"Остатки стр.{row} пустая", xl.evaluate(f"Остатки!G{row}"), "")
    same("остатки: итого стоимость", xl.evaluate(f"Остатки!K{tot}"), total_value)
    same("остатки: позиций к заказу", xl.evaluate(f"Остатки!I{tot}"),
         f"{sum(1 for n_ in names if ref[n_]['start'] + sum(m['qty'] for m in moves if m['product'] == n_ and m['op'] == 'Приход') - sum(m['qty'] for m in moves if m['product'] == n_ and m['op'] == 'Расход') < ref[n_]['min'])} поз. к заказу")

    tot_sold = 0
    for i, mth in enumerate((8, 9, 10)):
        row = 4 + i
        mm = [m for m in moves if m["date"].month == mth]
        buy = sum(m["qty"] * ref[m["product"]]["buy"] for m in mm if m["op"] == "Приход")
        sold = sum(m["qty"] * ref[m["product"]]["sell"] for m in mm if m["op"] == "Расход")
        cost = sum(m["qty"] * ref[m["product"]]["buy"] for m in mm if m["op"] == "Расход")
        ev = lambda c: xl.evaluate(f"Итоги!{c}{row}")
        same(f"итоги {mth}: закуплено", ev("C"), buy)
        same(f"итоги {mth}: продано", ev("D"), sold)
        same(f"итоги {mth}: себестоимость", ev("E"), cost)
        same(f"итоги {mth}: прибыль", ev("F"), sold - cost)
        same(f"итоги {mth}: маржа", ev("G"), (sold - cost) / sold)
        same(f"итоги {mth}: продаж", ev("H"), sum(1 for m in mm if m["op"] == "Расход"))
        tot_sold += sold
    ws_i = wb["Итоги"]
    tot_row = next(r for r in range(4, 100) if ws_i[f"A{r}"].value == "Итого")
    for row in range(7, tot_row):  # будущие месяцы учётного года — пока нули
        same(f"итоги стр.{row} ({ws_i[f'A{row}'].value}) продано", xl.evaluate(f"Итоги!D{row}"), 0)
    same("итоги: всего продано", xl.evaluate(f"Итоги!D{tot_row}"), tot_sold)
    # матрица товар × месяц, шт.
    hdr = next(r for r in range(tot_row, 200) if ws_i[f"A{r}"].value == "Товар")
    for i, name in enumerate(names):
        row = hdr + 1 + i
        for j, mth in enumerate((8, 9, 10)):
            q = sum(m["qty"] for m in moves if m["product"] == name and m["op"] == "Расход" and m["date"].month == mth)
            same(f"матрица {name} / {mth}", xl.evaluate(f"Итоги!{'CDE'[j]}{row}"), q)
        same(f"матрица {name}: товар", xl.evaluate(f"Итоги!A{row}"), name)
        same(f"матрица {name}: итого", xl.evaluate(f"Итоги!O{row}"),
             sum(m["qty"] for m in moves if m["product"] == name and m["op"] == "Расход"))


if __name__ == "__main__":
    check_dashboard()
    check_uchet()
    print(f"\nПроверено значений: {checked}, расхождений: {errors}")
    sys.exit(1 if errors else 0)
