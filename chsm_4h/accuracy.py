# -*- coding: utf-8 -*-
"""Проверка по факту АТС: как часто ЧСМ попадает в 4-часовой блок разгрузки.

ЧСМ суток — час максимума потребления субъекта ВНУТРИ плановых пиковых
часов СО ЕЭС. Блок — 4 подряд идущих часа, целиком лежащих в окне СО.
Попадание с учётом разгрузки: регион в часах блока снижен на суммарную
глубину заводов региона, и максимум окна всё равно остаётся в блоке.
"""
from __future__ import annotations

import collections

BLOCK = 4


def parse_window(s):
    hours = set()
    for part in s.split(","):
        a, b = part.split("-")
        hours.update(range(int(a), int(b) + 1))
    return hours


def blocks_in(window, n=BLOCK):
    return [tuple(range(h, h + n)) for h in range(1, 25 - n + 1)
            if all(x in window for x in range(h, h + n))]


def chsm(day, window, block=(), cut=0.0):
    """Час максимума окна после снижения блока на cut МВт."""
    return max(sorted(window), key=lambda h: day[h] - (cut if h in block else 0.0))


def month_stats(days, window, cut):
    """Для каждого блока месяца: (попаданий без разгрузки, с разгрузкой)."""
    res = {}
    for b in blocks_in(window):
        hit0 = sum(chsm(d, window) in b for d in days)
        hit1 = sum(chsm(d, window, b, cut) in b for d in days)
        res[b] = (hit0, hit1)
    return res


def analyse(region_days, window_for_month, cut):
    """Помесячно: лучший блок месяца (в выборке) и блок прошлого месяца."""
    by_m = collections.defaultdict(list)
    for d, hh in sorted(region_days.items()):
        by_m[d[:7]].append(hh)
    rows, prev_best = [], None
    for ym in sorted(by_m):
        days = by_m[ym]
        win = window_for_month(ym)
        st = month_stats(days, win, cut)
        best = max(st, key=lambda b: (st[b][1], st[b][0]))
        modes = collections.Counter(chsm(d, win) for d in days)
        wf = None
        if prev_best is not None and prev_best in st:
            wf = st[prev_best][1]
        rows.append(dict(ym=ym, n=len(days), window=win, best=best,
                         hit0=st[best][0], hit1=st[best][1],
                         prev_block=prev_best, hit_prev=wf, modes=modes))
        prev_best = best
    return rows


def oracle(region_days, window_for_month, cut):
    """Идеальный посуточный прогноз: на каждые сутки лучший блок.

    Запас блока = max(блок) − max(окно вне блока). Если запас ≥ глубины —
    вся глубина зачётна; иначе полный срез переносит ЧСМ за блок (эффект 0),
    а безопасный срез — только величина запаса.
    Возвращает помесячно: дней, дней с полной глубиной, средний безопасный срез.
    """
    by_m = collections.defaultdict(list)
    for d, hh in sorted(region_days.items()):
        by_m[d[:7]].append(hh)
    out = {}
    for ym, days in sorted(by_m.items()):
        win = window_for_month(ym)
        full = 0
        safe = 0.0
        for d in days:
            best_gap = 0.0
            for b in blocks_in(win):
                rest = [d[h] for h in win if h not in b]
                gap = max(d[h] for h in b) - (max(rest) if rest else 0.0)
                best_gap = max(best_gap, gap)
            full += best_gap >= cut
            safe += min(cut, best_gap)
        out[ym] = dict(n=len(days), full=full, safe=safe / len(days))
    return out
