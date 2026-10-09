# -*- coding: utf-8 -*-
"""Лестница срезов из паспортов (build_razgruzka_zavody.py) и посуточный вариант.

Лестница: средний профиль рабочих дней месяца (МСК), часы окна СО ЕЭС,
срезаются N верхних часов сверху вниз, каждый не ниже «нижний сосед + зазор»,
пол — (N+1)-й час. Срез региона ограничен суммой потолков активных заводов
и делится между ними пропорционально потолкам. В деньги — срез 1-го пика.
"""
from __future__ import annotations

import collections
import statistics


def ladder_cut(vals, n, gap, cap):
    """Срез 1-го пика по лестнице на N часов (vals — МВт часов окна)."""
    p = sorted(vals, reverse=True)
    support = p[n] if len(p) > n else min(p)
    for k in range(n - 1, 0, -1):        # пики N..2
        cut = max(0.0, min(p[k] - (support + gap), cap))
        support = p[k] - cut
    return max(0.0, min(p[0] - (support + gap), cap))


def monthly_profile(days):
    by_m = collections.defaultdict(list)
    for d, hh in days.items():
        by_m[d[:7]].append(hh)
    return {ym: {h: statistics.mean(x[h] for x in ds) for h in range(1, 25)}
            for ym, ds in by_m.items()}, {ym: len(ds) for ym, ds in by_m.items()}


def region_ladder(days, window_for_month, n, gap, cap):
    """{месяц: срез 1-го пика региона по среднему профилю}."""
    prof, _ = monthly_profile(days)
    return {ym: ladder_cut([pr[h] for h in window_for_month(ym)], n, gap, cap)
            for ym, pr in sorted(prof.items())}


def region_daily(days, window_for_month, n, gap, cap):
    """{месяц: средний по рабочим дням срез 1-го пика при посуточном прогнозе}.

    Та же лестница, но на профиле каждых суток: модуль прогноза знает форму
    графика дня и режет глубже там, где пик выражен.
    """
    acc = collections.defaultdict(list)
    for d, hh in sorted(days.items()):
        win = window_for_month(d[:7])
        acc[d[:7]].append(ladder_cut([hh[h] for h in win], n, gap, cap))
    return {ym: statistics.mean(v) for ym, v in acc.items()}
