# -*- coding: utf-8 -*-
"""Эффект модуля прогнозирования нагрузки сверх БП (слайд «Экономический эффект»).

Сравниваются три способа посчитать срез 1-го пика (ЧСМ) по одной методике
паспортов — лестница N верхних часов окна СО ЕЭС, зазор 5 МВт, срез региона
делится пропорционально потолкам, зачётное снижение = срез × 80 %:

* «2 ч, среднемес.» — как в паспортах тиражирования (2 часа, средний профиль);
* «4 ч, среднемес.» — как в паспорте единого диспетчера (4 часа, средний профиль);
* «4 ч, посуточно» — модуль прогноза: та же лестница на 4 часа, но на прогнозе
  профиля каждых суток, а не на среднемесячном профиле.

Период — 12 месяцев факта АТС (июнь 2025 — май 2026), все заводы активны
весь год, цена мощности 1 млн руб/МВт·мес (как в БП).

Запуск:
    python forecast_effect.py <папка raw с отчётами АТС> [out.json]
"""
from __future__ import annotations

import json
import statistics
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import accuracy as A  # noqa: E402
import ladder as L  # noqa: E402
import load_ats  # noqa: E402
from build_payback_4h import SO_WINDOWS  # noqa: E402

ACCURACY = 0.80
GAP = 5.0
PRICE = 1_000_000  # руб/МВт·мес

#: Эффект по БП со слайда 3, млн руб/год (ВгАЗ и КАЗ исключены).
BP = {"БрАЗ": 198.1, "ИркАЗ": 59.9, "НкАЗ": 147.3, "САЗ": 36.9, "ТАЗ": 33.8}

#: Регион -> (ценовая зона, {завод: потолок снижения, МВт}).
REGIONS = {
    "Иркутская область": (2, {"БрАЗ": 49.0, "ИркАЗ": 10.5, "ТАЗ": 6.0}),
    "Кемеровская область": (2, {"НкАЗ": 17.0}),
    "Республика Хакасия": (2, {"САЗ": 15.0}),
}

VARIANTS = {
    "2ч_среднемес": (L.region_ladder, 2),
    "4ч_среднемес": (L.region_ladder, 4),
    "4ч_посуточно": (L.region_daily, 4),
}


def compute(raw_dir):
    data = load_ats.load_raw(str(raw_dir))
    out = {}
    for region, (zone, plants) in REGIONS.items():
        q = sum(plants.values())
        wf = lambda ym, z=zone: A.parse_window(SO_WINDOWS[z][int(ym[5:7])])  # noqa: E731
        monthly = {key: fn(data[region], wf, n, GAP, q) for key, (fn, n) in VARIANTS.items()}
        for plant, cap in plants.items():
            row = {"cap": cap, "region": region, "bp": BP[plant], "months": {}}
            for key, by_m in monthly.items():
                cut = statistics.mean(by_m.values()) * cap / q          # МВт, срез завода
                row[key] = cut * ACCURACY * PRICE * 12 / 1e6             # млн руб/год
                row["months"][key] = {ym: v * cap / q for ym, v in by_m.items()}
            out[plant] = row
    return out


def main(argv):
    res = compute(argv[1])
    print(f"{'':6}{'БП':>7}" + "".join(f"{k:>14}" for k in VARIANTS) + f"{'сверх БП':>10}")
    for p, r in res.items():
        print(f"{p:6}{r['bp']:7.1f}" + "".join(f"{r[k]:14.1f}" for k in VARIANTS)
              + f"{r['4ч_посуточно'] - r['bp']:+10.1f}")
    if len(argv) > 2:
        Path(argv[2]).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv)
