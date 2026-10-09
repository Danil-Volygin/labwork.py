# -*- coding: utf-8 -*-
"""Чтение посуточного факта АТС (raw/YYYYMM01_<код>_fact_region.xls[x]).

Отчёт АТС содержит только рабочие дни; час — метка 1..24 МСК,
та же, что в плановых часах СО ЕЭС.
"""
from __future__ import annotations

import collections
import glob
import os
import re

import openpyxl
import xlrd

#: Код субъекта в имени файла АТС -> регион.
REGION_CODE = {
    "04": "Красноярский край",
    "18": "Волгоградская область",
    "25": "Иркутская область",
    "32": "Кемеровская область",
    "86": "Республика Карелия",
    "95": "Республика Хакасия",
}


def _rows(path):
    if path.endswith(".xls"):
        sh = xlrd.open_workbook(path).sheet_by_index(0)
        for i in range(sh.nrows):
            yield sh.row_values(i)
    else:
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        yield from ws.iter_rows(values_only=True)


def load_raw(raw_dir):
    """{регион: {'YYYY-MM-DD': {час: МВт}}}."""
    out = collections.defaultdict(lambda: collections.defaultdict(dict))
    for path in sorted(glob.glob(os.path.join(raw_dir, "*_fact_region.xls*"))):
        code = os.path.basename(path).split("_")[1]
        region = REGION_CODE.get(code)
        if not region:
            continue
        for row in _rows(path):
            if not row or not isinstance(row[0], str):
                continue
            m = re.fullmatch(r"(\d\d)\.(\d\d)\.(\d{4})", row[0].strip())
            if not m or row[1] in (None, "") or row[2] in (None, ""):
                continue
            d = f"{m[3]}-{m[2]}-{m[1]}"
            out[region][d][int(float(row[1]))] = float(row[2])
    return out
