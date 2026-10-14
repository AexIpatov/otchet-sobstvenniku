# src/presentation_builder_v2.py
# ============================================================
# Генератор Excel с диаграммами для презентации.
#
# ИЗМЕНЕНИЯ (сентябрь 2026):
#   1) _collect_breakeven теперь разделяет расходы на переменные
#      и постоянные по config.VARIABLE_EXPENSE_KEYS / FIXED_EXPENSE_KEYS.
#   2) Точка безубыточности считается по формуле:
#        ТБУ = Постоянные / (1 - Переменные/Выручка)
#      Если переменных нет (чистая аренда) — ТБУ = Постоянные.
#   3) Добавлена поддержка косвенных расходов, распределённых
#      на юнит по валовой прибыли (ВП1).
# ============================================================

import os
import re
import datetime as _dt

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.spreadsheet_drawing import (
    OneCellAnchor, AnchorMarker,
)
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import cm_to_EMU

import config


# ============================================================
# КОНСТАНТЫ
# ============================================================

_MONTHS_RU_LOWER = {
    1: "январь", 2: "февраль", 3: "март", 4: "апрель",
    5: "май", 6: "июнь", 7: "июль", 8: "август",
    9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь",
}
_MONTHS_RU_CAP = {
    1: "Янв", 2: "Фев", 3: "Мар", 4: "Апр", 5: "Май", 6: "Июн",
    7: "Июл", 8: "Авг", 9: "Сен", 10: "Окт", 11: "Ноя", 12: "Дек",
}

_UNITS = [
    "Latvia",
    "East-Восток",
    "Europe",
    "Nomiqa",
    "Unelma",
    "UK Estate",
]

_LATVIA_OBJECTS = [
    ("AN14 Антониас 14 (дом + парковка)", "Антонияс", 10),
    ("AC89 Чака 89 (дом + парковка)",     "Чака",     15),
    ("M81 - Matisa 81",                    "Матиса",   32),
    ("EKS_Esporta iela 12-113",            "Эспорта",  28),
]

_LATVIA_COMMERCIAL = [
    "D4 Парковка-Deglava4",
    "AC87 Гараж Чака",
    "B117 Бривибас, 117",
    "B78 Бривибас, 78",
    "C23 Цесу, 23",
    "DAR1_Darzauglu1",
    "DS1 Дзирнаву, 1",
    "G73 Гертрудес, 73",
    "H5 Хоспиталю",
    "MP1_Marupe",
    "MU3 - Mucenieku 3 - 4",
    "OZ1 Озолниеки",
    "SK3-Skunju 3",
    "UK_Latvia",
    "V22 К. Валдемара 22",
    "BRN_Brunieku",
]

_UNIT_SLIDE_START = {
    "Latvia":        4,
    "East-Восток":  38,
    "Europe":       49,
    "Nomiqa":       71,
    "Unelma":       67,
    "UK Estate":    56,
}

_HEADER_FILL   = PatternFill("solid", fgColor="1E3A8A")
_HEADER_FONT   = Font(bold=True, color="FFFFFF", size=11)
_TITLE_FONT    = Font(bold=True, size=14, color="1E3A8A")
_SUBTITLE_FONT = Font(italic=True, size=10, color="64748B")
_CELL_BORDER   = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1"),
)

_COLOR_REVENUE = "FFC000"
_COLOR_NET     = "70AD47"
_COLOR_LINE_RED = "C00000"
_COLOR_LINE_YEL = "FFC000"
_COLOR_LINE_BLU = "4472C4"


# ============================================================
# ВСПОМОГАТЕЛЬНОЕ
# ============================================================

def _clean(s):
    if s is None:
        return ""
    s = str(s).replace("\u00a0", " ")
    s = re.sub(r"[\s·]+$", "", s)
    s = re.sub(r"^[\s·]+", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _level(s):
    if s is None:
        return 0
    m = re.match(r"^[\s·]*", str(s))
    return m.group(0).count("·") if m else 0


def _num(ws, r, c):
    v = ws.cell(row=r, column=c).value
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
    if s in ("", "-"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _safe_sheet(name):
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31]


def _find_month_columns(ws, header_row, months):
    result = {}
    for c in range(2, min(ws.max_column, 25) + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        if isinstance(v, (_dt.datetime, _dt.date)):
            if v.month in months and v.month not in result:
                result[v.month] = c
            continue
        sv = str(v).lower()
        for m in months:
            if m in result:
                continue
            if _MONTHS_RU_LOWER[m] in sv:
                result[m] = c
                break
    return result


# ============================================================
# ШАГ 1. ПЛОСКИЙ СПИСОК СТРОК ОПиУ
# ============================================================

def _read_flat_rows(opiu_path, months):
    if not opiu_path or not os.path.exists(opiu_path):
        return [], {}

    wb = openpyxl.load_workbook(opiu_path, data_only=True)
    ws = wb.active

    header_row = None
    for r in range(1, 6):
        hits = 0
        for c in range(2, min(ws.max_column, 25) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            if isinstance(v, (_dt.datetime, _dt.date)):
                hits += 1
                continue
            sv = str(v).lower()
            if any(m in sv for m in _MONTHS_RU_LOWER.values()):
                hits += 1
        if hits >= 3:
            header_row = r
            break

    if header_row is None:
        print("[presentation_builder] шапка с месяцами не найдена")
        return [], {}

    month_col = _find_month_columns(ws, header_row, months)

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue
        title = _clean(raw)
        if title == "":
            continue
        lvl = _level(raw)
        values = {m: _num(ws, r, col) for m, col in month_col.items()}
        rows.append((r, lvl, title, values))

    return rows, month_col


# ============================================================
# ШАГ 2. ИЗВЛЕЧЕНИЕ ДАННЫХ
# ============================================================

_SECTION_REVENUE = "Выручка"
_SECTION_NET     = "Чистая прибыль"
_SECTION_PROD    = "Производственные расходы"
_SECTION_COMM    = "Коммерческие расходы"


def _find_row(rows, section, unit, object_name=None):
    cur_section = None
    cur_unit = None
    cur_object = None

    for r, lvl, title, values in rows:
        if lvl == 0:
            cur_section = title
            cur_unit = None
            cur_object = None
            continue

        if lvl == 1:
            cur_unit = title
            cur_object = None
            if (object_name is None
                    and cur_section == section
                    and cur_unit == unit):
                return dict(values)
            continue

        if lvl == 2:
            cur_object = title
            if (object_name is not None
                    and cur_section == section
                    and cur_unit == unit
                    and cur_object == object_name):
                return dict(values)
            continue

    return {}


def _is_inside_section(rows, idx, section):
    for i in range(idx - 1, -1, -1):
        r, lvl, title, values = rows[i]
        if lvl == 0:
            return title == section
    return False


def _find_object_fot(rows, section, unit, object_name):
    result = {}

    unit_idx = None
    for i, (r, lvl, title, values) in enumerate(rows):
        if title == unit:
            if _is_inside_section(rows, i, section):
                unit_idx = i
                break
    if unit_idx is None:
        return result

    obj_idx = None
    for i in range(unit_idx + 1, len(rows)):
        r, lvl, title, values = rows[i]
        if lvl == 0 and title != section:
            break
        if title == object_name:
            obj_idx = i
            break
    if obj_idx is None:
        return result

    obj_lvl = rows[obj_idx][1]

    for i in range(obj_idx + 1, len(rows)):
        r, lvl, title, values = rows[i]
        if lvl <= obj_lvl:
            break
        if lvl == 0 and title != section:
            break

        tl = title.lower()
        if "фот производственного" in tl or "фот коммерческого" in tl:
            for m, v in values.items():
                result[m] = result.get(m, 0.0) + v

    return result


def _collect_unit_expenses(rows, unit_name, months):
    sections = (
        "Производственные расходы",
        "Косвенные расходы",
        "Коммерческие расходы",
    )
    result = {m: 0.0 for m in months}

    cur_section = None
    cur_unit_level = None
    inside_target = False

    for r, lvl, title, values in rows:
        if lvl == 0:
            cur_section = title
            cur_unit_level = None
            inside_target = False
            continue

        if cur_section not in sections:
            continue

        if title == unit_name:
            cur_unit_level = lvl
            inside_target = True
            continue

        if not inside_target:
            continue

        if lvl <= cur_unit_level:
            inside_target = False
            continue

        if lvl == cur_unit_level + 1:
            for m in months:
                v = values.get(m, 0.0) or 0.0
                result[m] += abs(v)

    return result


def _parse_investments_all(ref_path):
    result = {}
    if not ref_path or not os.path.exists(ref_path):
        return result

    wb = openpyxl.load_workbook(ref_path, data_only=True)
    ws = wb.active

    for r in range(1, ws.max_row + 1):
        name = _clean(ws.cell(row=r, column=1).value)
        if not name:
            continue

        if "наименование" in name.lower():
            continue

        v = ws.cell(row=r, column=2).value

        if isinstance(v, str) and v.startswith("="):
            continue

        amount = _num(ws, r, 2)

        if amount == 0.0 and v in (None, "", 0):
            if v == 0:
                result[name] = 0.0
            continue

        result[name] = float(amount)

    return result


def _sum_investments(investments_all, object_names):
    total = 0.0
    for name in object_names:
        total += investments_all.get(name, 0.0)
    return total


def _parse_debts(debts_path, months):
    result = {
        "debts": {},
        "headcount": {m: 0.0 for m in months},
    }
    if not debts_path or not os.path.exists(debts_path):
        print(f"[_parse_debts] файл не найден: {debts_path}")
        return result

    wb = openpyxl.load_workbook(debts_path, data_only=True)
    ws = wb.active

    print(f"[_parse_debts] лист: {ws.title}, "
          f"строк: {ws.max_row}, колонок: {ws.max_column}")

    header_row = None
    for r in range(1, min(30, ws.max_row + 1)):
        v = ws.cell(row=r, column=1).value
        if not v:
            continue
        sv = str(v).strip().lower()
        if sv != "объект":
            continue
        has_dates = any(
            isinstance(ws.cell(row=r, column=c).value,
                       (_dt.datetime, _dt.date))
            for c in range(2, min(ws.max_column, 25) + 1)
        )
        if has_dates:
            header_row = r
            break

    if header_row is None:
        print("[_parse_debts] не найдена шапка «Объект» с датами")
        for rr in range(1, min(6, ws.max_row + 1)):
            print(f"[_parse_debts]   A{rr} = "
                  f"{ws.cell(row=rr, column=1).value!r}")
        return result
    print(f"[_parse_debts] шапка «Объект» с датами — строка {header_row}")

    month_col = {}
    for c in range(2, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, (_dt.datetime, _dt.date)):
            if v.year == config.REPORT_YEAR and v.month in months:
                month_col[v.month] = c
            continue
        sv = str(v).strip().lower()
        for m in months:
            if m in month_col:
                continue
            month_name = _MONTHS_RU_LOWER[m]
            if month_name in sv and str(config.REPORT_YEAR) in sv:
                month_col[m] = c
                break
    print(f"[_parse_debts] колонок с датами в шапке (год "
          f"{config.REPORT_YEAR}): {len(month_col)}")
    print(f"[_parse_debts]   month_col = {month_col}")

    if not month_col:
        print("[_parse_debts] в шапке нет дат нужных месяцев")
        return result

    headcount_header_row = None
    for r in range(header_row + 1, ws.max_row + 1):
        name = _clean(ws.cell(row=r, column=1).value)
        if not name:
            continue
        if "численность" in name.lower():
            headcount_header_row = r
            break

        vals = {}
        for m, col in month_col.items():
            vals[m] = _num(ws, r, col)
        result["debts"][name] = vals

    print(f"[_parse_debts] объектов с долгами: {len(result['debts'])}")
    for k in result["debts"]:
        print(f"[_parse_debts]   • {k}")

    if headcount_header_row is not None:
        print(f"[_parse_debts] блок «Численность» — строка {headcount_header_row}")

        hc_dates_row = None
        hc_values_row = None

        for rr in range(headcount_header_row + 1,
                        min(headcount_header_row + 10, ws.max_row + 1)):
            a_val = ws.cell(row=rr, column=1).value
            has_dates = any(
                isinstance(ws.cell(row=rr, column=c).value, (_dt.datetime, _dt.date))
                for c in range(2, min(ws.max_column, 25) + 1)
            )
            has_numbers = any(
                isinstance(ws.cell(row=rr, column=c).value, (int, float))
                and not isinstance(ws.cell(row=rr, column=c).value, bool)
                for c in range(2, min(ws.max_column, 25) + 1)
            )

            if has_dates and hc_dates_row is None:
                hc_dates_row = rr
            elif has_numbers and hc_dates_row is not None and hc_values_row is None:
                hc_values_row = rr

        print(f"[_parse_debts]   строка с датами: {hc_dates_row}")
        print(f"[_parse_debts]   строка с числами: {hc_values_row}")

        if hc_values_row is not None:
            hc_month_col = {}
            if hc_dates_row is not None:
                for c in range(2, ws.max_column + 1):
                    v = ws.cell(row=hc_dates_row, column=c).value
                    if isinstance(v, (_dt.datetime, _dt.date)):
                        if v.year == config.REPORT_YEAR and v.month in months:
                            hc_month_col[v.month] = c
                        continue
                    sv = str(v).strip().lower()
                    for m in months:
                        if m in hc_month_col:
                            continue
                        month_name = _MONTHS_RU_LOWER[m]
                        if month_name in sv and str(config.REPORT_YEAR) in sv:
                            hc_month_col[m] = c
                            break
            if not hc_month_col:
                hc_month_col = month_col

            for m, col in hc_month_col.items():
                result["headcount"][m] = _num(ws, hc_values_row, col)

            print(f"[_parse_debts] численность: {result['headcount']}")
        else:
            print("[_parse_debts] не найдена строка с числами в блоке «Численность»")
    else:
        print("[_parse_debts] блок «Численность» не найден")

    return result


def _sheet_debts_bars(wb, sheet_title, chart_title, subtitle,
                      objects, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц"] + [o["name"] for o in objects]
    rows_data = []
    for m in months:
        row = [_MONTHS_RU_CAP[m]]
        for o in objects:
            row.append(round(o["values"].get(m, 0.0) or 0.0, 2))
        rows_data.append(row)
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=1 + len(objects),
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.overlap = -10
    chart.gapWidth = 60
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    for idx, o in enumerate(objects):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.solidFill = o["color"]
            chart.series[idx].graphicalProperties.line.solidFill = o["color"]

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "outEnd"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")


def _sheet_debts_lines(wb, sheet_title, chart_title, subtitle,
                       objects, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц"] + [o["name"] for o in objects]
    rows_data = []
    for m in months:
        row = [_MONTHS_RU_CAP[m]]
        for o in objects:
            row.append(round(o["values"].get(m, 0.0) or 0.0, 2))
        rows_data.append(row)
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=1 + len(objects),
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    for idx, o in enumerate(objects):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.line.solidFill = o["color"]
            chart.series[idx].graphicalProperties.line.width = 25000
            chart.series[idx].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "t"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")


def _sheet_headcount_analysis(wb, sheet_title, chart_title, subtitle,
                              headcount, total_revenue, total_fot, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц", "Численность",
               "Выручка на 1 сотрудника",
               "Средняя ЗП на 1 сотрудника"]
    rows_data = []
    for m in months:
        hc = headcount.get(m, 0.0) or 0.0
        rev = total_revenue.get(m, 0.0) or 0.0
        fot = total_fot.get(m, 0.0) or 0.0
        rev_per = (rev / hc) if hc else 0.0
        fot_per = (abs(fot) / hc) if hc else 0.0
        rows_data.append([
            _MONTHS_RU_CAP[m],
            round(hc, 2),
            round(rev_per, 2),
            round(fot_per, 2),
        ])
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=4,
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    colors = ["4472C4", "C00000", "70AD47"]
    for idx, c in enumerate(colors):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.line.solidFill = c
            chart.series[idx].graphicalProperties.line.width = 25000
            chart.series[idx].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "t"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")


def _sheet_roi(wb, sheet_title, chart_title, subtitle,
               series_data, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц"] + [s["name"] for s in series_data]
    rows_data = []
    for m in months:
        row = [_MONTHS_RU_CAP[m]]
        for s in series_data:
            row.append(round(s["values"].get(m, 0.0) or 0.0, 4))
        rows_data.append(row)
    _write_month_table(ws, header_row, headers, rows_data, number_fmt='0.00%')

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=1 + len(series_data),
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None
    chart.y_axis.numFmt = '0.00%'

    for idx, s in enumerate(series_data):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.line.solidFill = s["color"]
            chart.series[idx].graphicalProperties.line.width = 25000
            chart.series[idx].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '0.00%'
    chart.dLbls.position = "t"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")


# ============================================================
# НОВЫЙ РАСЧЁТ ТБУ С РАЗДЕЛЕНИЕМ РАСХОДОВ
# ============================================================

def _collect_breakeven(rows, months, fixed_override=None):
    """
    Возвращает {unit_name: {month: value}} — точку безубыточности (ТБУ).

    Методика «Нескучных финансов»:
      ТБУ = Постоянные / (1 - Переменные / Выручка)
      Если переменных нет — ТБУ = Постоянные.

    Структура ОПиУ:
      Производственные расходы (level 0)
        · Прямые производственные (level 1)
          ·· Latvia (level 2)          ← юнит
            ··· AN14 (level 3)         ← объект — АГРЕГАТ
              ···· <статьи> (level 4)
      Косвенные расходы (level 0)
        · Административные расходы (level 1)
          ·· Latvia (level 2)          ← юнит
            ··· SK3 (level 3)
      Коммерческие расходы (level 0)
        ·· Nomiqa (level 2)            ← юнит
          ··· BNQ_BAKU (level 3)
    """
    _SECTIONS = (
        "Производственные расходы",
        "Косвенные расходы",
        "Коммерческие расходы",
    )

    # Множества нормализованных названий статей
    var_keys_norm = {
        _norm_key(k) for k in config.VARIABLE_EXPENSE_KEYS
    }
    fix_keys_norm = {
        _norm_key(k) for k in config.FIXED_EXPENSE_KEYS
    }

    revenue = {u: {m: 0.0 for m in months} for u in _UNITS}
    variable = {u: {m: 0.0 for m in months} for u in _UNITS}
    fixed = {u: {m: 0.0 for m in months} for u in _UNITS}

    cur_section = None
    cur_unit = None
    cur_unit_level = None

    # Стек уровней, чтобы понимать, куда относить статьи.
    for r, lvl, title, values in rows:
        # level 0 — смена секции
        if lvl == 0:
            cur_section = title
            cur_unit = None
            cur_unit_level = None
            continue

        # level 1 — юнит или не-юнит
        if lvl == 1:
            if title in _UNITS and cur_section in _SECTIONS:
                cur_unit = title
                cur_unit_level = 1
            else:
                cur_unit = None
                cur_unit_level = None
            continue

        # юнит ниже level 1 (например, Latvia внутри
        # «Прямые производственные» — на level 2)
        if title in _UNITS:
            if cur_section in _SECTIONS or cur_section == "Выручка":
                cur_unit = title
                cur_unit_level = lvl
            continue

        if cur_unit is None:
            continue

        # вышли за пределы юнита
        if lvl <= cur_unit_level:
            cur_unit = None
            cur_unit_level = None
            continue

        u = cur_unit

        # выручка юнита
        if cur_section == "Выручка":
            if lvl == cur_unit_level + 1:
                for m in months:
                    revenue[u][m] += values.get(m, 0.0) or 0.0
            continue

        # расходы: только прямые дети юнита
        if cur_section in _SECTIONS:
            if lvl == cur_unit_level + 1:
                title_norm = _norm_key(title)
                for m in months:
                    v = abs(values.get(m, 0.0) or 0.0)

                    # Классификация статьи:
                    #   - переменная (в var_keys_norm)
                    #   - постоянная (в fix_keys_norm)
                    #   - по умолчанию — постоянная
                    if title_norm in var_keys_norm:
                        variable[u][m] += v
                    elif title_norm in fix_keys_norm:
                        fixed[u][m] += v
                    else:
                        fixed[u][m] += v

                    if config.DEBUG and m == 1:
                        print(f"[ТБУ-СТАТЬЯ] юнит={u}, "
                              f"секция={cur_section}, lvl={lvl}, "
                              f"title={title!r}, "
                              f"norm={title_norm!r}, "
                              f"var={title_norm in var_keys_norm}, "
                              f"fix={title_norm in fix_keys_norm}, "
                              f"value={v:.2f}")
            continue

    # --- финальный расчёт ТБУ ---
    result = {u: {m: 0.0 for m in months} for u in _UNITS}
    for u in _UNITS:
        for m in months:
            rev = revenue[u].get(m, 0.0) or 0.0
            var = variable[u].get(m, 0.0) or 0.0
            fix = fixed[u].get(m, 0.0) or 0.0

            if rev <= 0:
                # Нет выручки — ТБУ не определена
                result[u][m] = 0.0
                continue

            # КМД = (Выручка − Переменные) / Выручка
            kmd = (rev - var) / rev

            if kmd <= 0:
                # Переменные съедают всю выручку — ТБУ не определена
                result[u][m] = 0.0
            else:
                result[u][m] = fix / kmd

    if config.DEBUG:
        for u in _UNITS:
            for m in months:
                print(f"[ТБУ] {u} м.{m}: "
                      f"выручка={revenue[u][m]:.2f}, "
                      f"перем={variable[u][m]:.2f}, "
                      f"пост={fixed[u][m]:.2f}, "
                      f"ТБУ={result[u][m]:.2f}")

    return result


def _norm_key(s):
    """
    Нормализация названия статьи для сопоставления с ключами
    config.VARIABLE_EXPENSE_KEYS / config.FIXED_EXPENSE_KEYS.
    Убирает пробелы, дефисы, точки, приводит к нижнему регистру.
    """
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[\s\-–—.,]+", "", s)
    return s


# ============================================================
# ШАГ 3. СБОРКА ДАННЫХ
# ============================================================

def _collect_units(rows, months):
    units = {u: {
        "revenue": {m: 0.0 for m in months},
        "net":     {m: 0.0 for m in months},
        "fot":     {m: 0.0 for m in months},
        "objects": {},
    } for u in _UNITS}

    cur_section = None
    cur_unit = None
    for r, lvl, title, values in rows:
        if lvl == 0:
            cur_section = title
            cur_unit = None
            continue
        if lvl == 1:
            cur_unit = title
            continue
        if (lvl == 2
                and cur_section in (_SECTION_REVENUE, _SECTION_NET)
                and cur_unit in _UNITS):
            if title not in units[cur_unit]["objects"]:
                units[cur_unit]["objects"][title] = {
                    "revenue": {m: 0.0 for m in months},
                    "net":     {m: 0.0 for m in months},
                    "fot":     {m: 0.0 for m in months},
                }

    for u_name, u in units.items():
        for obj_name in list(u["objects"].keys()):
            rev = _find_row(rows, _SECTION_REVENUE, u_name, object_name=obj_name)
            net = _find_row(rows, _SECTION_NET, u_name, object_name=obj_name)
            for m in months:
                u["objects"][obj_name]["revenue"][m] = rev.get(m, 0.0)
                u["objects"][obj_name]["net"][m]     = net.get(m, 0.0)

    for u_name, u in units.items():
        for obj_name in list(u["objects"].keys()):
            fot_prod = _find_object_fot(rows, _SECTION_PROD, u_name, obj_name)
            fot_comm = _find_object_fot(rows, _SECTION_COMM, u_name, obj_name)
            for m in months:
                u["objects"][obj_name]["fot"][m] = (
                    fot_prod.get(m, 0.0) + fot_comm.get(m, 0.0)
                )

    for u_name, u in units.items():
        for m in months:
            u["net"][m]     = sum(o["net"][m]     for o in u["objects"].values())
            u["revenue"][m] = sum(o["revenue"][m] for o in u["objects"].values())
            u["fot"][m]     = sum(o["fot"][m]     for o in u["objects"].values())

    if "Latvia" in units:
        lat = units["Latvia"]
        commercial = {
            "revenue": {m: 0.0 for m in months},
            "net":     {m: 0.0 for m in months},
            "fot":     {m: 0.0 for m in months},
        }
        for obj_name in _LATVIA_COMMERCIAL:
            obj = lat["objects"].get(obj_name)
            if not obj:
                continue
            for m in months:
                commercial["revenue"][m] += obj["revenue"][m]
                commercial["net"][m]     += obj["net"][m]
                commercial["fot"][m]     += obj["fot"][m]
        lat["objects"]["Коммерческие помещения LV"] = commercial

    return units


# ============================================================
# ПОСТРОЕНИЕ ЛИСТОВ
# ============================================================

def _write_header(ws, title, subtitle=""):
    ws.cell(row=1, column=1, value=title).font = _TITLE_FONT
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = _SUBTITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=13)


def _write_month_table(ws, start_row, headers, rows_data, number_fmt='#,##0'):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _CELL_BORDER
    for r_off, row_vals in enumerate(rows_data, start=1):
        for c, v in enumerate(row_vals, start=1):
            cell = ws.cell(row=start_row + r_off, column=c, value=v)
            cell.border = _CELL_BORDER
            if isinstance(v, (int, float)) and c > 1:
                cell.number_format = number_fmt


def _sheet_revenue_profit(wb, sheet_title, chart_title, subtitle,
                          color_rev, color_net,
                          data, months):
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц", "Выручка", "Чистая прибыль"]
    rows_data = []
    for m in months:
        rows_data.append([
            _MONTHS_RU_CAP[m],
            round(data["revenue"][m], 2),
            round(data["net"][m], 2),
        ])
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=3,
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.overlap = -10
    chart.gapWidth = 60
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)

    chart.title = chart_title
    chart.width = 20
    chart.height = 10
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    chart.series[0].graphicalProperties.solidFill = color_rev
    chart.series[0].graphicalProperties.line.solidFill = color_rev
    chart.series[1].graphicalProperties.solidFill = color_net
    chart.series[1].graphicalProperties.line.solidFill = color_net

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "outEnd"

    chart.legend.position = "b"
    chart.legend.overlay = False

    ws.add_chart(chart, "E4")


def _sheet_ebitda(wb, sheet_title, chart_title, subtitle,
                  line_color, data, months):
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц", "EBITDA margin"]
    rows_data = []
    for m in months:
        rev = data["revenue"][m]
        net = data["net"][m]
        margin = (net / rev) if rev else 0.0
        rows_data.append([_MONTHS_RU_CAP[m], round(margin, 4)])
    _write_month_table(ws, header_row, headers, rows_data, number_fmt='0.0%')

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=2,
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 20
    chart.height = 10
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    chart.series[0].graphicalProperties.line.solidFill = line_color
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True
    chart.y_axis.numFmt = '0.0%'

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '0.0%'
    chart.dLbls.position = "t"

    chart.legend = None

    ws.add_chart(chart, "E4")


def _sheet_ebitda_multi(wb, sheet_title, chart_title, subtitle,
                        series_data, months):
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц"] + [s["name"] for s in series_data]
    rows_data = []
    for m in months:
        row = [_MONTHS_RU_CAP[m]]
        for s in series_data:
            row.append(round(s["values"].get(m, 0.0) or 0.0, 4))
        rows_data.append(row)
    _write_month_table(ws, header_row, headers, rows_data, number_fmt='0.0%')

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=1 + len(series_data),
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None
    chart.y_axis.numFmt = '0.0%'

    for idx, s in enumerate(series_data):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.line.solidFill = s["color"]
            chart.series[idx].graphicalProperties.line.width = 25000
            chart.series[idx].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '0.0%'
    chart.dLbls.position = "t"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")

    return ws


def _sheet_fot(wb, sheet_title, chart_title, subtitle,
               line_color, data, months):
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц", "Доля ФОТ"]
    rows_data = []
    for m in months:
        rev = data["revenue"][m]
        fot = data["fot"][m]
        share = (abs(fot) / rev) if rev else 0.0
        rows_data.append([_MONTHS_RU_CAP[m], round(share, 4)])
    _write_month_table(ws, header_row, headers, rows_data, number_fmt='0.0%')

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=2,
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 20
    chart.height = 10
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    chart.series[0].graphicalProperties.line.solidFill = line_color
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True
    chart.y_axis.numFmt = '0.0%'

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '0.0%'
    chart.dLbls.position = "t"

    chart.legend = None

    ws.add_chart(chart, "E4")


def _sheet_breakeven(wb, sheet_title, chart_title, subtitle,
                     data_revenue, data_breakeven, months):
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц", "Выручка", "Точка безубыточности"]
    rows_data = []
    for m in months:
        rows_data.append([
            _MONTHS_RU_CAP[m],
            round(data_revenue.get(m, 0.0) or 0.0, 2),
            round(data_breakeven.get(m, 0.0) or 0.0, 2),
        ])
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=3,
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 20
    chart.height = 10
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    chart.series[0].graphicalProperties.line.solidFill = "4472C4"
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True

    chart.series[1].graphicalProperties.line.solidFill = "70AD47"
    chart.series[1].graphicalProperties.line.width = 25000
    chart.series[1].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "t"

    chart.legend.position = "t"
    chart.legend.overlay = False

    ws.add_chart(chart, "E4")


def _sheet_dds(wb, sheet_title, chart_title, subtitle,
               data_inflow, data_outflow, months):
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц", "Поступления", "Выбытия"]
    rows_data = []
    for m in months:
        rows_data.append([
            _MONTHS_RU_CAP[m],
            round(data_inflow.get(m, 0.0) or 0.0, 2),
            round(data_outflow.get(m, 0.0) or 0.0, 2),
        ])
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=3,
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.overlap = -10
    chart.gapWidth = 60
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 20
    chart.height = 10
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    chart.series[0].graphicalProperties.solidFill = "70AD47"
    chart.series[0].graphicalProperties.line.solidFill = "70AD47"
    chart.series[1].graphicalProperties.solidFill = "C00000"
    chart.series[1].graphicalProperties.line.solidFill = "C00000"

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "outEnd"

    chart.legend.position = "b"
    chart.legend.overlay = False

    ws.add_chart(chart, "E4")


def _sheet_revenue_by_units(wb, sheet_title, chart_title, subtitle,
                            series_data, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц"] + [s["name"] for s in series_data]
    rows_data = []
    for m in months:
        row = [_MONTHS_RU_CAP[m]]
        for s in series_data:
            row.append(round(s["values"].get(m, 0.0) or 0.0, 2))
        rows_data.append(row)
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=1 + len(series_data),
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    for idx, s in enumerate(series_data):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.line.solidFill = s["color"]
            chart.series[idx].graphicalProperties.line.width = 25000
            chart.series[idx].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "t"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")


def _sheet_ebitda_units(wb, sheet_title, chart_title, subtitle,
                        series_data, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц"] + [s["name"] for s in series_data]
    rows_data = []
    for m in months:
        row = [_MONTHS_RU_CAP[m]]
        for s in series_data:
            row.append(round(s["values"].get(m, 0.0) or 0.0, 4))
        rows_data.append(row)
    _write_month_table(ws, header_row, headers, rows_data, number_fmt='0.0%')

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=1 + len(series_data),
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None
    chart.y_axis.numFmt = '0.0%'

    for idx, s in enumerate(series_data):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.line.solidFill = s["color"]
            chart.series[idx].graphicalProperties.line.width = 25000
            chart.series[idx].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '0.0%'
    chart.dLbls.position = "t"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")


def _sheet_fot_share_full(wb, sheet_title, chart_title, subtitle,
                          series_data, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц"] + [s["name"] for s in series_data]
    rows_data = []
    for m in months:
        row = [_MONTHS_RU_CAP[m]]
        for s in series_data:
            row.append(round(s["values"].get(m, 0.0) or 0.0, 4))
        rows_data.append(row)
    _write_month_table(ws, header_row, headers, rows_data, number_fmt='0.0%')

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=1 + len(series_data),
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None
    chart.y_axis.numFmt = '0.0%'

    for idx, s in enumerate(series_data):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.line.solidFill = s["color"]
            chart.series[idx].graphicalProperties.line.width = 25000
            chart.series[idx].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '0.0%'
    chart.dLbls.position = "t"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")


def _sheet_ebitda_common(wb, sheet_title, chart_title, subtitle,
                         data_values, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц", "EBITDA margin"]
    rows_data = []
    for m in months:
        rows_data.append([
            _MONTHS_RU_CAP[m],
            round(data_values.get(m, 0.0) or 0.0, 4),
        ])
    _write_month_table(ws, header_row, headers, rows_data, number_fmt='0.0%')

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=2,
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None
    chart.y_axis.numFmt = '0.0%'

    chart.series[0].graphicalProperties.line.solidFill = "70AD47"
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '0.0%'
    chart.dLbls.position = "t"

    chart.legend = None
    ws.add_chart(chart, "H4")


def _sheet_profit_dynamics(wb, sheet_title, chart_title, subtitle,
                           series_data, months):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц"] + [s["name"] for s in series_data]
    rows_data = []
    for m in months:
        row = [_MONTHS_RU_CAP[m]]
        for s in series_data:
            row.append(round(s["values"].get(m, 0.0) or 0.0, 2))
        rows_data.append(row)
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=1 + len(series_data),
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.gapWidth = 60
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 24
    chart.height = 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    for idx, s in enumerate(series_data):
        if idx < len(chart.series):
            chart.series[idx].graphicalProperties.solidFill = s["color"]
            chart.series[idx].graphicalProperties.line.solidFill = "FFFFFF"
            chart.series[idx].graphicalProperties.line.width = 10000

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "ctr"

    chart.legend.position = "t"
    chart.legend.overlay = False
    ws.add_chart(chart, "H4")


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def _sheet_name_with_slide(slide_no, base_name):
    return _safe_sheet(f"Сл{slide_no:02d}_{base_name}")


def _parse_bdr_plan(bdr_path, unit_name, months):
    result = {m: 0.0 for m in months}
    if not bdr_path or not os.path.exists(bdr_path):
        return result

    wb = openpyxl.load_workbook(bdr_path, data_only=True)
    ws = wb.active

    header_row = None
    for r in range(1, 6):
        hits = 0
        for c in range(2, min(ws.max_column, 20) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            sv = str(v).lower()
            if any(m in sv for m in _MONTHS_RU_LOWER.values()):
                hits += 1
        if hits >= 3:
            header_row = r
            break

    if header_row is None:
        return result

    month_col = _find_month_columns(ws, header_row, months)

    cur_section = None
    for r in range(header_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue
        title = _clean(raw)
        lvl = _level(raw)

        if lvl == 0:
            cur_section = title
            continue
        if lvl == 1 and cur_section == "Чистая прибыль" and title == unit_name:
            for m, col in month_col.items():
                result[m] = _num(ws, r, col)
            return result

    return result


def _parse_bdr_plan_object(bdr_path, unit_name, object_name, months):
    result = {m: 0.0 for m in months}
    if not bdr_path or not os.path.exists(bdr_path):
        return result

    wb = openpyxl.load_workbook(bdr_path, data_only=True)
    ws = wb.active

    header_row = None
    for r in range(1, 6):
        hits = 0
        for c in range(2, min(ws.max_column, 20) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            sv = str(v).lower()
            if any(m in sv for m in _MONTHS_RU_LOWER.values()):
                hits += 1
        if hits >= 3:
            header_row = r
            break

    if header_row is None:
        return result

    month_col = _find_month_columns(ws, header_row, months)

    cur_section = None
    cur_unit = None
    for r in range(header_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue
        title = _clean(raw)
        lvl = _level(raw)

        if lvl == 0:
            cur_section = title
            cur_unit = None
            continue
        if lvl == 1:
            cur_unit = title
            continue
        if (lvl == 2
                and cur_section == "Чистая прибыль"
                and cur_unit == unit_name
                and title == object_name):
            for m, col in month_col.items():
                result[m] = _num(ws, r, col)
            return result

    return result


def _sheet_plan_fact(wb, sheet_title, chart_title, subtitle,
                     data_plan, data_fact, months):
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Месяц", "План", "Факт"]
    rows_data = []
    for m in months:
        rows_data.append([
            _MONTHS_RU_CAP[m],
            round(data_plan.get(m, 0.0), 2),
            round(data_fact.get(m, 0.0), 2),
        ])
    _write_month_table(ws, header_row, headers, rows_data)

    cats = Reference(ws, min_col=1,
                     min_row=header_row + 1,
                     max_row=header_row + len(months))
    data_ref = Reference(ws, min_col=2, max_col=3,
                         min_row=header_row,
                         max_row=header_row + len(months))

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.overlap = -10
    chart.gapWidth = 60
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = chart_title
    chart.width = 20
    chart.height = 10
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = None

    chart.series[0].graphicalProperties.solidFill = _COLOR_LINE_YEL
    chart.series[1].graphicalProperties.solidFill = _COLOR_NET

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "outEnd"

    chart.legend.position = "b"
    chart.legend.overlay = False

    ws.add_chart(chart, "E4")


def _sheet_operational_balance(wb, sheet_title, chart_title, subtitle,
                               rows_data, months):
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    header_row = 4
    headers = ["Направление"] + [_MONTHS_RU_CAP[m] for m in months]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True)
        cell.border = _CELL_BORDER

    for i, row in enumerate(rows_data):
        r = header_row + 1 + i
        name_cell = ws.cell(row=r, column=1, value=row["name"])
        name_cell.border = _CELL_BORDER
        name_cell.alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=False)

        if row.get("is_total"):
            name_cell.font = Font(bold=True, size=11, color="1E3A8A")

        for j, m in enumerate(months):
            v = row["values"].get(m, 0.0)
            cell = ws.cell(row=r, column=2 + j, value=round(v, 2))
            cell.border = _CELL_BORDER
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal="right", vertical="center")

            if v > 0.005:
                cell.font = Font(color="00B050", bold=True)
            elif v < -0.005:
                cell.font = Font(color="0070C0", bold=True)

            if row.get("is_total"):
                cell.font = Font(
                    color=("00B050" if v > 0.005 else
                           "0070C0" if v < -0.005 else "000000"),
                    bold=True,
                )

    ws.column_dimensions["A"].width = 45
    for c in range(2, 2 + len(months)):
        col_letter = ws.cell(row=header_row, column=c).column_letter
        ws.column_dimensions[col_letter].width = 12


def _find_ku_sheet(wb_ku, month_name):
    def _looks_like_ku(ws):
        header_row = None
        for r in range(1, 20):
            v = _clean(ws.cell(row=r, column=1).value)
            if v and "объекты" in v.lower():
                header_row = r
                break
        if header_row is None:
            return False
        has_total = False
        for r in range(header_row + 1,
                       min(header_row + 30, ws.max_row + 1)):
            v = _clean(ws.cell(row=r, column=1).value)
            if v and "итого" in v.lower():
                has_total = True
                break
        if not has_total:
            return False
        for r in range(header_row + 1,
                       min(header_row + 5, ws.max_row + 1)):
            e_val = _num(ws, r, 5)
            if e_val > 0:
                return True
            if e_val < 0:
                return False
        return False

    for ws in wb_ku.worksheets:
        if _looks_like_ku(ws):
            return ws

    if month_name in wb_ku.sheetnames:
        return wb_ku[month_name]

    for ws in wb_ku.worksheets:
        for r in range(1, 20):
            v = _clean(ws.cell(row=r, column=1).value)
            if v and "объекты" in v.lower():
                return ws

    return wb_ku.active


def _find_odds_month_col(ws_ku, month_name):
    header_row = None
    for r in range(1, 8):
        v = _clean(ws_ku.cell(row=r, column=1).value)
        if v and "направление" in v.lower():
            header_row = r
            break
    if header_row is None:
        return None

    mn = month_name.lower()
    for c in range(2, ws_ku.max_column + 1):
        v = ws_ku.cell(row=header_row, column=c).value
        if v is None:
            continue
        sv = str(v).lower()
        if mn in sv:
            return c
    return None


def _sheet_utility_reimbursement(wb, sheet_title, chart_title, subtitle,
                                  ku_path, month_name, subtract_object=None):
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    header_row_1 = 3
    header_row_2 = 4
    header_row_3 = 5
    header_row = header_row_3

    row3_headers = [
        "EUR",
        "Разница между выставленными счетами и фактическими коммунальными расходами с НДС с начала года",
        "Выставлено в сентябре (за август 2026) арендаторам",
        "в т.ч. НДС",
        "1.2.10.1 Мусор*",
        "1.2.10.2 Газ*",
        "1.2.10.3 Вода*",
        "1.2.10.4 Отопление*",
        "1.2.10.5 Электричество*",
        "1.2.10.6 Коммунальные УК дома*",
        "Всего расходов на коммунальные услуги*",
        "Разница между выставленными счетами и фактическими коммунальными расходами за отчётный месяц (гр. 3 -гр.11)",
        "1.1.2.3 Компенсация по коммунальным расходам-Поступившая на счет",
        "Задолженность по возмещению КУ за отчётный месяц (гр.13 - гр.3)",
    ]
    for c, h in enumerate(row3_headers, start=1):
        cell = ws.cell(row=header_row_1, column=c, value=h)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True)
        cell.border = _CELL_BORDER

    ws.cell(row=header_row_2, column=1, value="Объекты").fill = _HEADER_FILL
    ws.cell(row=header_row_2, column=1).font = _HEADER_FONT
    ws.cell(row=header_row_2, column=1).alignment = Alignment(
        horizontal="center", vertical="center")
    ws.cell(row=header_row_2, column=1).border = _CELL_BORDER
    for c in range(2, 15):
        cell = ws.cell(row=header_row_2, column=c)
        cell.fill = _HEADER_FILL
        cell.border = _CELL_BORDER

    for c in range(1, 15):
        cell = ws.cell(row=header_row_3, column=c, value=c)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _CELL_BORDER

    if not ku_path or not os.path.exists(ku_path):
        print(f"[presentation_builder] не найден файл КУ: {ku_path}")
        return ws

    wb_ku = openpyxl.load_workbook(ku_path, data_only=True)
    ws_ku = wb_ku.active

    month_col = _find_odds_month_col(ws_ku, month_name)
    if month_col is None:
        print(f"[presentation_builder] в файле {ku_path} "
              f"не найдена колонка для месяца «{month_name}»")
        return ws

    def _category(t):
        tl = t.lower()
        if "1.2.10.1" in tl:
            return "musor"
        if "1.2.10.2" in tl:
            return "gaz"
        if "1.2.10.3" in tl:
            return "voda"
        if "1.2.10.4" in tl:
            return "otopl"
        if "1.2.10.5" in tl:
            return "elektr"
        if "1.2.10.6" in tl:
            return "uk_doma"
        return None

    data_by_obj = {}
    cur_obj = None
    cur_obj_level = None

    for r in range(1, ws_ku.max_row + 1):
        raw = ws_ku.cell(row=r, column=1).value
        if raw is None:
            continue
        title = _clean(raw)
        lvl = _level(raw)

        if lvl == 0:
            cur_obj = None
            cur_obj_level = None
            continue

        if lvl == 1:
            cur_obj = title
            cur_obj_level = lvl
            if cur_obj not in data_by_obj:
                data_by_obj[cur_obj] = {
                    "musor": 0.0, "gaz": 0.0, "voda": 0.0,
                    "otopl": 0.0, "elektr": 0.0, "uk_doma": 0.0,
                }
            continue

        if cur_obj is None:
            continue
        if lvl <= cur_obj_level:
            cur_obj = None
            cur_obj_level = None
            continue

        cat = _category(title)
        if cat is None:
            continue
        val = _num(ws_ku, r, month_col)
        data_by_obj[cur_obj][cat] += abs(val)

    row_out = header_row + 1
    data_rows = []

    for ku_name in config.KU_OBJECTS:
        odd_name = config.KU_TO_ODDS_NAME.get(ku_name, ku_name)

        if subtract_object and ku_name == subtract_object:
            continue

        d = data_by_obj.get(odd_name)
        if d is None:
            ws.cell(row=row_out, column=1, value=ku_name).border = _CELL_BORDER
            for c in range(2, 15):
                ws.cell(row=row_out, column=c).border = _CELL_BORDER
            data_rows.append(row_out)
            row_out += 1
            continue

        ws.cell(row=row_out, column=1, value=ku_name).border = _CELL_BORDER
        ws.cell(row=row_out, column=2).border = _CELL_BORDER
        ws.cell(row=row_out, column=3).border = _CELL_BORDER
        ws.cell(row=row_out, column=4).border = _CELL_BORDER

        ws.cell(row=row_out, column=5,
                value=round(d["musor"], 2)).border = _CELL_BORDER
        ws.cell(row=row_out, column=6,
                value=round(d["gaz"], 2)).border = _CELL_BORDER
        ws.cell(row=row_out, column=7,
                value=round(d["voda"], 2)).border = _CELL_BORDER
        ws.cell(row=row_out, column=8,
                value=round(d["otopl"], 2)).border = _CELL_BORDER
        ws.cell(row=row_out, column=9,
                value=round(d["elektr"], 2)).border = _CELL_BORDER
        ws.cell(row=row_out, column=10,
                value=round(d["uk_doma"], 2)).border = _CELL_BORDER

        ws.cell(row=row_out, column=11,
                value=f"=SUM(E{row_out}:J{row_out})").border = _CELL_BORDER
        ws.cell(row=row_out, column=11).number_format = '#,##0'

        ws.cell(row=row_out, column=12).border = _CELL_BORDER
        ws.cell(row=row_out, column=13).border = _CELL_BORDER
        ws.cell(row=row_out, column=14).border = _CELL_BORDER

        for c in range(5, 11):
            ws.cell(row=row_out, column=c).number_format = '#,##0'

        data_rows.append(row_out)
        row_out += 1

    total_row = row_out
    ws.cell(row=total_row, column=1, value="Итого").font = Font(bold=True)
    ws.cell(row=total_row, column=1).border = _CELL_BORDER
    for c in range(2, 15):
        ws.cell(row=total_row, column=c).border = _CELL_BORDER

    if data_rows:
        first_data = data_rows[0]
        last_data = data_rows[-1]
        for c in range(5, 12):
            col_letter = get_column_letter(c)
            formula = f"=SUM({col_letter}{first_data}:{col_letter}{last_data})"
            cell = ws.cell(row=total_row, column=c, value=formula)
            cell.font = Font(bold=True)
            cell.number_format = '#,##0'
            cell.border = _CELL_BORDER
    else:
        for c in range(5, 12):
            ws.cell(row=total_row, column=c, value=0).border = _CELL_BORDER

    ws.column_dimensions["A"].width = 30
    for c in range(2, 15):
        ws.column_dimensions[get_column_letter(c)].width = 14

    return ws


def _collect_dds(odds_path, unit_name, months):
    result = {
        "inflow":  {m: 0.0 for m in months},
        "outflow": {m: 0.0 for m in months},
    }
    if not odds_path or not os.path.exists(odds_path):
        return result

    wb = openpyxl.load_workbook(odds_path, data_only=True)
    ws = wb.active

    header_row = None
    for r in range(1, 8):
        v = ws.cell(row=r, column=1).value
        if v and "направление" in str(v).lower():
            header_row = r
            break
    if header_row is None:
        return result

    month_col = _find_month_columns(ws, header_row, months)

    unit_row = None
    for r in range(header_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue
        title = _clean(raw)
        lvl = _level(raw)
        if lvl == 0 and title == unit_name:
            unit_row = r
            break

    if unit_row is None:
        return result

    unit_values = {}
    for m, col in month_col.items():
        unit_values[m] = _num(ws, unit_row, col) or 0.0

    has_positive_in_unit_row = any(v > 0 for v in unit_values.values())

    if has_positive_in_unit_row:
        inflow_row  = unit_row
        outflow_row = unit_row + 1

        for m, col in month_col.items():
            result["inflow"][m]  = _num(ws, inflow_row, col) or 0.0
            result["outflow"][m] = abs(_num(ws, outflow_row, col) or 0.0)

        if config.DEBUG:
            print(f"[DDS] {unit_name}: вариант A "
                  f"(строка {unit_row} = Поступления, "
                  f"строка {outflow_row} = Списания)")

    else:
        outflow_row = unit_row + 1
        inflow_row  = unit_row + 2

        for m, col in month_col.items():
            result["outflow"][m] = abs(_num(ws, outflow_row, col) or 0.0)
            result["inflow"][m]  = _num(ws, inflow_row, col) or 0.0

        if config.DEBUG:
            print(f"[DDS] {unit_name}: вариант B "
                  f"(строка {outflow_row} = Списания, "
                  f"строка {inflow_row} = Поступления)")

    return result


def _parse_odds_balance(odds_path, unit_name, months):
    result = {"rows": []}
    if not odds_path or not os.path.exists(odds_path):
        return result

    wb = openpyxl.load_workbook(odds_path, data_only=True)
    ws = wb.active

    header_row = None
    for r in range(1, 8):
        v = ws.cell(row=r, column=1).value
        if v and "направление" in str(v).lower():
            header_row = r
            break
    if header_row is None:
        return result

    month_col = _find_month_columns(ws, header_row, months)

    unit_row = None
    for r in range(header_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue
        title = _clean(raw)
        lvl = _level(raw)
        if lvl == 0 and title == unit_name:
            unit_row = r
            break

    if unit_row is None:
        return result

    objects = []
    cur_object_name = None
    cur_object_level = None

    for r in range(unit_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue
        title = _clean(raw)
        lvl = _level(raw)

        if lvl == 0 and title != unit_name:
            break

        if lvl == 1:
            if title.lower() in ("сальдо", "итого",
                                 "сальдо по направлению"):
                values = {m: _num(ws, r, col) for m, col in month_col.items()}
                result["rows"].append({
                    "name": "ИТОГО",
                    "values": values,
                    "is_total": True,
                })
                cur_object_name = None
                continue

            cur_object_name = title
            cur_object_level = lvl
            continue

        if lvl == 2 and cur_object_name is not None:
            if title.lower() in ("сальдо", "итого"):
                values = {m: _num(ws, r, col) for m, col in month_col.items()}
                objects.append({
                    "name": cur_object_name,
                    "values": values,
                })
                cur_object_name = None
                cur_object_level = None
            continue

    rows_out = []
    total_row = None
    for row in result["rows"]:
        if row.get("is_total"):
            total_row = row
        else:
            rows_out.append(row)

    for o in objects:
        rows_out.append(o)

    if total_row is not None:
        rows_out.append(total_row)

    result["rows"] = rows_out
    return result


def _sheet_cumulative_plan_fact(wb, sheet_title, chart_title,
                                data_plan, data_fact, months,
                                plan_label, fact_label):
    ws = wb.create_sheet(sheet_title)

    header_row = 2
    fact_row   = 3
    plan_row   = 4

    ws.cell(row=header_row, column=1, value="Показатель").font = Font(bold=True)
    for i, m in enumerate(months, start=2):
        ws.cell(row=header_row, column=i,
                value=_MONTHS_RU_LOWER[m].capitalize()).font = Font(bold=True)

    ws.cell(row=fact_row, column=1, value="Факт").font = Font(bold=True)
    ws.cell(row=plan_row, column=1, value="План").font = Font(bold=True)

    for i, m in enumerate(months, start=2):
        hc = ws.cell(row=header_row, column=i)
        hc.font = Font(color="FFFFFF", size=8)

        fc = ws.cell(row=fact_row, column=i,
                     value=float(data_fact.get(m, 0.0) or 0.0))
        fc.number_format = '#,##0'
        fc.font = Font(color="FFFFFF", size=8)

        pc = ws.cell(row=plan_row, column=i,
                     value=float(data_plan.get(m, 0.0) or 0.0))
        pc.number_format = '#,##0'
        pc.font = Font(color="FFFFFF", size=8)

    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100

    data_ref = Reference(
        ws,
        min_col=2, max_col=1 + len(months),
        min_row=header_row, max_row=plan_row,
    )
    cats_ref = Reference(
        ws,
        min_col=1,
        min_row=fact_row, max_row=plan_row,
    )

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    for idx, series in enumerate(chart.series):
        if idx < len(months):
            m = months[idx]
            color = _MONTH_FILL.get(m, "808080")
            series.graphicalProperties.solidFill = color
            series.graphicalProperties.line.solidFill = "FFFFFF"
            series.graphicalProperties.line.width = 10000

    chart.width = 32
    chart.height = 11
    chart.title = chart_title

    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showSerName = False
    chart.dLbls.showCatName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.numFmt = '#,##0'
    chart.dLbls.position = "ctr"

    chart.legend.position = "t"
    chart.legend.overlay = False

    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.numFmt = '#,##0'
    chart.x_axis.majorTickMark = "out"
    chart.x_axis.tickLblPos = "nextTo"
    chart.x_axis.title = "Периоды"

    chart.y_axis.majorTickMark = "out"
    chart.y_axis.tickLblPos = "nextTo"

    chart.x_axis.majorGridlines = None
    chart.y_axis.majorGridlines = None

    from openpyxl.utils import column_index_from_string
    anchor_col_idx = column_index_from_string("A") - 1
    anchor_row_idx = 6 - 1
    anchor = OneCellAnchor(
        _from=AnchorMarker(
            col=anchor_col_idx,
            colOff=0,
            row=anchor_row_idx,
            rowOff=0,
        ),
        ext=XDRPositiveSize2D(
            cx=cm_to_EMU(chart.width),
            cy=cm_to_EMU(chart.height),
        ),
    )
    chart.anchor = anchor

    ws.add_chart(chart)

    sum_plan = sum(float(data_plan.get(m, 0.0) or 0.0) for m in months)
    sum_fact = sum(float(data_fact.get(m, 0.0) or 0.0) for m in months)
    pct = (sum_fact / sum_plan) if sum_plan else 0.0

    box_col = max(13, len(months) + 5)
    box_row = 2

    for i, (val, fmt) in enumerate([
        (round(sum_plan, 0), '#,##0'),
        (round(sum_fact, 0), '#,##0'),
        (round(pct, 4),      '0.0%'),
    ]):
        c_val = ws.cell(row=box_row + i, column=box_col, value=val)
        c_val.font = Font(bold=True, color="FFC000", size=14)
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.fill = PatternFill("solid", fgColor="1F3864")
        c_val.border = _CELL_BORDER
        c_val.number_format = fmt

    ws.column_dimensions["A"].width = 22
    for i in range(2, 2 + len(months)):
        ws.column_dimensions[get_column_letter(i)].width = 10
    ws.column_dimensions[get_column_letter(box_col)].width = 14

    ws.sheet_view.showGridLines = False
    return ws


_MONTH_FILL = {
    1: "4472C4", 2: "8FAADC", 3: "FFC000", 4: "548235",
    5: "ED7D31", 6: "9DC3E6", 7: "7030A0", 8: "BFB100",
    9: "2E4C99", 10: "C55A11", 11: "375623", 12: "203864",
}

_TOTAL_BAR_COLUMNS = 50


def build_presentation_data(opiu_path, bdr_path, forecast_path,
                            year, month, output_path):
    months = list(range(1, month + 1))

    rows, _ = _read_flat_rows(opiu_path, months)
    if not rows:
        print("[presentation_builder] нет данных — файл пустой")
        return output_path

    units = _collect_units(rows, months)
    if not units:
        print("[presentation_builder] _collect_units вернул пусто")
        return output_path

    # НОВЫЙ расчёт ТБУ — с разделением на переменные и постоянные
    breakeven_data = _collect_breakeven(rows, months)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    unit_slides = {
        "Latvia":       (_UNIT_SLIDE_START["Latvia"],     5,  7),
        "East-Восток":  (_UNIT_SLIDE_START["East-Восток"], 42, 40),
        "Europe":       (_UNIT_SLIDE_START["Europe"],     51, 52),
        "Nomiqa":       (_UNIT_SLIDE_START["Nomiqa"],     73, 0),
        "Unelma":       (_UNIT_SLIDE_START["Unelma"],     68, 0),
        "UK Estate":    (0, 0, 0),
    }

    for unit_name in _UNITS:
        u = units.get(unit_name)
        if not u:
            continue

        sl_opiu, sl_ebitda, sl_fot = unit_slides.get(
            unit_name, (_UNIT_SLIDE_START.get(unit_name, 1), 0, 0)
        )

        if unit_name == "Latvia":
            ebitda_color = _COLOR_LINE_RED
            fot_color    = _COLOR_LINE_RED
        elif unit_name == "East-Восток":
            ebitda_color = _COLOR_LINE_YEL
            fot_color    = _COLOR_LINE_RED
        else:
            ebitda_color = _COLOR_LINE_YEL
            fot_color    = _COLOR_LINE_RED

        if unit_name != "UK Estate":
            _sheet_revenue_profit(
                wb,
                sheet_title=_sheet_name_with_slide(sl_opiu, f"ОПиУ_{unit_name}"),
                chart_title=f"Выручка & Чистая прибыль {unit_name}, без НДС",
                subtitle=f"ОПиУ {unit_name}",
                color_rev=_COLOR_REVENUE,
                color_net=_COLOR_NET,
                data=u,
                months=months,
            )

        if sl_ebitda:
            if unit_name == "Latvia":
                ant_objects = ["AN14 Антониас 14 (дом + парковка)"]
                chaka_objects = ["AC89 Чака 89 (дом + парковка)"]
                commercial_objects = [
                    "D4 Парковка-Deglava4",
                    "AC87 Гараж Чака",
                    "B117 Бривибас, 117",
                    "B78 Бривибас, 78",
                    "C23 Цесу, 23",
                    "DAR1_Darzauglu1",
                    "DS1 Дзирнаву, 1",
                    "G73 Гертрудес, 73",
                    "H5 Хоспиталю",
                    "MP1_Marupe",
                    "MU3 - Mucenieku 3 - 4",
                    "OZ1 Озолниеки",
                    "SK3-Skunju 3",
                    "UK_Latvia",
                    "V22 К. Валдемара 22",
                    "BRN_Brunieku",
                ]

                def _sum_objects(obj_list, key):
                    result = {m: 0.0 for m in months}
                    for obj_name in obj_list:
                        obj = u["objects"].get(obj_name)
                        if not obj:
                            continue
                        for m in months:
                            result[m] += obj[key].get(m, 0.0) or 0.0
                    return result

                series_lat = [
                    {
                        "name": "Latvia",
                        "color": _COLOR_LINE_RED,
                        "net": u["net"].copy(),
                        "revenue": u["revenue"].copy(),
                    },
                    {
                        "name": "Антонияс",
                        "color": "70AD47",
                        "net": _sum_objects(ant_objects, "net"),
                        "revenue": _sum_objects(ant_objects, "revenue"),
                    },
                    {
                        "name": "Чака 89",
                        "color": "FFC000",
                        "net": _sum_objects(chaka_objects, "net"),
                        "revenue": _sum_objects(chaka_objects, "revenue"),
                    },
                    {
                        "name": "Коммерческие",
                        "color": "4472C4",
                        "net": _sum_objects(commercial_objects, "net"),
                        "revenue": _sum_objects(commercial_objects, "revenue"),
                    },
                ]

                series_data_05 = []
                for s in series_lat:
                    margin = {}
                    for m in months:
                        rev = s["revenue"].get(m, 0.0) or 0.0
                        net = s["net"].get(m, 0.0) or 0.0
                        margin[m] = (net / rev) if rev else 0.0
                    series_data_05.append({
                        "name": s["name"],
                        "color": s["color"],
                        "values": margin,
                    })

                _sheet_ebitda_multi(
                    wb,
                    sheet_title=_sheet_name_with_slide(
                        sl_ebitda, f"EBITDA_{unit_name}"),
                    chart_title="Рентабельность по чистой прибыли Latvia",
                    subtitle="Рентабельность по ЧП (ЧП / Выручка × 100%), %",
                    series_data=series_data_05,
                    months=months,
                )
            else:
                _sheet_ebitda(
                    wb,
                    sheet_title=_sheet_name_with_slide(
                        sl_ebitda, f"EBITDA_{unit_name}"),
                    chart_title=f"Операционная рентабельность (EBITDA margin) {unit_name}",
                    subtitle=f"EBITDA margin {unit_name}, %",
                    line_color=ebitda_color,
                    data=u,
                    months=months,
                )

        if sl_fot:
            _sheet_fot(
                wb,
                sheet_title=_sheet_name_with_slide(
                    sl_fot, f"Доля_ФОТ_{unit_name}"),
                chart_title=f"Доля ФОТ в выручке {unit_name}",
                subtitle=f"Анализ ОПиУ Доля ФОТ {unit_name}",
                line_color=fot_color,
                data=u,
                months=months,
            )

        if bdr_path and os.path.exists(bdr_path):
            plan_vals = _parse_bdr_plan(bdr_path, unit_name, months)
            if any(abs(v) > 0.01 for v in plan_vals.values()):
                sl_plan_fact_map = {
                    "Latvia":       8,
                    "East-Восток":  41,
                    "Europe":       53,
                    "Nomiqa":       75,
                    "Unelma":       70,
                    "UK Estate":    0,
                }
                sl_plan_fact = sl_plan_fact_map.get(unit_name, 0)

                if sl_plan_fact:
                    _sheet_plan_fact(
                        wb,
                        sheet_title=_sheet_name_with_slide(
                            sl_plan_fact, f"ПланФакт_{unit_name}"),
                        chart_title=f"Выполнение годового плана по ЧП {unit_name}",
                        subtitle=f"План (БДиР) vs Факт (ОПиУ) — {unit_name}",
                        data_plan=plan_vals,
                        data_fact=u["net"],
                        months=months,
                    )

                sl_cumulative = {
                    "Latvia":      9,
                    "East-Восток": 43,
                    "Europe":      0,
                    "Nomiqa":      74,
                    "Unelma":      69,
                    "UK Estate":   0,
                }.get(unit_name)

                if sl_cumulative:
                    _sheet_cumulative_plan_fact(
                        wb,
                        sheet_title=_sheet_name_with_slide(
                            sl_cumulative, f"Итог_{unit_name}"),
                        chart_title=f"Выполнение годового плана по чистой прибыли {unit_name}",
                        data_plan=plan_vals,
                        data_fact=u["net"],
                        months=months,
                        plan_label=f"ЧП {unit_name} план (с НДС)",
                        fact_label=f"ЧП {unit_name} факт (с НДС)",
                    )

    # ... остальной код без изменений ...

    # СОРТИРОВКА И СОХРАНЕНИЕ
    _sort_sheets_by_slide_number(wb)
    wb.save(output_path)
    print(f"[OK] Диаграммы сохранены: {output_path}")
    return output_path


def _sort_sheets_by_slide_number(wb):
    import re as _re

    def _slide_key(ws):
        m = _re.search(r"Сл(\d+)", ws.title)
        if m:
            return (0, int(m.group(1)), ws.title)
        return (1, 10**9, ws.title)

    try:
        desired = sorted(wb.worksheets, key=_slide_key)
        wb._sheets = desired
        print(f"[presentation_builder] листы отсортированы, порядок: "
              f"{[ws.title[:10] for ws in wb.worksheets]}")
    except Exception as e:
        print(f"[presentation_builder] не удалось отсортировать листы: {e}")