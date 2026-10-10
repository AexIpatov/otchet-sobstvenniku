# src/presentation_builder.py
# ============================================================
# Генератор Excel с диаграммами для презентации.
#
# Читает ОПиУ формата «статьи × месяцы» с древовидной структурой
# (·, ··, ···, ····) и строит книгу Excel.
#
# Имена листов содержат номер слайда презентации:
#     Сл04_ОПиУ_Латвия
#     Сл05_EBITDA_Латвия
#     Сл07_Доля_ФОТ_Латвия
#     ...
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

# Юниты: ТОЧНЫЕ названия строк уровня 1 из файла
_UNITS = [
    "Latvia",
    "East-Восток",
    "Europe",
    "Nomiqa",
    "Unelma",
    "UK Estate",
]

# Объекты Латвии для отдельных листов:
#   (имя в файле ОПиУ, имя для листа, номер слайда презентации)
_LATVIA_OBJECTS = [
    ("AN14 Антониас 14 (дом + парковка)", "Антонияс", 10),
    ("AC89 Чака 89 (дом + парковка)",     "Чака",     15),
    ("M81 - Matisa 81",                    "Матиса",   32),
    ("EKS_Esporta iela 12-113",            "Эспорта",  28),
]

# Коммерческие объекты Латвии (собираем в один виртуальный блок)
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
]

# Номера слайдов для юнитов (первый слайд с ОПиУ каждого юнита)
_UNIT_SLIDE_START = {
    "Latvia":        4,   # слайд 4 — ОПиУ Латвия
    "East-Восток":  38,   # слайд 38 — ОПиУ East-Восток
    "Europe":       49,   # слайд 49 — ОПиУ Европа
    "Nomiqa":       71,   # слайд 71 — ОПиУ Nomiqa
    "Unelma":       67,   # слайд 67 — ОПиУ Унелма
    "UK Estate":    56,   # слайд 56 — ОПиУ Расходы УК
}

# Стили
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

# Цвета серий (как в презентации)
_COLOR_REVENUE = "FFC000"   # жёлтый  (Выручка)
_COLOR_NET     = "70AD47"   # зелёный (ЧП)
_COLOR_LINE_RED = "C00000"  # красный (EBITDA Латвия, Доля ФОТ Латвия)
_COLOR_LINE_YEL = "FFC000"  # жёлтый  (EBITDA Чака, EBITDA Коммерческие, EBITDA Матиса)
_COLOR_LINE_BLU = "4472C4"  # синий   (Доля ФОТ — общий цвет)


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
    """Возвращает {month: col_idx} по названиям месяцев в шапке."""
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

def _collect_breakeven(rows, months):
    """
    Возвращает {unit_name: {month: value}} — точку безубыточности
    по каждому юниту.

    Логика:
      ТБУ = |Производственные расходы| + |Косвенные расходы|.

    ВАЖНО: в ОПиУ секции «Производственные расходы» и «Косвенные расходы»
    имеют РАЗНУЮ структуру вложенности:

      Производственные расходы             ← level 0
      · Прямые производственные            ← level 1 (промежуточный)
      ·· Latvia                            ← level 2 (юнит)
      ··· AN14 …                           ← level 3 (объект)
      ···· 1.2.10 Коммунальные платежи     ← level 4 (статья)

      Косвенные расходы                    ← level 0
      · Административные расходы           ← level 1 (категория)
      ·· Latvia                            ← level 2 (юнит)
      ··· AN14 …                           ← level 3 (объект)

    Поэтому ищем юнит по ИМЕНИ (входит в _UNITS), а не по уровню.
    Учитываем только строку-объект, но НЕ его дочерние статьи,
    чтобы не удвоить сумму.
    """
    _SECTION_PROD = "Производственные расходы"
    _SECTION_INDIRECT = "Косвенные расходы"

    result = {u: {m: 0.0 for m in months} for u in _UNITS}

    cur_section = None
    cur_unit = None
    cur_unit_level = None

    for r, lvl, title, values in rows:
        # level 0 — переключение секции
        if lvl == 0:
            cur_section = title
            cur_unit = None
            cur_unit_level = None
            continue

        # Ищем юнит ПО ИМЕНИ (не по уровню)
        # Обновляем cur_unit, если это имя из _UNITS.
        if title in _UNITS:
            cur_unit = title
            cur_unit_level = lvl
            continue

        # Работаем только в нужных секциях и внутри какого-то юнита
        if cur_section not in (_SECTION_PROD, _SECTION_INDIRECT):
            continue
        if cur_unit is None:
            continue

        # Берём только СЛЕДУЮЩИЙ уровень после юнита
        # (это объект). Не заходим глубже — иначе задвоим.
        if lvl != cur_unit_level + 1:
            continue

        # Складываем по модулю
        for m in months:
            v = values.get(m, 0.0) or 0.0
            result[cur_unit][m] += abs(v)

    return result

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

    # 1) объекты уровня 2 в секциях Выручка / Чистая прибыль
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

    # 2) revenue / net каждого объекта
    for u_name, u in units.items():
        for obj_name in list(u["objects"].keys()):
            rev = _find_row(rows, _SECTION_REVENUE, u_name, object_name=obj_name)
            net = _find_row(rows, _SECTION_NET, u_name, object_name=obj_name)
            for m in months:
                u["objects"][obj_name]["revenue"][m] = rev.get(m, 0.0)
                u["objects"][obj_name]["net"][m]     = net.get(m, 0.0)

    # 3) ФОТ объектов
    for u_name, u in units.items():
        for obj_name in list(u["objects"].keys()):
            fot_prod = _find_object_fot(rows, _SECTION_PROD, u_name, obj_name)
            fot_comm = _find_object_fot(rows, _SECTION_COMM, u_name, obj_name)
            for m in months:
                u["objects"][obj_name]["fot"][m] = (
                    fot_prod.get(m, 0.0) + fot_comm.get(m, 0.0)
                )

    # 4) агрегаты юнита
    for u_name, u in units.items():
        # «Чистая прибыль» юнита = сумма по его объектам.
        # Это работает для всех юнитов: для Unelma/Nomiqa
        # (у которых один объект), для Latvia/East-Восток/Europe
        # (у которых много объектов), для UK Estate (объектов нет — 0).
        for m in months:
            u["net"][m]     = sum(o["net"][m]     for o in u["objects"].values())
            u["revenue"][m] = sum(o["revenue"][m] for o in u["objects"].values())
            u["fot"][m]     = sum(o["fot"][m]     for o in u["objects"].values())

    # 5) виртуальный юнит «Коммерческие»
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


# ------------------------------------------------------------
# Лист: ОПиУ юнита/объекта (Выручка + ЧП, столбики)
# ------------------------------------------------------------
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


# ------------------------------------------------------------
# Лист: EBITDA (линия)
# ------------------------------------------------------------
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


# ------------------------------------------------------------
# Лист: Доля ФОТ (линия)
# ------------------------------------------------------------
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
    """
    Лист «Точка безубыточности» — две линии:
      • синяя — Выручка
      • зелёная — Точка безубыточности

    Данные берём помесячно: для каждого месяца своё значение.
    """
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

    # Синяя линия — Выручка
    chart.series[0].graphicalProperties.line.solidFill = "4472C4"
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True

    # Зелёная линия — Точка безубыточности
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

# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def _sheet_name_with_slide(slide_no, base_name):
    """Формирует имя листа вида 'Сл04_ОПиУ_Латвия'."""
    return _safe_sheet(f"Сл{slide_no:02d}_{base_name}")

def _parse_bdr_plan(bdr_path, unit_name, months):
    """
    Читает БДиР (формат «статьи × месяцы») и возвращает
    {month: value} — план по чистой прибыли.

    В БДиР строка «Чистая прибыль» → «· <Юнит>» — это план ЧП
    за каждый месяц. Для 2026 года в БДиР колонки B..M — это
    Август..Декабрь + …
    """
    result = {m: 0.0 for m in months}
    if not bdr_path or not os.path.exists(bdr_path):
        return result

    wb = openpyxl.load_workbook(bdr_path, data_only=True)
    ws = wb.active

    # Ищем шапку (строка с названиями месяцев).
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

    # Ищем строку «Чистая прибыль» (level 0), затем «· <Юнит>».
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
    """
    Ищет план ЧП ОБЪЕКТА в БДиР.

    Структура БДиР:
      Чистая прибыль (level 0)
      · Latvia       (level 1 — юнит)
      ·· AN14 ...    (level 2 — объект)
    """
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
    """
    Лист «Выполнение годового плана» — столбики:
      • жёлтый  — План (БДиР)
      • зелёный — Факт (ОПиУ)
    """
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

    chart.series[0].graphicalProperties.solidFill = _COLOR_LINE_YEL  # план — жёлтый
    chart.series[1].graphicalProperties.solidFill = _COLOR_NET       # факт — зелёный

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

# ------------------------------------------------------------
# Лист: Таблица «Операционное сальдо» (слайды 2 и 37)
# ------------------------------------------------------------
def _sheet_operational_balance(wb, sheet_title, chart_title, subtitle,
                               rows_data, months):
    """
    Табличный лист «Операционное сальдо».

    rows_data — список словарей:
      {"name": str, "values": {month: value}, "is_total": bool}
    """
    ws = wb.create_sheet(sheet_title)

    _write_header(ws, chart_title, subtitle)

    # --- Шапка таблицы ---
    header_row = 4
    headers = ["Направление"] + [_MONTHS_RU_CAP[m] for m in months]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True)
        cell.border = _CELL_BORDER

    # --- Строки данных ---
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

            # Цвет числа:
            #   положительное (>0) — зелёный
            #   отрицательное (<0) — синий
            #   ноль или пусто — обычный
            if v > 0.005:
                cell.font = Font(color="00B050", bold=True)   # зелёный
            elif v < -0.005:
                cell.font = Font(color="0070C0", bold=True)   # синий

            if row.get("is_total"):
                cell.font = Font(
                    color=("00B050" if v > 0.005 else
                           "0070C0" if v < -0.005 else "000000"),
                    bold=True,
                )

    # --- Немного отрегулируем ширину колонок ---
    ws.column_dimensions["A"].width = 45
    for c in range(2, 2 + len(months)):
        col_letter = ws.cell(row=header_row, column=c).column_letter
        ws.column_dimensions[col_letter].width = 12

def _parse_odds_balance(odds_path, unit_name, months):
    """
    Читает ОДДС-файл и возвращает:
      {
        "rows": [
           {"name": "AN14 Антониас 14 (дом + парковка)", "values": {...}},
           ...
           {"name": "ИТОГО", "values": {...}, "is_total": True},
        ]
      }

    unit_name — "Latvia" или "East-Восток".

    ВАЖНО: для каждого объекта берём строку «·· Сальдо» (level 2),
    которая лежит ВНУТРИ объекта. Например:

      · AN14 Антониас 14 (дом + парковка)   ← level 1 — объект
      ·· Поступления                         ← level 2
      ··· 1.1.1 ...                          ← level 3
      ·· Списания                            ← level 2
      ··· 1.2.1 ...                          ← level 3
      ·· Сальдо                              ← level 2 ← ЭТО то, что нужно
    """
    result = {"rows": []}
    if not odds_path or not os.path.exists(odds_path):
        return result

    wb = openpyxl.load_workbook(odds_path, data_only=True)
    ws = wb.active

    # 1) Шапка
    header_row = None
    for r in range(1, 8):
        v = ws.cell(row=r, column=1).value
        if v and "направление" in str(v).lower():
            header_row = r
            break
    if header_row is None:
        return result

    month_col = _find_month_columns(ws, header_row, months)

    # 2) Ищем строку с названием юнита (level 0)
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

    # 3) Идём по строкам после unit_row.
    #    Для каждого объекта (level 1, НЕ «Сальдо»)
    #    запоминаем имя и ищем дочернюю строку «·· Сальдо» (level 2).
    objects = []
    cur_object_name = None
    cur_object_level = None

    for r in range(unit_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue
        title = _clean(raw)
        lvl = _level(raw)

        # Дошли до следующего юнита — стоп
        if lvl == 0 and title != unit_name:
            break

        # level 1 — либо объект, либо «Сальдо» юнита
        if lvl == 1:
            # Если нашли «Сальдо» уровня 1 — это итог юнита
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

            # Иначе — новый объект
            cur_object_name = title
            cur_object_level = lvl
            continue

        # level 2 — проверяем, не «Сальдо» ли это внутри объекта
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

    # 4) Объекты — перед «ИТОГО»
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

    # Собираем данные для точек безубыточности.
    # Это делается один раз, а используется ниже
    # при построении листов Сл06, Сл44, Сл63.
    breakeven_data = _collect_breakeven(rows, months)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --------------------------------------------------------
    # 1. Юниты — ОПиУ / EBITDA / Доля ФОТ
    # --------------------------------------------------------
    # Соответствие «юнит → номер слайда ОПиУ / EBITDA / Доля ФОТ»
    # Номера слайдов даны по образцу презентации (сентябрь 2026).
    # Для юнитов без отдельного слайда EBITDA/ФОТ — используем
    # базовый стартовый слайд.
    unit_slides = {
        # (ОПиУ, EBITDA, Доля ФОТ)
        "Latvia":       (_UNIT_SLIDE_START["Latvia"],     5,  7),
        "East-Восток":  (_UNIT_SLIDE_START["East-Восток"], 42, 40),
        "Europe":       (_UNIT_SLIDE_START["Europe"],     51, 52),
        "Nomiqa":       (_UNIT_SLIDE_START["Nomiqa"],     73, 0),  # ← Доля ФОТ = 0 (нет отдельного слайда)
        "Unelma":       (_UNIT_SLIDE_START["Unelma"],     68, 0),  # ← Доля ФОТ = 0 (нет отдельного слайда)
        "UK Estate":    (_UNIT_SLIDE_START["UK Estate"],  57, 58),
    }

    for unit_name in _UNITS:
        u = units.get(unit_name)
        if not u:
            continue

        sl_opiu, sl_ebitda, sl_fot = unit_slides.get(
            unit_name, (_UNIT_SLIDE_START.get(unit_name, 1), 0, 0)
        )

        # Определяем цвет линии для EBITDA по юниту
        if unit_name == "Latvia":
            ebitda_color = _COLOR_LINE_RED
            fot_color    = _COLOR_LINE_RED
        elif unit_name == "East-Восток":
            ebitda_color = _COLOR_LINE_YEL
            fot_color    = _COLOR_LINE_RED
        else:
            ebitda_color = _COLOR_LINE_YEL
            fot_color    = _COLOR_LINE_RED

        # ОПиУ юнита (кроме UK Estate — для него ОПиУ Расходы УК
        # создается отдельно в конце, на слайде 56)
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

        # EBITDA юнита
        if sl_ebitda:
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

        # Доля ФОТ юнита
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

        # ----------------------------------------------------
        # План/Факт по чистой прибыли (слайд «Выполнение годового плана»)
        # Диаграмма строится, только если есть файл БДиР.
        # ----------------------------------------------------
        if bdr_path and os.path.exists(bdr_path):
            plan_vals = _parse_bdr_plan(bdr_path, unit_name, months)
            # Проверяем, что план не пустой (иначе смысла в листе нет)
            if any(abs(v) > 0.01 for v in plan_vals.values()):
                # Номер слайда «ПланФакт» — для каждого юнита свой
                # (см. карту слайдов презентации 2026_09):
                #   Latvia       — 8
                #   East-Восток  — 41
                #   Europe       — 53
                #   Nomiqa       — 75
                #   Unelma       — 70
                #   UK Estate    — 59
                sl_plan_fact_map = {
                    "Latvia":       8,
                    "East-Восток":  41,
                    "Europe":       53,
                    "Nomiqa":       75,
                    "Unelma":       70,
                    "UK Estate":    59,
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

                            # Нарастающий итог (слайд 9 для Латвии и т.д.)
                sl_cumulative = {
                    "Latvia":      9,
                    "East-Восток": 43,
                    "Europe":      0,   # ← Нет отдельного слайда с Итогом
                    "Nomiqa":      74,
                    "Unelma":      69,
                    "UK Estate":   60,
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

    # --------------------------------------------------------
    # 1b. Точки безубыточности (слайды 6, 44, 63)
    # --------------------------------------------------------
    # Слайд 6 — ТБУ Латвия
    if "Latvia" in units:
        lat = units["Latvia"]
        _sheet_breakeven(
            wb,
            sheet_title=_sheet_name_with_slide(6, "ТБУ_Latvia"),
            chart_title="Точка безубыточности Латвия",
            subtitle="ТБУ Латвия",
            data_revenue=lat["revenue"],
            data_breakeven=breakeven_data.get("Latvia", {}),
            months=months,
        )

    # Слайд 44 — ТБУ East-Восток
    if "East-Восток" in units:
        east = units["East-Восток"]
        _sheet_breakeven(
            wb,
            sheet_title=_sheet_name_with_slide(44, "ТБУ_East-Восток"),
            chart_title="Точка безубыточности Estate - Восток",
            subtitle="ТБУ Восток (Все юниты, УК)",
            data_revenue=east["revenue"],
            data_breakeven=breakeven_data.get("East-Восток", {}),
            months=months,
        )

    # Слайд 63 — ТБУ Estate (консолидированно: Latvia + East-Восток + Europe)
    estate_rev = {m: 0.0 for m in months}
    estate_be = {m: 0.0 for m in months}
    for u_name in ("Latvia", "East-Восток", "Europe"):
        if u_name in units:
            for m in months:
                estate_rev[m] += units[u_name]["revenue"].get(m, 0.0)
                estate_be[m]  += breakeven_data.get(u_name, {}).get(m, 0.0)

    _sheet_breakeven(
        wb,
        sheet_title=_sheet_name_with_slide(63, "ТБУ_Estate"),
        chart_title="Точка безубыточности Estate",
        subtitle="ТБУ Estate",
        data_revenue=estate_rev,
        data_breakeven=estate_be,
        months=months,
    )

    # --------------------------------------------------------
    # 2. Объекты Латвии
    # --------------------------------------------------------
    if "Latvia" in units:
        latvia = units["Latvia"]
        for obj_full, obj_short, sl_opiu in _LATVIA_OBJECTS:
            obj = latvia["objects"].get(obj_full)
            if not obj:
                continue

            # EBITDA Антонияс — 11 (красный), Чака — 16 (жёлтый),
            # Матиса — 33 (жёлтый), Эспорта — нет (0)
            sl_ebitda_map = {
                "Антонияс": 11,
                "Чака":     16,
                "Матиса":   33,
                "Эспорта":  29,
            }
            ebitda_color_map = {
                "Антонияс": _COLOR_LINE_RED,
                "Чака":     _COLOR_LINE_YEL,
                "Матиса":   _COLOR_LINE_YEL,
                "Эспорта":  _COLOR_LINE_YEL,
            }
            sl_ebitda = sl_ebitda_map.get(obj_short, 0)
            ebitda_color = ebitda_color_map.get(obj_short, _COLOR_LINE_YEL)

            # ОПиУ объекта
            _sheet_revenue_profit(
                wb,
                sheet_title=_sheet_name_with_slide(
                    sl_opiu, f"ОПиУ_{obj_short}"),
                chart_title=f"Выручка & Чистая прибыль {obj_short}, без НДС",
                subtitle=f"ОПиУ {obj_short}",
                color_rev=_COLOR_REVENUE,
                color_net=_COLOR_NET,
                data=obj,
                months=months,
            )

            # EBITDA объекта
            if sl_ebitda:
                _sheet_ebitda(
                    wb,
                    sheet_title=_sheet_name_with_slide(
                        sl_ebitda, f"EBITDA_{obj_short}"),
                    chart_title=f"Операционная рентабельность (EBITDA margin) {obj_short}",
                    subtitle=f"EBITDA margin {obj_short}, %",
                    line_color=ebitda_color,
                    data=obj,
                    months=months,
                )

            # ----------------------------------------------------
            # План/Факт по ЧП объекта Латвии
            # ----------------------------------------------------
            # Для Эспорты в БДиР нет плана — все значения плана будут нулевые.
            # Так и задумано: объект открылся только в августе 2026.
            sl_plan_fact_obj_map = {
                "Антонияс": 12,
                "Чака":     17,
                "Матиса":   34,
                "Эспорта":  30,
            }
            sl_plan_fact_obj = sl_plan_fact_obj_map.get(obj_short, 0)

            if sl_plan_fact_obj:
                # Пытаемся получить план из БДиР.
                # Если файла нет или строка не найдена — вернётся словарь нулей.
                plan_obj = _parse_bdr_plan_object(
                    bdr_path, "Latvia", obj_full, months
                )
                _sheet_plan_fact(
                    wb,
                    sheet_title=_sheet_name_with_slide(
                        sl_plan_fact_obj, f"ПланФакт_{obj_short}"),
                    chart_title=f"Выполнение годового плана по ЧП {obj_short}",
                    subtitle=f"План (БДиР) vs Факт (ОПиУ) — {obj_short}",
                    data_plan=plan_obj,
                    data_fact=obj["net"],
                    months=months,
                )

                            # Нарастающий итог для объекта
                # Для Эспорты нет отдельного слайда с итогом —
                # на слайде 30 уже есть ПланФакт, поэтому ставим 0,
                # чтобы лист «Итог» не создавался.
                sl_cumulative_obj = {
                    "Антонияс": 13,
                    "Чака":     18,
                    "Матиса":   35,
                    "Эспорта":  0,
                }.get(obj_short)

                if sl_cumulative_obj:
                    _sheet_cumulative_plan_fact(
                        wb,
                        sheet_title=_sheet_name_with_slide(
                            sl_cumulative_obj, f"Итог_{obj_short}"),
                        chart_title=f"Выполнение годового плана по чистой прибыли {obj_short}",
                        data_plan=plan_obj,
                        data_fact=obj["net"],
                        months=months,
                        plan_label=f"ЧП {obj_short} план (с НДС)",
                        fact_label=f"ЧП {obj_short} факт (с НДС)",
                    )

        # ----------------------------------------------------
        # 3. Блок «Коммерческие»
        # ----------------------------------------------------
        commercial = latvia["objects"].get("Коммерческие помещения LV")
        if commercial:
            _sheet_revenue_profit(
                wb,
                sheet_title=_sheet_name_with_slide(
                    20, "ОПиУ_Коммерческие"),
                chart_title="Выручка & Чистая прибыль Коммерческие Латвия, без НДС",
                subtitle="ОПиУ Коммерческие",
                color_rev=_COLOR_REVENUE,
                color_net=_COLOR_NET,
                data=commercial,
                months=months,
            )
            _sheet_ebitda(
                wb,
                sheet_title=_sheet_name_with_slide(
                    21, "EBITDA_Коммерческие"),
                chart_title="Операционная рентабельность (EBITDA margin) Коммерческие помещения",
                subtitle="EBITDA margin Коммерческие, %",
                line_color=_COLOR_LINE_YEL,
                data=commercial,
                months=months,
            )

            # План/Факт по Коммерческим (слайд 22)
            plan_comm = _parse_bdr_plan_object(
                bdr_path, "Latvia", "Коммерческие помещения LV", months
            )
            # В БДиР «Коммерческие» — это сумма по объектам;
            # чтобы получить её, суммируем по каждому коммерческому объекту.
            if not any(abs(v) > 0.01 for v in plan_comm.values()):
                # Пробуем сумму по каждому объекту из _LATVIA_COMMERCIAL
                for obj_comm in _LATVIA_COMMERCIAL:
                    part = _parse_bdr_plan_object(
                        bdr_path, "Latvia", obj_comm, months
                    )
                    for m in months:
                        plan_comm[m] += part.get(m, 0.0)

            _sheet_plan_fact(
                wb,
                sheet_title=_sheet_name_with_slide(
                    22, "ПланФакт_Коммерческие"),
                chart_title="Выполнение годового плана по ЧП Коммерческие Латвия",
                subtitle="План (БДиР) vs Факт (ОПиУ) — Коммерческие",
                data_plan=plan_comm,
                data_fact=commercial["net"],
                months=months,
            )

                        # Нарастающий итог для Коммерческих
            _sheet_cumulative_plan_fact(
                wb,
                sheet_title=_sheet_name_with_slide(23, "Итог_Коммерческие"),
                chart_title="Выполнение годового плана по чистой прибыли Коммерческие Латвия",
                data_plan=plan_comm,
                data_fact=commercial["net"],
                months=months,
                plan_label="ЧП Коммерческие план (с НДС)",
                fact_label="ЧП Коммерческие факт (с НДС)",
            )

    # --------------------------------------------------------
    # 3b. Табличные слайды — «Операционное сальдо»
    #     Слайд 2  — Латвия
    #     Слайд 37 — East-Восток
    # --------------------------------------------------------
    # Путь к ОДДС — единый файл в TEMP_FORECAST_DIR
    # Путь к ОДДС — единый файл в TEMP_FORECAST_DIR.
    # Формат имени: «ОДДС 01.01.2026-30.09.2026.xlsx» (с последним днём месяца).
    # Определяем последний день отчётного месяца через calendar.monthrange.
    import calendar as _cal
    _last_day = _cal.monthrange(year, month)[1]
    odds_path = config.TEMP_FORECAST_DIR / f"ОДДС 01.01.{year}-{_last_day:02d}.{month:02d}.{year}.xlsx"

    if odds_path.exists():
        # --- Латвия (слайд 2) ---
        lat_bal = _parse_odds_balance(odds_path, "Latvia", months)
        if lat_bal["rows"]:
            _sheet_operational_balance(
                wb,
                sheet_title=_sheet_name_with_slide(2, "Сальдо_Latvia"),
                chart_title="Операционное сальдо Латвия (без НДС)",
                subtitle=f"Факт за {_MONTHS_RU_LOWER.get(month, '')} {year}",
                rows_data=lat_bal["rows"],
                months=months,
            )

        # --- East-Восток (слайд 37) ---
        east_bal = _parse_odds_balance(odds_path, "East-Восток", months)
        if east_bal["rows"]:
            _sheet_operational_balance(
                wb,
                sheet_title=_sheet_name_with_slide(37, "Сальдо_East-Восток"),
                chart_title="Операционное сальдо East-Восток",
                subtitle=f"Факт за {_MONTHS_RU_LOWER.get(month, '')} {year}",
                rows_data=east_bal["rows"],
                months=months,
            )

    # --------------------------------------------------------
    # 4. Специальные листы: Расходы УК (слайд 56)
    # --------------------------------------------------------
    if "UK Estate" in units:
        uk = units["UK Estate"]
        _sheet_revenue_profit(
            wb,
            sheet_title=_sheet_name_with_slide(56, "ОПиУ_Расходы_УК"),
            chart_title="Расходы УК R1",
            subtitle="ОПиУ Расходы УК",
            color_rev=_COLOR_REVENUE,
            color_net=_COLOR_NET,
            data=uk,
            months=months,
        )

    # --------------------------------------------------------
    # СОРТИРОВКА ЛИСТОВ ПО НОМЕРАМ СЛАЙДОВ ПРЕЗЕНТАЦИИ
    # --------------------------------------------------------
    # В процессе создания листы добавляются в порядке вызовов
    # функций, и он не совпадает с порядком слайдов в презентации
    # (например, для East-Восток: 38, 42, 40, 41).
    # Пересортировываем вкладки по числу после «Сл».
    _sort_sheets_by_slide_number(wb)

    wb.save(output_path)
    print(f"[OK] Диаграммы сохранены: {output_path}")
    return output_path


def _sort_sheets_by_slide_number(wb):
    """
    Пересортировывает листы книги по номеру слайда в имени листа.

    Использует wb.move_sheet() — официальный API openpyxl.
    Приём `wb._sheets = ...` не работает в свежих версиях openpyxl,
    потому что листы хранятся как связный список внутри workbook.
    """
    import re as _re

    def _slide_key(ws):
        m = _re.search(r"Сл(\d+)", ws.title)
        if m:
            return (0, int(m.group(1)), ws.title)
        # Листы без номера слайда — в конец
        return (1, 10**9, ws.title)

    try:
        desired = sorted(wb.worksheets, key=_slide_key)

        for target_index, ws in enumerate(desired):
            cur_index = wb.worksheets.index(ws)
            if cur_index != target_index:
                # Сдвигаем лист на позицию target_index
                offset = target_index - cur_index
                wb.move_sheet(ws, offset=offset)

        print(f"[presentation_builder] листы отсортированы, порядок: "
              f"{[ws.title[:8] for ws in wb.worksheets]}")
    except Exception as e:
        print(f"[presentation_builder] не удалось отсортировать листы: {e}")

# ============================================================
# ЛИСТ «НАРАСТАЮЩИМ ИТОГОМ» (слайды 9, 13, 18, 23, 35, ...)
# ============================================================

# Цвета месяцев (приближены к PowerPoint)
_MONTH_FILL = {
    1: "4472C4",   # январь  — синий
    2: "8FAADC",   # февраль — голубой
    3: "FFC000",   # март    — жёлтый
    4: "548235",   # апрель  — зелёный
    5: "ED7D31",   # май     — оранжевый
    6: "9DC3E6",   # июнь    — светло-голубой
    7: "7030A0",   # июль    — фиолетовый
    8: "BFB100",   # август  — оливковый
    9: "2E4C99",   # сентябрь — тёмно-синий
    10: "C55A11",
    11: "375623",
    12: "203864",
}

_TOTAL_BAR_COLUMNS = 50   # сколько узких столбцов используем для полос


def _sheet_cumulative_plan_fact(wb, sheet_title, chart_title,
                                data_plan, data_fact, months,
                                plan_label, fact_label):
    """
    Лист «Нарастающим итогом» — горизонтальная stacked-диаграмма.
    План сверху, Факт снизу.
    """
    ws = wb.create_sheet(sheet_title)

    # ---- 1. Данные для диаграммы ----
    header_row = 2   # шапка (названия месяцев)
    fact_row   = 3   # Факт (окажется ВНИЗУ диаграммы)
    plan_row   = 4   # План (окажется СВЕРХУ диаграммы)

    ws.cell(row=header_row, column=1, value="Показатель").font = Font(bold=True)
    for i, m in enumerate(months, start=2):
        ws.cell(row=header_row, column=i,
                value=_MONTHS_RU_LOWER[m].capitalize()).font = Font(bold=True)

    # Короткие подписи категорий — «План» и «Факт».
    # Их Excel возьмёт как подписи оси Y.
    ws.cell(row=fact_row, column=1, value="Факт").font = Font(bold=True)
    ws.cell(row=plan_row, column=1, value="План").font = Font(bold=True)

    # Заполняем данными. Шрифт делаем БЕЛЫМ и мелким — на белом фоне
    # числа визуально не будут видны, но Excel их использует для диаграммы.
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

    # ---- 2. Диаграмма ----
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

    # Цвета серий — по месяцам
    for idx, series in enumerate(chart.series):
        if idx < len(months):
            m = months[idx]
            color = _MONTH_FILL.get(m, "808080")
            series.graphicalProperties.solidFill = color
            series.graphicalProperties.line.solidFill = "FFFFFF"
            series.graphicalProperties.line.width = 10000

    chart.width = 32
    chart.height = 11

    # Заголовок диаграммы (как в PowerPoint: сверху по центру).
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

    # Включаем оси и задаём настройки, чтобы Excel гарантированно
    # нарисовал ось X и подписи категорий («План» / «Факт»).
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.numFmt = '#,##0'
    chart.x_axis.majorTickMark = "out"
    chart.x_axis.tickLblPos = "nextTo"
    chart.x_axis.title = "Периоды"

    chart.y_axis.majorTickMark = "out"
    chart.y_axis.tickLblPos = "nextTo"

    # Формат линий — убираем серую сетку, чтобы было как в PowerPoint
    chart.x_axis.majorGridlines = None
    chart.y_axis.majorGridlines = None

    # ВАЖНО: используем абсолютный якорь вместо "A6".
    # Строковый якорь иногда «съезжает» в скрытые строки
    # и openpyxl на этапе save() выбрасывает диаграмму.
    from openpyxl.utils import column_index_from_string
    anchor_col_idx = column_index_from_string("A") - 1   # 0-based
    anchor_row_idx = 6 - 1                              # 0-based, строка 6
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

    # ---- 3. Блок итогов справа ----
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

    # ---- 4. Строки 2-4 оставляем нормальной высоты ----
    # Excel использует их как источник подписей категорий («План» / «Факт»)
    # и заголовков серий (месяцы). Если их сжать, Excel теряет подписи
    # и ось X. Числа в этих строках уже сделаны белыми — визуально не видны.
    #
    # Ничего не делаем — просто оставляем строки как есть.
    pass

    ws.sheet_view.showGridLines = False
    return ws
