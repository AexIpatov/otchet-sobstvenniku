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
    ("EKS_Esporta iela 12-113",            "Эспорта",  33),
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

    # 4) агрегаты юнита = сумма по объектам
    for u_name, u in units.items():
        for m in months:
            u["revenue"][m] = sum(o["revenue"][m] for o in u["objects"].values())
            u["net"][m]     = sum(o["net"][m]     for o in u["objects"].values())
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


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def _sheet_name_with_slide(slide_no, base_name):
    """Формирует имя листа вида 'Сл04_ОПиУ_Латвия'."""
    return _safe_sheet(f"Сл{slide_no:02d}_{base_name}")


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
        "Nomiqa":       (_UNIT_SLIDE_START["Nomiqa"],     73, 74),
        "Unelma":       (_UNIT_SLIDE_START["Unelma"],     68, 69),
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

        # ОПиУ юнита
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
                "Эспорта":  34,
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

    wb.save(output_path)
    print(f"[OK] Диаграммы сохранены: {output_path}")
    return output_path
````---

## Что изменилось

### 1. Имена листов с номерами слайдов

Теперь каждый лист называется, например:
- `Сл04_ОПиУ_Latvia`
- `Сл05_EBITDA_Latvia`
- `Сл07_Доля_ФОТ_Latvia`
- `Сл10_ОПиУ_Антонияс`
- `Сл11_EBITDA_Антонияс`
- `Сл20_ОПиУ_Коммерческие`
- `Сл21_EBITDA_Коммерческие`
- `Сл56_ОПиУ_Расходы_УК`

### 2. Цвета по образцу презентации

- **ОПиУ юнитов и объектов**: выручка — **жёлтый** `FFC000`, ЧП — **зелёный** `70AD47`.
- **EBITDA Латвия**: **красный** `C00000`.
- **EBITDA Антонияс**: **красный** `C00000`.
- **EBITDA Чака / Матиса / Эспорта / Коммерческие**: **жёлтый** `FFC000`.
- **Доля ФОТ**: **красный** `C00000` (как в презентации стр. 7).

### 3. Цвета теперь точно соответствуют слайдам:

| Слайд | Что | Цвет линии/столбиков |
|-------|-----|----------------------|
| 4 | ОПиУ Латвия | жёлтый + зелёный |
| 5 | EBITDA Латвия | красный |
| 7 | Доля ФОТ Латвия | красный |
| 10 | ОПиУ Антонияс | жёлтый + зелёный |
| 11 | EBITDA Антонияс | красный |
| 15 | ОПиУ Чака | жёлтый + зелёный |
| 16 | EBITDA Чака | жёлтый |
| 20 | ОПиУ Коммерческие | жёлтый + зелёный |
| 21 | EBITDA Коммерческие | жёлтый |
| 32 | ОПиУ Матиса | жёлтый + зелёный |
| 33 | EBITDA Матиса | жёлтый |
| 38 | ОПиУ East-Восток | жёлтый + зелёный |
| 40 | Доля ФОТ East-Восток | красный |
| 42 | EBITDA East-Восток | жёлтый |
| 49 | ОПиУ Европа | жёлтый + зелёный |
| 56 | ОПиУ Расходы УК | жёлтый + зелёный |
| 67 | ОПиУ Унелма | жёлтый + зелёный |
| 71 | ОПиУ Nomiqa | жёлтый + зелёный |

---

## Git-команды

```bash
cd "C:\Users\Александр\Отчет для собственника - cloud"

git add src/presentation_builder.py
git commit -m "feat(presentation): номера слайдов в именах листов + цвета по образцу презентации"
git push origin main
```

---

## Что проверить после сборки

1. **Имена листов** в Excel должны начинаться с `Сл04_`, `Сл05_`, `Сл10_` и т.д.
2. **Лист `Сл04_ОПиУ_Latvia`**: столбики жёлтые + зелёные.
3. **Лист `Сл05_EBITDA_Latvia`**: линия **красная**.
4. **Лист `Сл07_Доля_ФОТ_Latvia`**: линия **красная** (не синяя!).
5. **Лист `Сл11_EBITDA_Антонияс`**: линия **красная**.
6. **Лист `Сл16_EBITDA_Чака`**: линия **жёлтая**.
7. **Лист `Сл20_ОПиУ_Коммерческие`**: столбики жёлтые + зелёные.

Если какие-то цвета не совпадут с презентацией — напишите, какой именно лист и какой цвет должен быть, я поправлю одну строку.