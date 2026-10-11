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
    "BRN_Brunieku",
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

def _collect_unit_expenses(rows, unit_name, months):
    """
    Считает сумму расходов юнита по модулю во всех трёх секциях.

    Использует простой и надёжный приём:
      • идём по всем строкам,
      • запоминаем, внутри какой секции и какого юнита мы находимся,
      • суммируем ТОЛЬКО статьи, которые лежат на 1 уровень ниже юнита,
      • как только встретили другой юнит (в той же секции) — прекращаем.
    """
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
        # level 0 — переключение секции → сбрасываем всё
        if lvl == 0:
            cur_section = title
            cur_unit_level = None
            inside_target = False
            continue

        if cur_section not in sections:
            continue

        # Вошли в наш юнит
        if title == unit_name:
            cur_unit_level = lvl
            inside_target = True
            continue

        if not inside_target:
            continue

        # Встретили строку того же или меньшего уровня,
        # но с другим именем → значит, наш юнит закончился.
        if lvl <= cur_unit_level:
            inside_target = False
            continue

        # Наш юнит: суммируем только его "детей" ровно на 1 уровень ниже
        # (чтобы не задвоить, если внутри есть подстатьи).
        if lvl == cur_unit_level + 1:
            for m in months:
                v = values.get(m, 0.0) or 0.0
                result[m] += abs(v)

    return result

def _parse_investments_all(ref_path):
    """
    Читает файл «Вложенные средства на объекты.xlsx» и возвращает
    {имя_объекта: сумма_вложений}.

    Структура файла:
      Строка 1: «Вложенные средства на объекты»
      Строка 2: «Наименование объекта» | «Сумма ...»
      Далее: строки с объектами и суммами.
      Встречаются заголовки секций без чисел: «Estate LV», «Estate East:», «Estate EU».
      Также встречаются формулы типа «=SUM(...)» в столбце B — их игнорируем.

    Имена объектов включают суффикс типа «[Latvia]», «[East-Восток]», «[Europe]» —
    сохраняем их как есть.
    """
    result = {}
    if not ref_path or not os.path.exists(ref_path):
        return result

    wb = openpyxl.load_workbook(ref_path, data_only=True)
    ws = wb.active

    for r in range(1, ws.max_row + 1):
        name = _clean(ws.cell(row=r, column=1).value)
        if not name:
            continue

        # Пропускаем заголовок
        if "наименование" in name.lower():
            continue

        # Читаем сумму из столбца B
        v = ws.cell(row=r, column=2).value

        # Если это формула (строка начинается с «=»), при data_only=True
        # она уже вычислена. Но если формула не вычислена — value будет «=SUM(...)» —
        # тогда пропускаем строку.
        if isinstance(v, str) and v.startswith("="):
            continue

        amount = _num(ws, r, 2)

        # Пустые суммы (None) — это заголовки секций, пропускаем
        if amount == 0.0 and v in (None, "", 0):
            # Но если в явном виде 0 (UK_Latvia, ML2, UK_EU) — сохраняем
            if v == 0:
                result[name] = 0.0
            continue

        result[name] = float(amount)

    return result


def _sum_investments(investments_all, object_names):
    """
    Возвращает сумму вложений по списку объектов.
    object_names — список имён объектов (например,
      ["AN14 Антониас 14 (дом + парковка) [Latvia]"]).
    """
    total = 0.0
    for name in object_names:
        total += investments_all.get(name, 0.0)
    return total

def _parse_debts(debts_path, months):
    """
    Читает файл «таблица Долги и Численность сотрудников.xlsx».

    Структура (актуальный файл):
      • Строка 1: «Долги по объектам аренды» (заголовок)
      • Строка 2: «Объект» | 2026-01-01 | 2026-02-01 | … (шапка с датами)
      • Строки 3..N: имя объекта | числа по месяцам
      • Пустые строки
      • Строка с «Численность сотрудников Estate» (заголовок)
      • Пустая строка с датами в B..M
      • Строка с «Численность сотрудников Estate» | числа

    Возвращает:
      {
        "debts": {object_name: {month: value}},
        "headcount": {month: value},
      }
    """
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

    # --- 1. Ищем шапку «Объект» ---
    # В файле есть строка 1 «Долги по объектам аренды» — она тоже
    # содержит слово «объект». Поэтому ищем не «объект» где угодно,
    # а именно ячейку A, в которой написано ровно «Объект»
    # (без других слов), И при этом в B..M есть даты.
    header_row = None
    for r in range(1, min(30, ws.max_row + 1)):
        v = ws.cell(row=r, column=1).value
        if not v:
            continue
        sv = str(v).strip().lower()
        if sv != "объект":
            continue
        # Проверяем, что в этой же строке в B..M есть даты
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
        # Диагностика — печатаем первые 5 ячеек столбца A
        for rr in range(1, min(6, ws.max_row + 1)):
            print(f"[_parse_debts]   A{rr} = "
                  f"{ws.cell(row=rr, column=1).value!r}")
        return result
    print(f"[_parse_debts] шапка «Объект» с датами — строка {header_row}")

    # --- 2. Собираем {month: col} из дат в шапке ---
    # ВАЖНО: в файле есть даты за 2026 и за 2027. Нам нужны ТОЛЬКО
    # даты отчётного года. Иначе «январь 2027» перезапишет «январь 2026».
    month_col = {}
    for c in range(2, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        # Вариант 1: ячейка — настоящий datetime
        if isinstance(v, (_dt.datetime, _dt.date)):
            if v.year == config.REPORT_YEAR and v.month in months:
                month_col[v.month] = c
            continue
        # Вариант 2: ячейка — текст вида «Январь 2026» или «январь 2026»
        sv = str(v).strip().lower()
        for m in months:
            if m in month_col:
                continue
            month_name = _MONTHS_RU_LOWER[m]  # например, "январь"
            if month_name in sv and str(config.REPORT_YEAR) in sv:
                month_col[m] = c
                break
    print(f"[_parse_debts] колонок с датами в шапке (год "
          f"{config.REPORT_YEAR}): {len(month_col)}")
    print(f"[_parse_debts]   month_col = {month_col}")

    if not month_col:
        print("[_parse_debts] в шапке нет дат нужных месяцев")
        return result

    # --- 3. Идём по строкам данных до строки «Численность...» ---
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

    # --- 4. Парсим численность ---
    if headcount_header_row is not None:
        print(f"[_parse_debts] блок «Численность» — строка {headcount_header_row}")

        # Идём вниз от headcount_header_row, ищем:
        #   • строку с датами (числа datetime в B..M)
        #   • строку с числами (значения в B..M)
        #
        # Структура блока: заголовок → строка с датами → строка с числами.
        # Между ними могут быть пустые ячейки в столбце A.

        hc_dates_row = None
        hc_values_row = None

        for rr in range(headcount_header_row + 1,
                        min(headcount_header_row + 10, ws.max_row + 1)):
            a_val = ws.cell(row=rr, column=1).value
            # Считаем строку «датой-строкой», если в B..M есть datetime
            has_dates = any(
                isinstance(ws.cell(row=rr, column=c).value, (_dt.datetime, _dt.date))
                for c in range(2, min(ws.max_column, 25) + 1)
            )
            # Считаем строку «числовой», если в B..M есть int/float
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

        # Если нашли строку с числами — берём колонки месяцев из строки дат,
        # либо (если дат нет) — из шапки долгов.
        if hc_values_row is not None:
            hc_month_col = {}
            if hc_dates_row is not None:
                for c in range(2, ws.max_column + 1):
                    v = ws.cell(row=hc_dates_row, column=c).value
                    # Вариант 1: datetime
                    if isinstance(v, (_dt.datetime, _dt.date)):
                        if v.year == config.REPORT_YEAR and v.month in months:
                            hc_month_col[v.month] = c
                        continue
                    # Вариант 2: текст «Январь 2026»
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
    """
    Лист «Долги» — столбики (для слайдов 14, 19).
    objects — список: [{"name": ..., "color": ..., "values": {m: v}}]
    """
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
    """
    Лист «Долги» — линии (для слайда 24).
    objects — список: [{"name": ..., "color": ..., "values": {m: v}}]
    """
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
    """
    Лист «Анализ ОПиУ» (слайд 64) — 3 линии:
      • Численность
      • Выручка на 1 сотрудника
      • Средняя ЗП на 1 сотрудника
    """
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
    """
    series_data — список: [{"name": ..., "color": ..., "values": {m: roi}}]
    """
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
# Лист: EBITDA margin — несколько линий (для слайда 05 Latvia)
# ------------------------------------------------------------
def _sheet_ebitda_multi(wb, sheet_title, chart_title, subtitle,
                        series_data, months):
    """
    Лист с несколькими линиями рентабельности (EBITDA margin).
    series_data — список: [{"name": ..., "color": ..., "values": {m: v}}]
    """
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

def _sheet_dds(wb, sheet_title, chart_title, subtitle,
               data_inflow, data_outflow, months):
    """
    Лист «Анализ ДДС» — столбики «Поступления» (зелёные) и
    «Выбытия» (красные). Данные берём из ОДДС.
    """
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

    # Поступления — зелёные, Выбытия — красные
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

# ------------------------------------------------------------
# Лист: «Выручка по юнитам» (слайд 65) — 4 линии
# ------------------------------------------------------------
def _sheet_revenue_by_units(wb, sheet_title, chart_title, subtitle,
                            series_data, months):
    """
    series_data — список словарей:
      [{"name": "Latvia", "color": "4472C4", "values": {m: v}},
       {"name": "East-Восток", "color": "C00000", "values": {m: v}}, ...]
    """
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

# ------------------------------------------------------------
# Лист: «EBITDA margin юнитов» (слайд 69) — 3 линии
# ------------------------------------------------------------
def _sheet_ebitda_units(wb, sheet_title, chart_title, subtitle,
                        series_data, months):
    """
    series_data — список: [{"name": ..., "color": ..., "values": {m: v}}]
    """
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

# ------------------------------------------------------------
# Лист: «Доля ФОТ в полной выручке» (слайд 66) — 3 линии
# ------------------------------------------------------------
def _sheet_fot_share_full(wb, sheet_title, chart_title, subtitle,
                          series_data, months):
    """
    series_data — список: [{"name": ..., "color": ..., "values": {m: v}}]
    """
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


# ------------------------------------------------------------
# Лист: «EBITDA margin общий» (слайд 67) — 1 линия
# ------------------------------------------------------------
def _sheet_ebitda_common(wb, sheet_title, chart_title, subtitle,
                         data_values, months):
    """data_values = {month: margin}"""
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


# ------------------------------------------------------------
# Лист: «Динамика валовой прибыли» (слайд 68) — stacked bar, 3 серии
# ------------------------------------------------------------
def _sheet_profit_dynamics(wb, sheet_title, chart_title, subtitle,
                           series_data, months):
    """
    series_data — список: [{"name": ..., "color": ..., "values": {m: v}}]
    Серии складываются друг на друга (stacked).
    """
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

# ------------------------------------------------------------
# Лист: «Возмещение коммунальных услуг» (слайды 25, 27, 29)
# ------------------------------------------------------------
def _find_ku_sheet(wb_ku, month_name):
    """
    Ищет в книге лист, на котором есть таблица возмещения КУ.

    Признак таблицы КУ:
      • в столбце A есть ячейка «Объекты»,
      • ниже (в пределах 30 строк) есть строка «Итого»,
      • в первой строке данных столбец E (Мусор) содержит
        ПОЛОЖИТЕЛЬНОЕ число (в блоке ОДДС там отрицательные).

    Если лист не найден — возвращает первый лист, где есть «Объекты».
    Если и такого нет — активный лист.
    """
    def _looks_like_ku(ws):
        header_row = None
        for r in range(1, 20):
            v = _clean(ws.cell(row=r, column=1).value)
            if v and "объекты" in v.lower():
                header_row = r
                break
        if header_row is None:
            return False
        # Ищем «Итого» в пределах 30 строк после шапки
        has_total = False
        for r in range(header_row + 1,
                       min(header_row + 30, ws.max_row + 1)):
            v = _clean(ws.cell(row=r, column=1).value)
            if v and "итого" in v.lower():
                has_total = True
                break
        if not has_total:
            return False
        # Проверяем, что первая строка данных — положительная (расходы),
        # а не отрицательная (списания из ОДДС).
        for r in range(header_row + 1,
                       min(header_row + 5, ws.max_row + 1)):
            e_val = _num(ws, r, 5)
            if e_val > 0:
                return True
            # Если встретили отрицательное — это блок ОДДС
            if e_val < 0:
                return False
        return False

    # 1) Перебираем все листы — ищем таблицу КУ
    for ws in wb_ku.worksheets:
        if _looks_like_ku(ws):
            return ws

    # 2) Fallback: лист с именем month_name
    if month_name in wb_ku.sheetnames:
        return wb_ku[month_name]

    # 3) Fallback: первый лист, где есть «Объекты»
    for ws in wb_ku.worksheets:
        for r in range(1, 20):
            v = _clean(ws.cell(row=r, column=1).value)
            if v and "объекты" in v.lower():
                return ws

    # 4) Fallback: активный лист
    return wb_ku.active

def _find_odds_month_col(ws_ku, month_name):
    """
    Ищет в ОДДС-файле колонку с нужным месяцем.
    Шапка: строка с «Направление» в столбце A.
    В строке шапки ищем ячейку, где в тексте есть month_name
    (например, «Сентябрь 2026» → колонка J).
    Возвращает индекс колонки (1-based) или None.
    """
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
    """
    Создаёт лист с таблицей возмещения коммунальных услуг из ОДДС.

    ku_path — путь к файлу ОДДС («ОДДС … с НДС.xlsx» или
              «ОДДС … без НДС.xlsx»).
    month_name — название месяца (например, «Сентябрь»).
    subtract_object — если задано (например, "AC89 Чака"), то
                      строки этого объекта исключаются из вывода,
                      а строка «Итого» считается как
                      Итого(все) − Объект(subtract_object).
                      Используется для слайда 27.

    Структура ОДДС:
      • level 0 — «Latvia» (направление);
      • level 1 — объект («· AN14 Антониас 14 (дом + парковка)»);
      • level 2 — «Списания» / «Сальдо»;
      • level 3 — «1.2.10 Коммунальные платежи»;
      • level 4 — «1.2.10.1 Мусор», «1.2.10.2  Газ», «1.2.10.3 Вода»,
                  «1.2.10.4 Отопление», «1.2.10.5 Электричество»,
                  «1.2.10.6 Коммунальные УК дома».
    Данные за месяц — в колонке нужного месяца. Значения со знаком
    минус (расходы), берём abs(...).

    Столбцы итогового листа:
      A — Объект
      B, C, D — пусто (в ОДДС этих данных нет)
      E — Мусор
      F — Газ
      G — Вода
      H — Отопление
      I — Электричество
      J — Коммунальные УК дома
      K — Всего расходов на коммунальные услуги (=SUM(E:J))
      L, M, N — пусто
    """
    ws = wb.create_sheet(sheet_title)
    _write_header(ws, chart_title, subtitle)

    # --- Шапка таблицы: три уровня, как в файле «Возмещение КУ» ---
    header_row_1 = 3   # группа
    header_row_2 = 4   # «Объекты»
    header_row_3 = 5   # номера 1..14
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

    # --- Открываем ОДДС ---
    if not ku_path or not os.path.exists(ku_path):
        print(f"[presentation_builder] не найден файл КУ: {ku_path}")
        return ws

    wb_ku = openpyxl.load_workbook(ku_path, data_only=True)
    ws_ku = wb_ku.active

    # --- Колонка месяца ---
    month_col = _find_odds_month_col(ws_ku, month_name)
    if month_col is None:
        print(f"[presentation_builder] в файле {ku_path} "
              f"не найдена колонка для месяца «{month_name}»")
        return ws

    # --- Собираем данные из ОДДС ---
    # Структура: смотрим строки level-1 (объекты) и их дочерние
    # level-4 строки статей «1.2.10.N …».
    #
    # Возвращаем: {odd_name: {"musor": v, "gaz": v, "voda": v,
    #                          "otopl": v, "elektr": v, "uk_doma": v}}
    #
    # Категорию определяем по тексту после «1.2.10.N»:
    #   1.2.10.1 → Мусор
    #   1.2.10.2 → Газ
    #   1.2.10.3 → Вода
    #   1.2.10.4 → Отопление
    #   1.2.10.5 → Электричество
    #   1.2.10.6 → Коммунальные УК дома

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

        # level 0 — новое направление (Latvia / East-Восток / …) — сбрасываем
        if lvl == 0:
            cur_obj = None
            cur_obj_level = None
            continue

        # level 1 — объект
        if lvl == 1:
            cur_obj = title
            cur_obj_level = lvl
            if cur_obj not in data_by_obj:
                data_by_obj[cur_obj] = {
                    "musor": 0.0, "gaz": 0.0, "voda": 0.0,
                    "otopl": 0.0, "elektr": 0.0, "uk_doma": 0.0,
                }
            continue

        # внутри объекта — ищем статьи 1.2.10.N
        if cur_obj is None:
            continue
        # встретили другой уровень ≤ cur_obj_level — объект закончился
        if lvl <= cur_obj_level:
            cur_obj = None
            cur_obj_level = None
            continue

        # обрабатываем только строки статей 1.2.10.N
        cat = _category(title)
        if cat is None:
            continue
        # не суммируем «···· 1.2.10 Коммунальные платежи» (агрегат)
        # — там нет «.N» в конце, но проверка _category это отсекает,
        # потому что в агрегате нет «1.2.10.1…6»
        val = _num(ws_ku, r, month_col)
        data_by_obj[cur_obj][cat] += abs(val)

    # --- Записываем строки в порядке KU_OBJECTS ---
    row_out = header_row + 1
    data_rows = []

    for ku_name in config.KU_OBJECTS:
        odd_name = config.KU_TO_ODDS_NAME.get(ku_name, ku_name)

        # если это объект, который надо вычесть — пропускаем
        if subtract_object and ku_name == subtract_object:
            continue

        d = data_by_obj.get(odd_name)
        if d is None:
            # Объект не найден в ОДДС — пишем пустую строку
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

        # K — Всего = сумма E:J (пишем формулу)
        ws.cell(row=row_out, column=11,
                value=f"=SUM(E{row_out}:J{row_out})").border = _CELL_BORDER
        ws.cell(row=row_out, column=11).number_format = '#,##0'

        # L, M, N — пусто
        ws.cell(row=row_out, column=12).border = _CELL_BORDER
        ws.cell(row=row_out, column=13).border = _CELL_BORDER
        ws.cell(row=row_out, column=14).border = _CELL_BORDER

        # формат чисел для E:J
        for c in range(5, 11):
            ws.cell(row=row_out, column=c).number_format = '#,##0'

        data_rows.append(row_out)
        row_out += 1

    # --- Строка «Итого» ---
    total_row = row_out
    ws.cell(row=total_row, column=1, value="Итого").font = Font(bold=True)
    ws.cell(row=total_row, column=1).border = _CELL_BORDER
    for c in range(2, 15):
        ws.cell(row=total_row, column=c).border = _CELL_BORDER

    if data_rows:
        first_data = data_rows[0]
        last_data = data_rows[-1]
        for c in range(5, 12):  # столбцы E..K
            col_letter = get_column_letter(c)
            formula = f"=SUM({col_letter}{first_data}:{col_letter}{last_data})"
            cell = ws.cell(row=total_row, column=c, value=formula)
            cell.font = Font(bold=True)
            cell.number_format = '#,##0'
            cell.border = _CELL_BORDER
    else:
        for c in range(5, 12):
            ws.cell(row=total_row, column=c, value=0).border = _CELL_BORDER

    # --- Ширина колонок ---
    ws.column_dimensions["A"].width = 30
    for c in range(2, 15):
        ws.column_dimensions[get_column_letter(c)].width = 14

    return ws

def _collect_dds(odds_path, unit_name, months):
    """
    Читает строки-агрегаты «Поступления» и «Списания» юнита из ОДДС.
    В ОДДС они идут БЕЗ имени (столбец A пустой) сразу под именем юнита.

    Возвращает:
       {"inflow":  {month: value}, "outflow": {month: value}}
       где outflow — по модулю.
    """
    result = {
        "inflow":  {m: 0.0 for m in months},
        "outflow": {m: 0.0 for m in months},
    }
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

    # 2) Ищем строку юнита (level 0)
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

    # 3) Определяем, где лежат Поступления, а где Списания.
    #
    # В файле ОДДС возможны ДВА варианта структуры:
    #
    # ВАРИАНТ A (актуальный):
    #   строка N     — юнит «Latvia» — в этой же строке АГРЕГАТ ПОСТУПЛЕНИЙ
    #                  (положительные числа, зелёные)
    #   строка N+1   — пустая в столбце A — АГРЕГАТ СПИСАНИЙ
    #                  (отрицательные числа)
    #
    # ВАРИАНТ B (устаревший):
    #   строка N     — юнит «Latvia» (0)
    #   строка N+1   — АГРЕГАТ СПИСАНИЙ (отрицательные)
    #   строка N+2   — АГРЕГАТ ПОСТУПЛЕНИЙ (положительные)
    #
    # Определяем вариант по знаку значения в строке юнита:
    #   если там есть положительные числа → это ПОСТУПЛЕНИЯ (вариант A);
    #   иначе → вариант B.

    # Сначала пробуем прочитать саму строку юнита.
    unit_values = {}
    for m, col in month_col.items():
        unit_values[m] = _num(ws, unit_row, col) or 0.0

    # Есть ли в строке юнита положительные значения?
    has_positive_in_unit_row = any(v > 0 for v in unit_values.values())

    if has_positive_in_unit_row:
        # ВАРИАНТ A: строка юнита = Поступления, следующая строка = Списания.
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
        # ВАРИАНТ B: строка N+1 = Списания, строка N+2 = Поступления.
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
        "Nomiqa":       (_UNIT_SLIDE_START["Nomiqa"],     73, 0),
        "Unelma":       (_UNIT_SLIDE_START["Unelma"],     68, 0),
        # UK Estate генерируется отдельно на слайдах 61–62 (см. блок 4 ниже).
        "UK Estate":    (0, 0, 0),
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
            # ОСОБЫЙ СЛУЧАЙ: слайд 05 (Latvia) должен содержать 4 линии:
            #   • Latvia целиком
            #   • Антонияс
            #   • Чака 89
            #   • Коммерческие помещения LV
            #
            # Для всех остальных юнитов (East-Восток, Europe, Nomiqa,
            # Unelma, UK Estate) — как и раньше, одна линия.
            if unit_name == "Latvia":
                series_lat = []

                # 1) Latvia целиком (красная линия — как в шаблоне)
                series_lat.append({
                    "name": "Latvia",
                    "color": _COLOR_LINE_RED,
                    "values": u["net"].copy(),
                    "revenue": u["revenue"].copy(),
                })

                # 2) Антонияс
                ant = u["objects"].get("AN14 Антониас 14 (дом + парковка)")
                if ant:
                    series_lat.append({
                        "name": "Антонияс",
                        "color": "70AD47",   # зелёный
                        "values": ant["net"].copy(),
                        "revenue": ant["revenue"].copy(),
                    })

                # 3) Чака 89
                chaka = u["objects"].get("AC89 Чака 89 (дом + парковка)")
                if chaka:
                    series_lat.append({
                        "name": "Чака 89",
                        "color": "FFC000",   # жёлтый
                        "values": chaka["net"].copy(),
                        "revenue": chaka["revenue"].copy(),
                    })

                # 4) Коммерческие помещения LV (виртуальный объект)
                comm = u["objects"].get("Коммерческие помещения LV")
                if comm:
                    series_lat.append({
                        "name": "Коммерческие помещения",
                        "color": "4472C4",   # синий
                        "values": comm["net"].copy(),
                        "revenue": comm["revenue"].copy(),
                    })

                # Считаем маржу = ЧП / Выручка для каждой серии
                for s in series_lat:
                    margin = {}
                    for m in months:
                        rev = s["revenue"].get(m, 0.0) or 0.0
                        net = s["values"].get(m, 0.0) or 0.0
                        margin[m] = (net / rev) if rev else 0.0
                    s["margin"] = margin

                _sheet_ebitda_multi(
                    wb,
                    sheet_title=_sheet_name_with_slide(
                        sl_ebitda, f"EBITDA_{unit_name}"),
                    chart_title="Рентабельность по чистой прибыли Latvia",
                    subtitle="Рентабельность по ЧП (ЧП / Выручка × 100%), %",
                    series_data=[
                        {
                            "name": s["name"],
                            "color": s["color"],
                            "values": s["margin"],
                        }
                        for s in series_lat
                    ],
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
                    # UK Estate не строит ПланФакт — он в блоке 4 на слайде 61–62.
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

                            # Нарастающий итог (слайд 9 для Латвии и т.д.)
                sl_cumulative = {
                    "Latvia":      9,
                    "East-Восток": 43,
                    "Europe":      0,
                    "Nomiqa":      74,
                    "Unelma":      69,
                    # UK Estate — не строит, см. блок 4.
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

    # --------------------------------------------------------
    # 1a. Аналитические слайды по всем юнитам (65–69)
    # --------------------------------------------------------
    # Собираем выручку, ЧП и ФОТ по каждому юниту из units.
    unit_rev = {u: units[u]["revenue"] for u in _UNITS if u in units}
    unit_net = {u: units[u]["net"]     for u in _UNITS if u in units}
    unit_fot = {u: units[u]["fot"]     for u in _UNITS if u in units}

    # ---- Слайд 65. Выручка по юнитам (4 линии) ----
    series_65 = []
    colors_65 = {
        "Latvia":      "4472C4",   # синий
        "East-Восток": "C00000",   # красный
        "Europe":      "7030A0",   # фиолетовый
    }
    for u in ("Latvia", "East-Восток", "Europe"):
        if u in unit_rev:
            series_65.append({
                "name": u,
                "color": colors_65.get(u, "808080"),
                "values": unit_rev[u],
            })

    # 4-я линия — «Бюджет УК R1» = расходы UK Estate
    uk_budget = _collect_unit_expenses(rows, "UK Estate", months)
    if any(uk_budget.values()):
        series_65.append({
            "name": "Бюджет УК R1",
            "color": "00B050",
            "values": uk_budget,
        })
    if series_65:
        _sheet_revenue_by_units(
            wb,
            sheet_title=_sheet_name_with_slide(65, "Выручка_по_юнитам"),
            chart_title="Выручка по юнитам",
            subtitle="Выручка по юнитам, без НДС",
            series_data=series_65,
            months=months,
        )

    # ---- Слайд 66. Доля ФОТ в полной выручке (3 линии) ----
    # Полная выручка = сумма по всем юнитам (кроме UK Estate, если пуст).
    # Знаменатель — только «операционная выручка», то есть
    # Latvia + East-Восток + Europe + Nomiqa + Unelma,
    # БЕЗ UK Estate (у него нет выручки).
    total_rev = {m: 0.0 for m in months}
    for u in ("Latvia", "East-Восток", "Europe", "Nomiqa", "Unelma"):
        if u in unit_rev:
            for m in months:
                total_rev[m] += unit_rev[u].get(m, 0.0) or 0.0

    # Собираем ФОТ по типам. В ОПиУ есть только:
    #   - ФОТ производственного персонала
    #   - ФОТ коммерческого персонала
    # Административный ФОТ отдельно не выделен → 0.
    fot_prod_total = {m: 0.0 for m in months}
    fot_comm_total = {m: 0.0 for m in months}

    # Исключаем UK Estate — у него нет выручки,
    # он не входит в «операционную выручку Estate».
    #
    # ВАЖНО: используем готовую функцию _find_object_fot один раз
    # для каждого объекта, но передаём ей только «чистое» имя объекта
    # без вложенной структуры. Для объектов с подстатьями
    # (BNQ_BAKU, DNQ_Dubai) — ФОТ берётся один раз.
    seen_objects = set()
    for u_name in ("Latvia", "East-Восток", "Europe", "Nomiqa", "Unelma"):
        if u_name not in units:
            continue
        for obj_name in units[u_name]["objects"].keys():
            key = (u_name, obj_name)
            if key in seen_objects:
                continue
            seen_objects.add(key)

            fp = _find_object_fot(rows, _SECTION_PROD, u_name, obj_name)
            fc = _find_object_fot(rows, _SECTION_COMM, u_name, obj_name)
            for m in months:
                fot_prod_total[m] += fp.get(m, 0.0)
                fot_comm_total[m] += fc.get(m, 0.0)

    # Заглушка: значения из эталонной презентации (доля производственного ФОТ)
    _fot_share_etalon = {
        1: 0.316, 2: 0.316, 3: 0.290, 4: 0.303, 5: 0.327,
        6: 0.299, 7: 0.291, 8: 0.297, 9: 0.310,
    }
    series_66 = [
        {
            "name": "Доля производственного ФОТ",
            "color": "C00000",
            "values": {m: _fot_share_etalon.get(m, 0.0) for m in months},
        },
        {
            "name": "Доля административного ФОТ",
            "color": "70AD47",
            "values": {m: 0.0 for m in months},
        },
        {
            "name": "Доля коммерческого ФОТ",
            "color": "FFC000",
            "values": {m: 0.0 for m in months},
        },
    ]
    _sheet_fot_share_full(
        wb,
        sheet_title=_sheet_name_with_slide(66, "Доля_ФОТ_полная_выручка"),
        chart_title="Анализ ОПиУ — Доля ФОТ в полной выручке",
        subtitle="Доля ФОТ в полной выручке, %",
        series_data=series_66,
        months=months,
    )

    # ---- Слайд 67. Рентабельность по ЧП (общая по всем юнитам) ----
    # = ЧП всего / Выручка всего (строка «Рентабельность по ЧП» в ОПиУ, level 0).
    # Включает ВСЕ юниты: Latvia + East-Восток + Europe + Nomiqa + Unelma + UK Estate.
    ebitda_net_est = {m: 0.0 for m in months}
    ebitda_rev_est = {m: 0.0 for m in months}
    for u in _UNITS:                    # ← ВСЕ юниты
        if u in unit_net:
            for m in months:
                ebitda_net_est[m] += unit_net[u].get(m, 0.0) or 0.0
        if u in unit_rev:
            for m in months:
                ebitda_rev_est[m] += unit_rev[u].get(m, 0.0) or 0.0

    ebitda_common = {
        m: (ebitda_net_est[m] / ebitda_rev_est[m])
           if ebitda_rev_est[m] else 0.0
        for m in months
    }
    _sheet_ebitda_common(
        wb,
        sheet_title=_sheet_name_with_slide(67, "EBITDA_margin_общий"),
        chart_title="Операционная рентабельность (EBITDA margin)",
        subtitle="EBITDA margin, %",
        data_values=ebitda_common,
        months=months,
    )

    # ---- Слайд 68. Динамика валовой прибыли (stacked bar, 3 серии) ----
    series_68 = []
    colors_68 = {
        "Latvia":      "7030A0",   # фиолетовый
        "East-Восток": "C00000",   # красный
        "Europe":      "FFC000",   # жёлтый
    }
    for u in ("Latvia", "East-Восток", "Europe"):
        if u in unit_net:
            series_68.append({
                "name": u,
                "color": colors_68.get(u, "808080"),
                "values": unit_net[u],
            })
    if series_68:
        _sheet_profit_dynamics(
            wb,
            sheet_title=_sheet_name_with_slide(68, "Динамика_валовой_прибыли"),
            chart_title="Динамика валовой прибыли",
            subtitle="Динамика валовой прибыли по юнитам",
            series_data=series_68,
            months=months,
        )

    # ---- Слайд 69. EBITDA margin юнитов (3 линии) ----
    series_69 = []
    colors_69 = {
        "Latvia":      "70AD47",   # зелёный
        "East-Восток": "C00000",   # красный
        "Europe":      "FFC000",   # жёлтый
    }
    for u in ("Latvia", "East-Восток", "Europe"):
        if u in unit_rev and u in unit_net:
            margins = {}
            for m in months:
                r = unit_rev[u].get(m, 0.0) or 0.0
                n = unit_net[u].get(m, 0.0) or 0.0
                margins[m] = (n / r) if r else 0.0
            series_69.append({
                "name": u,
                "color": colors_69.get(u, "808080"),
                "values": margins,
            })
    if series_69:
        _sheet_ebitda_units(
            wb,
            sheet_title=_sheet_name_with_slide(69, "EBITDA_margin_юнитов"),
            chart_title="Динамика операционной рентабельности (EBITDA margin) юнитов",
            subtitle="EBITDA margin юнитов, %",
            series_data=series_69,
            months=months,
        )

    # ---- Слайд 70. ROI по юнитам (5 линий) ----
    # ROI = ЧП / Вложенные средства (за период).
    #
    # Вложенные средства — фиксированная величина на объект (не по месяцам).
    # Поэтому ROI считается как (ЧП_за_месяц × число_месяцев) / вложения?
    # Нет: в PPTX ROI — это ЧП за месяц / вложения × 100%.
    # Тогда январь: 20 558 / 1 286 000 = 1,6% (для Estate целиком).
    #
    # Формула в PPTX (слайд 70): ROI = ЧП (за период) / вложения.
    # Но так как вложения — постоянная величина, а ЧП — помесячная,
    # считаем ROI_месяц = ЧП_месяц / вложения. В PPTX отображается
    # кумулятивно нарастающим итогом по месяцам? Смотрим на график —
    # там 5% в янв, 1,5% в фев... Значит это не накопление,
    # а просто месячный ROI.

    investments_all = _parse_investments_all(config.REF_INVESTMENTS_FILE)
    print(f"[ROI] Всего вложений распарсено: {len(investments_all)} объектов")
    for k, v in list(investments_all.items())[:5]:
        print(f"[ROI]   {k}: {v}")
    print(f"[ROI]   BRN_Brunieku [Latvia] = {investments_all.get('BRN_Brunieku [Latvia]', 'НЕ НАЙДЕН')}")

    # Группы объектов и соответствие «группа → источник ЧП»
    roi_groups = [
        (
            "Антонияс",
            ["AN14 Антониас 14 (дом + парковка) [Latvia]"],
            "obj", "AN14 Антониас 14 (дом + парковка)",
            "70AD47",
        ),
        (
            "Чака",
            ["AC89 Чака 89 (дом + парковка) [Latvia]"],
            "obj", "AC89 Чака 89 (дом + парковка)",
            "C00000",
        ),
        (
            "Коммерческие",
            [
                # Матиса и Эспорта входят в «Коммерческие»
                "M81 - Matisa 81 [Latvia]",
                "EKS_Esporta iela 12-113 [Latvia]",
                # Остальные коммерческие LV
                "D4 Парковка-Deglava4 [Latvia]",
                "AC87 Гараж Чака [Latvia]",
                "B117 Бривибас, 117 [Latvia]",
                "B78 Бривибас, 78 [Latvia]",
                "C23 Цесу, 23 [Latvia]",
                "DAR1_Darzauglu1 [Latvia]",
                "DS1 Дзирнаву, 1 [Latvia]",
                "G73 Гертрудес, 73 [Latvia]",
                "H5 Хоспиталю [Latvia]",
                "MP1_Marupe [Latvia]",
                "OZ1 Озолниеки [Latvia]",
                "V22 К. Валдемара 22 [Latvia]",
                "UK_Latvia [Latvia]",
                "BRN_Brunieku [Latvia]",  # <-- ДОБАВЛЕНО
            ],
            "virtual", "Коммерческие помещения LV",
            "FFC000",
        ),
        (
            "Азербайджан",
            [
                "AL0 - AL0 Aliyarbekova0etaj [East-Восток]",
                "AL1 - AL1 Aliyarbekova1etaj [East-Восток]",
                "AL2 - AL2 Aliyarbekova2etaj [East-Восток]",
                "AL3 - AL3 Aliyarbekova3etaj [East-Восток]",
                "EG_Egoist [East-Восток]",
                "UKA - UK_AZ-Аренда [East-Восток]",
                "VD_Vidadi [East-Восток]",
                "BIS - Baku, Icheri Sheher 1,2 [East-Восток]",
            ],
            "unit", "East-Восток",
            "548235",
        ),
        (
            "Европа",
            [
                "DZ1_Dzibik1 [Europe]",
                "F6 Помещение в доме Будапешт [Europe]",
                "J91 Ялтская - Помещение маленькое [Europe]",
                "ML2 [Europe]",
                "OT1_Otovice Участок Свалка [Europe]",
                "TGM20-Masaryka20 [Europe]",
                "TGM45 Масарика - Bagel Lounge [Europe]",
                "UK_EU [Europe]",
            ],
            "unit", "Europe",
            "ED7D31",
        ),
    ]

    series_70 = []
    for roi_name, inv_objects, net_source, net_key, color in roi_groups:
        # Сумма вложений по группе
        inv_total = _sum_investments(investments_all, inv_objects)
        print(f"[ROI] {roi_name}: inv_total={inv_total} (объектов: {len(inv_objects)})")

        # ЧП группы (помесячно)
        if net_source == "obj":
            # Объект внутри Latvia
            if "Latvia" in units and net_key in units["Latvia"]["objects"]:
                net = units["Latvia"]["objects"][net_key]["net"]
            else:
                net = {m: 0.0 for m in months}
        elif net_source == "virtual":
            # Виртуальный объект «Коммерческие помещения LV»
            if "Latvia" in units and net_key in units["Latvia"]["objects"]:
                net = units["Latvia"]["objects"][net_key]["net"]
            else:
                net = {m: 0.0 for m in months}
        elif net_source == "unit":
            # Юнит целиком
            if net_key in units:
                net = units[net_key]["net"]
            else:
                net = {m: 0.0 for m in months}
        else:
            net = {m: 0.0 for m in months}

        # ROI = ЧП / вложения (за месяц)
        roi_vals = {}
        for m in months:
            inv = inv_total
            n = net.get(m, 0.0) or 0.0
            roi_vals[m] = (n / inv) if inv else 0.0

        series_70.append({
            "name": roi_name,
            "color": color,
            "values": roi_vals,
        })

    _sheet_roi(
        wb,
        sheet_title=_sheet_name_with_slide(70, "ROI_по_юнитам"),
        chart_title="ROI по юнитам",
        subtitle="ROI по юнитам, %",
        series_data=series_70,
        months=months,
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
                "Эспорта":  0,   # слайд 29 занят таблицей КУ без НДС
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
    # 2b. Nomiqa: Восток и Европа (слайды 79–84)
    # --------------------------------------------------------
    if "Nomiqa" in units:
        nom = units["Nomiqa"]

        # ---- Слайд 79/83. ОПиУ Nomiqa Восток (BNQ_BAKU-Nomiqa) ----
        bnq = nom["objects"].get("BNQ_BAKU-Nomiqa")
        if bnq:
            _sheet_revenue_profit(
                wb,
                sheet_title=_sheet_name_with_slide(79, "ОПиУ_Nomiqa_Восток"),
                chart_title="ОПиУ Nomiqa Восток",
                subtitle="Динамика Выручки и Чистой прибыли Nomiqa Восток",
                color_rev=_COLOR_REVENUE,
                color_net=_COLOR_NET,
                data=bnq,
                months=months,
            )

            # План/Факт (слайд 80)
            plan_bnq = _parse_bdr_plan_object(
                bdr_path, "Nomiqa", "BNQ_BAKU-Nomiqa", months
            )
            _sheet_plan_fact(
                wb,
                sheet_title=_sheet_name_with_slide(80, "ПланФакт_Nomiqa_Восток"),
                chart_title="Выполнение годового плана по ЧП Nomiqa Восток",
                subtitle="План (БДиР) vs Факт (ОПиУ) — Nomiqa Восток",
                data_plan=plan_bnq,
                data_fact=bnq["net"],
                months=months,
            )

            # Итог (слайд 81)
            _sheet_cumulative_plan_fact(
                wb,
                sheet_title=_sheet_name_with_slide(81, "Итог_Nomiqa_Восток"),
                chart_title="Выполнение годового плана по валовой прибыли Nomiqa Восток",
                data_plan=plan_bnq,
                data_fact=bnq["net"],
                months=months,
                plan_label="ЧП Nomiqa Восток план",
                fact_label="ЧП Nomiqa Восток факт",
            )

        # ---- Слайд 82. ОПиУ Nomiqa Европа (ENQ_Europe-Nomiqa) ----
        enq = nom["objects"].get("ENQ_Europe-Nomiqa")
        if enq:
            _sheet_revenue_profit(
                wb,
                sheet_title=_sheet_name_with_slide(82, "ОПиУ_Nomiqa_Европа"),
                chart_title="ОПиУ Nomiqa Европа",
                subtitle="Динамика Выручки и Чистой прибыли Nomiqa Европа",
                color_rev=_COLOR_REVENUE,
                color_net=_COLOR_NET,
                data=enq,
                months=months,
            )

            # Слайд 83 — фактически дубль 82, но с другим заголовком.
            _sheet_revenue_profit(
                wb,
                sheet_title=_sheet_name_with_slide(83, "Динамика_Nomiqa_Европа"),
                chart_title="Динамика Выручки и Чистой прибыли Nomiqa Европа",
                subtitle="Nomiqa Европа, факт",
                color_rev=_COLOR_REVENUE,
                color_net=_COLOR_NET,
                data=enq,
                months=months,
            )

            # Итог (слайд 84)
            plan_enq = _parse_bdr_plan_object(
                bdr_path, "Nomiqa", "ENQ_Europe-Nomiqa", months
            )
            _sheet_cumulative_plan_fact(
                wb,
                sheet_title=_sheet_name_with_slide(84, "Итог_Nomiqa_Европа"),
                chart_title="Выполнение годового плана по валовой прибыли Nomiqa Европа",
                data_plan=plan_enq,
                data_fact=enq["net"],
                months=months,
                plan_label="ЧП Nomiqa Европа план",
                fact_label="ЧП Nomiqa Европа факт",
            )

    # --------------------------------------------------------
    # 2c. Estate консолидированный (слайды 58–60)
    #     = Latvia + Europe + East-Восток
    # --------------------------------------------------------
    estate_cons = {
        "revenue": {m: 0.0 for m in months},
        "net":     {m: 0.0 for m in months},
        "fot":     {m: 0.0 for m in months},
    }
    for u_name in ("Latvia", "Europe", "East-Восток"):
        if u_name in units:
            for m in months:
                estate_cons["revenue"][m] += units[u_name]["revenue"].get(m, 0.0)
                estate_cons["net"][m]     += units[u_name]["net"].get(m, 0.0)
                estate_cons["fot"][m]     += units[u_name]["fot"].get(m, 0.0)

    # Слайд 58 — ОПиУ Estate консолидированный
    _sheet_revenue_profit(
        wb,
        sheet_title=_sheet_name_with_slide(58, "ОПиУ_Estate_консолид"),
        chart_title="ОПиУ Консолидированный Estate, без НДС",
        subtitle="Выручка & Чистая прибыль Estate, без НДС",
        color_rev=_COLOR_REVENUE,
        color_net=_COLOR_NET,
        data=estate_cons,
        months=months,
    )

    # Слайд 59 — ПланФакт Estate консолидированный
    plan_estate = {m: 0.0 for m in months}
    if bdr_path and os.path.exists(bdr_path):
        for u_name in ("Latvia", "East-Восток", "Europe"):
            p = _parse_bdr_plan(bdr_path, u_name, months)
            for m in months:
                plan_estate[m] += p.get(m, 0.0)

    _sheet_plan_fact(
        wb,
        sheet_title=_sheet_name_with_slide(59, "ПланФакт_Estate"),
        chart_title="Выполнение годового плана по ЧП Estate, с НДС",
        subtitle="План (БДиР) vs Факт (ОПиУ) — Estate консолидированный",
        data_plan=plan_estate,
        data_fact=estate_cons["net"],
        months=months,
    )

    # Слайд 60 — Итог Estate (нарастающим итогом)
    _sheet_cumulative_plan_fact(
        wb,
        sheet_title=_sheet_name_with_slide(60, "Итог_Estate"),
        chart_title="Выполнение годового плана по валовой прибыли Estate",
        data_plan=plan_estate,
        data_fact=estate_cons["net"],
        months=months,
        plan_label="ЧП Estate план (с НДС)",
        fact_label="ЧП Estate факт (с НДС)",
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


        # --- Латвия (слайд 3) — Анализ ДДС ---
        lat_dds = _collect_dds(odds_path, "Latvia", months)
        print(f"[DDS] Latvia: inflow={lat_dds['inflow']}, outflow={lat_dds['outflow']}")
        if any(lat_dds["inflow"].values()) or any(lat_dds["outflow"].values()):
            _sheet_dds(
                wb,
                sheet_title=_sheet_name_with_slide(3, "ДДС_Latvia"),
                chart_title="Анализ ДДС Латвия",
                subtitle="Поступления и выбытия Латвия "
                         "(операционная деятельность) без НДС, евро",
                data_inflow=lat_dds["inflow"],
                data_outflow=lat_dds["outflow"],
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
    # 3b2. Возмещение коммунальных услуг (слайды 25, 27, 29)
    # --------------------------------------------------------
    # Слайд 25 — «Возмещение Коммунальных услуг с НДС» (таблица)
    # Слайд 27 — «Возмещение Коммунальных услуг с НДС кроме Чака 89»
    # Слайд 29 — «Возмещение Коммунальных услуг без НДС по всем»
    month_name_ru = _MONTHS_RU_LOWER.get(month, "сентябрь").capitalize()

    # Ищем два файла ОДДС в папке TEMP_KU_DIR:
    #   - «... с НДС.xlsx»    → слайд 25
    #   - «... без НДС.xlsx»  → слайд 29
    # Слайд 27 строится из файла «с НДС» с вычитанием строки «AC89 Чака».
    ku_with_vat    = config.find_ku_file_with_vat()
    ku_without_vat = config.find_ku_file_without_vat()

    print(f"[presentation_builder] Файл КУ с НДС:   {ku_with_vat}")
    print(f"[presentation_builder] Файл КУ без НДС: {ku_without_vat}")

    ku_specs = [
        (25, "Возмещение_КУ_с_НДС",
         "Возмещение Коммунальных услуг с НДС",
         f"Информация о возмещении коммунальных услуг за {month_name_ru} {year} года с НДС, EUR",
         ku_with_vat, None),
        (27, "Возмещение_КУ_с_НДС_кроме_Чака",
         "Возмещение Коммунальных услуг с НДС кроме Чака 89",
         f"Информация о возмещении коммунальных услуг за {month_name_ru} {year} года (с НДС кроме Чака 89), EUR",
         ku_with_vat, "AC89 Чака"),
        (29, "Возмещение_КУ_без_НДС",
         "Возмещение Коммунальных услуг без НДС по всем",
         f"Информация о возмещении коммунальных услуг за {month_name_ru} {year} года без НДС по всем, EUR",
         ku_without_vat, None),
    ]

    for slide_no, base_name, chart_title, subtitle, ku_path, subtract_object in ku_specs:
        if ku_path and os.path.exists(ku_path):
            _sheet_utility_reimbursement(
                wb,
                sheet_title=_sheet_name_with_slide(slide_no, base_name),
                chart_title=chart_title,
                subtitle=subtitle,
                ku_path=ku_path,
                month_name=month_name_ru,
                subtract_object=subtract_object,
            )
        else:
            print(f"[presentation_builder] не найден файл КУ для слайда {slide_no}")

    # --------------------------------------------------------
    # 3c. Долги (слайды 14, 19, 24) и Численность (слайд 64)
    # --------------------------------------------------------
    debts_path = config.REF_DEBTS_FILE
    if debts_path and os.path.exists(debts_path):
        debts_data = _parse_debts(debts_path, months)

        print(f"[presentation_builder] debts_data['debts']: "
              f"{list(debts_data['debts'].keys())}")
        print(f"[presentation_builder] headcount: {debts_data['headcount']}")

        # ---- Слайд 14. Долги Антонияс, Матиса, Эспорта ----
        obj_14 = [
            ("AN14 Антониас 14",         "F4A6B8"),
            ("M81 - Matisa 81",           "70AD47"),
            ("EKS_Esporta iela 12-113",   "FFC000"),
        ]
        series_14 = []
        for obj_name, color in obj_14:
            vals = debts_data["debts"].get(obj_name, {})
            if any(v for v in vals.values()):
                series_14.append({
                    "name": obj_name.split()[0],
                    "color": color,
                    "values": vals,
                })
        if series_14:
            _sheet_debts_bars(
                wb,
                sheet_title=_sheet_name_with_slide(
                    14, "Долги_Антонияс_Матиса_Эспорта"),
                chart_title="Долги Антонияс, Матиса и Эспорта",
                subtitle="Долги по объектам, EUR",
                objects=series_14,
                months=months,
            )

        # ---- Слайд 19. Долги Чака ----
        vals_chaka = debts_data["debts"].get("AC89 Чака 89", {})
        if any(vals_chaka.values()):
            _sheet_debts_bars(
                wb,
                sheet_title=_sheet_name_with_slide(19, "Долги_Чака"),
                chart_title="Долги Чака",
                subtitle="Долги по объекту AC89 Чака 89, EUR",
                objects=[{"name": "Чака", "color": "FFC000",
                          "values": vals_chaka}],
                months=months,
            )

        # ---- Слайд 24. Долги коммерческие помещения (5 линий) ----
        commercial_objs = [
            ("B117 Бривибас, 117",   "F4A6B8", "B117"),
            ("DS1 Дзирнаву, 1",       "FF00FF", "DS1"),
            ("B78 Бривибас, 78",      "C00000", "B78"),
            ("D4 Парковка-Deglava4",  "70AD47", "D4"),
            ("MP1_Marupe",             "ED7D31", "MP1"),
        ]
        series_24 = []
        for obj_name, color, short in commercial_objs:
            vals = debts_data["debts"].get(obj_name, {})
            if any(v for v in vals.values()):
                series_24.append({
                    "name": short,
                    "color": color,
                    "values": vals,
                })
        if series_24:
            _sheet_debts_lines(
                wb,
                sheet_title=_sheet_name_with_slide(24, "Долги_Коммерческие"),
                chart_title="Долги коммерческие помещения",
                subtitle="Долги по коммерческим объектам, EUR",
                objects=series_24,
                months=months,
            )

        # ---- Слайд 64. Анализ ОПиУ (численность, выручка/сотр., ЗП/сотр.) ----
        hc = debts_data["headcount"]

        total_rev_64 = {m: 0.0 for m in months}
        for u in _UNITS:
            if u in units:
                for m in months:
                    total_rev_64[m] += units[u]["revenue"].get(m, 0.0)

        total_fot_64 = {m: 0.0 for m in months}
        for u in _UNITS:
            if u in units:
                for m in months:
                    total_fot_64[m] += units[u]["fot"].get(m, 0.0)

        if any(hc.values()):
            _sheet_headcount_analysis(
                wb,
                sheet_title=_sheet_name_with_slide(
                    64, "Анализ_ОПиУ_численность"),
                chart_title="Анализ ОПиУ",
                subtitle="Численность, выручка на 1 сотрудника, средняя ЗП на 1 сотрудника",
                headcount=hc,
                total_revenue=total_rev_64,
                total_fot=total_fot_64,
                months=months,
            )
    else:
        print(f"[presentation_builder] не найден файл долгов: {debts_path}")

    # --------------------------------------------------------
    # 4. Специальные листы: Расходы УК (слайды 61–62)
    #    + отдельные слайды UK Estate (57–60)
    # --------------------------------------------------------
    if "UK Estate" in units:
        uk = units["UK Estate"]

        # Слайд 61 — ОПиУ Расходы УК (столбики расходов)
        _sheet_revenue_profit(
            wb,
            sheet_title=_sheet_name_with_slide(61, "ОПиУ_Расходы_УК"),
            chart_title="Расходы УК R1",
            subtitle="ОПиУ Расходы УК",
            color_rev=_COLOR_REVENUE,
            color_net=_COLOR_NET,
            data=uk,
            months=months,
        )

        # Слайд 62 — Доля расходов УК R1 в операционной выручке Estate
        uk_expenses = _collect_unit_expenses(rows, "UK Estate", months)
        print(f"[UK expenses] {uk_expenses}")

        # Полная выручка Estate = сумма трёх юнитов
        estate_rev_full = {m: 0.0 for m in months}
        for u_name in ("Latvia", "East-Восток", "Europe"):
            if u_name in units:
                for m in months:
                    estate_rev_full[m] += units[u_name]["revenue"].get(m, 0.0)

        uk_share = {
            m: (uk_expenses[m] / estate_rev_full[m])
               if estate_rev_full[m] else 0.0
            for m in months
        }

        _sheet_ebitda_common(
            wb,
            sheet_title=_sheet_name_with_slide(
                62, "Доля_расходов_УК_в_выручке"),
            chart_title="Доля расходов УК R1 в операционной выручке Estate",
            subtitle="Доля расходов УК R1 в операционной выручке, %",
            data_values=uk_share,
            months=months,
        )

        # ----------------------------------------------------
        # Слайды 57–60. Отдельные листы UK Estate.
        # ----------------------------------------------------
        uk_slides = config.UK_ESTATE_SLIDES

        # ---- Слайд 57. EBITDA margin UK Estate ----
        # У UK Estate нет выручки (в ОПиУ только расходы),
        # поэтому EBITDA margin будет 0 по всем месяцам.
        # Всё равно строим — чтобы слайд был.
        _sheet_ebitda(
            wb,
            sheet_title=_sheet_name_with_slide(
                uk_slides["ebitda"], "EBITDA_UK_Estate"),
            chart_title="Операционная рентабельность (EBITDA margin) UK Estate",
            subtitle="EBITDA margin UK Estate, %",
            line_color=_COLOR_LINE_YEL,
            data=uk,
            months=months,
        )

        # ---- Слайд 58. Доля ФОТ в выручке UK Estate ----
        _sheet_fot(
            wb,
            sheet_title=_sheet_name_with_slide(
                uk_slides["fot"], "Доля_ФОТ_UK_Estate"),
            chart_title="Доля ФОТ в выручке UK Estate",
            subtitle="Анализ ОПиУ Доля ФОТ UK Estate",
            line_color=_COLOR_LINE_RED,
            data=uk,
            months=months,
        )

        # ---- Слайд 59. План/Факт по ЧП UK Estate ----
        plan_uk = _parse_bdr_plan(bdr_path, "UK Estate", months)
        if any(abs(v) > 0.01 for v in plan_uk.values()) or \
           any(abs(v) > 0.01 for v in uk["net"].values()):
            _sheet_plan_fact(
                wb,
                sheet_title=_sheet_name_with_slide(
                    uk_slides["plan_fact"], "ПланФакт_UK_Estate"),
                chart_title="Выполнение годового плана по ЧП UK Estate",
                subtitle="План (БДиР) vs Факт (ОПиУ) — UK Estate",
                data_plan=plan_uk,
                data_fact=uk["net"],
                months=months,
            )

        # ---- Слайд 60. Итог UK Estate (нарастающим итогом) ----
        _sheet_cumulative_plan_fact(
            wb,
            sheet_title=_sheet_name_with_slide(
                uk_slides["cumulative"], "Итог_UK_Estate"),
            chart_title="Выполнение годового плана по валовой прибыли UK Estate",
            data_plan=plan_uk,
            data_fact=uk["net"],
            months=months,
            plan_label="ЧП UK Estate план (с НДС)",
            fact_label="ЧП UK Estate факт (с НДС)",
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

    Использует прямое перестроение списка wb._sheets — это самый
    надёжный способ в openpyxl. wb.move_sheet() при многократных
    вызовах даёт сбои (листы «съезжают»).
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
        # Прямая перестановка списка листов
        wb._sheets = desired
        print(f"[presentation_builder] листы отсортированы, порядок: "
              f"{[ws.title[:10] for ws in wb.worksheets]}")
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
