# src/presentation_builder.py
# ============================================================
# Генератор Excel с диаграммами для презентации.
#
# Читает ОПиУ формата «статьи × месяцы» с древовидной структурой
# (·, ··, ···, ····) и строит книгу Excel.
#
# ГЛАВНАЯ ИДЕЯ ПАРСИНГА:
#   1) Один раз пробегаем файл и строим плоский список строк:
#         [(row_idx, level, title_clean, {месяц: значение})]
#   2) Для каждой искомой величины (выручка юнита, ЧП юнита,
#      ФОТ объекта и т.п.) проходим по этому списку с ЯВНЫМ
#      условием «находимся внутри секции X, юнит Y, объект Z».
#      Это исключает любые сдвиги контекста.
# ============================================================

import os
import re
import datetime as _dt

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
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

# Объекты Латвии для отдельных листов: (имя в файле, имя для листа)
_LATVIA_OBJECTS = [
    ("AN14 Антониас 14 (дом + парковка)", "Антонияс"),
    ("AC89 Чака 89 (дом + парковка)",     "Чака"),
    ("M81 - Matisa 81",                    "Матиса"),
    ("EKS_Esporta iela 12-113",            "Эспорта"),
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

_COLOR_REVENUE = "FFC000"   # жёлтый
_COLOR_NET     = "70AD47"   # зелёный
_COLOR_LINE_1  = "C00000"   # красный
_COLOR_LINE_2  = "4472C4"   # синий


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
    """
    Возвращает:
        rows = [(row_idx, level, title_clean, {month: value})]
        month_col = {month: col_idx}

    Никакой интерпретации — просто дамп всех строк с уровнями.
    """
    if not opiu_path or not os.path.exists(opiu_path):
        return [], {}

    wb = openpyxl.load_workbook(opiu_path, data_only=True)
    ws = wb.active

    # Ищем шапку: строка, где минимум 3 колонки содержат названия месяцев
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
# ШАГ 2. ИЗВЛЕЧЕНИЕ ДАННЫХ ИЗ ПЛОСКОГО СПИСКА
# ============================================================

# Секции (level 0), в которых мы ищем:
#   - Выручка              → revenue
#   - Чистая прибыль       → net
#   - Производственные расходы → ФОТ (производственный)
#   - Коммерческие расходы     → ФОТ (коммерческий)
_SECTION_REVENUE = "Выручка"
_SECTION_NET     = "Чистая прибыль"
_SECTION_PROD    = "Производственные расходы"
_SECTION_COMM    = "Коммерческие расходы"


def _find_row(rows, section, unit, object_name=None, level_expected=None):
    """
    Возвращает словарь {month: value} для строки, которая находится:
      - в секции `section` (level 0),
      - затем идёт `unit` (level 1, точное совпадение),
      - затем, если задан, `object_name` (level 2, точное совпадение),
      - и сама строка имеет уровень `level_expected`.

    Логика:
      - проходим список rows по порядку;
      - отслеживаем текущую секцию (level 0),
        текущего юнита (level 1),
        текущий объект (level 2);
      - когда все три совпали и уровень строки = level_expected —
        возвращаем значения.

    Если object_name=None, ищем строку, у которой level == 1
    (юнит) в нужной секции.
    """
    cur_section = None
    cur_unit = None
    cur_object = None

    for r, lvl, title, values in rows:
        if lvl == 0:
            cur_section = title
            cur_unit = None
            cur_object = None
            continue

        # Секция сменилась — прерываемся, если уже нашли нужное
        if cur_section != section:
            if cur_section is not None and cur_section != section:
                # Возможно, мы ушли из нужной секции — но могли
                # вернуться позже. Просто продолжаем.
                pass

        if lvl == 1:
            cur_unit = title
            cur_object = None
            # Проверяем: может, это искомая строка (юнит без объекта)
            if (object_name is None
                    and cur_section == section
                    and cur_unit == unit):
                return dict(values)
            continue

        if lvl == 2:
            cur_object = title
            # Юнит и объект без секции нам не подходят
            continue

    return {}


def _find_object_fot(rows, section, unit, object_name):
    """
    Возвращает {month: value} для ФОТ объекта.
    ФОТ — строка уровня 4 (···) с текстом, содержащим «ФОТ
    производственного персонала» или «ФОТ коммерческого персонала»,
    внутри блока ·· <object_name> внутри · <unit> в секции `section`.
    """
    cur_section = None
    cur_unit = None
    cur_object = None
    result = {}

    for r, lvl, title, values in rows:
        if lvl == 0:
            cur_section = title
            cur_unit = None
            cur_object = None
            continue
        if lvl == 1:
            cur_unit = title
            cur_object = None
            continue
        if lvl == 2:
            cur_object = title
            continue

        # Уровень 3+ внутри объекта
        if (cur_section == section
                and cur_unit == unit
                and cur_object == object_name
                and lvl >= 3):
            tl = title.lower()
            if "фот производственного" in tl or \
               "фот коммерческого" in tl:
                for m, v in values.items():
                    result[m] = result.get(m, 0.0) + v

    return result


# ============================================================
# ШАГ 3. СБОРКА ДАННЫХ ПО ЮНИТАМ И ОБЪЕКТАМ
# ============================================================

def _collect_units(rows, months):
    """
    Возвращает:
    {
        "Latvia": {
            "revenue": {m: v},
            "net":     {m: v},
            "fot":     {m: v},
            "objects": {
                "AN14 ...": {"revenue": {...}, "net": {...}, "fot": {...}},
                ...
            }
        },
        ...
    }
    """


# ============================================================
# ПОСТРОЕНИЕ ЛИСТОВ EXCEL
# ============================================================

def _write_header(ws, title, subtitle=""):
    ws.cell(row=1, column=1, value=title).font = _TITLE_FONT
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = _SUBTITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=13)


def _write_table(ws, start_row, headers, rows):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _CELL_BORDER
    for r_off, row_vals in enumerate(rows, start=1):
        for c, v in enumerate(row_vals, start=1):
            cell = ws.cell(row=start_row + r_off, column=c, value=v)
            cell.border = _CELL_BORDER
            if isinstance(v, (int, float)):
                cell.number_format = '#,##0'


def _style_bar(chart, title, w=20, h=10):
    chart.title = title
    chart.width = w
    chart.height = h
    chart.style = 10
    chart.x_axis.delete = False
    chart.y_axis.delete = False


def _style_line(chart, title, w=20, h=10):
    chart.title = title
    chart.width = w
    chart.height = h
    chart.style = 12
    chart.x_axis.delete = False
    chart.y_axis.delete = False


def _sheet_revenue_profit(wb, name, data, months):
    sheet_name = _safe_sheet(f"ОПиУ_{name}")
    ws = wb.create_sheet(sheet_name) if sheet_name not in wb.sheetnames else wb[sheet_name]

    _write_header(ws, f"ОПиУ {name}", "Выручка и чистая прибыль по месяцам")

    headers = ["Показатель"] + [_MONTHS_RU_CAP[m] for m in months]
    rows = [
        ["Выручка"]         + [round(data["revenue"][m], 2) for m in months],
        ["Чистая прибыль"]  + [round(data["net"][m], 2)     for m in months],
    ]
    _write_table(ws, 4, headers, rows)

    data_ref = Reference(ws, min_col=2, max_col=1 + len(months), min_row=4, max_row=6)
    cats = Reference(ws, min_col=2, max_col=1 + len(months), min_row=4, max_row=4)

    chart = BarChart()
    chart.type = "col"
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    _style_bar(chart, f"Выручка и ЧП — {name}")
    chart.series[0].graphicalProperties.solidFill = _COLOR_REVENUE
    chart.series[1].graphicalProperties.solidFill = _COLOR_NET
    ws.add_chart(chart, "A9")


def _sheet_ebitda(wb, name, data, months):
    sheet_name = _safe_sheet(f"EBITDA_{name}")
    ws = wb.create_sheet(sheet_name) if sheet_name not in wb.sheetnames else wb[sheet_name]

    _write_header(ws, f"Операционная рентабельность — {name}", "EBITDA margin, %")

    margins = {}
    for m in months:
        rev = data["revenue"][m]
        net = data["net"][m]
        margins[m] = round(net / rev, 4) if rev else 0.0

    headers = ["Показатель"] + [_MONTHS_RU_CAP[m] for m in months]
    rows = [["EBITDA margin"] + [margins[m] for m in months]]
    _write_table(ws, 4, headers, rows)

    for c in range(2, 2 + len(months)):
        ws.cell(row=5, column=c).number_format = '0.00%'

    data_ref = Reference(ws, min_col=2, max_col=1 + len(months), min_row=4, max_row=5)
    cats = Reference(ws, min_col=2, max_col=1 + len(months), min_row=4, max_row=4)

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    _style_line(chart, f"EBITDA margin — {name}")
    chart.series[0].graphicalProperties.line.solidFill = _COLOR_LINE_1
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True
    chart.y_axis.numFmt = '0%'
    ws.add_chart(chart, "A9")


def _sheet_fot(wb, name, data, months):
    sheet_name = _safe_sheet(f"Доля_ФОТ_{name}")
    ws = wb.create_sheet(sheet_name) if sheet_name not in wb.sheetnames else wb[sheet_name]

    _write_header(ws, f"Доля ФОТ в выручке — {name}", "%")

    shares = {}
    for m in months:
        rev = data["revenue"][m]
        fot = data["fot"][m]
        # ФОТ в файле отрицательный — берём модуль
        shares[m] = round(abs(fot) / rev, 4) if rev else 0.0

    headers = ["Показатель"] + [_MONTHS_RU_CAP[m] for m in months]
    rows = [["Доля ФОТ"] + [shares[m] for m in months]]
    _write_table(ws, 4, headers, rows)

    for c in range(2, 2 + len(months)):
        ws.cell(row=5, column=c).number_format = '0.00%'

    data_ref = Reference(ws, min_col=2, max_col=1 + len(months), min_row=4, max_row=5)
    cats = Reference(ws, min_col=2, max_col=1 + len(months), min_row=4, max_row=4)

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    _style_line(chart, f"Доля ФОТ — {name}")
    chart.series[0].graphicalProperties.line.solidFill = _COLOR_LINE_2
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True
    chart.y_axis.numFmt = '0%'
    ws.add_chart(chart, "A9")


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def build_presentation_data(opiu_path, bdr_path, forecast_path,
                            year, month, output_path):
    months = list(range(1, month + 1))

    rows, month_col = _read_flat_rows(opiu_path, months)
    if not rows:
        print("[presentation_builder] нет данных — файл пустой")
        return output_path

    units = _collect_units(rows, months)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for unit_name in _UNITS:
        u = units.get(unit_name)
        if not u:
            continue
        _sheet_revenue_profit(wb, unit_name, u, months)
        _sheet_ebitda(wb, unit_name, u, months)
        _sheet_fot(wb, unit_name, u, months)

    if "Latvia" in units:
        latvia = units["Latvia"]
        for obj_full, obj_short in _LATVIA_OBJECTS:
            obj = latvia["objects"].get(obj_full)
            if not obj:
                continue
            _sheet_revenue_profit(wb, obj_short, obj, months)
            _sheet_ebitda(wb, obj_short, obj, months)
            _sheet_fot(wb, obj_short, obj, months)

        commercial = latvia["objects"].get("Коммерческие помещения LV")
        if commercial:
            _sheet_revenue_profit(wb, "Коммерческие", commercial, months)
            _sheet_ebitda(wb, "Коммерческие", commercial, months)
            _sheet_fot(wb, "Коммерческие", commercial, months)

    wb.save(output_path)
    print(f"[OK] Диаграммы сохранены: {output_path}")
    return output_path