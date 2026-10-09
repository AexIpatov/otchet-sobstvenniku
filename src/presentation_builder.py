# src/presentation_builder.py
# ============================================================
# Генератор Excel с диаграммами для презентации.
#
# Читает ОПиУ формата «статьи × месяцы» с древовидной структурой
# (·, ··, ···, ····) и строит книгу Excel:
#   - на каждом листе — один слайд презентации:
#       * таблица данных
#       * живой график (openpyxl.chart)
#
# Типы листов:
#   ОПиУ_<Юнит>       — столбики: Выручка (жёлтый) + ЧП (зелёный)
#   EBITDA_<Юнит>     — линия: EBITDA margin, %
#   Доля_ФОТ_<Юнит>   — линия: доля ФОТ в выручке, %
#   ОПиУ_<Объект>     — столбики по объектам Латвии
#   EBITDA_<Объект>   — линия по объектам Латвии
#   Доля_ФОТ_<Объект> — линия по объектам Латвии
#
# Поддерживает любой год и месяц (не только 2026-09).
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

# Юниты (уровень 1, «· <Юнит>»)
_UNITS = [
    "Latvia",
    "East-Восток",
    "Europe",
    "Nomiqa",
    "Unelma",
    "UK Estate",
]

# Объекты Латвии (уровень 2, «·· <Объект>»), для отдельных листов
_LATVIA_OBJECTS = [
    ("AN14 Антониас 14 (дом + парковка)", "Антонияс"),
    ("AC89 Чака 89 (дом + парковка)",     "Чака"),
    ("M81 - Matisa 81",                    "Матиса"),
    ("EKS_Esporta iela 12-113",            "Эспорта"),
]

# Объекты, которые собираем в один «виртуальный» блок «Коммерческие»
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

# Цвета оформления
_COLOR_REVENUE = "FFC000"   # жёлтый
_COLOR_NET     = "70AD47"   # зелёный
_COLOR_LINE_1  = "C00000"   # красный
_COLOR_LINE_2  = "4472C4"   # синий


# ============================================================
# ВСПОМОГАТЕЛЬНОЕ
# ============================================================

def _clean_title(s):
    if s is None:
        return ""
    s = str(s).replace("\u00a0", " ")
    s = re.sub(r"^[\s·]+", "", s)
    s = re.sub(r"[\s·]+$", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _level(s):
    if s is None:
        return 0
    m = re.match(r"^[\s·]*", str(s))
    return m.group(0).count("·") if m else 0


def _cell_float(ws, r, c):
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


def _safe_sheet_name(name):
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31]


def _month_to_col(ws, header_row, target_month, max_col=20):
    """Возвращает индекс колонки (1-based) для нужного месяца.
    Ищет по названию («Январь 2026», «сентябрь 2026», «Sep 2026» и т.п.).
    """
    target_ru = _MONTHS_RU_LOWER.get(target_month, "").lower()
    target_en = ["jan", "feb", "mar", "apr", "may", "jun",
                 "jul", "aug", "sep", "oct", "nov", "dec"][target_month - 1]

    for c in range(2, min(ws.max_column, max_col) + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        if isinstance(v, _dt.datetime) or isinstance(v, _dt.date):
            if v.month == target_month:
                return c
            continue
        sv = str(v).lower()
        if target_ru and target_ru in sv:
            return c
        if target_en in sv:
            return c
    return None


# ============================================================
# ГЛАВНЫЙ ПАРСЕР ОПиУ
# ============================================================

def _parse_opiu(opiu_path, months):
    """
    Читает ОПиУ формата «статьи × месяцы» с древовидной структурой.

    Возвращает:
    {
        "units": {
            "Latvia": {
                "revenue": {m: value},
                "net":     {m: value},
                "fot":     {m: value},   # ФОТ = сумма по всем объектам
                "objects": {
                    "AN14 ...": {
                        "revenue": {m: value},
                        "net":     {m: value},
                        "fot":     {m: value},
                    },
                    ...
                }
            },
            ...
        }
    }
    """
    if not opiu_path or not os.path.exists(opiu_path):
        return {"units": {}}

    wb = openpyxl.load_workbook(opiu_path, data_only=True)
    ws = wb.active

    # --- 1. Шапка: ищем строку с месяцами ---
    header_row = None
    for r in range(1, 6):
        found = 0
        for c in range(2, min(ws.max_column, 20) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            if isinstance(v, (_dt.datetime, _dt.date)):
                found += 1
                continue
            sv = str(v).lower()
            if any(m in sv for m in _MONTHS_RU_LOWER.values()):
                found += 1
        if found >= 3:
            header_row = r
            break

    if header_row is None:
        print(f"[WARN] [presentation_builder] не найдена шапка в {opiu_path}")
        return {"units": {}}

    # --- 2. Карта колонок месяцев ---
    month_col = {}
    for m in months:
        col = _month_to_col(ws, header_row, m)
        if col:
            month_col[m] = col

    if not month_col:
        print("[WARN] не найдены колонки месяцев")
        return {"units": {}}

    # --- 3. Один проход: собираем данные с учётом уровней ---
    units = {u: {
        "revenue": {m: 0.0 for m in months},
        "net":     {m: 0.0 for m in months},
        "fot":     {m: 0.0 for m in months},
        "objects": {},
    } for u in _UNITS}

    # Текущий контекст
    cur_section = None       # "Выручка" / "Производственные расходы" / ... / "Чистая прибыль"
    cur_unit = None          # имя юнита (совпадает с _UNITS)
    cur_object = None        # имя объекта (уровень 2, «·· »)
    cur_object_is_fot = False  # мы сейчас внутри строки «ФОТ производственного/коммерческого»

    for r in range(header_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue

        lvl = _level(raw)
        title = _clean_title(raw)

        # --- Уровень 0: секция ---
        if lvl == 0:
            cur_section = title
            cur_unit = None
            cur_object = None
            cur_object_is_fot = False
            continue

        # --- Уровень 1: юнит ---
        if lvl == 1:
            cur_unit = title if title in _UNITS else None
            cur_object = None
            cur_object_is_fot = False
            continue

        # --- Уровень 2: объект ---
        if lvl == 2:
            cur_object = title
            # заводим слот объекта (только для юнитов, которые нас интересуют)
            if cur_unit:
                if cur_object not in units[cur_unit]["objects"]:
                    units[cur_unit]["objects"][cur_object] = {
                        "revenue": {m: 0.0 for m in months},
                        "net":     {m: 0.0 for m in months},
                        "fot":     {m: 0.0 for m in months},
                    }
            continue

        # --- Уровень 3: статья. Может быть ФОТ? ---
        # Уровень 3 — это подстатья. Нам она как «··· ФОТ производственного персонала» не встречается,
        # но на всякий случай проверяем.
        if lvl == 3:
            title_low = title.lower()
            if "фот производственного" in title_low or \
               "фот коммерческого" in title_low:
                # Это ФОТ объекта. Читаем значения.
                if cur_unit and cur_object:
                    for m, col in month_col.items():
                        units[cur_unit]["objects"][cur_object]["fot"][m] += _cell_float(ws, r, col)
            continue

        # --- Уровень 4: ФОТ (главный случай) ---
        if lvl >= 4:
            title_low = title.lower()
            is_fot = ("фот производственного" in title_low or
                      "фот коммерческого" in title_low)
            if is_fot and cur_unit and cur_object:
                for m, col in month_col.items():
                    units[cur_unit]["objects"][cur_object]["fot"][m] += _cell_float(ws, r, col)
            continue

    # --- 4. Второй проход: собираем revenue и net по юнитам и объектам ---
    cur_section = None
    cur_unit = None
    cur_object = None

    for r in range(header_row + 1, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue

        lvl = _level(raw)
        title = _clean_title(raw)

        if lvl == 0:
            cur_section = title
            cur_unit = None
            cur_object = None
            continue

        if lvl == 1:
            cur_unit = title if title in _UNITS else None
            cur_object = None
            # Если это раздел «Выручка» или «Чистая прибыль», пишем в юнит
            if cur_unit and cur_section in ("Выручка", "Чистая прибыль"):
                key = "revenue" if cur_section == "Выручка" else "net"
                for m, col in month_col.items():
                    units[cur_unit][key][m] = _cell_float(ws, r, col)
            continue

        if lvl == 2:
            cur_object = title
            if cur_unit and cur_section in ("Выручка", "Чистая прибыль"):
                key = "revenue" if cur_section == "Выручка" else "net"
                if cur_object in units[cur_unit]["objects"]:
                    for m, col in month_col.items():
                        units[cur_unit]["objects"][cur_object][key][m] = _cell_float(ws, r, col)
            continue

        # Уровень 3+ для revenue/net не нужен — там только детализация

    # --- 5. Пересчёт ФОТ на уровне юнита: сумма по всем объектам ---
    for u_name, u in units.items():
        for m in months:
            u["fot"][m] = sum(
                obj["fot"][m] for obj in u["objects"].values()
            )

    # --- 6. Виртуальный юнит «Коммерческие» (для Латвии) ---
    if "Latvia" in units:
        latvia = units["Latvia"]
        commercial = {
            "revenue": {m: 0.0 for m in months},
            "net":     {m: 0.0 for m in months},
            "fot":     {m: 0.0 for m in months},
        }
        for obj_name in _LATVIA_COMMERCIAL:
            obj = latvia["objects"].get(obj_name)
            if not obj:
                continue
            for m in months:
                commercial["revenue"][m] += obj["revenue"][m]
                commercial["net"][m]     += obj["net"][m]
                commercial["fot"][m]     += obj["fot"][m]
        latvia["objects"]["Коммерческие помещения LV"] = commercial

    return {"units": units}


# ============================================================
# ПОСТРОЕНИЕ ЛИСТОВ
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
    return start_row + len(rows)


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


# ------------------------------------------------------------
# Лист: ОПиУ юнита/объекта (Выручка + ЧП)
# ------------------------------------------------------------
def _sheet_revenue_profit(wb, name, data, months):
    sheet_name = _safe_sheet_name(f"ОПиУ_{name}")
    ws = wb.create_sheet(sheet_name) if sheet_name not in wb.sheetnames else wb[sheet_name]

    _write_header(ws, f"ОПиУ {name}", "Выручка и чистая прибыль по месяцам")

    headers = ["Показатель"] + [_MONTHS_RU_CAP[m] for m in months]
    rows = [
        ["Выручка"]         + [round(data["revenue"][m], 2) for m in months],
        ["Чистая прибыль"]  + [round(data["net"][m], 2)     for m in months],
    ]
    _write_table(ws, 4, headers, rows)

    # Диаграмма
    data_ref = Reference(ws, min_col=2, max_col=1 + len(months),
                         min_row=4, max_row=6)
    cats = Reference(ws, min_col=2, max_col=1 + len(months),
                     min_row=4, max_row=4)

    chart = BarChart()
    chart.type = "col"
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    _style_bar(chart, f"Выручка и ЧП — {name}")
    chart.series[0].graphicalProperties.solidFill = _COLOR_REVENUE
    chart.series[1].graphicalProperties.solidFill = _COLOR_NET
    ws.add_chart(chart, "A9")


# ------------------------------------------------------------
# Лист: EBITDA margin (линия)
# ------------------------------------------------------------
def _sheet_ebitda(wb, name, data, months):
    sheet_name = _safe_sheet_name(f"EBITDA_{name}")
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

    data_ref = Reference(ws, min_col=2, max_col=1 + len(months),
                         min_row=4, max_row=5)
    cats = Reference(ws, min_col=2, max_col=1 + len(months),
                     min_row=4, max_row=4)

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    _style_line(chart, f"EBITDA margin — {name}")
    chart.series[0].graphicalProperties.line.solidFill = _COLOR_LINE_1
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True
    chart.y_axis.numFmt = '0%'
    ws.add_chart(chart, "A9")


# ------------------------------------------------------------
# Лист: Доля ФОТ (линия)
# ------------------------------------------------------------
def _sheet_fot(wb, name, data, months):
    sheet_name = _safe_sheet_name(f"Доля_ФОТ_{name}")
    ws = wb.create_sheet(sheet_name) if sheet_name not in wb.sheetnames else wb[sheet_name]

    _write_header(ws, f"Доля ФОТ в выручке — {name}", "%")

    shares = {}
    for m in months:
        rev = data["revenue"][m]
        fot = data["fot"][m]
        shares[m] = round(fot / rev, 4) if rev else 0.0

    headers = ["Показатель"] + [_MONTHS_RU_CAP[m] for m in months]
    rows = [["Доля ФОТ"] + [shares[m] for m in months]]
    _write_table(ws, 4, headers, rows)

    for c in range(2, 2 + len(months)):
        ws.cell(row=5, column=c).number_format = '0.00%'

    data_ref = Reference(ws, min_col=2, max_col=1 + len(months),
                         min_row=4, max_row=5)
    cats = Reference(ws, min_col=2, max_col=1 + len(months),
                     min_row=4, max_row=4)

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
    """
    Собирает Excel-файл с диаграммами для презентации.

    opiu_path     — путь к ОПиУ 01.{year}-{month}.{year}.xlsx
    bdr_path      — путь к БДиР (пока не используется)
    forecast_path — путь к Прогнозам (пока не используется)
    year, month   — отчётный год и месяц
    output_path   — куда сохранить итоговый Excel
    """
    months = list(range(1, month + 1))

    parsed = _parse_opiu(opiu_path, months)
    units = parsed["units"]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for unit_name in _UNITS:
        u = units.get(unit_name)
        if not u:
            continue

        _sheet_revenue_profit(wb, unit_name, u, months)
        _sheet_ebitda(wb, unit_name, u, months)
        _sheet_fot(wb, unit_name, u, months)

    # Отдельные листы по объектам Латвии
    if "Latvia" in units:
        latvia = units["Latvia"]
        for obj_full_name, obj_short in _LATVIA_OBJECTS:
            obj = latvia["objects"].get(obj_full_name)
            if not obj:
                continue
            _sheet_revenue_profit(wb, obj_short, obj, months)
            _sheet_ebitda(wb, obj_short, obj, months)
            _sheet_fot(wb, obj_short, obj, months)

        # Блок «Коммерческие»
        commercial = latvia["objects"].get("Коммерческие помещения LV")
        if commercial:
            _sheet_revenue_profit(wb, "Коммерческие", commercial, months)
            _sheet_ebitda(wb, "Коммерческие", commercial, months)
            _sheet_fot(wb, "Коммерческие", commercial, months)

    wb.save(output_path)
    print(f"[OK] Диаграммы сохранены: {output_path}")
    return output_path