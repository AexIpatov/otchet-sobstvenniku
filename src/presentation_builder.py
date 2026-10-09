# src/presentation_builder.py
# ============================================================
# Генератор Excel с диаграммами для презентации.
#
# Читает ОПиУ + БДиР и строит книгу Excel:
#   - на каждом листе — один слайд презентации:
#       * таблица данных
#       * живой график (openpyxl.chart)
#
# Поддерживает любой год и месяц (не только 2026-09).
# ============================================================

import os
import re
import calendar
from datetime import datetime

import openpyxl
from openpyxl.chart import (
    BarChart, LineChart, Reference, Series,
)
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import config


# ============================================================
# ВСПОМОГАТЕЛЬНОЕ
# ============================================================

_MONTHS_RU = {
    1: "январь", 2: "февраль", 3: "март", 4: "апрель",
    5: "май", 6: "июнь", 7: "июль", 8: "август",
    9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь",
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


def _safe_sheet_name(name):
    """Делает имя листа валидным для Excel (до 31 символа)."""
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31]


def _read_monthly_series(path, object_cols, row_keywords,
                        months_to_read, sheet=None):
    """
    Читает ОПиУ-файл, где данные лежат «объекты × месяцы».
    object_cols — список объектов, которые нужно суммировать (по вхождению).
    row_keywords — список ключей, одно из которых должно быть в строке.
                   Если несколько — берётся ПЕРВОЕ совпадение (по приоритету).
    Возвращает список значений по месяцам months_to_read.

    ВАЖНО: месяц определяется по позиции колонки (C=январь, D=февраль, ...).
    """
    result = {m: 0.0 for m in months_to_read}

    if not path or not os.path.exists(path):
        return [0.0 for _ in months_to_read]

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if (sheet and sheet in wb.sheetnames) else wb.active

    # Ищем строку с названиями месяцев в первых 5 строках
    # (в новом ОПиУ формат «статьи × месяцы»).
    header_row = None
    for r in range(1, 6):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, str) and "январ" in v.lower():
            header_row = r
            break

    if header_row is None:
        return [0.0 for _ in months_to_read]

    # Определяем колонку для каждого месяца по шапке
    month_col = {}
    for m in months_to_read:
        target = _MONTHS_RU[m].lower()
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            if v and target in str(v).lower():
                month_col[m] = c
                break

    # Ищем строки, подходящие под row_keywords, в первых 2 колонках
    for r in range(header_row + 1, ws.max_row + 1):
        v1 = ws.cell(row=r, column=1).value
        v2 = ws.cell(row=r, column=2).value

        for v in (v1, v2):
            if v is None:
                continue
            sv = str(v).lower()
            for kw in row_keywords:
                if kw.lower() in sv:
                    # Нашли нужную строку — собираем значения по месяцам
                    for m in months_to_read:
                        col = month_col.get(m)
                        if col:
                            val = ws.cell(row=r, column=col).value
                            try:
                                result[m] += float(val) if val is not None else 0.0
                            except (ValueError, TypeError):
                                pass
                    # Прерываемся на первом совпавшем keyword
                    return [result[m] for m in months_to_read]

    return [result[m] for m in months_to_read]


def _read_opiu_by_object(path, object_name_part, months):
    """
    Читает ОПиУ формата «объекты × месяцы» (первая строка — шапка с объектами,
    столбец A — название статьи).
    object_name_part — часть имени объекта (например, "AN14").
    Возвращает dict: {месяц: {'выручка': x, 'чистая': y, 'ebitda': z, 'фот': w}}
    """
    result = {m: {"выручка": 0.0, "чистая": 0.0, "ebitda": 0.0, "фот": 0.0}
              for m in months}

    if not path or not os.path.exists(path):
        return result

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    # В новом ОПиУ формат «месяц × показатели» (см. приложенный файл):
    # строка 1 — шапка (январь 2026, февраль 2026, ...),
    # столбец A — название статьи/объекта.
    # Ищем нужную строку по ключевому слову и суммируем по месяцам.

    target = object_name_part.lower()
    # Ищем заголовок
    header_row = None
    for r in range(1, 6):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, str) and "январ" in v.lower():
            header_row = r
            break
    if header_row is None:
        return result

    # Колонки месяцев
    month_col = {}
    for m in months:
        tgt = _MONTHS_RU[m].lower()
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            if v and tgt in str(v).lower():
                month_col[m] = c
                break

    # Ищем строку с нужным объектом
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        if target in str(v).lower():
            for m in months:
                col = month_col.get(m)
                if col:
                    val = ws.cell(row=r, column=col).value
                    try:
                        result[m]["выручка"] = float(val) if val is not None else 0.0
                    except (ValueError, TypeError):
                        result[m]["выручка"] = 0.0
            break

    return result


# ============================================================
# ЗАГРУЗКА ДАННЫХ ПО ЮНИТАМ
# ============================================================

def _load_unit_data(opiu_path, months):
    """
    Загружает данные для всех юнитов из одного файла ОПиУ
    формата «месяц × показатели».

    Возвращает dict: {unit_name: {metric: [values_by_month]}}
    """
    units = {
        "Латвия":        {"row": "· Latvia",     "type": "section"},
        "Антонияс":      {"row": "AN14",         "type": "object"},
        "Чака":          {"row": "AC89",         "type": "object"},
        "Матиса":        {"row": "M81",          "type": "object"},
        "Эспорта":       {"row": "EKS_Esporta",  "type": "object"},
        "Estate-Восток": {"row": "· East-Восток","type": "section"},
        "Европа":        {"row": "· Europe",     "type": "section"},
        "Nomiqa":        {"row": "· Nomiqa",     "type": "section"},
        "Унелма":        {"row": "· Unelma",     "type": "section"},
        "UK Estate":     {"row": "· UK Estate",  "type": "section"},
    }

    # Все данные читаем из ОПиУ 09.2026 (там все месяцы 01-09).
    # Формат: столбец A — статья/объект, столбцы B..K — месяцы 01-09.
    if not opiu_path or not os.path.exists(opiu_path):
        return {}

    wb = openpyxl.load_workbook(opiu_path, data_only=True)
    ws = wb.active

    # Определяем шапку (строка 1)
    header_row = 1

    # Колонки месяцев
    month_col = {}
    for m in months:
        tgt = _MONTHS_RU[m].lower()
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            if v and tgt in str(v).lower():
                month_col[m] = c
                break

    # Собираем все строки в список (row_idx, title_lower)
    all_rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        all_rows.append((r, str(v).strip()))

    out = {}
    for unit_name, spec in units.items():
        key = spec["row"].lower()
        unit_data = {
            "выручка":  [0.0] * len(months),
            "ЧП":       [0.0] * len(months),
            "EBITDA":   [0.0] * len(months),
            "ФОТ":      [0.0] * len(months),
            "Расходы":  [0.0] * len(months),
        }

        # Ищем строку, где title == key (section) или начинается с key (object)
        for r, title in all_rows:
            t_low = title.lower()
            if spec["type"] == "section":
                # Точное совпадение по префиксу «· »
                if t_low == key or t_low.startswith(key + " ") or t_low.startswith(key + "\t"):
                    # для секций берём значения из этой строки
                    for i, m in enumerate(months):
                        col = month_col.get(m)
                        if col:
                            val = ws.cell(row=r, column=col).value
                            try:
                                unit_data["выручка"][i] = float(val) if val is not None else 0.0
                            except (ValueError, TypeError):
                                pass
                    break
            else:
                if key in t_low:
                    for i, m in enumerate(months):
                        col = month_col.get(m)
                        if col:
                            val = ws.cell(row=r, column=col).value
                            try:
                                unit_data["выручка"][i] = float(val) if val is not None else 0.0
                            except (ValueError, TypeError):
                                pass
                    break

        out[unit_name] = unit_data

    return out


# ============================================================
# ПОСТРОЕНИЕ ЛИСТОВ
# ============================================================

def _write_header(ws, title, subtitle=""):
    ws.cell(row=1, column=1, value=title).font = _TITLE_FONT
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = _SUBTITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=14)


def _write_table(ws, row, headers, rows):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _CELL_BORDER
    for r_off, row_vals in enumerate(rows, start=1):
        for c, v in enumerate(row_vals, start=1):
            cell = ws.cell(row=row + r_off, column=c, value=v)
            cell.border = _CELL_BORDER
            if isinstance(v, (int, float)):
                cell.number_format = '#,##0'
    return row + len(rows)


def _style_bar_chart(chart, title, w=18, h=10):
    chart.title = title
    chart.width = w
    chart.height = h
    chart.style = 10
    chart.y_axis.majorGridlines = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False


def _style_line_chart(chart, title, w=18, h=10):
    chart.title = title
    chart.width = w
    chart.height = h
    chart.style = 12
    chart.y_axis.majorGridlines = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False


# ------------------------------------------------------------
# ЛИСТ 1. ОПиУ юнита: Выручка + ЧП (двойной столбик)
# ------------------------------------------------------------
def _sheet_opiu_revenue_profit(wb, unit_name, unit_data, months):
    sheet_name = _safe_sheet_name(f"ОПиУ_{unit_name}")
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)

    _write_header(ws,
                  f"ОПиУ {unit_name}",
                  "Выручка & Чистая прибыль, без НДС")

    # Таблица: Заголовки — месяцы, строки — выручка и ЧП
    header_row = 4
    headers = ["Показатель"] + [_MONTHS_RU[m].capitalize() for m in months]
    rows = [
        ["Выручка"] + unit_data["выручка"],
        ["Чистая прибыль"] + unit_data["ЧП"],
    ]
    _write_table(ws, header_row, headers, rows)

    # Данные для графика
    data_min_col = 2
    data_max_col = 1 + len(months)
    chart_data = Reference(ws,
                           min_col=data_min_col, max_col=data_max_col,
                           min_row=header_row, max_row=header_row + 2)
    cats = Reference(ws,
                     min_col=data_min_col, max_col=data_max_col,
                     min_row=header_row, max_row=header_row)

    chart = BarChart()
    chart.type = "col"
    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(cats)
    _style_bar_chart(chart, f"Выручка и ЧП {unit_name}")

    # Цвета: выручка — жёлтый, ЧП — зелёный
    chart.series[0].graphicalProperties.solidFill = "FFC000"
    chart.series[1].graphicalProperties.solidFill = "70AD47"

    ws.add_chart(chart, "A10")
    return ws


# ------------------------------------------------------------
# ЛИСТ 2. EBITDA margin (линия %)
# ------------------------------------------------------------
def _sheet_ebitda(wb, unit_name, unit_data, months):
    sheet_name = _safe_sheet_name(f"EBITDA_{unit_name}")
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)

    _write_header(ws,
                  f"Операционная рентабельность {unit_name}",
                  "EBITDA margin, %")

    # Считаем EBITDA = ЧП / Выручка
    ebitda = []
    for i in range(len(months)):
        v = unit_data["выручка"][i]
        c = unit_data["ЧП"][i]
        ebitda.append(round(c / v, 4) if v else 0.0)

    header_row = 4
    headers = ["Показатель"] + [_MONTHS_RU[m].capitalize() for m in months]
    rows = [["EBITDA margin"] + ebitda]
    _write_table(ws, header_row, headers, rows)

    # Формат ячеек — процент
    for c in range(2, 2 + len(months)):
        ws.cell(row=header_row + 1, column=c).number_format = '0.00%'

    data_ref = Reference(ws,
                         min_col=2, max_col=1 + len(months),
                         min_row=header_row, max_row=header_row + 1)
    cats = Reference(ws,
                     min_col=2, max_col=1 + len(months),
                     min_row=header_row, max_row=header_row)

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    _style_line_chart(chart, f"EBITDA margin {unit_name}")

    chart.series[0].graphicalProperties.line.solidFill = "FFC000"
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True
    chart.y_axis.numFmt = '0%'

    ws.add_chart(chart, "A10")
    return ws


# ------------------------------------------------------------
# ЛИСТ 3. Доля ФОТ (линия %)
# ------------------------------------------------------------
def _sheet_fot_share(wb, unit_name, unit_data, months):
    sheet_name = _safe_sheet_name(f"Доля_ФОТ_{unit_name}")
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)

    _write_header(ws,
                  f"Доля ФОТ в выручке {unit_name}",
                  "%")

    share = []
    for i in range(len(months)):
        v = unit_data["выручка"][i]
        f = unit_data["ФОТ"][i]
        share.append(round(f / v, 4) if v else 0.0)

    header_row = 4
    headers = ["Показатель"] + [_MONTHS_RU[m].capitalize() for m in months]
    rows = [["Доля ФОТ"] + share]
    _write_table(ws, header_row, headers, rows)

    for c in range(2, 2 + len(months)):
        ws.cell(row=header_row + 1, column=c).number_format = '0.00%'

    data_ref = Reference(ws,
                         min_col=2, max_col=1 + len(months),
                         min_row=header_row, max_row=header_row + 1)
    cats = Reference(ws,
                     min_col=2, max_col=1 + len(months),
                     min_row=header_row, max_row=header_row)

    chart = LineChart()
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    _style_line_chart(chart, f"Доля ФОТ {unit_name}")

    chart.series[0].graphicalProperties.line.solidFill = "C00000"
    chart.series[0].graphicalProperties.line.width = 25000
    chart.series[0].smooth = True
    chart.y_axis.numFmt = '0%'

    ws.add_chart(chart, "A10")
    return ws


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ СБОРКИ
# ============================================================

def build_presentation_data(opiu_path, bdr_path, forecast_path,
                            year, month, output_path):
    """
    Собирает Excel-файл с диаграммами для презентации.

    opiu_path     — путь к ОПиУ 01.{year}-{month}.{year}.xlsx
    bdr_path      — путь к БДиР 01.{year}-12.{year}.xlsx (пока не используется)
    forecast_path — путь к Прогнозам (пока не используется)
    year, month   — отчётный год и месяц
    output_path   — куда сохранить итоговый Excel
    """
    months = list(range(1, month + 1))   # 1..9 для сентября

    # Загружаем данные
    units_data = _load_unit_data(opiu_path, months)

    # Создаём книгу
    wb = openpyxl.Workbook()
    # Убираем дефолтный лист
    default = wb.active
    wb.remove(default)

    # Строим листы по каждому юниту
    for unit_name, unit_data in units_data.items():
        _sheet_opiu_revenue_profit(wb, unit_name, unit_data, months)
        _sheet_ebitda(wb, unit_name, unit_data, months)
        _sheet_fot_share(wb, unit_name, unit_data, months)

    # Сохраняем
    wb.save(output_path)
    print(f"[OK] Презентация данных сохранена: {output_path}")
    return output_path