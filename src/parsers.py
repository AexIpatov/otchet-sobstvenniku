# src/parsers.py
# ============================================================
# Парсеры исходных файлов.
# Импортируем сам модуль config, чтобы видеть актуальные
# значения после config.set_period(...).
# ============================================================

import os
import re
import openpyxl

import config


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

def _clean_title(title):
    if title is None:
        return ""
    s = str(title)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"^[\s·]+", "", s)
    s = re.sub(r"[\s·]+$", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _get_level(title):
    if title is None:
        return 0
    s = str(title)
    m = re.match(r"^[\s·]*", s)
    prefix = m.group(0) if m else ""
    return prefix.count("·")


def _normalize(name):
    if name is None:
        return ""
    s = _clean_title(name).lower()
    s = s.replace("-", " ").replace(".", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _cell_value(ws, row, col_idx):
    v = ws.cell(row=row, column=col_idx + 1).value
    if v is None:
        return 0.0
    if isinstance(v, str):
        v = v.strip().replace("\u00a0", "").replace(" ", "")
        if v == "" or v == "-":
            return 0.0
        v = v.replace(",", ".")
        try:
            return float(v)
        except ValueError:
            return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ---------------------------------------------------------------------------
# Маппинг «сырых» статей на строки шаблона
# ---------------------------------------------------------------------------

def _map_to_template(raw, row_map):
    norm_raw = {}
    for k, v in raw.items():
        nk = _normalize(k)
        norm_raw[nk] = norm_raw.get(nk, 0.0) + v

    result = {}

    for tpl_row, rule in row_map.items():
        if isinstance(rule, dict):
            sum_list      = rule.get("sum", []) or []
            any_list      = rule.get("any", []) or []
            fallback_list = rule.get("fallback", []) or []
        else:
            sum_list      = []
            any_list      = list(rule) if rule else []
            fallback_list = []

        # 2) any: первое ненулевое совпадение
        chosen_any = None
        for alias in any_list:
            na = _normalize(alias)
            if na in norm_raw:
                v = norm_raw[na]
                if v != 0.0:
                    chosen_any = v
                    break
                if chosen_any is None:
                    chosen_any = 0.0

        # 3) sum: складываем ВСЕ совпадения
        total_sum = 0.0
        sum_hits = 0
        for alias in sum_list:
            na = _normalize(alias)
            if na in norm_raw:
                total_sum += norm_raw[na]
                sum_hits += 1

        # 4) fallback: используется ТОЛЬКО если ни any, ни sum не дали результата
        chosen_fb = None
        if chosen_any is None and sum_hits == 0:
            for alias in fallback_list:
                na = _normalize(alias)
                if na in norm_raw:
                    v = norm_raw[na]
                    if v != 0.0:
                        chosen_fb = v
                        break
                    if chosen_fb is None:
                        chosen_fb = 0.0

        # 5) Итог
        if chosen_any is not None and chosen_any != 0.0:
            final = chosen_any
        elif sum_hits > 0:
            final = total_sum
        elif chosen_any is not None:
            final = chosen_any
        elif chosen_fb is not None:
            final = chosen_fb
        else:
            continue

        result[tpl_row] = final

    return result


def _require_file(path, kind):
    if not os.path.exists(path):
        msg = f"[ERROR] Не найден файл {kind}: {path}. Загрузите и повторите."
        print(msg)
        raise FileNotFoundError(msg)


# ---------------------------------------------------------------------------
# ОПиУ
# ---------------------------------------------------------------------------

def parse_opiu(path, object_names, row_map, month_name_ru=None):
    """
    Читает файл ОПиУ. Поддерживает ДВА формата:

    Формат A (объекты × статьи):
      - в шапке (обычно строка 1 или 2) перечислены имена объектов
        (например, "AN14 Антониас 14 (дом + парковка) [Latvia]");
      - столбец A — наименования статей;
      - данные — в столбце нужного объекта.
      Так устроены ОПиУ Латвия, Европа, Estate AZE, Nomiqa, Unelma.

    Формат B (статьи × месяцы):
      - в шапке — названия месяцев ("Август 2026", "Сентябрь 2026");
      - столбец A — № п/п, столбец B — наименование статьи;
      - данные — в столбце нужного месяца.
      Так устроен ОПиУ UK Estate.

    Формат определяется автоматически. Если format A — используется
    object_names. Если format B — используется month_name_ru.
    """
    _require_file(path, "ОПиУ")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    # ---------- 1. Определяем формат ----------
    _month_re = re.compile(
        r"("
        r"январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр"
        r"|"
        r"jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec"
        r")",
        re.IGNORECASE,
    )

    header_row = None
    header_format = None   # "A" или "B"

    for r in range(1, min(ws.max_row, 10) + 1):
        hits_obj = 0    # ячейки вида "[...]" или "итого"
        hits_month = 0  # ячейки вида "Август 2026"
        for c in range(2, min(ws.max_column, 30) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            sv = str(v)
            if "[" in sv and "]" in sv:
                hits_obj += 1
            elif "итого" in sv.lower():
                hits_obj += 1
            elif _month_re.search(sv):
                hits_month += 1

        # Формат A: достаточно ОДНОГО объекта в шапке (со скобками
        # вида "[Latvia]", "[Unelma]" и т.п.) или ячейки "итого".
        # В файле ОПиУ Unelma в шапке только один объект
        # "UK_Unelma [Unelma]", и раньше парсер его не находил.
        if hits_obj >= 1:
            header_row = r
            header_format = "A"
            break

        # Формат B: достаточно ОДНОГО месяца в шапке.
        # В UK Estate в шапке только одна ячейка с месяцем
        # (например, "Август 2026"), потому что файл содержит
        # данные только за один месяц.
        if hits_month >= 1:
            header_row = r
            header_format = "B"
            break

    if header_row is None or header_format is None:
        # Не нашли шапку — не угадываем формат, возвращаем пусто.
        print(f"[WARN] [parse_opiu] {os.path.basename(path)}: "
              f"не удалось определить формат ОПиУ (шапка не найдена). "
              f"Возвращаем пусто.")
        return {}

    if config.DEBUG:
        print(f"[parse_opiu] {os.path.basename(path)}: "
              f"формат {header_format}, шапка в строке {header_row}")

    # ---------- 2. Формат A: объекты × статьи ----------
    if header_format == "A":
        # Карта объектов: {normalized_name: col_idx(0-based)}
        col_map = {}
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            if v is None:
                continue
            name = str(v).strip()
            if name == "" or name.lower() == "итого":
                continue
            col_map[_normalize(name)] = c - 1

        if config.DEBUG:
            print(f"[parse_opiu] Найдено объектов в шапке: {len(col_map)}")

        # Определяем нужные колонки по object_names.
        wanted_cols = []
        for obj in (object_names or []):
            no = _normalize(obj)
            if no in col_map:
                wanted_cols.append(col_map[no])
            else:
                if config.DEBUG:
                    print(f"[parse_opiu] Объект не найден в шапке: {obj!r}")

        if not wanted_cols:
            if config.DEBUG:
                print(f"[parse_opiu] Ни один объект не найден. "
                      f"Возвращаем пусто.")
            return {}

        if config.DEBUG:
            print(f"[parse_opiu] Используем колонки (0-based): {wanted_cols}")

        # Собираем raw: суммируем значения по имени статьи,
        # НО НЕ суммируем родительские и дочерние строки с одинаковым
        # именем (например, «1.2.16.1 Налог на недвижимость» —
        # родительская и дочерняя с тем же именем).
        #
        # Логика: если строка с таким именем уже была добавлена
        # РАНЬШЕ (то есть это родительская строка), дочернюю
        # игнорируем. Если строка с таким именем встречается
        # несколько раз подряд на РАЗНЫХ уровнях — берём только
        # первую (родительскую).
        #
        # ВАЖНО: это НЕ ломает случай, когда две НЕЗАВИСИМЫЕ
        # строки имеют одинаковое имя (например, две «Зарплата»
        # под разными родителями). Такие строки НЕ являются
        # родитель-потомок, потому что у них нет родительской
        # строки с тем же именем. Их нужно суммировать.
        #
        # Как отличить: если строка с именем X уже встречалась
        # РАНЬШЕ на более высоком уровне (меньше `·`), то текущая —
        # дочерняя, и её нужно игнорировать. Если встречалась
        # на том же уровне или ниже — это независимая строка,
        # её нужно суммировать.
        raw = {}
        raw_first_level = {}  # {title: level первой встречи}
        for r in range(header_row + 1, ws.max_row + 1):
            title_cell = ws.cell(row=r, column=1).value
            if title_cell is None:
                continue
            title = _clean_title(title_cell)
            if title == "":
                continue

            total = 0.0
            for col_idx in wanted_cols:
                total += _cell_value(ws, r, col_idx)

            if total == 0.0:
                continue

            current_level = _get_level(title_cell)

            if title in raw:
                prev_level = raw_first_level[title]
                if current_level > prev_level:
                    # Текущая строка глубже — это дочерняя,
                    # игнорируем (родительская уже содержит агрегат).
                    if config.DEBUG:
                        print(f"[parse_opiu] пропускаю дочернюю строку "
                              f"{title!r} (level {current_level} > {prev_level})")
                    continue
                elif current_level < prev_level:
                    # Текущая строка ВЫШЕ — значит, ранее мы взяли
                    # дочернюю, а теперь нашли родительскую.
                    # Заменяем значение родительским (НЕ суммируем).
                    if config.DEBUG:
                        print(f"[parse_opiu] заменяю дочернюю строку "
                              f"{title!r} родительской "
                              f"(level {current_level} < {prev_level})")
                    raw[title] = total
                    raw_first_level[title] = current_level
                else:
                    # Одинаковый уровень — независимые строки, суммируем.
                    raw[title] = raw[title] + total
            else:
                raw[title] = total
                raw_first_level[title] = current_level

        if config.DEBUG:
            print(f"[parse_opiu] Собрано {len(raw)} статей.")

        return _map_to_template(raw, row_map)

    # ---------- 3. Формат B: статьи × месяцы ----------
    if header_format == "B":
        if not month_name_ru:
            # Месяц не передан — ищем первый столбец с данными.
            # Но лучше — явно требовать month_name_ru.
            print(f"[WARN] [parse_opiu] Формат B, но не передан "
                  f"month_name_ru. Возвращаем пусто.")
            return {}

        target_norm = _normalize(month_name_ru)
        wanted_col = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            if v is None:
                continue
            if target_norm in _normalize(v):
                wanted_col = c
                break

        if wanted_col is None:
            print(f"[WARN] [parse_opiu] {os.path.basename(path)}: "
                  f"столбец для месяца {month_name_ru!r} не найден. "
                  f"Доступные значения в шапке: "
                  f"{[ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column+1)]}")
            return {}

        if config.DEBUG:
            print(f"[parse_opiu] {os.path.basename(path)}: "
                  f"месяц {month_name_ru!r} → столбец {wanted_col}")

        # Определяем столбец с наименованием статьи (A, B или C).
        # Ищем тот, где минимум 3 непустые строки подряд в первых
        # 20 строках под шапкой.
        title_col = None
        for c in (1, 2, 3):
            cnt = 0
            for r in range(header_row + 1, min(header_row + 30, ws.max_row) + 1):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip() != "":
                    cnt += 1
            if cnt >= 3:
                title_col = c
                break

        if title_col is None:
            title_col = 2   # по умолчанию — B

        if config.DEBUG:
            print(f"[parse_opiu] {os.path.basename(path)}: "
                  f"столбец наименований = {title_col}")

        raw = {}
        for r in range(header_row + 1, ws.max_row + 1):
            title_cell = ws.cell(row=r, column=title_col).value
            if title_cell is None:
                continue
            title = _clean_title(title_cell)
            if title == "":
                continue
            val = _cell_value(ws, r, wanted_col - 1)   # 0-based
            if val == 0.0:
                continue
            if title not in raw:
                raw[title] = val

        if config.DEBUG:
            print(f"[parse_opiu] {os.path.basename(path)}: "
                  f"собрано статей {len(raw)} (формат B)")

        return _map_to_template(raw, row_map)

    return {}


# ---------------------------------------------------------------------------
# БДиР
# ---------------------------------------------------------------------------

def parse_bdr(path, object_names, month, row_map):
    _require_file(path, "БДиР")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    col_idx = config.BDR_MONTH_COL.get(month)
    if col_idx is None:
        fallback_month = getattr(config, "REPORT_MONTH", None)
        if fallback_month is not None:
            col_idx = config.BDR_MONTH_COL.get(fallback_month)
            if config.DEBUG:
                print(f"[parse_bdr] Для месяца {month} нет колонки — "
                      f"использую план на отчётный месяц {fallback_month}")
    if col_idx is None:
        if config.DEBUG:
            print(f"[parse_bdr] Для месяца {month} нет колонки в BDR_MONTH_COL")
        return {}

    wanted = {_normalize(o) for o in (object_names or [])}

    raw = {}
    current_local = None
    current_object_ok = False
    object_level = None
    found_objects = []

    for r in range(1, ws.max_row + 1):
        cell = ws.cell(row=r, column=1).value
        if cell is None:
            continue
        title = str(cell)
        level = _get_level(title)
        name = _clean_title(title)
        if name == "":
            continue

        name_norm = _normalize(name)

        # 1) Строка-объект из wanted
        if name_norm in wanted:
            # Игнорируем повторные вхождения того же объекта
            # (в БДиР Nomiqa имена BNQ_BAKU-Nomiqa и DNQ_Dubai-Nomiqa
            # повторяются на разных уровнях — level 2 и level 3).
            # Обрабатываем только ПЕРВОЕ вхождение (level 2).
            if name_norm in [n for _, _, n in found_objects]:
                if config.DEBUG:
                    print(f"[parse_bdr] пропускаю повторный объект: {name!r} (стр.{r}, level={level})")
                continue

            if current_local is not None:
                for k, v in current_local.items():
                    raw[k] = raw.get(k, 0.0) + v
            current_local = {}
            current_object_ok = True
            object_level = level
            found_objects.append((r, level, name))
            if config.DEBUG:
                print(f"[parse_bdr] найден объект: {name!r} (стр.{r}, level={level})")
            continue

        # 2) Внутри объекта
        if current_object_ok:
            if level <= object_level:
                if current_local is not None:
                    for k, v in current_local.items():
                        raw[k] = raw.get(k, 0.0) + v
                current_local = None
                current_object_ok = False
                object_level = None
                continue

            val = _cell_value(ws, r, col_idx)
            if val == 0.0:
                continue

            if name in current_local and current_local[name] != 0.0:
                continue
            current_local[name] = val

    # 3) закрываем последний объект
    if current_local is not None:
        for k, v in current_local.items():
            raw[k] = raw.get(k, 0.0) + v

    if config.DEBUG:
        print(f"[parse_bdr] {os.path.basename(path)}: "
              f"найдено объектов: {len(found_objects)}, "
              f"собрано статей: {len(raw)}, месяц={month}")
        # ДИАГНОСТИКА: показываем статьи, связанные с ФОТ и Зарплатой
        for k, v in raw.items():
            kn = _normalize(k)
            if "зарплат" in kn or "фот" in kn or "relat" in kn:
                print(f"[parse_bdr]   raw['{k}'] = {v}")

    return _map_to_template(raw, row_map)


# ---------------------------------------------------------------------------
# Прогнозы
# ---------------------------------------------------------------------------

def parse_forecast(path, sheet_name, month, row_map):
    """
    Читает файл «Прогнозы месячные».

    ВАЖНО: если запрошенный лист не найден, НЕ подставляем молча
    wb.active — это может привести к чтению чужих данных (как это
    и произошло с Estate AZE). Вместо этого выводим предупреждение
    и возвращаем пусто.
    """
    _require_file(path, "Прогнозы месячные")

    wb = openpyxl.load_workbook(path, data_only=True)

    if not sheet_name:
        if config.DEBUG:
            print(f"[parse_forecast] Не задано имя листа. Возвращаем пусто.")
        return {}

    if sheet_name not in wb.sheetnames:
        # Пытаемся найти лист с похожим именем (регистронезависимо,
        # пробелы/скобки не учитываем) — но только если найдено ровно одно
        # совпадение, иначе не угадываем.
        def _norm_sheet(s):
            return re.sub(r"\s+", " ", str(s).strip().lower())

        target = _norm_sheet(sheet_name)
        matches = [sn for sn in wb.sheetnames if _norm_sheet(sn) == target]

        if len(matches) == 1:
            sheet_name = matches[0]
            if config.DEBUG:
                print(f"[parse_forecast] Лист '{sheet_name}' найден "
                      f"по нечёткому совпадению.")
        else:
            print(f"[WARN] [parse_forecast] Лист '{sheet_name}' не найден "
                  f"в файле {os.path.basename(path)}. "
                  f"Доступные листы: {wb.sheetnames}. "
                  f"Возвращаем пусто (данные не будут заполнены).")
            return {}

    ws = wb[sheet_name]

    col_idx = config.FORECAST_MONTH_COL.get(month)
    if col_idx is None:
        fallback_month = getattr(config, "REPORT_MONTH", None)
        if fallback_month is not None:
            col_idx = config.FORECAST_MONTH_COL.get(fallback_month)
            if config.DEBUG:
                print(f"[parse_forecast] Для месяца {month} нет колонки — "
                      f"использую прогноз на отчётный месяц {fallback_month}")
    if col_idx is None:
        if config.DEBUG:
            print(f"[parse_forecast] Для месяца {month} нет колонки")
        return {}

    if config.DEBUG:
        print(f"[parse_forecast] {os.path.basename(path)} [{sheet_name}], "
              f"месяц={month}, колонка (0-based)={col_idx}")

    raw = {}
    for r in range(1, ws.max_row + 1):
        name_cell = ws.cell(row=r, column=2).value
        if name_cell is None:
            continue
        name = _clean_title(name_cell)
        if name == "":
            continue

        if name.lower() in ("показатель", "итого"):
            continue

        val = _cell_value(ws, r, col_idx)
        if val == 0.0:
            continue

        raw[name] = raw.get(name, 0.0) + val

    if config.DEBUG:
        print(f"[parse_forecast] собрано статей: {len(raw)}, "
              f"примеры: {list(raw.items())[:5]}")

    return _map_to_template(raw, row_map)


# ---------------------------------------------------------------------------
# Вложенные средства
# ---------------------------------------------------------------------------

def parse_investments(path):
    if not os.path.exists(path):
        if config.DEBUG:
            print(f"[parse_investments] Файл не найден: {path}")
        return {}

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    result = {}
    for r in range(1, ws.max_row + 1):
        name_cell = ws.cell(row=r, column=1).value
        if name_cell is None:
            continue
        name = _clean_title(name_cell)
        if name == "":
            continue

        val = None
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (int, float)):
                val = float(v)
                break
            if isinstance(v, str):
                s = v.strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
                if s in ("", "-"):
                    continue
                try:
                    val = float(s)
                    break
                except ValueError:
                    continue

        if val is not None:
            result[name] = val

    return result


# ---------------------------------------------------------------------------
# НДС: файлы «НДС ММ.ГГГГ <Статья>.xlsx» и справочная строка БДиР/Прогнозов
# ---------------------------------------------------------------------------

def _read_month_cell(value):
    """
    Приводит значение ячейки «Месяц начисления» к tuple (год, месяц).
    Понимает:
      - datetime / date из openpyxl;
      - строки "09.2026", "9.2026", "2026-09";
      - "сентябрь 2026".
    Возвращает (year, month) или None.
    """
    import datetime as _dt

    if value is None:
        return None
    if isinstance(value, (_dt.datetime, _dt.date)):
        return (value.year, value.month)

    s = str(value).strip()
    if s == "":
        return None

    # "09.2026" или "9.2026"
    m = re.match(r"^(\d{1,2})[./\-](\d{4})$", s)
    if m:
        month, year = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return (year, month)
        return None

    # "2026-09" или "2026.09"
    m = re.match(r"^(\d{4})[./\-](\d{1,2})$", s)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return (year, month)
        return None

    # "сентябрь 2026"
    _ru = {
        "январ": 1, "феврал": 2, "март": 3, "апрел": 4,
        "ма": 5, "июн": 6, "июл": 7, "август": 8,
        "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
    }
    low = s.lower()
    for pref, num in _ru.items():
        if pref in low:
            my = re.search(r"(\d{4})", low)
            if my:
                return (int(my.group(1)), num)
    return None


def _parse_vat_file(path, target_year, target_month, wanted_objects):
    """
    Читает один файл НДС. Возвращает {'НДС': float, 'с_НДС': float}
    только по строкам, совпавшим с месяцем и направлением.
    """
    result = {"НДС": 0.0, "с_НДС": 0.0}

    if not path or not os.path.exists(path):
        return result

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        print(f"[WARN] [vat] не открыть {path}: {e}")
        return result

    if config.VAT_SHEET_NAME in wb.sheetnames:
        ws = wb[config.VAT_SHEET_NAME]
    else:
        ws = wb.active

    needed = {
        _normalize(config.VAT_COL_MONTH),
        _normalize(config.VAT_COL_OBJECT),
        _normalize(config.VAT_COL_VAT),
    }
    needed.discard("")

    # Ищем шапку в первых 20 строках: должна содержать
    # минимум 2 из нужных колонок.
    header_row = None
    header_map = {}
    for r in range(1, min(ws.max_row, 20) + 1):
        cur = {}
        for c in range(1, min(ws.max_column, 30) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            nv = _normalize(v)
            if nv:
                cur[nv] = c
        if sum(1 for k in needed if k in cur) >= 2:
            header_row = r
            header_map = cur
            break

    if header_row is None:
        print(f"[WARN] [vat] {os.path.basename(path)}: шапка не найдена")
        return result

    c_month  = header_map.get(_normalize(config.VAT_COL_MONTH))
    c_object = header_map.get(_normalize(config.VAT_COL_OBJECT))
    c_vat    = header_map.get(_normalize(config.VAT_COL_VAT))
    c_sumvat = header_map.get(_normalize("Сумма с НДС"))

    if c_month is None or c_object is None or c_vat is None:
        print(f"[WARN] [vat] {os.path.basename(path)}: нет колонок "
              f"Месяц/Направление/НДС")
        return result

    for r in range(header_row + 1, ws.max_row + 1):
        ym = _read_month_cell(ws.cell(row=r, column=c_month).value)
        if ym != (target_year, target_month):
            continue

        obj_val = ws.cell(row=r, column=c_object).value
        if obj_val is None:
            continue
        obj_norm = _normalize(obj_val)

        if wanted_objects and obj_norm not in wanted_objects:
            continue

        result["НДС"]   += _cell_value(ws, r, c_vat - 1)
        if c_sumvat is not None:
            result["с_НДС"] += _cell_value(ws, r, c_sumvat - 1)

    return result


def parse_vat(folder, year, month, object_names):
    """
    Читает ВСЕ файлы НДС из папки, складывает НДС по строкам,
    подходящим по месяцу и объектам.

    Возвращает {'НДС': float, 'с_НДС': float}.
    """
    total = {"НДС": 0.0, "с_НДС": 0.0}

    if not folder or not os.path.isdir(folder):
        return total

    wanted = set()
    for obj in (object_names or []):
        clean = str(obj).split(" [")[0].strip()
        if clean:
            wanted.add(_normalize(clean))

    files = []
    for f in os.listdir(folder):
        if f.startswith("~$"):
            continue
        if not f.lower().endswith(config.VAT_FILE_SUFFIX):
            continue
        if config.VAT_FILE_KEYWORD.lower() not in f.lower():
            continue
        files.append(os.path.join(folder, f))

    for p in files:
        part = _parse_vat_file(p, year, month, wanted)
        total["НДС"]   += part["НДС"]
        total["с_НДС"] += part["с_НДС"]

    if config.DEBUG:
        print(f"[vat] {year}-{month:02d}: НДС={total['НДС']:.2f}, "
              f"с_НДС={total['с_НДС']:.2f}")
    return total


def _month_by_name(name):
    """
    Определяет номер месяца (1..12) по названию.
    Поддерживает русский и английский.
    """
    if name is None:
        return None
    low = str(name).strip().lower()

    ru = {
        "январ": 1, "феврал": 2, "март": 3, "апрел": 4,
        "май": 5, "мая": 5, "июн": 6, "июл": 7, "август": 8,
        "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
    }
    for pref, num in ru.items():
        if pref in low:
            return num

    en = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    }
    for pref, num in en.items():
        if pref in low:
            return num
    return None


def _pick_month_column(ws, anchor_row, month):
    """
    Определяет индекс (1-based) колонки данных для нужного месяца.

    Ищет в таком порядке:
      1. В строках 1..6 листа — значения типа datetime (в Прогнозах
         даты обычно в строке 2).
      2. Если не нашли — в строках 1..6 листа ищем ТЕКСТОВЫЕ названия
         месяцев ("Сентябрь 2026", "September 2026") — это формат БДиР.
      3. Если не нашли — в диапазоне (anchor_row - 10 .. anchor_row)
         ищем datetime (страховка на случай нестандартного файла).
    """
    import datetime as _dt

    # 1) datetime в первых 6 строках
    for r in range(1, min(ws.max_row, 6) + 1):
        for c in range(2, min(ws.max_column, 20) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (_dt.datetime, _dt.date)) and v.month == month:
                return c

    # 2) Текстовое название месяца в первых 6 строках
    for r in range(1, min(ws.max_row, 6) + 1):
        for c in range(2, min(ws.max_column, 20) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            if isinstance(v, (_dt.datetime, _dt.date)):
                continue
            if _month_by_name(v) == month:
                return c

    # 3) datetime в диапазоне над якорем (страховка)
    for r in range(max(1, anchor_row - 10), anchor_row):
        for c in range(2, min(ws.max_column, 20) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (_dt.datetime, _dt.date)) and v.month == month:
                return c

    return None


def _parse_vat_ref_from_file(path, sheet_name, month,
                             wanted_objects=None, subtract_objects=None):
    """
    Универсальный парсер строки «Справочно по НДС выручки:» из БДиР
    или Прогнозов.

    Если subtract_objects задан — берём агрегат и вычитаем из него
    указанные объекты.
    """
    if not path or not os.path.exists(path):
        return None

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        print(f"[WARN] [vat_ref] не открыть {path}: {e}")
        return None

    if sheet_name:
        if sheet_name not in wb.sheetnames:
            return None
        ws = wb[sheet_name]
    else:
        ws = wb.active

    # Якорь «Справочно по НДС выручки:» может находиться
    # в колонке A (БДиР) или в колонке B (некоторые Прогнозы).
    # Ищем сразу в первых двух колонках.
    anchor_row = None
    anchor_col = None
    anchor_norm = _normalize(config.VAT_REF_ROW_TITLE)
    for r in range(1, ws.max_row + 1):
        for c in (1, 2):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            if _normalize(v) == anchor_norm:
                anchor_row = r
                anchor_col = c
                break
        if anchor_row is not None:
            break
    if anchor_row is None:
        return None

    col_idx = _pick_month_column(ws, anchor_row, month)
    if col_idx is None:
        return None

    # ============================================================
    # Логика парсинга справочной строки НДС:
    #
    # 1) Сначала читаем значение из САМОЙ строки-якоря
    #    (в Прогнозах месячных объекты под якорем НЕ перечислены —
    #    там стоит одно число на весь лист).
    #
    # 2) Если под якорем всё-таки есть строки с объектами
    #    (в БДиР они есть), используем их по старой логике:
    #    - для простых блоков суммируем значения по wanted_objects;
    #    - для специальных блоков берём агрегат и вычитаем subtract.
    # ============================================================

    anchor_value = _cell_value(ws, anchor_row, col_idx - 1)

    wanted_norm = set()
    for o in (wanted_objects or []):
        clean = str(o).split(" [")[0].strip()
        if clean:
            wanted_norm.add(_normalize(clean).replace(" ", ""))

    aggregate_name = None
    subtract_list  = []
    if isinstance(subtract_objects, dict):
        aggregate_name = subtract_objects.get("aggregate")
        subtract_list  = subtract_objects.get("subtract", []) or []

    aggregate_norm = _normalize(aggregate_name) if aggregate_name else None
    subtract_norm  = {_normalize(x) for x in subtract_list}

    aggregate_value = None
    subtract_sum = 0.0
    found_objects_under_anchor = False

    for r in range(anchor_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=anchor_col).value
        if v is None:
            continue
        name_norm = _normalize(v)

        if not name_norm:
            continue
        if name_norm == anchor_norm:
            break

        val = _cell_value(ws, r, col_idx - 1)

        name_norm_ns = name_norm.replace(" ", "")
        aggregate_norm_ns = aggregate_norm.replace(" ", "") if aggregate_norm else None
        subtract_norm_ns = {x.replace(" ", "") for x in subtract_norm}

        if aggregate_norm_ns and name_norm_ns == aggregate_norm_ns:
            aggregate_value = val
            found_objects_under_anchor = True
        elif name_norm_ns in subtract_norm_ns:
            subtract_sum += val
            found_objects_under_anchor = True
        elif not aggregate_norm_ns and name_norm_ns in wanted_norm:
            if aggregate_value is None:
                aggregate_value = 0.0
            aggregate_value += val
            found_objects_under_anchor = True

    # --- Если под якорем были объекты (БДиР) — используем старую логику ---
    if found_objects_under_anchor:
        if aggregate_norm:
            if aggregate_value is None:
                return None
            return aggregate_value - subtract_sum
        return aggregate_value

    # --- Если под якорем объектов НЕ было (Прогнозы) — берём значение
    #     из самой строки-якоря. Это одно число на весь лист.
    return anchor_value


def parse_vat_ref_bdr(path, month, block_name, objects_for_block):
    """
    НДС из БДиР для указанного блока и месяца.
    Учитывает специальные правила вычитания.
    """
    spec = config.VAT_SPECIAL_BLOCKS.get(block_name)
    if spec:
        return _parse_vat_ref_from_file(
            path, None, month,
            subtract_objects={
                "aggregate": spec["aggregate"],
                "subtract":  spec["subtract"],
            },
        )
    return _parse_vat_ref_from_file(
        path, None, month,
        wanted_objects=objects_for_block,
    )


def parse_vat_ref_forecast(path, sheet_name, month, block_name, objects_for_block):
    """
    НДС из Прогнозов месячных для указанного блока и месяца.
    Учитывает специальные правила вычитания.
    """
    spec = config.VAT_SPECIAL_BLOCKS.get(block_name)
    if spec:
        return _parse_vat_ref_from_file(
            path, sheet_name, month,
            subtract_objects={
                "aggregate": spec["aggregate"],
                "subtract":  spec["subtract"],
            },
        )
    return _parse_vat_ref_from_file(
        path, sheet_name, month,
        wanted_objects=objects_for_block,
    )