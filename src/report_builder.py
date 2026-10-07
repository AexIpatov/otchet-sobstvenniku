# src/report_builder.py
# ============================================================
# Сборка отчётов.
#
# ВАЖНО: этот модуль импортирует config и parsers. Обратный импорт
# (parsers → report_builder или config → report_builder) недопустим —
# это создаст циклический импорт.
# ============================================================

import os
import re
import openpyxl

import config
from src.parsers import (
    parse_opiu, parse_bdr, parse_forecast, parse_investments,
)


# ============================================================
# БЕЗОПАСНЫЕ ОБЁРТКИ ПАРСЕРОВ
# ============================================================

def _parse_opiu_safe(path, objects, row_map, month_name_ru=None):
    if not path or not os.path.exists(path):
        if config.DEBUG:
            print(f"[safe] файл ОПиУ не найден — колонка будет нулевой: {path}")
        return {}
    try:
        return parse_opiu(path, objects, row_map, month_name_ru=month_name_ru)
    except FileNotFoundError:
        return {}


def _parse_bdr_safe(path, objects, month, row_map):
    if not path or not os.path.exists(path):
        if config.DEBUG:
            print(f"[safe] файл БДиР не найден — колонка будет нулевой: {path}")
        return {}
    try:
        return parse_bdr(path, objects, month, row_map)
    except FileNotFoundError:
        return {}


def _parse_forecast_safe(path, sheet, month, row_map):
    if not path or not os.path.exists(path):
        if config.DEBUG:
            print(f"[safe] файл Прогнозов не найден — колонка будет нулевой: {path}")
        return {}
    try:
        return parse_forecast(path, sheet, month, row_map)
    except FileNotFoundError:
        return {}


# ============================================================
# НОРМАЛИЗАЦИЯ ЮНИКОДА
# ============================================================

_HOMOGLYPH_MAP = {
    "\u0061": "\u0430", "\u0063": "\u0441", "\u0065": "\u0435",
    "\u006f": "\u043e", "\u0070": "\u0440", "\u0078": "\u0445",
    "\u0079": "\u0443", "\u0073": "\u0441",
    "\u0041": "\u0410", "\u0042": "\u0412", "\u0043": "\u0421",
    "\u0045": "\u0415", "\u0048": "\u041d", "\u004b": "\u041a",
    "\u004d": "\u041c", "\u004f": "\u041e", "\u0050": "\u0420",
    "\u0054": "\u0422", "\u0058": "\u0425", "\u0059": "\u0423",
}


def _normalize_unicode(s):
    if s is None:
        return ""
    return "".join(_HOMOGLYPH_MAP.get(ch, ch) for ch in str(s))


# Расширенный regex для всех видов пробельных и невидимых символов.
_ALL_SPACES_RE = re.compile(
    r"[\s\u00a0\u2000-\u200b\u2009\u200a\u202f\u205f\u3000\ufeff]+",
    flags=re.UNICODE,
)


def _norm_for_compare(s):
    """
    Нормализация для сопоставления имён.
    Убирает ВСЕ виды пробелов и невидимых символов
    (обычные, неразрывные, тонкие, bidi-маркеры, BOM и т.п.),
    чтобы 'Выручка Антонияс' и 'ВыручкаАнтонияс' считались
    одинаковыми и не зависели от типографских деталей.
    """
    if s is None:
        return ""
    s = _normalize_unicode(s)
    s = _ALL_SPACES_RE.sub("", s)
    return s.lower()


# ============================================================
# ВСПОМОГАТЕЛЬНОЕ
# ============================================================

def _sum_dicts(*dicts):
    result = {}
    for d in dicts:
        if not d:
            continue
        for k, v in d.items():
            result[k] = result.get(k, 0.0) + (v or 0.0)
    return result


def _empty_period():
    return {
        "prev_fact": {},
        "plan_year": {},
        "plan_month": {},
        "fact_month": {},
        "next_forecast": {},
    }


# ============================================================
# ВЛОЖЕННЫЕ СРЕДСТВА
# ============================================================

def _investments_for_simple_block(block, investments):
    if block not in config.BLOCK_SOURCES:
        return 0.0
    # >>> ИЗМЕНЕНО:
    # Блок «УК» — управляющая компания, но по требованию
    # пользователя в него нужно подтянуть ОБЩУЮ сумму вложений
    # по всему Estate (Estate LV + EU + East).
    # Реализовано через отдельную ветку _total_investments_for_block
    # для консолидированного блока — чтобы не было рекурсии,
    # здесь возвращаем 0, а в _total_investments_for_block
    # обрабатываем отдельно.
    if block == "УК":
        return 0.0
    total = 0.0
    for obj in config.BLOCK_SOURCES[block].get("opiu_objects", []) or []:
        clean = obj.split(" [")[0].strip()
        if not clean:
            continue
        for inv_name, inv_val in investments.items():
            if clean == inv_name or clean in inv_name or inv_name in clean:
                total += float(inv_val or 0.0)
                break
    return total


def _total_investments_for_block(block, investments, cache):
    if block in cache:
        return cache[block]

    # >>> ДОБАВЛЕНО:
    # Для блока «УК» нужно вернуть ОБЩУЮ сумму вложений
    # по всему Estate (Estate LV + EU + East). Это делается
    # через рекурсивный вызов для консолидированного блока.
    if block == "УК":
        total = _total_investments_for_block(
            "ESTATE КОНСОЛИДИРОВАННЫЙ", investments, cache
        )
        cache[block] = total
        if config.DEBUG:
            print(f"[invest] УК = {total} (взято из ESTATE КОНСОЛИДИРОВАННЫЙ)")
        return total

    if block in config.BLOCK_SOURCES:
        total = _investments_for_simple_block(block, investments)
        cache[block] = total
        if config.DEBUG and total > 0:
            print(f"[invest] {block} = {total}")
        return total
    if block in config.TEMPLATE_CONSOLIDATION:
        total = 0.0
        for sub in config.TEMPLATE_CONSOLIDATION[block]:
            total += _total_investments_for_block(sub, investments, cache)
        cache[block] = total
        if config.DEBUG:
            print(f"[invest] {block} = {total} (сумма по подблокам)")
        return total
    cache[block] = 0.0
    return 0.0


# ============================================================
# СБОР ДАННЫХ ПО БЛОКАМ
# ============================================================

def _collect_block(block, sources):
    if block not in sources:
        raise KeyError(f"Блок '{block}' не найден в BLOCK_SOURCES")

    src = sources[block]
    result = _empty_period()
    row_map = config.TEMPLATE_ROW_MAP

    if src["opiu_file"]:
        path_prev = config.OPIU_FILES[src["opiu_file"]].get(config.PREV_MONTH)
        month_ru = config.MONTHS_RU.get(config.PREV_MONTH, "")
        result["prev_fact"] = _parse_opiu_safe(
            path_prev, src["opiu_objects"] or [], row_map,
            month_name_ru=month_ru,
        )

    if src["bdr_file"] and src["bdr_objects"]:
        path_bdr = config.BDR_FILES.get(src["bdr_file"])
        result["plan_year"] = _parse_bdr_safe(
            path_bdr, src["bdr_objects"], config.REPORT_MONTH, row_map
        )

    if src["forecast_file"] and src["forecast_sheet"]:
        path_fc = config.FORECAST_FILES.get(src["forecast_file"])
        result["plan_month"] = _parse_forecast_safe(
            path_fc, src["forecast_sheet"], config.REPORT_MONTH, row_map
        )

    if src["opiu_file"]:
        path_cur = config.OPIU_FILES[src["opiu_file"]].get(config.REPORT_MONTH)
        month_ru = config.MONTHS_RU.get(config.REPORT_MONTH, "")
        result["fact_month"] = _parse_opiu_safe(
            path_cur, src["opiu_objects"] or [], row_map,
            month_name_ru=month_ru,
        )

    if src["forecast_file"] and src["forecast_sheet"]:
        path_fc = config.FORECAST_FILES.get(src["forecast_file"])
        result["next_forecast"] = _parse_forecast_safe(
            path_fc, src["forecast_sheet"], config.NEXT_MONTH, row_map
        )

    return result


def _consolidate(source_blocks, collected):
    result = _empty_period()
    for period in result.keys():
        for b in source_blocks:
            if b not in collected:
                continue
            result[period] = _sum_dicts(result[period], collected[b][period])
    return result


def collect_all_blocks(template_name):
    blocks = config.TEMPLATE_BLOCKS[template_name]
    collected = {}

    for b in blocks:
        if b in config.BLOCK_SOURCES and b not in config.TEMPLATE_CONSOLIDATION:
            if config.DEBUG:
                print(f"[collect] простой блок: {b}")
            collected[b] = _collect_block(b, config.BLOCK_SOURCES)

    for b in blocks:
        if b not in config.TEMPLATE_CONSOLIDATION:
            continue
        if b in collected:
            continue
        sources_b = config.TEMPLATE_CONSOLIDATION[b]
        missing = [s for s in sources_b if s not in collected]
        if missing:
            raise RuntimeError(f"Блок '{b}' зависит от {missing}, но они не собраны.")
        if config.DEBUG:
            print(f"[collect] консолидированный: {b} ← {sources_b}")
        collected[b] = _consolidate(sources_b, collected)

    return collected


# ============================================================
# СПИСОК КЛЮЧЕЙ ВЫРУЧКИ
# ============================================================

_REVENUE_KEYS = [k for k in config.TEMPLATE_ROW_MAP if k.startswith("Выручка")]


def _revenue_sum(block_data, period):
    d = block_data.get(period, {})
    return sum(float(d.get(k, 0.0) or 0.0) for k in _REVENUE_KEYS)


# ============================================================
# ПОИСК ЗАГОЛОВКА БЛОКА В ШАБЛОНЕ
# ============================================================

def _find_block_start(ws, block_name):
    target = _norm_for_compare(block_name)
    for r in range(1, ws.max_row + 1):
        a_val = ws.cell(row=r, column=1).value
        if a_val is None:
            continue
        if not (isinstance(a_val, str) and _norm_for_compare(a_val) == "№п/п"):
            continue
        b_val = ws.cell(row=r, column=2).value
        if b_val is None:
            continue
        if _norm_for_compare(b_val) == target:
            return r
    return 0


def _find_block_end(ws, block_start):
    """
    Возвращает номер последней строки блока
    (строки перед следующим заголовком '№ п/п').
    """
    end_row = block_start
    for r in range(block_start + 2, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if a is not None and _norm_for_compare(a) == "№п/п":
            break
        end_row = r
    return end_row


# ============================================================
# ОБНОВЛЕНИЕ ДАТ В ШАПКЕ БЛОКА (СТРОКА 2)
# ============================================================

_MONTHS_RU_CAP = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}


def _fmt_period(year, month):
    name = _MONTHS_RU_CAP.get(month, str(month))
    return f"{name} {year}"


def _write_block_header_dates(ws, block_start):
    header_row = block_start + 1

    if config.REPORT_MONTH is None or config.REPORT_YEAR is None:
        return

    prev_year  = config.PREV_YEAR  if config.PREV_YEAR  else config.REPORT_YEAR
    prev_month = config.PREV_MONTH if config.PREV_MONTH else config.REPORT_MONTH
    next_year  = config.NEXT_YEAR  if config.NEXT_YEAR  else config.REPORT_YEAR
    next_month = config.NEXT_MONTH if config.NEXT_MONTH else config.REPORT_MONTH

    text_prev = _fmt_period(prev_year, prev_month)
    text_rep  = _fmt_period(config.REPORT_YEAR, config.REPORT_MONTH)
    text_next = _fmt_period(next_year, next_month)

    ws.cell(row=header_row, column=3).value  = text_prev
    ws.cell(row=header_row, column=6).value  = text_rep
    ws.cell(row=header_row, column=8).value  = text_rep
    ws.cell(row=header_row, column=10).value = text_next

    if config.DEBUG:
        print(f"[header] блок {block_start}: "
              f"C2={text_prev!r} F2={text_rep!r} "
              f"H2={text_rep!r} J2={text_next!r}")


# ============================================================
# СЛОВАРЬ КОРОТКИХ ИМЁН → ПОЛНЫХ КЛЮЧЕЙ TEMPLATE_ROW_MAP
# ============================================================

_SHORT_TO_FULL = {
    # ============================================================
    # ФОТ и налоги (важно для шаблона Unelma и ОПиУ Unelma/Nomiqa)
    # ============================================================
    # Строка "- ФОТ" в шаблоне Unelma после _norm_for_compare
    # превращается в "-фот" (дефис сохраняется, пробелы убираются).
    "-фот": "Зарплата",
    "- фот": "Зарплата",
    "фот": "Зарплата",
    "фотпроизводственногоперсонала": "Зарплата",
    "фоткоммерческогоперсонала": "Зарплата",
    "фот[сделки]": "Зарплата",

    "-налоги": "НДФЛ",
    "- налоги": "НДФЛ",
    "налоги": "НДФЛ",
    "взносывфонды": "НДФЛ",

    # ============================================================
    # Выручка Nomiqa (в ОПиУ августа статьи называются иначе)
    # ============================================================
    "выручкаbnq_baku-nomiqa": "Выручка от продажи недвижимости",
    "выручкаdnq_dubai-nomiqa": "Выручка от продажи недвижимости",
    "выручкаenq_europe-nomiqa": "Выручка от продажи недвижимости",
    "выручкаuk_latvia": "Выручка от аренды",
    "выручкаук_latvia": "Выручка от аренды",

    # ============================================================
    # Выручка (общие)
    # ============================================================
    "выручкаотаренды": "Выручка от аренды",
    "выручкаотвозмещениекоммунальныхуслуг": "Выручка от возмещение коммунальных услуг",
    "выручкаотвозмещениякоммунальныхуслуг": "Выручка от возмещение коммунальных услуг",
    "выручкаотпродажинедвижимости": "Выручка от продажи недвижимости",
    "выручкапродажанедвижимости": "Выручка от продажи недвижимости",
    "выручкапродожанедвижимости": "Выручка от продажи недвижимости",
    "выручкаподополнительным услугам": "Выручка по дополнительным услугам",
    "выручкаподополнительнымуслугам": "Выручка по дополнительным услугам",
    "выручкапостроительным проектам": "Выручка строительных проектов",
    "выручкапостроительныхпроектов": "Выручка строительных проектов",

    # ============================================================
    # Расходы (общие)
    # ============================================================
    "налоги": "НДФЛ",
    "коммунальныерасходы": "1.2.10 Коммунальные платежи",
    "налогнанедвижимость": "1.2.16.1 Налог на недвижимость",
    "рко": "1.2.17 РКО",
    "комиссиявнешнимриэлторам и партнерам": "1.2.6 Комиссия внешним риэлторам и партнерам",
    "комиссиявнешнимриэлторамипартнерам": "1.2.6 Комиссия внешним риэлторам и партнерам",
    "закупки,обслуживаниеиремонт:": "1.2.1 Закупка до 1000 евро",
    "закупкиобслуживаниеиремонт": "1.2.1 Закупка до 1000 евро",
    "бытовоеоборудование,инвентарь": "1.2.1.1 Бытовое оборудование, инвентарь",
    "бытовоеоборудованиеинвентарь": "1.2.1.1 Бытовое оборудование, инвентарь",
    "мебель": "1.2.1.2 Мебель",
    "оргтехника": "1.2.1.3 Оргтехника",
    "хоз.принадлежности": "1.2.1.4 Хоз.принадлежности",
    "хозпринадлежности": "1.2.1.4 Хоз.принадлежности",
    "прочиемелкиетмц": "1.2.1.5 Прочие мелкие ТМЦ",
    "обслуживаниеобъектов(бытовыевопросы,безремонта)": "1.2.8.1 Обслуживание объектов (бытовые вопросы, без ремонта)",
    "обслуживаниеобъектов": "1.2.8.1 Обслуживание объектов (бытовые вопросы, без ремонта)",
    "страхование": "1.2.8.2 Страхование",
    "ремонттекущий": "1.2.22 Ремонт текущий",
    "командировки": "1.2.2 Командировочные расходы",
    "объявление,реклама": "1.2.3.1 Платное продвижение - Instagram",
    "объявлениереклама": "1.2.3.1 Платное продвижение - Instagram",
    "такси": "1.2.5.2 Трансфер сотрудников, такси",
    "трансферсотрудников,такси": "1.2.5.2 Трансфер сотрудников, такси",
    "расходынаперсонал": "1.2.5.4 Прочие расходы на персонал",
    "бухгалтер": "1.2.12 Бухгалтер",
    "интернет,программы,телефоны": "1.2.9.1 Связь , интернет, TV",
    "интернетпрограммытелефоны": "1.2.9.1 Связь , интернет, TV",
    "представительскиерасходы": "1.2.11 Представительские расходы",
    "арендаофиса": "1.2.21.1 Аренда офиса",
    "покупкивофис": "1.2.21.2 Административные офисные расходы",
    "коммунальныеофисныерасходы": "1.2.21.3 Коммунальные офисные расходы",
    "налогнаприбыль": "1.2.16.2 Налог на прибыль",
    "ячейки": "1.2.18 Банковская комиссия за использование ячейки",
    "лицензии,госпошлины,разрешения": "1.2.20 Лицензии, гос.пошлины, разрешения",
    "лицензиигоспошлиныразрешения": "1.2.20 Лицензии, гос.пошлины, разрешения",
    "непредвиденныерасходы": "1.2.33 Непредвиденные расходы",
    "вознаграждениеинвестора": "1.2.34 Вознаграждение инвестора",
    "расходыпосодержаниюавтотранспорта": "1.2.19 Расходы по содержанию автотранспорта",
    "юридическиеуслуги(голландия)": "1.2.13 Юридические услуги (Голландия)",
    "юридическиеуслугиголландия": "1.2.13 Юридические услуги (Голландия)",
    "юридическиеуслуги": "1.2.14 Юридические услуги",
    "возвратыклиентам": "1.2.23 Возвраты клиентам",
    "транспортныеуслуги": "1.2.25 Транспортные услуги",
    "сайты(разработкаиобслуживание)": "1.2.30 Сайты (разработка и обслуживание)",
    "сайтыразработкаиобслуживание": "1.2.30 Сайты (разработка и обслуживание)",
    "оффлайнмероприятия": "1.2.31 Оффлайн мероприятия",
    "консультационныеуслуги": "1.2.7 Консультационные услуги",
    "туристическийналог": "1.2.16.4 Туристический налог",
    "расходыпонаправлениюunelma": "1.2.36 Расходы по направлению",
    "маркетинговыерасходы": "1.2.3.1 Платное продвижение - Instagram",
}


def _strip_prefix(title):
    """
    Убирает префикс «2.1.1. », «· », «- » и т.п.
    Для известных коротких имён возвращает полный ключ из TEMPLATE_ROW_MAP.

    Логика проверок (по порядку):
      1. Сырая строка целиком (например, "- ФОТ" → "-фот").
      2. Сырая строка без ведущих «·».
      3. Без нумерации «2.1.1. » (но с дефисом, если он есть).
      4. Без нумерации и без ведущего дефиса «- ».
      5. Если ничего не найдено — возвращаем строку без нумерации.
    """
    raw = str(title).strip()

    # 1) Сырая строка целиком
    raw_norm = _norm_for_compare(raw)
    if raw_norm in _SHORT_TO_FULL:
        return _SHORT_TO_FULL[raw_norm]

    # 2) Убираем ведущие «·»
    t1 = raw.lstrip("·").strip()
    t1_norm = _norm_for_compare(t1)
    if t1_norm in _SHORT_TO_FULL:
        return _SHORT_TO_FULL[t1_norm]

    # 3) Убираем нумерацию «2.1.1. »
    t2 = re.sub(r"^\d+(\.\d+)*\.\s*", "", t1).strip()
    t2_norm = _norm_for_compare(t2)
    if t2_norm in _SHORT_TO_FULL:
        return _SHORT_TO_FULL[t2_norm]

    # 4) Убираем ведущий дефис «- »
    t3 = re.sub(r"^[-–—]\s*", "", t2).strip()
    t3_norm = _norm_for_compare(t3)
    if t3_norm in _SHORT_TO_FULL:
        return _SHORT_TO_FULL[t3_norm]

    # 5) Ничего не нашли — возвращаем строку без нумерации
    #    (нормализованную, чтобы дальше можно было искать в TEMPLATE_ROW_MAP)
    return t2


def _is_formula(cell):
    v = cell.value
    return isinstance(v, str) and v.startswith("=")


# ============================================================
# ОПРЕДЕЛЕНИЕ СЕКЦИЙ (доход / расход / стоп)
# ============================================================
# Используется в _write_block, чтобы понять, в какой секции
# находится строка: доход, расход или стоп (конец таблицы).
#
#   - "доход"  — строки выручки (значения не инвертируются);
#   - "расход" — строки расходов (значения берутся по модулю);
#   - "стоп"   — агрегатные строки (Валовая Прибыль, Рентабельность,
#                ROI, Вложенные средства), после которых цикл
#                _write_block прерывается.
# ============================================================

_SECTION_MARKERS = {
    "выручка": "доход",
    "расходы": "расход",
    "прочиерасходы": "расход",
    "валоваяприбыль": "стоп",
    "рентабельность": "стоп",
    "roiнавложенныесредства,вмесяц": "стоп",
    "roiнавложенныесредствавмесяц": "стоп",
    "планируемыйroiнагод": "стоп",
    "вложенныесредстванаобъекты": "стоп",
}


def _detect_section(title_raw):
    """
    Определяет секцию строки по её названию.
    Возвращает 'доход', 'расход', 'стоп' или None.
    """
    t = str(title_raw).strip().lstrip("·").strip()
    if t.startswith("- "):
        t = t[2:].strip()
    t = re.sub(r"^\d+(\.\d+)*\.\s*", "", t).strip()
    return _SECTION_MARKERS.get(_norm_for_compare(t))


# ============================================================
# КЛАССИФИКАЦИЯ ФОРМУЛ
# ============================================================
# Формулы-агрегаты — это формулы, которые Excel должен
# пересчитывать сам из подстрок или из других ячеек.
# Их НЕЛЬЗЯ перезаписывать значениями парсера.
#
# Признаки агрегата:
#   1) =SUM(...)            — сумма диапазона
#   2) =SUMIF(...)          — сумма по условию
#   3) =A1+B1+C1+...        — арифметика из ссылок
#   4) =A1-B1, =A1*B1, ...  — арифметика из двух ссылок
#   5) =IFERROR(...)        — обёртка, но внутри обычно ссылка
#
# Простая формула-ссылка — это =A1 (одна ссылка без арифметики).
# Битая формула — содержит #REF!, #ДЕЛ/0!, #ЗНАЧ! и т.п.
# ============================================================

_AGGREGATE_PATTERNS = [
    re.compile(r"^=\s*SUM\s*\(", re.IGNORECASE),
    re.compile(r"^=\s*SUMIF\s*\(", re.IGNORECASE),
    re.compile(r"^=\s*SUMIFS\s*\(", re.IGNORECASE),
    re.compile(r"^=\s*IFERROR\s*\(\s*SUM", re.IGNORECASE),
    re.compile(
        r"^=\s*IFERROR\s*\(\s*\$?[A-Z]+\$?\d+\s*[+\-*/]\s*\$?[A-Z]+\$?\d+",
        re.IGNORECASE,
    ),
]

_ARITHMETIC_RE = re.compile(
    r"^=\s*\$?[A-Z]+\$?\d+\s*[+\-*/]\s*\$?[A-Z]+\$?\d+",
    re.IGNORECASE,
)

_BROKEN_TOKENS = ("#REF!", "#ДЕЛ/0!", "#ЗНАЧ!", "#ИМЯ?", "#ЧИСЛО!", "#ПУСТО!", "#Н/Д")

_SIMPLE_REF_RE = re.compile(
    r"^=\s*\$?[A-Z]+\$?\d+\s*$",
    re.IGNORECASE,
)


def _is_aggregate_formula(cell):
    """
    Формула-агрегат: =SUM(...), =SUMIF(...), =A1+B1, =C4+C5+C6,
    =SUM(...)-C23, =IFERROR(SUM(...),0) и т.п.
    Простая ссылка =A1 агрегатом НЕ считается.
    """
    if not _is_formula(cell):
        return False
    v = cell.value.strip()

    # Простая ссылка =A1 — не агрегат.
    if _SIMPLE_REF_RE.match(v):
        return False

    # Явные агрегатные паттерны.
    for pat in _AGGREGATE_PATTERNS:
        if pat.match(v):
            return True
    if _ARITHMETIC_RE.match(v):
        return True

    # Расширенная проверка: формула содержит арифметику со ссылками
    # (+, -, *, /) и хотя бы одну ссылку на ячейку.
    # Покрывает =C7+C8+C9, =SUM(C7:C22)-C23 и т.п.
    if re.search(r"[+\-*/]", v) and re.search(r"\$?[A-Z]+\$?\d+", v):
        return True

    return False

def _is_broken_formula(cell):
    if not _is_formula(cell):
        return False
    v = cell.value
    return any(tok in v for tok in _BROKEN_TOKENS)


def _is_simple_ref_formula(cell):
    if not _is_formula(cell):
        return False
    return bool(_SIMPLE_REF_RE.match(cell.value.strip()))


# ============================================================
# ЗАПИСЬ ЗНАЧЕНИЙ В БЛОК
# ============================================================

def _write_cell(cell, val, force_overwrite=False):
    """
    Записывает значение в ячейку.

    force_overwrite=True — перезаписывать даже агрегаты.
    Используется для колонок E, F, J (план/прогноз), где источник —
    Прогнозы или БДиР, и агрегатная формула шаблона может дать
    неверный результат (0), если подстроки в источнике пустые.
    """
    if not config.PRESERVE_AGGREGATE_FORMULAS:
        cell.value = val
        return

    if not _is_formula(cell):
        cell.value = val
        return

    # Если force_overwrite — пишем значение, даже если это агрегат,
    # НО только если значение ненулевое. Если значение 0, а в ячейке
    # формула-агрегат (=SUM, =E9+E8 и т.п.) — оставляем формулу,
    # чтобы Excel сам пересчитал её из подстрок.
    #
    # Это важно для строк-заголовков подсекций (2.1, 2.4, 3),
    # у которых в шаблоне формула-агрегат, а в источнике (БДиР)
    # такой статьи нет (plan_year.get(...) = 0).
    if force_overwrite:
        if val == 0 and _is_aggregate_formula(cell):
            if config.DEBUG:
                print(f"[write_cell] сохраняю агрегат {cell.coordinate} "
                      f"(force, val=0): {cell.value!r}")
            return
        cell.value = val
        return

    if _is_aggregate_formula(cell):
        if config.DEBUG:
            print(f"[write_cell] сохраняю агрегат {cell.coordinate}: {cell.value!r}")
        return

    if _is_broken_formula(cell):
        cell.value = val
        return

    if _is_simple_ref_formula(cell):
        if val == 0:
            if config.DEBUG:
                print(f"[write_cell] сохраняю простую ссылку "
                      f"{cell.coordinate}: {cell.value!r} (val=0)")
            return
        cell.value = val
        return

    # Прочие формулы (сложные, но не агрегаты) — перезаписываем значением,
    # потому что они, вероятно, неверные (например, =C6/$C$5 в шаблоне
    # Unelma, где должно быть =C6/$C$3).
    cell.value = val


# Формат отображения: целое с разделителем тысяч.
# Значение ячейки остаётся дробным (с копейками), меняется
# только то, как Excel показывает число на экране.
_NUMBER_FORMAT = '#,##0'


def _apply_number_format(cell):
    """Ставит ячейке формат «целое с разделителем тысяч».
    Само значение ячейки не трогает."""
    try:
        cell.number_format = _NUMBER_FORMAT
    except Exception:
        # Если по какой-то причине не получилось — не падаем,
        # просто оставляем формат по умолчанию.
        pass


def _safe_set_value(ws, row, col, value):
    """
    Безопасно записывает значение в ячейку.

    Проблема: в Excel-шаблоне есть ОБЪЕДИНЁННЫЕ ячейки (merged).
    openpyxl представляет все ячейки диапазона, кроме верхней левой,
    объектами MergedCell. У них атрибут .value — read-only, и запись
    вызывает ошибку: 'MergedCell' object attribute 'value' is read-only.

    Решение: если ячейка — MergedCell, находим верхнюю левую ячейку
    её диапазона и пишем в неё. Если не удалось — тихо пропускаем.

    Возвращает True, если запись удалась, иначе False.
    """
    cell = ws.cell(row=row, column=col)

    # Быстрая проверка: обычная ячейка (не MergedCell)?
    if type(cell).__name__ != "MergedCell":
        try:
            cell.value = value
            return True
        except AttributeError:
            # На всякий случай — вдруг всё-таки MergedCell.
            pass

    # Это MergedCell (или запись не удалась). Ищем родителя.
    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= row <= rng.max_row
                and rng.min_col <= col <= rng.max_col):
            top_left = ws.cell(row=rng.min_row, column=rng.min_col)
            try:
                top_left.value = value
                return True
            except AttributeError:
                return False

    # Не MergedCell, но запись не удалась по другой причине.
    return False


def _safe_apply_format(ws, row, col):
    """
    Безопасно применяет формат к ячейке. Для MergedCell — к верхней
    левой ячейке диапазона.
    """
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ != "MergedCell":
        _apply_number_format(cell)
        return

    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= row <= rng.max_row
                and rng.min_col <= col <= rng.max_col):
            top_left = ws.cell(row=rng.min_row, column=rng.min_col)
            _apply_number_format(top_left)
            return


def _write_block(ws, block_start, block_data, investments_amount):
    if config.DEBUG:
        print(f"[write] блок начинается в строке {block_start}")

    end_row = _find_block_end(ws, block_start)

    if config.DEBUG:
        print(f"[write] конец блока: строка {end_row}")

    prev = block_data["prev_fact"]
    plan_year = block_data["plan_year"]
    plan_month = block_data["plan_month"]
    fact_month = block_data["fact_month"]
    next_fc = block_data["next_forecast"]

    skip_titles_norm = {_norm_for_compare(x) for x in (
        "Валовая Прибыль", "Рентабельность",
        "ROI на вложенные средства, в месяц",
        "ПЛАНИРУЕМЫЙ  ROI на ГОД", "ПЛАНИРУЕМЫЙ ROI на ГОД",
        "Вложенные средства на объекты",
    )}

    current_section = None

    for r in range(block_start + 1, end_row + 1):
        title_raw = ws.cell(row=r, column=2).value
        if title_raw is None:
            continue

        section = _detect_section(title_raw)
        if section == "стоп":
            break
        if section is not None:
            current_section = section
            continue

        title = _strip_prefix(title_raw)
        if not title:
            continue
        if _norm_for_compare(title) in skip_titles_norm:
            continue

        # ДИАГНОСТИКА: показываем, что именно пишем в строку «Зарплата».
        # Помогает найти, почему в отчёте 0 вместо ожидаемого значения.
        if config.DEBUG and _norm_for_compare(title) == _norm_for_compare("Зарплата"):
            print(f"[write_block] строка {r}: title_raw={title_raw!r}, "
                  f"title={title!r}, "
                  f"prev={prev.get(title, 0.0)}, "
                  f"plan_year={plan_year.get(title, 0.0)}, "
                  f"plan_m={plan_month.get(title, 0.0)}, "
                  f"fact_m={fact_month.get(title, 0.0)}, "
                  f"next_fc={next_fc.get(title, 0.0)}")

        val_prev      = float(prev.get(title, 0.0) or 0.0)
        val_plan_year = float(plan_year.get(title, 0.0) or 0.0)
        val_plan_m    = float(plan_month.get(title, 0.0) or 0.0)
        val_fact_m    = float(fact_month.get(title, 0.0) or 0.0)
        val_next_fc   = float(next_fc.get(title, 0.0) or 0.0)

        if current_section in ("расход", "доход"):
            val_prev      = abs(val_prev)
            val_plan_year = abs(val_plan_year)
            val_plan_m    = abs(val_plan_m)
            val_fact_m    = abs(val_fact_m)
            val_next_fc   = abs(val_next_fc)

        for col_idx, val, force in (
            (3, val_prev, False),        # C — факт прошлого месяца (ОПиУ)
            (5, val_plan_year, True),    # E — план года (БДиР) — перезаписываем агрегаты
            (6, val_plan_m, True),       # F — план месяца (Прогнозы) — перезаписываем агрегаты
            (8, val_fact_m, False),      # H — факт текущего месяца (ОПиУ)
            (10, val_next_fc, True),     # J — прогноз будущего (Прогнозы) — перезаписываем агрегаты
        ):
            cell = ws.cell(row=r, column=col_idx)
            if type(cell).__name__ == "MergedCell":
                _safe_set_value(ws, r, col_idx, val)
                _safe_apply_format(ws, r, col_idx)
            else:
                _write_cell(cell, val, force_overwrite=force)
                _apply_number_format(cell)

    for r in range(block_start, end_row + 1):
        title_raw = ws.cell(row=r, column=2).value
        if title_raw is None:
            continue
        if _norm_for_compare(_strip_prefix(title_raw)) == _norm_for_compare("Вложенные средства на объекты"):
            # >>> ИЗМЕНЕНО: вложения пишем ТОЛЬКО в колонку C.
            # Раньше писали во все колонки (C, E, F, H, J) — это нужно
            # было для ROI блока «УК», но визуально неправильно:
            # вложения не меняются от месяца к месяцу.
            #
            # Для блока «УК» ROI будет считаться по колонке C
            # (см. _fix_uk_formulas), поэтому E/F/H/J там не нужны.
            _safe_set_value(ws, r, 3, investments_amount or 0.0)
            break

    # Применяем числовой формат «целое с разделителем тысяч» ко всем
    # числовым ячейкам блока. Это НЕ меняет хранимые значения —
    # только то, как Excel показывает число (без копеек).
    for r in range(block_start, end_row + 1):
        for c in range(3, 17):   # C..P
            cell = ws.cell(row=r, column=c)
            if type(cell).__name__ == "MergedCell":
                _safe_apply_format(ws, r, c)
            else:
                _apply_number_format(cell)

    # >>> Процентный формат для строк «Рентабельность», «ROI», «ПЛАНИРУЕМЫЙ ROI».
    pct_titles_norm = {
        _norm_for_compare("Рентабельность"),
        _norm_for_compare("ROI на вложенные средства, в месяц"),
        _norm_for_compare("ПЛАНИРУЕМЫЙ ROI на ГОД"),
        _norm_for_compare("ПЛАНИРУЕМЫЙ  ROI на ГОД"),
    }
    for r in range(block_start, end_row + 1):
        title_raw = ws.cell(row=r, column=2).value
        if title_raw is None:
            continue
        t_norm = _norm_for_compare(_strip_prefix(title_raw))
        if t_norm not in pct_titles_norm:
            continue
        for c in (3, 5, 6, 8, 10):
            cell = ws.cell(row=r, column=c)
            if type(cell).__name__ == "MergedCell":
                continue
            try:
                cell.number_format = '0.00%'
            except Exception:
                pass



# ============================================================
# ДЕТАЛИЗАЦИЯ СТРОК
# ============================================================

def _write_detailed_rows(ws, block, block_start, collected):
    if block not in config.DETAILED_CONSOLIDATION:
        return
    detail_map = config.DETAILED_CONSOLIDATION[block]
    detail_map_norm = {_norm_for_compare(k): v for k, v in detail_map.items()}

    end_row = _find_block_end(ws, block_start)

    for r in range(block_start + 1, end_row + 1):
        title_raw = ws.cell(row=r, column=2).value
        if title_raw is None:
            continue
        title = _strip_prefix(title_raw)
        title_norm = _norm_for_compare(title)
        if title_norm not in detail_map_norm:
            continue

        source_block = detail_map_norm[title_norm]
        if source_block not in collected:
            continue

        src_data = collected[source_block]
        vals = {
            "prev_fact":     _revenue_sum(src_data, "prev_fact"),
            "plan_year":     _revenue_sum(src_data, "plan_year"),
            "plan_month":    _revenue_sum(src_data, "plan_month"),
            "fact_month":    _revenue_sum(src_data, "fact_month"),
            "next_forecast": _revenue_sum(src_data, "next_forecast"),
        }
        for col_idx, val in (
            (3, vals["prev_fact"]),
            (5, vals["plan_year"]),
            (6, vals["plan_month"]),
            (8, vals["fact_month"]),
            (10, vals["next_forecast"]),
        ):
            cell = ws.cell(row=r, column=col_idx)
            if type(cell).__name__ == "MergedCell":
                _safe_set_value(ws, r, col_idx, val)
            else:
                _write_cell(cell, val)
                _apply_number_format(cell)


# ============================================================
# ПРИНУДИТЕЛЬНОЕ ИСПРАВЛЕНИЕ ФОРМУЛЫ «ВАЛОВАЯ ПРИБЫЛЬ»
# ============================================================

def _fix_valovaya_pribyl_formulas(ws, block_start):
    vp_row = None
    for r in range(block_start + 1, ws.max_row + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        a_val = ws.cell(row=r, column=1).value
        if a_val is not None and _norm_for_compare(a_val) == "№п/п" and r != block_start:
            break
        if _norm_for_compare(_strip_prefix(title)) == _norm_for_compare("Валовая Прибыль"):
            vp_row = r
            break

    if vp_row is None:
        return

    revenue_row = None
    expenses_row = None
    for r in range(block_start + 1, vp_row):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t_norm = _norm_for_compare(_strip_prefix(title))
        if t_norm == _norm_for_compare("Выручка") and revenue_row is None:
            revenue_row = r
        if t_norm == _norm_for_compare("Расходы") and expenses_row is None:
            expenses_row = r

    if revenue_row is None or expenses_row is None:
        return

    for col_letter, col_idx in (("C", 3), ("E", 5), ("F", 6), ("H", 8), ("J", 10)):
        _safe_set_value(
            ws, vp_row, col_idx,
            f"={col_letter}{revenue_row}-{col_letter}{expenses_row}",
        )

    if config.DEBUG:
        print(f"[fix VP] блок {block_start}: ВП в строке {vp_row}, "
              f"Выручка={revenue_row}, Расходы={expenses_row}")


# ============================================================
# СКРЫТИЕ ПУСТЫХ СТРОК
# ============================================================

def _cell_numeric_or_none(cell):
    v = cell.value
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


def _is_row_empty(ws, r):
    for col in (3, 5, 6, 8, 10):
        cell = ws.cell(row=r, column=col)
        v = cell.value

        if isinstance(v, str) and v.startswith("="):
            return False

        num = _cell_numeric_or_none(cell)
        if num is None:
            continue
        if num != 0.0:
            return False

    return True


def _hide_empty_rows(ws, template_name):
    total_hidden = 0

    for block in config.TEMPLATE_BLOCKS[template_name]:
        start = _find_block_start(ws, block)
        if start == 0:
            continue

        end_row = _find_block_end(ws, start)

        for r in range(start + 2, end_row + 1):
            title_raw = ws.cell(row=r, column=2).value
            if title_raw is None:
                continue

            title_norm = _norm_for_compare(_strip_prefix(title_raw))
            if title_norm in {
                "валоваяприбыль", "рентабельность",
                "roiнавложенныесредства,вмесяц",
                "планируемыйroiнагод",
                "вложенныесредстванаобъекты",
            }:
                continue

            if _is_row_empty(ws, r):
                ws.row_dimensions[r].hidden = True
                total_hidden += 1

    if config.DEBUG and total_hidden:
        print(f"[hide] {template_name}: скрыто пустых строк — {total_hidden}")


# ============================================================
# ПОСТПРОЦЕССИНГ ФОРМУЛ
# ============================================================

_NOMIQA_REF_FIX = [
    ("$C$54", "$C$105"),
    ("$E$54", "$E$105"),
    ("$F$54", "$F$105"),
    ("$H$54", "$H$105"),
    ("$J$54", "$J$105"),
]


def _postprocess_nomiqa_formulas(ws, template_name):
    if template_name != "Nomiqa":
        return
    block_name = "Nomiqa (консолидированный)"
    start = _find_block_start(ws, block_name)
    if start == 0:
        return

    end_row = _find_block_end(ws, start)

    fixed = 0
    for r in range(start, end_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            new_v = v
            for old, new in _NOMIQA_REF_FIX:
                if old in new_v:
                    new_v = new_v.replace(old, new)
            if new_v != v:
                _safe_set_value(ws, r, c, new_v)
                fixed += 1
    if config.DEBUG and fixed:
        print(f"[postproc] Nomiqa: исправлено формул: {fixed}")


# >>> ИЗМЕНЕНО:
# Раньше _postprocess_unelma_formulas обрабатывал только строку 2.1
# ("ФОТ + налоги на ФОТ") и заменял =X6/$C$3*100 на =X7/$C$3*100.
# Но в шаблоне Unelma формулы долей выглядят как =X<n>/$C$5 (деление
# на Расходы), а должны быть =X<n>/$C$3 (деление на Выручку).
#
# Теперь функция обрабатывает ВСЕ строки блока Unelma и заменяет
# все формулы вида =X<n>/$C$5 на =X<n>/$C$3.
#
# Паттерн: =<col><row>/$C$5, где col ∈ {C, E, F, H, J}, row — число.
_UNELMA_DENOM_PATTERN = re.compile(
    r"^=(C|E|F|H|J)(\d+)(\/\$C\$5)$"
)

# На всякий случай — старый паттерн с $C$3*100 (может встретиться).
_UNELMA_ROW21_OLD_PATTERN = re.compile(r"^=(C|E|F|H|J)(\d+)(\/\$C\$3\*100)$")


def _postprocess_unelma_formulas(ws, template_name):
    if template_name != "Unelma":
        return
    block_name = "Unelma"
    start = _find_block_start(ws, block_name)
    if start == 0:
        return

    end_row = _find_block_end(ws, start)

    fixed_denom = 0
    fixed_old = 0

    for r in range(start, end_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue

            # 1) =X<n>/$C$5 → =X<n>/$C$3
            m = _UNELMA_DENOM_PATTERN.match(v.strip())
            if m:
                col = m.group(1)
                row = m.group(2)
                _safe_set_value(ws, r, c, f"={col}{row}/$C$3")
                fixed_denom += 1
                continue

            # 2) =X<n>/$C$3*100 → =X<n>/$C$3
            m2 = _UNELMA_ROW21_OLD_PATTERN.match(v.strip())
            if m2:
                col = m2.group(1)
                row = m2.group(2)
                _safe_set_value(ws, r, c, f"={col}{row}/$C$3")
                fixed_old += 1
                continue

    if config.DEBUG and (fixed_denom or fixed_old):
        print(f"[postproc] Unelma: заменено $C$5→$C$3: {fixed_denom}, "
              f"$C$3*100→$C$3: {fixed_old}")


def _postprocess_formulas(ws, template_name):
    # ВАЖНО: в новых шаблонах Estate/Unelma/Nomiqa формулы долей
    # уже правильные (D7=C7/$C$6, G7=F7/$F$6, I7=H7/$H$6, K7=J7/$J$6).
    # Никакой постпроцессинг не нужен — только испортим.
    return


# ============================================================
# ПОСТПРОЦЕССИНГ: ДОПОЛНИТЕЛЬНАЯ КОНСОЛИДАЦИЯ СТРОК
# ============================================================
# Для блоков из config.EXTRA_CONSOLIDATION пересчитывает указанные
# строки как сумму одноимённых строк из подблоков.
#
# Логика получения значения из подблока (в порядке приоритета):
#   1. Если в ws лежит ЧИСЛО — берём его.
#   2. Если в ws лежит формула =SUM(X:Y) — разворачиваем её
#      вручную, читая значения диапазона из ws.
#   3. Если формула другого вида — берём закешированное значение
#      из ws_values (data_only=True).
#   4. Если ничего не вышло — вклад 0.
# ============================================================

_SUM_FORMULA_RE = re.compile(
    r"^\s*=\s*SUM\s*\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$",
    flags=re.IGNORECASE,
)


def _col_letter_to_idx(letter):
    """A → 1, B → 2, ..., Z → 26, AA → 27."""
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx


def _extract_sum_from_formula(formula, target_ws):
    """
    Если формула вида =SUM(X1:X2) — считаем сумму значений
    диапазона из target_ws. Возвращает число или None.
    """
    m = _SUM_FORMULA_RE.match(formula)
    if not m:
        return None
    c1l, r1s, c2l, r2s = m.groups()
    c1 = _col_letter_to_idx(c1l)
    c2 = _col_letter_to_idx(c2l)
    r1 = int(r1s)
    r2 = int(r2s)
    total = 0.0
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            v = target_ws.cell(row=rr, column=cc).value
            if isinstance(v, (int, float)):
                total += float(v)
            elif isinstance(v, str):
                s = v.strip().replace(",", ".")
                try:
                    total += float(s)
                except ValueError:
                    pass
    return total


def _postprocess_extra_consolidation(ws, template_name, ws_values=None):
    rules = getattr(config, "EXTRA_CONSOLIDATION", {})
    if not rules:
        return

    for block_name, row_rules in rules.items():
        if block_name not in config.TEMPLATE_BLOCKS.get(template_name, []):
            continue

        block_start = _find_block_start(ws, block_name)
        if block_start == 0:
            if config.DEBUG:
                print(f"[extra-cons] блок не найден в шаблоне: {block_name}")
            continue

        end_row = _find_block_end(ws, block_start)

        for row_title, source_blocks in row_rules.items():
            # Нормализованное название без номера — используется
            # и для поиска целевой строки, и для поиска источников.
            # В шаблоне строка называется просто «Прочие расходы»,
            # в подблоках — «3. Прочие расходы»; после _strip_prefix
            # оба названия дают «Прочие расходы».
            row_title_norm = _norm_for_compare(_strip_prefix(row_title))

            if config.DEBUG:
                print(f"[extra-cons] ищу row_title={row_title!r}, "
                      f"norm={row_title_norm!r}")

            # 1) Ищем целевую строку в блоке.
            target_row = None
            for r in range(block_start + 1, end_row + 1):
                title = ws.cell(row=r, column=2).value
                if title is None:
                    continue
                if config.DEBUG:
                    print(f"[extra-cons]   стр.{r}: {title!r} "
                          f"→ {_norm_for_compare(_strip_prefix(title))!r}")
                if _norm_for_compare(_strip_prefix(title)) == row_title_norm:
                    target_row = r
                    break
            if target_row is None:
                if config.DEBUG:
                    print(f"[extra-cons] {block_name}: строка '{row_title}' не найдена")
                continue

            # 2) Ищем ту же строку в каждом подблоке.
            source_rows = []
            for sb in source_blocks:
                sb_start = _find_block_start(ws, sb)
                if sb_start == 0:
                    if config.DEBUG:
                        print(f"[extra-cons] подблок не найден: {sb}")
                    continue
                sb_end = _find_block_end(ws, sb_start)

                for r in range(sb_start + 1, sb_end + 1):
                    title = ws.cell(row=r, column=2).value
                    if title is None:
                        continue
                    t_norm = _norm_for_compare(_strip_prefix(title))
                    if t_norm == row_title_norm:
                        source_rows.append((sb, r))
                        break

            if not source_rows:
                if config.DEBUG:
                    print(f"[extra-cons] {block_name}: строка '{row_title}' — "
                          f"нет ни одного источника")
                continue

            # 3) Складываем значения по колонкам.
            totals = {}
            for col_idx in (3, 5, 6, 8, 10):
                total = 0.0
                for sb, sr in source_rows:
                    v = ws.cell(row=sr, column=col_idx).value
                    if isinstance(v, (int, float)):
                        total += float(v)
                        continue
                    if isinstance(v, str) and v.startswith("="):
                        # 1) Разворачиваем =SUM(X:Y) вручную
                        val = _extract_sum_from_formula(v, ws)
                        if val is not None:
                            total += val
                            continue
                        # 2) Fallback — закешированное значение
                        if ws_values is not None:
                            v_cached = ws_values.cell(row=sr, column=col_idx).value
                            if isinstance(v_cached, (int, float)):
                                total += float(v_cached)
                                continue
                    # Если формула не =SUM(...) и кеша нет — вклад 0.
                totals[col_idx] = total

            for col_idx, total in totals.items():
                _safe_set_value(ws, target_row, col_idx, total)

            if config.DEBUG:
                src_rows = [r for _, r in source_rows]
                print(f"[extra-cons] {block_name}: '{row_title}' "
                      f"(стр.{target_row}) = "
                      f"C={totals[3]}, E={totals[5]}, F={totals[6]}, "
                      f"H={totals[8]}, J={totals[10]} "
                      f"из строк {src_rows}")


# ============================================================
# ПРОВЕРКА ФОРМУЛ
# ============================================================

_BROKEN_FORMULA_TOKENS = (
    "#REF!", "#ДЕЛ/0!", "#ЗНАЧ!", "#ИМЯ?", "#ЧИСЛО!", "#ПУСТО!", "#Н/Д",
)


def _verify_formulas(ws, template_name):
    if not config.VERIFY_FORMULAS:
        return
    bad = 0
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if not isinstance(v, str):
                continue
            for token in _BROKEN_FORMULA_TOKENS:
                if token in v:
                    print(f"[WARN] {template_name}: битая формула в {cell.coordinate}: {v!r}")
                    bad += 1
                    break
    if config.DEBUG and bad == 0:
        print(f"[verify] {template_name}: битых формул не найдено")

# ============================================================
# ВЫЧИСЛЕНИЕ АГРЕГАТНЫХ ФОРМУЛ
# ============================================================
# После записи значений в блок нужно вычислить формулы-агрегаты
# (=SUM(...), =A1+B1, =A1-B1 и т.п.), потому что openpyxl не
# пересчитывает формулы, а Excel пересчитает только при открытии.
# Без этого зависимые формулы (=IFERROR(C8/$C$6,0)) дадут
# #ДЕЛ/0! или 0.

_SIMPLE_ARITH_RE = re.compile(
    r"^=\s*\$?([A-Z]+)\$?(\d+)\s*([+\-*/])\s*\$?([A-Z]+)\$?(\d+)\s*$"
)

_CELL_REF_RE = re.compile(r"\$?([A-Z]+)\$?(\d+)")


def _col_letter_to_idx(letter):
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx


def _read_cell_value(ws, ref):
    """Читает значение ячейки по ссылке вида C6 или $C$6.
    Если в ячейке формула-агрегат — разворачиваем её рекурсивно.
    """
    m = _CELL_REF_RE.fullmatch(ref.strip())
    if not m:
        return None
    col = _col_letter_to_idx(m.group(1))
    row = int(m.group(2))
    v = ws.cell(row=row, column=col).value
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        if v.startswith("="):
            # >>> ДОБАВЛЕНО: ячейка содержит формулу — разворачиваем.
            # Это нужно, чтобы ROI УК (=E645/E649) посчитался,
            # даже если E645 — это формула =E599-E604.
            return _eval_formula(ws, v)
        s = v.strip().replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _eval_formula(ws, formula, depth=0):
    """
    Рекурсивно вычисляет простые формулы:
      - =SUM(X1:X2)
      - =A1+B1, =A1-B1, =A1*B1, =A1/B1
      - =IFERROR(expr, fallback)
    Возвращает число или None, если не удалось.
    """
    if depth > 20:
        return None
    if not isinstance(formula, str) or not formula.startswith("="):
        return None

    f = formula.strip()

    # >>> ДОБАВЛЕНО: простая ссылка =C604 (без арифметики).
    # Без этого _force_numeric_denominator не может развернуть
    # строку «Расходы» в консолидированном блоке, где формула
    # вида =C604+C607+C608+C609+C618+C619+C620, а C604 —
    # это формула =C605+C606.
    m_ref = _SIMPLE_REF_RE.match(f)
    if m_ref:
        ref = f[1:].strip()
        return _read_cell_value(ws, ref)

    # IFERROR(expr, fallback)
    m = re.match(r"^=IFERROR\s*\((.*)\)\s*$", f, re.IGNORECASE)
    if m:
        inner = m.group(1)
        # Разбиваем по запятой верхнего уровня
        depth_p = 0
        parts = []
        cur = ""
        for ch in inner:
            if ch == "(":
                depth_p += 1
            elif ch == ")":
                depth_p -= 1
            if ch == "," and depth_p == 0:
                parts.append(cur)
                cur = ""
            else:
                cur += ch
        parts.append(cur)
        if len(parts) >= 2:
            val = _eval_formula(ws, "=" + parts[0].strip(), depth + 1)
            if val is None or val == 0:
                fb = parts[1].strip()
                try:
                    return float(fb)
                except ValueError:
                    return _eval_formula(ws, "=" + fb, depth + 1)
            return val
        return None

    # SUM(X1:X2)
    m = re.match(r"^=SUM\s*\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$",
                 f, re.IGNORECASE)
    if m:
        c1 = _col_letter_to_idx(m.group(1))
        r1 = int(m.group(2))
        c2 = _col_letter_to_idx(m.group(3))
        r2 = int(m.group(4))
        total = 0.0
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                v = ws.cell(row=rr, column=cc).value
                if isinstance(v, (int, float)):
                    total += float(v)
                elif isinstance(v, str) and v.startswith("="):
                    sub = _eval_formula(ws, v, depth + 1)
                    if sub is not None:
                        total += sub
        return total

    # SUM(A1,B1,C1,...)
    m = re.match(r"^=SUM\s*\((.*)\)\s*$", f, re.IGNORECASE)
    if m:
        inner = m.group(1)
        total = 0.0
        for part in inner.split(","):
            part = part.strip()
            if ":" in part:
                sub_m = re.match(r"\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)", part)
                if sub_m:
                    c1 = _col_letter_to_idx(sub_m.group(1))
                    r1 = int(sub_m.group(2))
                    c2 = _col_letter_to_idx(sub_m.group(3))
                    r2 = int(sub_m.group(4))
                    for rr in range(r1, r2 + 1):
                        for cc in range(c1, c2 + 1):
                            v = ws.cell(row=rr, column=cc).value
                            if isinstance(v, (int, float)):
                                total += float(v)
                continue
            v = _read_cell_value(ws, part)
            if v is not None:
                total += v
        return total

    # Простая арифметика A1+B1, A1-B1, A1*B1, A1/B1
    m = _SIMPLE_ARITH_RE.match(f)
    if m:
        v1 = _read_cell_value(ws, m.group(1) + m.group(2))
        v2 = _read_cell_value(ws, m.group(4) + m.group(5))
        if v1 is None or v2 is None:
            return None
        op = m.group(3)
        if op == "+":
            return v1 + v2
        if op == "-":
            return v1 - v2
        if op == "*":
            return v1 * v2
        if op == "/":
            return v1 / v2 if v2 != 0 else None

    # >>> ДОБАВЛЕНО: ссылка и число: =C647*12, =C647/12, =C647+12, =C647-12.
    # Также число и ссылка: =12*C647, =12-C647 и т.п.
    m2 = re.match(
        r"^=\s*(\$?[A-Z]+\$?\d+|\d+(?:\.\d+)?)"
        r"\s*([+\-*/])\s*"
        r"(\$?[A-Z]+\$?\d+|\d+(?:\.\d+)?)\s*$",
        f,
    )
    if m2:
        def _resolve(tok):
            tok = tok.strip()
            # число?
            try:
                return float(tok)
            except ValueError:
                pass
            # ссылка?
            val = _read_cell_value(ws, tok)
            if val is not None:
                return val
            # в ячейке формула? — раскрываем рекурсивно
            m_ref = _CELL_REF_RE.fullmatch(tok)
            if m_ref:
                col = _col_letter_to_idx(m_ref.group(1))
                row = int(m_ref.group(2))
                inner = ws.cell(row=row, column=col).value
                if isinstance(inner, str) and inner.startswith("="):
                    return _eval_formula(ws, inner, depth + 1)
            return None

        v1 = _resolve(m2.group(1))
        v2 = _resolve(m2.group(3))
        if v1 is None or v2 is None:
            return None
        op = m2.group(2)
        if op == "+":
            return v1 + v2
        if op == "-":
            return v1 - v2
        if op == "*":
            return v1 * v2
        if op == "/":
            return v1 / v2 if v2 != 0 else None

    # Цепочка A1+B1+C1+... или SUM(X:Y)-C23
    # Пробуем вычислить как сумму/разность слагаемых.
    if re.match(r"^=", f):
        # Разбиваем по + и - на верхнем уровне (не внутри скобок).
        depth_p = 0
        parts = []
        cur = ""
        cur_sign = "+"
        for ch in f[1:]:
            if ch == "(":
                depth_p += 1
            elif ch == ")":
                depth_p -= 1
            if ch in "+-" and depth_p == 0:
                parts.append((cur_sign, cur.strip()))
                cur_sign = ch
                cur = ""
            else:
                cur += ch
        parts.append((cur_sign, cur.strip()))

        if len(parts) >= 2:
            total = 0.0
            ok = True
            for sign, part in parts:
                if not part:
                    continue
                # Пытаемся вычислить часть как формулу (SUM, IFERROR, ссылка).
                val = None
                if part.startswith("SUM") or part.startswith("IFERROR"):
                    val = _eval_formula(ws, "=" + part, depth + 1)
                else:
                    # Простая ссылка или число.
                    try:
                        val = float(part)
                    except ValueError:
                        val = _read_cell_value(ws, part)
                        if val is None:
                            val = _eval_formula(ws, "=" + part, depth + 1)
                if val is None:
                    ok = False
                    break
                if sign == "-":
                    total -= val
                else:
                    total += val
            if ok:
                return total

    return None

def _fix_prochie_raskhody(ws, block_start, block_end):
    """
    Пересчитывает строку «Прочие расходы» как сумму строк 3.x
    (после «Прочие расходы» и до «Валовая Прибыль»).
    """
    pr_row = None
    for r in range(block_start + 1, block_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t_norm = _norm_for_compare(_strip_prefix(title))
        if t_norm == _norm_for_compare("Прочие расходы"):
            pr_row = r
            break
    if pr_row is None:
        return

    end_pr = pr_row
    for r in range(pr_row + 1, block_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t_norm = _norm_for_compare(_strip_prefix(title))
        if t_norm in {"валоваяприбыль", "рентабельность"}:
            break
        end_pr = r

    # Если после «Прочие расходы» нет ни одной подстроки 3.x —
    # НЕ трогаем ячейку (в консолидированном блоке она заполняется
    # из подблоков через _postprocess_extra_consolidation).
    if end_pr <= pr_row:
        if config.DEBUG:
            print(f"[fix PR] блок {block_start}: после «Прочие расходы» "
                  f"(стр.{pr_row}) нет подстрок — пропускаю.")
        return

    for col_idx in (3, 5, 6, 8, 10):
        total = 0.0
        for r in range(pr_row + 1, end_pr + 1):
            v = ws.cell(row=r, column=col_idx).value
            if isinstance(v, (int, float)):
                total += float(v)
        _safe_set_value(ws, pr_row, col_idx, total)
    if config.DEBUG:
        print(f"[fix PR] блок {block_start}: «Прочие расходы» = строка {pr_row}, "
              f"сумма строк {pr_row + 1}..{end_pr}")

def _force_numeric_denominator(ws, block_start, block_end):
    """
    Принудительно вычисляет строку «Расходы» (знаменатель долей)
    в столбцах C, E, F, H, J. Если формула не вычисляется —
    оставляет как есть, но пишет предупреждение в DEBUG.
    """
    denom_row = None
    for r in range(block_start + 1, block_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t_norm = _norm_for_compare(_strip_prefix(title))
        if t_norm == _norm_for_compare(config.SHARE_DENOMINATOR_TITLE):
            denom_row = r
            break
    if denom_row is None:
        return

    for col_idx in (3, 5, 6, 8, 10):
        cell = ws.cell(row=denom_row, column=col_idx)
        if type(cell).__name__ == "MergedCell":
            continue
        v = cell.value
        if not (isinstance(v, str) and v.startswith("=")):
            continue
        # Пытаемся вычислить формулу.
        val = _eval_formula(ws, v)
        if val is not None:
            _safe_set_value(ws, denom_row, col_idx, val)
            if config.DEBUG:
                print(f"[force denom] строка {denom_row}, колонка {col_idx}: "
                      f"{v!r} → {val}")
        elif config.DEBUG:
            print(f"[force denom] НЕ удалось вычислить {cell.coordinate}: {v!r}")

def _fix_share_formulas(ws, block_start, block_end):
    """
    Заполняет столбцы D и I долями расходов.

    Логика:
      - доля считается ТОЛЬКО для строк секции «расход»
        (строки 2.x и 3.x), включая саму строку «3. Прочие расходы»,
        но ИСКЛЮЧАЯ заголовок «2. Расходы»;
      - знаменатель — строка «Расходы» блока
        (config.SHARE_DENOMINATOR_TITLE);
      - доля ВСЕГДА записывается ЧИСЛОМ. Если числитель или
        знаменатель — формула, она сначала вычисляется через
        _eval_formula. Если вычислить не удалось — пишем 0.
    """
    # --- 1. Находим строку «Расходы» (знаменатель) ---
    denom_row = None
    for r in range(block_start + 1, block_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t_norm = _norm_for_compare(_strip_prefix(title))
        if t_norm == _norm_for_compare(config.SHARE_DENOMINATOR_TITLE):
            denom_row = r
            break
    if denom_row is None:
        return

    # --- 2. Определяем границы секций в блоке ---
    sections = []
    for r in range(block_start + 1, block_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        sec = _detect_section(title)
        if sec is not None:
            sections.append((r, sec))

    def _section_of(row):
        cur = None
        for sr, name in sections:
            if sr <= row:
                cur = name
            else:
                break
        return cur

    # --- 3. Строки, для которых доли НЕ считаем ---
    skip_norm = {
        "валоваяприбыль", "рентабельность",
        "roiнавложенныесредства,вмесяц",
        "планируемыйroiнагод",
        "вложенныесредстванаобъекты",
        "расходы",     # знаменатель
        "выручка",     # заголовок секции доход
    }

    # --- 4. Заполняем доли ---
    for r in range(block_start + 1, block_end + 1):
        if r == denom_row:
            continue

        title = ws.cell(row=r, column=2).value
        if title is None:
            continue

        t_norm = _norm_for_compare(_strip_prefix(title))
        if t_norm in skip_norm:
            continue

        if _section_of(r) != "расход":
            continue

        # Пары (числитель, доля):
        #   C(3) → D(4)   — доля по факту прошлого месяца
        #   F(6) → G(7)   — доля по плану текущего месяца
        #   H(8) → I(9)   — доля по факту текущего месяца
        #   J(10) → K(11) — доля по прогнозу будущего месяца
        for src_col, tgt_col in ((3, 4), (6, 7), (8, 9), (10, 11)):
            num_val = ws.cell(row=r, column=src_col).value
            denom_val = ws.cell(row=denom_row, column=src_col).value

            # Если числитель — формула, вычисляем его.
            if isinstance(num_val, str) and num_val.startswith("="):
                num_val = _eval_formula(ws, num_val)
            # Если знаменатель — формула, вычисляем его.
            if isinstance(denom_val, str) and denom_val.startswith("="):
                denom_val = _eval_formula(ws, denom_val)

            try:
                num = float(num_val) if num_val is not None else 0.0
                denom = float(denom_val) if denom_val is not None else 0.0
            except (ValueError, TypeError):
                num, denom = 0.0, 0.0

            share = (num / denom) if denom != 0 else 0.0

            _safe_set_value(ws, r, tgt_col, share)

            # Формат "0.00%" — Excel сам умножит значение (0..1) на 100
            # и покажет "38.37%". Внутри ячейки остаётся доля 0.3837.
            tgt_cell = ws.cell(row=r, column=tgt_col)
            if type(tgt_cell).__name__ != "MergedCell":
                try:
                    tgt_cell.number_format = '0.00%'
                except Exception:
                    pass

    if config.DEBUG:
        print(f"[fix share] блок {block_start}: знаменатель «Расходы» в строке "
              f"{denom_row}, доли записаны в D и I (числами)")

def _recalc_aggregates(ws, block_start, block_end):
    """
    Проходит по строкам блока и вычисляет все формулы-агрегаты
    (=SUM, =A1+B1, =A1-B1 и т.п.), заменяя их на значения.
    Делает несколько проходов, пока что-то меняется
    (для вложенных агрегатов).

    ВАЖНО: этот метод вычисляет ТОЛЬКО агрегаты. Простые ссылки
    вида =C4 обрабатывает отдельная функция _resolve_simple_refs,
    которую нужно вызывать СРАЗУ ПОСЛЕ этой.
    """
    # Столбцы D (4) и I (9) — это формулы долей, которые
    # Excel пересчитает сам. Мы их НЕ трогаем.
    SKIP_COLS = {4, 9}

    for _ in range(10):  # максимум 10 проходов
        changed = False
        for r in range(block_start, block_end + 1):
            for c in range(3, ws.max_column + 1):
                if c in SKIP_COLS:
                    continue
                cell = ws.cell(row=r, column=c)
                # Пропускаем MergedCell — их .value read-only.
                if type(cell).__name__ == "MergedCell":
                    continue
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                if not _is_aggregate_formula(cell):
                    continue
                # Пытаемся вычислить
                val = _eval_formula(ws, v)
                if val is not None:
                    _safe_set_value(ws, r, c, val)
                    changed = True
        if not changed:
            break


# ============================================================
# ВЫЧИСЛЕНИЕ ПРОСТЫХ ССЫЛОК =C4
# ============================================================
# В шаблоне Unelma есть формулы-ссылки =C4, =E4, =F4, =H4, =J4
# в строке «1. Выручка», которые ссылаются на строку 4
# («Выручка строительных проектов»). Программа записывает
# данные в строку 3 («1. Выручка»), а не в строку 4, поэтому
# ссылка =C4 даёт 0.
#
# Эта функция принудительно заменяет простые ссылки на значения.
# Вызывать ПОСЛЕ _recalc_aggregates.

def _resolve_simple_refs(ws, block_start, block_end):
    """
    Заменяет формулы-ссылки вида =C4, =$C$4 на их значения.
    Для агрегатов (SUM, A1+B1 и т.п.) ничего не делает.
    """
    for r in range(block_start, block_end + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if type(cell).__name__ == "MergedCell":
                continue
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            # НЕ трогаем агрегаты
            if _is_aggregate_formula(cell):
                continue
            # Проверяем, что это простая ссылка =C4 или =$C$4
            m = _SIMPLE_REF_RE.match(v.strip())
            if not m:
                continue
            ref = v.strip()[1:].strip()   # "C4" или "$C$4"
            val = _read_cell_value(ws, ref)
            if val is not None:
                _safe_set_value(ws, r, c, val)
                if config.DEBUG:
                    print(f"[resolve ref] {cell.coordinate}: {v!r} → {val}")


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def build_report(template_name):
    print(f"\n{'=' * 60}")
    print(f"СБОРКА ОТЧЁТА: {template_name}")
    print(f"{'=' * 60}")

    template_path = config.TEMPLATE_FILES[template_name]
    output_path = config.OUTPUT_FILES[template_name]

    if not template_path.exists():
        msg = f"[ERROR] Шаблон не найден: {template_path}"
        print(msg)
        raise FileNotFoundError(msg)

    collected = collect_all_blocks(template_name)

    wb = openpyxl.load_workbook(template_path)
    if config.TEMPLATE_SHEET_NAME in wb.sheetnames:
        ws = wb[config.TEMPLATE_SHEET_NAME]
    else:
        ws = wb.active

    # Вторая загрузка того же файла с data_only=True:
    # openpyxl с data_only возвращает вычисленные Excel значения
    # вместо формул. Нужно для fallback-чтения закешированных
    # значений в постпроцессинге консолидации.
    wb_values = openpyxl.load_workbook(template_path, data_only=True)
    if config.TEMPLATE_SHEET_NAME in wb_values.sheetnames:
        ws_values = wb_values[config.TEMPLATE_SHEET_NAME]
    else:
        ws_values = wb_values.active

    investments = {}
    try:
        investments = parse_investments(config.REF_INVESTMENTS_FILE)
    except FileNotFoundError as e:
        print(f"[WARN] {e}")

    investments_cache = {}

    # ============================================================
    # ЦИКЛ ПО БЛОКАМ
    # ============================================================
    for block in config.TEMPLATE_BLOCKS[template_name]:
        start = _find_block_start(ws, block)
        if start == 0:
            print(f"[WARN] блок не найден в шаблоне: {block}")
            continue
        if config.DEBUG:
            end_preview = _find_block_end(ws, start)
            print(f"[block] {block!r}: строки {start}..{end_preview}")
        if block not in collected:
            print(f"[WARN] данные не собраны: {block}")
            continue

        _write_block_header_dates(ws, start)
        block_investments = _total_investments_for_block(
            block, investments, investments_cache
        )
        _write_block(ws, start, collected[block], block_investments)
        _write_detailed_rows(ws, block, start, collected)

        _fix_valovaya_pribyl_formulas(ws, start)

        _block_end = _find_block_end(ws, start)

        # Пересчитываем «Прочие расходы» как сумму подстрок 3.x.
        # НО: для блоков из EXTRA_CONSOLIDATION эту строку
        # заполняет _postprocess_extra_consolidation, поэтому
        # _fix_prochie_raskhody здесь вызывать НЕЛЬЗЯ (иначе
        # она перезапишет результат консолидации нулём).
        _extra_cons_blocks = set(
            getattr(config, "EXTRA_CONSOLIDATION", {}).keys()
        )
        if block not in _extra_cons_blocks:
            _fix_prochie_raskhody(ws, start, _block_end)

        # 1. Сначала вычисляем агрегаты (SUM, A1+B1 и т.п.),
        #    чтобы «Расходы» стали числом.
        _recalc_aggregates(ws, start, _block_end)

        # 1a. ← ДОБАВЛЕНО: вычисляем простые ссылки =C4
        #     (нужно для шаблона Unelma).
        _resolve_simple_refs(ws, start, _block_end)

        # 1b. Принудительно раскрываем строку «Расходы» (знаменатель
        #     долей) в числа, если она всё ещё формула.
        _force_numeric_denominator(ws, start, _block_end)

        # 2. Только ПОСЛЕ этого записываем формулы долей в D и I.
        _fix_share_formulas(ws, start, _block_end)

        # 3. ← ДОБАВЛЕНО: повторно вычисляем агрегаты — теперь,
        #    когда доли записаны числами, некоторые формулы
        #    могли стать вычислимыми.
        _recalc_aggregates(ws, start, _block_end)

        # 3a. ← ДОБАВЛЕНО: снова резолвим простые ссылки —
        #     на случай, если после второго _recalc_aggregates
        #     появились новые значения.
        _resolve_simple_refs(ws, start, _block_end)

    # ============================================================
    # ПОСТОБРАБОТКА КОНСОЛИДАЦИИ СТРОК
    # ============================================================
    # Передаём ws_values для fallback-чтения формул в подблоках.
    _postprocess_extra_consolidation(ws, template_name, ws_values)

    # После консолидации строк («Прочие расходы» и т.п.) нужно
    # ПЕРЕСЧИТАТЬ доли в тех блоках, куда мы только что что-то
    # записали. Иначе доля для строки «Прочие расходы» останется 0.
    extra_cons_blocks = set(
        getattr(config, "EXTRA_CONSOLIDATION", {}).keys()
    )
    for block in config.TEMPLATE_BLOCKS[template_name]:
        if block not in extra_cons_blocks:
            continue
        b_start = _find_block_start(ws, block)
        if b_start == 0:
            continue
        b_end = _find_block_end(ws, b_start)

        _recalc_aggregates(ws, b_start, b_end)
        _force_numeric_denominator(ws, b_start, b_end)
        _fix_share_formulas(ws, b_start, b_end)
        if config.DEBUG:
            print(f"[post-extra-cons] пересчитал агрегаты и доли "
                  f"в блоке {block!r}")

    # ============================================================
    # ИСПРАВЛЕНИЕ ROI ВО ВСЕХ БЛОКАХ ESTATE
    # ============================================================
    # Вложения теперь пишутся только в колонку C, а формулы ROI
    # в шаблоне ссылаются на E8/F8/H8/J8. Нужно переписать их так,
    # чтобы знаменателем всегда была колонка C.
    if template_name == "Estate":
        _fix_roi_formulas_for_all_blocks(ws)

    # ============================================================
    # СПЕЦИАЛЬНАЯ ОБРАБОТКА БЛОКА «УК» (только для Estate)
    # ============================================================
    # Валовая Прибыль УК = ВП консолидированного блока − Расходы УК.
    # ROI = ВП УК / Вложения, ПЛАНИРУЕМЫЙ ROI = ROI * 12.
    if template_name == "Estate":
        # Разворачиваем ВСЕ формулы в блоке ESTATE КОНСОЛИДИРОВАННЫЙ
        # в числа — чтобы ВП УК считалась из уже посчитанных значений.
        est_start = _find_block_start(ws, "ESTATE КОНСОЛИДИРОВАННЫЙ")
        if est_start != 0:
            est_end = _find_block_end(ws, est_start)
            for _pass in range(20):
                changed = False
                for r in range(est_start, est_end + 1):
                    for c in (3, 5, 6, 8, 10):
                        cell = ws.cell(row=r, column=c)
                        if type(cell).__name__ == "MergedCell":
                            continue
                        v = cell.value
                        if not (isinstance(v, str) and v.startswith("=")):
                            continue
                        val = _eval_formula(ws, v)
                        if val is not None:
                            _safe_set_value(ws, r, c, val)
                            changed = True
                        elif config.DEBUG:
                            print(f"[fix UK] НЕ развернулась {cell.coordinate}: {v!r}")
                if not changed:
                    break
            if config.DEBUG:
                print(f"[fix UK] развернул формулы блока "
                      f"ESTATE КОНСОЛИДИРОВАННЫЙ (строки {est_start}..{est_end})")

            # Рентабельность ESTATE КОНСОЛИДИРОВАННЫЙ
            # = Валовая Прибыль / Выручка (внутри того же блока).
            est_row_rev = None
            est_row_vp = None
            est_row_rent = None
            for r in range(est_start + 1, est_end + 1):
                title = ws.cell(row=r, column=2).value
                if title is None:
                    continue
                t = _norm_for_compare(_strip_prefix(title))
                if t == _norm_for_compare("Выручка") and est_row_rev is None:
                    est_row_rev = r
                elif t == _norm_for_compare("Валовая Прибыль") and est_row_vp is None:
                    est_row_vp = r
                elif t == _norm_for_compare("Рентабельность") and est_row_rent is None:
                    est_row_rent = r

            if (est_row_rev is not None
                    and est_row_vp is not None
                    and est_row_rent is not None):
                for col_letter, col_idx in (("C", 3), ("E", 5), ("F", 6), ("H", 8), ("J", 10)):
                    _safe_set_value(
                        ws, est_row_rent, col_idx,
                        f"={col_letter}{est_row_vp}/{col_letter}{est_row_rev}",
                    )
                # Разворачиваем формулу в число (или оставляем формулу —
                # Excel пересчитает при открытии).
                for _pass in range(5):
                    changed = False
                    for col_idx in (3, 5, 6, 8, 10):
                        cell = ws.cell(row=est_row_rent, column=col_idx)
                        if type(cell).__name__ == "MergedCell":
                            continue
                        v = cell.value
                        if not (isinstance(v, str) and v.startswith("=")):
                            continue
                        val = _eval_formula(ws, v)
                        if val is not None:
                            _safe_set_value(ws, est_row_rent, col_idx, val)
                            changed = True
                    if not changed:
                        break
                # Процентный формат для строки «Рентабельность».
                for col_idx in (3, 5, 6, 8, 10):
                    cell = ws.cell(row=est_row_rent, column=col_idx)
                    if type(cell).__name__ == "MergedCell":
                        continue
                    try:
                        cell.number_format = '0.00%'
                    except Exception:
                        pass
                if config.DEBUG:
                    print(f"[fix UK] Рентабельность ESTATE КОНСОЛИДИРОВАННЫЙ "
                          f"(стр.{est_row_rent}) = "
                          f"=C{est_row_vp}/C{est_row_rev} "
                          f"(ВП / Выручка)")

        # Формулы блока УК
        _fix_uk_formulas(ws)

        # Пересчитываем формулы, которые только что проставили.
        uk_start = _find_block_start(ws, "УК")
        if uk_start != 0:
            uk_end = _find_block_end(ws, uk_start)
            # Многораундовый пересчёт для УК.
            for _pass in range(20):
                changed = False
                for r in range(uk_start, uk_end + 1):
                    for c in (3, 5, 6, 8, 10):
                        cell = ws.cell(row=r, column=c)
                        if type(cell).__name__ == "MergedCell":
                            continue
                        v = cell.value
                        if not (isinstance(v, str) and v.startswith("=")):
                            continue
                        val = _eval_formula(ws, v)
                        if val is not None:
                            _safe_set_value(ws, r, c, val)
                            changed = True
                        elif config.DEBUG:
                            print(f"[fix UK] НЕ развернулась {cell.coordinate}: {v!r}")
                if not changed:
                    break
            _force_numeric_denominator(ws, uk_start, uk_end)

    # ============================================================
    # ФИНАЛЬНАЯ ПРОВЕРКА И СОХРАНЕНИЕ
    # ============================================================
    _postprocess_formulas(ws, template_name)
    _verify_formulas(ws, template_name)

    try:
        wb.save(output_path)
        print(f"[OK] Сохранено: {output_path}")
    except PermissionError:
        msg = (f"[ERROR] Файл открыт в Excel и не может быть перезаписан: "
               f"{output_path}\n        Закройте его и повторите сборку.")
        print(msg)
        raise

    return output_path

def _fix_roi_formulas_for_all_blocks(ws):
    """
    Для всех блоков Estate переписывает формулы ROI и ПЛАНИРУЕМЫЙ ROI
    так, чтобы знаменатель был всегда C<вложения> (абсолютная ссылка).

    Зачем: после правки «вложения только в C» в E/F/H/J вложения пусты,
    и формулы =E_ВП/E_Вложения дают 0. Теперь будет:
      ROI в C = C_ВП/C_Вложения
      ROI в E = E_ВП/C_Вложения
      ROI в F = F_ВП/C_Вложения
      ROI в H = H_ВП/C_Вложения
      ROI в J = J_ВП/C_Вложения
    """
    for block_name in config.TEMPLATE_BLOCKS.get("Estate", []):
        start = _find_block_start(ws, block_name)
        if start == 0:
            continue
        end = _find_block_end(ws, start)

        row_vp = None            # «4. Валовая Прибыль»
        row_roi = None           # «6. ROI на вложенные средства, в месяц»
        row_roi_year = None      # «7. ПЛАНИРУЕМЫЙ ROI на ГОД»
        row_investments = None   # «8. Вложенные средства на объекты»

        for r in range(start + 1, end + 1):
            title = ws.cell(row=r, column=2).value
            if title is None:
                continue
            t = _norm_for_compare(_strip_prefix(title))
            if t == _norm_for_compare("Валовая Прибыль") and row_vp is None:
                row_vp = r
            elif t == _norm_for_compare("ROI на вложенные средства, в месяц") and row_roi is None:
                row_roi = r
            elif t == _norm_for_compare("ПЛАНИРУЕМЫЙ ROI на ГОД") and row_roi_year is None:
                row_roi_year = r
            elif t == _norm_for_compare("Вложенные средства на объекты") and row_investments is None:
                row_investments = r

        if row_vp is None or row_roi is None or row_investments is None:
            if config.DEBUG:
                print(f"[fix ROI] {block_name!r}: не хватает строк — пропускаю "
                      f"(ВП={row_vp}, ROI={row_roi}, Вложения={row_investments})")
            continue

        # ROI: =<col>ВП / $C$<вложения>
        for col_letter, col_idx in (("C", 3), ("E", 5), ("F", 6), ("H", 8), ("J", 10)):
            _safe_set_value(
                ws, row_roi, col_idx,
                f"={col_letter}{row_vp}/$C${row_investments}",
            )

        # ПЛАНИРУЕМЫЙ ROI: =<col>ROI * 12
        if row_roi_year is not None:
            for col_letter, col_idx in (("C", 3), ("E", 5), ("F", 6), ("H", 8), ("J", 10)):
                _safe_set_value(
                    ws, row_roi_year, col_idx,
                    f"={col_letter}{row_roi}*12",
                )

        # Разворачиваем формулы в числа (или оставляем — Excel пересчитает)
        for _pass in range(5):
            changed = False
            for rr in (row_roi, row_roi_year):
                if rr is None:
                    continue
                for col_idx in (3, 5, 6, 8, 10):
                    cell = ws.cell(row=rr, column=col_idx)
                    if type(cell).__name__ == "MergedCell":
                        continue
                    v = cell.value
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    val = _eval_formula(ws, v)
                    if val is not None:
                        _safe_set_value(ws, rr, col_idx, val)
                        changed = True
                    elif config.DEBUG:
                        print(f"[fix ROI]   НЕ развернул {cell.coordinate}: {v!r}")
            if not changed:
                break

        # Финальная страховка: если формула ROI всё ещё не развернулась —
        # считаем её вручную через прямое чтение ячеек ВП и Вложений.
        for col_idx in (3, 5, 6, 8, 10):
            cell = ws.cell(row=row_roi, column=col_idx)
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                vp_cell = ws.cell(row=row_vp, column=col_idx)
                inv_cell = ws.cell(row=row_investments, column=3)  # всегда C
                vp_val = vp_cell.value
                inv_val = inv_cell.value
                if isinstance(vp_val, str) and vp_val.startswith("="):
                    vp_val = _eval_formula(ws, vp_val)
                if isinstance(inv_val, str) and inv_val.startswith("="):
                    inv_val = _eval_formula(ws, inv_val)
                try:
                    vp_num = float(vp_val) if vp_val is not None else 0.0
                    inv_num = float(inv_val) if inv_val is not None else 0.0
                    if inv_num != 0:
                        _safe_set_value(ws, row_roi, col_idx, vp_num / inv_num)
                except (ValueError, TypeError):
                    pass

            # И то же для ПЛАНИРУЕМЫЙ ROI = ROI * 12
            if row_roi_year is not None:
                cell_y = ws.cell(row=row_roi_year, column=col_idx)
                v_y = cell_y.value
                if isinstance(v_y, str) and v_y.startswith("="):
                    roi_cell = ws.cell(row=row_roi, column=col_idx)
                    roi_val = roi_cell.value
                    if isinstance(roi_val, str) and roi_val.startswith("="):
                        roi_val = _eval_formula(ws, roi_val)
                    try:
                        roi_num = float(roi_val) if roi_val is not None else 0.0
                        _safe_set_value(ws, row_roi_year, col_idx, roi_num * 12)
                    except (ValueError, TypeError):
                        pass

        if config.DEBUG:
            print(f"[fix ROI] {block_name!r}: ROI стр.{row_roi}, "
                  f"ROI_год стр.{row_roi_year}, "
                  f"вложения стр.{row_investments} (делён. на $C${row_investments})")

def _fix_uk_formulas(ws):
    """
    Явно проставляет формулы для блока «УК»:
      - Валовая Прибыль = ВП блока ESTATE КОНСОЛИДИРОВАННЫЙ - Расходы УК
      - ROI = Валовая Прибыль УК / Вложенные средства
      - ПЛАНИРУЕМЫЙ ROI = ROI * 12

    Строки определяются по названиям, а не по номерам, — чтобы
    работать при любом шаблоне.
    """
    # 1) Находим начало и конец блока УК.
    uk_start = _find_block_start(ws, "УК")
    if uk_start == 0:
        if config.DEBUG:
            print("[fix UK] блок 'УК' не найден — пропускаю")
        return
    uk_end = _find_block_end(ws, uk_start)

    # 2) Находим начало и конец блока ESTATE КОНСОЛИДИРОВАННЫЙ.
    est_start = _find_block_start(ws, "ESTATE КОНСОЛИДИРОВАННЫЙ")
    if est_start == 0:
        if config.DEBUG:
            print("[fix UK] блок 'ESTATE КОНСОЛИДИРОВАННЫЙ' не найден — пропускаю")
        return
    est_end = _find_block_end(ws, est_start)

    # 3) В блоке УК ищем нужные строки.
    uk_row_expenses = None       # «2. Расходы»
    uk_row_valprofit = None      # «4. Валовая Прибыль»
    uk_row_roi = None            # «6. ROI на вложенные средства, в месяц»
    uk_row_roi_year = None       # «7. ПЛАНИРУЕМЫЙ ROI на ГОД»
    uk_row_investments = None    # «8. Вложенные средства на объекты»

    for r in range(uk_start + 1, uk_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t = _norm_for_compare(_strip_prefix(title))
        if t == _norm_for_compare("Расходы") and uk_row_expenses is None:
            uk_row_expenses = r
        elif t == _norm_for_compare("Валовая Прибыль") and uk_row_valprofit is None:
            uk_row_valprofit = r
        elif t == _norm_for_compare("ROI на вложенные средства, в месяц") and uk_row_roi is None:
            uk_row_roi = r
        elif t == _norm_for_compare("ПЛАНИРУЕМЫЙ ROI на ГОД") and uk_row_roi_year is None:
            uk_row_roi_year = r
        elif t == _norm_for_compare("Вложенные средства на объекты") and uk_row_investments is None:
            uk_row_investments = r

    # 4) В блоке ESTATE КОНСОЛИДИРОВАННЫЙ ищем строку «Валовая Прибыль».
    est_row_valprofit = None
    for r in range(est_start + 1, est_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t = _norm_for_compare(_strip_prefix(title))
        if t == _norm_for_compare("Валовая Прибыль"):
            est_row_valprofit = r
            break

    if config.DEBUG:
        print(f"[fix UK] УК: Расходы={uk_row_expenses}, ВП={uk_row_valprofit}, "
              f"ROI={uk_row_roi}, ROI_год={uk_row_roi_year}, "
              f"Вложения={uk_row_investments}; "
              f"ВП консолидированного={est_row_valprofit}")

    # 5) Подставляем формулы.
    if (uk_row_valprofit is None
            or uk_row_expenses is None
            or est_row_valprofit is None):
        if config.DEBUG:
            print("[fix UK] не хватает одной из строк — ВП УК не будет задана")
    else:
        for col_letter, col_idx in (("C", 3), ("E", 5), ("F", 6), ("H", 8), ("J", 10)):
            _safe_set_value(
                ws, uk_row_valprofit, col_idx,
                f"={col_letter}{est_row_valprofit}-{col_letter}{uk_row_expenses}",
            )


    if (uk_row_roi is None
            or uk_row_valprofit is None
            or uk_row_investments is None):
        if config.DEBUG:
            print("[fix UK] не хватает одной из строк — ROI УК не будет задан")
    else:
        # >>> ИЗМЕНЕНО: ROI УК во всех колонках делится на C-вложения.
        # Вложения записаны только в колонке C (там реальные данные),
        # в E/F/H/J их нет. Поэтому для всех колонок используем
        # абсолютную ссылку $C${uk_row_investments}.
        for col_letter, col_idx in (("C", 3), ("E", 5), ("F", 6), ("H", 8), ("J", 10)):
            _safe_set_value(
                ws, uk_row_roi, col_idx,
                f"={col_letter}{uk_row_valprofit}/$C${uk_row_investments}",
            )

    if uk_row_roi_year is None or uk_row_roi is None:
        if config.DEBUG:
            print("[fix UK] не хватает одной из строк — ПЛАНИРУЕМЫЙ ROI УК не будет задан")
    else:
        for col_letter, col_idx in (("C", 3), ("E", 5), ("F", 6), ("H", 8), ("J", 10)):
            _safe_set_value(
                ws, uk_row_roi_year, col_idx,
                f"={col_letter}{uk_row_roi}*12",
            )

    # >>> ДОБАВЛЕНО: Рентабельность УК = ВП УК / Расходы УК.
    uk_row_rent = None
    for r in range(uk_start + 1, uk_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t = _norm_for_compare(_strip_prefix(title))
        if t == _norm_for_compare("Рентабельность"):
            uk_row_rent = r
            break

    # >>> ИЗМЕНЕНО: Рентабельность УК = ВП УК / Выручка консолидированная.
    # (Раньше было ВП УК / Расходы УК — это неверно.)
    # В блоке ESTATE КОНСОЛИДИРОВАННЫЙ ищем строку «Выручка».
    est_row_revenue = None
    for r in range(est_start + 1, est_end + 1):
        title = ws.cell(row=r, column=2).value
        if title is None:
            continue
        t = _norm_for_compare(_strip_prefix(title))
        if t == _norm_for_compare("Выручка"):
            est_row_revenue = r
            break

    if (uk_row_rent is not None
            and uk_row_valprofit is not None
            and est_row_revenue is not None):
        for col_letter, col_idx in (("C", 3), ("E", 5), ("F", 6), ("H", 8), ("J", 10)):
            _safe_set_value(
                ws, uk_row_rent, col_idx,
                f"={col_letter}{uk_row_valprofit}/{col_letter}{est_row_revenue}",
            )
        if config.DEBUG:
            print(f"[fix UK] Рентабельность УК (стр.{uk_row_rent}) = "
                  f"=C{uk_row_valprofit}/C{est_row_revenue} "
                  f"(ВП УК / Выручка консолидированная)")

    # >>> ДОБАВЛЕНО: ставим процентный формат для строк 5 (Рентабельность),
    # 6 (ROI) и 7 (ПЛАНИРУЕМЫЙ ROI) в блоке УК.
    # Значение в ячейке остаётся долей (0.0012), а Excel показывает
    # его как процент (0.12%).
    for row_num in (uk_row_rent, uk_row_roi, uk_row_roi_year):
        if row_num is None:
            continue
        for col_idx in (3, 5, 6, 8, 10):   # C, E, F, H, J
            cell = ws.cell(row=row_num, column=col_idx)
            if type(cell).__name__ == "MergedCell":
                continue
            try:
                cell.number_format = '0.00%'
            except Exception:
                pass

def build_all_reports(reports=None):
    if reports is None:
        reports = ["Estate", "Unelma", "Nomiqa"]

    for name in reports:
        if name not in config.TEMPLATE_FILES:
            print(f"[ERROR] Неизвестный отчёт: {name}")
            continue

        missing = config.check_inputs(name)
        missing_all = []
        for key in ("template", "opiu_prev", "opiu_cur", "bdr", "forecast"):
            for p in missing[key]:
                missing_all.append((key, p))

        only_missing_template = all(k == "template" for k, _ in missing_all)

        if missing_all and only_missing_template:
            print(f"\n[ERROR] {name}: не найден шаблон — отчёт НЕ строится.")
            for key, p in missing_all:
                print(f"    • Шаблон: {p}")
            continue

        if missing_all:
            print(f"\n[WARN] {name}: часть входных файлов отсутствует — "
                  f"соответствующие колонки будут нулевыми.")
            _labels = {
                "template":  "Шаблон",
                "opiu_prev": "ОПиУ за прошлый месяц",
                "opiu_cur":  "ОПиУ за текущий месяц",
                "bdr":       "БДиР",
                "forecast":  "Прогноз месячный",
            }
            for key, p in missing_all:
                print(f"    • {_labels[key]}: {p}")

        try:
            build_report(name)
        except Exception as e:
            print(f"[ERROR] {name}: {e}")