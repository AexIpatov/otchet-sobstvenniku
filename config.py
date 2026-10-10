# config.py
# ============================================================
# Все настройки, маппинги, пути для программы «Отчет для собственника».
#
# ВАЖНО: этот файл содержит ТОЛЬКО данные и настройки.
# Никакой логики (парсинг, сборка отчётов) здесь быть не должно —
# вся логика живёт в src/parsers.py и src/report_builder.py.
#
# ОБЛАЧНАЯ ВЕРСИЯ:
#   - BASE_DIR вычисляется автоматически (папка, где лежит этот файл).
#   - ОПиУ, БДиР и Прогнозы читаются из ВРЕМЕННОЙ папки (TEMP_DIR),
#     куда app.py сохраняет файлы, загруженные пользователем.
#   - Шаблоны и Справочники — из папок внутри проекта.
# ============================================================

import os
import tempfile
from pathlib import Path


# ============================================================
# 1. БАЗОВЫЕ ПУТИ (динамические)
# ============================================================

# Папка, где лежит этот файл (config.py). Это и есть корень проекта.
BASE_DIR = Path(__file__).resolve().parent

DATA_DIR      = BASE_DIR / "data"
INPUT_DIR     = DATA_DIR / "input"
OUTPUT_DIR    = DATA_DIR / "output"

# Папки внутри проекта (заливаются в GitHub)
TEMPLATES_DIR = BASE_DIR / "templates"
REF_DIR       = INPUT_DIR / "Справочники"

# >>> ВРЕМЕННАЯ ПАПКА для файлов, загруженных пользователем.
# На Streamlit Cloud (Linux) — это /tmp. На Windows — папка TEMP.
# Файлы здесь живут, пока работает сессия, потом удаляются системой.
TEMP_DIR      = Path(tempfile.gettempdir()) / "otchet_sobstvenniku"

# Подпапки внутри временной папки
TEMP_OPIU_DIR         = TEMP_DIR / "ОПиУ"
TEMP_FORECAST_DIR     = TEMP_DIR / "Прогнозы"
TEMP_VAT_DIR          = TEMP_DIR / "НДС"
TEMP_PRESENTATION_DIR = TEMP_DIR / "Презентация"
TEMP_KU_DIR           = TEMP_DIR / "Возмещение КУ"

# Создаём папки (если их нет)
for _d in (OUTPUT_DIR, TEMP_DIR, TEMP_OPIU_DIR, TEMP_FORECAST_DIR,
           TEMP_VAT_DIR, TEMP_PRESENTATION_DIR, TEMP_KU_DIR):
    _d.mkdir(parents=True, exist_ok=True)


# ============================================================
# 2. ПЕРИОД ПО УМОЛЧАНИЮ
# ============================================================

_DEFAULT_YEAR  = 2026
_DEFAULT_MONTH = 9


# ============================================================
# 3. НАЗВАНИЯ МЕСЯЦЕВ
# ============================================================

MONTHS_RU = {
    1: "январь", 2: "февраль", 3: "март", 4: "апрель",
    5: "май", 6: "июнь", 7: "июль", 8: "август",
    9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь",
}


# ============================================================
# 4. ЛИСТ В ШАБЛОНЕ
# ============================================================

TEMPLATE_SHEET_NAME = "Лист1"


# ============================================================
# 5. ФЛАГИ
# ============================================================

# >>> В облаке DEBUG лучше выключить, чтобы не засорять логи.
# Если что-то не работает — включи DEBUG=True и посмотри логи в Streamlit Cloud.
DEBUG                = False
POSTPROCESS_FORMULAS = True
VERIFY_FORMULAS      = True

PRESERVE_AGGREGATE_FORMULAS = True
SHARE_DENOMINATOR_TITLE = "Расходы"
FORCE_OVERWRITE_FORMULAS = False


# ============================================================
# 6. КОНСТАНТЫ
# ============================================================

REF_OBJECTS_FILE     = REF_DIR / "Список Объектов Estate.xlsx"
REF_INVESTMENTS_FILE = REF_DIR / "Вложенные средства на объекты.xlsx"
REF_DEBTS_FILE       = REF_DIR / "таблица Долги и Численность сотрудников.xlsx"

TEMPLATE_FILES = {
    "Estate": TEMPLATES_DIR / "Шаблон Отчета Estate.xlsx",
    "Unelma": TEMPLATES_DIR / "Шаблон Отчета Unelma.xlsx",
    "Nomiqa": TEMPLATES_DIR / "Шаблон Отчета Nomiqa.xlsx",
}


# ============================================================
# 6b. НДС
# ============================================================
# Логика:
#   - Колонка C (факт прошлого месяца)  — НДС из файлов НДС за прошлый месяц
#   - Колонка E (план года из БДиР)     — НДС из БДиР, строка «Справочно по НДС выручки:»
#   - Колонка F (план месяца из Прогнозов) — НДС из Прогнозов, та же строка
#   - Колонка H (факт текущего месяца)  — НДС из файлов НДС за текущий месяц
#   - Колонка J (прогноз будущего)      — НДС из Прогнозов, та же строка
#
# Если месяц колонки < VAT_START_MONTH (сентябрь 2026) — НДС = 0,
# и «Выручка с НДС» = «Выручка без НДС».

# С какого месяца заполняем строку «Выручка с НДС» реальным НДС.
VAT_START_YEAR  = 2026
VAT_START_MONTH = 9

# Названия строк в шаблоне
VAT_ROW_TITLE        = "Выручка с НДС"
VAT_ROW_TITLE_NO_VAT = "Выручка без НДС"

# Строка со справочной суммой НДС в БДиР и Прогнозах
VAT_REF_ROW_TITLE = "Справочно по НДС выручки:"

# Имя листа в файлах НДС
VAT_SHEET_NAME = "Статьи свободного ввода в ОПиУ"

# Названия колонок в файлах НДС
VAT_COL_ITEM   = "Статья"
VAT_COL_MONTH  = "Месяц начисления"
VAT_COL_OBJECT = "Направление"
VAT_COL_VAT    = "НДС"

# Файлы НДС лежат в TEMP_VAT_DIR, имя любое, но:
#   - расширение .xlsx
#   - в имени есть слово "ндс" (без учёта регистра)
VAT_FILE_SUFFIX  = ".xlsx"
VAT_FILE_KEYWORD = "ндс"

# Правила вычитания НДС для отдельных блоков.
# НДС(блок) = НДС(агрегат) − НДС(вычитаемые объекты).
#
# Для блока «Коммерческие помещения LV без Матиса и Эспорта»:
# в БДиР справочная строка даёт НДС по «Коммерческие помещения LV»
# (это агрегат), из которого нужно вычесть Матису и Эспорту.
#
# Для «Estate EU» и «Baku-Nomiqa и Dibai-Nomiqa» справочная строка
# содержит уже готовую сумму — берём её целиком, ничего не вычитаем.
VAT_SPECIAL_BLOCKS = {
    "Коммерческие помещения LV без Матиса и Эспорта": {
        "aggregate": "Коммерческие помещения LV",
        "subtract": [
            "M81 - Matisa 81",
            "EKS_Esporta iela 12-113",
        ],
    },
    "Estate EU": {
        "aggregate": "Estate EU",
        "subtract": [],
    },
    "Baku-Nomiqa и Dibai-Nomiqa": {
        "aggregate": "Baku-Nomiqa и Dibai-Nomiqa",
        "subtract": [],
    },
}


DIRECTIONS          = ["Латвия", "Европа", "Estate_AZE", "Nomiqa", "UK_Estate", "Unelma"]
DIRECTIONS_BDR      = ["Латвия", "Европа", "Estate_AZE", "Nomiqa", "UK_Estate", "Unelma"]
DIRECTIONS_FORECAST = ["Латвия", "Европа", "Estate_AZE", "Nomiqa", "UK_Estate", "Unelma"]

DIRECTION_TO_FILENAME = {
    "Латвия":     "Латвия",
    "Европа":     "Европа",
    "Estate_AZE": "Estate AZE",
    "Nomiqa":     "Nomiqa",
    "UK_Estate":  "UK Estate",
    "Unelma":     "Unelma",
    "Estate":     "Estate",
}

BDR_DIRECTION_TO_FILENAME = {
    "Латвия":     "Латвия",
    "Европа":     "Estate EU",
    "Estate_AZE": "Estate AZE",
    "Nomiqa":     "Nomiqa",
    "Unelma":     "Unelma",
    "UK_Estate":  "UK Estate",
}

REPORT_KEYS = ["Estate", "Unelma", "Nomiqa"]

OPIU_FLAT_FILES = {"UK_Estate"}


# ============================================================
# 7. МАППИНГИ МЕСЯЦЕВ → КОЛОНОК
# ============================================================

FORECAST_MONTH_COL = {m: m + 1 for m in range(1, 13)}

BDR_MONTH_COL_2026 = {8: 1, 9: 2, 10: 3, 11: 4, 12: 5}
BDR_MONTH_COL_FULL = {m: m for m in range(1, 13)}

BDR_MONTH_COL = dict(BDR_MONTH_COL_2026)


# ============================================================
# 8. СЛОВАРИ ПУТЕЙ
# ============================================================

REPORT_YEAR  = _DEFAULT_YEAR
REPORT_MONTH = _DEFAULT_MONTH
PREV_MONTH   = None
PREV_YEAR    = None
NEXT_MONTH   = None
NEXT_YEAR    = None

REPORT_MONTH_RU = None
PREV_MONTH_RU   = None
NEXT_MONTH_RU   = None

OPIU_FILES     = {}
BDR_FILES      = {}
FORECAST_FILES = {}
OUTPUT_FILES   = {}

# Дополнительные пути для генератора презентации:
# ОПиУ-за-все-месяцы, БДиР, Прогнозы и выходной Excel с диаграммами.
PRESENTATION_OPIU_FILE       = None   # заполняется в set_period()
PRESENTATION_OUTPUT_FILE     = None   # заполняется в set_period()
PRESENTATION_BDR_FILE        = None   # заполняется в set_period()
PRESENTATION_FORECAST_FILE   = None   # заполняется в set_period()


# ============================================================
# 9. ФУНКЦИЯ ПЕРЕСБОРКИ ПУТЕЙ
# ============================================================

def _prev_month(year, month):
    return (year - 1, 12) if month == 1 else (year, month - 1)


def _next_month(year, month):
    return (year + 1, 1) if month == 12 else (year, month + 1)


def _opiu_path(direction, year, month):
    fname_dir = DIRECTION_TO_FILENAME.get(direction, direction)
    name = f"ОПиУ {month:02d}.{year}-{month:02d}.{year} {fname_dir}.xlsx"
    # >>> Читаем из ВРЕМЕННОЙ папки (туда app.py сохранит файлы пользователя)
    return TEMP_OPIU_DIR / name


def _presentation_opiu_path(year, month):
    """
    Путь к ОПиУ за много месяцев (используется в презентации).
    Формат: «ОПиУ 01.2026-09.2026.xlsx» — от января до отчётного месяца.
    """
    name = f"ОПиУ 01.{year}-{month:02d}.{year}.xlsx"
    return TEMP_OPIU_DIR / name


def _bdr_path(direction, year):
    fname_dir = BDR_DIRECTION_TO_FILENAME.get(direction, direction)
    if year <= 2026:
        name = f"БДиР 08.{year}-12.{year} {fname_dir}.xlsx"
    else:
        name = f"БДиР 01.{year}-12.{year} {fname_dir}.xlsx"
    return TEMP_FORECAST_DIR / name


def _forecast_path(direction, year):
    fname_dir = DIRECTION_TO_FILENAME.get(direction, direction)
    if year <= 2026:
        name = f"Прогнозы месячные {fname_dir}.xlsx"
    else:
        name = f"Прогнозы месячные {fname_dir} {year}.xlsx"
    return TEMP_FORECAST_DIR / name


def set_period(year, month):
    global REPORT_YEAR, REPORT_MONTH
    global PREV_MONTH, PREV_YEAR, NEXT_MONTH, NEXT_YEAR
    global REPORT_MONTH_RU, PREV_MONTH_RU, NEXT_MONTH_RU
    global OPIU_FILES, BDR_FILES, FORECAST_FILES, OUTPUT_FILES
    global BDR_MONTH_COL
    global PRESENTATION_OPIU_FILE, PRESENTATION_OUTPUT_FILE
    global PRESENTATION_BDR_FILE, PRESENTATION_FORECAST_FILE

    if not (1 <= month <= 12):
        raise ValueError(f"Месяц должен быть 1..12, получено: {month}")

    REPORT_YEAR  = year
    REPORT_MONTH = month
    PREV_YEAR, PREV_MONTH = _prev_month(year, month)
    NEXT_YEAR, NEXT_MONTH = _next_month(year, month)

    REPORT_MONTH_RU = MONTHS_RU[month]
    PREV_MONTH_RU   = MONTHS_RU[PREV_MONTH]
    NEXT_MONTH_RU   = MONTHS_RU[NEXT_MONTH]

    if year <= 2026:
        BDR_MONTH_COL = dict(BDR_MONTH_COL_2026)
    else:
        BDR_MONTH_COL = dict(BDR_MONTH_COL_FULL)

    OPIU_FILES.clear()
    for d in DIRECTIONS:
        OPIU_FILES[d] = {
            PREV_MONTH:   _opiu_path(d, PREV_YEAR, PREV_MONTH),
            REPORT_MONTH: _opiu_path(d, year, month),
        }

    OPIU_FILES["Estate_ЕДИНЫЙ"] = {
        PREV_MONTH:   _opiu_path("Estate", PREV_YEAR, PREV_MONTH),
        REPORT_MONTH: _opiu_path("Estate", year, month),
    }

    BDR_FILES.clear()
    for d in DIRECTIONS_BDR:
        BDR_FILES[d] = _bdr_path(d, year)

    FORECAST_FILES.clear()
    for d in DIRECTIONS_FORECAST:
        FORECAST_FILES[d] = _forecast_path(d, year)

    OUTPUT_FILES.clear()
    for key in REPORT_KEYS:
        fname = f"Отчет_для_собственника_{key}_{year}_{month:02d}.xlsx"
        OUTPUT_FILES[key] = OUTPUT_DIR / fname

    OUTPUT_FILES.clear()
    for key in REPORT_KEYS:
        fname = f"Отчет_для_собственника_{key}_{year}_{month:02d}.xlsx"
        OUTPUT_FILES[key] = OUTPUT_DIR / fname

    # Пути для генератора презентации
    PRESENTATION_OPIU_FILE = _presentation_opiu_path(year, month)
    PRESENTATION_OUTPUT_FILE = OUTPUT_DIR / f"Диаграммы для презентации {year}_{month:02d}.xlsx"

    # БДиР для презентации План/Факт — ЕДИНЫЙ сводный файл за год
    # (формат «БДиР 01.2026-12.2026.xlsx»). Пользователь загружает
    # его в отдельном загрузчике (см. app.py).
    if year <= 2026:
        PRESENTATION_BDR_FILE = TEMP_FORECAST_DIR / f"БДиР 01.{year}-12.{year}.xlsx"
    else:
        PRESENTATION_BDR_FILE = TEMP_FORECAST_DIR / f"БДиР 01.{year}-12.{year}.xlsx"

    # Прогнозы месячные в презентации не участвуют
    PRESENTATION_FORECAST_FILE = None

set_period(_DEFAULT_YEAR, _DEFAULT_MONTH)


# ============================================================
# 10. ПРОВЕРКА ВХОДНЫХ ФАЙЛОВ
# ============================================================

def _directions_used_in(template_name):
    opiu_dirs, bdr_dirs, forecast_dirs = set(), set(), set()

    def _add(block_name):
        if block_name in BLOCK_SOURCES:
            src = BLOCK_SOURCES[block_name]
            if src.get("opiu_file"):     opiu_dirs.add(src["opiu_file"])
            if src.get("bdr_file"):      bdr_dirs.add(src["bdr_file"])
            if src.get("forecast_file"): forecast_dirs.add(src["forecast_file"])

    for block in TEMPLATE_BLOCKS[template_name]:
        if block in BLOCK_SOURCES:
            _add(block)
        elif block in TEMPLATE_CONSOLIDATION:
            for sub in TEMPLATE_CONSOLIDATION[block]:
                _add(sub)

    return {"opiu": opiu_dirs, "bdr": bdr_dirs, "forecast": forecast_dirs}


def check_inputs(template_name):
    used = _directions_used_in(template_name)
    missing = {
        "template":  [],
        "opiu_prev": [],
        "opiu_cur":  [],
        "bdr":       [],
        "forecast":  [],
        "vat":       [],
    }

    tpl = TEMPLATE_FILES[template_name]
    if not tpl.exists():
        missing["template"].append(tpl)

    for d in used["opiu"]:
        p_prev = OPIU_FILES[d].get(PREV_MONTH)
        if p_prev and not p_prev.exists():
            missing["opiu_prev"].append(p_prev)
        p_cur = OPIU_FILES[d].get(REPORT_MONTH)
        if p_cur and not p_cur.exists():
            missing["opiu_cur"].append(p_cur)

    for d in used["bdr"]:
        p = BDR_FILES.get(d)
        if p and not p.exists():
            missing["bdr"].append(p)

    for d in used["forecast"]:
        p = FORECAST_FILES.get(d)
        if p and not p.exists():
            missing["forecast"].append(p)

    # НДС: файлы необязательны, но если их нет в папке —
    # предупреждаем (не блокируем сборку).
    try:
        vat_files = [
            f for f in os.listdir(TEMP_VAT_DIR)
            if f.lower().endswith(VAT_FILE_SUFFIX)
            and VAT_FILE_KEYWORD.lower() in f.lower()
            and not f.startswith("~$")
        ]
        if not vat_files:
            missing["vat"].append(TEMP_VAT_DIR / "(нет файлов НДС)")
    except Exception:
        missing["vat"].append(TEMP_VAT_DIR / "(нет файлов НДС)")

    return missing


def has_missing_inputs(template_name):
    missing = check_inputs(template_name)
    return any(missing[k] for k in missing)


# ============================================================
# 11. СПИСОК БЛОКОВ
# ============================================================

TEMPLATE_BLOCKS = {
    "Estate": [
        "Antonijas",
        "Чака",
        "Матиса",
        "Эспорта",
        "Коммерческие помещения LV без Матиса и Эспорта",
        "Коммерческие помещения LV все",
        "Estate LV",
        "Estate EU",
        "Estate AZE (Estate East, Алярбекова)",
        "Estate AZE (Baku Revelton)",
        "Estate East",
        "ESTATE КОНСОЛИДИРОВАННЫЙ",
        "УК",
    ],
    "Unelma": ["Unelma"],
    "Nomiqa": [
        "Baku-Nomiqa и Dibai-Nomiqa",
        "Europe-Nomiqa",
        "Nomiqa (консолидированный)",
    ],
}


# ============================================================
# 12. МАППИНГ: БЛОК → ИСТОЧНИК
# ============================================================

BLOCK_SOURCES = {
    "Antonijas": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": ["AN14 Антониас 14 (дом + парковка) [Latvia]"],
        "bdr_file": "Латвия",
        "bdr_objects": ["AN14 Антониас 14 (дом + парковка)"],
        "forecast_file": "Латвия",
        "forecast_sheet": "Antonijas",
    },
    "Чака": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": ["AC89 Чака 89 (дом + парковка) [Latvia]"],
        "bdr_file": "Латвия",
        "bdr_objects": ["AC89 Чака 89 (дом + парковка)"],
        "forecast_file": "Латвия",
        "forecast_sheet": "Чака",
    },
    "Матиса": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": ["M81 - Matisa 81 [Latvia]"],
        "bdr_file": "Латвия",
        "bdr_objects": ["M81 - Matisa 81"],
        "forecast_file": "Латвия",
        "forecast_sheet": "Матиса",
    },
    "Эспорта": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": ["EKS_Esporta iela 12-113 [Latvia]"],
        "bdr_file": "Латвия",
        "bdr_objects": ["EKS_Esporta iela 12-113"],
        "forecast_file": "Латвия",
        "forecast_sheet": "Эспорта",
    },
    "Коммерческие помещения LV без Матиса и Эспорта": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": [
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
        ],
        "bdr_file": "Латвия",
        "bdr_objects": [
            "D4 Парковка-Deglava4", "AC87 Гараж Чака",
            "B117 Бривибас, 117", "B78 Бривибас, 78",
            "C23 Цесу, 23", "DAR1_Darzauglu1",
            "DS1 Дзирнаву, 1", "G73 Гертрудес, 73",
            "H5 Хоспиталю", "MP1_Marupe",
            "OZ1 Озолниеки", "V22 К. Валдемара 22", "UK_Latvia",
        ],
        "forecast_file": "Латвия",
        "forecast_sheet": "Коммерческие помещения LV",
    },
    "Estate EU": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": [
            "DZ1_Dzibik1 [Europe]",
            "F6 Помещение в доме Будапешт [Europe]",
            "J91 Ялтская - Помещение маленькое [Europe]",
            "ML2 [Europe]",
            "OT1_Otovice Участок Свалка [Europe]",
            "TGM20-Masaryka20 [Europe]",
            "TGM45 Масарика - Bagel Lounge [Europe]",
            "UK_EU [Europe]",
        ],
        "bdr_file": "Европа",
        "bdr_objects": [
            "DZ1_Dzibik1", "F6 Помещение в доме Будапешт",
            "J91 Ялтская - Помещение маленькое", "ML2",
            "OT1_Otovice Участок Свалка", "TGM20-Masaryka20",
            "TGM45 Масарика - Bagel Lounge", "UK_EU",
        ],
        "forecast_file": "Европа",
        "forecast_sheet": "Estate EU",
    },
    "Estate AZE (Estate East, Алярбекова)": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": [
            "AL0 - AL0 Aliyarbekova0etaj [East-Восток]",
            "AL1 - AL1 Aliyarbekova1etaj [East-Восток]",
            "AL2 - AL2 Aliyarbekova2etaj [East-Восток]",
            "AL3 - AL3 Aliyarbekova3etaj [East-Восток]",
            "EG_Egoist [East-Восток]",
            "UKA - UK_AZ-Аренда [East-Восток]",
            "VD_Vidadi [East-Восток]",
        ],
        "bdr_file": "Estate_AZE",
        "bdr_objects": [
            "AL0 - AL0 Aliyarbekova0etaj", "AL1 - AL1 Aliyarbekova1etaj",
            "AL2 - AL2 Aliyarbekova2etaj", "AL3 - AL3 Aliyarbekova3etaj",
            "EG_Egoist", "UKA - UK_AZ-Аренда", "VD_Vidadi",
        ],
        "forecast_file": "Estate_AZE",
        "forecast_sheet": "Estate AZE (Baku Estate)",
    },
    "Estate AZE (Baku Revelton)": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": ["BIS - Baku, Icheri Sheher 1,2 [East-Восток]"],
        "bdr_file": "Estate_AZE",
        "bdr_objects": ["BIS - Baku, Icheri Sheher 1,2"],
        "forecast_file": "Estate_AZE",
        "forecast_sheet": "Estate AZE (Baku Revelton)",
    },
    "УК": {
        "opiu_file": "Estate_ЕДИНЫЙ",
        "opiu_objects": ["UK Estate"],
        "bdr_file": "UK_Estate",
        "bdr_objects": ["UK Estate"],
        "forecast_file": "UK_Estate",
        "forecast_sheet": "УК (UK Estate)",
    },
    "Unelma": {
        "opiu_file": "Unelma",
        "opiu_objects": ["UK_Unelma [Unelma]"],
        "bdr_file": "Unelma",
        "bdr_objects": ["UK_Unelma"],
        "forecast_file": "Unelma",
        "forecast_sheet": "Unelma",
    },
    "Baku-Nomiqa и Dibai-Nomiqa": {
        "opiu_file": "Nomiqa",
        "opiu_objects": ["BNQ_BAKU-Nomiqa [Nomiqa]", "DNQ_Dubai-Nomiqa [Nomiqa]"],
        "bdr_file": "Nomiqa",
        "bdr_objects": ["BNQ_BAKU-Nomiqa", "DNQ_Dubai-Nomiqa"],
        "forecast_file": "Nomiqa",
        "forecast_sheet": "Baku-Nomiqa и Dibai-Nomiqa",
    },
    "Europe-Nomiqa": {
        "opiu_file": "Nomiqa",
        "opiu_objects": ["ENQ_Europe-Nomiqa [Nomiqa]"],
        "bdr_file": "Nomiqa",
        "bdr_objects": ["ENQ_Europe-Nomiqa"],
        "forecast_file": "Nomiqa",
        "forecast_sheet": "Nomiqa-Europe",
    },
}


# ============================================================
# 13. КОНСОЛИДАЦИЯ
# ============================================================

TEMPLATE_CONSOLIDATION = {
    "Коммерческие помещения LV все": [
        "Коммерческие помещения LV без Матиса и Эспорта", "Матиса", "Эспорта",
    ],
    "Estate LV": ["Antonijas", "Чака", "Коммерческие помещения LV все"],
    "Estate East": [
        "Estate AZE (Estate East, Алярбекова)",
        "Estate AZE (Baku Revelton)",
    ],
    "ESTATE КОНСОЛИДИРОВАННЫЙ": ["Estate LV", "Estate EU", "Estate East"],
    "Nomiqa (консолидированный)": [
        "Baku-Nomiqa и Dibai-Nomiqa", "Europe-Nomiqa",
    ],
}

DETAILED_CONSOLIDATION = {
    "Estate LV": {
        "Выручка Антонияс": "Antonijas",
        "Выручка Чака": "Чака",
        "Выручка Коммерческие помещения": "Коммерческие помещения LV все",
    },
    "ESTATE КОНСОЛИДИРОВАННЫЙ": {
        "Estate LV": "Estate LV",
        "Estate EU": "Estate EU",
        "Estate AZE (Estate East, Алярбекова)": "Estate AZE (Estate East, Алярбекова)",
        "Estate AZE (Baku Revelton)": "Estate AZE (Baku Revelton)",
    },
}

SPECIAL_ROWS = {}


# ============================================================
# 13b. ДОПОЛНИТЕЛЬНАЯ КОНСОЛИДАЦИЯ СТРОК
# ============================================================

EXTRA_CONSOLIDATION = {
    "ESTATE КОНСОЛИДИРОВАННЫЙ": {
        "Прочие расходы": [
            "Estate LV",
            "Estate EU",
            "Estate AZE (Estate East, Алярбекова)",
            "Estate AZE (Baku Revelton)",
        ],
    },
}


# ============================================================
# 13c. ПОСТПРОЦЕССИНГ ФОРМУЛ ДОЛЕЙ В UNELMA
# ============================================================

UNELMA_DENOMINATOR_OLD = "$C$5"
UNELMA_DENOMINATOR_NEW = "$C$3"


# ============================================================
# 14. НАЗВАНИЯ СТРОК
# ============================================================

ROW_NAMES = {
    "revenue": "Выручка",
    "expenses": "Расходы",
    "fot": "Зарплата",
    "ndfl": "НДФЛ",
    "utilities": "1.2.10 Коммунальные платежи",
    "commission": "1.2.6 Комиссия внешним риэлторам и партнерам",
    "repair": "1.2.22 Ремонт текущий",
    "property_tax": "1.2.16.1 Налог на недвижимость",
    "rko": "1.2.17 РКО",
    "accountant": "1.2.12 Бухгалтер",
}


# ============================================================
# 15. МАППИНГ: СТРОКА ШАБЛОНА → ИСТОЧНИК
# ============================================================

TEMPLATE_ROW_MAP = {
    "Выручка от аренды": {
        "sum": [
            "Выручка от аренды",
            "Выручка от аренды Антониас",
            "Выручка от аренды ЧАКА",
            "Выручка от аренды Парковки Д4",
            "Выручка от аренды Матиса",
            "Выручка от аренды Муциниеку",
            "Выручка [сделки]",
            "Выручка UK_Latvia",
        ],
        "fallback": ["Аренда"],
    },
    "Выручка от возмещение коммунальных услуг": [
        "Выручка от возмещение коммунальных услуг",
        "Выручка от возмещения коммунальных услуг",
        "Возмещение коммунальных услуг",
    ],
    "Выручка от продажи недвижимости": [
        "Выручка от продажи недвижимости",
        "Выручка продажа недвижимости",
        "Выручка BNQ_BAKU-Nomiqa",
        "Выручка DNQ_Dubai-Nomiqa",
        "Выручка ENQ_Europe-Nomiqa",
    ],
    "Выручка по дополнительным услугам": ["Выручка по дополнительным услугам"],
    "Выручка строительных проектов": ["Выручка строительных проектов"],

    "Зарплата": {
        "sum": ["Зарплата"],
        "fallback": [
            "ФОТ", "- ФОТ", "ФОТ производственного персонала",
            "ФОТ коммерческого персонала", "ФОТ [сделки]",
        ],
    },
    "НДФЛ": ["НДФЛ", "Налоги", "- Налоги", "Взносы в фонды"],

    "1.2.10 Коммунальные платежи": [
        "1.2.10 Коммунальные платежи", "Коммунальные расходы",
    ],

    "1.2.6 Комиссия внешним риэлторам и партнерам": {
        "sum": [
            "1.2.6 Комиссия внешним риэлторам и партнерам",
            "Комиссия за сдачу",
            "Комиссия внешним риэлторам и партнерам",
        ],
    },

    "1.2.1 Закупка до 1000 евро": {
        "any": ["Закупки, обслуживание и ремонт:"],
        "sum": [
            "1.2.1.1 Бытовое оборудование, инвентарь",
            "1.2.1.2 Мебель",
            "1.2.1.3 Оргтехника",
            "1.2.1.4 Хоз.принадлежности",
            "1.2.1.5 Прочие мелкие ТМЦ",
            "1.2.8.1 Обслуживание объектов (бытовые вопросы, без ремонта)",
            "1.2.8.2 Страхование",
            "1.2.22 Ремонт текущий",
        ],
    },
    "1.2.1.1 Бытовое оборудование, инвентарь": ["1.2.1.1 Бытовое оборудование, инвентарь"],
    "1.2.1.2 Мебель": ["1.2.1.2 Мебель"],
    "1.2.1.3 Оргтехника": ["1.2.1.3 Оргтехника"],
    "1.2.1.4 Хоз.принадлежности": ["1.2.1.4 Хоз.принадлежности"],
    "1.2.1.5 Прочие мелкие ТМЦ": ["1.2.1.5 Прочие мелкие ТМЦ"],
    "1.2.8.1 Обслуживание объектов (бытовые вопросы, без ремонта)": [
        "1.2.8.1 Обслуживание объектов (бытовые вопросы, без ремонта)",
    ],
    "1.2.8.2 Страхование": ["1.2.8.2 Страхование"],
    "1.2.22 Ремонт текущий": ["1.2.22 Ремонт текущий", "Ремонт текущий"],

    "1.2.16.1 Налог на недвижимость": {
        "any": ["Налог на недвижимость"],
        "sum": ["1.2.16.1 Налог на недвижимость"],
    },
    "1.2.17 РКО": ["1.2.17 РКО", "РКО"],
    "1.2.12 Бухгалтер": ["1.2.12 Бухгалтер", "Бухгалтер"],

    "1.2.9.1 Связь , интернет, TV": {
        "sum": [
            "1.2.9.1 Связь , интернет, TV",
            "1.2.9.2 IT внедрение",
            "1.2.9.3 IT сервисы",
            "Интернет, программы,телефоны",
        ],
    },

    "1.2.2 Командировочные расходы": [
        "1.2.2 Командировочные расходы", "Командировки",
    ],

    "1.2.3.1 Платное продвижение - Instagram": {
        "sum": [
            "1.2.3.1 Платное продвижение  - Instagram",
            "1.2.3.2 Платное продвижение - Linkedin",
            "1.2.3.3 Платное продвижение  - Youtube",
            "1.2.3.4 Платное продвижение  - Google ads",
            "1.2.3.5 Платное продвижение  - Тестирование гипотезы",
            "1.2.3.6 Платное продвижение - TikTok",
            "1.2.4.1 Услуги подрядчиков по созданию контента для соц.сетей",
            "1.2.4.2 Контентный маркетинг - Instagram",
            "1.2.4.3 Контентный маркетинг - Linkedin",
            "1.2.4.4 Контентный маркетинг - Youtube",
            "1.2.4.5 Контентный маркетинг - SEO сайта",
            "1.2.4.6  Контентный маркетинг - Тестирование гипотезы",
            "1.2.4.7 Контентный маркетинг - Telegram",
            "1.2.29 Площадки для объявлений объектов",
            "1.2.32 Общие услуги подрядных организаций для маркетинга (рекламы, макеты, подрядчики)",
            "Объявление, реклама",
            "Маркетинговые расходы",
        ],
    },

    "1.2.5.2 Трансфер сотрудников, такси": [
        "1.2.5.2 Трансфер сотрудников, такси", "Такси",
    ],

    "1.2.5.4 Прочие расходы на персонал": {
        "sum": [
            "1.2.5.1 Поиск и найм персонала",
            "1.2.5.3  Оформление разрешения и прочих документов для сотрудников",
            "1.2.5.4 Прочие расходы на персонал",
            "1.2.5.5 Связь и интернет персонал",
            "Расходы на персонал",
        ],
    },
    "1.2.5.1 Поиск и найм персонала": ["1.2.5.1 Поиск и найм персонала"],
    "1.2.5.3  Оформление разрешения и прочих документов для сотрудников": [
        "1.2.5.3  Оформление разрешения и прочих документов для сотрудников",
    ],
    "1.2.5.5 Связь и интернет персонал": ["1.2.5.5 Связь и интернет персонал"],

    "1.2.7 Консультационные услуги": ["1.2.7 Консультационные услуги"],
    "1.2.9.2 IT внедрение": ["1.2.9.2 IT внедрение"],
    "1.2.9.3 IT сервисы": ["1.2.9.3 IT сервисы"],
    "1.2.11 Представительские расходы": [
        "1.2.11 Представительские расходы", "Представительские расходы",
    ],
    "1.2.13 Юридические услуги (Голландия)": ["1.2.13 Юридические услуги (Голландия)"],
    "1.2.14 Юридические услуги": ["1.2.14 Юридические услуги"],
    "1.2.16.2 Налог на прибыль": [
        "1.2.16.2 Налог на прибыль", "Налог на прибыль",
    ],
    "1.2.16.4 Туристический налог": ["1.2.16.4 Туристический налог"],
    "1.2.18 Банковская комиссия за использование ячейки": [
        "1.2.18 Банковская комиссия за использование ячейки", "Ячейки",
    ],

    "1.2.19 Расходы по содержанию автотранспорта": [
        "1.2.19 Расходы по содержанию автотранспорта",
        "Расходы по содержанию автотранспорта",
    ],

    "1.2.20 Лицензии, гос.пошлины, разрешения": [
        "1.2.20 Лицензии, гос.пошлины, разрешения",
        "Лицензии, госпошлины, разрешения",
    ],
    "1.2.21.1 Аренда офиса": ["1.2.21.1 Аренда офиса", "Аренда офиса"],
    "1.2.21.2 Административные офисные расходы": [
        "1.2.21.2 Административные офисные расходы",
        "Покупки в офис",
    ],
    "1.2.21.3 Коммунальные офисные расходы": [
        "1.2.21.3 Коммунальные офисные расходы",
    ],
    "1.2.23 Возвраты клиентам": ["1.2.23 Возвраты клиентам"],
    "1.2.25 Транспортные услуги": ["1.2.25 Транспортные услуги"],
    "1.2.29 Площадки для объявлений объектов": [
        "1.2.29 Площадки для объявлений объектов",
    ],
    "1.2.30 Сайты (разработка и обслуживание)": [
        "1.2.30 Сайты (разработка и обслуживание)",
    ],
    "1.2.31 Оффлайн мероприятия": ["1.2.31 Оффлайн мероприятия"],
    "1.2.32 Общие услуги подрядных организаций для маркетинга (рекламы, макеты, подрядчики)": [
        "1.2.32 Общие услуги подрядных организаций для маркетинга (рекламы, макеты, подрядчики)",
    ],
    "1.2.33 Непредвиденные расходы": [
        "1.2.33 Непредвиденные расходы", "Непредвиденные расходы",
    ],
    "1.2.34 Вознаграждение инвестора": [
        "1.2.34 Вознаграждение инвестора", "Вознаграждение инвестора",
    ],
    "1.2.36 Расходы по направлению": {
        "sum": [
            "1.2.36.1 Услуги субподрядчиков (строительные, инженерные)",
            "1.2.36.2 Аудит и технический надзор",
            "1.2.36.3 Лицензии и разрешения (строительные)",
            "1.2.36.4 Закупка оборудования \"умный дом\"",
            "1.2.36.5 Услуги по установке и программированию",
            "1.2.36.6 Тестирование и сертификация систем",
            "1.2.36.7 Закупки материалов и оборудования",
        ],
    },
}

# ============================================================
# 16. ФАЙЛЫ ВОЗМЕЩЕНИЯ КОММУНАЛЬНЫХ УСЛУГ (КУ)
# ============================================================
# Используются для слайдов 25, 27, 29 презентации.
# Пользователь загружает ДВА файла:
#   - «ОДДС 01.01.2026-30.09.2026 с НДС.xlsx»
#   - «ОДДС 01.01.2026-30.09.2026 без НДС.xlsx»
# Слайд 27 («с НДС кроме Чака 89») формируется программно
# из файла «с НДС» — из строки «Итого» вычитается строка «AC89 Чака».
#
# Имена файлов содержат диапазон дат, который меняется каждый месяц.
# Поэтому используем маски (glob) — ищем по ключевым словам.
def find_ku_file_with_vat():
    """Ищет в TEMP_KU_DIR файл «... с НДС.xlsx» (но не «без НДС» и не «кроме Чака»)."""
    if not TEMP_KU_DIR.exists():
        return None
    candidates = []
    for f in TEMP_KU_DIR.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        n = f.name.lower()
        if "без ндс" in n:
            continue
        if "кроме чака" in n:
            continue
        if "с ндс" in n:
            candidates.append(f)
    return candidates[0] if candidates else None


def find_ku_file_without_vat():
    """Ищет в TEMP_KU_DIR файл «... без НДС.xlsx»."""
    if not TEMP_KU_DIR.exists():
        return None
    for f in TEMP_KU_DIR.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        if "без ндс" in f.name.lower():
            return f
    return None

