# app.py
# ============================================================
# Streamlit-интерфейс «Отчет для собственника» (облачная версия).
#
# ОТЛИЧИЯ ОТ ЛОКАЛЬНОЙ ВЕРСИИ:
#   1) Файлы ОПиУ, БДиР и Прогнозов пользователь загружает через
#      браузер (st.file_uploader). Программа сохраняет их во
#      временную папку и только потом запускает сборку.
#   2) Шаблоны и Справочники уже лежат в репозитории — их
#      загружать не нужно.
#   3) Пароли хранятся в st.secrets (для Streamlit Cloud) или
#      в переменных окружения (для локального запуска).
#   4) Готовые отчёты отдаются через st.download_button.
# ============================================================

import os
import hashlib
import traceback
from pathlib import Path

import streamlit as st

import config
from src.report_builder import build_all_reports, build_report
from src.presentation_builder import build_presentation_data


# ============================================================
# НАСТРОЙКИ СТРАНИЦЫ
# ============================================================

st.set_page_config(
    page_title="Отчет для собственника",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# АВТОРИЗАЦИЯ
# ============================================================
# Пароли хранятся в виде SHA-256 хешей.
#
# Где брать хеши:
#   python -c "import hashlib; print(hashlib.sha256(b'твой_пароль').hexdigest())"
#
# Где хранить:
#   - На Streamlit Cloud: в настройках приложения → Secrets (файл .streamlit/secrets.toml).
#   - Локально: можно просто вписать в словарь DEFAULT_USERS ниже.
#
# Формат секретов на Streamlit Cloud (в Secrets):
#   [users]
#   Alex = "хеш_пароля_для_Alex"
#   colleague = "хеш_пароля_для_colleague"
# ============================================================

# Запасной вариант — используется, если секреты не настроены.
# ВНИМАНИЕ: это публичный хеш от пароля "AlEx_1708".
# Для реальной работы настрой секреты на Streamlit Cloud!
DEFAULT_USERS = {
    "Alex": hashlib.sha256("AlEx_1708".encode("utf-8")).hexdigest(),
    "colleague": hashlib.sha256("AlEx_1708".encode("utf-8")).hexdigest(),
}


def _get_users():
    """Возвращает словарь {логин: хеш_пароля}. Сначала пробует st.secrets, потом DEFAULT_USERS."""
    try:
        if "users" in st.secrets:
            return {k: str(v) for k, v in st.secrets["users"].items()}
    except Exception:
        # st.secrets может быть недоступен локально — это нормально.
        pass
    return DEFAULT_USERS


SESSION_KEY = "authenticated_user"


def _check_password(user: str, password: str) -> bool:
    h = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return _get_users().get(user) == h


def _login_form():
    """Показывает форму входа. Возвращает имя пользователя, если вход удался."""
    st.markdown(
        """
        <style>
        .login-card {
            max-width: 420px;
            margin: 5rem auto 0 auto;
            background: linear-gradient(135deg, rgba(30, 41, 59, 0.95) 0%, rgba(15, 23, 42, 0.95) 100%);
            border: 1px solid rgba(148, 163, 184, 0.2);
            border-radius: 16px;
            padding: 2rem 2rem 1.5rem 2rem;
            box-shadow: 0 12px 40px rgba(0,0,0,0.4);
        }
        .login-title {
            font-size: 1.6rem;
            font-weight: 800;
            background: linear-gradient(90deg, #60a5fa 0%, #22d3ee 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-align: center;
            margin-bottom: 0.3rem;
        }
        .login-sub {
            text-align: center;
            color: #94a3b8;
            font-size: 0.9rem;
            margin-bottom: 1.4rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col_a, col_b, col_c = st.columns([1, 1.2, 1])
    with col_b:
        st.markdown(
            """
            <div class="login-card">
                <div class="login-title">📊 Отчет для собственника</div>
                <div class="login-sub">Введите логин и пароль</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form", clear_on_submit=False):
            user = st.text_input("Логин", key="login_user")
            pwd  = st.text_input("Пароль", type="password", key="login_pwd")
            submitted = st.form_submit_button("Войти", use_container_width=True)

        if submitted:
            if not user or not pwd:
                st.error("Заполните оба поля")
            elif _check_password(user, pwd):
                st.session_state[SESSION_KEY] = user
                st.rerun()
            else:
                st.error("Неверный логин или пароль")


def require_login():
    """Проверяет вход. Если пользователь не залогинен — показывает форму и останавливает приложение."""
    if st.session_state.get(SESSION_KEY):
        return st.session_state[SESSION_KEY]
    _login_form()
    st.stop()


current_user = require_login()


# ============================================================
# КАСТОМНЫЙ CSS — ОФОРМЛЕНИЕ
# ============================================================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    .stApp {
        background:
            radial-gradient(circle at 0% 0%, #1e3a8a 0%, transparent 45%),
            radial-gradient(circle at 100% 0%, #0f766e 0%, transparent 45%),
            linear-gradient(180deg, #0f172a 0%, #111827 100%);
        color: #e5e7eb;
    }

    h1, h2, h3, h4, h5, h6 { color: #f9fafb !important; letter-spacing: -0.02em; }

    .main-title {
        font-size: 2.4rem; font-weight: 800;
        background: linear-gradient(90deg, #60a5fa 0%, #22d3ee 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text; margin: 0; line-height: 1.1;
    }
    .main-subtitle { font-size: 1rem; color: #94a3b8; margin-top: 0.35rem; font-weight: 400; }

    .section-title {
        font-size: 1.25rem; font-weight: 700; color: #f1f5f9;
        margin: 0.6rem 0 0.4rem 0; display: flex; align-items: center; gap: 0.5rem;
    }
    .section-hint { color: #94a3b8; font-size: 0.85rem; margin-bottom: 1rem; }

    .metric-card {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.9) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(148, 163, 184, 0.15);
        border-radius: 14px; padding: 1.1rem 1.3rem;
        position: relative; overflow: hidden;
        transition: transform 0.18s ease, border-color 0.18s ease;
        height: 100%;
    }
    .metric-card:hover { transform: translateY(-2px); border-color: rgba(96, 165, 250, 0.5); }
    .metric-card::before {
        content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px;
        background: linear-gradient(180deg, #60a5fa, #22d3ee);
    }
    .metric-label {
        color: #94a3b8; font-size: 0.78rem; font-weight: 500;
        text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 0.35rem;
    }
    .metric-value { color: #f8fafc; font-size: 1.55rem; font-weight: 700; line-height: 1.2; }
    .metric-sub { color: #64748b; font-size: 0.75rem; margin-top: 0.2rem; }

    .status-ok { border-color: rgba(34, 197, 94, 0.4) !important; }
    .status-ok::before { background: linear-gradient(180deg, #22c55e, #4ade80) !important; }
    .status-warn::before { background: linear-gradient(180deg, #f59e0b, #fbbf24) !important; }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0b1220 0%, #0f172a 100%);
        border-right: 1px solid rgba(148, 163, 184, 0.1);
    }
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #e2e8f0 !important; font-size: 1rem;
        letter-spacing: 0.05em; text-transform: uppercase;
    }

    .stButton > button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: #fff; border: none; border-radius: 10px;
        padding: 0.55rem 1.1rem; font-weight: 600; font-size: 0.92rem;
        transition: transform 0.15s ease, box-shadow 0.15s ease, filter 0.15s ease;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25);
    }
    .stButton > button:hover {
        transform: translateY(-1px); filter: brightness(1.1);
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.45);
    }
    .stButton > button:active { transform: translateY(0); }
    .stButton > button[kind="secondary"] {
        background: rgba(148, 163, 184, 0.12); color: #e2e8f0;
        box-shadow: none; border: 1px solid rgba(148, 163, 184, 0.25);
    }
    .stButton > button[kind="secondary"]:hover {
        background: rgba(148, 163, 184, 0.2);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    }

    .stDownloadButton > button {
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);
        color: #fff; border: none; border-radius: 10px;
        padding: 0.55rem 1.1rem; font-weight: 600; font-size: 0.92rem; width: 100%;
        transition: transform 0.15s ease, filter 0.15s ease;
        box-shadow: 0 4px 12px rgba(13, 148, 136, 0.25);
    }
    .stDownloadButton > button:hover { transform: translateY(-1px); filter: brightness(1.1); }

    details[data-testid="stExpander"] {
        background: rgba(15, 23, 42, 0.6);
        border: 1px solid rgba(148, 163, 184, 0.15);
        border-radius: 12px; margin-bottom: 0.6rem; overflow: hidden;
        transition: border-color 0.18s ease;
    }
    details[data-testid="stExpander"]:hover { border-color: rgba(96, 165, 250, 0.35); }
    details[data-testid="stExpander"] summary { color: #e2e8f0 !important; font-weight: 600; padding: 0.75rem 1rem; }

    div[data-baseweb="select"] > div {
        background: rgba(15, 23, 42, 0.8) !important;
        border-color: rgba(148, 163, 184, 0.25) !important;
        color: #f1f5f9 !important; border-radius: 10px !important;
    }
    div[data-baseweb="select"] svg { fill: #94a3b8 !important; }
    label { color: #cbd5e1 !important; font-weight: 500 !important; font-size: 0.85rem !important; }

    hr { border-color: rgba(148, 163, 184, 0.12); }
    .stSpinner > div { border-top-color: #60a5fa !important; }
    div[data-testid="stAlert"] { border-radius: 10px; border: 1px solid rgba(148, 163, 184, 0.2); }

    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header[data-testid="stHeader"] { background: transparent; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# СПРАВОЧНИКИ МЕСЯЦЕВ
# ============================================================

_MONTH_NAMES = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}

_YEAR_OPTIONS = [2026, 2027]


# ============================================================
# ВЫБОР ПЕРИОДА (SIDEBAR)
# ============================================================

st.sidebar.markdown("### Период отчёта")

_year = st.sidebar.selectbox(
    "Год",
    options=_YEAR_OPTIONS,
    index=0,
    key="report_year",
)

_month = st.sidebar.selectbox(
    "Месяц",
    options=list(_MONTH_NAMES.keys()),
    format_func=lambda m: _MONTH_NAMES[m],
    index=config.REPORT_MONTH - 1,
    key="report_month",
)

config.set_period(_year, _month)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"📅 Период: **{_MONTH_NAMES[_month]} {_year}**  \n"
    f"📄 Отчётов: **3** (Estate · Unelma · Nomiqa)"
)

st.sidebar.markdown("---")
st.sidebar.caption(f"👤 Пользователь: **{current_user}**")
if st.sidebar.button("Выйти", use_container_width=True):
    st.session_state.pop(SESSION_KEY, None)
    st.rerun()


# ============================================================
# ДИАГНОСТИКА ПО ФАЙЛАМ
# ============================================================

def _file_status_table():
    result = {}
    for report_key in ("Estate", "Unelma", "Nomiqa"):
        missing = config.check_inputs(report_key)
        flat = []
        for k in ("opiu_prev", "opiu_cur", "bdr", "forecast", "vat"):
            for p in missing.get(k, []):
                flat.append((k, Path(p)))
        result[report_key] = flat
    return result


# ============================================================
# ЗАГОЛОВОК
# ============================================================

st.markdown(
    f"""
    <div style="padding: 0.5rem 0 0.5rem 0;">
        <div class="main-title">📊 Отчет для собственника</div>
        <div class="main-subtitle">
            Период: <b style="color:#93c5fd;">{_MONTH_NAMES[_month]} {_year}</b> · Подготовка финансовых отчётов Estate · Unelma · Nomiqa
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# ВЕРХНЯЯ ПАНЕЛЬ С КРАТКОЙ СТАТИСТИКОЙ (КАРТОЧКИ)
# ============================================================

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">📅 Период</div>
            <div class="metric-value">{_MONTH_NAMES[_month]} {_year}</div>
            <div class="metric-sub">отчётный месяц</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col2:
    st.markdown(
        """
        <div class="metric-card">
            <div class="metric-label">📄 Всего отчётов</div>
            <div class="metric-value">3</div>
            <div class="metric-sub">Estate · Unelma · Nomiqa</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col3:
    # Показываем реальное количество загруженных файлов.
    # Считаем во временной папке — так счётчик не сбрасывается
    # при повторном рендере страницы.
    try:
        _n_opiu_now = len(list(config.TEMP_OPIU_DIR.glob("*.xlsx")))
        _n_fc_now   = len(list(config.TEMP_FORECAST_DIR.glob("*.xlsx")))
        _n_vat_now  = len(list(config.TEMP_VAT_DIR.glob("*.xlsx")))
        _uploaded_now = _n_opiu_now + _n_fc_now + _n_vat_now
    except Exception:
        _uploaded_now = 0

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">📤 Загружено файлов</div>
            <div class="metric-value">{_uploaded_now}</div>
            <div class="metric-sub">ОПиУ · БДиР · Прогнозы · НДС</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col4:
    st.markdown(
        """
        <div class="metric-card status-warn">
            <div class="metric-label">⏳ Статус</div>
            <div class="metric-value">Ожидание</div>
            <div class="metric-sub">загрузите файлы для сборки</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.markdown("<br>", unsafe_allow_html=True)


# ============================================================
# БЛОК 1. ЗАГРУЗКА ФАЙЛОВ
# ============================================================

st.markdown(
    '<div class="section-title">📥 Шаг 1 · Загрузка файлов</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="section-hint">'
    'Загрузите файлы ОПиУ, БДиР, Прогнозов и НДС. '
    'Шаблоны и справочники уже встроены в программу — их загружать не нужно.'
    '</div>',
    unsafe_allow_html=True,
)

with st.expander("📗 ОПиУ (отчёт о прибылях и убытках)", expanded=True):
    st.caption(
        "Загрузите файлы ОПиУ за **прошлый месяц** и за **текущий месяц** "
        "для всех направлений: Латвия, Европа, Estate AZE, Nomiqa, UK Estate, Unelma, Estate (единый)."
    )
    opiu_files = st.file_uploader(
        "Выберите один или несколько файлов ОПиУ (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_opiu",
    )

    st.caption(
        "**Дополнительно:** загрузите ОПиУ за **все месяцы года** "
        "(например, `ОПиУ 01.2026-09.2026.xlsx`) — это нужно для генератора диаграмм презентации."
    )
    opiu_full_year_files = st.file_uploader(
        "Выберите ОПиУ за все месяцы (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_opiu_full_year",
    )

with st.expander("📘 БДиР (бюджет доходов и расходов)", expanded=True):
    st.caption(
        "Загрузите годовые файлы БДиР для направлений: Латвия, Европа (Estate EU), "
        "Estate AZE, Nomiqa, UK Estate, Unelma."
    )
    bdr_files = st.file_uploader(
        "Выберите один или несколько файлов БДиР (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_bdr",
    )

    st.caption(
        "**Дополнительно для презентации:** загрузите **единый сводный** БДиР "
        "за весь год — `БДиР 01.ГГГГ-12.ГГГГ.xlsx`. Он используется для диаграмм "
        "«План/Факт» в презентации."
    )
    bdr_full_year_files = st.file_uploader(
        "Выберите БДиР за все месяцы (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_bdr_full_year",
    )

    st.caption(
        "**Дополнительно для презентации:** загрузите **ОДДС** за весь год — "
        "`ОДДС 01.01.ГГГГ-ММ.ГГГГ.xlsx`. Он используется для таблиц "
        "«Операционное сальдо»."
    )
    odds_files = st.file_uploader(
        "Выберите ОДДС (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_odds",
    )

    st.caption(
        "**Дополнительно для презентации:** загрузите файл вложений "
        "`Вложенные средства на объекты.xlsx` — он используется для "
        "расчёта ROI на слайде 70."
    )
    investment_files = st.file_uploader(
        "Выберите файл вложений (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        key="upload_investments",
    )

    st.caption(
        "**Дополнительно для презентации:** загрузите файл "
        "`таблица Долги и Численность сотрудников.xlsx` — он используется "
        "для слайдов 14, 19, 24, 64 (долги и численность сотрудников)."
    )
    debts_files = st.file_uploader(
        "Выберите файл долгов и численности (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        key="upload_debts",
    )

with st.expander("📙 Прогнозы месячные", expanded=True):
    st.caption(
        "Загрузите файлы «Прогнозы месячные» для тех же направлений, что и БДиР."
    )
    forecast_files = st.file_uploader(
        "Выберите один или несколько файлов Прогнозов (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_forecast",
    )

with st.expander("📗 ОДДС для слайдов 25/27/29 (Возмещение КУ)", expanded=True):
    st.caption(
        "Загрузите **два файла ОДДС за весь год** — они используются "
        "для слайдов 25, 27, 29 презентации (таблицы возмещения КУ).\n"
        "Имена файлов могут быть такими (главное, чтобы в имени были "
        "слова «с НДС» и «без НДС»):\n"
        "- `ОДДС 01.01.2026-30.09.2026 с НДС.xlsx`\n"
        "- `ОДДС 01.01.2026-30.09.2026 без НДС.xlsx`\n\n"
        "Слайд 27 («с НДС кроме Чака 89») формируется программно — "
        "из файла «с НДС» исключается строка «AC89 Чака»."
    )
    ku_files = st.file_uploader(
        "Выберите файлы ОДДС для слайдов 25/27/29 (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_ku",
    )

with st.expander("📕 НДС (файлы «НДС ММ.ГГГГ <Статья>.xlsx»)", expanded=True):
    st.caption(
        "Загрузите файлы НДС **за прошлый месяц** (для колонки C) "
        "и **за текущий месяц** (для колонки H). "
        "Файлы должны содержать в имени слово «НДС» и расширение `.xlsx`. "
        "Например: `НДС 09.2026 Выручка от аренды.xlsx`."
    )
    vat_files = st.file_uploader(
        "Выберите один или несколько файлов НДС (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_vat",
    )


# ============================================================
# СОХРАНЕНИЕ ЗАГРУЖЕННЫХ ФАЙЛОВ ВО ВРЕМЕННУЮ ПАПКУ
# ============================================================

def _save_uploaded(files, target_dir):
    """Сохраняет загруженные файлы в указанную папку. Возвращает количество."""
    if not files:
        return 0
    target_dir.mkdir(parents=True, exist_ok=True)
    saved = 0
    for f in files:
        try:
            path = target_dir / f.name
            with open(path, "wb") as out:
                out.write(f.getbuffer())
            saved += 1
        except Exception as e:
            st.warning(f"Не удалось сохранить {f.name}: {e}")
    return saved


n_opiu = _save_uploaded(opiu_files, config.TEMP_OPIU_DIR)
n_opiu_full = _save_uploaded(opiu_full_year_files, config.TEMP_OPIU_DIR)
n_bdr = _save_uploaded(bdr_files, config.TEMP_FORECAST_DIR)
n_bdr_full = _save_uploaded(bdr_full_year_files, config.TEMP_FORECAST_DIR)
n_odds = _save_uploaded(odds_files, config.TEMP_FORECAST_DIR)
n_invest = _save_uploaded(
    [investment_files] if investment_files else [],
    config.REF_DIR,
)
n_debts = 0
if debts_files is not None:
    try:
        config.REF_DIR.mkdir(parents=True, exist_ok=True)
        debts_path = config.REF_DIR / "таблица Долги и Численность сотрудников.xlsx"
        with open(debts_path, "wb") as out:
            out.write(debts_files.getbuffer())
        n_debts = 1
    except Exception as e:
        st.warning(f"Не удалось сохранить файл долгов: {e}")
n_forecast = _save_uploaded(forecast_files, config.TEMP_FORECAST_DIR)
n_ku = _save_uploaded(ku_files, config.TEMP_KU_DIR)
n_vat = _save_uploaded(vat_files, config.TEMP_VAT_DIR)

total_uploaded = (n_opiu + n_opiu_full + n_bdr + n_bdr_full +
                  n_odds + n_invest + n_debts + n_forecast + n_ku + n_vat)

if total_uploaded > 0:
    st.success(
        f"✅ Сохранено во временную папку: ОПиУ (по месяцам) — {n_opiu}, "
        f"ОПиУ (за все месяцы) — {n_opiu_full}, "
        f"БДиР (по направлениям) — {n_bdr}, "
        f"БДиР (единый) — {n_bdr_full}, "
        f"ОДДС — {n_odds}, "
        f"Вложений — {n_invest}, "
        f"Долгов — {n_debts}, "
        f"Прогнозов — {n_forecast}, "
        f"Возмещение КУ — {n_ku}, НДС — {n_vat}."
    )


# ============================================================
# БЛОК 2. ПРОВЕРКА ГОТОВНОСТИ
# ============================================================

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">🔍 Шаг 2 · Проверка готовности</div>',
    unsafe_allow_html=True,
)

status = _file_status_table()

for report_key in ("Estate", "Unelma", "Nomiqa"):
    missing = status[report_key]
    if not missing:
        st.success(f"**{report_key}**: все входные файлы на месте — можно собирать.")
    else:
        labels = {
            "opiu_prev": "ОПиУ за прошлый месяц",
            "opiu_cur":  "ОПиУ за текущий месяц",
            "bdr":       "БДиР",
            "forecast":  "Прогноз месячный",
            "vat":       "НДС",
        }
        with st.expander(f"⚠️ {report_key}: не хватает {len(missing)} файлов", expanded=False):
            for k, p in missing:
                st.text(f"• {labels.get(k, k)}: {p.name}")


# ============================================================
# БЛОК 3. КНОПКА СБОРКИ
# ============================================================

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">🚀 Шаг 3 · Сборка отчётов</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="section-hint">Соберите все три отчёта одной кнопкой или выборочно.</div>',
    unsafe_allow_html=True,
)

col_build_all, col_build_estate, col_build_unelma, col_build_nomiqa = st.columns([1.4, 1, 1, 1])


def _run_build(report_name):
    """Обёртка для сборки с показом ошибок."""
    with st.spinner(f"Собираю {report_name}..."):
        try:
            if report_name == "ALL":
                build_all_reports()
                st.success("Готово! Все три отчёта собраны.")
            else:
                build_report(report_name)
                st.success(f"{report_name} собран.")
        except Exception as e:
            st.error(f"Ошибка при сборке {report_name}: {e}")
            with st.expander("🔧 Подробности ошибки (для разработчика)"):
                st.code(traceback.format_exc())


with col_build_all:
    if st.button("🔨 Собрать все отчёты", type="primary", use_container_width=True):
        _run_build("ALL")

with col_build_estate:
    if st.button("Estate", use_container_width=True, type="secondary"):
        _run_build("Estate")

with col_build_unelma:
    if st.button("Unelma", use_container_width=True, type="secondary"):
        _run_build("Unelma")

with col_build_nomiqa:
    if st.button("Nomiqa", use_container_width=True, type="secondary"):
        _run_build("Nomiqa")


# ============================================================
# БЛОК 4. СКАЧИВАНИЕ ГОТОВЫХ ОТЧЁТОВ
# ============================================================

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">💾 Шаг 4 · Скачать готовые отчёты</div>',
    unsafe_allow_html=True,
)

col_dl_1, col_dl_2, col_dl_3 = st.columns(3)

for col, report_key in zip(
    (col_dl_1, col_dl_2, col_dl_3), ("Estate", "Unelma", "Nomiqa")
):
    path = config.OUTPUT_FILES[report_key]
    with col:
        if path.exists():
            with open(path, "rb") as f:
                data = f.read()
            st.download_button(
                label=f"⬇️ Скачать {report_key}",
                data=data,
                file_name=path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{report_key}",
                use_container_width=True,
            )
            st.caption(f"`{path.name}`")
        else:
            st.info(f"Файл ещё не собран: `{path.name}`")


# ============================================================
# БЛОК 5. ПРЕЗЕНТАЦИЯ (ДИАГРАММЫ В EXCEL)
# ============================================================

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">📊 Шаг 5 · Диаграммы для презентации</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="section-hint">'
    'Excel-файл с диаграммами для вставки в презентацию. '
    'Использует ОПиУ за все месяцы года (загрузите его в Шаге 1).'
    '</div>',
    unsafe_allow_html=True,
)

if st.button("📊 Собрать диаграммы для презентации",
             type="primary", use_container_width=True,
             key="build_presentation"):
    if not config.PRESENTATION_OPIU_FILE or \
       not config.PRESENTATION_OPIU_FILE.exists():
        st.error(
            f"❌ Не найден файл ОПиУ за все месяцы: "
            f"`{config.PRESENTATION_OPIU_FILE.name if config.PRESENTATION_OPIU_FILE else '—'}`. "
            f"Загрузите его в Шаге 1 (блок ОПиУ → «ОПиУ за все месяцы»)."
        )
    else:
        with st.spinner("Собираю диаграммы..."):
            try:
                build_presentation_data(
                    opiu_path=config.PRESENTATION_OPIU_FILE,
                    bdr_path=config.PRESENTATION_BDR_FILE,
                    forecast_path=config.PRESENTATION_FORECAST_FILE,
                    year=config.REPORT_YEAR,
                    month=config.REPORT_MONTH,
                    output_path=config.PRESENTATION_OUTPUT_FILE,
                )
                st.success("Готово! Файл с диаграммами собран.")
            except Exception as e:
                st.error(f"Ошибка при сборке диаграмм: {e}")
                with st.expander("🔧 Подробности ошибки (для разработчика)"):
                    st.code(traceback.format_exc())

# Кнопка скачивания
pres_path = config.PRESENTATION_OUTPUT_FILE
if pres_path and pres_path.exists():
    with open(pres_path, "rb") as f:
        pres_data = f.read()
    st.download_button(
        label="⬇️ Скачать «Диаграммы для презентации»",
        data=pres_data,
        file_name=pres_path.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_presentation",
        use_container_width=True,
    )
    st.caption(f"`{pres_path.name}`")
else:
    st.info(
        f"Файл ещё не собран: `{pres_path.name if pres_path else '—'}`"
    )


# ============================================================
# ПОДВАЛ
# ============================================================

st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(
    f"""
    <div style="text-align:center; color:#475569; font-size:0.8rem; padding: 0.5rem;">
        Временная папка: <code style="color:#94a3b8;">{config.TEMP_DIR}</code> ·
        Выходная папка: <code style="color:#94a3b8;">{config.OUTPUT_DIR.name}</code> ·
        Streamlit {st.__version__}
    </div>
    """,
    unsafe_allow_html=True,
)