# app.py
# ============================================================
# Streamlit-интерфейс «Отчет для собственника» (облачная версия).
#
# Две вкладки:
#   1) 📊 Отчёты для собственника — 3 отчёта (Estate, Unelma, Nomiqa).
#   2) 📈 Диаграммы для презентации — Excel с 80+ диаграммами.
#
# В каждой вкладке — крупные блоки загрузки файлов.
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

USERS = {
    "Alex": hashlib.sha256("AlEx_1708".encode("utf-8")).hexdigest(),
    "colleague": hashlib.sha256("AlEx_1708".encode("utf-8")).hexdigest(),
}

SESSION_KEY = "authenticated_user"


def _check_password(user: str, password: str) -> bool:
    h = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return USERS.get(user) == h


def _login_form():
    st.markdown(
        """
        <style>
        .login-card {
            max-width: 420px; margin: 5rem auto 0 auto;
            background: linear-gradient(135deg, rgba(30, 41, 59, 0.95) 0%, rgba(15, 23, 42, 0.95) 100%);
            border: 1px solid rgba(148, 163, 184, 0.2);
            border-radius: 16px; padding: 2rem 2rem 1.5rem 2rem;
            box-shadow: 0 12px 40px rgba(0,0,0,0.4);
        }
        .login-title {
            font-size: 1.6rem; font-weight: 800;
            background: linear-gradient(90deg, #60a5fa 0%, #22d3ee 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            text-align: center; margin-bottom: 0.3rem;
        }
        .login-sub {
            text-align: center; color: #94a3b8;
            font-size: 0.9rem; margin-bottom: 1.4rem;
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
    if st.session_state.get(SESSION_KEY):
        return st.session_state[SESSION_KEY]
    _login_form()
    st.stop()


current_user = require_login()


# ============================================================
# КАСТОМНЫЙ CSS
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

    h1, h2, h3, h4, h5, h6 { color: #f9fafb !important; }

    .main-title {
        font-size: 2.4rem; font-weight: 800;
        background: linear-gradient(90deg, #60a5fa 0%, #22d3ee 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text; margin: 0; line-height: 1.1;
    }
    .main-subtitle { font-size: 1rem; color: #94a3b8; margin-top: 0.35rem; }

    .section-title {
        font-size: 1.25rem; font-weight: 700; color: #f1f5f9;
        margin: 0.6rem 0 0.4rem 0;
    }
    .section-hint { color: #94a3b8; font-size: 0.85rem; margin-bottom: 1rem; }

    /* ===== КАРТОЧКИ МЕТРИК ===== */
    .metric-card {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.9) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(148, 163, 184, 0.15);
        border-radius: 14px; padding: 1.1rem 1.3rem;
        position: relative; overflow: hidden; height: 100%;
    }
    .metric-card::before {
        content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px;
        background: linear-gradient(180deg, #60a5fa, #22d3ee);
    }
    .metric-label {
        color: #94a3b8; font-size: 0.78rem; font-weight: 500;
        text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 0.35rem;
    }
    .metric-value { color: #f8fafc; font-size: 1.55rem; font-weight: 700; }
    .metric-sub { color: #64748b; font-size: 0.75rem; margin-top: 0.2rem; }

    .status-ok { border-color: rgba(34, 197, 94, 0.4) !important; }
    .status-ok::before { background: linear-gradient(180deg, #22c55e, #4ade80) !important; }
    .status-warn::before { background: linear-gradient(180deg, #f59e0b, #fbbf24) !important; }

    /* ===== ЗАГОЛОВКИ БЛОКОВ ЗАГРУЗКИ ===== */
    .upload-block-title {
        display: flex; align-items: center; gap: 0.8rem;
        font-size: 1.55rem; font-weight: 800; color: #f8fafc;
        letter-spacing: -0.015em;
        margin: 0;
        padding: 0.2rem 0;
    }
    .upload-block-icon {
        font-size: 2rem; line-height: 1;
        filter: drop-shadow(0 4px 10px rgba(96, 165, 250, 0.35));
    }
    .upload-block-desc {
        color: #94a3b8; font-size: 0.95rem;
        margin: 0.4rem 0 0.9rem 0;
    }
    .upload-block {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.85) 0%, rgba(15, 23, 42, 0.85) 100%);
        border: 1px solid rgba(148, 163, 184, 0.2);
        border-radius: 18px;
        padding: 1.5rem 1.7rem 1.3rem 1.7rem;
        margin-bottom: 1.4rem;
        position: relative;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.25);
    }
    .upload-block::before {
        content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 6px;
        border-radius: 18px 0 0 18px;
        background: linear-gradient(180deg, #60a5fa, #22d3ee, #a78bfa);
    }

    /* ===== БОЛЬШИЕ КНОПКИ / ВКЛАДКИ ===== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 1.2rem; justify-content: center;
        border-bottom: 2px solid rgba(148, 163, 184, 0.15);
        padding: 0.6rem 0 1.2rem 0;
        margin-bottom: 1.4rem;
    }
    .stTabs [data-baseweb="tab"] {
        background: linear-gradient(180deg,
            rgba(51, 65, 85, 0.85) 0%,
            rgba(30, 41, 59, 0.95) 100%);
        border: 1px solid rgba(148, 163, 184, 0.3);
        border-radius: 16px;
        padding: 1.1rem 2.2rem;
        color: #e2e8f0; font-weight: 800;
        transition: all 0.22s ease;
        min-height: 3.6rem;
        box-shadow:
            inset 0 1px 0 rgba(255, 255, 255, 0.1),
            inset 0 -2px 0 rgba(0, 0, 0, 0.25),
            0 6px 14px rgba(0, 0, 0, 0.35);
    }
    .stTabs [data-baseweb="tab"] p {
        font-size: 1.35rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        line-height: 1.2 !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background: linear-gradient(180deg,
            rgba(59, 130, 246, 0.35) 0%,
            rgba(37, 99, 235, 0.5) 100%);
        border-color: rgba(147, 197, 253, 0.8);
        transform: translateY(-3px) scale(1.015);
        box-shadow: 0 14px 28px rgba(37, 99, 235, 0.45);
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(180deg,
            #3b82f6 0%, #2563eb 50%, #1d4ed8 100%) !important;
        border-color: #60a5fa !important;
        color: #ffffff !important;
        transform: translateY(-3px) scale(1.015);
        box-shadow:
            inset 0 1px 0 rgba(255, 255, 255, 0.35),
            inset 0 -3px 0 rgba(0, 0, 0, 0.3),
            0 16px 36px rgba(37, 99, 235, 0.55),
            0 6px 16px rgba(96, 165, 250, 0.4);
    }
    .stTabs [aria-selected="true"] p {
        color: #ffffff !important;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0b1220 0%, #0f172a 100%);
        border-right: 1px solid rgba(148, 163, 184, 0.1);
    }

    .stButton > button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: #fff; border: none; border-radius: 10px;
        padding: 0.55rem 1.1rem; font-weight: 600;
    }
    .stButton > button:hover { filter: brightness(1.1); }

    .stDownloadButton > button {
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);
        color: #fff; border: none; border-radius: 10px;
        padding: 0.55rem 1.1rem; font-weight: 600; width: 100%;
    }

    div[data-baseweb="select"] > div {
        background: rgba(15, 23, 42, 0.8) !important;
        color: #f1f5f9 !important;
    }
    label { color: #cbd5e1 !important; }

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
# ВСПОМОГАТЕЛЬНОЕ: КРУПНЫЙ БЛОК ЗАГРУЗКИ
# ============================================================

def _upload_block_html(title, icon, desc):
    """HTML-карточка заголовка блока загрузки."""
    return f"""
    <div class="upload-block">
        <div class="upload-block-title">
            <span class="upload-block-icon">{icon}</span>
            <span>{title}</span>
        </div>
        <div class="upload-block-desc">{desc}</div>
    </div>
    """


# ============================================================
# ВЫБОР ПЕРИОДА (SIDEBAR)
# ============================================================

st.sidebar.markdown("### 📅 Период отчёта")

_year = st.sidebar.selectbox(
    "Год", options=_YEAR_OPTIONS, index=0, key="report_year",
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
st.sidebar.caption(f"📅 Период: **{_MONTH_NAMES[_month]} {_year}**")
st.sidebar.markdown("---")
st.sidebar.caption(f"👤 Пользователь: **{current_user}**")

if st.sidebar.button("Выйти", use_container_width=True):
    st.session_state.pop(SESSION_KEY, None)
    st.rerun()


# ============================================================
# ВСПОМОГАТЕЛЬНОЕ: ДИАГНОСТИКА
# ============================================================

def _file_status_table():
    result = {}
    for report_key in ("Estate", "Unelma", "Nomiqa"):
        try:
            missing = config.check_inputs(report_key)
        except Exception as e:
            result[report_key] = [("error", Path(str(e)))]
            continue
        flat = []
        for k in ("template", "opiu_prev", "opiu_cur", "bdr",
                  "forecast", "vat"):
            for p in missing.get(k, []):
                flat.append((k, Path(p)))
        result[report_key] = flat
    return result


def _save_uploaded(files, target_dir):
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


# ============================================================
# ЗАГОЛОВОК
# ============================================================

st.markdown(
    f"""
    <div style="padding: 0.5rem 0 1.0rem 0;">
        <div class="main-title">📊 Отчет для собственника</div>
        <div class="main-subtitle">
            Период: <b style="color:#93c5fd;">{_MONTH_NAMES[_month]} {_year}</b>
            · Estate · Unelma · Nomiqa
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# ДВЕ ВКЛАДКИ
# ============================================================

tab_reports, tab_presentation = st.tabs([
    "📊  Отчёты для собственника",
    "📈  Диаграммы для презентации",
])


# ============================================================
# ВКЛАДКА 1: ОТЧЁТЫ ДЛЯ СОБСТВЕННИКА
# ============================================================

with tab_reports:

    st.markdown(
        '<div class="section-title">📥 Шаг 1 · Загрузка файлов</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-hint">'
        'Загрузите файлы ОПиУ, БДиР, Прогнозов и НДС. '
        'Шаблоны уже встроены в программу — их загружать не нужно.'
        '</div>',
        unsafe_allow_html=True,
    )

    # --- ОПиУ ---
    st.markdown(
        _upload_block_html(
            "ОПиУ (отчёт о прибылях и убытках)",
            "📗",
            "Загрузите файлы ОПиУ за <b>прошлый</b> и <b>текущий</b> месяц "
            "для направлений: Латвия, Европа, Estate AZE, Nomiqa, "
            "UK Estate, Unelma, Estate (единый).",
        ),
        unsafe_allow_html=True,
    )
    opiu_files = st.file_uploader(
        "Выберите один или несколько файлов ОПиУ (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="upload_opiu",
        label_visibility="collapsed",
    )

    # --- БДиР ---
    st.markdown(
        _upload_block_html(
            "БДиР (бюджет доходов и расходов)",
            "📘",
            "Загрузите годовые файлы БДиР по направлениям: Латвия, "
            "Европа (Estate EU), Estate AZE, Nomiqa, UK Estate, Unelma.",
        ),
        unsafe_allow_html=True,
    )
    bdr_files = st.file_uploader(
        "Выберите один или несколько файлов БДиР (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="upload_bdr",
        label_visibility="collapsed",
    )

    # --- Прогнозы ---
    st.markdown(
        _upload_block_html(
            "Прогнозы месячные",
            "📙",
            "Загрузите файлы «Прогнозы месячные» по тем же направлениям, "
            "что и БДиР.",
        ),
        unsafe_allow_html=True,
    )
    forecast_files = st.file_uploader(
        "Выберите один или несколько файлов Прогнозов (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="upload_forecast",
        label_visibility="collapsed",
    )

    # --- НДС ---
    st.markdown(
        _upload_block_html(
            "НДС (файлы «НДС ММ.ГГГГ <Статья>.xlsx»)",
            "📕",
            "Загрузите файлы НДС за <b>прошлый</b> и <b>текущий</b> месяц. "
            "В имени файла должно быть слово «НДС» и расширение <code>.xlsx</code>.",
        ),
        unsafe_allow_html=True,
    )
    vat_files = st.file_uploader(
        "Выберите один или несколько файлов НДС (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="upload_vat",
        label_visibility="collapsed",
    )

    # --- Сохраняем загруженные файлы ---
    n_opiu     = _save_uploaded(opiu_files, config.TEMP_OPIU_DIR)
    n_bdr      = _save_uploaded(bdr_files, config.TEMP_FORECAST_DIR)
    n_forecast = _save_uploaded(forecast_files, config.TEMP_FORECAST_DIR)
    n_vat      = _save_uploaded(vat_files, config.TEMP_VAT_DIR)

    total_reports_uploaded = n_opiu + n_bdr + n_forecast + n_vat
    if total_reports_uploaded > 0:
        st.success(
            f"✅ Сохранено: ОПиУ — {n_opiu}, БДиР — {n_bdr}, "
            f"Прогнозов — {n_forecast}, НДС — {n_vat}."
        )

    # --- Карточки метрик ---
    status = _file_status_table()
    ready = sum(1 for k in status if not status[k])
    total_missing = sum(len(v) for v in status.values())

    st.markdown("<br>", unsafe_allow_html=True)
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
        cls = "metric-card status-ok" if ready == 3 else "metric-card"
        st.markdown(
            f"""
            <div class="{cls}">
                <div class="metric-label">✅ Готово к сборке</div>
                <div class="metric-value">{ready} / 3</div>
                <div class="metric-sub">все входные файлы на месте</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col4:
        if total_missing == 0:
            cls, val, sub = ("metric-card status-ok", "всё на месте", "можно собирать")
        else:
            cls = "metric-card status-warn"
            val = f"не хватает: {total_missing}"
            sub = "часть файлов отсутствует"
        st.markdown(
            f"""
            <div class="{cls}">
                <div class="metric-label">📥 Входные файлы</div>
                <div class="metric-value">{val}</div>
                <div class="metric-sub">{sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # --- Проверка готовности ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">🔍 Шаг 2 · Проверка готовности</div>',
        unsafe_allow_html=True,
    )

    for report_key in ("Estate", "Unelma", "Nomiqa"):
        missing = status[report_key]
        if not missing:
            st.success(f"**{report_key}**: все входные файлы на месте — можно собирать.")
        else:
            labels = {
                "template":  "Шаблон",
                "opiu_prev": "ОПиУ за прошлый месяц",
                "opiu_cur":  "ОПиУ за текущий месяц",
                "bdr":       "БДиР",
                "forecast":  "Прогноз месячный",
                "vat":       "НДС",
                "error":     "Ошибка проверки",
            }
            with st.expander(
                f"⚠️ {report_key}: не хватает {len(missing)} файлов",
                expanded=False,
            ):
                for k, p in missing:
                    st.text(f"• {labels.get(k, k)}: {p.name}")

    # --- Сборка отчётов ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">🚀 Шаг 3 · Сборка отчётов</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-hint">'
        'Соберите все три отчёта одной кнопкой или выборочно.'
        '</div>',
        unsafe_allow_html=True,
    )

    def _run_build(report_name):
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
                with st.expander("🔧 Подробности ошибки"):
                    st.code(traceback.format_exc())

    col_all, col_est, col_un, col_nom = st.columns([1.4, 1, 1, 1])

    with col_all:
        if st.button("🔨 Собрать все отчёты", type="primary",
                     use_container_width=True):
            _run_build("ALL")
    with col_est:
        if st.button("Estate", use_container_width=True, type="secondary"):
            _run_build("Estate")
    with col_un:
        if st.button("Unelma", use_container_width=True, type="secondary"):
            _run_build("Unelma")
    with col_nom:
        if st.button("Nomiqa", use_container_width=True, type="secondary"):
            _run_build("Nomiqa")

    # --- Скачивание отчётов ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">💾 Шаг 4 · Скачать готовые отчёты</div>',
        unsafe_allow_html=True,
    )

    col_dl_1, col_dl_2, col_dl_3 = st.columns(3)
    for col, key in zip((col_dl_1, col_dl_2, col_dl_3),
                        ("Estate", "Unelma", "Nomiqa")):
        path = config.OUTPUT_FILES[key]
        with col:
            if path.exists():
                with open(path, "rb") as f:
                    data = f.read()
                st.download_button(
                    label=f"⬇️ Скачать {key}",
                    data=data,
                    file_name=path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{key}",
                    use_container_width=True,
                )
                st.caption(f"`{path.name}`")
            else:
                st.info(f"Файл ещё не собран: `{path.name}`")


# ============================================================
# ВКЛАДКА 2: ДИАГРАММЫ ДЛЯ ПРЕЗЕНТАЦИИ
# ============================================================

with tab_presentation:

    st.markdown(
        '<div class="section-title">📥 Шаг 1 · Загрузка файлов</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-hint">'
        'Загрузите ОПиУ за все месяцы года, единый БДиР, ОДДС, файлы '
        'возмещения КУ и два справочника.'
        '</div>',
        unsafe_allow_html=True,
    )

    # --- ОПиУ за все месяцы ---
    st.markdown(
        _upload_block_html(
            "ОПиУ за все месяцы года",
            "📗",
            "Например: <code>ОПиУ 01.2026-09.2026.xlsx</code> — "
            "нужен для построения всех диаграмм презентации.",
        ),
        unsafe_allow_html=True,
    )
    opiu_full_year_files = st.file_uploader(
        "Выберите ОПиУ за все месяцы (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="upload_opiu_full_year",
        label_visibility="collapsed",
    )

    # --- Единый БДиР ---
    st.markdown(
        _upload_block_html(
            "БДиР единый сводный",
            "📘",
            "Например: <code>БДиР 01.2026-12.2026.xlsx</code> — "
            "нужен для диаграмм План/Факт.",
        ),
        unsafe_allow_html=True,
    )
    bdr_full_year_files = st.file_uploader(
        "Выберите БДиР за все месяцы (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="upload_bdr_full_year",
        label_visibility="collapsed",
    )

    # --- ОДДС ---
    st.markdown(
        _upload_block_html(
            "ОДДС (операционное сальдо и ДДС)",
            "📒",
            "Например: <code>ОДДС 01.01.2026-30.09.2026.xlsx</code> — "
            "для слайдов «Операционное сальдо» и «Анализ ДДС».",
        ),
        unsafe_allow_html=True,
    )
    odds_files = st.file_uploader(
        "Выберите ОДДС (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="upload_odds",
        label_visibility="collapsed",
    )

    # --- ОДДС для слайдов 25/27/29 ---
    st.markdown(
        _upload_block_html(
            "ОДДС для слайдов 25/27/29 (Возмещение КУ)",
            "📗",
            "Два файла ОДДС за весь год для таблиц возмещения КУ. "
            "В именах должны быть слова «с НДС» и «без НДС».",
        ),
        unsafe_allow_html=True,
    )
    ku_files = st.file_uploader(
        "Выберите файлы ОДДС для слайдов 25/27/29 (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="upload_ku",
        label_visibility="collapsed",
    )

    # --- Справочники ---
    st.markdown(
        _upload_block_html(
            "Справочник: Вложенные средства на объекты",
            "📓",
            "Файл <code>Вложенные средства на объекты.xlsx</code> — "
            "для расчёта ROI на слайде 70.",
        ),
        unsafe_allow_html=True,
    )
    investment_files = st.file_uploader(
        "Выберите файл вложений (.xlsx)",
        type=["xlsx"], accept_multiple_files=False, key="upload_investments",
        label_visibility="collapsed",
    )

    st.markdown(
        _upload_block_html(
            "Справочник: Долги и численность сотрудников",
            "📓",
            "Файл <code>таблица Долги и Численность сотрудников.xlsx</code> — "
            "для слайдов 14, 19, 24, 64.",
        ),
        unsafe_allow_html=True,
    )
    debts_files = st.file_uploader(
        "Выберите файл долгов и численности (.xlsx)",
        type=["xlsx"], accept_multiple_files=False, key="upload_debts",
        label_visibility="collapsed",
    )

    # --- Сохраняем ---
    n_opiu_full = _save_uploaded(opiu_full_year_files, config.TEMP_OPIU_DIR)
    n_bdr_full  = _save_uploaded(bdr_full_year_files, config.TEMP_FORECAST_DIR)
    n_odds      = _save_uploaded(odds_files, config.TEMP_FORECAST_DIR)
    n_ku        = _save_uploaded(ku_files, config.TEMP_KU_DIR)

    n_invest = 0
    if investment_files is not None:
        try:
            config.REF_DIR.mkdir(parents=True, exist_ok=True)
            inv_path = config.REF_DIR / "Вложенные средства на объекты.xlsx"
            with open(inv_path, "wb") as out:
                out.write(investment_files.getbuffer())
            n_invest = 1
        except Exception as e:
            st.warning(f"Не удалось сохранить файл вложений: {e}")

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

    total_pres_uploaded = (
        n_opiu_full + n_bdr_full + n_odds + n_ku + n_invest + n_debts
    )
    if total_pres_uploaded > 0:
        st.success(
            f"✅ Сохранено: ОПиУ (за год) — {n_opiu_full}, "
            f"БДиР (сводный) — {n_bdr_full}, ОДДС — {n_odds}, "
            f"Возмещение КУ — {n_ku}, Вложений — {n_invest}, "
            f"Долгов — {n_debts}."
        )

    # --- Сборка диаграмм ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">🚀 Шаг 2 · Сборка диаграмм</div>',
        unsafe_allow_html=True,
    )

    if st.button("📊 Собрать диаграммы для презентации",
                 type="primary", use_container_width=True,
                 key="build_presentation"):
        opiu_full_path = config.PRESENTATION_OPIU_FILE
        if not opiu_full_path or not opiu_full_path.exists():
            st.error(
                f"❌ Не найден файл ОПиУ за все месяцы: "
                f"`{opiu_full_path.name if opiu_full_path else '—'}`. "
                f"Загрузите его в блоке «ОПиУ за все месяцы года»."
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
                    with st.expander("🔧 Подробности ошибки"):
                        st.code(traceback.format_exc())

    # --- Скачивание ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">💾 Шаг 3 · Скачать готовый файл</div>',
        unsafe_allow_html=True,
    )

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
    <div style="text-align:center; color:#475569; font-size:0.8rem;">
        Временная папка: <code>{config.TEMP_DIR}</code> ·
        Streamlit {st.__version__}
    </div>
    """,
    unsafe_allow_html=True,
)