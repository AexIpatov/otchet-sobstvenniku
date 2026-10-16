# app.py
# ============================================================
# Streamlit-интерфейс «Отчет для собственника» (облачная версия).
#
# Дизайн:
#   • Крупный заголовок с градиентом.
#   • Две большие овальные вкладки.
#   • Декоративный SVG-фон.
#   • Крупные блоки загрузки файлов с форматом.
#   • Выбор периода — под заголовком.
#   • Кнопка «Выйти» — в правом верхнем углу.
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
    initial_sidebar_state="collapsed",   # сайдбар не нужен
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
            font-family: 'Playfair Display', serif;
            font-size: 2rem; font-weight: 900;
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
# ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: HTML-БЛОК ЗАГРУЗКИ
# ============================================================

def _upload_block_html(icon, title, desc, format_hint):
    """Крупная карточка-заголовок для блока загрузки файлов."""
    return f"""
    <div class="upload-block">
        <div class="upload-block-title">
            <span class="upload-block-icon">{icon}</span>
            <span>{title}</span>
        </div>
        <div class="upload-block-desc">{desc}</div>
        <div class="upload-block-format">
            <span class="format-label">Формат файла:</span>
            <code class="format-code">{format_hint}</code>
        </div>
    </div>
    """


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


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800;900&family=Manrope:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Manrope', 'Inter', -apple-system, sans-serif;
    }

    /* ===== ФОН С УЗОРАМИ ===== */
    .stApp {
        background-color: #0a0f1c;
        background-image:
            radial-gradient(circle at 15% 20%, rgba(59, 130, 246, 0.18) 0%, transparent 40%),
            radial-gradient(circle at 85% 10%, rgba(34, 211, 238, 0.16) 0%, transparent 40%),
            radial-gradient(circle at 75% 85%, rgba(167, 139, 250, 0.14) 0%, transparent 45%),
            radial-gradient(circle at 20% 90%, rgba(20, 184, 166, 0.12) 0%, transparent 45%),
            url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='400' height='400' viewBox='0 0 400 400'><defs><pattern id='p' x='0' y='0' width='80' height='80' patternUnits='userSpaceOnUse'><path d='M40 0 L80 40 L40 80 L0 40 Z' fill='none' stroke='%233b82f6' stroke-opacity='0.05' stroke-width='1'/></pattern><pattern id='c' x='0' y='0' width='120' height='120' patternUnits='userSpaceOnUse'><circle cx='60' cy='60' r='30' fill='none' stroke='%2322d3ee' stroke-opacity='0.05' stroke-width='1'/></pattern></defs><rect width='400' height='400' fill='url(%23p)'/><rect width='400' height='400' fill='url(%23c)'/></svg>"),
            linear-gradient(180deg, #0a0f1c 0%, #0f172a 100%);
        background-attachment: fixed;
        color: #e5e7eb;
    }

    h1, h2, h3, h4, h5, h6 { color: #f9fafb !important; }

    /* ===== ЗАГОЛОВОК ===== */
    .hero-title {
        font-family: 'Playfair Display', 'Manrope', serif;
        font-size: 4.2rem;
        font-weight: 900;
        line-height: 1.05;
        letter-spacing: -0.025em;
        background: linear-gradient(90deg, #60a5fa 0%, #22d3ee 45%, #a78bfa 100%);
        background-size: 200% 200%;
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text;
        margin: 0;
        text-align: center;
        filter: drop-shadow(0 6px 26px rgba(96, 165, 250, 0.35));
        animation: heroShift 8s ease infinite;
    }
    @keyframes heroShift {
        0%   { background-position: 0% 50%; }
        50%  { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    .hero-sub {
        text-align: center;
        color: #94a3b8;
        font-size: 1.1rem;
        margin-top: 0.6rem;
        font-weight: 500;
    }

    /* ===== КНОПКА ВЫХОДА В ПРАВОМ ВЕРХНЕМ УГЛУ ===== */
    .logout-corner {
        position: fixed; top: 1rem; right: 1.4rem; z-index: 999;
    }
    .logout-corner button {
        background: rgba(30, 41, 59, 0.85) !important;
        color: #cbd5e1 !important;
        border: 1px solid rgba(148, 163, 184, 0.3) !important;
        border-radius: 10px !important;
        padding: 0.4rem 0.9rem !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
    }
    .logout-corner button:hover {
        background: rgba(239, 68, 68, 0.25) !important;
        color: #fff !important;
        border-color: rgba(239, 68, 68, 0.6) !important;
    }

    /* ===== ВКЛАДКИ — БОЛЬШИЕ, ОВАЛЬНЫЕ ===== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
        justify-content: center;
        border-bottom: none;
        padding: 0.4rem 0 1.6rem 0;
        margin-bottom: 1.4rem;
    }
    .stTabs [data-baseweb="tab"] {
        background: linear-gradient(180deg,
            rgba(51, 65, 85, 0.9) 0%,
            rgba(30, 41, 59, 0.95) 100%);
        border: 1px solid rgba(148, 163, 184, 0.3);
        border-radius: 60px;
        padding: 1.6rem 3.6rem;
        color: #e2e8f0; font-weight: 800;
        transition: all 0.22s ease;
        min-height: 5.5rem;
        box-shadow:
            inset 0 1px 0 rgba(255, 255, 255, 0.1),
            inset 0 -3px 0 rgba(0, 0, 0, 0.25),
            0 8px 18px rgba(0, 0, 0, 0.4);
    }
    .stTabs [data-baseweb="tab"] p {
        font-size: 2rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        line-height: 1.2 !important;
        letter-spacing: -0.01em !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background: linear-gradient(180deg,
            rgba(59, 130, 246, 0.4) 0%,
            rgba(37, 99, 235, 0.55) 100%);
        border-color: rgba(147, 197, 253, 0.9);
        transform: translateY(-4px) scale(1.02);
        box-shadow: 0 16px 32px rgba(37, 99, 235, 0.5);
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(180deg,
            #3b82f6 0%, #2563eb 55%, #1d4ed8 100%) !important;
        border-color: #60a5fa !important;
        color: #ffffff !important;
        transform: translateY(-4px) scale(1.02);
        box-shadow:
            inset 0 1px 0 rgba(255, 255, 255, 0.35),
            inset 0 -4px 0 rgba(0, 0, 0, 0.3),
            0 18px 40px rgba(37, 99, 235, 0.55),
            0 8px 18px rgba(96, 165, 250, 0.4);
    }
    .stTabs [aria-selected="true"] p { color: #ffffff !important; }

    /* ===== ЗАГОЛОВКИ СЕКЦИЙ ===== */
    .section-title {
        font-size: 1.5rem; font-weight: 800; color: #f1f5f9;
        margin: 1.2rem 0 0.5rem 0;
        letter-spacing: -0.015em;
    }
    .section-hint {
        color: #94a3b8; font-size: 0.95rem;
        margin-bottom: 1.2rem;
    }

    /* ===== КРУПНЫЙ БЛОК ЗАГРУЗКИ ===== */
    .upload-block {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.9) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(148, 163, 184, 0.2);
        border-radius: 20px;
        padding: 1.6rem 1.9rem 1.4rem 1.9rem;
        margin-bottom: 0.6rem;
        margin-top: 1.4rem;
        position: relative;
        box-shadow: 0 12px 28px rgba(0, 0, 0, 0.3);
    }
    .upload-block::before {
        content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 8px;
        border-radius: 20px 0 0 20px;
        background: linear-gradient(180deg, #60a5fa, #22d3ee, #a78bfa);
    }
    .upload-block-title {
        display: flex; align-items: center; gap: 1rem;
        font-size: 1.7rem; font-weight: 800; color: #f8fafc;
        letter-spacing: -0.02em;
        margin: 0;
    }
    .upload-block-icon {
        font-size: 2.2rem; line-height: 1;
        filter: drop-shadow(0 4px 12px rgba(96, 165, 250, 0.4));
    }
    .upload-block-desc {
        color: #94a3b8; font-size: 1rem;
        margin: 0.6rem 0 0.8rem 0;
        line-height: 1.45;
    }
    .upload-block-format {
        display: flex; align-items: center; gap: 0.7rem;
        flex-wrap: wrap;
        padding: 0.7rem 1rem;
        background: rgba(96, 165, 250, 0.08);
        border: 1px solid rgba(96, 165, 250, 0.25);
        border-radius: 12px;
        margin-bottom: 0.9rem;
    }
    .format-label {
        color: #93c5fd; font-weight: 700; font-size: 0.88rem;
        text-transform: uppercase; letter-spacing: 0.08em;
    }
    .format-code {
        background: rgba(15, 23, 42, 0.9);
        color: #fbbf24;
        padding: 0.2rem 0.6rem;
        border-radius: 6px;
        font-family: 'Menlo', 'Consolas', monospace;
        font-size: 0.85rem;
    }

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

    /* ===== КНОПКИ ===== */
    .stButton > button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: #fff; border: none; border-radius: 12px;
        padding: 0.7rem 1.3rem; font-weight: 700;
        font-size: 1rem;
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.3);
        transition: all 0.18s ease;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        filter: brightness(1.1);
        box-shadow: 0 10px 26px rgba(37, 99, 235, 0.5);
    }

    .stDownloadButton > button {
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);
        color: #fff; border: none; border-radius: 12px;
        padding: 0.7rem 1.3rem; font-weight: 700;
        font-size: 1rem; width: 100%;
        box-shadow: 0 6px 18px rgba(13, 148, 136, 0.3);
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px); filter: brightness(1.1);
    }

    /* ===== SELECTBOX ===== */
    div[data-baseweb="select"] > div {
        background: rgba(15, 23, 42, 0.85) !important;
        border: 1px solid rgba(96, 165, 250, 0.35) !important;
        color: #f1f5f9 !important;
        border-radius: 12px !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
    }
    label { color: #cbd5e1 !important; font-weight: 600 !important; }

    /* ===== FILE UPLOADER ВНУТРИ БЛОКА ===== */
    [data-testid="stFileUploader"] {
        background: rgba(15, 23, 42, 0.5);
        border: 1px dashed rgba(96, 165, 250, 0.35);
        border-radius: 14px;
        padding: 0.4rem 0.6rem;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: rgba(96, 165, 250, 0.6);
    }

    section[data-testid="stSidebar"] { display: none; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header[data-testid="stHeader"] { background: transparent; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# КНОПКА «ВЫЙТИ» В ПРАВОМ ВЕРХНЕМ УГЛУ
# ============================================================

col_logout_left, col_logout_right = st.columns([5, 1])
with col_logout_right:
    if st.button(f"Выйти ({current_user})", key="logout_top"):
        st.session_state.pop(SESSION_KEY, None)
        st.rerun()


# ============================================================
# ЗАГОЛОВОК
# ============================================================

st.markdown(
    """
    <div style="padding: 0.5rem 0 0.4rem 0;">
        <div class="hero-title">📊 Отчет для собственника</div>
        <div class="hero-sub">
            Финансовая отчётность Estate · Unelma · Nomiqa и диаграммы для презентации
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# ВЫБОР ПЕРИОДА — ПОД ЗАГОЛОВКОМ
# ============================================================

_MONTH_NAMES = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}

_YEAR_OPTIONS = [2026, 2027]

st.markdown("<br>", unsafe_allow_html=True)
col_yr, col_mo, col_sp = st.columns([1.2, 1.6, 3])

with col_yr:
    _year = st.selectbox(
        "Год", options=_YEAR_OPTIONS, index=0, key="report_year",
    )
with col_mo:
    _month = st.selectbox(
        "Месяц",
        options=list(_MONTH_NAMES.keys()),
        format_func=lambda m: _MONTH_NAMES[m],
        index=config.REPORT_MONTH - 1,
        key="report_month",
    )
with col_sp:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f'<div style="color:#93c5fd; font-weight:700; font-size:1.05rem;">'
        f'📅 Период отчёта: {_MONTH_NAMES[_month]} {_year}</div>',
        unsafe_allow_html=True,
    )

config.set_period(_year, _month)


# ============================================================
# ДВЕ ВКЛАДКИ
# ============================================================

st.markdown("<br>", unsafe_allow_html=True)
tab_reports, tab_presentation = st.tabs([
    "📊  Отчёты для собственника",
    "📈  Диаграммы для презентации",
])


# ============================================================
# ВКЛАДКА 1: ОТЧЁТЫ
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

    # ---- ОПиУ ----
    st.markdown(
        _upload_block_html(
            "📗", "ОПиУ (отчёт о прибылях и убытках)",
            "Файлы ОПиУ за <b>прошлый</b> и <b>текущий</b> месяц по каждому "
            "направлению: <b>Латвия</b>, <b>Европа</b>, <b>Estate AZE</b>, "
            "<b>Nomiqa</b>, <b>UK Estate</b>, <b>Unelma</b>, "
            "<b>Estate (единый)</b>.",
            "ОПиУ ММ.ГГГГ-ММ.ГГГГ <Направление>.xlsx",
        ),
        unsafe_allow_html=True,
    )
    opiu_files = st.file_uploader(
        "ОПиУ (.xlsx)", type=["xlsx"],
        accept_multiple_files=True, key="upload_opiu",
        label_visibility="collapsed",
    )

    # ---- БДиР ----
    st.markdown(
        _upload_block_html(
            "📘", "БДиР (бюджет доходов и расходов)",
            "Годовые файлы БДиР по направлениям: <b>Латвия</b>, "
            "<b>Европа (Estate EU)</b>, <b>Estate AZE</b>, <b>Nomiqa</b>, "
            "<b>UK Estate</b>, <b>Unelma</b>.",
            "БДиР 08.2026-12.2026 <Направление>.xlsx",
        ),
        unsafe_allow_html=True,
    )
    bdr_files = st.file_uploader(
        "БДиР (.xlsx)", type=["xlsx"],
        accept_multiple_files=True, key="upload_bdr",
        label_visibility="collapsed",
    )

    # ---- Прогнозы ----
    st.markdown(
        _upload_block_html(
            "📙", "Прогнозы месячные",
            "Файлы «Прогнозы месячные» по тем же направлениям, что и БДиР.",
            "Прогнозы месячные <Направление>.xlsx",
        ),
        unsafe_allow_html=True,
    )
    forecast_files = st.file_uploader(
        "Прогнозы (.xlsx)", type=["xlsx"],
        accept_multiple_files=True, key="upload_forecast",
        label_visibility="collapsed",
    )

    # ---- НДС ----
    st.markdown(
        _upload_block_html(
            "📕", "НДС",
            "Файлы НДС за <b>прошлый</b> и <b>текущий</b> месяц. "
            "В имени файла должно быть слово «НДС».",
            "НДС ММ.ГГГГ <Статья>.xlsx",
        ),
        unsafe_allow_html=True,
    )
    vat_files = st.file_uploader(
        "НДС (.xlsx)", type=["xlsx"],
        accept_multiple_files=True, key="upload_vat",
        label_visibility="collapsed",
    )

    # Сохраняем
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

    # Метрики
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
            """, unsafe_allow_html=True,
        )
    with col2:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-label">📄 Всего отчётов</div>
                <div class="metric-value">3</div>
                <div class="metric-sub">Estate · Unelma · Nomiqa</div>
            </div>
            """, unsafe_allow_html=True,
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
            """, unsafe_allow_html=True,
        )
    with col4:
        if total_missing == 0:
            cls, val, sub = ("metric-card status-ok",
                             "всё на месте", "можно собирать")
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
            """, unsafe_allow_html=True,
        )

    # Проверка
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
                "template": "Шаблон", "opiu_prev": "ОПиУ за прошлый месяц",
                "opiu_cur": "ОПиУ за текущий месяц", "bdr": "БДиР",
                "forecast": "Прогноз месячный", "vat": "НДС",
                "error": "Ошибка проверки",
            }
            with st.expander(
                f"⚠️ {report_key}: не хватает {len(missing)} файлов",
                expanded=False,
            ):
                for k, p in missing:
                    st.text(f"• {labels.get(k, k)}: {p.name}")

    # Сборка
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">🚀 Шаг 3 · Сборка отчётов</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-hint">Соберите все три отчёта или выборочно.</div>',
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

    c_all, c_est, c_un, c_nom = st.columns([1.4, 1, 1, 1])
    with c_all:
        if st.button("🔨 Собрать все отчёты", type="primary",
                     use_container_width=True):
            _run_build("ALL")
    with c_est:
        if st.button("Estate", use_container_width=True, type="secondary"):
            _run_build("Estate")
    with c_un:
        if st.button("Unelma", use_container_width=True, type="secondary"):
            _run_build("Unelma")
    with c_nom:
        if st.button("Nomiqa", use_container_width=True, type="secondary"):
            _run_build("Nomiqa")

    # Скачивание
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
                    data=data, file_name=path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{key}", use_container_width=True,
                )
                st.caption(f"`{path.name}`")
            else:
                st.info(f"Файл ещё не собран: `{path.name}`")


# ============================================================
# ВКЛАДКА 2: ДИАГРАММЫ
# ============================================================

with tab_presentation:

    st.markdown(
        '<div class="section-title">📥 Шаг 1 · Загрузка файлов</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-hint">'
        'Загрузите ОПиУ за все месяцы года, единый БДиР, ОДДС, '
        'файлы возмещения КУ и два справочника.'
        '</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        _upload_block_html(
            "📗", "ОПиУ за все месяцы года",
            "Сводный файл ОПиУ за январь–отчётный месяц. "
            "Нужен для построения всех диаграмм презентации.",
            "ОПиУ 01.2026-09.2026.xlsx",
        ),
        unsafe_allow_html=True,
    )
    opiu_full_year_files = st.file_uploader(
        "ОПиУ за все месяцы (.xlsx)", type=["xlsx"],
        accept_multiple_files=True, key="upload_opiu_full_year",
        label_visibility="collapsed",
    )

    st.markdown(
        _upload_block_html(
            "📘", "БДиР единый сводный",
            "Один файл БДиР за весь год. Нужен для диаграмм «План/Факт».",
            "БДиР 01.2026-12.2026.xlsx",
        ),
        unsafe_allow_html=True,
    )
    bdr_full_year_files = st.file_uploader(
        "БДиР за все месяцы (.xlsx)", type=["xlsx"],
        accept_multiple_files=True, key="upload_bdr_full_year",
        label_visibility="collapsed",
    )

    st.markdown(
        _upload_block_html(
            "📒", "ОДДС (операционное сальдо и ДДС)",
            "ОДДС за весь год. Нужен для слайдов «Операционное сальдо» "
            "и «Анализ ДДС».",
            "ОДДС 01.01.2026-30.09.2026.xlsx",
        ),
        unsafe_allow_html=True,
    )
    odds_files = st.file_uploader(
        "ОДДС (.xlsx)", type=["xlsx"],
        accept_multiple_files=True, key="upload_odds",
        label_visibility="collapsed",
    )

    st.markdown(
        _upload_block_html(
            "📗", "ОДДС для слайдов 25 / 27 / 29 (Возмещение КУ)",
            "Два файла ОДДС за весь год для таблиц возмещения КУ. "
            "В именах должны быть слова «с НДС» и «без НДС».",
            "ОДДС 01.01.2026-30.09.2026 с НДС.xlsx<br>"
            "ОДДС 01.01.2026-30.09.2026 без НДС.xlsx",
        ),
        unsafe_allow_html=True,
    )
    ku_files = st.file_uploader(
        "ОДДС для КУ (.xlsx)", type=["xlsx"],
        accept_multiple_files=True, key="upload_ku",
        label_visibility="collapsed",
    )

    st.markdown(
        _upload_block_html(
            "📓", "Справочник: Вложенные средства на объекты",
            "Один файл. Используется для расчёта ROI (слайд 70).",
            "Вложенные средства на объекты.xlsx",
        ),
        unsafe_allow_html=True,
    )
    investment_files = st.file_uploader(
        "Вложения (.xlsx)", type=["xlsx"],
        accept_multiple_files=False, key="upload_investments",
        label_visibility="collapsed",
    )

    st.markdown(
        _upload_block_html(
            "📓", "Справочник: Долги и численность сотрудников",
            "Один файл. Используется для слайдов 14, 19, 24, 64.",
            "таблица Долги и Численность сотрудников.xlsx",
        ),
        unsafe_allow_html=True,
    )
    debts_files = st.file_uploader(
        "Долги (.xlsx)", type=["xlsx"],
        accept_multiple_files=False, key="upload_debts",
        label_visibility="collapsed",
    )

    # Сохранение
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
        n_opiu_full + n_bdr_full + n_odds
        + n_ku + n_invest + n_debts
    )
    if total_pres_uploaded > 0:
        st.success(
            f"✅ Сохранено: ОПиУ (за год) — {n_opiu_full}, "
            f"БДиР (сводный) — {n_bdr_full}, "
            f"ОДДС — {n_odds}, "
            f"Возмещение КУ — {n_ku}, Вложений — {n_invest}, "
            f"Долгов — {n_debts}."
        )

    # Сборка
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

    # Скачивание
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
            data=pres_data, file_name=pres_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_presentation", use_container_width=True,
        )
        st.caption(f"`{pres_path.name}`")
    else:
        st.info(f"Файл ещё не собран: `{pres_path.name if pres_path else '—'}`")


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