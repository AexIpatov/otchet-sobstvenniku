# app.py
# ============================================================
# Streamlit-интерфейс «Отчет для собственника» (локальная версия).
# + Авторизация по логину и паролю.
# + Два таба: «Отчёты для собственника» и «Диаграммы для презентации».
# ============================================================

import os
import hashlib
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
# КАСТОМНЫЙ CSS — ОФОРМЛЕНИЕ
# ============================================================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    /* ===== ФОН С АНИМАЦИЕЙ ===== */
    .stApp {
        background:
            radial-gradient(circle at 0% 0%, #1e3a8a 0%, transparent 45%),
            radial-gradient(circle at 100% 0%, #0f766e 0%, transparent 45%),
            radial-gradient(circle at 50% 100%, #4c1d95 0%, transparent 55%),
            linear-gradient(180deg, #0f172a 0%, #111827 100%);
        background-size: 100% 100%, 100% 100%, 100% 100%, 100% 100%;
        background-attachment: fixed;
        color: #e5e7eb;
        animation: bgMove 30s ease-in-out infinite;
    }
    @keyframes bgMove {
        0%, 100% { background-position: 0% 0%, 100% 0%, 50% 100%, 0% 0%; }
        50%      { background-position: 20% 10%, 80% 20%, 60% 80%, 0% 0%; }
    }

    h1, h2, h3, h4, h5, h6 { color: #f9fafb !important; letter-spacing: -0.02em; }

    /* ===== ЗАГОЛОВОК ===== */
    .main-title {
        font-size: 3.2rem; font-weight: 900;
        background: linear-gradient(90deg, #60a5fa 0%, #22d3ee 50%, #a78bfa 100%);
        background-size: 200% 200%;
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text; margin: 0; line-height: 1.1;
        animation: gradientShift 6s ease infinite;
        letter-spacing: -0.03em;
        filter: drop-shadow(0 4px 20px rgba(96, 165, 250, 0.25));
    }
    @keyframes gradientShift {
        0%   { background-position: 0% 50%; }
        50%  { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    .main-subtitle {
        font-size: 1.15rem; color: #94a3b8;
        margin-top: 0.55rem; font-weight: 500;
    }

    .greeting {
        font-size: 1.5rem; color: #e2e8f0;
        margin-bottom: 0.5rem; font-weight: 500;
    }
    .greeting strong {
        background: linear-gradient(90deg, #f472b6 0%, #fbbf24 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text; font-weight: 800;
        font-size: 1.6rem;
    }

    /* ===== ЗАГОЛОВКИ СЕКЦИЙ ===== */
    .section-title {
        font-size: 1.7rem; font-weight: 800; color: #f1f5f9;
        margin: 1.2rem 0 0.5rem 0;
        display: flex; align-items: center; gap: 0.7rem;
        letter-spacing: -0.02em;
    }
    .section-hint {
        color: #94a3b8; font-size: 1rem;
        margin-bottom: 1.3rem; font-weight: 400;
    }

    /* ===== КАРТОЧКИ МЕТРИК ===== */
    .metric-card {
        background:
            linear-gradient(135deg, rgba(30, 41, 59, 0.95) 0%, rgba(15, 23, 42, 0.95) 100%);
        border: 1px solid rgba(148, 163, 184, 0.15);
        border-radius: 18px;
        padding: 1.5rem 1.6rem;
        position: relative; overflow: hidden;
        transition: transform 0.25s ease, border-color 0.25s ease, box-shadow 0.25s ease;
        height: 100%;
        animation: cardFadeIn 0.55s ease both;
    }
    @keyframes cardFadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to   { opacity: 1; transform: translateY(0); }
    }
    .metric-card::before {
        content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 5px;
        background: linear-gradient(180deg, #60a5fa, #22d3ee, #a78bfa);
        border-radius: 18px 0 0 18px;
    }
    .metric-card::after {
        content: ""; position: absolute; right: -30px; top: -30px;
        width: 110px; height: 110px; border-radius: 50%;
        background: radial-gradient(circle, rgba(96, 165, 250, 0.18), transparent 70%);
        pointer-events: none;
    }
    .metric-card:hover {
        transform: translateY(-5px);
        border-color: rgba(96, 165, 250, 0.55);
        box-shadow: 0 20px 42px rgba(37, 99, 235, 0.28);
    }
    .metric-label {
        color: #94a3b8; font-size: 0.86rem; font-weight: 600;
        text-transform: uppercase; letter-spacing: 0.1em;
        margin-bottom: 0.55rem;
    }
    .metric-value {
        color: #f8fafc; font-size: 2rem; font-weight: 800;
        line-height: 1.15; letter-spacing: -0.02em;
    }
    .metric-sub {
        color: #64748b; font-size: 0.85rem;
        margin-top: 0.4rem; font-weight: 500;
    }

    /* ===== СТАТУСЫ ===== */
    .status-ok { border-color: rgba(34, 197, 94, 0.5) !important; }
    .status-ok::before {
        background: linear-gradient(180deg, #22c55e, #4ade80) !important;
    }
    .status-ok::after {
        background: radial-gradient(circle, rgba(34, 197, 94, 0.22), transparent 70%) !important;
    }
    .status-warn { border-color: rgba(245, 158, 11, 0.5) !important; }
    .status-warn::before {
        background: linear-gradient(180deg, #f59e0b, #fbbf24) !important;
    }
    .status-pulse { animation: pulse 2.4s ease-in-out infinite; }
    @keyframes pulse {
        0%, 100% { box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.45); }
        50%      { box-shadow: 0 0 0 12px rgba(34, 197, 94, 0); }
    }

    /* ===== САЙДБАР ===== */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0b1220 0%, #0f172a 100%);
        border-right: 1px solid rgba(148, 163, 184, 0.1);
    }
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #e2e8f0 !important; font-size: 1.05rem;
        letter-spacing: 0.08em; text-transform: uppercase;
        font-weight: 800;
    }

    /* ===== КНОПКИ ===== */
    .stButton > button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: #fff; border: none; border-radius: 12px;
        padding: 0.8rem 1.3rem;
        font-weight: 700; font-size: 1.02rem;
        transition: transform 0.18s ease, box-shadow 0.18s ease, filter 0.18s ease;
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.3);
        letter-spacing: 0.01em;
    }
    .stButton > button:hover {
        transform: translateY(-2px) scale(1.015);
        filter: brightness(1.12);
        box-shadow: 0 10px 28px rgba(37, 99, 235, 0.5);
    }
    .stButton > button:active { transform: translateY(0) scale(1); }

    .stButton > button[kind="secondary"] {
        background: rgba(148, 163, 184, 0.12);
        color: #e2e8f0;
        box-shadow: none;
        border: 1px solid rgba(148, 163, 184, 0.3);
    }
    .stButton > button[kind="secondary"]:hover {
        background: rgba(96, 165, 250, 0.18);
        border-color: rgba(96, 165, 250, 0.6);
        box-shadow: 0 6px 18px rgba(96, 165, 250, 0.25);
    }

    .stDownloadButton > button {
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);
        color: #fff; border: none; border-radius: 12px;
        padding: 0.8rem 1.3rem;
        font-weight: 700; font-size: 1.02rem;
        width: 100%;
        transition: transform 0.18s ease, filter 0.18s ease, box-shadow 0.18s ease;
        box-shadow: 0 6px 18px rgba(13, 148, 136, 0.3);
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px) scale(1.015);
        filter: brightness(1.12);
        box-shadow: 0 10px 28px rgba(13, 148, 136, 0.5);
    }

    /* ===== EXPANDER ===== */
    details[data-testid="stExpander"] {
        background: rgba(15, 23, 42, 0.65);
        border: 1px solid rgba(148, 163, 184, 0.18);
        border-radius: 14px;
        margin-bottom: 0.8rem;
        overflow: hidden;
        transition: border-color 0.2s ease, box-shadow 0.2s ease;
    }
    details[data-testid="stExpander"]:hover {
        border-color: rgba(96, 165, 250, 0.45);
        box-shadow: 0 8px 24px rgba(37, 99, 235, 0.15);
    }
    details[data-testid="stExpander"] summary {
        color: #e2e8f0 !important;
        font-weight: 700; font-size: 1.02rem;
        padding: 0.9rem 1.1rem;
    }

    /* ===== SELECT ===== */
    div[data-baseweb="select"] > div {
        background: rgba(15, 23, 42, 0.8) !important;
        border-color: rgba(148, 163, 184, 0.3) !important;
        color: #f1f5f9 !important;
        border-radius: 12px !important;
        font-size: 1rem !important;
    }
    div[data-baseweb="select"] svg { fill: #94a3b8 !important; }
    label {
        color: #cbd5e1 !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
    }

    hr { border-color: rgba(148, 163, 184, 0.12); }
    .stSpinner > div { border-top-color: #60a5fa !important; }
    div[data-testid="stAlert"] {
        border-radius: 12px;
        border: 1px solid rgba(148, 163, 184, 0.2);
        font-size: 1rem;
    }

    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header[data-testid="stHeader"] { background: transparent; }

    /* ===== КРУЖКИ С НОМЕРАМИ ШАГОВ ===== */
    .step-badge {
        display: inline-flex; align-items: center; justify-content: center;
        width: 2.6rem; height: 2.6rem; border-radius: 50%;
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: #fff; font-weight: 900; font-size: 1.15rem;
        margin-right: 0.7rem;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.4);
        flex-shrink: 0;
    }

    /* ===== ВКЛАДКИ STREAMLIT — КРУПНЫЕ И ОБЪЁМНЫЕ ===== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 1.4rem;
        justify-content: center;
        border-bottom: 2px solid rgba(148, 163, 184, 0.15);
        padding: 0.6rem 0 1.4rem 0;
        margin-bottom: 1.6rem;
    }
    .stTabs [data-baseweb="tab"] {
        background: linear-gradient(180deg,
            rgba(51, 65, 85, 0.85) 0%,
            rgba(30, 41, 59, 0.95) 100%);
        border: 1px solid rgba(148, 163, 184, 0.3);
        border-radius: 18px;
        padding: 1.3rem 2.6rem;
        color: #e2e8f0;
        font-weight: 800;
        letter-spacing: 0.01em;
        transition: all 0.22s cubic-bezier(0.34, 1.56, 0.64, 1);
        min-height: 4rem;
        /* Объёмность: светлая полоска сверху, тёмная тень снизу + мягкое свечение */
        box-shadow:
            inset 0 1px 0 rgba(255, 255, 255, 0.1),
            inset 0 -2px 0 rgba(0, 0, 0, 0.25),
            0 6px 14px rgba(0, 0, 0, 0.35),
            0 2px 6px rgba(37, 99, 235, 0.12);
        position: relative;
        overflow: hidden;
    }
    .stTabs [data-baseweb="tab"] p {
        font-size: 1.45rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        letter-spacing: -0.01em !important;
        line-height: 1.2 !important;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.35);
    }
    .stTabs [data-baseweb="tab"]:hover {
        background: linear-gradient(180deg,
            rgba(59, 130, 246, 0.35) 0%,
            rgba(37, 99, 235, 0.5) 100%);
        border-color: rgba(147, 197, 253, 0.8);
        color: #ffffff;
        transform: translateY(-4px) scale(1.02);
        box-shadow:
            inset 0 1px 0 rgba(255, 255, 255, 0.25),
            inset 0 -2px 0 rgba(0, 0, 0, 0.2),
            0 14px 28px rgba(37, 99, 235, 0.45),
            0 4px 10px rgba(37, 99, 235, 0.3);
    }
    /* Активная вкладка — сильный объём и свечение */
    .stTabs [aria-selected="true"] {
        background: linear-gradient(180deg,
            #3b82f6 0%,
            #2563eb 50%,
            #1d4ed8 100%) !important;
        border-color: #60a5fa !important;
        color: #ffffff !important;
        transform: translateY(-4px) scale(1.02);
        box-shadow:
            inset 0 1px 0 rgba(255, 255, 255, 0.35),
            inset 0 -3px 0 rgba(0, 0, 0, 0.3),
            0 16px 36px rgba(37, 99, 235, 0.55),
            0 6px 16px rgba(96, 165, 250, 0.4),
            0 0 0 4px rgba(96, 165, 250, 0.15);
    }
    .stTabs [aria-selected="true"] p {
        color: #ffffff !important;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.4) !important;
    }
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
st.sidebar.caption(
    f"📁 Рабочая папка: `{config.BASE_DIR.name}`  \n"
    f"📤 Выходная папка: `data/output/`"
)
st.sidebar.markdown("---")
st.sidebar.caption(f"👤 Пользователь: **{current_user}**")
if st.sidebar.button("Выйти", use_container_width=True):
    st.session_state.pop(SESSION_KEY, None)
    st.rerun()


# ============================================================
# ВСПОМОГАТЕЛЬНОЕ: ДИАГНОСТИКА ПО ФАЙЛАМ
# ============================================================

def _file_status_table():
    result = {}
    for report_key in ("Estate", "Unelma", "Nomiqa"):
        missing = config.check_inputs(report_key)
        flat = []
        for k in ("template", "opiu_prev", "opiu_cur", "bdr", "forecast"):
            for p in missing[k]:
                flat.append((k, Path(p)))
        result[report_key] = flat
    return result


# ============================================================
# ЗАГОЛОВОК
# ============================================================

st.markdown(
    f"""
    <div style="padding: 1.2rem 0 0.6rem 0; text-align: center;">
        <div class="greeting">👋 Привет, <strong>{current_user}</strong>!</div>
        <div class="main-title">📊 Отчет для собственника</div>
        <div class="main-subtitle">
            Период: <b style="color:#93c5fd;">{_MONTH_NAMES[_month]} {_year}</b>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)


# ============================================================
# ДВА ТАБА: ОТЧЁТЫ И ДИАГРАММЫ
# ============================================================

tab_reports, tab_presentation = st.tabs([
    "📊 Отчёты для собственника",
    "📈 Диаграммы для презентации",
])


# ============================================================
# ТАБ 1: ОТЧЁТЫ ДЛЯ СОБСТВЕННИКА
# ============================================================

with tab_reports:

    status = _file_status_table()

    col1, col2, col3, col4 = st.columns(4)
    ready = sum(1 for k in status if not status[k])
    total_missing = sum(len(v) for v in status.values())

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
        cls = "metric-card status-ok status-pulse" if ready == 3 else "metric-card"
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
            cls = "metric-card status-ok"
            val = "всё на месте"
            sub = "можно собирать"
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

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Шаг 1: диагностика ---
    with st.expander("🔍 Шаг 1 · Проверка входных файлов", expanded=False):
        st.caption(
            "Файлы читаются из `data/input/`. "
            "Если чего-то не хватает — положите нужный файл в соответствующую папку "
            "и обновите страницу."
        )
        for report_key in ("Estate", "Unelma", "Nomiqa"):
            missing = status[report_key]
            if not missing:
                st.success(f"**{report_key}**: все входные файлы на месте")
            else:
                st.warning(f"**{report_key}**: не хватает {len(missing)} файлов")
                labels = {
                    "template":  "Шаблон",
                    "opiu_prev": "ОПиУ за прошлый месяц",
                    "opiu_cur":  "ОПиУ за текущий месяц",
                    "bdr":       "БДиР",
                    "forecast":  "Прогноз месячный",
                }
                for k, p in missing:
                    st.text(f"   • {labels.get(k, k)}: {p}")

    # --- Шаг 2: сборка ---
    st.markdown(
        '<div class="section-title"><span class="step-badge">2</span>Сборка отчётов</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-hint">Соберите все три отчёта одной кнопкой или выборочно.</div>',
        unsafe_allow_html=True,
    )

    col_build_all, col_build_estate, col_build_unelma, col_build_nomiqa = st.columns([1.4, 1, 1, 1])

    with col_build_all:
        if st.button("🔨 Собрать все отчёты", type="primary", use_container_width=True, key="btn_all"):
            with st.spinner("⏳ Собираю все три отчёта..."):
                try:
                    build_all_reports()
                    st.success("✅ Готово! Файлы сохранены в `data/output/`.")
                    st.balloons()
                except Exception as e:
                    st.error(f"Ошибка при сборке: {e}")

    with col_build_estate:
        if st.button("Estate", use_container_width=True, type="secondary", key="btn_est"):
            with st.spinner("⏳ Собираю Estate..."):
                try:
                    build_report("Estate")
                    st.success("✅ Estate собран.")
                except Exception as e:
                    st.error(f"Ошибка: {e}")

    with col_build_unelma:
        if st.button("Unelma", use_container_width=True, type="secondary", key="btn_un"):
            with st.spinner("⏳ Собираю Unelma..."):
                try:
                    build_report("Unelma")
                    st.success("✅ Unelma собран.")
                except Exception as e:
                    st.error(f"Ошибка: {e}")

    with col_build_nomiqa:
        if st.button("Nomiqa", use_container_width=True, type="secondary", key="btn_nom"):
            with st.spinner("⏳ Собираю Nomiqa..."):
                try:
                    build_report("Nomiqa")
                    st.success("✅ Nomiqa собран.")
                except Exception as e:
                    st.error(f"Ошибка: {e}")

    # --- Шаг 3: скачивание ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title"><span class="step-badge">3</span>Скачать готовые отчёты</div>',
        unsafe_allow_html=True,
    )

    col_dl_1, col_dl_2, col_dl_3 = st.columns(3)
    for col, report_key in zip((col_dl_1, col_dl_2, col_dl_3), ("Estate", "Unelma", "Nomiqa")):
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
# ТАБ 2: ДИАГРАММЫ ДЛЯ ПРЕЗЕНТАЦИИ
# ============================================================

with tab_presentation:

    # --- Карточки ---
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
                <div class="metric-label">📄 Листов в файле</div>
                <div class="metric-value">80+</div>
                <div class="metric-sub">с диаграммами</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col3:
        opiu_ok = config.PRESENTATION_OPIU_FILE and config.PRESENTATION_OPIU_FILE.exists()
        cls = "metric-card status-ok" if opiu_ok else "metric-card status-warn"
        val = "на месте" if opiu_ok else "нет файла"
        sub = "ОПиУ за все месяцы" if opiu_ok else "положите файл в data/input/ОПиУ/"
        st.markdown(
            f"""
            <div class="{cls}">
                <div class="metric-label">📥 ОПиУ</div>
                <div class="metric-value">{val}</div>
                <div class="metric-sub">{sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col4:
        bdr_ok = config.PRESENTATION_BDR_FILE and config.PRESENTATION_BDR_FILE.exists()
        cls = "metric-card status-ok" if bdr_ok else "metric-card status-warn"
        val = "на месте" if bdr_ok else "нет файла"
        sub = "БДиР за год" if bdr_ok else "положите файл в data/input/Прогнозы/"
        st.markdown(
            f"""
            <div class="{cls}">
                <div class="metric-label">📥 БДиР</div>
                <div class="metric-value">{val}</div>
                <div class="metric-sub">{sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Что собирается ---
    with st.expander("📋 Что собирается в файле «Диаграммы для презентации»", expanded=False):
        st.markdown(
            """
            **Файл содержит 80+ листов с диаграммами:**
            - **ОПиУ, EBITDA, Доля ФОТ** — по каждому юниту и объекту
            - **ДДС, Операционное сальдо** — по Латвии и Востоку
            - **Точка безубыточности (ТБУ)** — Латвия, Восток, Estate
            - **Рентабельность по ЧП** — Латвия (4 линии)
            - **План/Факт по ЧП** — все юниты и объекты
            - **Нарастающий итог по ЧП** — все юниты и объекты
            - **Долги по объектам** — Антонияс, Чака, Коммерческие
            - **Возмещение коммунальных услуг** — с НДС / без НДС
            - **Анализ ОПиУ** — численность, выручка/чел., ЗП/чел.
            - **Выручка по юнитам, EBITDA margin, ROI, Доля ФОТ** — аналитические слайды
            """
        )

    # --- Сборка ---
    st.markdown(
        '<div class="section-title">'
        '<span class="step-badge" style="background: linear-gradient(135deg, #a78bfa 0%, #7c3aed 100%); '
        'box-shadow: 0 6px 16px rgba(167, 139, 250, 0.4);">1</span>'
        'Сборка диаграмм</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-hint">'
        'Используются: ОПиУ за все месяцы года, БДиР за год, ОДДС и справочники.'
        '</div>',
        unsafe_allow_html=True,
    )

    if st.button("📊 Собрать диаграммы для презентации",
                 type="primary", use_container_width=True,
                 key="build_presentation"):
        if not config.PRESENTATION_OPIU_FILE or not config.PRESENTATION_OPIU_FILE.exists():
            st.error(
                f"❌ Не найден файл ОПиУ за все месяцы: "
                f"`{config.PRESENTATION_OPIU_FILE.name if config.PRESENTATION_OPIU_FILE else '—'}`. "
                f"Положите его в папку `data/input/ОПиУ/`."
            )
        else:
            with st.spinner("⏳ Собираю диаграммы... Это может занять 1–2 минуты."):
                try:
                    build_presentation_data(
                        opiu_path=config.PRESENTATION_OPIU_FILE,
                        bdr_path=config.PRESENTATION_BDR_FILE,
                        forecast_path=config.PRESENTATION_FORECAST_FILE,
                        year=config.REPORT_YEAR,
                        month=config.REPORT_MONTH,
                        output_path=config.PRESENTATION_OUTPUT_FILE,
                    )
                    st.success("✅ Готово! Файл с диаграммами собран.")
                    st.balloons()
                except Exception as e:
                    st.error(f"Ошибка при сборке диаграмм: {e}")
                    import traceback
                    with st.expander("🔧 Подробности ошибки (для разработчика)"):
                        st.code(traceback.format_exc())

    # --- Скачивание ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">'
        '<span class="step-badge" style="background: linear-gradient(135deg, #a78bfa 0%, #7c3aed 100%); '
        'box-shadow: 0 6px 16px rgba(167, 139, 250, 0.4);">2</span>'
        'Скачать готовый файл</div>',
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
    <div style="text-align:center; color:#475569; font-size:0.8rem; padding: 0.5rem;">
        Рабочая папка: <code style="color:#94a3b8;">{config.BASE_DIR}</code> ·
        Выходная папка: <code style="color:#94a3b8;">{config.OUTPUT_DIR}</code> ·
        Streamlit {st.__version__}
    </div>
    """,
    unsafe_allow_html=True,
)