"""
app.py — Streamlit-приложение для сверочного отчёта цветочного магазина.
Запуск: streamlit run app.py
"""

import streamlit as st
from report_engine import run

# ─── Настройка страницы ───────────────────────────────────────────────────────

st.set_page_config(
    page_title="Сверочный отчёт | Цветочный магазин",
    page_icon="🌸",
    layout="centered",
)

# ─── CSS ─────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .main { background-color: #f9f9f9; }
    .block-container { max-width: 780px; padding-top: 2rem; }
    .stButton > button {
        background-color: #1F4E79;
        color: white;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        font-size: 16px;
        font-weight: bold;
        border: none;
        width: 100%;
        margin-top: 1rem;
    }
    .stButton > button:hover {
        background-color: #2E6DA4;
    }
    .upload-box {
        background: white;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        margin-bottom: 1rem;
    }
    h1 { color: #1F4E79; }
    h3 { color: #2E4057; margin-top: 0; }
    .success-box {
        background: #EAF3DE;
        border-left: 4px solid #276221;
        padding: 1rem 1.2rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    .error-box {
        background: #FFC7CE;
        border-left: 4px solid #9C0006;
        padding: 1rem 1.2rem;
        border-radius: 8px;
        margin: 1rem 0;
        color: #9C0006;
    }
    .info-box {
        background: #EFF6FF;
        border-left: 4px solid #1F4E79;
        padding: 0.8rem 1.2rem;
        border-radius: 8px;
        margin: 1rem 0;
        font-size: 14px;
        color: #1F4E79;
    }
</style>
""", unsafe_allow_html=True)

# ─── Заголовок ────────────────────────────────────────────────────────────────

st.markdown("# 🌸 Сверочный отчёт")
st.markdown("**Цветочный магазин** — сверка Битрикс с таблицей продаж")
st.divider()

# ─── Инструкция ───────────────────────────────────────────────────────────────

st.markdown("""
<div class="info-box">
📌 <b>Как пользоваться:</b><br>
1. Загрузите файл из Битрикс (Deal_*.xls)<br>
2. Загрузите таблицу продаж (*.xlsx)<br>
3. Нажмите <b>«Сформировать отчёт»</b><br>
4. Скачайте готовый Excel одной кнопкой
</div>
""", unsafe_allow_html=True)

# ─── Загрузка файлов ─────────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📂 Файл Битрикс")
    bitrix_file = st.file_uploader(
        "Загрузите Deal_*.xls",
        type=["xls", "html", "htm"],
        key="bitrix",
        label_visibility="collapsed"
    )
    if bitrix_file:
        st.markdown(f'<div class="success-box">✅ <b>{bitrix_file.name}</b><br>{bitrix_file.size // 1024} КБ</div>',
                    unsafe_allow_html=True)

with col2:
    st.markdown("### 📊 Таблица продаж")
    sales_file = st.file_uploader(
        "Загрузите *.xlsx",
        type=["xlsx"],
        key="sales",
        label_visibility="collapsed"
    )
    if sales_file:
        st.markdown(f'<div class="success-box">✅ <b>{sales_file.name}</b><br>{sales_file.size // 1024} КБ</div>',
                    unsafe_allow_html=True)

st.divider()

# ─── Кнопка и генерация ───────────────────────────────────────────────────────

both_uploaded = bitrix_file is not None and sales_file is not None

if not both_uploaded:
    st.markdown(
        '<div style="text-align:center; color:#888; font-size:15px;">⬆️ Загрузите оба файла для формирования отчёта</div>',
        unsafe_allow_html=True
    )

if both_uploaded:
    if st.button("🚀 Сформировать отчёт", use_container_width=True):
        with st.spinner("⏳ Обрабатываем данные, формируем Excel..."):
            try:
                bitrix_bytes = bitrix_file.read()
                sales_bytes  = sales_file.read()

                excel_bytes = run(bitrix_bytes, sales_bytes)

                st.session_state["excel_bytes"] = excel_bytes
                st.session_state["ready"] = True

            except Exception as e:
                st.markdown(
                    f'<div class="error-box">❌ <b>Ошибка при обработке:</b><br>{str(e)}</div>',
                    unsafe_allow_html=True
                )
                st.session_state["ready"] = False

# ─── Кнопка скачивания ───────────────────────────────────────────────────────

if st.session_state.get("ready"):
    st.markdown(
        '<div class="success-box">✅ <b>Отчёт готов!</b> Нажмите кнопку ниже чтобы скачать.</div>',
        unsafe_allow_html=True
    )
    st.download_button(
        label="⬇️ Скачать отчёт Excel",
        data=st.session_state["excel_bytes"],
        file_name="Сверочный_отчёт.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ─── Подвал ───────────────────────────────────────────────────────────────────

st.divider()
st.markdown(
    '<div style="text-align:center; color:#aaa; font-size:12px;">'
    '🌸 Сверочный отчёт · Данные обрабатываются локально · Никуда не передаются'
    '</div>',
    unsafe_allow_html=True
)
