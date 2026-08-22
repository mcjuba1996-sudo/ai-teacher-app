import io
import json
import urllib.parse
import re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.shared import Inches, Pt
import google.generativeai as genai
from PIL import Image
import pandas as pd
import streamlit as st

import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.ensemble import RandomForestClassifier

# ==========================================
# 0. СЛОВАРЬ ПЕРЕВОДОВ (РУССКИЙ / ҚАЗАҚША)
# ==========================================
translations = {
    "ru": {
        "page_title": "AI-Помощник Учителя",
        "sidebar_title": "🎓 AI-Помощник",
        "api_subheader": "🔑 Доступ к ИИ",
        "api_help": "Введите ключ один раз для всех инструментов",
        "api_expander": "ℹ️ Как получить API ключ бесплатно?",
        "tools_subheader": "🛠️ Инструменты",
        "footer": "✨ Создано для учителей с ❤️",
        "menu": [
            "📝 Генератор карточек",
            "📅 AI-Генератор КТП",
            "📋 AI-Конструктор КСП",
            "📊 Анализ и визуализация (EDA)",
            "🤖 ML-Прогноз уровня ученика",
            "📷 AI-Проверка по фото",
            "👤 Генератор характеристик",
            "⚡ Разминки и интерактивы",
        ],
        "no_key": "Пожалуйста, введите ваш рабочий Gemini API Key в боковом меню слева!",
    },
    "kk": {
        "page_title": "AI Мұғалім Көмекшісі",
        "sidebar_title": "🎓 AI Көмекші",
        "api_subheader": "🔑 Жасанды интеллект кілті",
        "api_help": "Барлық құралдар үшін кілтті бір рет енгізіңіз",
        "api_expander": "ℹ️ API кілтін қалай тегін алуға болады?",
        "tools_subheader": "🛠️ Құралдар",
        "footer": "✨ Мұғалімдер үшін махаббатпен жасалған ❤️",
        "menu": [
            "📝 Тапсырма карточкаларын жасау",
            "📅 КТП AI-Генераторы",
            "📋 ҚМЖ (КСП) AI-Конструкторы",
            "📊 Талдау және визуализация (EDA)",
            "🤖 Оқушы деңгейін ML болжау",
            "📷 Фото арқылы AI тексеру",
            "👤 Мінездеме генераторы",
            "⚡ Сергіту сәттері мен интерактив",
        ],
        "no_key": "Сол жақ бүйірлік мәзірге жұмыс істейтін Gemini API Key енгізіңіз!",
    }
}

# ==========================================
# 1. НАСТРОЙКИ ДИЗАЙНА И ЯЗЫКА
# ==========================================
st.set_page_config(page_title="AI-Помощник Учителя", page_icon="🎓", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;800&display=swap');
    html, body, [class*="css"] { font-family: 'Nunito', sans-serif; }
    .stApp { background: url("https://images.unsplash.com/photo-1557683316-973673baf926?q=80&w=2029&auto=format&fit=crop") no-repeat center center fixed; background-size: cover; }
    .block-container { background-color: rgba(255, 255, 255, 0.93); backdrop-filter: blur(10px); border-radius: 20px; padding: 2.5rem 3rem !important; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.1); }
    [data-testid="stSidebar"] { background-color: rgba(255, 255, 255, 0.85); backdrop-filter: blur(15px); }
    .stButton>button { border-radius: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white !important; font-weight: 700; transition: 0.3s; }
    .stButton>button:hover { transform: translateY(-3px); box-shadow: 0 8px 25px rgba(118, 75, 162, 0.5); }
    h1 { background: -webkit-linear-gradient(45deg, #1e3c72, #2a5298); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

DEFAULT_API_KEY = ""

# ==========================================
# 2. БОКОВОЕ МЕНЮ И ВЫБОР ЯЗЫКА
# ==========================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/1972/1972413.png", width=80)

# Переключатель языка
lang_choice = st.sidebar.selectbox("🌐 Тіл / Язык:", ["Русский", "Қазақша"], index=0)
lang = "ru" if lang_choice == "Русский" else "kk"

st.sidebar.title(translations[lang]["sidebar_title"])
st.sidebar.markdown("---")

st.sidebar.subheader(translations[lang]["api_subheader"])
global_api_key = st.sidebar.text_input("Gemini API Key:", type="password", help=translations[lang]["api_help"])

with st.sidebar.expander(translations[lang]["api_expander"]):
    st.markdown("""
    1. Зайдите на [Google AI Studio](https://aistudio.google.com/app/apikey).
    2. Войдите через Google-аккаунт.
    3. Нажмите синюю кнопку **Create API key**.
    4. Скопируйте ключ и вставьте в поле выше.
    """)

active_key = global_api_key if global_api_key else DEFAULT_API_KEY

st.sidebar.markdown("---")
st.sidebar.subheader(translations[lang]["tools_subheader"])

menu_choice = st.sidebar.radio(
    "Навигация:",
    translations[lang]["menu"],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.caption(translations[lang]["footer"])


# ==========================================
# МОДУЛЬ 1: ГЕНЕРАТОР КАРТОЧЕК
# ==========================================
if menu_choice in ["📝 Генератор карточек", "📝 Тапсырма карточкаларын жасау"]:
    st.title("📝 " + ("Генератор карточек с заданиями" if lang=="ru" else "Тапсырма карточкаларын жасау генераторы"))
    st.markdown("#### " + ("Автоматическая генерация индивидуальных вариантов в Word" if lang=="ru" else "Word форматында жеке нұсқаларды автоматты түрде жасау"))
    st.divider()

    source_type = st.radio("Источник данных / Дереккөз:" if lang=="ru" else "Дереккөз:", ["Google Таблица (ссылка)" if lang=="ru" else "Google Кесте (сілтеме)", "Excel-файл (.xlsx)"], horizontal=True)
    df_questions, df_students = None, None

    if "Google" in source_type:
        default_url = "https://docs.google.com/spreadsheets/d/1fJKlRP7YY3r6DFjd_PuLXFIKkg3GdSRAM9Rxwq502e8/edit?usp=sharing"
        sheet_url = st.text_input("🔗 " + ("Вставьте ссылку на Google Таблицу:" if lang=="ru" else "Google кесте сілтемесін енгізіңіз:"))
        def extract_sheet_id(url):
            try: return url.split("/d/")[1].split("/")[0] if "/d/" in url else url
            except: return None
        if sheet_url:
            sheet_id = extract_sheet_id(sheet_url)
            if sheet_id:
                try:
                    def get_url(name):
                        enc = urllib.parse.quote(name)
                        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={enc}"
                    df_questions = pd.read_csv(get_url("Банк_вопросов"))
                    df_students = pd.read_csv(get_url("Ученики"))
                except: pass
    else:
        uploaded_excel = st.file_uploader("📂 " + ("Загрузите Excel-файл (листы: Банк_вопросов, Ученики):" if lang=="ru" else "Excel файлын жүктеңіз:"), type=["xlsx"])
        if uploaded_excel:
            try:
                xls = pd.ExcelFile(uploaded_excel)
                df_questions = pd.read_excel(xls, 'Банк_вопросов')
                df_students = pd.read_excel(xls, 'Ученики')
                st.success("Excel сәтті оқылды!" if lang=="kk" else "Excel-файл успешно прочитан!")
            except Exception as e: st.error(f"Ошибка: {e}")

    if df_questions is not None and df_students is not None:
        st.markdown("#### ⚙️ " + ("Настройка теста" if lang=="ru" else "Тест баптаулары"))
        col1, col2, col3 = st.columns(3)
        with col1: count_easy = st.number_input("🟢 " + ("Легких:" if lang=="ru" else "Жеңіл:"), min_value=0, max_value=5, value=1)
        with col2: count_med = st.number_input("🟡 " + ("Средних:" if lang=="ru" else "Орташа:"), min_value=0, max_value=5, value=1)
        with col3: count_hard = st.number_input("🔴 " + ("Сложных:" if lang=="ru" else "Қиын:"), min_value=0, max_value=5, value=1)

        if st.button("🚀 " + ("Сгенерировать варианты в Word" if lang=="ru" else "Word форматында жасау"), type="primary", use_container_width=True):
            if not active_key: st.warning(translations[lang]["no_key"])
            else:
                try:
                    with st.spinner("⏳ ИИ формирует варианты..."):
                        df_questions.columns = df_questions.columns.astype(str).str.strip().str.lower()
                        df_students.columns = df_students.columns.astype(str).str.strip().str.lower()
                        rename_dict = {}
                        for col in df_questions.columns:
                            if "сложн" in col or "күрдел" in col: rename_dict[col] = "сложность"
                            elif "тем" in col: rename_dict[col] = "тема"
                            elif "вопрос" in col or "сұрақ" in col: rename_dict[col] = "вопрос"
                            elif "ответ" in col or "жауап" in col: rename_dict[col] = "ответ"
                        df_questions = df_questions.rename(columns=rename_dict)
                        student_col = [c for c in df_students.columns if "фио" in c or "аты" in c]
                        student_col_name = student_col[0] if student_col else df_students.columns[0]
                        students = df_students[student_col_name].dropna().tolist()

                    doc_students = Document()
                    structure = {"Легкий": count_easy, "Средний": count_med, "Сложный": count_hard}
                    teacher_keys = []
                    for student in students:
                        variant_questions = []
                        for level, count in structure.items():
                            if count > 0:
                                subset = df_questions[df_questions["сложность"].astype(str).str.strip().str.capitalize() == level]
                                if len(subset) == 0: subset = df_questions
                                variant_questions.append(subset.sample(n=min(count, len(subset))))
                        student_variant = pd.concat(variant_questions).reset_index(drop=True)
                        title = doc_students.add_heading("Проверочная работа", level=2)
                        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_info = doc_students.add_paragraph()
                        p_info.add_run(f"Ученик(ца) / Оқушы: {student}").bold = True
                        for idx, row in student_variant.iterrows():
                            p_q = doc_students.add_paragraph()
                            p_q.add_run(f"Задание {idx + 1} ").bold = True
                            p_q.add_run(f"{row['вопрос']}\n")
                            p_q.add_run("Ответ / Жауап: ____________________")
                            teacher_keys.append({"Ученик": student, "№ Задания": idx + 1, "Ответ": row["ответ"]})
                        doc_students.add_paragraph("--------------------------------------------------")

                    bio_students = io.BytesIO()
                    doc_students.save(bio_students)
                    bio_students.seek(0)
                    doc_teacher = Document()
                    doc_teacher.add_heading("КЛЮЧИ (ДЛЯ УЧИТЕЛЯ)", level=1)
                    df_keys = pd.DataFrame(teacher_keys)
                    curr_st = ""
                    for _, row in df_keys.iterrows():
                        if row["Ученик"] != curr_st:
                            curr_st = row["Ученик"]
                            doc_teacher.add_paragraph().add_run(f"\n👤 {row['Ученик']}").bold = True
                        doc_teacher.add_paragraph(f"  • Задание {row['№ Задания']}: {row['Ответ']}")
                    bio_teacher = io.BytesIO()
                    doc_teacher.save(bio_teacher)
                    bio_teacher.seek(0)

                    st.success("🎉 Готово!" if lang=="ru" else "🎉 Дайын!")
                    col_d1, col_d2 = st.columns(2)
                    with col_d1: st.download_button("📄 " + ("Скачать Карточки" if lang=="ru" else "Карточкаларды жүктеу"), bio_students, "Карточки.docx", use_container_width=True)
                    with col_d2: st.download_button("🔑 " + ("Скачать Ключи" if lang=="ru" else "Клттерді жүктеу"), bio_teacher, "Ключи.docx", use_container_width=True)
                except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 2: AI-ГЕНЕРАТОР КТП
# ==========================================
elif menu_choice in ["📅 AI-Генератор КТП", "📅 КТП AI-Генераторы"]:
    st.title("📅 AI-Генератор КТП")
    st.divider()
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        subject = st.text_input("📚 Предмет / Пән:", "Информатика")
        grade = st.number_input("🏫 Класс / Сынып:", 1, 11, 8)
    with col_p2:
        quarters_count = st.selectbox("📅 " + ("Четвертей:" if lang=="ru" else "Тоқсан саны:"), [1, 2, 3, 4], index=0)
        hours_per_week = st.number_input("⏰ " + ("Часов в неделю:" if lang=="ru" else "Аптасына сағат:"), 1, 5, 2)
    quarters_weeks = {q: st.number_input(f"{q}-я четверть (недель):" if lang=="ru" else f"{q}-ші тоқсан (апта):", 1, 15, 8) for q in range(1, quarters_count + 1)}
    total_all_lessons = sum(q_w * hours_per_week for q_w in quarters_weeks.values())
    st.info(f"💡 Всего уроков / Барлық сағат: **{total_all_lessons}**")

    source_type = st.radio("Источник / Дереккөз:", ["Загрузить PDF" if lang=="ru" else "PDF жүктеу", "Текст" if lang=="ru" else "Мәтін"], horizontal=True)
    uploaded_pdf, textbook_content = None, ""
    if "PDF" in source_type: uploaded_pdf = st.file_uploader("📂 PDF:", type=["pdf"])
    else: textbook_content = st.text_area("📝 " + ("Темы:" if lang=="ru" else "Тақырыптар:"), "1. Инфо 2. Алгоритмы", height=100)

    if st.button("🚀 " + ("Сгенерировать КТП" if lang=="ru" else "КТП құру"), type="primary", use_container_width=True):
        if not active_key: st.warning(translations[lang]["no_key"])
        else:
            try:
                with st.spinner("⏳ ИИ анализирует..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"Составь КТП по предмету {subject}, {grade} класс, уроков: {total_all_lessons}. Темы: {textbook_content}. Верни строго JSON массив: [{{'quarter':1, 'lesson_num':1, 'topic':'', 'targets':'', 'homework':''}}]"
                    response = model.generate_content([prompt, uploaded_pdf] if uploaded_pdf else prompt)
                    match = re.search(r'\[.*\]', response.text, re.DOTALL)
                    ktp_data = json.loads(match.group(0) if match else response.text)
                    df_ktp = pd.DataFrame(ktp_data)
                    st.dataframe(df_ktp, use_container_width=True)
                    bio = io.BytesIO()
                    with pd.ExcelWriter(bio, engine="openpyxl") as w: df_ktp.to_excel(w, index=False)
                    bio.seek(0)
                    st.download_button("📊 Скачать КТП (Excel)", bio, f"КТП_{subject}.xlsx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 3: AI-КОНСТРУКТОР КСП
# ==========================================
elif menu_choice in ["📋 AI-Конструктор КСП", "📋 ҚМЖ (КСП) AI-Конструкторы"]:
    st.title("📋 AI-Конструктор КСП (ҚМЖ)")
    st.divider()
    col_k1, col_k2 = st.columns(2)
    with col_k1:
        teacher_name = st.text_input("ФИО учителя / Мұғалім:", "Иванов И.И.")
        subject_ksp = st.text_input("Предмет / Пән:", "Информатика")
        grade_ksp = st.number_input("Класс / Сынып:", 1, 11, 8)
    with col_k2:
        topic_ksp = st.text_input("Тема / Тақырып:", "Двоичная система")
        target_ksp = st.text_input("ЦО / ОМ:", "8.1.1.1 Перевод чисел")

    if st.button("🚀 " + ("Сгенерировать КСП" if lang=="ru" else "ҚМЖ құру"), type="primary", use_container_width=True):
        if not active_key: st.warning(translations[lang]["no_key"])
        else:
            try:
                with st.spinner("⏳ Создание КСП..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"Создай план урока по предмету {subject_ksp}, тема {topic_ksp}. Верни строго JSON: {{\"lesson_targets\":\"...\", \"eval_criteria\":\"...\", \"stages\":[{{\"time\":\"Начало\", \"teacher\":\"...\", \"student\":\"...\", \"eval\":\"...\", \"resources\":\"...\"}}]}}"
                    response = model.generate_content(prompt)
                    match = re.search(r'\{.*\}', response.text, re.DOTALL)
                    ksp_data = json.loads(match.group(0) if match else response.text)

                    doc_ksp = Document()
                    section = doc_ksp.sections[-1]
                    section.orientation = WD_ORIENT.LANDSCAPE
                    section.page_width, section.page_height = section.page_height, section.page_width
                    
                    title = doc_ksp.add_paragraph()
                    r = title.add_run("КРАТКОСРОЧНЫЙ ПЛАН УРОКА (ҚМЖ)")
                    r.font.bold = True
                    r.font.size = Pt(14)
                    r.font.name = 'Times New Roman'
                    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    t_table = doc_ksp.add_table(rows=7, cols=2)
                    t_table.style = 'Table Grid'
                    info = [("Мұғалім:", teacher_name), ("Пән:", subject_ksp), ("Сынып:", str(grade_ksp)), ("Тақырып:", topic_ksp), ("ОМ:", target_ksp), ("Максат:", ksp_data.get("lesson_targets","")), ("Критерий:", ksp_data.get("eval_criteria",""))]
                    for idx, (l, v) in enumerate(info):
                        t_table.rows[idx].cells[0].text = l
                        t_table.rows[idx].cells[1].text = str(v)

                    doc_ksp.add_paragraph()
                    s_table = doc_ksp.add_table(rows=1, cols=5)
                    s_table.style = 'Table Grid'
                    headers = ["Кезені", "Мұғалім әрекеті", "Оқушы әрекеті", "Бағалау", "Ресурстар"]
                    for i, h in enumerate(headers): s_table.rows[0].cells[i].text = h

                    for stg in ksp_data.get("stages", []):
                        row = s_table.add_row().cells
                        row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = stg.get("time",""), stg.get("teacher",""), stg.get("student",""), stg.get("eval",""), stg.get("resources","")

                    bio = io.BytesIO()
                    doc_ksp.save(bio)
                    bio.seek(0)
                    st.success("Готово!")
                    st.download_button("📄 Скачать КСП (Word)", bio, f"ҚМЖ_{topic_ksp}.docx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 4: EDA
# ==========================================
elif menu_choice in ["📊 Анализ и визуализация (EDA)", "📊 Талдау және визуализация (EDA)"]:
    st.title("📊 " + ("Анализ и визуализация успеваемости" if lang=="ru" else "Үлгерімді талдау және визуализация"))
    st.divider()
    uploaded_eda = st.file_uploader("📂 Excel (.xlsx)", type=["xlsx"])
    if uploaded_eda:
        df_eda = pd.read_excel(uploaded_eda)
        st.write(df_eda.head())
        num_cols = df_eda.select_dtypes(include=['number']).columns.tolist()
        if num_cols:
            col = st.selectbox("Колонка:", num_cols)
            if st.button("📈 График", type="primary"):
                with st.spinner("⏳ Анализ..."):
                    st.write(df_eda[col].describe())
                    fig, ax = plt.subplots(figsize=(6, 4))
                    sns.histplot(df_eda[col], kde=True, ax=ax, color='purple')
                    st.pyplot(fig)

# ==========================================
# МОДУЛЬ 5: ML
# ==========================================
elif menu_choice in ["🤖 ML-Прогноз уровня ученика", "🤖 Оқушы деңгейін ML болжау"]:
    st.title("🤖 " + ("ML-Прогноз уровня ученика" if lang=="ru" else "Оқушы деңгейін ML болжау"))
    st.divider()
    att = st.slider("Посещаемость (%):", 50, 100, 85)
    hw = st.slider("ДЗ (%):", 0, 100, 75)
    test = st.slider("Тест балл:", 0, 100, 80)
    if st.button("🔮 Болжау", type="primary"):
        X_train = [[60, 50, 55], [90, 85, 88], [70, 60, 65], [95, 95, 92]]
        y_train = ["Қолдау", "Жоғары", "Орташа", "Жоғары"]
        model = RandomForestClassifier().fit(X_train, y_train)
        pred = model.predict([[att, hw, test]])[0]
        st.success(f"🎯 Нәтиже / Результат: **{pred}**")

# ==========================================
# МОДУЛЬ 6: AI-ПРОВЕРКА ПО ФОТО
# ==========================================
elif menu_choice in ["📷 AI-Проверка по фото", "📷 Фото арқылы AI тексеру"]:
    st.title("📷 " + ("AI-Проверка работ" if lang=="ru" else "Жұмыстарды AI тексеру"))
    st.divider()
    img = st.file_uploader("Фото:", type=["jpg", "png"])
    if img and st.button("Тексеру", type="primary"):
        if not active_key: st.warning(translations[lang]["no_key"])
        else:
            with st.spinner("⏳ Тексеруде..."):
                genai.configure(api_key=active_key)
                model = genai.GenerativeModel("gemini-3.6-flash")
                res = model.generate_content(["Проверь работу ученика:", Image.open(img)])
                st.markdown(res.text)

# ==========================================
# МОДУЛЬ 7: ХАРАКТЕРИСТИКА
# ==========================================
elif menu_choice in ["👤 Генератор характеристик", "👤 Мінездеме генераторы"]:
    st.title("👤 " + ("Генератор характеристик" if lang=="ru" else "Мінездеме генераторы"))
    st.divider()
    name = st.text_input("Аты-жөні / ФИО:", "Иванов Иван")
    cls = st.text_input("Сынып / Класс:", "8 «А»")
    att = st.slider("Сабаққа қатысу / Посещаемость (%):", 0, 100, 90)
    perf = st.selectbox("Үлгерім / Успеваемость:", ["Үздік / Отличник", "Екпінді / Ударник", "Орташа / Средняя"])
    beh = st.selectbox("Тәртіп / Поведение:", ["Үлгілі / Примерное", "Жақсы / Хорошее"])
    traits = st.text_area("Қосымша / Дополнительно:", "Олимпиада қатысушысы...")

    if st.button("🚀 Жасау", type="primary"):
        if not active_key: st.warning(translations[lang]["no_key"])
        else:
            with st.spinner("⏳ Жасалуда..."):
                genai.configure(api_key=active_key)
                model = genai.GenerativeModel("gemini-3.6-flash")
                res = model.generate_content(f"Напиши характеристику на ученика {name}, класс {cls}. Посещаемость: {att}%, успеваемость: {perf}, поведение: {beh}, доп: {traits}.")
                st.markdown(res.text)

# ==========================================
# МОДУЛЬ 8: РАЗМИНКИ
# ==========================================
elif menu_choice in ["⚡ Разминки и интерактивы", "⚡ Сергіту сәттері мен интерактив"]:
    st.title("⚡ " + ("AI-Разминки" if lang=="ru" else "AI Сергіту сәттері"))
    st.divider()
    top = st.text_input("Тақырып / Тема:", "Алгоритмдер")
    tm = st.slider("Уақыт / Время (мин):", 2, 10, 5)
    if st.button("Таңдау", type="primary"):
        if not active_key: st.warning(translations[lang]["no_key"])
        else:
            with st.spinner("⏳ Дайындалуда..."):
                genai.configure(api_key=active_key)
                model = genai.GenerativeModel("gemini-3.6-flash")
                res = model.generate_content(f"Предложи 3 разминки на тему {top} на {tm} минут.")
                st.markdown(res.text)
