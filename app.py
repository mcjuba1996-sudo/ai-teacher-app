%%writefile app.py
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
import streamlit.components.v1 as components

import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.ensemble import RandomForestClassifier

# ==========================================
# 0. СЛОВАРЬ ПЕРЕВОДОВ ИНТЕРФЕЙСА (РАЗДЕЛЬНЫЙ)
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
        "ai_lang_prompt": "Напиши ответ строго на русском языке.",
        
        # Общие тексты
        "source": "Источник данных:",
        "source_options": ["Google Таблица (ссылка)", "Excel-файл (.xlsx)"],
        "sheet_link": "Вставьте ссылку на Google Таблицу:",
        "upload_excel": "Загрузите Excel-файл (листы: Банк_вопросов, Ученики):",
        "success_excel": "Excel-файл успешно прочитан!",
        "settings": "Настройка теста",
        "easy": "Легких вопросов:",
        "med": "Средних вопросов:",
        "hard": "Сложных вопросов:",
        "gen_word": "Сгенерировать варианты в Word",
        "wait_ai": "ИИ обрабатывает банк вопросов и формирует индивидуальные варианты...",
        "student_lbl": "Ученик(ца):",
        "task_lbl": "Задание",
        "answer_lbl": "Ответ: ____________________",
        "keys_title": "КЛЮЧИ (ДЛЯ УЧИТЕЛЯ)",
        "done": "Готово!",
        "download_cards": "Скачать Карточки (Word)",
        "download_keys": "Скачать Ключи (Word)",
        
        "subject": "Предмет:",
        "grade": "Класс:",
        "quarters": "Количество четвертей:",
        "hours": "Часов в неделю:",
        "total_lessons": "Всего уроков:",
        "source_pdf_text": "Источник тем:",
        "pdf_opt": ["Загрузить PDF-файл", "Ввести темы текстом"],
        "topics_lbl": "Темы:",
        "gen_ktp": "Сгенерировать КТП в Word",
        "wait_ktp": "ИИ анализирует учебные материалы и формирует КТП...",
        "download_ktp": "Скачать КТП (Word)",
        
        "teacher_name": "ФИО учителя:",
        "topic_lbl": "Тема урока:",
        "target_lbl": "Цели обучения (ЦО):",
        "gen_ksp": "Сгенерировать КСП в Word",
        "wait_ksp": "ИИ методист разрабатывает этапы урока, дескрипторы и дифференциацию...",
        "download_ksp": "Скачать КСП (Word)",
        
        "eda_title": "Анализ и визуализация успеваемости класса",
        "eda_sub": "Загрузите Excel-файл с оценками учеников для описательной статистики и графиков",
        "eda_load": "Загрузите файл с оценками (.xlsx)",
        "eda_select": "Выберите числовой показатель для анализа:",
        "eda_btn": "Построить графики и статистику",
        "eda_wait": "Выполняется расчет статистических метрик и генерация графиков...",
        "hist": "Гистограмма распределения",
        "box": "Ящик с усами (Boxplot)",
        
        "ml_title": "Машинное обучение в образовании (Scikit-learn)",
        "ml_sub": "Прогнозирование группы поддержки / продвинутого уровня с помощью Random Forest",
        "ml_txt": "Введите показатели ученика, чтобы модель машинного обучения определила его уровень:",
        "att": "Посещаемость уроков (%):",
        "hw": "Выполнение домашних заданий (%):",
        "test": "Средний балл по тестам:",
        "activity": "Классная активность:",
        "act_opts": ["Низкая", "Средняя", "Высокая"],
        "ml_btn": "Предсказать уровень ученика",
        "ml_wait": "ML-модель анализирует показатели и выстраивает классификацию...",
        "rec": "Рекомендация ML-модели:",
        
        "photo_title": "AI-Проверка письменных работ",
        "photo_load": "Фото работы:",
        "photo_check": "Проверить работу",
        "photo_wait": "Мультимодальный ИИ распознает почерк и проверяет ход решения...",
        
        "char_title": "Генератор педагогических характеристик",
        "char_sub": "Составление развернутого отчета с учетом посещаемости, успеваемости и поведения",
        "name_lbl": "ФИО ученика:",
        "cls_lbl": "Класс:",
        "att_lbl": "Посещаемость занятий (%):",
        "perf_lbl": "Успеваемость:",
        "perf_opts": ["Отличник", "Ударник", "Занимается средне", "Имеет академические задолженности"],
        "beh_lbl": "Поведение и дисциплина:",
        "beh_opts": ["Дисциплинирован, примерное поведение", "Спокойный, исполнительный", "Иногда нарушает дисциплину", "Требует повышенного педагогического внимания"],
        "traits_lbl": "Дополнительные достижения, склонности и заметки:",
        "char_btn": "Сгенерировать характеристику",
        "char_wait": "ИИ классный руководитель формирует текст педагогической характеристики...",
        
        "warm_title": "AI-Генератор разминок и Icebreakers",
        "warm_sub": "Интерактивы и разминки для начала урока с учетом тайминга",
        "warm_top": "Тема урока:",
        "warm_time": "Время на разминку (минут):",
        "warm_btn": "Подобрать разминку",
        "warm_wait": "ИИ подбирает креативные интерактивы...",
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
        "ai_lang_prompt": "Жауапты қатаң түрде қазақ тілінде жаз.",
        
        # Общие тексты
        "source": "Дереккөз:",
        "source_options": ["Google Кесте (сілтеме)", "Excel-файл (.xlsx)"],
        "sheet_link": "Google кестенің сілтемесін енгізіңіз:",
        "upload_excel": "Excel файлын жүктеңіз (парақтар: Банк_вопросов, Ученики):",
        "success_excel": "Excel файлы сәтті оқылды!",
        "settings": "Тест баптаулары",
        "easy": "Жеңіл сұрақтар:",
        "med": "Орташа сұрақтар:",
        "hard": "Қиын сұрақтар:",
        "gen_word": "Word форматында нұсқаларды жасау",
        "wait_ai": "ЖИ сұрақтар банкін өңдеп, жеке нұсқаларды дайындауда...",
        "student_lbl": "Оқушы:",
        "task_lbl": "Тапсырма",
        "answer_lbl": "Жауап: ____________________",
        "keys_title": "ЖАУАПТАР (МҰҒАЛІМГЕ ҮШІН)",
        "done": "Дайын!",
        "download_cards": "Карточкаларды жүктеу (Word)",
        "download_keys": "Жауаптарды жүктеу (Word)",
        
        "subject": "Пән:",
        "grade": "Сынып:",
        "quarters": "Тоқсан саны:",
        "hours": "Аптасына сағат:",
        "total_lessons": "Барлық сағат:",
        "source_pdf_text": "Тақырыптар көзі:",
        "pdf_opt": ["PDF файлын жүктеу", "Тақырыптарды мәтін түрінде енгізу"],
        "topics_lbl": "Тақырыптар:",
        "gen_ktp": "Word форматында КТП құру",
        "wait_ktp": "ЖИ оқу материалдарын талдап, КТП құруда...",
        "download_ktp": "КТП жүктеу (Word)",
        
        "teacher_name": "Мұғалімнің аты-жөні:",
        "topic_lbl": "Сабақ тақырыбы:",
        "target_lbl": "Оқыту мақсаттары (ОМ):",
        "gen_ksp": "Word форматында ҚМЖ құру",
        "wait_ksp": "ЖИ әдіскер сабақ кезеңдерін, дескрипторларды әзірлеуде...",
        "download_ksp": "ҚМЖ жүктеу (Word)",
        
        "eda_title": "Сыныптың үлгерімін талдау және визуализация",
        "eda_sub": "Сипаттамалық статистика мен графиктер үшін оқушылардың бағалары бар Excel файлын жүктеңіз",
        "eda_load": "Бағалар файлын жүктеңіз (.xlsx)",
        "eda_select": "Талдау үшін сандық көрсеткішті таңдаңыз:",
        "eda_btn": "Графиктер мен статистиканы құру",
        "eda_wait": "Статистикалық метрикалар есептеліп, графиктер жасалуда...",
        "hist": "Бөлініс гистограммасы",
        "box": "Жәшік диаграммасы (Boxplot)",
        
        "ml_title": "Білім берудегі машиналық оқыту (Scikit-learn)",
        "ml_sub": "Random Forest көмегімен қолдау тобын / жоғары деңгейді болжау",
        "ml_txt": "Машиналық оқыту моделі оқушының деңгейін анықтауы үшін көрсеткіштерін енгізіңіз:",
        "att": "Сабаққа қатысу (%):",
        "hw": "Үй жұмысын орындау (%):",
        "test": "Тесттердің орташа балы:",
        "activity": "Сыныптағы белсенділік:",
        "act_opts": ["Төмен", "Орташа", "Жоғары"],
        "ml_btn": "Оқушы деңгейін болжау",
        "ml_wait": "ML моделі көрсеткіштерді талдап, жіктеуді жүргізуде...",
        "rec": "ML моделінің ұсынысы:",
        
        "photo_title": "Жазбаша жұмыстарды AI тексеру",
        "photo_load": "Жұмыс фотосы:",
        "photo_check": "Жұмысты тексеру",
        "photo_wait": "Мультимодальды ЖИ қолжазбаны танып, шешу жолын тексеруде...",
        
        "char_title": "Педагогикалық мінездеме генераторы",
        "char_sub": "Сабаққа қатысуды, үлгерімді және тәртіпті ескере отырып мінездеме жасау",
        "name_lbl": "Оқушының аты-жөні:",
        "cls_lbl": "Сынып:",
        "att_lbl": "Сабаққа қатысу (%):",
        "perf_lbl": "Үлгерім:",
        "perf_opts": ["Үздік", "Екпінді", "Орташа оқиды", "Академиялық қарыздары бар"],
        "beh_lbl": "Тәртіп пен мінез-құлық:",
        "beh_opts": ["Тәртіпті, үлгілі мінез-құлық", "Сабырлы, жауапкершілікті", "Кейде тәртіпті бұзады", "Педагогикалық бақылауды қажет етеді"],
        "traits_lbl": "Қосымша жетістіктер, бейімділіктер және ескертпелер:",
        "char_btn": "Мінездеме құру",
        "char_wait": "ЖИ сынып жетекшісі мінездеме мәтінін дайындауда...",
        
        "warm_title": "AI Сергіту сәттері мен Icebreakers генераторы",
        "warm_sub": "Уақытты ескере отырып, сабақ басына арналған интерактивтер",
        "warm_top": "Сабақ тақырыбы:",
        "warm_time": "Сергіту сәтіне берілген уақыт (минут):",
        "warm_btn": "Сергіту сәтін таңдау",
        "warm_wait": "ЖИ креативті интерактивтерді таңдауда...",
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

lang_choice = st.sidebar.selectbox("🌐 Тіл / Язык:", ["Русский", "Қазақша"], index=0)
lang = "ru" if lang_choice == "Русский" else "kk"
t = translations[lang]

st.sidebar.title(t["sidebar_title"])
st.sidebar.markdown("---")

st.sidebar.subheader(t["api_subheader"])
global_api_key = st.sidebar.text_input("Gemini API Key:", type="password", help=t["api_help"])

with st.sidebar.expander(t["api_expander"]):
    st.markdown("""
    1. Зайдите на [Google AI Studio](https://aistudio.google.com/app/apikey).
    2. Войдите через Google-аккаунт.
    3. Нажмите синюю кнопку **Create API key**.
    4. Скопируйте ключ и вставьте в поле выше.
    """)

active_key = global_api_key if global_api_key else DEFAULT_API_KEY

st.sidebar.markdown("---")
st.sidebar.subheader(t["tools_subheader"])

menu_choice = st.sidebar.radio(
    "Навигация:",
    t["menu"],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.caption(t["footer"])


# ==========================================
# МОДУЛЬ 1: ГЕНЕРАТОР КАРТОЧЕК
# ==========================================
if menu_choice in ["📝 Генератор карточек", "📝 Тапсырма карточкаларын жасау"]:
    st.title("📝 " + menu_choice)
    st.markdown("#### " + ("Автоматическая генерация индивидуальных вариантов в Word" if lang=="ru" else "Word форматында жеке нұсқаларды автоматты түрде жасау"))
    st.divider()

    source_type = st.radio(t["source"], t["source_options"], horizontal=True)
    df_questions, df_students = None, None

    if "Google" in source_type:
        default_url = "https://docs.google.com/spreadsheets/d/1fJKlRP7YY3r6DFjd_PuLXFIKkg3GdSRAM9Rxwq502e8/edit?usp=sharing"
        sheet_url = st.text_input("🔗 " + t["sheet_link"], value=default_url)
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
        uploaded_excel = st.file_uploader("📂 " + t["upload_excel"], type=["xlsx"])
        if uploaded_excel:
            try:
                xls = pd.ExcelFile(uploaded_excel)
                df_questions = pd.read_excel(xls, 'Банк_вопросов')
                df_students = pd.read_excel(xls, 'Ученики')
                st.success(t["success_excel"])
            except Exception as e: st.error(f"Ошибка: {e}")

    if df_questions is not None and df_students is not None:
        st.markdown("#### ⚙️ " + t["settings"])
        col1, col2, col3 = st.columns(3)
        with col1: count_easy = st.number_input("🟢 " + t["easy"], min_value=0, max_value=5, value=1)
        with col2: count_med = st.number_input("🟡 " + t["med"], min_value=0, max_value=5, value=1)
        with col3: count_hard = st.number_input("🔴 " + t["hard"], min_value=0, max_value=5, value=1)

        if st.button("🚀 " + t["gen_word"], type="primary", use_container_width=True):
            if not active_key: st.warning(t["no_key"])
            else:
                try:
                    with st.spinner("⏳ " + t["wait_ai"]):
                        df_questions.columns = df_questions.columns.astype(str).str.strip().str.lower()
                        df_students.columns = df_students.columns.astype(str).str.strip().str.lower()
                        rename_dict = {}
                        for col in df_questions.columns:
                            if "сложн" in col or "күрдел" in col: rename_dict[col] = "сложность"
                            elif "тем" in col or "тақырып" in col: rename_dict[col] = "тема"
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
                        title = doc_students.add_heading("Проверочная работа" if lang=="ru" else "Бақылау жұмысы", level=2)
                        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_info = doc_students.add_paragraph()
                        p_info.add_run(f"{t['student_lbl']} {student}").bold = True
                        for idx, row in student_variant.iterrows():
                            p_q = doc_students.add_paragraph()
                            p_q.add_run(f"{t['task_lbl']} {idx + 1} ").bold = True
                            p_q.add_run(f"{row['вопрос']}\n")
                            p_q.add_run(t["answer_lbl"])
                            teacher_keys.append({"Ученик": student, "№ Задания": idx + 1, "Ответ": row["ответ"]})
                        doc_students.add_paragraph("--------------------------------------------------")

                    bio_students = io.BytesIO()
                    doc_students.save(bio_students)
                    bio_students.seek(0)
                    doc_teacher = Document()
                    doc_teacher.add_heading(t["keys_title"], level=1)
                    df_keys = pd.DataFrame(teacher_keys)
                    curr_st = ""
                    for _, row in df_keys.iterrows():
                        if row["Ученик"] != curr_st:
                            curr_st = row["Ученик"]
                            doc_teacher.add_paragraph().add_run(f"\n👤 {row['Ученик']}").bold = True
                        doc_teacher.add_paragraph(f"  • {t['task_lbl']} {row['№ Задания']}: {row['Ответ']}")
                    bio_teacher = io.BytesIO()
                    doc_teacher.save(bio_teacher)
                    bio_teacher.seek(0)

                    st.success("🎉 " + t["done"])
                    col_d1, col_d2 = st.columns(2)
                    with col_d1: st.download_button("📄 " + t["download_cards"], bio_students, "Карточки.docx", use_container_width=True)
                    with col_d2: st.download_button("🔑 " + t["download_keys"], bio_teacher, "Ключи.docx", use_container_width=True)
                except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 2: AI-ГЕНЕРАТОР КТП (WORD)
# ==========================================
elif menu_choice in ["📅 AI-Генератор КТП", "📅 КТП AI-Генераторы"]:
    st.title("📅 " + menu_choice)
    st.divider()
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        subject = st.text_input(t["subject"], "Информатика")
        grade = st.number_input(t["grade"], 1, 11, 8)
    with col_p2:
        quarters_count = st.selectbox(t["quarters"], [1, 2, 3, 4], index=0)
        hours_per_week = st.number_input(t["hours"], 1, 5, 2)
    quarters_weeks = {q: st.number_input(f"{q}-я четверть (недель):" if lang=="ru" else f"{q}-ші тоқсан (апта):", 1, 15, 8) for q in range(1, quarters_count + 1)}
    total_all_lessons = sum(q_w * hours_per_week for q_w in quarters_weeks.values())
    st.info(f"💡 {t['total_lessons']} **{total_all_lessons}**")

    source_type = st.radio(t["source_pdf_text"], t["pdf_opt"], horizontal=True)
    uploaded_pdf, textbook_content = None, ""
    if "PDF" in source_type or "PDF" in source_type: uploaded_pdf = st.file_uploader("📂 PDF:", type=["pdf"])
    else: textbook_content = st.text_area(t["topics_lbl"], "1. Инфо 2. Алгоритмы", height=100)

    if st.button("🚀 " + t["gen_ktp"], type="primary", use_container_width=True):
        if not active_key: st.warning(t["no_key"])
        else:
            try:
                with st.spinner("⏳ " + t["wait_ktp"]):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"{t['ai_lang_prompt']} Составь КТП по предмету {subject}, {grade} класс, уроков: {total_all_lessons}. Темы: {textbook_content}. Верни строго JSON массив: [{{'quarter':1, 'lesson_num':1, 'topic':'', 'targets':'', 'homework':''}}]"
                    response = model.generate_content([prompt, uploaded_pdf] if uploaded_pdf else prompt)
                    match = re.search(r'\[.*\]', response.text, re.DOTALL)
                    ktp_data = json.loads(match.group(0) if match else response.text)
                    
                    df_ktp = pd.DataFrame(ktp_data)
                    st.dataframe(df_ktp, use_container_width=True)

                    doc_ktp = Document()
                    section = doc_ktp.sections[-1]
                    section.orientation = WD_ORIENT.LANDSCAPE
                    section.page_width, section.page_height = section.page_height, section.page_width

                    title = doc_ktp.add_paragraph()
                    r = title.add_run("КАЛЕНДАРНО-ТЕМАТИЧЕСКОЕ ПЛАНИРОВАНИЕ (КТП)\n" if lang=="ru" else "КҮНТІЗБЕЛІК-ТАҚЫРЫПТЫҚ ЖОСПАР (КТП)\n")
                    r.font.bold = True
                    r.font.size = Pt(14)
                    r.font.name = 'Times New Roman'
                    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    doc_ktp.add_paragraph()

                    table = doc_ktp.add_table(rows=1, cols=5)
                    table.style = 'Table Grid'
                    headers = ["Четверть / Тоқсан", "№", "Тема урока / Сабақ тақырыбы", "Цели обучения / Оқыту мақсаттары", "ДЗ / Үй жұмысы"]
                    for i, h in enumerate(headers):
                        table.rows[0].cells[i].text = h

                    for item in ktp_data:
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(item.get("quarter", ""))
                        row_cells[1].text = str(item.get("lesson_num", ""))
                        row_cells[2].text = str(item.get("topic", ""))
                        row_cells[3].text = str(item.get("targets", ""))
                        row_cells[4].text = str(item.get("homework", ""))

                    bio = io.BytesIO()
                    doc_ktp.save(bio)
                    bio.seek(0)
                    
                    st.success("🎉 " + t["done"])
                    st.download_button("📄 " + t["download_ktp"], bio, f"КТП_{subject}.docx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 3: AI-КОНСТРУКТОР КСП
# ==========================================
elif menu_choice in ["📋 AI-Конструктор КСП", "📋 ҚМЖ (КСП) AI-Конструкторы"]:
    st.title("📋 " + menu_choice)
    st.divider()
    col_k1, col_k2 = st.columns(2)
    with col_k1:
        teacher_name = st.text_input(t["teacher_name"], "Иванов И.И.")
        subject_ksp = st.text_input(t["subject"], "Информатика")
        grade_ksp = st.number_input(t["grade"], 1, 11, 8)
    with col_k2:
        topic_ksp = st.text_input(t["topic_lbl"], "Двоичная система")
        target_ksp = st.text_input(t["target_lbl"], "8.1.1.1 Перевод чисел")

    if st.button("🚀 " + t["gen_ksp"], type="primary", use_container_width=True):
        if not active_key: st.warning(t["no_key"])
        else:
            try:
                with st.spinner("⏳ " + t["wait_ksp"]):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"{t['ai_lang_prompt']} Создай план урока по предмету {subject_ksp}, тема {topic_ksp}. Верни строго JSON: {{\"lesson_targets\":\"...\", \"eval_criteria\":\"...\", \"stages\":[{{\"time\":\"Начало\", \"teacher\":\"...\", \"student\":\"...\", \"eval\":\"...\", \"resources\":\"...\"}}]}}"
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
                    info = [("Мұғалім / Учитель:", teacher_name), ("Пән / Предмет:", subject_ksp), ("Сынып / Класс:", str(grade_ksp)), ("Тақырып / Тема:", topic_ksp), ("ОМ / ЦО:", target_ksp), ("Мақсат / Цели:", ksp_data.get("lesson_targets","")), ("Критерий:", ksp_data.get("eval_criteria",""))]
                    for idx, (l, v) in enumerate(info):
                        t_table.rows[idx].cells[0].text = l
                        t_table.rows[idx].cells[1].text = str(v)

                    doc_ksp.add_paragraph()
                    s_table = doc_ksp.add_table(rows=1, cols=5)
                    s_table.style = 'Table Grid'
                    headers = ["Этап / Кезең", "Действия учителя", "Действия ученика", "Оценивание", "Ресурсы"]
                    for i, h in enumerate(headers): s_table.rows[0].cells[i].text = h

                    for stg in ksp_data.get("stages", []):
                        row = s_table.add_row().cells
                        row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = stg.get("time",""), stg.get("teacher",""), stg.get("student",""), stg.get("eval",""), stg.get("resources","")

                    bio = io.BytesIO()
                    doc_ksp.save(bio)
                    bio.seek(0)
                    st.success("🎉 " + t["done"])
                    st.download_button("📄 " + t["download_ksp"], bio, f"ҚМЖ_{topic_ksp}.docx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 4: EDA
# ==========================================
elif menu_choice in ["📊 Анализ и визуализация (EDA)", "📊 Талдау және визуализация (EDA)"]:
    st.title("📊 " + menu_choice)
    st.markdown("#### " + t["eda_sub"])
    st.divider()
    uploaded_eda = st.file_uploader("📂 " + t["eda_load"], type=["xlsx"])
    if uploaded_eda:
        df_eda = pd.read_excel(uploaded_eda)
        st.write(df_eda.head())
        num_cols = df_eda.select_dtypes(include=['number']).columns.tolist()
        if num_cols:
            col = st.selectbox(t["eda_select"], num_cols)
            if st.button("📈 " + t["eda_btn"], type="primary"):
                with st.spinner("⏳ " + t["eda_wait"]):
                    st.write(df_eda[col].describe())
                    col_g1, col_g2 = st.columns(2)
                    with col_g1:
                        st.markdown(f"##### {t['hist']}")
                        fig, ax = plt.subplots(figsize=(6, 4))
                        sns.histplot(df_eda[col], kde=True, ax=ax, color='purple')
                        st.pyplot(fig)
                    with col_g2:
                        st.markdown(f"##### {t['box']}")
                        fig, ax = plt.subplots(figsize=(6, 4))
                        sns.boxplot(y=df_eda[col], ax=ax, color='skyblue')
                        st.pyplot(fig)

# ==========================================
# МОДУЛЬ 5: ML
# ==========================================
elif menu_choice in ["🤖 ML-Прогноз уровня ученика", "🤖 Оқушы деңгейін ML болжау"]:
    st.title("🤖 " + menu_choice)
    st.markdown("#### " + t["ml_sub"])
    st.divider()
    st.write(t["ml_txt"])
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        att = st.slider(t["att"], 50, 100, 85)
        hw = st.slider(t["hw"], 0, 100, 75)
    with col_m2:
        test = st.slider(t["test"], 0, 100, 80)
        activity = st.selectbox(t["activity"], t["act_opts"])
        act_val = 1 if activity in ["Низкая", "Төмен"] else (2 if activity in ["Средняя", "Орташа"] else 3)

    if st.button("🔮 " + t["ml_btn"], type="primary", use_container_width=True):
        with st.spinner("⏳ " + t["ml_wait"]):
            X_train = [[60, 50, 55, 1], [90, 85, 88, 3], [70, 60, 65, 2], [95, 95, 92, 3]]
            y_train = ["Группа поддержки" if lang=="ru" else "Қолдау тобы", "Продвинутый" if lang=="ru" else "Жоғары", "Стандартный" if lang=="ru" else "Стандартты", "Продвинутый" if lang=="ru" else "Жоғары"]
            model = RandomForestClassifier(random_state=42).fit(X_train, y_train)
            pred = model.predict([[att, hw, test, act_val]])[0]
        st.success(f"🎯 {t['rec']} **{pred}**")

# ==========================================
# МОДУЛЬ 6: AI-ПРОВЕРКА ПО ФОТО
# ==========================================
elif menu_choice in ["📷 AI-Проверка по фото", "📷 Фото арқылы AI тексеру"]:
    st.title("📷 " + menu_choice)
    st.divider()
    img = st.file_uploader("📂 " + t["photo_load"], type=["jpg", "png"])
    if img and st.button(t["photo_check"], type="primary"):
        if not active_key: st.warning(t["no_key"])
        else:
            with st.spinner("⏳ " + t["photo_wait"]):
                genai.configure(api_key=active_key)
                model = genai.GenerativeModel("gemini-3.6-flash")
                res = model.generate_content([f"{t['ai_lang_prompt']} Проверь работу ученика:", Image.open(img)])
                st.markdown(res.text)

# ==========================================
# МОДУЛЬ 7: ХАРАКТЕРИСТИКА
# ==========================================
elif menu_choice in ["👤 Генератор характеристик", "👤 Мінездеме генераторы"]:
    st.title("👤 " + menu_choice)
    st.markdown("#### " + t["char_sub"])
    st.divider()
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        name = st.text_input(t["name_lbl"], "Иванов Иван")
        cls = st.text_input(t["cls_lbl"], "8 «А»")
        att = st.slider(t["att_lbl"], 0, 100, 90)
    with col_h2:
        perf = st.selectbox(t["perf_lbl"], t["perf_opts"])
        beh = st.selectbox(t["beh_lbl"], t["beh_opts"])

    traits = st.text_area(t["traits_lbl"], "...")

    if st.button("🚀 " + t["char_btn"], type="primary", use_container_width=True):
        if not active_key: st.warning(t["no_key"])
        else:
            with st.spinner("⏳ " + t["char_wait"]):
                genai.configure(api_key=active_key)
                model = genai.GenerativeModel("gemini-3.6-flash")
                prompt = f"{t['ai_lang_prompt']} Напиши официальную характеристику на ученика {name}, класс {cls}. Посещаемость: {att}%, успеваемость: {perf}, поведение: {beh}, доп: {traits}."
                res = model.generate_content(prompt)
                st.markdown(res.text)

# ==========================================
# МОДУЛЬ 8: РАЗМИНКИ
# ==========================================
elif menu_choice in ["⚡ Разминки и интерактивы", "⚡ Сергіту сәттері мен интерактив"]:
    st.title("⚡ " + menu_choice)
    st.markdown("#### " + t["warm_sub"])
    st.divider()
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        top = st.text_input(t["warm_top"], "Алгоритмы")
    with col_w2:
        tm = st.slider(t["warm_time"], 2, 10, 5)

    if st.button("🚀 " + t["warm_btn"], type="primary", use_container_width=True):
        if not active_key: st.warning(t["no_key"])
        else:
            with st.spinner("⏳ " + t["warm_wait"]):
                genai.configure(api_key=active_key)
                model = genai.GenerativeModel("gemini-3.6-flash")
                prompt = f"{t['ai_lang_prompt']} Предложи 3 разминки на тему {top} на {tm} минут."
                res = model.generate_content(prompt)
                st.markdown(res.text)

# ==========================================
# 📊 GOOGLE ANALYTICS СЧЕТЧИК
# ==========================================
GA_TRACKING_ID = "G-0EW4TYEDKE" 
ga_component = f"""
<!DOCTYPE html>
<html>
<head>
    <script async src="https://www.googletagmanager.com/gtag/js?id={GA_TRACKING_ID}"></script>
    <script>
      window.dataLayer = window.dataLayer || [];
      function gtag(){{dataLayer.push(arguments);}}
      gtag('js', new Date());
      gtag('config', '{GA_TRACKING_ID}');
    </script>
</head>
<body>
</body>
</html>
"""
components.html(ga_component, height=0, width=0)
