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
# 0. НАСТРОЙКИ АНАЛИТИКИ И ДИЗАЙНА
# ==========================================
def track_event(action, category, label):
    components.html(f"""
        <script>
            gtag('event', '{action}', {{
                'event_category': '{category}',
                'event_label': '{label}'
            }});
        </script>
    """, height=0, width=0)
# ==========================================
# Google Analytics (замените G-XXXXXXXXXX на ваш ID!)
GA_ID = "G-G-0EW4TYEDKE" 
components.html(f"""
    <script async src="https://www.googletagmanager.com/gtag/js?id={GA_ID}"></script>
    <script>
        window.dataLayer = window.dataLayer || [];
        function gtag(){{dataLayer.push(arguments);}}
        gtag('js', new Date());
        gtag('config', '{GA_ID}');
    </script>
""", height=0, width=0)
# 1. НАСТРОЙКИ ДИЗАЙНА (CSS)
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
# 2. БОКОВОЕ МЕНЮ И НАСТРОЙКИ
# ==========================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/1972/1972413.png", width=80)
st.sidebar.title("🎓 AI-Помощник")
st.sidebar.markdown("---")

st.sidebar.subheader("🔑 Доступ к ИИ")
global_api_key = st.sidebar.text_input("Gemini API Key:", type="password", help="Введите ключ один раз для всех инструментов")

with st.sidebar.expander("ℹ️ Как получить API ключ бесплатно?"):
    st.markdown("""
    1. Зайдите на [Google AI Studio](https://aistudio.google.com/app/apikey).
    2. Войдите через Google-аккаунт.
    3. Нажмите синюю кнопку **Create API key**.
    4. Скопируйте ключ и вставьте в поле выше.
    """)

active_key = global_api_key if global_api_key else DEFAULT_API_KEY

st.sidebar.markdown("---")
st.sidebar.subheader("🛠️ Инструменты")
menu_choice = st.sidebar.radio(
    "Навигация:",
    [
        "📝 Генератор карточек",
        "📅 AI-Генератор КТП",
        "📋 AI-Конструктор КСП",
        "📊 Анализ и визуализация (EDA)",
        "🤖 ML-Прогноз уровня ученика",
        "📷 AI-Проверка по фото",
        "👤 Генератор характеристик",
        "⚡ Разминки и интерактивы",
    ],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.caption("✨ Создано для учителей с ❤️")


# ==========================================
# МОДУЛЬ 1: ГЕНЕРАТОР КАРТОЧЕК
# ==========================================
if menu_choice == "📝 Генератор карточек":
    st.title("📝 Генератор карточек с заданиями")
    st.markdown("#### Автоматическая генерация индивидуальных вариантов в Word")
    st.divider()

    source_type = st.radio("Источник данных:", ["Google Таблица (ссылка)", "Excel-файл (.xlsx)"], horizontal=True)
    df_questions, df_students = None, None

    if source_type == "Google Таблица (ссылка)":
        default_url = "https://docs.google.com/spreadsheets/d/1fJKlRP7YY3r6DFjd_PuLXFIKkg3GdSRAM9Rxwq502e8/edit?usp=sharing"
        sheet_url = st.text_input("🔗 Вставьте ссылку на Google Таблицу:", value=default_url)
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
        uploaded_excel = st.file_uploader("📂 Загрузите Excel-файл (листы: Банк_вопросов, Ученики):", type=["xlsx"])
        if uploaded_excel:
            try:
                xls = pd.ExcelFile(uploaded_excel)
                df_questions = pd.read_excel(xls, 'Банк_вопросов')
                df_students = pd.read_excel(xls, 'Ученики')
                st.success("Excel-файл успешно прочитан!")
            except Exception as e: st.error(f"Ошибка чтения Excel: {e}")

    if df_questions is not None and df_students is not None:
        st.markdown("#### ⚙️ Настройка теста")
        col1, col2, col3 = st.columns(3)
        with col1: count_easy = st.number_input("🟢 Легких вопросов:", min_value=0, max_value=5, value=1)
        with col2: count_med = st.number_input("🟡 Средних вопросов:", min_value=0, max_value=5, value=1)
        with col3: count_hard = st.number_input("🔴 Сложных вопросов:", min_value=0, max_value=5, value=1)

        if st.button("🚀 Сгенерировать варианты в Word", type="primary", use_container_width=True):
            track_event('generate_cards', 'Module_Action', 'Cards_Generator')
            if not active_key: st.warning("Пожалуйста, введите API Key в боковом меню!")
            else:
                try:
                    with st.spinner("⏳ ИИ обрабатывает банк вопросов и формирует индивидуальные варианты..."):
                        df_questions.columns = df_questions.columns.astype(str).str.strip().str.lower()
                        df_students.columns = df_students.columns.astype(str).str.strip().str.lower()
                        rename_dict = {}
                        for col in df_questions.columns:
                            if "сложн" in col: rename_dict[col] = "сложность"
                            elif "тем" in col: rename_dict[col] = "тема"
                            elif "вопрос" in col: rename_dict[col] = "вопрос"
                            elif "ответ" in col: rename_dict[col] = "ответ"
                        df_questions = df_questions.rename(columns=rename_dict)
                        student_col = [c for c in df_students.columns if "фио" in c]
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
                        p_info.add_run(f"Ученик(ца): {student}").bold = True
                        for idx, row in student_variant.iterrows():
                            p_q = doc_students.add_paragraph()
                            p_q.add_run(f"Задание {idx + 1} ").bold = True
                            p_q.add_run(f"{row['вопрос']}\n")
                            p_q.add_run("Ответ: ____________________")
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

                    st.success("🎉 Документы успешно созданы!")
                    col_d1, col_d2 = st.columns(2)
                    with col_d1: st.download_button("📄 Скачать Карточки (Word)", bio_students, "Карточки.docx", use_container_width=True)
                    with col_d2: st.download_button("🔑 Скачать Ключи (Word)", bio_teacher, "Ключи.docx", use_container_width=True)
                except Exception as e: st.error(f"Ошибка обработки: {e}")

# ==========================================
# МОДУЛЬ 2: AI-ГЕНЕРАТОР КТП
# ==========================================
elif menu_choice == "📅 AI-Генератор КТП":
    st.title("📅 AI-Генератор КТП")
    st.markdown("#### Сформируйте КТП из PDF-учебника или текста с выгрузкой в Excel")
    st.divider()
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        subject = st.text_input("📚 Предмет:", "Информатика")
        grade = st.number_input("🏫 Класс:", 1, 11, 8)
    with col_p2:
        quarters_count = st.selectbox("📅 Количество четвертей:", [1, 2, 3, 4], index=0)
        hours_per_week = st.number_input("⏰ Часов в неделю:", 1, 5, 2)
    quarters_weeks = {q: st.number_input(f"{q}-я четверть (недель):", 1, 15, 8) for q in range(1, quarters_count + 1)}
    total_all_lessons = sum(q_w * hours_per_week for q_w in quarters_weeks.values())
    st.info(f"💡 Всего уроков: **{total_all_lessons}**")

    source_type = st.radio("Источник:", ["Загрузить PDF-файл", "Ввести темы текстом"], horizontal=True)
    uploaded_pdf, textbook_content = None, ""
    if source_type == "Загрузить PDF-файл": uploaded_pdf = st.file_uploader("📂 PDF:", type=["pdf"])
    else: textbook_content = st.text_area("📝 Темы:", "1. Инфо 2. Двоичная система", height=100)

    if st.button("🚀 Сгенерировать КТП", type="primary", use_container_width=True):
        track_event('generate_ktp', 'Module_Action', 'KTP_Generator')
        if not active_key: st.warning("Пожалуйста, введите API Key в боковом меню!")
        else:
            try:
                with st.spinner("⏳ ИИ анализирует учебные материалы и формирует календарно-тематическое планирование..."):
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
elif menu_choice == "📋 AI-Конструктор КСП":
    st.title("📋 AI-Конструктор КСП")
    st.markdown("#### Поурочный план в формате Word")
    st.divider()
    col_k1, col_k2 = st.columns(2)
    with col_k1:
        teacher_name = st.text_input("ФИО учителя:", "Иванов И.И.")
        subject_ksp = st.text_input("Предмет:", "Информатика")
        grade_ksp = st.number_input("Класс:", 1, 11, 8)
    with col_k2:
        topic_ksp = st.text_input("Тема урока:", "Двоичная система счисления")
        target_ksp = st.text_input("ЦО:", "8.1.1.1 Перевод чисел")

    if st.button("🚀 Сгенерировать КСП", type="primary", use_container_width=True):
        track_event('generate_ksp', 'Module_Action', 'KSP_Generator')
        if not active_key: st.warning("Пожалуйста, введите API Key в боковом меню!")
        else:
            try:
                with st.spinner("⏳ ИИ методист разрабатывает этапы урока, дескрипторы и дифференциацию..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"Создай план урока по предмету {subject_ksp}, тема {topic_ksp}. Верни строго JSON без markdown: {{\"lesson_targets\":\"...\", \"eval_criteria\":\"...\", \"stages\":[{{\"time\":\"Начало\", \"teacher\":\"...\", \"student\":\"...\", \"eval\":\"...\", \"resources\":\"...\"}}]}}"
                    response = model.generate_content(prompt)
                    match = re.search(r'\{.*\}', response.text, re.DOTALL)
                    ksp_data = json.loads(match.group(0) if match else response.text)

                    doc_ksp = Document()
                    section = doc_ksp.sections[-1]
                    section.orientation = WD_ORIENT.LANDSCAPE
                    section.page_width, section.page_height = section.page_height, section.page_width
                    
                    title = doc_ksp.add_paragraph()
                    r = title.add_run("КРАТКОСРОЧНЫЙ ПЛАН УРОКА (КСП)")
                    r.font.bold = True
                    r.font.size = Pt(14)
                    r.font.name = 'Times New Roman'
                    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    t_table = doc_ksp.add_table(rows=7, cols=2)
                    t_table.style = 'Table Grid'
                    info = [("ФИО учителя:", teacher_name), ("Предмет:", subject_ksp), ("Класс:", str(grade_ksp)), ("Тема:", topic_ksp), ("ЦО:", target_ksp), ("Цели урока:", ksp_data.get("lesson_targets","")), ("Критерии:", ksp_data.get("eval_criteria",""))]
                    for idx, (l, v) in enumerate(info):
                        t_table.rows[idx].cells[0].text = l
                        t_table.rows[idx].cells[1].text = str(v)

                    doc_ksp.add_paragraph()
                    s_table = doc_ksp.add_table(rows=1, cols=5)
                    s_table.style = 'Table Grid'
                    headers = ["Этап / Время", "Действия учителя", "Действия ученика", "Оценивание", "Ресурсы"]
                    for i, h in enumerate(headers): s_table.rows[0].cells[i].text = h

                    for stg in ksp_data.get("stages", []):
                        row = s_table.add_row().cells
                        row[0].text, row[1].text, row[2].text, row[3].text, row[4].text = stg.get("time",""), stg.get("teacher",""), stg.get("student",""), stg.get("eval",""), stg.get("resources","")

                    bio = io.BytesIO()
                    doc_ksp.save(bio)
                    bio.seek(0)
                    st.success("Готово!")
                    st.download_button("📄 Скачать КСП (Word)", bio, f"КСП_{topic_ksp}.docx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 4: АНАЛИЗ И ВИЗУАЛИЗАЦИЯ ДАННЫХ (EDA)
# ==========================================
elif menu_choice == "📊 Анализ и визуализация (EDA)":
    st.title("📊 Анализ и визуализация успеваемости класса")
    st.markdown("#### Загрузите Excel-файл с оценками учеников для описательной статистики и графиков")
    st.divider()

    uploaded_eda = st.file_uploader("📂 Загрузите файл с оценками (.xlsx)", type=["xlsx"])
    if uploaded_eda:
        df_eda = pd.read_excel(uploaded_eda)
        st.write("📋 Предпросмотр данных:", df_eda.head())

        numeric_cols = df_eda.select_dtypes(include=['number']).columns.tolist()
        if numeric_cols:
            selected_col = st.selectbox("Выберите числовой показатель для анализа:", numeric_cols)
            
            if st.button("📈 Построить графики и статистику", type="primary", use_container_width=True):
                track_event('run_eda', 'Module_Action', 'EDA_Analysis')
                with st.spinner("⏳ Выполняется расчет статистических метрик и генерация графиков..."):
                    st.subheader("📌 Описательная статистика")
                    st.write(df_eda[selected_col].describe())

                    col_g1, col_g2 = st.columns(2)
                    
                    with col_g1:
                        st.markdown("##### Гистограмма распределения")
                        fig, ax = plt.subplots(figsize=(6, 4))
                        sns.histplot(df_eda[selected_col], kde=True, ax=ax, color='purple')
                        st.pyplot(fig)

                    with col_g2:
                        st.markdown("##### Ящик с усами (Boxplot)")
                        fig, ax = plt.subplots(figsize=(6, 4))
                        sns.boxplot(y=df_eda[selected_col], ax=ax, color='skyblue')
                        st.pyplot(fig)
        else:
            st.warning("В файле не найдено числовых колонок для построения графиков.")

# ==========================================
# МОДУЛЬ 5: ML-ПРОГНОЗ УРОВНЯ УЧЕНИКА
# ==========================================
elif menu_choice == "🤖 ML-Прогноз уровня ученика":
    st.title("🤖 Машинное обучение в образовании (Scikit-learn)")
    st.markdown("#### Прогнозирование группы поддержки / продвинутого уровня с помощью Random Forest")
    st.divider()

    st.write("Введите показатели ученика, чтобы модель машинного обучения определила его уровень:")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        att = st.slider("Посещаемость уроков (%):", 50, 100, 85)
        hw = st.slider("Выполнение домашних заданий (%):", 0, 100, 75)
    with col_m2:
        test_score = st.slider("Средний балл по тестам:", 0, 100, 80)
        activity = st.selectbox("Классная активность:", ["Низкая", "Средняя", "Высокая"])
        act_val = 1 if activity == "Низкая" else (2 if activity == "Средняя" else 3)

    if st.button("🔮 Предсказать уровень ученика", type="primary", use_container_width=True):
        track_event('ml_predict', 'Module_Action', 'ML_Prediction')
        with st.spinner("⏳ ML-модель анализирует показатели и выстраивает классификацию..."):
            X_train = [[60, 50, 55, 1], [90, 85, 88, 3], [70, 60, 65, 2], [95, 95, 92, 3], [55, 40, 45, 1], [85, 80, 82, 2]]
            y_train = ["Группа поддержки", "Продвинутый", "Стандартный", "Продвинутый", "Группа поддержки", "Стандартный"]

            model = RandomForestClassifier(random_state=42)
            model.fit(X_train, y_train)

            prediction = model.predict([[att, hw, test_score, act_val]])[0]

        st.success(f"🎯 Рекомендация ML-модели: **{prediction}**")
        if prediction == "Группа поддержки":
            st.info("💡 Рекомендация преподавателю: Рекомендуется выдать дифференцированные карточки Уровня А и уделить внимание базовым темам.")
        elif prediction == "Стандартный":
            st.info("💡 Рекомендация преподавателю: Ученик стабильно усваивает материал. Подходят стандартные задания Уровня В.")
        else:
            st.info("💡 Рекомендация преподавателю: Ученик опережает программу. Предложите задачи повышенной сложности (Уровень С).")

# ==========================================
# МОДУЛЬ 6: AI-ПРОВЕРКА ПО ФОТО
# ==========================================
elif menu_choice == "📷 AI-Проверка по фото":
    st.title("📷 AI-Проверка письменных работ")
    st.divider()
    uploaded_image = st.file_uploader("📂 Фото работы:", type=["jpg", "png"])
    subject_check = st.selectbox("Предмет:", ["Информатика", "Математика", "Химия"])
    if uploaded_image and st.button("Проверить", type="primary"):
        track_event('check_photo', 'Module_Action', 'Photo_Correction')
        if not active_key: st.warning("Пожалуйста, введите API Key в боковом меню!")
        else:
            try:
                with st.spinner("⏳ Мультимодальный ИИ распознает почерк и проверяет ход решения..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    res = model.generate_content(["Проверь работу ученика, найди ошибки и дай оценку:", Image.open(uploaded_image)])
                st.markdown(res.text)
            except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 7: ГЕНЕРАТОР ХАРАКТЕРИСТИК (ОБНОВЛЕННЫЙ)
# ==========================================
elif menu_choice == "👤 Генератор характеристик":
    st.title("👤 Генератор педагогических характеристик")
    st.markdown("#### Составление развернутого отчета на основе ключевых параметров")
    st.divider()

    col_h1, col_h2 = st.columns(2)
    with col_h1:
        s_name = st.text_input("👤 ФИО ученика:", "Иванов Иван")
        s_class = st.text_input("🏫 Класс:", "8 «А»")
    with col_h2:
        s_traits = st.text_area("🔑 Ключевые особенности и достижения:", placeholder="Пример: Ответственный, увлекается программированием, староста класса, иногда пропускает тренировки...")

    if st.button("🚀 Сгенерировать характеристику", type="primary", use_container_width=True):
        track_event('generate_char', 'Module_Action', 'Char_Generator')
        if not active_key: st.warning("Пожалуйста, введите API Key в боковом меню!")
        else:
            try:
                with st.spinner("⏳ ИИ классный руководитель формирует текст педагогической характеристики..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"Составь официальную развернутую педагогическую характеристику на ученика {s_name}, обучающегося в {s_class} классе. Учти следующие ключевые особенности и качества: {s_traits}. Напиши в официально-деловом стиле (3-4 абзаца)."
                    res = model.generate_content(prompt)
                
                st.success("Характеристика готова!")
                st.markdown(res.text)
            except Exception as e: st.error(f"Ошибка: {e}")

# ==========================================
# МОДУЛЬ 8: РАЗМИНКИ И ИНТЕРАКТИВЫ (ОБНОВЛЕННЫЙ)
# ==========================================
elif menu_choice == "⚡ Разминки и интерактивы":
    st.title("⚡ AI-Генератор разминок и Icebreakers")
    st.markdown("#### Интерактивы и разминки для начала урока с учетом тайминга")
    st.divider()

    col_w1, col_w2 = st.columns(2)
    with col_w1:
        w_top = st.text_input("📝 Тема урока:", "Алгоритмы и ветвления")
    with col_w2:
        w_time = st.slider("⏳ Время на разминку (минут):", 2, 10, 5)

    if st.button("🚀 Подобрать разминку", type="primary", use_container_width=True):
        track_event('generate_warmup', 'Module_Action', 'Warmup_Generator')
        if not active_key: st.warning("Пожалуйста, введите API Key в боковом меню!")
        else:
            try:
                with st.spinner(f"⏳ ИИ подбирает креативные интерактивы ровно на {w_time} минут..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"Предложи 3 интересных варианта разминки или icebreaker на начало урока по теме: '{w_top}'. Ограничение по времени: ровно {w_time} минут. Для каждого варианта укажи суть, правила и вопросы для класса."
                    res = model.generate_content(prompt)
                
                st.success("Идеи готовы!")
                st.markdown(res.text)
            except Exception as e: st.error(f"Ошибка: {e}")
