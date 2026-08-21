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

# ==========================================
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

DEFAULT_API_KEY = "AIzaSy..."

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
    st.markdown("#### Автоматическая генерация индивидуальных вариантов в Word!")
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
                except:
                    pass
    else:
        uploaded_excel = st.file_uploader("📂 Загрузите Excel-файл (листы: Банк_вопросов, Ученики):", type=["xlsx"])
        if uploaded_excel:
            try:
                xls = pd.ExcelFile(uploaded_excel)
                df_questions = pd.read_excel(xls, 'Банк_вопросов')
                df_students = pd.read_excel(xls, 'Ученики')
                st.success("Excel-файл успешно прочитан!")
            except Exception as e:
                st.error(f"Ошибка чтения Excel: {e}")

    if df_questions is not None and df_students is not None:
        st.markdown("#### ⚙️ Настройка теста")
        col1, col2, col3 = st.columns(3)
        with col1: count_easy = st.number_input("🟢 Легких вопросов:", min_value=0, max_value=5, value=1)
        with col2: count_med = st.number_input("🟡 Средних вопросов:", min_value=0, max_value=5, value=1)
        with col3: count_hard = st.number_input("🔴 Сложных вопросов:", min_value=0, max_value=5, value=1)

        st.markdown("<br>", unsafe_allow_html=True) 

        if st.button("🚀 Сгенерировать варианты в Word", type="primary", use_container_width=True):
            try:
                with st.spinner("Генерация документов..."):
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

            except Exception as e:
                st.error(f"Ошибка обработки: {e}")


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

    st.markdown("##### Недели в четвертях")
    quarters_weeks = {}
    col_q = st.columns(quarters_count)
    for q in range(1, quarters_count + 1):
        with col_q[q - 1]:
            quarters_weeks[q] = st.number_input(f"{q}-я четверть:", 1, 15, 8)
            
    total_all_lessons = sum(q_w * hours_per_week for q_w in quarters_weeks.values())
    st.info(f"💡 Всего уроков: **{total_all_lessons}**")

    st.markdown("#### 📖 Источник учебного материала")
    source_type = st.radio(
        "Как передать содержание учебника?",
        ["Загрузить PDF-файл (учебник / оглавление)", "Ввести темы вручную текстом"],
        horizontal=True
    )

    uploaded_pdf = None
    textbook_content = ""

    if source_type == "Загрузить PDF-файл (учебник / оглавление)":
        uploaded_pdf = st.file_uploader("📂 Загрузите PDF-документ:", type=["pdf"])
    else:
        textbook_content = st.text_area("📝 Введите список тем или оглавление:", "1. Инфо и свойства 2. Двоичная система", height=150)

    if st.button("🚀 Сгенерировать КТП", type="primary", use_container_width=True):
        if not active_key or active_key == "AIzaSy...": 
            st.warning("Пожалуйста, введите ваш рабочий Gemini API Key в боковом меню слева!")
        elif source_type == "Загрузить PDF-файл (учебник / оглавление)" and not uploaded_pdf:
            st.error("Пожалуйста, загрузите PDF-файл!")
        elif source_type == "Ввести темы вручную текстом" and not textbook_content.strip():
            st.error("Пожалуйста, введите текст тем!")
        else:
            try:
                with st.spinner("ИИ анализирует материалы и распределяет цели обучения..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    
                    quarters_info_text = ", ".join([f"Четверть {q}: {w*hours_per_week} уроков" for q, w in quarters_weeks.items()])

                    prompt = f"""
                    Ты профессиональный методист. Составь КТП по предмету "{subject}" ({grade} класс).
                    Всего четвертей: {quarters_count}. Распределение: {quarters_info_text}. Всего уроков: {total_all_lessons}.
                    
                    {f'Темы из текста: {textbook_content}' if source_type == 'Ввести темы вручную текстом' else 'Изучи прикрепленный PDF-файл и вытяни темы оттуда.'}

                    ТРЕБОВАНИЯ К ЦЕЛЯМ ОБУЧЕНИЯ (ЦО):
                    - Каждый урок должен иметь 1-2 цели.
                    - Нумерация цели должна начинаться с номера четверти (Пример для 1 четверти: {grade}.1.1.1, для 2 четверти: {grade}.2.1.1).

                    Верни СТРОГО JSON-массив из объектов (без ```json):
                    [
                      {{
                        "quarter": 1,
                        "lesson_num": 1,
                        "topic": "Тема урока",
                        "targets": "{grade}.1.1.1 Цель обучения",
                        "homework": "Параграф 1"
                      }}
                    ]
                    """

                    if uploaded_pdf is not None:
                        pdf_bytes = uploaded_pdf.read()
                        pdf_part = {"mime_type": "application/pdf", "data": pdf_bytes}
                        response = model.generate_content([prompt, pdf_part])
                    else:
                        response = model.generate_content(prompt)

                    match = re.search(r'\[.*\]', response.text, re.DOTALL)
                    clean_json = match.group(0) if match else response.text.replace("```json", "").replace("```", "")
                    ktp_data = json.loads(clean_json)

                    df_ktp = pd.DataFrame(ktp_data)
                    df_ktp.columns = ["Четверть", "№", "Тема", "Цели (ЦО)", "Д/З"]
                    st.dataframe(df_ktp, use_container_width=True)

                    bio_excel = io.BytesIO()
                    with pd.ExcelWriter(bio_excel, engine="openpyxl") as writer: df_ktp.to_excel(writer, index=False)
                    bio_excel.seek(0)
                    st.download_button("📊 Скачать КТП (Excel)", bio_excel, f"КТП_{subject}_{grade}класс.xlsx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка ИИ: {e}")


# ==========================================
# МОДУЛЬ 3: AI-КОНСТРУКТОР КСП
# ==========================================
elif menu_choice == "📋 AI-Конструктор КСП":
    st.title("📋 AI-Конструктор КСП")
    st.markdown("#### Получите идеальный поурочный план в формате Word")
    st.divider()

    col_k1, col_k2 = st.columns(2)
    with col_k1:
        teacher_name = st.text_input("👤 ФИО учителя:", "Иванов И.И.")
        subject_ksp = st.text_input("📚 Предмет:", "Информатика")
        grade_ksp = st.number_input("🏫 Класс:", 1, 11, 8)
    with col_k2:
        topic_ksp = st.text_input("📝 Тема урока:", "Двоичная система счисления")
        target_ksp = st.text_input("🎯 Цель обучения (ЦО):", "8.1.1.1 Осуществлять перевод чисел")

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 Сгенерировать красивый КСП", type="primary", use_container_width=True):
        if not active_key or active_key == "AIzaSy...": st.warning("Пожалуйста, введите ваш рабочий Gemini API Key в боковом меню слева!")
        else:
            try:
                with st.spinner("Создаю структурированную таблицу урока..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")

                    prompt = f"""
                    Ты строгий методист. Создай план урока. Предмет: {subject_ksp}. Класс: {grade_ksp}. Тема: {topic_ksp}. ЦО: {target_ksp}.
                    КАТЕГОРИЧЕСКИ ЗАПРЕЩАЕТСЯ использовать Markdown (никаких звездочек **, решеток ###, жирного текста).
                    ВЕРНИ ОТВЕТ СТРОГО В ФОРМАТЕ JSON:
                    {{
                      "lesson_targets": "Все: ... Большинство: ... Некоторые: ...",
                      "eval_criteria": "Знает... Умеет... Применяет...",
                      "stages": [
                        {{ "time": "Начало урока (7 мин)", "teacher": "Орг. момент...", "student": "Слушают...", "eval": "Формативное", "resources": "Слайд 1" }},
                        {{ "time": "Середина урока (30 мин)", "teacher": "Объяснение...", "student": "Решают...", "eval": "Взаимооценивание", "resources": "Карточки" }},
                        {{ "time": "Конец урока (8 мин)", "teacher": "Рефлексия...", "student": "Заполняют стикеры...", "eval": "Самооценивание", "resources": "Стикеры" }}
                      ]
                    }}
                    """
                    response = model.generate_content(prompt)
                    match = re.search(r'\{.*\}', response.text, re.DOTALL)
                    clean_json = match.group(0) if match else response.text.replace("```json", "").replace("```", "")
                    ksp_data = json.loads(clean_json)

                    doc_ksp = Document()
                    section = doc_ksp.sections[-1]
                    section.orientation = WD_ORIENT.LANDSCAPE
                    section.page_width, section.page_height = section.page_height, section.page_width
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.5)
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)

                    title_p = doc_ksp.add_paragraph()
                    run_title = title_p.add_run("КРАТКОСРОЧНЫЙ ПЛАН УРОКА (КСП)")
                    run_title.font.bold = True
                    run_title.font.size = Pt(14)
                    run_title.font.name = 'Times New Roman'
                    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    passport_table = doc_ksp.add_table(rows=7, cols=2)
                    passport_table.style = 'Table Grid'
                    info_rows = [
                        ("ФИО учителя:", teacher_name), ("Предмет:", subject_ksp), ("Класс:", str(grade_ksp)),
                        ("Тема урока:", topic_ksp), ("Цель обучения (ЦО):", target_ksp),
                        ("Цели урока:", ksp_data.get("lesson_targets", "")), ("Критерии оценивания:", ksp_data.get("eval_criteria", ""))
                    ]

                    for idx, (label, val) in enumerate(info_rows):
                        passport_table.rows[idx].cells[0].text = label
                        passport_table.rows[idx].cells[0].paragraphs[0].runs[0].font.name = 'Times New Roman'
                        passport_table.rows[idx].cells[1].text = str(val)
                        if passport_table.rows[idx].cells[1].paragraphs[0].runs:
                            passport_table.rows[idx].cells[1].paragraphs[0].runs[0].font.name = 'Times New Roman'

                    doc_ksp.add_paragraph() 
                    stages_table = doc_ksp.add_table(rows=1, cols=5)
                    stages_table.style = 'Table Grid'
                    headers = ["Этап урока / Время", "Действия педагога", "Действия ученика", "Оценивание", "Ресурсы"]
                    for i, h_text in enumerate(headers):
                        stages_table.rows[0].cells[i].text = h_text
                        stages_table.rows[0].cells[i].paragraphs[0].runs[0].font.bold = True
                        stages_table.rows[0].cells[i].paragraphs[0].runs[0].font.name = 'Times New Roman'

                    for stage in ksp_data.get("stages", []):
                        row_cells = stages_table.add_row().cells
                        row_cells[0].text, row_cells[1].text, row_cells[2].text, row_cells[3].text, row_cells[4].text = (
                            str(stage.get("time", "")), str(stage.get("teacher", "")), str(stage.get("student", "")), str(stage.get("eval", "")), str(stage.get("resources", ""))
                        )
                        for cell in row_cells:
                            for p in cell.paragraphs:
                                for run in p.runs:
                                    run.font.name = 'Times New Roman'
                                    run.font.size = Pt(11)

                    bio_ksp = io.BytesIO()
                    doc_ksp.save(bio_ksp)
                    bio_ksp.seek(0)
                    st.success("✅ КСП успешно сгенерирован!")
                    st.download_button("📄 Скачать КСП (Word)", bio_ksp, f"КСП_{grade_ksp}класс.docx", use_container_width=True)

            except Exception as e: st.error(f"Ошибка: {e}")


# ==========================================
# МОДУЛЬ 4: AI-ПРОВЕРКА РАБОТ ПО ФОТО
# ==========================================
elif menu_choice == "📷 AI-Проверка по фото":
    st.title("📷 AI-Проверка письменных работ")
    st.divider()

    col_img1, col_img2 = st.columns([1, 1])
    with col_img1:
        st.markdown("#### Шаг 1. Ввод данных")
        uploaded_image = st.file_uploader("📂 Загрузите фото:", type=["jpg", "jpeg", "png"])
        subject_check = st.selectbox("📚 Предмет:", ["Информатика", "Математика / Алгебра", "Химия", "Физика", "Другой"])
        teacher_notes = st.text_input("💡 Правильный ответ (необязательно):", placeholder="x^2 - 4 = 0")
        
        if uploaded_image:
            img = Image.open(uploaded_image)
            st.image(img, use_container_width=True, caption="Предпросмотр")

    with col_img2:
        st.markdown("#### Шаг 2. Анализ")
        if uploaded_image and st.button("🔍 Проверить работу", type="primary", use_container_width=True):
            if not active_key or active_key == "AIzaSy...": st.warning("Пожалуйста, введите ваш рабочий Gemini API Key в боковом меню слева!")
            else:
                try:
                    with st.spinner("ИИ распознает и проверяет..."):
                        genai.configure(api_key=active_key)
                        model = genai.GenerativeModel("gemini-3.6-flash")
                        prompt = f"""
                        Ты учитель. Предмет: {subject_check}. Изучи фото. Условие: {teacher_notes}.
                        Разбор: 1. Распознанный текст 2. Ход решения 3. Ошибки 4. Итоговая оценка.
                        """
                        response = model.generate_content([prompt, img])
                        st.info("Результат готов!")
                        st.markdown(response.text)
                except Exception as e: st.error(f"Ошибка: {e}")


# ==========================================
# МОДУЛЬ 5: ГЕНЕРАТОР ХАРАКТЕРИСТИК
# ==========================================
elif menu_choice == "👤 Генератор характеристик":
    st.title("👤 Генератор характеристик учеников")
    st.divider()

    col_ch1, col_ch2 = st.columns(2)
    with col_ch1:
        student_name = st.text_input("👤 ФИО ученика:", "Иванов Иван")
        student_grade = st.number_input("🏫 Класс:", 1, 11, 8)
        academic_perf = st.selectbox("📊 Успеваемость:", ["Отличник", "Ударник", "Средняя", "Низкая"])
    with col_ch2:
        behavior = st.selectbox("🎭 Поведение:", ["Примерное", "Хорошее", "Нарушает", "Требует внимания"])
        activity = st.multiselect("🌟 Активность:", ["Активист", "Олимпиады", "Спорт", "Творчество"], default=["Активист"])
    
    extra_info = st.text_area("📝 Доп. заметки:", placeholder="Любит математику...")

    if st.button("🚀 Создать характеристику", type="primary", use_container_width=True):
        if not active_key or active_key == "AIzaSy...": st.warning("Пожалуйста, введите ваш рабочий Gemini API Key в боковом меню слева!")
        else:
            try:
                with st.spinner("ИИ пишет текст..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"Ты классный руководитель. Напиши характеристику. Ученик: {student_name}, класс: {student_grade}, успеваемость: {academic_perf}, поведение: {behavior}, активность: {', '.join(activity)}. Заметки: {extra_info}."
                    response = model.generate_content(prompt)
                    
                    st.success("Характеристика готова!")
                    with st.expander("Посмотреть текст", expanded=True):
                        st.markdown(response.text)

                    doc_char = Document()
                    doc_char.add_heading(f"Характеристика: {student_name}", level=1)
                    doc_char.add_paragraph(response.text)
                    bio_char = io.BytesIO()
                    doc_char.save(bio_char)
                    bio_char.seek(0)
                    st.download_button("📄 Скачать (Word)", bio_char, f"Характеристика_{student_name}.docx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка: {e}")


# ==========================================
# МОДУЛЬ 6: РАЗМИНКИ И ИНТЕРАКТИВЫ
# ==========================================
elif menu_choice == "⚡ Разминки и интерактивы":
    st.title("⚡ AI-Разминки на урок")
    st.divider()

    col_w1, col_w2 = st.columns(2)
    with col_w1:
        warm_subject = st.text_input("📚 Предмет:", "Информатика")
        warm_topic = st.text_input("📝 Тема:", "Алгоритмы")
    with col_w2:
        warm_type = st.selectbox("🎭 Формат:", ["Викторина", "Загадка", "Двигательная", "True/False"])
        duration = st.slider("⏳ Время (мин):", 2, 10, 5)

    if st.button("🚀 Подобрать", type="primary", use_container_width=True):
        if not active_key or active_key == "AIzaSy...": st.warning("Пожалуйста, введите ваш рабочий Gemini API Key в боковом меню слева!")
        else:
            try:
                with st.spinner("ИИ генерирует идеи..."):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"Предмет: {warm_subject}, Тема: {warm_topic}, Формат: {warm_type}, Время: {duration} мин. Предложи 3 интерактива."
                    response = model.generate_content(prompt)
                    
                    st.success("Идеи готовы!")
                    st.markdown(response.text)
            except Exception as e: st.error(f"Ошибка: {e}")
