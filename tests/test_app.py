"""Offline Streamlit integration tests for both languages and all eight tools.

The Gemini boundary is mocked; uploaded files and generated documents stay in
memory. Network clients are blocked as an extra guard against accidental calls.
"""

from datetime import date
import io
import json
import os
from pathlib import Path
import unittest
from unittest.mock import patch
import zipfile

# Load scientific libraries before AppTest starts its per-rerun timeout.
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot  # noqa: F401
import pandas as pd
from PIL import Image
from docx import Document
from sklearn.ensemble import RandomForestClassifier  # noqa: F401
from streamlit.testing.v1 import AppTest

import ai_service
import bilim_core as core
from ui_helpers import ERRORS


APP = Path(__file__).resolve().parents[1] / "app.py"
TEST_KEY = "mock-integration-key-not-a-secret"
MODULES = ("cards", "ktp", "ksp", "eda", "ml", "photo", "characteristic", "warmup")


def document_text(raw):
    document = Document(io.BytesIO(raw))
    paragraphs = [paragraph.text for paragraph in document.paragraphs]
    paragraphs.extend(cell.text for table in document.tables for row in table.rows for cell in row.cells)
    return "\n".join(paragraphs)


def ksp_response(lang):
    kk = lang == "kk"
    return json.dumps({
        "section": "Алгоритмдер" if kk else "Алгоритмы",
        "lesson_targets": "Алгоритмді құрады және тексереді." if kk else "Составляет и проверяет алгоритм.",
        # The user-supplied target must take precedence over this model invention.
        "learning_targets": "UNTRUSTED MODEL TARGET",
        "stages": [{
            "time": time,
            "teacher": "Тапсырманы түсіндіреді." if kk else "Объясняет задание.",
            "student": "Шешімді тексереді." if kk else "Проверяет решение.",
            "resources": "Тапсырма парағы" if kk else "Лист заданий",
        } for time in (
            ("Басы — 5 минут", "Ортасы — 35 минут", "Соңы — 5 минут") if kk else
            ("Начало — 5 минут", "Основная часть — 35 минут", "Конец — 5 минут")
        )],
    }, ensure_ascii=False)


def ktp_response(quarter=1, lang="ru"):
    return json.dumps([{
        "quarter": quarter, "lesson_num": 1,
        "section": "Алгоритмдер" if lang == "kk" else "Алгоритмы",
        "topic": ("Алгоритм құру " if lang == "kk" else "Составление алгоритма ") + str(quarter),
        "targets": "Сызықтық алгоритм құру." if lang == "kk" else "Составлять линейный алгоритм.",
    }], ensure_ascii=False)


def demo_png():
    buffer = io.BytesIO()
    Image.new("RGB", (80, 60), color="white").save(buffer, format="PNG")
    return buffer.getvalue()


class AppIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.answer = "Проверенный учебный результат.\nОқу нәтижесі дайын."
        self.key_patch = patch.dict(os.environ, {
            "GEMINI_API_KEY": "", "GOOGLE_API_KEY": "", "DEFAULT_API_KEY": "",
            "GEMINI_MODEL": "",
        })
        self.key_patch.start()
        self.addCleanup(self.key_patch.stop)
        self.ai_patch = patch("ai_service.generate_text", side_effect=self.generate_answer)
        self.ai = self.ai_patch.start()
        self.addCleanup(self.ai_patch.stop)
        self.client_patch = patch("ai_service.genai.Client", side_effect=AssertionError("Network disabled in tests"))
        self.client = self.client_patch.start()
        self.addCleanup(self.client_patch.stop)
        self.url_patch = patch("urllib.request.urlopen", side_effect=AssertionError("Network disabled in tests"))
        self.urlopen = self.url_patch.start()
        self.addCleanup(self.url_patch.stop)

    def tearDown(self):
        self.client.assert_not_called()
        self.urlopen.assert_not_called()

    def generate_answer(self, prompt, api_key, **kwargs):
        if not api_key:
            raise ai_service.AIError("missing_key")
        if isinstance(self.answer, Exception):
            raise self.answer
        if isinstance(self.answer, list):
            return self.answer.pop(0)
        return self.answer

    def new_app(self, module="cards", lang="ru", with_key=False):
        app = AppTest.from_file(APP, default_timeout=45)
        # Never inherit any real Streamlit credentials from this workstation.
        for key in ("GEMINI_API_KEY", "GOOGLE_API_KEY", "GEMINI_MODEL"):
            app.secrets[key] = ""
        app.secrets["DEFAULT_API_KEY"] = TEST_KEY if with_key else ""
        app.run()
        self.assert_no_exception(app)
        if lang != "ru":
            app.selectbox("language").select(lang)
        if module != "cards":
            app.radio("module").set_value(module)
        if lang != "ru" or module != "cards":
            app.run()
        self.assert_no_exception(app)
        return app

    def assert_no_exception(self, app):
        self.assertEqual([item.message for item in app.exception], [])

    def click(self, app, module):
        app.button(f"{module}_generate").click().run()
        self.assert_no_exception(app)

    def assert_error(self, app, code, lang):
        expected = ERRORS[code][lang == "kk"]
        self.assertIn(expected, [item.value for item in app.error])
        self.assertEqual(len(app.success), 0)

    def assert_no_result(self, app, module, lang):
        self.assertNotIn(f"result_{module}_{lang}", app.session_state)

    def test_all_eight_modules_open_in_both_languages(self):
        app = self.new_app()
        for lang in ("ru", "kk"):
            app.selectbox("language").select(lang).run()
            for module in MODULES:
                with self.subTest(lang=lang, module=module):
                    app.radio("module").set_value(module).run()
                    self.assert_no_exception(app)
                    self.assertEqual(len(app.error), 0)
                    self.assertTrue(app.button(f"{module}_generate").label)
                    self.assertTrue(app.title[0].value)
        self.ai.assert_not_called()

    def test_demo_cards_work_without_key_and_downloads_persist(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app(lang=lang)
                self.click(app, "cards")
                self.assertEqual(len(app.error), 0)
                result = app.session_state[f"result_cards_{lang}"]
                self.assertEqual(len(result["variants"]), 3)
                self.assertTrue(all(len(variant["questions"]) == 3 for variant in result["variants"]))
                downloads = {button.key for button in app.download_button}
                self.assertTrue({"cards_download_cards", "cards_download_keys", "cards_download_zip"} <= downloads)
                with zipfile.ZipFile(io.BytesIO(result["zip"])) as archive:
                    self.assertEqual(set(archive.namelist()), {"cards.docx", "answers.docx"})
                    self.assertEqual(archive.read("cards.docx"), result["cards"])
                    self.assertEqual(archive.read("answers.docx"), result["keys"])
                self.assertIn("Оқушы" if lang == "kk" else "Ученик", document_text(result["cards"]))
                cards = result["cards"]
                app.download_button("cards_download_cards").click().run()
                self.assert_no_exception(app)
                app.run()
                self.assertEqual(app.session_state[f"result_cards_{lang}"]["cards"], cards)
                self.assertIn("cards_download_keys", {button.key for button in app.download_button})
                other_lang = "kk" if lang == "ru" else "ru"
                app.selectbox("language").select(other_lang).run()
                self.assertNotIn("cards_download_cards", {button.key for button in app.download_button})
                app.selectbox("language").select(lang).run()
                self.assertEqual(app.session_state[f"result_cards_{lang}"]["cards"], cards)
                app.radio("module").set_value("ml").run()
                app.radio("module").set_value("cards").run()
                self.assert_no_exception(app)
                self.assertEqual(app.session_state[f"result_cards_{lang}"]["cards"], cards)
                self.assertIn("cards_download_cards", {button.key for button in app.download_button})
        self.ai.assert_not_called()

    def test_zero_card_count_is_rejected_in_both_languages(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app(lang=lang)
                for level in core.LEVELS:
                    app.number_input(f"cards_{level}").set_value(0)
                self.click(app, "cards")
                self.assert_error(app, "counts_empty", lang)
                self.assert_no_result(app, "cards", lang)
        self.ai.assert_not_called()

    def test_ai_cards_respect_level_counts_and_language(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                self.answer = core.demo_bank(lang).to_json(orient="records", force_ascii=False)
                app = self.new_app(lang=lang, with_key=True)
                app.radio("cards_source").set_value("ai").run()
                app.number_input("cards_easy").set_value(0)
                app.number_input("cards_medium").set_value(2)
                app.number_input("cards_hard").set_value(1)
                app.number_input("cards_grade").set_value(7)
                self.click(app, "cards")
                self.assertEqual(len(app.error), 0)
                result = app.session_state[f"result_cards_{lang}"]
                for variant in result["variants"]:
                    levels = [item["difficulty"] for item in variant["questions"]]
                    self.assertEqual(levels.count("easy"), 0)
                    self.assertEqual(levels.count("medium"), 2)
                    self.assertEqual(levels.count("hard"), 1)
                call = self.ai.call_args
                self.assertIn("7 класса", call.args[0])
                self.assertIn("қазақ тілінде" if lang == "kk" else "русском языке", call.args[0])
                self.assertTrue(call.kwargs["json_mode"])

    def test_cards_excel_upload_works_without_ai(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app(lang=lang)
                app.radio("cards_source").set_value("excel").run()
                app.file_uploader("cards_upload").upload("demo.xlsx", core.make_template(lang),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").run()
                self.click(app, "cards")
                self.assertEqual(len(app.error), 0)
                self.assertTrue(app.session_state[f"result_cards_{lang}"]["cards"])
        self.ai.assert_not_called()

    def test_ksp_uses_selected_grade_target_and_exports_user_target(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                self.answer = ksp_response(lang)
                app = self.new_app("ksp", lang, with_key=True)
                target = "8.3.2.1 Циклдің жұмысын түсіндіру." if lang == "kk" else "8.3.2.1 Объяснять работу цикла."
                app.number_input("ksp_grade").set_value(8)
                app.text_area(f"ksp_target_{lang}").set_value(target)
                self.click(app, "ksp")
                self.assertEqual(len(app.error), 0)
                prompt = self.ai.call_args.args[0]
                self.assertIn('"grade": 8', prompt)
                self.assertIn(target, prompt)
                self.assertIn("қазақ тілінде" if lang == "kk" else "русском языке", prompt)
                result = app.session_state[f"result_ksp_{lang}"]
                exported = document_text(result["document"])
                self.assertIn(target, exported)
                self.assertIn("Информатика", exported)
                self.assertNotIn("UNTRUSTED MODEL TARGET", exported)
                self.assertIn("Қысқа мерзімді жоспар" if lang == "kk" else "Краткосрочный план", exported)
                self.assertEqual(len(result["data"]["stages"]), 3)
                self.assertIn("ksp_download", {button.key for button in app.download_button})
                app.run()
                self.assert_no_exception(app)
                self.assertEqual(app.session_state[f"result_ksp_{lang}"]["document"], result["document"])

    def test_ksp_rejects_empty_target_before_ai(self):
        app = self.new_app("ksp", with_key=True)
        app.text_area("ksp_target_ru").set_value("  ")
        self.click(app, "ksp")
        self.assert_error(app, "field_required", "ru")
        self.assert_no_result(app, "ksp", "ru")
        self.ai.assert_not_called()

    def test_ktp_default_quarters_and_year_export_correctly(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app("ktp", lang, with_key=True)
                self.assertEqual(app.selectbox("ktp_quarters").value, 4)
                today = date.today()
                start = today.year if today.month >= 8 else today.year - 1
                self.assertEqual(app.text_input("ktp_year").value, f"{start}–{start + 1}")
                selected_year = f"{start + 1}–{start + 2}"
                app.text_input("ktp_year").set_value(selected_year)
                for quarter in range(1, 5):
                    app.number_input(f"ktp_weeks_{quarter}").set_value(1)
                self.answer = [ktp_response(q, lang) for q in range(1, 5)]
                self.ai.reset_mock()
                self.click(app, "ktp")
                self.assertEqual(len(app.error), 0)
                self.assertEqual(self.ai.call_count, 4)
                result = app.session_state[f"result_ktp_{lang}"]
                self.assertEqual([row["lesson_num"] for row in result["data"]], [1, 2, 3, 4])
                self.assertEqual([row["quarter"] for row in result["data"]], [1, 2, 3, 4])
                exported = document_text(result["document"])
                self.assertIn(selected_year, exported)
                self.assertIn("Тоқсан" if lang == "kk" else "Четверть", exported)
                if lang == "kk":
                    self.assertNotIn("учебный год", exported)
                    self.assertNotIn("Цель обучения", exported)
                self.assertIn("ktp_download", {button.key for button in app.download_button})

    def test_ktp_rejects_malformed_and_wrong_count_without_success(self):
        for lang in ("ru", "kk"):
            for answer, code in [("not JSON", "json_invalid"), ("[]", "ktp_count_invalid")]:
                with self.subTest(lang=lang, code=code):
                    self.answer = answer
                    app = self.new_app("ktp", lang, with_key=True)
                    app.selectbox("ktp_quarters").select(1).run()
                    app.number_input("ktp_weeks_1").set_value(1)
                    self.click(app, "ktp")
                    self.assert_error(app, code, lang)
                    self.assert_no_result(app, "ktp", lang)
                    self.assertFalse(any("Готовлю четверть" in hint.value or "дайындалуда" in hint.value for hint in app.info))
                    self.assertNotIn("ktp_download", {button.key for button in app.download_button})

    def test_pdf_upload_is_sent_as_pdf_bytes(self):
        raw = b"%PDF-1.4\n% offline attachment fixture\n%%EOF"
        app = self.new_app("ktp", with_key=True)
        app.selectbox("ktp_quarters").select(1).run()
        app.number_input("ktp_weeks_1").set_value(1)
        app.radio("ktp_source").set_value("pdf").run()
        app.file_uploader("ktp_pdf").upload("topics.pdf", raw, "application/pdf").run()
        self.answer = ktp_response()
        self.click(app, "ktp")
        self.assertEqual(len(app.error), 0)
        self.assertEqual(self.ai.call_args.kwargs["attachments"], [("application/pdf", raw)])
        self.assertTrue(self.ai.call_args.kwargs["json_mode"])

    def test_language_switch_resends_selected_plan_and_source_to_browser(self):
        app = self.new_app("ktp")
        app.radio("ktp_source").set_value("pdf").run()
        for language in ("kk", "ru"):
            app.selectbox("language").select(language).run()
            # The browser retains localized labels across reruns unless explicitly updated.
            for key, value in (("module", "ktp"), ("ktp_source", "pdf")):
                radio = app.radio(key)
                self.assertEqual(radio.value, value)
                self.assertTrue(radio.proto.set_value)
                self.assertIn(radio.proto.raw_value, radio.proto.options)
            app.number_input("ktp_grade").set_value(7).run()
            self.assert_no_exception(app)
            self.assertEqual(app.radio("module").value, "ktp")
            self.assertIn("ktp_pdf", {upload.key for upload in app.file_uploader})

    def test_eda_demo_metrics_and_ml_extremes_without_key(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app("eda", lang)
                self.click(app, "eda")
                self.assertEqual([item.value for item in app.metric], ["6", "77.50", "77.50"])
                self.assertEqual(len(app.session_state[f"result_eda_{lang}"]["values"]), 6)
                self.assertIn("eda_download", {button.key for button in app.download_button})
                app.radio("module").set_value("ml").run()
                self.assertTrue(app.warning)
                for value, activity, expected in [(100, 3, 2), (0, 1, 0)]:
                    for name in ("attendance", "homework", "test"):
                        app.slider(f"ml_{name}").set_value(value)
                    app.selectbox("ml_activity").select(activity)
                    self.click(app, "ml")
                    self.assertEqual(app.session_state[f"result_ml_{lang}"]["prediction"], expected)
                    self.assertEqual(len(app.error), 0)
        self.ai.assert_not_called()

    def test_photo_upload_sends_image_and_exports_result(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app("photo", lang, with_key=True)
                app.file_uploader("photo_upload").upload("exercise.png", demo_png(), "image/png").run()
                criteria = "2 + 2 = 4"
                app.text_area(f"photo_criteria_{lang}").set_value(criteria)
                self.click(app, "photo")
                self.assertEqual(len(app.error), 0)
                self.assertIn(criteria, self.ai.call_args.args[0])
                attachments = self.ai.call_args.kwargs["attachments"]
                self.assertEqual(len(attachments), 1)
                self.assertEqual(attachments[0][0], "image/jpeg")
                with Image.open(io.BytesIO(attachments[0][1])) as uploaded:
                    self.assertEqual(uploaded.size, (80, 60))
                result = app.session_state[f"result_photo_{lang}"]
                self.assertIn(self.answer.splitlines()[0], document_text(result["document"]))
                self.assertIn("photo_download", {button.key for button in app.download_button})

    def test_eda_excel_upload_excludes_missing_values_and_handles_text_only(self):
        numeric_file = io.BytesIO()
        pd.DataFrame({"DEMO": ["DEMO-01", "DEMO-02", "DEMO-03"], "Score": [10, 20, None]}).to_excel(numeric_file, index=False)
        text_file = io.BytesIO()
        pd.DataFrame({"DEMO": ["DEMO-01"], "Comment": ["Учебный пример"]}).to_excel(text_file, index=False)
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app("eda", lang)
                app.radio("eda_source").set_value("excel").run()
                app.file_uploader("eda_upload").upload("grades.xlsx", numeric_file.getvalue(), mime).run()
                app.selectbox(f"eda_column_{lang}").select("Score")
                self.click(app, "eda")
                result = app.session_state[f"result_eda_{lang}"]
                self.assertEqual(result["values"], [10, 20])
                self.assertEqual(result["excluded"], 1)
                self.assertEqual([item.value for item in app.metric], ["2", "15.00", "15.00"])
                app.file_uploader("eda_upload").upload("comments.xlsx", text_file.getvalue(), mime).run()
                self.assert_no_exception(app)
                self.assertTrue(app.warning)
                self.assertIn("Сандық бағандар жоқ" if lang == "kk" else "Нет числовых столбцов", app.warning[0].value)
                self.assertNotIn("eda_generate", {button.key for button in app.button})
        self.ai.assert_not_called()

    def test_photo_missing_and_invalid_files_show_localized_errors(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app("photo", lang, with_key=True)
                self.click(app, "photo")
                self.assert_error(app, "file_missing", lang)
                app.file_uploader("photo_upload").upload("bad.png", b"not an image", "image/png").run()
                self.click(app, "photo")
                self.assert_error(app, "image_invalid", lang)
                self.assert_no_result(app, "photo", lang)
        self.ai.assert_not_called()

    def test_characteristic_and_warmup_generate_persist_and_export(self):
        for lang in ("ru", "kk"):
            for module in ("characteristic", "warmup"):
                with self.subTest(lang=lang, module=module):
                    app = self.new_app(module, lang, with_key=True)
                    self.click(app, module)
                    self.assertEqual(len(app.error), 0)
                    self.assertEqual(len(app.success), 1)
                    self.assertIn("қазақ тілінде" if lang == "kk" else "русском языке", self.ai.call_args.args[0])
                    result = app.session_state[f"result_{module}_{lang}"]
                    self.assertEqual(result["text"], self.answer)
                    self.assertIn(self.answer.splitlines()[0], document_text(result["document"]))
                    self.assertTrue({f"{module}_download", f"{module}_text_download"} <= {button.key for button in app.download_button})
                    count = self.ai.call_count
                    app.run()
                    self.assert_no_exception(app)
                    self.assertEqual(self.ai.call_count, count)
                    self.assertEqual(app.session_state[f"result_{module}_{lang}"]["document"], result["document"])

    def test_ai_modules_missing_key_show_localized_error(self):
        for lang in ("ru", "kk"):
            for module in ("cards", "ktp", "ksp", "photo", "characteristic", "warmup"):
                with self.subTest(lang=lang, module=module):
                    app = self.new_app(module, lang)
                    if module == "cards":
                        app.radio("cards_source").set_value("ai").run()
                    elif module == "photo":
                        app.file_uploader("photo_upload").upload("exercise.png", demo_png(), "image/png").run()
                    self.click(app, module)
                    self.assert_error(app, "missing_key", lang)
                    self.assert_no_result(app, module, lang)

    def test_provider_failures_do_not_expose_secret_or_replace_saved_result(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.new_app("warmup", lang, with_key=True)
                self.answer = "Previously generated lesson"
                self.click(app, "warmup")
                previous = app.session_state[f"result_warmup_{lang}"]["document"]
                for code in ("auth", "quota", "timeout", "blocked", "invalid_response"):
                    self.answer = ai_service.AIError(code)
                    self.click(app, "warmup")
                    self.assert_error(app, code, lang)
                    self.assertNotIn(TEST_KEY, "\n".join(item.value for item in app.error))
                    self.assertEqual(app.session_state[f"result_warmup_{lang}"]["document"], previous)

    def test_kazakh_widget_labels_and_options(self):
        expected = {
            "cards": {"cards_source": "Сұрақтар көзі", "cards_easy": "Жеңіл", "cards_medium": "Орташа", "cards_hard": "Қиын", "cards_generate": "Карточкалар мен жауаптарды жасау"},
            "ktp": {"ktp_subject_kk": "Пән", "ktp_grade": "Сынып", "ktp_year": "Оқу жылы", "ktp_quarters": "Тоқсан саны", "ktp_hours": "Аптасына сағат", "ktp_source": "Тақырыптар көзі", "ktp_topics_kk": "Өз бағдарламаңыздағы тақырыптар", "ktp_generate": "КТЖ жасау"},
            "ksp": {"ksp_teacher": "Педагог (міндетті емес)", "ksp_subject_kk": "Пән", "ksp_grade": "Сынып", "ksp_topic_kk": "Сабақ тақырыбы", "ksp_target_kk": "Өз бағдарламаңыздағы оқу мақсаты", "ksp_generate": "ҚМЖ жасау"},
            "eda": {"eda_source": "Деректер", "eda_column_kk": "Көрсеткіш", "eda_generate": "Талдау жасау"},
            "ml": {"ml_attendance": "Қатысу (%)", "ml_homework": "Үй тапсырмасын орындау (%)", "ml_test": "Тесттердің орташа балы (0–100)", "ml_activity": "Белсенділік", "ml_generate": "Демонстрацияны іске қосу"},
            "photo": {"photo_upload": "Фото — JPG немесе PNG, 10 МБ дейін", "photo_criteria_kk": "Тапсырма шарты мен критерийлер (бар болса)", "photo_generate": "Жұмысты тексеру"},
            "characteristic": {"char_name_kk": "Оқушының шартты атауы", "char_class": "Сынып", "char_attendance": "Қатысу (%)", "char_performance": "Үлгерім", "char_behaviour": "Мінез-құлық", "char_traits_kk": "Байқалған жетістіктер мен қасиеттер", "characteristic_generate": "Мінездеме жасау"},
            "warmup": {"warmup_topic_kk": "Сабақ тақырыбы", "warmup_grade": "Сынып", "warmup_minutes": "Бір сергітудің ұзақтығы (мин)", "warmup_generate": "Сергіту сәттерін таңдау"},
        }
        app = self.new_app(lang="kk")
        for module, labels in expected.items():
            with self.subTest(module=module):
                app.radio("module").set_value(module).run()
                self.assert_no_exception(app)
                for key, label in labels.items():
                    self.assertEqual(app.get_by_key(key).label, label)
                self.assertEqual(app.radio("module").label, "Құралдар")
                if module == "ml":
                    self.assertEqual(app.selectbox("ml_activity").options, ["Төмен", "Орташа", "Жоғары"])
                elif module == "characteristic":
                    self.assertEqual(app.selectbox("char_performance").options, ["Өте жақсы", "Жақсы", "Орташа", "Қолдау қажет"])
                    self.assertEqual(app.selectbox("char_behaviour").options, ["Жауапты", "Сабырлы", "Өзін-өзі реттеуге қолдау қажет"])

    def test_kazakh_labels_on_optional_upload_and_ai_paths(self):
        app = self.new_app(lang="kk")
        self.assertEqual(app.radio("cards_source").options, ["Демонстрациялық қор", "Excel", "Google Sheets", "ЖИ арқылы жасау"])
        for source, labels in {
            "ai": {"cards_topic_kk": "Тақырып", "cards_grade": "Сынып", "cards_students_kk": "Оқушылар: әр жолға біреу"},
            "excel": {"cards_upload": "Сұрақтары мен оқушылары бар Excel"},
            "google": {"cards_google_url": "Google кестесінің сілтемесі", "cards_google_load": "Кестені жүктеу"},
        }.items():
            app.radio("cards_source").set_value(source).run()
            self.assert_no_exception(app)
            for key, label in labels.items():
                self.assertEqual(app.get_by_key(key).label, label)
        app.radio("module").set_value("ktp").run()
        for quarter in range(1, 5):
            self.assertEqual(app.number_input(f"ktp_weeks_{quarter}").label, f"Апта: {quarter}-тоқсан")
        self.assertEqual(app.radio("ktp_source").options, ["Мәтін", "PDF"])
        app.radio("ktp_source").set_value("pdf").run()
        self.assertEqual(app.file_uploader("ktp_pdf").label, "Бағдарлама немесе тақырыптар — PDF, 10 МБ дейін")
        app.radio("module").set_value("eda").run()
        app.radio("eda_source").set_value("excel").run()
        self.assertEqual(app.file_uploader("eda_upload").label, "Сандық көрсеткіштері бар кесте")
        self.assert_no_exception(app)
        self.ai.assert_not_called()


if __name__ == "__main__":
    unittest.main()
