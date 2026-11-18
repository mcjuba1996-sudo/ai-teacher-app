"""Exercise the Python taskbook through its Streamlit controls without network access."""

import json
import os
from pathlib import Path
import unittest
from unittest.mock import patch

import streamlit as st
from streamlit.testing.v1 import AppTest

import ai_service
from ui_helpers import ERRORS


APP = Path(__file__).resolve().parents[1] / "app.py"
TEST_KEY = "offline-taskbook-test-key"
REQUESTS = {
    "ru": "Составь задачу о числе колёс у велосипедов в школьном дворе.",
    "kk": "Мектеп ауласындағы велосипедтердің дөңгелек саны туралы есеп құрастыр.",
}


def exercise(lang):
    return {
        "title": "Велосипеды во дворе" if lang == "ru" else "Ауладағы велосипедтер",
        "statement": ("Введите число двухколёсных велосипедов и выведите общее число колёс."
                      if lang == "ru" else "Екі дөңгелекті велосипедтер санын енгізіп, дөңгелектердің жалпы санын шығарыңыз."),
        "examples": [{"input": "4", "output": "8"}, {"input": "7", "output": "14"}],
        "hint": "Умножьте число на два." if lang == "ru" else "Санды екіге көбейтіңіз.",
        "solution": "bikes = int(input())\nprint(bikes * 2)",
    }


def page_text(app):
    return "\n".join(item.value for elements in
                     (app.title, app.header, app.subheader, app.markdown, app.code, app.text)
                     for item in elements)


def downloads(app):
    return {button.key for button in app.download_button}


class TaskbookAppTests(unittest.TestCase):
    def setUp(self):
        self.settings = dict.fromkeys(
            ("GEMINI_API_KEY", "GOOGLE_API_KEY", "DEFAULT_API_KEY", "GEMINI_MODEL"), "")
        self.enterContext(patch.dict(os.environ, self.settings))
        self.client = self.enterContext(patch("ai_service.genai.Client"))

    def tearDown(self):
        self.client.assert_not_called()

    def open_taskbook(self, lang="ru", with_key=False):
        app = AppTest.from_file(APP, default_timeout=45)
        for key, value in self.settings.items():
            app.secrets[key] = value
        if with_key:
            app.secrets["DEFAULT_API_KEY"] = TEST_KEY
        app.run()
        app.selectbox("language").select(lang)
        app.radio("module").set_value("python").run()
        self.assertFalse(app.exception)
        return app

    def test_bank_switches_stories_tasks_and_languages_without_gemini(self):
        app = self.open_taskbook()
        with patch("ai_service.generate_text") as generate:
            for lang in ("ru", "kk"):
                with self.subTest(lang=lang):
                    app.selectbox("language").select(lang).run()
                    app.selectbox("py_story").select("pizza").run()
                    self.assertEqual(len(app.selectbox("py_number").options), 15)
                    app.selectbox("py_number").select(1).run()
                    self.assertIn("500000\n120000", [item.value for item in app.code])
                    self.assertIn("4 20000", [item.value for item in app.code])
                    app.selectbox("py_number").select(15).run()
                    expected = "Доступ открыт" if lang == "ru" else "Кіруге рұқсат берілді"
                    self.assertIn(expected, [item.value for item in app.code])
                    # The previously selected number lies outside the second story.
                    app.selectbox("py_story").select("mars").run()
                    self.assertFalse(app.exception)
                    self.assertEqual(len(app.selectbox("py_number").options), 10)
                    app.selectbox("py_number").select(1).run()
                    self.assertIn("AAABBBCC", page_text(app))
                    self.assertIn("A3B3C2", page_text(app))
                    app.selectbox("py_number").select(10).run()
                    self.assertIn("abcabcbb", page_text(app))
                    self.assertIn("py_bank_download", downloads(app))
                    self.assertFalse(app.error)
            generate.assert_not_called()

    def test_bank_example_seeds_an_editable_request_without_calling_gemini(self):
        app = self.open_taskbook(lang="kk")
        app.selectbox("py_story").select("mars").run()
        app.selectbox("py_number").select(5).run()
        title = app.subheader[0].value
        with patch("ai_service.generate_text") as generate:
            app.button("py_variation").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.radio("py_mode").value, "ai")
            self.assertEqual(app.selectbox("py_level").value, "hard")
            self.assertIn(title, app.text_area("py_request_kk").value)
            app.text_area("py_request_kk").set_value(REQUESTS["kk"]).run()
            self.assertEqual(app.text_area("py_request_kk").value, REQUESTS["kk"])
            generate.assert_not_called()

    def test_generation_uses_request_and_keeps_downloads_separate_by_language(self):
        app = self.open_taskbook(with_key=True)
        app.radio("py_mode").set_value("ai").run()
        with (patch("ai_service.generate_text") as generate,
              patch("streamlit.download_button", wraps=st.download_button) as download):
            for lang in ("ru", "kk"):
                with self.subTest(lang=lang):
                    app.selectbox("language").select(lang).run()
                    self.assertEqual(app.radio("py_mode").value, "ai")
                    self.assertNotIn("py_student", downloads(app))
                    app.text_area(f"py_request_{lang}").set_value(REQUESTS[lang])
                    app.number_input("py_grade").set_value(8)
                    app.selectbox("py_level").select("hard")
                    answer = exercise(lang)
                    generate.return_value = json.dumps(answer, ensure_ascii=False)
                    app.button("py_generate").click().run()
                    self.assertFalse(app.exception)
                    self.assertFalse(app.error)
                    prompt, key = generate.call_args.args
                    self.assertEqual(key, TEST_KEY)
                    self.assertIn(REQUESTS[lang], prompt)
                    self.assertIn('"grade": 8', prompt)
                    self.assertIn('"level": "hard"', prompt)
                    self.assertIn("қазақ тілінде" if lang == "kk" else "русском языке", prompt)
                    self.assertTrue(generate.call_args.kwargs["json_mode"])
                    for field in ("title", "statement", "hint", "solution"):
                        self.assertIn(answer[field], page_text(app))
                    self.assertTrue({"py_student", "py_teacher"} <= downloads(app))
                    files = {call.kwargs["key"]: call.args[1] for call in download.call_args_list}
                    self.assertIn(answer["statement"], files["py_student"])
                    self.assertIn("14", files["py_student"])
                    for field in ("hint", "solution"):
                        self.assertNotIn(answer[field], files["py_student"])
                        self.assertIn(answer[field], files["py_teacher"])
                    app.download_button("py_student").click().run()
                    self.assertIn("py_teacher", downloads(app))
            app.selectbox("language").select("ru").run()
            self.assertIn(exercise("ru")["title"], page_text(app))
            self.assertNotIn(exercise("kk")["title"], page_text(app))
            app.radio("module").set_value("cards").run()
            app.radio("module").set_value("python").run()
            app.radio("py_mode").set_value("ai").run()
            self.assertIn(exercise("ru")["title"], page_text(app))
            self.assertEqual(generate.call_count, 2)

    def test_invalid_answers_and_service_failure_keep_previous_exercise(self):
        app = self.open_taskbook(lang="kk", with_key=True)
        app.radio("py_mode").set_value("ai").run()
        app.text_area("py_request_kk").set_value(REQUESTS["kk"])
        valid = exercise("kk")
        with patch("ai_service.generate_text", return_value=json.dumps(valid)) as generate:
            app.button("py_generate").click().run()
            self.assertTrue({"py_student", "py_teacher"} <= downloads(app))
            invalid = dict(valid, title="invalid exercise")
            for response in ("malformed", json.dumps(dict(invalid, examples=valid["examples"][:1])),
                             json.dumps(dict(invalid, examples=[valid["examples"][0]] * 2)),
                             json.dumps(dict(invalid, solution="def :"))):
                with self.subTest(response=response):
                    generate.return_value = response
                    app.button("py_generate").click().run()
                    self.assertFalse(app.exception)
                    self.assertEqual(app.error[0].value, ERRORS["json_invalid"][1])
                    self.assertFalse(app.success)
                    self.assertIn(valid["title"], page_text(app))
                    self.assertNotIn(invalid["title"], page_text(app))
            generate.side_effect = ai_service.AIError("quota")
            app.button("py_generate").click().run()
            self.assertEqual(app.error[0].value, ERRORS["quota"][1])
            self.assertTrue({"py_student", "py_teacher"} <= downloads(app))
            self.assertIn(valid["title"], page_text(app))

    def test_empty_request_and_missing_key_show_localized_errors(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                app = self.open_taskbook(lang)
                app.radio("py_mode").set_value("ai").run()
                app.text_area(f"py_request_{lang}").set_value("   ")
                with patch("ai_service.generate_text") as generate:
                    app.button("py_generate").click().run()
                    generate.assert_not_called()
                self.assertEqual(app.error[0].value, ERRORS["field_required"][lang == "kk"])
                app.text_area(f"py_request_{lang}").set_value(REQUESTS[lang])
                app.button("py_generate").click().run()
                self.assertFalse(app.exception)
                self.assertEqual(app.error[0].value, ERRORS["missing_key"][lang == "kk"])
                self.assertNotIn("py_student", downloads(app))
