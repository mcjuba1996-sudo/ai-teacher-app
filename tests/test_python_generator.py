"""Teacher brief, response schema and lossless Python sample validation."""
import json
import unittest

from bilim_core import ValidationError, parse_json_response
from python_generator import MAX_REQUEST_LENGTH, build_task_prompt, validate_task


def echo_task():
    return {
        "title": "Копия строки",
        "statement": "Прочитайте весь ввод и напечатайте его, добавив перевод строки.",
        "examples": [{"input": "", "output": "\n"}, {"input": "  A ", "output": "  A \n"}],
        "hint": "Сохраните пробелы в начале и конце строки.",
        "solution": "import sys\nprint(sys.stdin.read())\n",
    }


class TaskPromptTests(unittest.TestCase):
    def test_teacher_brief_and_chosen_language(self):
        brief = 'Задача про "снег".\nПрименить for, без sum.'
        for lang, phrase in (("ru", "на русском языке"), ("kk", "қазақ тілінде")):
            with self.subTest(lang=lang):
                prompt = build_task_prompt(brief, 7, "medium", lang)
                data = json.loads(next(line for line in prompt.splitlines() if line.startswith("{")))
                self.assertEqual(data, {"request": brief, "grade": 7, "level": "medium"})
                self.assertIn(phrase, prompt)
                for key in ("title", "statement", "examples", "input", "output", "hint", "solution"):
                    self.assertIn(key, prompt)

    def test_invalid_inputs(self):
        for brief in (None, "   ", "...", "a\x00b", "x" * (MAX_REQUEST_LENGTH + 1)):
            with self.subTest(brief_type=type(brief)), self.assertRaises(ValidationError) as caught:
                build_task_prompt(brief, 8, "easy", "ru")
            self.assertEqual(caught.exception.code, "field_required")
        for grade, level, lang in ((True, "easy", "ru"), (0, "easy", "ru"), (12, "hard", "ru"),
                                   (8, "extreme", "ru"), (8, "easy", "en")):
            with self.subTest(grade=grade, level=level, lang=lang), self.assertRaises(ValidationError):
                build_task_prompt("Циклы", grade, level, lang)


class TaskValidationTests(unittest.TestCase):
    def assert_invalid(self, data):
        with self.assertRaises(ValidationError) as caught:
            validate_task(data)
        self.assertEqual(caught.exception.code, "json_invalid")

    def test_parsed_task_preserves_all_whitespace(self):
        data = echo_task()
        data["examples"].append({"input": "A\tB\n", "output": "A\tB\n\n"})
        self.assertEqual(validate_task(parse_json_response(json.dumps(data))), data)

    def test_rejects_missing_or_malformed_content(self):
        self.assert_invalid([])
        self.assert_invalid({"title": "One"})
        data = echo_task()
        for field, value in (("title", ""), ("hint", "..."), ("statement", []),
                             ("statement", "A\x00B"), ("title", "\ud800"),
                             ("solution", "def ():\n"), ("solution", "# A comment"),
                             ("solution", "a" * 10000), ("statement", "a" * 10000)):
            with self.subTest(field=field, value_type=type(value)):
                self.assert_invalid({**data, field: value})
        self.assert_invalid({**data, "unexpected": [{}]})

    def test_examples_have_two_or_three_distinct_inputs_and_bounded_strings(self):
        data = echo_task()
        samples = data["examples"]
        for examples in (None, samples[:1], samples * 2, [samples[0], samples[0]],
                         [samples[0], {"input": 1, "output": "1"}],
                         [samples[0], {"input": "x"}],
                         [samples[0], {"input": "x", "output": ""}],
                         [samples[0], {"input": "x", "output": 0}],
                         [samples[0], {"input": "x" * 3000, "output": "x"}],
                         [{"input": "A\n", "output": "1"}, {"input": "A\r\n", "output": "1"}]):
            with self.subTest(examples_type=type(examples)):
                self.assert_invalid({**data, "examples": examples})

    def test_solution_is_parsed_without_execution(self):
        data = echo_task()
        data["solution"] = 'raise RuntimeError("must never run")'
        self.assertEqual(validate_task(data)["solution"], data["solution"])


if __name__ == "__main__":
    unittest.main()
