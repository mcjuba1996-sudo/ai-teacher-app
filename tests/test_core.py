"""Regression checks for teacher inputs, bilingual exports and safe validation."""
import copy
import io
import json
import unittest
from collections import Counter

import pandas as pd
from docx import Document

from bilim_core import (
    ValidationError, build_ksp_prompt, build_ktp_prompt, cards_documents,
    demo_bank, demo_students, ksp_document, ktp_document, make_template,
    normalize_bank, parse_json_response, parse_students, read_card_workbook,
    select_variants, validate_ksp, validate_ktp,
)


def document_text(blob):
    document = Document(io.BytesIO(blob))
    return "\n".join([p.text for p in document.paragraphs] + [cell.text for table in document.tables for row in table.rows for cell in row.cells])


def ksp_sample():
    return {
        "section": "Алгоритмы",
        "lesson_targets": "Составить и проверить линейный алгоритм.",
        "learning_targets": "MODEL-INVENTED-TARGET",
        "stages": [
            {"time": "Начало — 5 мин", "teacher": "Учитель задаёт вопрос об алгоритмах.", "student": "Ученики приводят примеры алгоритмов.", "resources": "Доска"},
            {"time": "Основная часть — 35 мин", "teacher": "Учитель выдаёт задачу и помогает проверить решение.", "student": "Ученики составляют алгоритм и проверяют его на примерах.", "resources": "Карточки"},
            {"time": "Завершение — 5 мин", "teacher": "Учитель предлагает объяснить порядок команд.", "student": "Ученики объясняют решение и исправляют ошибки.", "resources": "Тетрадь"},
        ],
    }


def ktp_sample():
    return [{"quarter": 1 if n <= 2 else 2, "lesson_num": n, "section": "Алгоритмдер",
             "topic": f"Алгоритмді тексеру {n}", "targets": "Алгоритмді құру және тексеру"} for n in range(1, 5)]


class ErrorAssertions(unittest.TestCase):
    def assert_error(self, code, function, *args, **kwargs):
        with self.assertRaises(ValidationError) as caught:
            function(*args, **kwargs)
        self.assertEqual(caught.exception.code, code)
        return caught.exception


class JsonValidationTests(ErrorAssertions):
    def test_valid_json_and_fence(self):
        self.assertEqual(parse_json_response('{"section":"Алгоритмдер"}'), {"section": "Алгоритмдер"})
        self.assertEqual(parse_json_response('```json\n[{"quarter": 1}]\n```'), [{"quarter": 1}])

    def test_malformed_response_never_becomes_empty_success(self):
        for text in ("not JSON", '{"stages": [}', '[{"topic":"Algorithms"},', "", "null", '{"value": NaN}'):
            with self.subTest(text=text):
                self.assert_error("json_invalid", parse_json_response, text)

    def test_missing_semantic_content_is_rejected_after_parsing(self):
        self.assert_error("field_required", validate_ksp, parse_json_response("{}"))
        self.assert_error("ktp_count_invalid", validate_ktp, parse_json_response("[]"), {1: 2})


class BankTests(ErrorAssertions):
    def test_both_languages_have_canonical_difficulties(self):
        for language, names, levels in (
            ("ru", ["Вопрос", "Ответ", "Сложность"], ["Лёгкий", "Средний", "Сложный"]),
            ("kk", ["Сұрақ", "Жауап", "Қиындық"], ["Жеңіл", "Орташа", "Қиын"]),
        ):
            with self.subTest(language=language):
                source = pd.DataFrame([[f"Question {n}", f"Answer {n}", level] for n, level in enumerate(levels)], columns=names)
                normalized = normalize_bank(source)
                self.assertEqual(normalized.columns.tolist(), ["question", "answer", "difficulty"])
                self.assertEqual(normalized["difficulty"].tolist(), ["easy", "medium", "hard"])

    def test_identical_duplicates_removed_and_conflicts_rejected(self):
        source = demo_bank("ru").iloc[:1]
        duplicate = pd.concat([source, source], ignore_index=True)
        self.assertEqual(len(normalize_bank(duplicate)), 1)
        duplicate.loc[1, "answer"] = "A conflicting answer"
        self.assert_error("bank_question_conflict", normalize_bank, duplicate)

    def test_case_and_whitespace_inside_programming_questions_remain_significant(self):
        bank = pd.DataFrame([
            {"question": 'What does print("A") output?', "answer": "A", "difficulty": "easy"},
            {"question": 'What does print("a") output?', "answer": "a", "difficulty": "easy"},
            {"question": 'What does print("a  b") output?', "answer": "a  b", "difficulty": "easy"},
            {"question": 'What does print("a b") output?', "answer": "a b", "difficulty": "easy"},
        ])
        self.assertEqual(len(normalize_bank(bank)), 4)

    def test_missing_fields_invalid_levels_and_empty_cells(self):
        bank = demo_bank("kk")
        self.assert_error("bank_columns_missing", normalize_bank, bank.drop(columns="answer"))
        bank.loc[0, "difficulty"] = "unknown"
        self.assert_error("bank_difficulty_invalid", normalize_bank, bank)
        bank = demo_bank("ru")
        bank.loc[0, "question"] = "   "
        self.assert_error("bank_cell_empty", normalize_bank, bank)

    def test_bilingual_workbooks_round_trip_and_use_demo_students(self):
        for lang in ("ru", "kk"):
            with self.subTest(lang=lang):
                blob = make_template(lang)
                bank, students = read_card_workbook(blob)
                pd.testing.assert_frame_equal(bank, demo_bank(lang))
                self.assertEqual(students, demo_students())
                self.assertEqual(Counter(bank["difficulty"]), {"easy": 5, "medium": 5, "hard": 5})

    def test_invalid_workbook_is_a_domain_error(self):
        self.assert_error("workbook_invalid", read_card_workbook, b"not an xlsx")

    def test_nested_model_values_are_rejected_as_a_domain_error(self):
        bank = pd.DataFrame([{"question": ["unexpected", "list"], "answer": "Answer", "difficulty": "easy"}])
        self.assert_error("bank_cell_invalid", normalize_bank, bank)

    def test_student_names_are_deduplicated_and_empty_input_rejected(self):
        self.assertEqual(parse_students("DEMO-01, DEMO-02\nDEMO-01\n"), ["DEMO-01", "DEMO-02"])
        self.assert_error("students_empty", parse_students, " , \n ")


class CardTests(ErrorAssertions):
    def test_each_variant_has_exact_requested_difficulties_without_duplicates(self):
        counts = {"easy": 2, "medium": 2, "hard": 1}
        for lang in ("ru", "kk"):
            variants = select_variants(demo_bank(lang), demo_students(), counts, seed=42)
            self.assertEqual(len(variants), 3)
            for variant in variants:
                self.assertEqual(Counter(q["difficulty"] for q in variant["questions"]), counts)
                self.assertEqual(len({q["question"] for q in variant["questions"]}), 5)
            self.assertEqual(variants, select_variants(demo_bank(lang), demo_students(), counts, seed=42))

    def test_medium_only_uses_medium_questions(self):
        variants = select_variants(demo_bank("ru"), ["DEMO-01"], {"medium": 3}, seed=1)
        self.assertEqual([q["difficulty"] for q in variants[0]["questions"]], ["medium"] * 3)

    def test_insufficient_bank_never_substitutes_another_difficulty(self):
        error = self.assert_error("bank_insufficient", select_variants, demo_bank("ru"), demo_students(), {"hard": 6})
        self.assertEqual(error.params, {"difficulty": "hard", "requested": 6, "available": 5})

    def test_zero_and_negative_counts_fail_before_generation(self):
        self.assert_error("counts_empty", select_variants, demo_bank("ru"), demo_students(), {})
        self.assert_error("integer_invalid", select_variants, demo_bank("ru"), demo_students(), {"easy": -1})

    def test_card_and_key_documents_share_exact_variants(self):
        bank = pd.DataFrame([{"question": "Question unique 1", "answer": "ANSWER-SENTINEL-901", "difficulty": "medium"}])
        variants = select_variants(bank, ["DEMO-01", "DEMO-02"], {"medium": 1}, seed=2)
        for lang, task in (("ru", "Задание"), ("kk", "Тапсырма")):
            cards, keys = cards_documents(variants, lang)
            card_text, key_text = document_text(cards), document_text(keys)
            self.assertEqual(card_text.count("Question unique 1"), 2)
            self.assertEqual(key_text.count("Question unique 1"), 2)
            self.assertNotIn("ANSWER-SENTINEL-901", card_text)
            self.assertEqual(key_text.count("ANSWER-SENTINEL-901"), 2)
            self.assertIn(f"{task} 1", card_text)
            for student in demo_students()[:2]:
                self.assertIn(student, card_text)
                self.assertIn(student, key_text)


class LessonPlanTests(ErrorAssertions):
    def test_ksp_prompt_changes_with_grade_and_learning_target(self):
        args = {"subject": "Информатика", "topic": "Алгоритмы", "lang": "ru"}
        first = build_ksp_prompt(grade=5, target="Составить линейный алгоритм", **args)
        second = build_ksp_prompt(grade=11, target="Спроектировать нейронную сеть", **args)
        self.assertNotEqual(first, second)
        self.assertIn('"grade": 5', first)
        self.assertIn("Составить линейный алгоритм", first)
        self.assertIn('"grade": 11', second)
        self.assertIn("Спроектировать нейронную сеть", second)

    def test_ksp_requires_exactly_three_complete_stages(self):
        self.assertEqual(len(validate_ksp(ksp_sample())["stages"]), 3)
        invalid = ksp_sample()
        invalid["stages"].pop()
        self.assert_error("ksp_stages_invalid", validate_ksp, invalid)
        invalid = ksp_sample()
        invalid["stages"][1]["teacher"] = "..."
        self.assert_error("field_required", validate_ksp, invalid)

    def test_ksp_rejects_broken_timing_before_exporting_a_new_plan(self):
        valid_sets = [
            ["Басы — 5 минут", "Ортасы — 30 минут", "Соңы — 10 минут"],
            ["Начало 0–5 мин", "Практика 5–40 мин", "Итог 40–45 мин"],
        ]
        invalid_sets = [
            ["Басы — 5 минут", "Ортасы — 35 минут", "Соңы — 20 минут"],
            ["Начало 0–5 мин", "Практика 5–40 мин", "Итог 38–45 мин"],
            ["Начало", "Практика", "Итог"],
        ]
        for valid, labels in [(True, v) for v in valid_sets] + [(False, v) for v in invalid_sets]:
            with self.subTest(labels=labels):
                data = ksp_sample()
                for stage, label in zip(data["stages"], labels):
                    stage["time"] = label
                if valid:
                    self.assertEqual(validate_ksp(data, require_timing=True)["stages"], data["stages"])
                else:
                    self.assert_error("ksp_timing_invalid", validate_ksp, data, require_timing=True)

    def test_ksp_export_preserves_teacher_target_and_has_three_stage_rows(self):
        target = "TEACHER-SUPPLIED-TARGET-EXACT"
        for lang, heading in (("ru", "Цель обучения"), ("kk", "Оқу мақсаты")):
            blob = ksp_document(ksp_sample(), "DEMO-TEACHER", 5, "Алгоритмы", target, lang)
            text = document_text(blob)
            self.assertIn(target, text)
            self.assertNotIn("MODEL-INVENTED-TARGET", text)
            self.assertIn(heading, text)
            self.assertEqual(len(Document(io.BytesIO(blob)).tables[1].rows), 4)

    def test_blank_optional_teacher_is_exported_as_blank(self):
        blob = ksp_document(ksp_sample(), "", 5, "Алгоритмы", "Составить алгоритм.", "ru")
        document = Document(io.BytesIO(blob))
        self.assertEqual(document.tables[0].rows[1].cells[1].text, "")

    def test_ksp_keeps_the_selected_subject_separate_from_the_topic(self):
        for lang, label, subject, topic in (("kk", "Пән", "Физика", "Тығыздық"),
                                             ("ru", "Предмет", "Физика", "Плотность")):
            with self.subTest(lang=lang):
                blob = ksp_document(ksp_sample(), "", 7, topic, "Цель", lang, subject=subject)
                doc = Document(io.BytesIO(blob))
                self.assertIn([label, subject], [[c.text for c in row.cells] for row in doc.tables[0].rows])
                self.assertIn(topic, [p.text for p in doc.paragraphs])

    def test_ktp_prompt_encodes_distribution_and_year(self):
        args = {"subject": "Информатика", "grade": 7, "hours_per_week": 1, "topics": "Алгоритмы", "lang": "kk", "academic_year": "2026–2027"}
        first = build_ktp_prompt(quarters_weeks={1: 10}, **args)
        second = build_ktp_prompt(quarters_weeks={1: 5, 2: 5}, **args)
        self.assertNotEqual(first, second)
        self.assertIn('"quarter_hours": {"1": 5, "2": 5}', second)
        self.assertIn("2026–2027", second)

    def test_ktp_validates_total_distribution_and_numbering(self):
        sample = ktp_sample()
        self.assertEqual(validate_ktp(sample, {1: 2, 2: 2}), sample)
        self.assert_error("ktp_count_invalid", validate_ktp, sample[:-1], {1: 2, 2: 2})
        wrong_quarter = copy.deepcopy(sample)
        wrong_quarter[-1]["quarter"] = 1
        self.assert_error("ktp_quarters_invalid", validate_ktp, wrong_quarter, {1: 2, 2: 2})
        wrong_number = copy.deepcopy(sample)
        wrong_number[-1]["lesson_num"] = 3
        self.assert_error("ktp_numbering_invalid", validate_ktp, wrong_number, {1: 2, 2: 2})

    def test_single_quarter_response_uses_local_numbering_even_for_second_quarter(self):
        part = ktp_sample()[2:]
        self.assert_error("ktp_numbering_invalid", validate_ktp, part, {2: 2})
        part[0]["lesson_num"], part[1]["lesson_num"] = 1, 2
        self.assertEqual([row["quarter"] for row in validate_ktp(part, {2: 2})], [2, 2])

    def test_second_quarter_prompt_does_not_give_a_first_quarter_row_example(self):
        prompt = build_ktp_prompt("Физика", 7, {2: 3}, 2, "Масса и объём", "ru", "2026–2027")
        row_example, _ = json.JSONDecoder().raw_decode(prompt.split("Структура строки: ", 1)[1])
        self.assertEqual((row_example["quarter"], row_example["lesson_num"]), (2, 1))

    def test_multi_page_calendar_table_repeats_the_legend_and_lists_hours_per_quarter(self):
        blob = ktp_document(ktp_sample(), "Информатика", 7, 1, "2026–2027", "kk")
        doc = Document(io.BytesIO(blob))
        self.assertTrue(doc.tables[0].rows[0]._tr.xpath("./w:trPr/w:tblHeader"))
        self.assertIn("1-тоқсан — 2 сағ", document_text(blob))
        self.assertIn("2-тоқсан — 2 сағ", document_text(blob))
        self.assertGreater(doc.tables[0].columns[4].width, doc.tables[0].columns[1].width)

    def test_ktp_kazakh_export_localizes_headers_and_uses_requested_year(self):
        blob = ktp_document(ktp_sample(), "Информатика", 7, 1, "2026–2027", "kk")
        text = document_text(blob)
        self.assertIn("2026–2027 оқу жылы", text)
        self.assertIn("Тоқсан", text)
        self.assertIn("Сабақ тақырыбы", text)
        self.assertNotIn("2024-2025", text)
        self.assertNotIn("учебный год", text)
        self.assertNotIn("Цель обучения", text)
        table = Document(io.BytesIO(blob)).tables[0]
        self.assertEqual(len(table.rows), 5)
        self.assertEqual([row.cells[0].text for row in table.rows[1:]], ["1", "2", "3", "4"])


if __name__ == "__main__":
    unittest.main()
