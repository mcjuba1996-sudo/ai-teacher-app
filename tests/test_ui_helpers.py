"""File/URL boundary checks with in-memory uploads and no network calls."""
import io
import urllib.error
import urllib.parse
import unittest
from unittest.mock import patch

import pandas as pd
from docx import Document
from PIL import Image

from bilim_core import ValidationError
import ui_helpers as helpers


class ErrorAssertions(unittest.TestCase):
    def assert_error(self, code, function, *args, **kwargs):
        with self.assertRaises(ValidationError) as caught:
            function(*args, **kwargs)
        self.assertEqual(caught.exception.code, code)
        return caught.exception


class HelperTests(ErrorAssertions):
    def test_upload_accepts_nonempty_bytes_and_rejects_missing_or_oversize(self):
        self.assertEqual(helpers.upload_bytes(io.BytesIO(b"data")), b"data")
        self.assert_error("file_missing", helpers.upload_bytes, None)
        self.assert_error("file_size", helpers.upload_bytes, io.BytesIO())
        with patch.object(helpers, "MAX_UPLOAD", 4):
            self.assert_error("file_size", helpers.upload_bytes, io.BytesIO(b"12345"))
            self.assertEqual(helpers.upload_bytes(io.BytesIO(b"1234")), b"1234")

    def test_fingerprint_is_order_independent_and_sensitive_to_content(self):
        self.assertEqual(helpers.fingerprint({"a": 1, "b": 2}), helpers.fingerprint({"b": 2, "a": 1}))
        self.assertNotEqual(helpers.fingerprint(["Тақырып", 5]), helpers.fingerprint(["Тақырып", 6]))

    def test_text_document_preserves_cyrillic_and_line_breaks(self):
        document = Document(io.BytesIO(helpers.text_document("Оқу материалы", "Первая строка\n\nҚазақша жауап")))
        self.assertEqual([p.text for p in document.paragraphs], ["Оқу материалы", "Первая строка", "", "Қазақша жауап"])

    def test_image_is_decoded_resized_and_encoded_as_rgb_jpeg(self):
        upload = io.BytesIO()
        Image.new("RGBA", (2200, 1100), (30, 50, 100, 128)).save(upload, "PNG")
        mime, data = helpers.image_attachment(upload)
        self.assertEqual(mime, "image/jpeg")
        with Image.open(io.BytesIO(data)) as picture:
            self.assertEqual(picture.format, "JPEG")
            self.assertEqual(picture.mode, "RGB")
            self.assertEqual(picture.size, (2048, 1024))

    def test_image_orientation_is_applied_before_export(self):
        upload = io.BytesIO()
        exif = Image.Exif()
        exif[274] = 6
        Image.new("RGB", (40, 20), "blue").save(upload, "JPEG", exif=exif)
        _, data = helpers.image_attachment(upload)
        with Image.open(io.BytesIO(data)) as picture:
            self.assertEqual(picture.size, (20, 40))
            self.assertNotEqual(picture.getexif().get(274), 6)

    def test_invalid_image_and_missing_upload_have_readable_errors(self):
        self.assert_error("image_invalid", helpers.image_attachment, io.BytesIO(b"not an image"))
        self.assert_error("file_missing", helpers.image_attachment, None)
        with patch.object(helpers, "MAX_UPLOAD", 4):
            self.assert_error("file_size", helpers.image_attachment, io.BytesIO(b"12345"))


class GoogleCardTests(ErrorAssertions):
    URL = "https://docs.google.com/spreadsheets/d/DEMO_test-123/edit?usp=sharing#gid=0"

    def csv(self, frame):
        return frame.to_csv(index=False).encode("utf-8-sig")

    def responses(self, bank=None, students=None, language="ru"):
        if bank is None:
            bank = pd.DataFrame({"Сұрақ": ["Алгоритм дегеніміз не?"], "Жауап": ["Командалар тізбегі"], "Қиындық": ["Орташа"]}) if language == "kk" else pd.DataFrame({"Вопрос": ["Что такое алгоритм?"], "Ответ": ["Последовательность команд"], "Сложность": ["Средний"]})
        if students is None:
            students = pd.DataFrame({"Аты-жөні" if language == "kk" else "ФИО": ["DEMO-01", "DEMO-02", "DEMO-01"]})
        bodies = {"Сұрақтар_банкі" if language == "kk" else "Банк_вопросов": self.csv(bank),
                  "Оқушылар" if language == "kk" else "Ученики": self.csv(students)}

        def open_response(url, timeout):
            self.assertEqual(timeout, 15)
            parsed = urllib.parse.urlparse(url)
            self.assertEqual(parsed.scheme, "https")
            self.assertEqual(parsed.hostname, "docs.google.com")
            self.assertEqual(parsed.path, "/spreadsheets/d/DEMO_test-123/gviz/tq")
            self.assertEqual(parsed.fragment, "")
            query = urllib.parse.parse_qs(parsed.query)
            self.assertEqual(query["tqx"], ["out:csv"])
            return io.BytesIO(bodies[query["sheet"][0]])

        return open_response

    def test_google_import_normalizes_difficulty_and_deduplicates_names(self):
        with patch("ui_helpers.urllib.request.urlopen", side_effect=self.responses()) as request:
            bank, students = helpers.read_google_cards(self.URL)
        self.assertEqual(request.call_count, 2)
        self.assertEqual(bank.columns.tolist(), ["question", "answer", "difficulty"])
        self.assertEqual(bank["difficulty"].tolist(), ["medium"])
        self.assertEqual(students, ["DEMO-01", "DEMO-02"])

    def test_kazakh_template_sheet_and_column_names_are_supported(self):
        with patch("ui_helpers.urllib.request.urlopen", side_effect=self.responses(language="kk")) as request:
            bank, students = helpers.read_google_cards(self.URL)
        self.assertEqual(request.call_count, 4)
        self.assertEqual(bank["difficulty"].tolist(), ["medium"])
        self.assertIn("Алгоритм", bank.iloc[0]["question"])
        self.assertEqual(students, ["DEMO-01", "DEMO-02"])

    def test_wrong_host_scheme_or_document_path_is_rejected_before_network(self):
        urls = [
            "http://docs.google.com/spreadsheets/d/test/edit",
            "https://docs.google.com.evil.example/spreadsheets/d/test/edit",
            "https://evil.example/spreadsheets/d/test/edit",
            "https://docs.google.com/document/d/test/edit",
            "https://docs.google.com/spreadsheets/d/",
            "",
        ]
        with patch("ui_helpers.urllib.request.urlopen") as request:
            for url in urls:
                with self.subTest(url=url):
                    self.assert_error("google_url_invalid", helpers.read_google_cards, url)
            request.assert_not_called()

    def test_network_error_is_sanitized(self):
        with patch("ui_helpers.urllib.request.urlopen", side_effect=urllib.error.URLError("private provider detail")):
            error = self.assert_error("google_load_failed", helpers.read_google_cards, self.URL)
        self.assertNotIn("private provider detail", str(error))

    def test_html_signin_page_is_not_treated_as_csv(self):
        with patch("ui_helpers.urllib.request.urlopen", return_value=io.BytesIO(b"  <!doctype html><html>Sign in</html>")):
            self.assert_error("google_load_failed", helpers.read_google_cards, self.URL)

    def test_oversize_response_is_rejected(self):
        with patch.object(helpers, "MAX_UPLOAD", 4), patch("ui_helpers.urllib.request.urlopen", return_value=io.BytesIO(b"12345")):
            self.assert_error("google_load_failed", helpers.read_google_cards, self.URL)

    def test_malformed_csv_is_rejected(self):
        with patch("ui_helpers.urllib.request.urlopen", return_value=io.BytesIO(b'question,answer\n"unterminated')):
            self.assert_error("google_load_failed", helpers.read_google_cards, self.URL)

    def test_empty_roster_and_missing_question_fields_are_rejected(self):
        with patch("ui_helpers.urllib.request.urlopen", side_effect=self.responses(students=pd.DataFrame({"ФИО": []}))):
            self.assert_error("google_load_failed", helpers.read_google_cards, self.URL)
        with patch("ui_helpers.urllib.request.urlopen", side_effect=self.responses(bank=pd.DataFrame({"unexpected": ["value"]}))):
            self.assert_error("google_load_failed", helpers.read_google_cards, self.URL)

    def test_first_column_is_not_silently_assumed_to_be_student_names(self):
        with patch("ui_helpers.urllib.request.urlopen", side_effect=self.responses(students=pd.DataFrame({"Question": ["Not a student"]}))):
            self.assert_error("google_load_failed", helpers.read_google_cards, self.URL)


if __name__ == "__main__":
    unittest.main()
