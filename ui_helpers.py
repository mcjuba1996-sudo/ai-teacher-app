"""Small UI-independent helpers for uploaded files and readable errors."""
import hashlib
import io
import json
import re
import urllib.parse
import urllib.request

import pandas as pd
from docx import Document
from PIL import Image, ImageOps

import bilim_core as core

MAX_UPLOAD = 10 * 1024 * 1024
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def fingerprint(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, default=str).encode()).hexdigest()


def upload_bytes(upload):
    if upload is None:
        raise core.ValidationError("file_missing")
    data = upload.getvalue()
    if not data or len(data) > MAX_UPLOAD:
        raise core.ValidationError("file_size")
    return data


def text_document(title, text):
    doc = Document()
    doc.add_heading(title, 0)
    for line in text.splitlines():
        doc.add_paragraph(line)
    result = io.BytesIO()
    doc.save(result)
    return result.getvalue()


def image_attachment(upload):
    raw = upload_bytes(upload)
    try:
        with Image.open(io.BytesIO(raw)) as original:
            original.load()
            picture = ImageOps.exif_transpose(original).convert("RGB")
            picture.thumbnail((2048, 2048))
            output = io.BytesIO()
            picture.save(output, format="JPEG", quality=90)
        return ("image/jpeg", output.getvalue())
    except (OSError, ValueError, Image.DecompressionBombError):
        raise core.ValidationError("image_invalid") from None


def read_google_cards(url):
    parsed = urllib.parse.urlparse(url.strip())
    match = re.match(r"^/spreadsheets/d/([A-Za-z0-9_-]+)(?:/|$)", parsed.path)
    if parsed.scheme != "https" or parsed.hostname != "docs.google.com" or not match:
        raise core.ValidationError("google_url_invalid")
    def read_sheet(sheet):
        target = f"https://docs.google.com/spreadsheets/d/{match[1]}/gviz/tq?" + urllib.parse.urlencode({"tqx": "out:csv", "sheet": sheet})
        try:
            with urllib.request.urlopen(target, timeout=15) as response:
                raw = response.read(MAX_UPLOAD + 1)
            if len(raw) > MAX_UPLOAD or raw.lstrip().startswith(b"<"):
                raise ValueError()
            return pd.read_csv(io.BytesIO(raw))
        except Exception:
            raise core.ValidationError("google_load_failed") from None
    bank = students = None
    for sheet in ("Банк_вопросов", "Сұрақтар_банкі"):
        try:
            bank = core.normalize_bank(read_sheet(sheet))
            break
        except core.ValidationError:
            continue
    aliases = {"фио", "атыжөні", "атә", "оқушы", "student", "name"}
    for sheet in ("Ученики", "Оқушылар"):
        try:
            frame = read_sheet(sheet)
            column = next((c for c in frame.columns if re.sub(r"[\s_.\-]+", "", str(c).casefold()) in aliases), None)
            if column is not None:
                students = core.parse_students("\n".join(frame[column].dropna().astype(str)))
                break
        except core.ValidationError:
            continue
    if bank is None or students is None:
        raise core.ValidationError("google_load_failed")
    return bank, students


ERRORS = {
    "missing_key": ("Для этой функции нужен Gemini API-ключ. Введите его в боковой панели.", "Бұл функцияға Gemini API кілті қажет. Оны бүйірлік панельге енгізіңіз."),
    "quota": ("Лимит ИИ исчерпан. Попробуйте позже или используйте свой ключ.", "ЖИ лимиті аяқталды. Кейін қайталаңыз немесе өз кілтіңізді қолданыңыз."),
    "auth": ("Сервис не принял ключ. Проверьте ключ и его доступ.", "Сервис кілтті қабылдамады. Кілтті және қолжетімділікті тексеріңіз."),
    "timeout": ("ИИ не ответил вовремя. Повторите запрос или сократите материал.", "ЖИ уақытында жауап бермеді. Қайталаңыз немесе материалды қысқартыңыз."),
    "unavailable": ("Модель ИИ недоступна. Попробуйте позже; владельцу стоит проверить GEMINI_MODEL.", "ЖИ моделі қолжетімсіз. Кейін қайталаңыз; қосымша иесі GEMINI_MODEL баптауын тексере алады."),
    "blocked": ("Сервис не смог обработать материал. Измените запрос или выберите другой учебный пример.", "Сервис материалды өңдей алмады. Сұрауды өзгертіңіз немесе басқа оқу мысалын таңдаңыз."),
    "empty": ("ИИ вернул пустой ответ. Повторите генерацию.", "ЖИ бос жауап қайтарды. Қайта жасаңыз."),
    "invalid_response": ("Ответ ИИ неполный или имеет неверный формат. Документ не создан; повторите запрос.", "ЖИ жауабы толық емес немесе пішімі қате. Құжат жасалмады; қайталаңыз."),
    "request": ("Запрос ИИ не выполнен. Проверьте соединение и повторите попытку.", "ЖИ сұрауы орындалмады. Байланысты тексеріп, қайталаңыз."),
    "json_invalid": ("Ответ ИИ не прошёл проверку формата. Повторите генерацию.", "ЖИ жауабының пішімі тексеруден өтпеді. Қайта жасаңыз."),
    "field_required": ("Заполните все обязательные поля содержательным текстом.", "Барлық міндетті өрістерді мазмұнды мәтінмен толтырыңыз."),
    "students_empty": ("Добавьте хотя бы одного ученика с учебным обозначением.", "Кемінде бір оқушының шартты атауын енгізіңіз."),
    "bank_empty": ("Загрузите непустой банк вопросов.", "Бос емес сұрақтар қорын жүктеңіз."),
    "bank_columns_missing": ("Не хватает столбцов вопроса, ответа или сложности. Используйте шаблон.", "Сұрақ, жауап немесе қиындық бағаны жоқ. Үлгіні пайдаланыңыз."),
    "bank_cell_empty": ("В банке есть пустой вопрос, ответ или уровень сложности.", "Қорда бос сұрақ, жауап немесе қиындық деңгейі бар."),
    "bank_cell_invalid": ("Вопрос, ответ и сложность должны быть текстом, а не списком или объектом.", "Сұрақ, жауап және қиындық тізім не объект емес, мәтін болуы тиіс."),
    "bank_difficulty_invalid": ("Используйте уровни Легкий/Средний/Сложный или Жеңіл/Орташа/Қиын.", "Жеңіл/Орташа/Қиын немесе Легкий/Средний/Сложный деңгейлерін қолданыңыз."),
    "bank_question_conflict": ("Один вопрос указан с разными ответами или сложностью. Исправьте дубликаты.", "Бір сұрақтың жауаптары немесе деңгейлері әртүрлі. Қайталанған жолдарды түзетіңіз."),
    "bank_duplicate_column": ("В таблице повторяются обязательные столбцы. Используйте шаблон.", "Кестеде міндетті бағандар қайталанған. Үлгіні пайдаланыңыз."),
    "bank_insufficient": ("Недостаточно вопросов выбранной сложности. Уменьшите количество или дополните банк.", "Таңдалған деңгейдегі сұрақтар жеткіліксіз. Санды азайтыңыз немесе қорды толықтырыңыз."),
    "counts_empty": ("Выберите хотя бы одно задание.", "Кемінде бір тапсырма таңдаңыз."),
    "counts_invalid": ("Количество заданий должно быть целым и неотрицательным.", "Тапсырмалар саны бүтін және теріс емес болуы тиіс."),
    "workbook_invalid": ("Excel не прочитан. Используйте шаблон и сохраните его как .xlsx.", "Excel оқылмады. Үлгіні қолданып, .xlsx түрінде сақтаңыз."),
    "workbook_sheets_missing": ("В Excel нужны два листа: банк вопросов и ученики. Скачайте шаблон.", "Excel ішінде сұрақтар қоры және оқушылар парақтары керек. Үлгіні жүктеңіз."),
    "students_column_missing": ("На листе учеников нет столбца с именами/учебными обозначениями.", "Оқушылар парағында аты-жөні/шартты атаулары бар баған жоқ."),
    "ksp_invalid": ("ИИ вернул неполный КСП. Документ не создан; повторите генерацию.", "ЖИ толық емес ҚМЖ қайтарды. Құжат жасалмады; қайта жасаңыз."),
    "ksp_stages_invalid": ("В плане должны быть три содержательных этапа. Повторите генерацию.", "Жоспарда мазмұнды үш кезең болуы тиіс. Қайта жасаңыз."),
    "ktp_invalid": ("ИИ вернул неполный КТП. Повторите генерацию.", "ЖИ толық емес КТЖ қайтарды. Қайта жасаңыз."),
    "ktp_count_invalid": ("Количество уроков в ответе ИИ не совпало с заданным. Повторите генерацию.", "ЖИ жауабындағы сабақ саны берілген санға сәйкес емес. Қайта жасаңыз."),
    "ktp_numbering_invalid": ("ИИ нарушил нумерацию уроков. Повторите генерацию.", "ЖИ сабақтардың нөмірленуін бұзды. Қайта жасаңыз."),
    "ktp_quarters_invalid": ("ИИ неверно распределил уроки по четвертям. Повторите генерацию.", "ЖИ сабақтарды тоқсандарға қате бөлді. Қайта жасаңыз."),
    "file_missing": ("Сначала выберите файл.", "Алдымен файлды таңдаңыз."),
    "file_size": ("Файл должен быть непустым и не больше 10 МБ.", "Файл бос болмауы және 10 МБ-тан аспауы тиіс."),
    "pdf_invalid": ("Файл не соответствует формату PDF.", "Файл PDF пішіміне сәйкес емес."),
    "image_invalid": ("Изображение не прочитано. Загрузите исправный JPG или PNG.", "Сурет оқылмады. Жарамды JPG немесе PNG жүктеңіз."),
    "google_url_invalid": ("Нужна ссылка https://docs.google.com/spreadsheets/d/… на учебную таблицу.", "Оқу кестесінің https://docs.google.com/spreadsheets/d/… сілтемесі қажет."),
    "google_load_failed": ("Таблица не прочитана. Проверьте доступ по ссылке и названия листов/столбцов по русскому или казахскому шаблону.", "Кесте оқылмады. Сілтеме арқылы қолжетімділікті және парақтар/бағандар атауларын қазақша не орысша үлгімен салыстырыңыз."),
}
