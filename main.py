# main.py
from __future__ import annotations  # позволяет писать типы "как строки", чтобы не было ошибок с порядком объявления классов

import csv  # библиотека для чтения CSV-файлов (таблички с запятыми)
import hashlib  # библиотека для подсчёта хэша MD5
import io  # работа с данными "в памяти" как с файлом
import json  # сохранение и чтение файла snapshot.json
import os  # доступ к переменным окружения (там лежит строка подключения к БД)
import sqlite3  # встроенная база данных (читаем список таблиц)
import sys  # настройка вывода в консоль; аргументы командной строки
import zipfile  # чтение zip-архивов
from dataclasses import dataclass, asdict, field  # dataclass - "умный" класс, который сам делает __init__
from datetime import datetime, timezone  # работа с датами и временем
from pathlib import Path  # удобная работа с путями к файлам
from typing import Iterator, Optional  # подсказки для типов: "или значение, или None", "эта функция отдаёт значения по одному"

import chardet  # угадывает кодировку текстового файла
import pymupdf as fitz  # PyMuPDF: вытаскивает текст из PDF 
import psycopg2  # драйвер для подключения к PostgreSQL
import py7zr  # чтение 7z-архивов
import rarfile  # чтение rar-архивов
import yaml  # чтение YAML-файлов
from bs4 import BeautifulSoup  # вытаскивает текст из HTML
from docx import Document  # читает .docx
from hachoir.metadata import extractMetadata  # универсальный "вытаскиватель" метаданных (если формат незнакомый)
from hachoir.parser import createParser  # создаёт парсер для hachoir (для бинарных файлов)
from mutagen import File as MutagenFile  # читает теги из аудио
from odf import teletype  # вытаскивает текст из ODF-документов
from odf.opendocument import load as load_odf  # открывает ODF-документы
from openpyxl import load_workbook  # читает .xlsx
from PIL import Image  # читает картинки
from pptx import Presentation  # читает .pptx

# ------------------------- КОНФИГ -------------------------
# Путь сканирования: приоритет - аргумент командной строки, потом переменная окружения,
# потом константа (на случай ручного запуска без interf).
_cli_root = None
if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):  # первый аргумент - не флаг, значит это путь
    _cli_root = sys.argv[1]  # запоминаем путь из командной строки

_env_root = os.environ.get("ROOT_PATH_OVERRIDE")  # путь, который мог передать interf через переменную окружения

ROOT_PATH = _cli_root or _env_root or r"E:\MAX_files"  # итоговый путь: аргумент > env > значение по умолчанию

ENABLE_DB_SYNC = True  # False = писать результаты скана в терминал; True = писать результаты скана в PostgreSQL
if os.environ.get("ENABLE_DB_SYNC_OVERRIDE"):  # interf может переопределить режим через переменную окружения
    ENABLE_DB_SYNC = os.environ["ENABLE_DB_SYNC_OVERRIDE"] == "1"  # "1" -> True, иначе False

DB_DSN = os.environ.get(  # строка подключения к БД: из переменной окружения DB_DSN или значение по умолчанию
    "DB_DSN",  # имя переменной окружения с параметрами подключения
    "host=localhost port=5432 dbname=file_storage user=postgres password=12345",  # значение по умолчанию: локальная база file_storage
)

# Снимок прошлого скана привязан к папке (по хэшу пути), чтобы сканы разных папок не смешивались.
# Лежит в %LOCALAPPDATA%\ScannerApp\ (на Windows) или ~/.local/share (на Linux).
_cache_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ScannerApp"  # папка для снимков
_cache_dir.mkdir(parents=True, exist_ok=True)  # создаём её, если ещё нет
_key = hashlib.md5(str(Path(ROOT_PATH).resolve()).encode("utf-8")).hexdigest()[:12]  # короткий хэш пути
SNAPSHOT_FILE = _cache_dir / f"snapshot_{_key}.json"  # файл-снимок именно этой папки

HASH_CHUNK_SIZE = 64 * 1024  # читаем файл кусками по 64 КБ, чтобы не грузить его целиком в память
MAX_CONTENT_SIZE = 5 * 1024 * 1024  # файлы больше 5 МБ не пытаемся читать для извлечения текста

# Только реально текстовые форматы. CSV/YAML/HTML обрабатываются отдельными ветками.
TEXT_EXTENSIONS = {".txt", ".py", ".json", ".md", ".log", ".xml"}  # расширения простых текстовых файлов

try:  # настраиваем консольный вывод, чтобы русский текст не превращался в кашу
    sys.stdout.reconfigure(encoding="utf-8")  # переключаем стандартный вывод в UTF-8
except Exception:  # если переключить не удалось (например, вывод перенаправлен) -
    pass  # ничего страшного: просто продолжаем работу

# ------------------------- СТАТУСЫ -------------------------
class FileStatus:  # возможные статусы файла при сравнении с прошлым сканом
    NEW = "new"  # файл встречен впервые
    UNCHANGED = "unchanged"  # хэш совпал со снимком - содержимое не менялось
    MODIFIED = "modified"  # хэш отличается - файл был изменён
    DELETED = "deleted"  # файла больше нет на диске (был в прошлом снимке)

    ALL = (NEW, UNCHANGED, MODIFIED, DELETED)  # все статусы одним кортежем - удобно перебирать

    RU = {  # русские подписи статусов для вывода в консоль
        NEW: "НОВЫЙ",
        UNCHANGED: "БЕЗ ИЗМЕНЕНИЙ",
        MODIFIED: "ИЗМЕНЁН",
        DELETED: "УДАЛЁН",
    }

# ------------------------- МОДЕЛЬ -------------------------
@dataclass  # декоратор dataclass: сам создаст __init__ и другие служебные методы
class FileRecord:  # одна запись о файле - всё, что программа знает о файле
    """
    Соответствие полям таблицы files в PostgreSQL:
        id          BIGSERIAL     (только в БД)
        path        TEXT          NOT NULL
        name        VARCHAR(255)  NOT NULL
        extension   VARCHAR(32)
        size_bytes  BIGINT        NOT NULL
        created_at  TIMESTAMPTZ   NOT NULL
        modified_at TIMESTAMPTZ   NOT NULL
        scanned_at  TIMESTAMPTZ   NOT NULL DEFAULT now()
        content     TEXT
        status      VARCHAR(16)   NOT NULL
        md5_hash    CHAR(32)
        change_count INTEGER
    """
    path: str  # полный путь к файлу
    name: str  # имя файла с расширением
    extension: str  # расширение в нижнем регистре
    size_bytes: int  # размер файла в байтах
    created_at: datetime  # дата создания файла
    modified_at: datetime  # дата последнего изменения файла
    status: str = FileStatus.NEW  # статус файла; по умолчанию - новый
    scanned_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))  # момент сканирования подставляется автоматически
    content: Optional[str] = None  # извлечённый текст; None - если достать не удалось
    md5_hash: Optional[str] = None  # контрольная сумма MD5
    file_id: Optional[int] = None  # id из таблицы files (появляется только после записи в БД)
    change_count: Optional[int] = None  # сколько раз менялся файл (счётчик из БД)

    def to_jsonable(self) -> dict:  # версия записи для сохранения в JSON
        d = asdict(self)  # превращаем запись в обычный словарь
        for k in ("created_at", "modified_at", "scanned_at"):  # даты в JSON не сохранить напрямую - переводим их в текст
            d[k] = d[k].isoformat()  # формат ISO 8601, например 2026-10-05T20:00:00+00:00
        return d  # возвращаем готовый словарь

    def to_row(self) -> dict:  # словарь значений для SQL-запроса (ключи = имена полей в запросе)
        return {  # начинаем собирать словарь
            "name": self.name,  # имя файла
            "extension": self.extension,  # расширение
            "size_bytes": self.size_bytes,  # размер в байтах
            "md5_hash": self.md5_hash,  # передаём MD5 в SQL-запрос
            "created_at": self.created_at,  # дата создания
            "modified_at": self.modified_at,  # дата изменения
            "scanned_at": self.scanned_at,  # момент сканирования
            "content": self.content,  # извлечённый текст
            "status": self.status,  # статус файла
        }

# ------------------------- ХРАНИЛИЩА -------------------------
class ConsoleStorage:  # хранилище №1: просто печатает результаты в консоль
    def save(self, record: FileRecord) -> None:  # сохранить запись (здесь - вывести на экран)
        created = record.created_at.astimezone().strftime("%Y-%m-%d %H:%M:%S")  # дата создания в местном времени
        modified = record.modified_at.astimezone().strftime("%Y-%m-%d %H:%M:%S")  # дата изменения в местном времени
        scanned = record.scanned_at.astimezone().strftime("%Y-%m-%d %H:%M:%S")  # момент сканирования в местном времени
        status_ru = FileStatus.RU.get(record.status, record.status)  # русская подпись статуса
        file_id_str = str(record.file_id) if record.file_id is not None else "-"  # id в консольном режиме отсутствует

        print(  # печатаем карточку файла
            f"[{status_ru}] {record.name}\n"  # строка 1: статус и имя файла
            f"    id:          {file_id_str}\n"  # id (в консольном режиме - прочерк)
            f"    status:      {record.status}\n"  # технический статус
            f"    путь:        {record.path}\n"  # полный путь
            f"    размер:      {record.size_bytes} B\n"  # размер в байтах
            f"    md5:         {record.md5_hash}\n"  # хэш MD5 - по нему отслеживаются изменения
            f"    создан:      {created}\n"  # дата создания
            f"    изменён:     {modified}\n"  # дата изменения
            f"    просканирован: {scanned}\n"  # момент сканирования
        )

    # Методы-«заглушки»: у консольного и БД-хранилища одинаковый набор методов,
    # поэтому main() работает с любым из них без проверок.
    def begin_scan(self, root: Path) -> None:  # «заглушка» начала скана
        pass

    def finish_scan(self, summary: dict) -> None:  # конец скана (в консольном режиме ничего делать не нужно)
        pass

    def rollback(self) -> None:  # «заглушка» - вызывается только при ошибке сканирования
        pass

    def close(self) -> None:  # «заглушка» - для единообразия с БД-хранилищем
        pass


class DatabaseStorage:  # хранилище №2: пишет результаты в PostgreSQL
    def __init__(self) -> None:  # при создании объекта сразу подключаемся к базе
        self.conn = psycopg2.connect(DB_DSN)  # открываем соединение с PostgreSQL по строке DB_DSN
        self.scan_id: Optional[int] = None  # id текущей строки в таблице scans
        self.root: Optional[Path] = None  # корневая папка текущего скана
        self.path_cache: dict[str, int] = {}  # «полный путь папки -> id в paths», чтобы не делать лишних запросов

    # ---- таблица scans: начало сканирования ----
    def begin_scan(self, root: Path) -> None:  # создаём новую строку в таблице scans (начало скана)
        self.root = root.resolve()  # запоминаем абсолютный путь корня
        with self.conn.cursor() as cur:  # курсор - через него выполняются все SQL-запросы
            cur.execute(  # вставляем запись о скане
                "INSERT INTO scans (root_path) VALUES (%s) RETURNING id",  # started_at заполнится сам
                (str(self.root),),  # значение подставляется безопасно вместо %s
            )
            self.scan_id = cur.fetchone()[0]  # RETURNING id вернул номер новой строки - запоминаем

    # ---- таблица paths: найти/создать папку и вернуть её id ----
    def get_path_id(self, directory: Path) -> int:  # ищем папку в БД или создаём её; возвращаем id
        key = str(directory)  # строковый вид пути
        cached = self.path_cache.get(key)  # смотрим, не создавали ли эту папку раньше
        if cached is not None:  # если id папки уже в кэше - возвращаем сразу
            return cached

        parent_id: Optional[int] = None  # у корневой папки родителя нет
        if directory != self.root and directory.parent != directory:  # для некорневых папок сначала нужен id родителя
            # Вся цепочка папок создаётся сверху вниз
            parent_id = self.get_path_id(directory.parent)  # рекурсивно находим/создаём родительскую папку

        with self.conn.cursor() as cur:
            cur.execute(  # вставляем папку; если такая уже есть - вернём её id
                """
                INSERT INTO paths (parent_id, name, full_path)
                VALUES (%s, %s, %s)
                ON CONFLICT (full_path) DO UPDATE SET name = EXCLUDED.name
                RETURNING id
                """,
                (parent_id, directory.name or key, key),  # у корня диска имя пустое -> берём полный путь
            )
            path_id = cur.fetchone()[0]  # id найденной или созданной папки

        self.path_cache[key] = path_id  # запоминаем в кэше
        return path_id

    # ---- таблицы paths + files: сохранить один файл ----
    def save(self, record: FileRecord) -> None:
        if record.status == FileStatus.DELETED:  # файл удалён с диска
            with self.conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE files
                       SET status = 'deleted', last_scan_id = %s, scanned_at = now()
                     WHERE name = %s
                       AND path_id = (SELECT id FROM paths WHERE full_path = %s)
                 RETURNING id, change_count
                    """,
                    # Строку не удаляем, а помечаем статусом 'deleted'. Файл ищем по имени
                    # и по папке (id папки находим подзапросом по full_path)
                    (self.scan_id, record.name, str(Path(record.path).parent)),
                )
                row = cur.fetchone()
                if row:  # если нашли - запомним id и change_count
                    record.file_id = row[0]
                    record.change_count = row[1]
            file_id_str = str(record.file_id) if record.file_id is not None else "-"
            print(f"[{FileStatus.RU[record.status]}] {record.name}  id={file_id_str}")
            return

        path_id = self.get_path_id(Path(record.path).parent)  # id папки файла (создаём, если её ещё нет)

        params = record.to_row()  # словарь значений из записи
        params["path_id"] = path_id  # добавляем id папки
        params["scan_id"] = self.scan_id  # добавляем id текущего сканирования
        if params["content"]:  # если есть извлечённый текст
            params["content"] = params["content"].replace("\x00", "")  # PostgreSQL не принимает символ NUL в тексте

        with self.conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO files
                    (path_id, name, extension, size_bytes, md5_hash,
                     created_at, modified_at, scanned_at, content,
                     status, first_scan_id, last_scan_id)
                VALUES
                    (%(path_id)s, %(name)s, %(extension)s, %(size_bytes)s, %(md5_hash)s,
                     %(created_at)s, %(modified_at)s,
                     %(scanned_at)s, %(content)s, %(status)s, %(scan_id)s, %(scan_id)s)
                ON CONFLICT (path_id, name) DO UPDATE SET
                    extension    = EXCLUDED.extension,
                    size_bytes   = EXCLUDED.size_bytes,
                    md5_hash     = EXCLUDED.md5_hash,
                    created_at   = EXCLUDED.created_at,
                    modified_at  = EXCLUDED.modified_at,
                    scanned_at   = EXCLUDED.scanned_at,
                    content      = EXCLUDED.content,
                    status       = EXCLUDED.status,
                    change_count = files.change_count + CASE
                        WHEN files.modified_at IS DISTINCT FROM EXCLUDED.modified_at THEN 1
                        ELSE 0
                    END,
                    last_scan_id = EXCLUDED.last_scan_id
                RETURNING id, change_count
                """,
                # Нового файла ещё нет -> вставляется строка, а first_scan_id
                # и last_scan_id = текущий скан. Файл уже есть -> обновляются его свойства и
                # last_scan_id; first_scan_id не трогаем.
                params,
            )
            file_id, change_count = cur.fetchone()
            record.file_id = file_id
            record.change_count = change_count

            json_record = record.to_jsonable()
            json_record["change_count"] = change_count
            json_record["file_id"] = file_id
            cur.execute(
                """
                INSERT INTO json (file_id, data, updated_at)
                VALUES (%s, %s::jsonb, now())
                ON CONFLICT (file_id) DO UPDATE
                SET data = EXCLUDED.data, updated_at = now()
                """,
                (file_id, json.dumps(json_record, ensure_ascii=False)),
            )

        print(f"[{FileStatus.RU.get(record.status, record.status)}] {record.name}  id={file_id}")

    # ---- таблица scans: конец сканирования ----
    def finish_scan(self, summary: dict) -> None:
        with self.conn.cursor() as cur:
            cur.execute(
                """
                UPDATE scans
                   SET finished_at     = now(),
                       files_total     = %s,
                       files_new       = %s,
                       files_unchanged = %s,
                       files_modified  = %s,
                       files_deleted   = %s
                 WHERE id = %s
                """,
                (
                    sum(summary.values()),  # всего файлов = сумма по всем статусам
                    summary[FileStatus.NEW],  # новых
                    summary[FileStatus.UNCHANGED],  # без изменений
                    summary[FileStatus.MODIFIED],  # изменённых
                    summary[FileStatus.DELETED],  # удалённых
                    self.scan_id,  # обновляем именно текущее сканирование
                ),
            )

    # При ошибке во время скана отменяем все изменения в БД и закрываем соединение.
    def rollback(self) -> None:
        self.conn.rollback()  # откат незавершённой транзакции
        self.conn.close()  # закрываем соединение

    def close(self) -> None:
        self.conn.commit()  # фиксируем всё, что записали за скан
        self.conn.close()  # закрываем соединение


storage = DatabaseStorage() if ENABLE_DB_SYNC else ConsoleStorage()  # выбираем хранилище один раз

# ------------------------- УТИЛИТЫ -------------------------
def compute_md5(path: Path) -> str:
    hasher = hashlib.md5()  # алгоритм MD5
    with path.open("rb") as f:  # открываем файл на чтение в бинарном режиме
        for chunk in iter(lambda: f.read(HASH_CHUNK_SIZE), b""):  # читаем кусками
            hasher.update(chunk)  # добавляем в хэш
    return hasher.hexdigest()  # возвращаем строку из 32 символов

# ------------------------- КОНВЕРТЕР СТАРЫХ ФОРМАТОВ -------------------------
def convert_office_file(path: Path) -> Optional[Path]:
    """Конвертирует старый файл Office в современный формат. Возвращает путь к копии или None."""
    ext = path.suffix.lower()  # расширение исходного файла (".doc" или ".xls")
    target = None  # сюда положим имя будущей конвертированной копии
    if ext == ".doc":  # документ Word старого формата
        target = path.with_suffix(".docx")  # копия рядом: то же имя, расширение .docx
    elif ext == ".xls":  # таблица Excel старого формата
        target = path.with_suffix(".xlsx")  # копия рядом: то же имя, расширение .xlsx
    else:
        return None  # другие расширения не конвертируем
    if target.exists():  # если конвертированная копия уже есть рядом
        return target  # повторно не конвертируем
    try:  # пробуем конвертировать через установленный Office
        import win32com.client  # pywin32: управляем установленным Word/Excel (COM)
        if ext == ".doc":  # ---- Word: .doc -> .docx ----
            app = win32com.client.Dispatch("Word.Application")  # запускаем Word
            app.Visible = False  # работаем в фоне
            app.DisplayAlerts = 0  # не задавать вопросов
            try:
                doc = app.Documents.Open(str(path))  # открываем старый документ
                doc.SaveAs2(str(target), 16)  # сохраняем как .docx
                doc.Close(False)  # закрываем документ, правки не сохраняем
            finally:
                app.Quit()  # закрываем Word в любом случае
        else:  # ---- Excel: .xls -> .xlsx ----
            app = win32com.client.Dispatch("Excel.Application")  # запускаем Excel
            app.Visible = False
            app.DisplayAlerts = False
            try:
                wb = app.Workbooks.Open(str(path))  # открываем старую книгу
                wb.SaveAs(str(target), 51)  # сохраняем как .xlsx
                wb.Close(False)  # закрываем книгу без сохранения
            finally:
                app.Quit()  # закрываем Excel в любом случае
        print(f"[КОНВЕРТАЦИЯ] {path.name} -> {target.name}")
        return target  # возвращаем путь к новой копии
    except Exception as e:  # Office недоступен или файл не открылся
        print(f"[CONVERT-SKIP] {path.name}: {type(e).__name__}: {e}")
        return None

 #Достаёт текст/метаданные из файла. None - если не умеем или слишком большой
def read_text_content(path: Path, size_bytes: int) -> Optional[str]:
    if size_bytes > MAX_CONTENT_SIZE:  # если файл больше 5 МБ
        return None  # не трогаем его - слишком тяжёлый

    ext = path.suffix.lower()  # берём расширение и приводим к нижнему регистру

    try:
        # ---- ТЕКСТОВЫЕ: определяем кодировку через chardet ----
        if ext in TEXT_EXTENSIONS:  # простой текстовый файл
            raw = path.read_bytes()  # читаем файл как набор байтов
            enc = chardet.detect(raw)["encoding"] or "utf-8"  # угадываем кодировку; если не смог - берём utf-8
            return raw.decode(enc, errors="replace").replace("\x00", "")  # байты -> текст; "битые" байты заменяем

        # ---- DOCX ----
        if ext == ".docx":  # документ Word
            doc = Document(str(path))
            return "\n".join(p.text for p in doc.paragraphs)  # текст всех абзацев

        # ---- XLSX ----
        if ext == ".xlsx":  # таблица Excel
            wb = load_workbook(str(path), read_only=True, data_only=True)  # экономный режим, только значения
            lines = []
            for ws in wb.worksheets:  # по всем листам
                for row in ws.iter_rows(values_only=True):  # по всем строкам
                    lines.append("\t".join("" if c is None else str(c) for c in row))  # пустые ячейки -> "", разделитель - таб
            return "\n".join(lines)

        # ---- PPTX ----
        if ext == ".pptx":  # презентация PowerPoint
            prs = Presentation(str(path))
            lines = []
            for i, slide in enumerate(prs.slides, 1):
                lines.append(f"--- Слайд {i} ---")
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        lines.append(shape.text_frame.text)
            return "\n".join(lines)

        # ---- PDF ----
        if ext == ".pdf":
            with fitz.open(str(path)) as doc:
                return "\n".join(page.get_text() for page in doc)  # текст со всех страниц

        # ---- CSV ----
        if ext == ".csv":
            raw = path.read_bytes()
            enc = chardet.detect(raw)["encoding"] or "utf-8"
            text = raw.decode(enc, errors="replace")
            rows = list(csv.reader(io.StringIO(text)))  # разбираем CSV построчно
            return "\n".join("\t".join(r) for r in rows)  # колонки через таб, строки через перевод строки

        # ---- YAML ----
        if ext in (".yaml", ".yml"):
            data = yaml.safe_load(path.read_text(encoding="utf-8", errors="replace"))  # безопасный парсинг
            return yaml.safe_dump(data, allow_unicode=True)  # обратно в строку, кириллица читаемая

        # ---- HTML ----
        if ext in (".html", ".htm"):
            soup = BeautifulSoup(path.read_text(encoding="utf-8", errors="replace"), "html.parser")
            return soup.get_text(separator="\n", strip=True)  # только текст без тегов

        # ---- ODF (odt/ods/odp) ----
        if ext in (".odt", ".ods", ".odp"):
            doc = load_odf(str(path))
            return teletype.extractText(doc)

        # ---- ИЗОБРАЖЕНИЯ ----
        if ext in (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"):
            with Image.open(path) as img:
                return f"Изображение {img.format} {img.width}x{img.height} mode={img.mode}"

        # ---- АУДИО ----
        if ext in (".mp3", ".flac", ".ogg", ".m4a", ".wav", ".wma"):
            audio = MutagenFile(str(path), easy=True)  # читаем теги
            if audio is None:
                return None
            return "\n".join(f"{k}: {v}" for k, v in audio.items())  # теги в виде «ключ: значение»

        # ---- АРХИВЫ ----
        if ext == ".zip":
            with zipfile.ZipFile(path) as z:
                return "\n".join(z.namelist())
        if ext == ".7z":
            with py7zr.SevenZipFile(path) as z:
                return "\n".join(z.getnames())
        if ext == ".rar":
            with rarfile.RarFile(path) as r:
                return "\n".join(r.namelist())

        # ---- SQL ----
        if ext in (".sqlite", ".db"):
            con = sqlite3.connect(str(path))
            try:
                cur = con.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
                return "\n".join(t[0] for t in cur.fetchall())  # список таблиц
            finally:
                con.close()

        # ---- СТАРЫЕ ФОРМАТЫ OFFICE: .doc/.xls сначала конвертируем ----
        if ext in (".doc", ".xls"):
            converted = convert_office_file(path)
            if converted is not None and converted.suffix.lower() == ".docx":
                doc = Document(str(converted))
                return "\n".join(p.text for p in doc.paragraphs)
            if converted is not None and converted.suffix.lower() == ".xlsx":
                wb = load_workbook(str(converted), read_only=True, data_only=True)
                lines = []
                for ws in wb.worksheets:
                    for row in ws.iter_rows(values_only=True):
                        lines.append("\t".join("" if c is None else str(c) for c in row))
                return "\n".join(lines)

        # ---- OLE (старый формат .ppt) ----
        if ext == ".ppt":
            import olefile
            if olefile.isOleFile(str(path)):
                ole = olefile.OleFileIO(str(path))
                return "\n".join("/".join(s) for s in ole.listdir())  # список внутренних потоков

        # ---- ХАОШИР: универсальные метаданные (только для форматов, где мы не пробовали специальный парсер) ----
        if ext not in (".doc", ".xls"):
            parser = createParser(str(path))
            if parser:
                meta = extractMetadata(parser)
                if meta:
                    return str(meta.exportDictionary())

    except Exception as e:  # если что-то сломалось при чтении
        print(f"[CONTENT-SKIP] {path.name}: {type(e).__name__}: {e}")
        return None

    return None  # формат незнакомый - ничего не извлекли


def build_record(path: Path) -> Optional[FileRecord]:  # собираем запись о файле
    try:
        st = path.stat()  # берём информацию о файле (размер, даты)
        return FileRecord(
            path=str(path.resolve()),  # полный путь
            name=path.name,  # имя файла
            extension=path.suffix.lower(),  # расширение
            size_bytes=st.st_size,  # размер в байтах
            md5_hash=compute_md5(path),  # считаем MD5
            created_at=datetime.fromtimestamp(st.st_ctime, tz=timezone.utc),  # дата создания
            modified_at=datetime.fromtimestamp(st.st_mtime, tz=timezone.utc),  # дата изменения
            content=read_text_content(path, st.st_size),  # пытаемся достать содержимое
            status=FileStatus.NEW,  # пока помечаем как "новый", потом уточним
        )
    except OSError as e:  # не получилось прочитать файл
        print(f"[SKIP] {path} -> {type(e).__name__}: {e}")
        return None

# ------------------------- СНИМОК -------------------------
def load_snapshot() -> dict[str, str]:  # читаем снимок прошлого скана
    if not SNAPSHOT_FILE.exists():  # файла нет (первый запуск для этой папки)
        return {}  # пустой словарь - все файлы будут "новыми"
    try:
        data = json.loads(SNAPSHOT_FILE.read_text(encoding="utf-8"))  # читаем JSON
        return {rec["path"]: rec["md5_hash"] for rec in data}  # путь -> хэш MD5
    except (json.JSONDecodeError, KeyError, OSError) as e:  # файл битый или не читается
        print(f"[WARN] Не удалось прочитать снимок: {e}")
        return {}  # считаем, что снимка нет


def save_snapshot(records: list[FileRecord]) -> None:  # сохраняем снимок текущего скана
    alive = [r for r in records if r.status != FileStatus.DELETED]  # удалённые файлы в снимок не пишем
    SNAPSHOT_FILE.write_text(
        json.dumps([r.to_jsonable() for r in alive], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

# ------------------------- СКАНЕР -------------------------
def scan_directory(root: Path) -> Iterator[FileRecord]:  # обходит папку рекурсивно и отдаёт записи (генератор)
    for path in root.rglob("*"):  # рекурсивно все файлы и папки внутри root
        if not path.is_file():  # если это не файл (а папка)
            continue  # пропускаем
        record = build_record(path)  # строим запись о файле
        if record is not None:
            yield record  # отдаём её наружу

# ------------------------- ОСНОВНАЯ ЛОГИКА -------------------------
def run_scan(root: Path) -> list[FileRecord]:  # основной проход
    previous = load_snapshot()  # прошлый снимок: path -> MD5
    seen_paths: set[str] = set()  # пути, которые встретили в этом скане
    results: list[FileRecord] = []  # сюда собираем все записи

    for record in scan_directory(root):  # идём по всем файлам
        seen_paths.add(record.path)  # помечаем, что этот путь видели

        old_hash = previous.get(record.path)  # ищем файл в прошлом снимке
        if old_hash is None:  # если его там нет
            record.status = FileStatus.NEW  # значит, файл новый
        elif old_hash == record.md5_hash:  # если MD5 совпал
            record.status = FileStatus.UNCHANGED  # файл не менялся
        else:  # хэш другой
            record.status = FileStatus.MODIFIED  # файл изменился

        storage.save(record)  # сохраняем запись (в БД или консоль)
        results.append(record)

    for old_path, old_hash in previous.items():  # теперь ищем удалённые файлы
        if old_path in seen_paths:  # если этот путь уже видели
            continue  # пропускаем
        deleted = FileRecord(  # файла больше нет - делаем запись-маркер
            path=old_path,
            name=Path(old_path).name,
            extension=Path(old_path).suffix.lower(),
            size_bytes=0,
            md5_hash=old_hash,  # оставляем старый MD5 из снимка
            created_at=datetime.now(timezone.utc),
            modified_at=datetime.now(timezone.utc),
            status=FileStatus.DELETED,
        )
        storage.save(deleted)
        results.append(deleted)

    return results


def main() -> None:  # главная функция
    root = Path(ROOT_PATH)  # папка для сканирования
    print(f"[СТАРТ] Сканирую папку: {root}")  # видно в консоли и в статусе interf
    print(f"[СНИМОК] {SNAPSHOT_FILE}")  # какой файл снимка используется

    if not root.exists():  # если папки нет
        print(f"Путь не существует: {root}")
        return

    try:
        storage.begin_scan(root)  # создаём запись о новом скане
        results = run_scan(root)  # запускаем скан

        summary = {s: 0 for s in FileStatus.ALL}  # счётчики по статусам
        for r in results:
            summary[r.status] = summary.get(r.status, 0) + 1

        storage.finish_scan(summary)  # закрываем скан: время окончания и счётчики
        storage.close()  # commit + закрытие соединения
    except Exception:
        storage.rollback()  # откатываем всё, что записали
        raise

    save_snapshot(results)  # сохраняем снимок - пригодится в следующий раз

    print("=" * 60)
    print(f"Всего обработано:     {len(results)}")
    print(f"  новых:              {summary[FileStatus.NEW]}")
    print(f"  без изменений:      {summary[FileStatus.UNCHANGED]}")
    print(f"  изменено:           {summary[FileStatus.MODIFIED]}")
    print(f"  удалено:            {summary[FileStatus.DELETED]}")
    print("=" * 60)


if __name__ == "__main__":  # если файл запущен напрямую
    main()
