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
    def save(self, record: FileRecord) -> None:  # объявляем метод save класса DatabaseStorage: принимает одну запись FileRecord, ничего не возвращает
        if record.status == FileStatus.DELETED:  # проверяем: если у записи статус "удалён" - идём по особой ветке
            with self.conn.cursor() as cur:  # открываем курсор БД через with; в конце блока он сам закроется
                cur.execute(  # выполняем SQL-запрос (текст ниже)
                    """
                    UPDATE files
                       SET status = 'deleted', last_scan_id = %s, scanned_at = now()
                     WHERE name = %s
                       AND path_id = (SELECT id FROM paths WHERE full_path = %s)
                 RETURNING id, change_count
                    """,
                    # Строку не удаляем, а помечаем статусом 'deleted'. Файл ищем по имени
                    # и по папке (id папки находим подзапросом по full_path)
                    (self.scan_id, record.name, str(Path(record.path).parent)),  # подставляем значения вместо %s: id текущего скана, имя файла, путь к его папке
                )  # закрываем вызов cur.execute
                row = cur.fetchone()  # забираем одну строку результата (или None, если ничего не нашли)
                if row:  # если строка есть - значит, файл в БД нашёлся
                    record.file_id = row[0]  # кладём в запись id файла из БД (первая колонка RETURNING)
                    record.change_count = row[1]  # кладём в запись счётчик изменений (вторая колонка RETURNING)
            file_id_str = str(record.file_id) if record.file_id is not None else "-"  # готовим id для печати: если id есть - строка, иначе прочерк
            print(f"[{FileStatus.RU[record.status]}] {record.name}  id={file_id_str}")  # печатаем в консоль: русский статус, имя, id
            return  # выходим из метода - ветка для удалённых закончена

        path_id = self.get_path_id(Path(record.path).parent)  # id папки файла (создаём, если её ещё нет)

        params = record.to_row()  # словарь значений из записи
        params["path_id"] = path_id  # добавляем id папки
        params["scan_id"] = self.scan_id  # добавляем id текущего сканирования
        if params["content"]:  # если есть извлечённый текст
            params["content"] = params["content"].replace("\x00", "")  # PostgreSQL не принимает символ NUL в тексте

        with self.conn.cursor() as cur:  # открываем курсор БД через with - в конце блока он сам закроется
            cur.execute(  # выполняем SQL-запрос (текст ниже)
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
                params,  # передаём словарь значений - psycopg2 сам подставит их вместо %(...)s
            )  # закрываем вызов cur.execute
            file_id, change_count = cur.fetchone()  # забираем из RETURNING две величины: id файла и счётчик изменений
            record.file_id = file_id  # кладём id файла в запись
            record.change_count = change_count  # кладём счётчик изменений в запись

            json_record = record.to_jsonable()  # превращаем запись в словарь, готовый для JSON
            json_record["change_count"] = change_count  # добавляем в словарь счётчик изменений
            json_record["file_id"] = file_id  # добавляем в словарь id файла
            cur.execute(  # выполняем второй SQL-запрос - пишем JSON-версию записи в таблицу json
                """
                INSERT INTO json (file_id, data, updated_at)
                VALUES (%s, %s::jsonb, now())
                ON CONFLICT (file_id) DO UPDATE
                SET data = EXCLUDED.data, updated_at = now()
                """,
                (file_id, json.dumps(json_record, ensure_ascii=False)),  # подставляем id файла и сам JSON (кириллица как есть)
            )  # закрываем вызов cur.execute

        print(f"[{FileStatus.RU.get(record.status, record.status)}] {record.name}  id={file_id}")  # печатаем в консоль: русский статус, имя, id
        
    # ---- таблица scans: конец сканирования ----
    def finish_scan(self, summary: dict) -> None:  # объявляем метод finish_scan: принимает словарь счётчиков, ничего не возвращает
        with self.conn.cursor() as cur:  # открываем курсор БД через with - в конце блока он сам закроется
            cur.execute(  # выполняем SQL-запрос - обновляем строку текущего скана (текст ниже)
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
                (  # начинаем кортеж значений, которые подставятся вместо %s
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


storage = DatabaseStorage() if ENABLE_DB_SYNC else ConsoleStorage()  # выбираем хранилище: если синхронизация с БД включена - DatabaseStorage, иначе ConsoleStorage

# ------------------------- УТИЛИТЫ -------------------------
def compute_md5(path: Path) -> str:  # объявляем функцию compute_md5: принимает путь, возвращает строку MD5
    hasher = hashlib.md5()  # создаём объект хэшера MD5 - в него будем подкладывать данные
    with path.open("rb") as f:  # открываем файл в бинарном режиме; with автоматически закроет его в конце
        for chunk in iter(lambda: f.read(HASH_CHUNK_SIZE), b""):  # в цикле читаем файл кусками по HASH_CHUNK_SIZE; когда f.read вернёт пустые байты - цикл остановится
            hasher.update(chunk)  # подкладываем очередной кусок в хэшер
    return hasher.hexdigest()  # возвращаем итоговый MD5 в виде строки из 32 хэш-символов

# ------------------------- КОНВЕРТЕР СТАРЫХ ФОРМАТОВ ------------------------- 
def convert_office_file(path: Path) -> Optional[Path]:  # объявляем функцию: принимает путь к старому файлу, возвращает путь к копии или None
    """Конвертирует старый файл Office в современный формат. Возвращает путь к копии или None."""  # текстовое описание функции
    ext = path.suffix.lower()  # берём расширение файла и делаем его маленькими буквами 
    target = None  # переменная для пути будущей конвертированной копии; пока пусто (None)
    if ext == ".doc":  # если это старый Word-документ
        target = path.with_suffix(".docx")  # путь копии: то же имя, но расширение .docx
    elif ext == ".xls":  # иначе если это старая Excel-таблица
        target = path.with_suffix(".xlsx")  # путь копии: то же имя, но расширение .xlsx
    else:  # если расширение не .doc и не .xls
        return None  # конвертировать нечего - выходим, вернув None
    if target.exists():  # если рядом уже лежит конвертированная копия
        return target  # второй раз не конвертируем - сразу отдаём её путь
    try:
        import win32com.client  # подключаем pywin32 - через него управляем Word/Excel
        if ext == ".doc":  # ---- ветка Word: .doc -> .docx ----
            app = win32com.client.Dispatch("Word.Application")  # запускаем Word
            app.Visible = False  # скрываем окно Word - работаем в фоне
            app.DisplayAlerts = 0  # отключаем всплывающие окошки-предупреждения
            try:  # внутренний try, чтобы Word точно закрылся
                doc = app.Documents.Open(str(path))  # открываем старый .doc
                doc.SaveAs2(str(target), 16)  # сохраняем как .docx
                doc.Close(False)  # закрываем документ без сохранения правок
            finally:  # этот блок выполнится в любом случае - даже если была ошибка
                app.Quit()  # закрываем Word
        else:  # ---- ветка Excel: .xls -> .xlsx ----
            app = win32com.client.Dispatch("Excel.Application")  # запускаем Excel
            app.Visible = False  # скрываем окно Excel
            app.DisplayAlerts = False  # отключаем всплывающие окошки
            try:
                wb = app.Workbooks.Open(str(path))  # открываем старую книгу .xls
                wb.SaveAs(str(target), 51)  # сохраняем как .xlsx
                wb.Close(False)  # закрываем книгу без сохранения правок
            finally:  # выполнится в любом случае
                app.Quit()  # закрываем Excel
        print(f"[КОНВЕРТАЦИЯ] {path.name} -> {target.name}")  # пишем в консоль, что конвертация прошла
        return target  # отдаём путь к новой копии
    except Exception as e:  # если что-то пошло не так
        print(f"[CONVERT-SKIP] {path.name}: {type(e).__name__}: {e}")  # сообщаем, что конвертацию пропустили, и почему
        return None  # отдаём None - не получилось
        
 #Достаёт текст/метаданные из файла. None - если не умеем или слишком большой 
def read_text_content(path: Path, size_bytes: int) -> Optional[str]:  # объявляем функцию: путь + размер, возвращает текст или None
    if size_bytes > MAX_CONTENT_SIZE:  # если файл больше 5 МБ
        return None  # не трогаем его - слишком тяжёлый, сразу выходим

    ext = path.suffix.lower()  # берём расширение файла маленькими буквами - пригодится для проверок ниже

    try:
        # ---- ТЕКСТОВЫЕ: определяем кодировку через chardet ----
        if ext in TEXT_EXTENSIONS:  # если расширение входит в набор простых текстовых
            raw = path.read_bytes()  # читаем файл целиком в байтах
            enc = chardet.detect(raw)["encoding"] or "utf-8"  # угадываем кодировку; если chardet не справился - берём utf-8
            return raw.decode(enc, errors="replace").replace("\x00", "")  # превращаем байты в текст (битые байты заменяем), убираем NUL-символы

        # ---- DOCX ----
        if ext == ".docx":  # если это .docx
            doc = Document(str(path))  # открываем документ через python-docx
            return "\n".join(p.text for p in doc.paragraphs)  # берём текст каждого абзаца и склеиваем через перевод строки

        # ---- XLSX ----
        if ext == ".xlsx":  # если это .xlsx
            wb = load_workbook(str(path), read_only=True, data_only=True)  # открываем книгу в экономном режиме и только значения
            lines = []  # сюда будем складывать строки текста
            for ws in wb.worksheets:  # перебираем все листы книги
                for row in ws.iter_rows(values_only=True):  # перебираем строки листа, получая только значения ячеек
                    lines.append("\t".join("" if c is None else str(c) for c in row))  # ячейки строки склеиваем через таб; пустые -> ""
            return "\n".join(lines)  # все строки склеиваем через перевод строки и отдаём

        # ---- PPTX ---- 
        if ext == ".pptx":  # если это .pptx
            prs = Presentation(str(path))  # открываем презентацию через python-pptx
            lines = []  # сюда будем складывать строки текста
            for i, slide in enumerate(prs.slides, 1):  # перебираем слайды, нумеруя с 1
                lines.append(f"--- Слайд {i} ---")  # добавляем строку-разделитель с номером слайда
                for shape in slide.shapes:  # перебираем все фигуры на слайде
                    if shape.has_text_frame:  # если у фигуры есть текстовый блок
                        lines.append(shape.text_frame.text)  # добавляем текст из этого блока
            return "\n".join(lines)  # склеиваем строки через перевод строки

        # ---- PDF ----
        if ext == ".pdf":  # если это .pdf
            with fitz.open(str(path)) as doc:  # открываем PDF через PyMuPDF; with сам закроет
                return "\n".join(page.get_text() for page in doc)  # берём текст со всех страниц и склеиваем

        # ---- CSV ----
        if ext == ".csv":  # если это .csv
            raw = path.read_bytes()  # читаем файл в байтах
            enc = chardet.detect(raw)["encoding"] or "utf-8"  # угадываем кодировку; fallback - utf-8
            text = raw.decode(enc, errors="replace")  # байты -> текст
            rows = list(csv.reader(io.StringIO(text)))  # парсим CSV построчно через StringIO
            return "\n".join("\t".join(r) for r in rows)  # колонки через таб, строки через перевод строки

        # ---- YAML ----
        if ext in (".yaml", ".yml"):  # если это YAML
            data = yaml.safe_load(path.read_text(encoding="utf-8", errors="replace"))  # безопасно парсим YAML
            return yaml.safe_dump(data, allow_unicode=True)  # обратно в строку, кириллица читаемая

        # ---- HTML ----
        if ext in (".html", ".htm"):  # если это HTML
            soup = BeautifulSoup(path.read_text(encoding="utf-8", errors="replace"), "html.parser")  # парсим HTML
            return soup.get_text(separator="\n", strip=True)  # берём только текст без тегов

        # ---- ODF (odt/ods/odp) ----
        if ext in (".odt", ".ods", ".odp"):  # если это ODF-документ
            doc = load_odf(str(path))  # открываем ODF-документ
            return teletype.extractText(doc)  # извлекаем из него текст

        # ---- ИЗОБРАЖЕНИЯ ----
        if ext in (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"):  # если это изображение
            with Image.open(path) as img:  # открываем картинку через PIL
                return f"Изображение {img.format} {img.width}x{img.height} mode={img.mode}"  # отдаём строку с форматом и размерами

        # ---- АУДИО ----
        if ext in (".mp3", ".flac", ".ogg", ".m4a", ".wav", ".wma"):  # если это аудио
            audio = MutagenFile(str(path), easy=True)  # читаем теги через mutagen в упрощённом режиме
            if audio is None:  # если теги прочитать не удалось
                return None  # отдаём None
            return "\n".join(f"{k}: {v}" for k, v in audio.items())  # теги в виде «ключ: значение», по одному на строку

        # ---- АРХИВЫ ----
        if ext == ".zip":  # если это ZIP
            with zipfile.ZipFile(path) as z:  # открываем архив
                return "\n".join(z.namelist())  # список имён внутри архива
        if ext == ".7z":  # если это 7z
            with py7zr.SevenZipFile(path) as z:  # открываем архив
                return "\n".join(z.getnames())  # список имён внутри
        if ext == ".rar":  # если это RAR
            with rarfile.RarFile(path) as r:  # открываем архив
                return "\n".join(r.namelist())  # список имён внутри

        # ---- SQL ---- 
        if ext in (".sqlite", ".db"):  # если это SQLite-база
            con = sqlite3.connect(str(path))  # подключаемся к базе
            try:
                cur = con.cursor()  # создаём курсор
                cur.execute("SELECT name FROM sqlite_master WHERE type='table'")  # получаем список таблиц
                return "\n".join(t[0] for t in cur.fetchall())  # имена таблиц по одной на строку
            finally:  # выполнится в любом случае
                con.close()  # закрываем соединение

        # ---- СТАРЫЕ ФОРМАТЫ OFFICE: .doc/.xls сначала конвертируем ----
        if ext in (".doc", ".xls"):  # если это старый формат Office
            converted = convert_office_file(path)  # пытаемся сконвертировать в новый формат
            if converted is not None and converted.suffix.lower() == ".docx":  # если получили .docx
                doc = Document(str(converted))  # открываем сконвертированный документ
                return "\n".join(p.text for p in doc.paragraphs)  # текст всех абзацев
            if converted is not None and converted.suffix.lower() == ".xlsx":  # если получили .xlsx
                wb = load_workbook(str(converted), read_only=True, data_only=True)  # открываем книгу в экономном режиме
                lines = []  # сюда будем складывать строки
                for ws in wb.worksheets:  # перебираем листы
                    for row in ws.iter_rows(values_only=True):  # перебираем строки
                        lines.append("\t".join("" if c is None else str(c) for c in row))  # ячейки через таб
                return "\n".join(lines)  # строки через перевод строки

        # ---- OLE (старый формат .ppt) ----
        if ext == ".ppt":  # если это старый формат PowerPoint
            import olefile  # подключаем библиотеку для OLE
            if olefile.isOleFile(str(path)):  # проверяем, что это OLE-файл
                ole = olefile.OleFileIO(str(path))  # открываем OLE-контейнер
                return "\n".join("/".join(s) for s in ole.listdir())  # список внутренних потоков

        # ---- ХАОШИР: универсальные метаданные (только для форматов, где мы не пробовали специальный парсер) ----
        if ext not in (".doc", ".xls"):  # для всех форматов, кроме .doc/.xls
            parser = createParser(str(path))  # создаём парсер hachoir
            if parser:  # если парсер создался
                meta = extractMetadata(parser)  # извлекаем метаданные
                if meta:  # если метаданные есть
                    return str(meta.exportDictionary())  # отдаём их в виде словаря-строки

    except Exception as e:  # если что-то сломалось при чтении
        print(f"[CONTENT-SKIP] {path.name}: {type(e).__name__}: {e}")  # сообщаем, что файл пропущен, и почему
        return None  # отдаём None

    return None  # формат незнакомый - ничего не извлекли, отдаём None

def build_record(path: Path) -> Optional[FileRecord]:  # объявляем функцию: по пути собирает FileRecord или None
    try: 
        st = path.stat()  # берём информацию о файле: размер, даты создания/изменения
        return FileRecord(  # создаём и отдаём запись FileRecord
            path=str(path.resolve()),  # полный абсолютный путь
            name=path.name,  # имя файла с расширением
            extension=path.suffix.lower(),  # расширение маленькими буквами
            size_bytes=st.st_size,  # размер в байтах
            md5_hash=compute_md5(path),  # считаем MD5 через нашу функцию
            created_at=datetime.fromtimestamp(st.st_ctime, tz=timezone.utc),  # дата создания в UTC
            modified_at=datetime.fromtimestamp(st.st_mtime, tz=timezone.utc),  # дата изменения в UTC
            content=read_text_content(path, st.st_size),  # пытаемся достать текст/метаданные
            status=FileStatus.NEW,  # пока ставим "новый" - потом уточним при сравнении
        )
    except OSError as e:  # если файл не удалось прочитать
        print(f"[SKIP] {path} -> {type(e).__name__}: {e}")  # сообщаем, что файл пропущен, и почему
        return None  # отдаём None

# ------------------------- СНИМОК -------------------------
def load_snapshot() -> dict[str, str]:  # объявляем функцию: читает снимок прошлого скана, возвращает словарь path -> md5
    if not SNAPSHOT_FILE.exists():  # если файла снимка нет
        return {}  # отдаём пустой словарь - все файлы будут "новыми"
    try:
        data = json.loads(SNAPSHOT_FILE.read_text(encoding="utf-8"))  # читаем JSON из файла снимка
        return {rec["path"]: rec["md5_hash"] for rec in data}  # превращаем в словарь путь -> md5
    except (json.JSONDecodeError, KeyError, OSError) as e:  # если JSON битый, нет нужного поля или файл не читается
        print(f"[WARN] Не удалось прочитать снимок: {e}")  # предупреждаем
        return {}  # считаем, что снимка нет - все файлы будут "новыми"

def save_snapshot(records: list[FileRecord]) -> None:  # объявляем функцию: сохраняет снимок текущего скана
    alive = [r for r in records if r.status != FileStatus.DELETED]  # берём только "живые" записи (удалённые в снимок не пишем)
    SNAPSHOT_FILE.write_text(  # пишем в файл снимка
        json.dumps([r.to_jsonable() for r in alive], ensure_ascii=False, indent=2),  # JSON: список записей, кириллица как есть, с отступами
        encoding="utf-8",  # кодировка файла - utf-8
    )
# ------------------------- СКАНЕР -------------------------
def scan_directory(root: Path) -> Iterator[FileRecord]:  # обходит папку и отдаёт записи по одной
    for path in root.rglob("*"):  # перебираем все файлы и папки внутри root
        if not path.is_file():  # если это не файл (а папка)
            continue  # пропускаем - папки не записываем
        record = build_record(path)  # собираем запись о файле
        if record is not None:  # если запись получилась (файл читается)
            yield record  # отдаём её наружу по одной

# ------------------------- ОСНОВНАЯ ЛОГИКА -------------------------
def run_scan(root: Path) -> list[FileRecord]:  # объявляем функцию: делает полный проход и возвращает список записей
    previous = load_snapshot()  # читаем прошлый снимок: path -> md5
    seen_paths: set[str] = set()  # множество путей, которые встретили в этом скане
    results: list[FileRecord] = []  # сюда будем складывать все записи

    for record in scan_directory(root):  # идём по всем файлам через генератор
        seen_paths.add(record.path)  # помечаем путь как "пройденный"

        old_hash = previous.get(record.path)  # ищем этот путь в прошлом снимке
        if old_hash is None:  # если в снимке его нет
            record.status = FileStatus.NEW  # значит, файл новый
        elif old_hash == record.md5_hash:  # если md5 совпал со снимком
            record.status = FileStatus.UNCHANGED  # файл не менялся
        else:  # иначе (md5 другой)
            record.status = FileStatus.MODIFIED  # файл изменился

        storage.save(record)  # сохраняем запись
        results.append(record)  # добавляем в общий список

    for old_path, old_hash in previous.items():  # теперь ищем удалённые: перебираем прошлый снимок
        if old_path in seen_paths:  # если этот путь уже встречался в текущем скане
            continue  # пропускаем - файл жив
        deleted = FileRecord(  # файла больше нет - делаем запись-маркер "удалён"
            path=old_path,  # путь из снимка
            name=Path(old_path).name,  # имя файла
            extension=Path(old_path).suffix.lower(),  # расширение
            size_bytes=0,  # размер уже не важен - ставим 0
            md5_hash=old_hash,  # оставляем старый md5 из снимка
            created_at=datetime.now(timezone.utc),  # подставляем текущее время
            modified_at=datetime.now(timezone.utc),  # подставляем текущее время
            status=FileStatus.DELETED,  # статус "удалён"
        )  # закрываем конструктор FileRecord
        storage.save(deleted)  # сохраняем маркер в хранилище
        results.append(deleted)  # добавляем в общий список

    return results  # отдаём все записи (новые/изменённые/без изменений/удалённые)

def main() -> None:  # объявляем главную функцию (ничего не возвращает)
    root = Path(ROOT_PATH)  # берём путь сканирования из глобальной настройки
    print(f"[СТАРТ] Сканирую папку: {root}")  # пишем в консоль, какую папку сканируем
    print(f"[СНИМОК] {SNAPSHOT_FILE}")  # пишем, какой файл снимка используется

    if not root.exists():  # если папки не существует
        print(f"Путь не существует: {root}")  # сообщаем об этом
        return  # и выходим

    try:
        storage.begin_scan(root)  # создаём запись о новом скане (в БД) или ничего не делаем (консоль)
        results = run_scan(root)  # запускаем полный проход и получаем список записей

        summary = {s: 0 for s in FileStatus.ALL}  # словарь счётчиков по статусам, все 0
        for r in results:  # перебираем все записи
            summary[r.status] = summary.get(r.status, 0) + 1  # увеличиваем счётчик нужного статуса на 1

        storage.finish_scan(summary)  # закрываем скан: время окончания и счётчики (в БД) или ничего (консоль)
        storage.close()  # фиксируем всё в БД и закрываем соединение (или ничего в консоли)
    except Exception:  # если на любом из шагов выше что-то упало
        storage.rollback()  # откатываем всё, что записали в БД, и закрываем соединение
        raise  # пробрасываем исключение наверх - пусть видно, что случилось

    save_snapshot(results)  # сохраняем снимок - пригодится в следующий раз

    print("=" * 60)  # рисуем разделитель в консоли
    print(f"Всего обработано:     {len(results)}")  # сколько всего записей
    print(f"  новых:              {summary[FileStatus.NEW]}")  # сколько новых
    print(f"  без изменений:      {summary[FileStatus.UNCHANGED]}")  # сколько без изменений
    print(f"  изменено:           {summary[FileStatus.MODIFIED]}")  # сколько изменено
    print(f"  удалено:            {summary[FileStatus.DELETED]}")  # сколько удалено
    print("=" * 60)  # ещё один разделитель

if __name__ == "__main__":  # если файл запущен напрямую
    main()  # вызываем главную функцию
