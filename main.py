# main.py
from __future__ import annotations

import csv
import hashlib
import io
import json
import os
import sqlite3
import sys
import zipfile
from dataclasses import dataclass, asdict, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator, Optional

import chardet
import pymupdf as fitz
import psycopg2
import py7zr
import rarfile
import yaml
from bs4 import BeautifulSoup
from docx import Document
from hachoir.metadata import extractMetadata
from hachoir.parser import createParser
from mutagen import File as MutagenFile
from odf import teletype
from odf.opendocument import load as load_odf
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation

# ------------------------- КОНФИГ -------------------------
_cli_root = None
if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
    _cli_root = sys.argv[1]

_env_root = os.environ.get("ROOT_PATH_OVERRIDE")

ROOT_PATH = _cli_root or _env_root or r"E:\MAX_files"

ENABLE_DB_SYNC = True
if os.environ.get("ENABLE_DB_SYNC_OVERRIDE"):
    ENABLE_DB_SYNC = os.environ["ENABLE_DB_SYNC_OVERRIDE"] == "1"

DB_DSN = os.environ.get(
    "DB_DSN",
    "host=localhost port=5432 dbname=file_storage user=postgres password=12345",
)

_cache_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ScannerApp"
_cache_dir.mkdir(parents=True, exist_ok=True)
_key = hashlib.md5(str(Path(ROOT_PATH).resolve()).encode("utf-8")).hexdigest()[:12]
SNAPSHOT_FILE = _cache_dir / f"snapshot_{_key}.json"

HASH_CHUNK_SIZE = 64 * 1024
MAX_CONTENT_SIZE = 5 * 1024 * 1024

TEXT_EXTENSIONS = {".txt", ".py", ".json", ".md", ".log", ".xml"}

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ------------------------- СТАТУСЫ -------------------------
class FileStatus:
    NEW = "new"
    UNCHANGED = "unchanged"
    MODIFIED = "modified"
    DELETED = "deleted"

    ALL = (NEW, UNCHANGED, MODIFIED, DELETED)

    RU = {
        NEW: "НОВЫЙ",
        UNCHANGED: "БЕЗ ИЗМЕНЕНИЙ",
        MODIFIED: "ИЗМЕНЁН",
        DELETED: "УДАЛЁН",
    }

# ------------------------- МОДЕЛЬ -------------------------
@dataclass
class FileRecord:
    path: str
    name: str
    extension: str
    size_bytes: int
    created_at: datetime
    modified_at: datetime
    status: str = FileStatus.NEW
    scanned_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    content: Optional[str] = None
    md5_hash: Optional[str] = None
    file_id: Optional[int] = None
    change_count: Optional[int] = None

    def to_jsonable(self) -> dict:
        d = asdict(self)
        for k in ("created_at", "modified_at", "scanned_at"):
            d[k] = d[k].isoformat()
        return d

    def to_row(self) -> dict:
        return {
            "name": self.name,
            "extension": self.extension,
            "size_bytes": self.size_bytes,
            "md5_hash": self.md5_hash,
            "created_at": self.created_at,
            "modified_at": self.modified_at,
            "scanned_at": self.scanned_at,
            "content": self.content,
            "status": self.status,
        }

# ------------------------- ХРАНИЛИЩА -------------------------
class ConsoleStorage:
    def save(self, record: FileRecord) -> None:
        created = record.created_at.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        modified = record.modified_at.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        scanned = record.scanned_at.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        status_ru = FileStatus.RU.get(record.status, record.status)
        file_id_str = str(record.file_id) if record.file_id is not None else "-"

        print(
            f"[{status_ru}] {record.name}\n"
            f"    id:          {file_id_str}\n"
            f"    status:      {record.status}\n"
            f"    путь:        {record.path}\n"
            f"    размер:      {record.size_bytes} B\n"
            f"    md5:         {record.md5_hash}\n"
            f"    создан:      {created}\n"
            f"    изменён:     {modified}\n"
            f"    просканирован: {scanned}\n"
        )

    def begin_scan(self, root: Path) -> None:
        pass

    def finish_scan(self, summary: dict) -> None:
        pass

    def rollback(self) -> None:
        pass

    def close(self) -> None:
        pass


class DatabaseStorage:
    def __init__(self) -> None:
        self.conn = psycopg2.connect(DB_DSN)
        self.scan_id: Optional[int] = None
        self.root: Optional[Path] = None
        self.path_cache: dict[str, int] = {}

    def begin_scan(self, root: Path) -> None:
        self.root = root.resolve()
        with self.conn.cursor() as cur:
            cur.execute(
                "INSERT INTO scans (root_path) VALUES (%s) RETURNING id",
                (str(self.root),),
            )
            self.scan_id = cur.fetchone()[0]

    def get_path_id(self, directory: Path) -> int:
        key = str(directory)
        cached = self.path_cache.get(key)
        if cached is not None:
            return cached

        parent_id: Optional[int] = None
        if directory != self.root and directory.parent != directory:
            parent_id = self.get_path_id(directory.parent)

        with self.conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO paths (parent_id, name, full_path)
                VALUES (%s, %s, %s)
                ON CONFLICT (full_path) DO UPDATE SET name = EXCLUDED.name
                RETURNING id
                """,
                (parent_id, directory.name or key, key),
            )
            path_id = cur.fetchone()[0]

        self.path_cache[key] = path_id
        return path_id

    def save(self, record: FileRecord) -> None:
        if record.status == FileStatus.DELETED:
            with self.conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE files
                       SET status = 'deleted', last_scan_id = %s, scanned_at = now()
                     WHERE name = %s
                       AND path_id = (SELECT id FROM paths WHERE full_path = %s)
                 RETURNING id, change_count
                    """,
                    (self.scan_id, record.name, str(Path(record.path).parent)),
                )
                row = cur.fetchone()
                if row:
                    record.file_id = row[0]
                    record.change_count = row[1]
            file_id_str = str(record.file_id) if record.file_id is not None else "-"
            print(f"[{FileStatus.RU[record.status]}] {record.name}  id={file_id_str}")
            return

        path_id = self.get_path_id(Path(record.path).parent)

        params = record.to_row()
        params["path_id"] = path_id
        params["scan_id"] = self.scan_id
        if params["content"]:
            params["content"] = params["content"].replace("\x00", "")

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
                    sum(summary.values()),
                    summary[FileStatus.NEW],
                    summary[FileStatus.UNCHANGED],
                    summary[FileStatus.MODIFIED],
                    summary[FileStatus.DELETED],
                    self.scan_id,
                ),
            )

    def rollback(self) -> None:
        self.conn.rollback()
        self.conn.close()

    def close(self) -> None:
        self.conn.commit()
        self.conn.close()


storage = DatabaseStorage() if ENABLE_DB_SYNC else ConsoleStorage()

# ------------------------- УТИЛИТЫ -------------------------
def compute_md5(path: Path) -> str:
    hasher = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(HASH_CHUNK_SIZE), b""):
            hasher.update(chunk)
    return hasher.hexdigest()

# ------------------------- КОНВЕРТЕР СТАРЫХ ФОРМАТОВ -------------------------
def convert_office_file(path: Path) -> Optional[Path]:
    ext = path.suffix.lower()
    target = None
    if ext == ".doc":
        target = path.with_suffix(".docx")
    elif ext == ".xls":
        target = path.with_suffix(".xlsx")
    else:
        return None
    if target.exists():
        return target
    try:
        import win32com.client
        if ext == ".doc":
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            try:
                doc = app.Documents.Open(str(path))
                doc.SaveAs2(str(target), 16)
                doc.Close(False)
            finally:
                app.Quit()
        else:
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                wb = app.Workbooks.Open(str(path))
                wb.SaveAs(str(target), 51)
                wb.Close(False)
            finally:
                app.Quit()
        print(f"[КОНВЕРТАЦИЯ] {path.name} -> {target.name}")
        return target
    except Exception as e:
        print(f"[CONVERT-SKIP] {path.name}: {type(e).__name__}: {e}")
        return None


def read_text_content(path: Path, size_bytes: int) -> Optional[str]:
    if size_bytes > MAX_CONTENT_SIZE:
        return None

    ext = path.suffix.lower()

    try:
        if ext in TEXT_EXTENSIONS:
            raw = path.read_bytes()
            enc = chardet.detect(raw)["encoding"] or "utf-8"
            return raw.decode(enc, errors="replace").replace("\x00", "")

        if ext == ".docx":
            doc = Document(str(path))
            return "\n".join(p.text for p in doc.paragraphs)

        if ext == ".xlsx":
            wb = load_workbook(str(path), read_only=True, data_only=True)
            lines = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    lines.append("\t".join("" if c is None else str(c) for c in row))
            return "\n".join(lines)

        if ext == ".pptx":
            prs = Presentation(str(path))
            lines = []
            for i, slide in enumerate(prs.slides, 1):
                lines.append(f"--- Слайд {i} ---")
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        lines.append(shape.text_frame.text)
            return "\n".join(lines)

        if ext == ".pdf":
            with fitz.open(str(path)) as doc:
                return "\n".join(page.get_text() for page in doc)

        if ext == ".csv":
            raw = path.read_bytes()
            enc = chardet.detect(raw)["encoding"] or "utf-8"
            text = raw.decode(enc, errors="replace")
            rows = list(csv.reader(io.StringIO(text)))
            return "\n".join("\t".join(r) for r in rows)

        if ext in (".yaml", ".yml"):
            data = yaml.safe_load(path.read_text(encoding="utf-8", errors="replace"))
            return yaml.safe_dump(data, allow_unicode=True)

        if ext in (".html", ".htm"):
            soup = BeautifulSoup(path.read_text(encoding="utf-8", errors="replace"), "html.parser")
            return soup.get_text(separator="\n", strip=True)

        if ext in (".odt", ".ods", ".odp"):
            doc = load_odf(str(path))
            return teletype.extractText(doc)

        if ext in (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"):
            with Image.open(path) as img:
                return f"Изображение {img.format} {img.width}x{img.height} mode={img.mode}"

        if ext in (".mp3", ".flac", ".ogg", ".m4a", ".wav", ".wma"):
            audio = MutagenFile(str(path), easy=True)
            if audio is None:
                return None
            return "\n".join(f"{k}: {v}" for k, v in audio.items())

        if ext == ".zip":
            with zipfile.ZipFile(path) as z:
                return "\n".join(z.namelist())
        if ext == ".7z":
            with py7zr.SevenZipFile(path) as z:
                return "\n".join(z.getnames())
        if ext == ".rar":
            with rarfile.RarFile(path) as r:
                return "\n".join(r.namelist())

        if ext in (".sqlite", ".db"):
            con = sqlite3.connect(str(path))
            try:
                cur = con.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
                return "\n".join(t[0] for t in cur.fetchall())
            finally:
                con.close()

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

        if ext == ".ppt":
            import olefile
            if olefile.isOleFile(str(path)):
                ole = olefile.OleFileIO(str(path))
                return "\n".join("/".join(s) for s in ole.listdir())

        if ext not in (".doc", ".xls"):
            parser = createParser(str(path))
            if parser:
                meta = extractMetadata(parser)
                if meta:
                    return str(meta.exportDictionary())

    except Exception as e:
        print(f"[CONTENT-SKIP] {path.name}: {type(e).__name__}: {e}")
        return None

    return None


def build_record(path: Path) -> Optional[FileRecord]:
    try:
        st = path.stat()
        return FileRecord(
            path=str(path.resolve()),
            name=path.name,
            extension=path.suffix.lower(),
            size_bytes=st.st_size,
            md5_hash=compute_md5(path),
            created_at=datetime.fromtimestamp(st.st_ctime, tz=timezone.utc),
            modified_at=datetime.fromtimestamp(st.st_mtime, tz=timezone.utc),
            content=read_text_content(path, st.st_size),
            status=FileStatus.NEW,
        )
    except OSError as e:
        print(f"[SKIP] {path} -> {type(e).__name__}: {e}")
        return None

# ------------------------- СНИМОК -------------------------
def load_snapshot() -> dict[str, str]:
    if not SNAPSHOT_FILE.exists():
        return {}
    try:
        data = json.loads(SNAPSHOT_FILE.read_text(encoding="utf-8"))
        return {rec["path"]: rec["md5_hash"] for rec in data}
    except (json.JSONDecodeError, KeyError, OSError) as e:
        print(f"[WARN] Не удалось прочитать снимок: {e}")
        return {}


def save_snapshot(records: list[FileRecord]) -> None:
    alive = [r for r in records if r.status != FileStatus.DELETED]
    SNAPSHOT_FILE.write_text(
        json.dumps([r.to_jsonable() for r in alive], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ------------------------- НОВОЕ: ПРЕДЫДУЩЕЕ СОСТОЯНИЕ ИЗ БД -------------------------
def load_previous_from_db(root: Path) -> dict[str, str]:
    """
    Читает из БД последнее известное состояние файлов ПОД данной корневой папкой.
    Возвращает {полный путь -> md5_hash}. Это позволяет сравнивать с БД даже если
    локальный snapshot.json потерян или скан был давно/на другой машине.
    """
    if not ENABLE_DB_SYNC:
        return {}

    root_str = str(root.resolve())
    result: dict[str, str] = {}
    try:
        with psycopg2.connect(DB_DSN) as conn:
            with conn.cursor() as cur:
                # Берём файлы, чей путь начинается с корня, и которые не помечены deleted.
                # Если файл удалён ранее — в previous его не кладём (иначе он «воскреснет»).
                cur.execute(
                    """
                    SELECT p.full_path, f.name, f.md5_hash
                      FROM files f
                      JOIN paths p ON p.id = f.path_id
                     WHERE (p.full_path = %s OR p.full_path LIKE %s)
                       AND f.status <> 'deleted'
                    """,
                    (root_str, root_str.rstrip("\\/") + "\\%"),
                )
                for full_path, name, md5_hash in cur.fetchall():
                    if md5_hash is None:
                        continue
                    # склеиваем так же, как record.path: Path(full_path) / name
                    rec_path = str(Path(full_path) / name)
                    result[rec_path] = md5_hash
    except Exception as e:
        print(f"[WARN] Не удалось прочитать предыдущее состояние из БД: {type(e).__name__}: {e}")
    return result


# ------------------------- СКАНЕР -------------------------
def scan_directory(root: Path) -> Iterator[FileRecord]:
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        record = build_record(path)
        if record is not None:
            yield record

# ------------------------- ОСНОВНАЯ ЛОГИКА -------------------------
def run_scan(root: Path) -> list[FileRecord]:
    # Приоритет: состояние из БД (актуально между любыми сканами) -> snapshot.json
    previous = load_previous_from_db(root)
    if not previous:
        previous = load_snapshot()

    seen_paths: set[str] = set()
    results: list[FileRecord] = []

    for record in scan_directory(root):
        seen_paths.add(record.path)

        old_hash = previous.get(record.path)
        if old_hash is None:
            record.status = FileStatus.NEW
        elif old_hash == record.md5_hash:
            record.status = FileStatus.UNCHANGED
        else:
            record.status = FileStatus.MODIFIED

        storage.save(record)
        results.append(record)

    for old_path, old_hash in previous.items():
        if old_path in seen_paths:
            continue
        # Если файл уже помечен в БД как deleted — повторно не создаём запись
        # (иначе каждый скан будет «удалять» его заново). Проверяем по БД.
        if ENABLE_DB_SYNC:
            try:
                with psycopg2.connect(DB_DSN) as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            """
                            SELECT 1 FROM files f
                              JOIN paths p ON p.id = f.path_id
                             WHERE p.full_path = %s AND f.name = %s AND f.status = 'deleted'
                            """,
                            (str(Path(old_path).parent), Path(old_path).name),
                        )
                        if cur.fetchone():
                            continue
            except Exception:
                pass

        deleted = FileRecord(
            path=old_path,
            name=Path(old_path).name,
            extension=Path(old_path).suffix.lower(),
            size_bytes=0,
            md5_hash=old_hash,
            created_at=datetime.now(timezone.utc),
            modified_at=datetime.now(timezone.utc),
            status=FileStatus.DELETED,
        )
        storage.save(deleted)
        results.append(deleted)

    return results


def main() -> None:
    root = Path(ROOT_PATH)
    print(f"[СТАРТ] Сканирую папку: {root}")
    print(f"[СНИМОК] {SNAPSHOT_FILE}")

    if not root.exists():
        print(f"Путь не существует: {root}")
        return

    try:
        storage.begin_scan(root)
        results = run_scan(root)

        summary = {s: 0 for s in FileStatus.ALL}
        for r in results:
            summary[r.status] = summary.get(r.status, 0) + 1

        storage.finish_scan(summary)
        storage.close()
    except Exception:
        storage.rollback()
        raise

    save_snapshot(results)

    print("=" * 60)
    print(f"Всего обработано:     {len(results)}")
    print(f"  новых:              {summary[FileStatus.NEW]}")
    print(f"  без изменений:      {summary[FileStatus.UNCHANGED]}")
    print(f"  изменено:           {summary[FileStatus.MODIFIED]}")
    print(f"  удалено:            {summary[FileStatus.DELETED]}")
    print("=" * 60)


if __name__ == "__main__":
    main()