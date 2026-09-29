# scanner.py
from __future__ import annotations

import hashlib
import json
import mimetypes
import sys
import os
import psycopg
from dataclasses import dataclass, asdict, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator, Optional

ENABLE_DB_SYNC = True
DB_DSN = os.environ.get(
    "DB_DSN",
    "host=localhost port=5432 dbname=file_storage user=postgres password=123",)
ROOT_PATH = r"D:\Новая папка"
SNAPSHOT_FILE = Path("snapshot.json")
HASH_CHUNK_SIZE = 64 * 1024
MAX_CONTENT_SIZE = 5 * 1024 * 1024
TEXT_EXTENSIONS = {".txt", ".py", ".json", ".csv", ".md", ".log", ".xml", ".html", ".yaml", ".yml"}

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


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


@dataclass
class FileRecord:
    path: str
    name: str
    extension: str
    size_bytes: int
    sha256_hash: str
    mime_type: Optional[str]
    created_at: datetime
    modified_at: datetime
    status: str = FileStatus.NEW      
    scanned_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    content: Optional[str] = None

    def to_jsonable(self) -> dict:
        d = asdict(self)
        for k in ("created_at", "modified_at", "scanned_at"):
            d[k] = d[k].isoformat()
        return d

    def to_db_tuple(self) -> tuple:
        return (
            self.path,
            self.name,
            self.extension,
            self.size_bytes,
            self.sha256_hash,
            self.mime_type,
            self.created_at,
            self.modified_at,
            self.scanned_at,
            self.content,
            self.status,
        )

    def to_row(self) -> dict:
        return {
            "path": self.path,
            "name": self.name,
            "extension": self.extension,
            "size_bytes": self.size_bytes,
            "sha256_hash": self.sha256_hash,
            "mime_type": self.mime_type,
            "created_at": self.created_at,
            "modified_at": self.modified_at,
            "scanned_at": self.scanned_at,
            "content": self.content,
            "status": self.status,
        }


class ConsoleStorage:
    def save(self, record: FileRecord) -> None:
        created = record.created_at.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        modified = record.modified_at.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        status_ru = FileStatus.RU.get(record.status, record.status)

        print(
            f"[{status_ru}] {record.name}\n"
            f"    status:      {record.status}\n"       
            f"    путь:        {record.path}\n"
            f"    размер:      {record.size_bytes} B\n"
            f"    sha256:      {record.sha256_hash[:16]}...\n"
            f"    создан:      {created}\n"
            f"    изменён:     {modified}\n"
        )


class DatabaseStorage:
    def __init__(self) -> None:
        self.conn = psycopg.connect(DB_DSN)

    def save(self, record: FileRecord) -> None:
        row = record.to_row()
        self.conn.execute(row)

    def close(self) -> None:
        self.conn.commit()  
        self.conn.close()


storage = DatabaseStorage() if ENABLE_DB_SYNC else ConsoleStorage()


def compute_sha256(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(HASH_CHUNK_SIZE), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def read_text_content(path: Path, size_bytes: int) -> Optional[str]:
    if size_bytes > MAX_CONTENT_SIZE:
        return None
    if path.suffix.lower() not in TEXT_EXTENSIONS:
        return None
    try:
          return path.read_text(encoding="utf-8", errors="replace").replace("\x00", "")
    except (OSError, UnicodeDecodeError):
        return None


def build_record(path: Path) -> Optional[FileRecord]:
    try:
        st = path.stat()
        return FileRecord(
            path=str(path.resolve()),
            name=path.name,
            extension=path.suffix.lower(),
            size_bytes=st.st_size,
            sha256_hash=compute_sha256(path),
            mime_type=mimetypes.guess_type(str(path))[0],
            created_at=datetime.fromtimestamp(st.st_ctime, tz=timezone.utc),
            modified_at=datetime.fromtimestamp(st.st_mtime, tz=timezone.utc),
            content=read_text_content(path, st.st_size),
            status=FileStatus.NEW,
        )
    except (PermissionError, FileNotFoundError, OSError) as e:
        print(f"[SKIP] {path} -> {type(e).__name__}: {e}")
        return None


def load_snapshot() -> dict[str, str]:
    if not SNAPSHOT_FILE.exists():
        return {}
    try:
        data = json.loads(SNAPSHOT_FILE.read_text(encoding="utf-8"))
        return {rec["path"]: rec["sha256_hash"] for rec in data}
    except (json.JSONDecodeError, KeyError, OSError) as e:
        print(f"[WARN] Не удалось прочитать снимок: {e}")
        return {}


def save_snapshot(records: list[FileRecord]) -> None:
    alive = [r for r in records if r.status != FileStatus.DELETED]
    SNAPSHOT_FILE.write_text(
        json.dumps([r.to_jsonable() for r in alive], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def scan_directory(root: Path) -> Iterator[FileRecord]:
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        record = build_record(path)
        if record is not None:
            yield record


def run_scan(root: Path) -> list[FileRecord]:
    previous = load_snapshot()
    seen_paths: set[str] = set()
    results: list[FileRecord] = []

    for record in scan_directory(root):
        seen_paths.add(record.path)

        old_hash = previous.get(record.path)
        if old_hash is None:
            record.status = FileStatus.NEW
        elif old_hash == record.sha256_hash:
            record.status = FileStatus.UNCHANGED
        else:
            record.status = FileStatus.MODIFIED

        storage.save(record)
        results.append(record)

    for old_path, old_hash in previous.items():
        if old_path in seen_paths:
            continue
        deleted = FileRecord(
            path=old_path,
            name=Path(old_path).name,
            extension=Path(old_path).suffix.lower(),
            size_bytes=0,
            sha256_hash=old_hash,
            mime_type=None,
            created_at=datetime.now(timezone.utc),
            modified_at=datetime.now(timezone.utc),
            status=FileStatus.DELETED,
        )
        storage.save(deleted)
        results.append(deleted)

    return results


def main() -> None:
    root = Path(ROOT_PATH)
    if not root.exists():
        print(f"Путь не существует: {root}")
        return

    results = run_scan(root)

    if ENABLE_DB_SYNC:
        storage.close()

    if not ENABLE_DB_SYNC:
        save_snapshot(results)

    summary = {s: 0 for s in FileStatus.ALL}
    for r in results:
        summary[r.status] = summary.get(r.status, 0) + 1

    print("=" * 60)
    print(f"Всего обработано:     {len(results)}")
    print(f"  новых:              {summary[FileStatus.NEW]}")
    print(f"  без изменений:      {summary[FileStatus.UNCHANGED]}")
    print(f"  изменено:           {summary[FileStatus.MODIFIED]}")
    print(f"  удалено:            {summary[FileStatus.DELETED]}")
    print("=" * 60)


if __name__ == "__main__":
    main()
