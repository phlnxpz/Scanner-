
# gui.py
from __future__ import annotations

import difflib
import json
import os
import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import Tk, StringVar, BooleanVar, END
from tkinter import ttk, filedialog, messagebox

import psycopg2

APP_TITLE = "Сканер файлов"
APP_VERSION = "scanner v1.1"
MAIN_SCRIPT = Path(__file__).with_name("main.py")

DB_DSN = os.environ.get(
    "DB_DSN",
    "host=localhost port=5432 dbname=file_storage user=postgres password=12345",
)

SETTINGS_FILE = Path(__file__).with_name("gui_settings.json")

DEFAULT_SETTINGS = {
    "reports_scan_dir": "",
    "reports_changes_dir": "",
    "theme": "light",
}

FILES_COLUMNS = (
    ("id",           "id",              70),
    ("path",         "Путь",           300),
    ("name",         "Имя",            170),
    ("extension",    "Расширение",      90),
    ("size_bytes",   "Размер (байт)",  110),
    ("md5_hash",     "MD5",            250),
    ("created_at",   "Создан",         150),
    ("modified_at",  "Изменён",        150),
    ("scanned_at",   "Просканирован",  150),
    ("status",       "Статус",         110),
    ("change_count", "Изменений",       90),
    ("content",      "Содержимое",     300),
)

SCANS_COLUMNS = (
    ("id",               "id",               60),
    ("root_path",        "Корень скана",    320),
    ("started_at",       "Начат",           150),
    ("finished_at",      "Завершён",        150),
    ("files_total",      "Всего",            80),
    ("files_new",        "Новых",            80),
    ("files_unchanged",  "Без изм.",         90),
    ("files_modified",   "Изменено",         90),
    ("files_deleted",    "Удалено",          90),
)

PATHS_COLUMNS = (
    ("id",         "id",          60),
    ("parent_id",  "parent_id",   90),
    ("name",       "Имя",        200),
    ("full_path",  "Полный путь", 500),
)

JSON_COLUMNS = (
    ("id",          "id",          60),
    ("file_id",     "file_id",     80),
    ("updated_at",  "Обновлён",   150),
    ("data",        "data",      1000),
)

STATUS_RU = {
    "new": "НОВЫЙ",
    "unchanged": "БЕЗ ИЗМЕНЕНИЙ",
    "modified": "ИЗМЕНЁН",
    "deleted": "УДАЛЁН",
}

THEMES = {
    "light": {
        "bg":        "#f5f5f5",
        "fg":        "#1e1e1e",
        "panel":     "#ffffff",
        "panel_fg":  "#1e1e1e",
        "sidebar":   "#e8e8e8",
        "sidebar_fg":"#1e1e1e",
        "accent":    "#0078d4",
        "accent_fg": "#ffffff",
        "tree_bg":   "#ffffff",
        "tree_fg":   "#1e1e1e",
        "footer":    "#dcdcdc",
        "footer_fg": "#555555",
    },
    "dark": {
        "bg":        "#1e1e1e",
        "fg":        "#e6e6e6",
        "panel":     "#252526",
        "panel_fg":  "#e6e6e6",
        "sidebar":   "#2d2d30",
        "sidebar_fg":"#e6e6e6",
        "accent":    "#0e639c",
        "accent_fg": "#ffffff",
        "tree_bg":   "#1e1e1e",
        "tree_fg":   "#e6e6e6",
        "footer":    "#2d2d30",
        "footer_fg": "#999999",
    },
}


def load_settings() -> dict:
    if SETTINGS_FILE.exists():
        try:
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            merged = dict(DEFAULT_SETTINGS)
            merged.update({k: v for k, v in data.items() if k in DEFAULT_SETTINGS})
            return merged
        except Exception:
            pass
    return dict(DEFAULT_SETTINGS)


def save_settings(settings: dict) -> None:
    try:
        SETTINGS_FILE.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        print(f"[WARN] Не удалось сохранить настройки: {e}")


class ScannerGUI:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1600x900")

        self.settings = load_settings()

        self.folder = StringVar(value="")
        self.process: subprocess.Popen | None = None
        self.log_queue: queue.Queue = queue.Queue()
        self.records: list[dict] = []
        self.last_scan_report: Path | None = None
        self.last_changes_report: Path | None = None

        self.previous_content: dict[str, str] = {}
        self.scan_root: str = ""

        self.scan_reports_dir = StringVar(value=self.settings.get("reports_scan_dir", ""))
        self.changes_reports_dir = StringVar(value=self.settings.get("reports_changes_dir", ""))
        self.dark_theme = BooleanVar(value=self.settings.get("theme", "light") == "dark")

        self._style = ttk.Style()
        self._build_ui()
        self._apply_theme()
        self._poll_log()

        self.root.after(50, self._on_startup)

    # =================================================================
    #  ПОСТРОЕНИЕ ИНТЕРФЕЙСА
    # =================================================================
    def _build_ui(self) -> None:
        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True)

        # ---------- САЙДБАР ----------
        self.sidebar = ttk.Frame(main, width=220, style="Sidebar.TFrame")
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        ttk.Label(
            self.sidebar, text="НАВИГАЦИЯ", font=("", 10, "bold"), style="Sidebar.TLabel"
        ).pack(anchor="w", padx=15, pady=(15, 10))

        self.nav_buttons: dict[str, ttk.Button] = {}
        nav_items = [
            ("scan",     "Текущее сканирование"),
            ("db",       "Данные в БД"),
            ("settings", "Настройки"),
        ]
        for key, label in nav_items:
            btn = ttk.Button(
                self.sidebar,
                text=label,
                style="Nav.TButton",
                command=lambda k=key: self._switch_section(k),
            )
            btn.pack(fill="x", padx=10, pady=3)
            self.nav_buttons[key] = btn

        # ---------- КОНТЕНТ ----------
        self.content = ttk.Frame(main)
        self.content.pack(side="left", fill="both", expand=True)

        self.sections: dict[str, ttk.Frame] = {}
        self.sections["scan"] = self._build_scan_section()
        self.sections["db"] = self._build_db_section()
        self.sections["settings"] = self._build_settings_section()

        # ---------- ФУТЕР ----------
        self.footer = ttk.Frame(self.root, height=28, style="Footer.TFrame")
        self.footer.pack(fill="x", side="bottom")
        self.footer.pack_propagate(False)
        ttk.Label(
            self.footer, text=APP_VERSION, style="Footer.TLabel"
        ).pack(side="right", padx=15)

        self._switch_section("scan")

    # =================================================================
    #  РАЗДЕЛ: ТЕКУЩЕЕ СКАНИРОВАНИЕ
    # =================================================================
    def _build_scan_section(self) -> ttk.Frame:
        frame = ttk.Frame(self.content, padding=10)

        top = ttk.Frame(frame)
        top.pack(fill="x", pady=(0, 6))

        ttk.Label(top, text="Папка:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.folder, width=70).grid(
            row=0, column=1, padx=5, sticky="we"
        )
        ttk.Button(top, text="Обзор…", command=self._choose_folder).grid(row=0, column=2)
        top.columnconfigure(1, weight=1)

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(0, 6))

        self.btn_start = ttk.Button(
            buttons, text="Запустить скан", command=self._start_scan
        )
        self.btn_start.pack(side="left")

        self.btn_stop = ttk.Button(
            buttons, text="Остановить", command=self._stop_scan, state="disabled"
        )
        self.btn_stop.pack(side="left", padx=6)

        ttk.Button(buttons, text="Очистить", command=self._clear_scan).pack(
            side="left", padx=6
        )

        self.status = StringVar(value="Готов к работе")
        ttk.Label(buttons, textvariable=self.status).pack(side="right")

        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 6))

        self.notebook_current = ttk.Notebook(frame)
        self.notebook_current.pack(fill="both", expand=True)

        self.cur_files = self._make_tab(self.notebook_current, "files", FILES_COLUMNS)
        self.cur_scans = self._make_tab(self.notebook_current, "scans", SCANS_COLUMNS)
        self.cur_paths = self._make_tab(self.notebook_current, "paths", PATHS_COLUMNS)
        self.cur_json  = self._make_tab(self.notebook_current, "json",  JSON_COLUMNS)

        reports = ttk.Frame(frame)
        reports.pack(fill="x", pady=(6, 0))

        self.scan_report_btn = ttk.Button(
            reports,
            text="Отчёт сканирования: —",
            style="Report.TButton",
            command=self._open_scan_report,
            state="disabled",
        )
        self.scan_report_btn.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.changes_report_btn = ttk.Button(
            reports,
            text="Отчёт изменений: —",
            style="Report.TButton",
            command=self._open_changes_report,
            state="disabled",
        )
        self.changes_report_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))

        return frame

    # =================================================================
    #  РАЗДЕЛ: ДАННЫЕ В БД
    # =================================================================
    def _build_db_section(self) -> ttk.Frame:
        frame = ttk.Frame(self.content, padding=10)

        top = ttk.Frame(frame)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="ВСЯ БАЗА ДАННЫХ", font=("", 10, "bold")).pack(side="left")
        ttk.Button(top, text="Обновить БД", command=self._refresh_db).pack(side="right")

        self.notebook_all = ttk.Notebook(frame)
        self.notebook_all.pack(fill="both", expand=True)

        self.all_files = self._make_tab(self.notebook_all, "files", FILES_COLUMNS)
        self.all_scans = self._make_tab(self.notebook_all, "scans", SCANS_COLUMNS)
        self.all_paths = self._make_tab(self.notebook_all, "paths", PATHS_COLUMNS)
        self.all_json  = self._make_tab(self.notebook_all, "json",  JSON_COLUMNS)

        return frame

    # =================================================================
    #  РАЗДЕЛ: НАСТРОЙКИ
    # =================================================================
    def _build_settings_section(self) -> ttk.Frame:
        frame = ttk.Frame(self.content, padding=20)

        ttk.Label(frame, text="НАСТРОЙКИ", font=("", 12, "bold")).pack(
            anchor="w", pady=(0, 15)
        )

        ttk.Label(frame, text="Папка для отчётов сканирования (обязательно):").pack(
            anchor="w", pady=(5, 2)
        )
        row1 = ttk.Frame(frame)
        row1.pack(fill="x", pady=(0, 12))
        ttk.Entry(row1, textvariable=self.scan_reports_dir, width=70).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            row1, text="Обзор…", command=lambda: self._pick_dir(self.scan_reports_dir)
        ).pack(side="left", padx=4)

        ttk.Label(frame, text="Папка для отчётов изменений (обязательно):").pack(
            anchor="w", pady=(5, 2)
        )
        row2 = ttk.Frame(frame)
        row2.pack(fill="x", pady=(0, 12))
        ttk.Entry(row2, textvariable=self.changes_reports_dir, width=70).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            row2, text="Обзор…", command=lambda: self._pick_dir(self.changes_reports_dir)
        ).pack(side="left", padx=4)

        theme_row = ttk.Frame(frame)
        theme_row.pack(fill="x", pady=(5, 12))
        ttk.Label(theme_row, text="Тёмная тема:").pack(side="left")
        ttk.Checkbutton(
            theme_row,
            variable=self.dark_theme,
            command=self._on_theme_toggle,
        ).pack(side="left", padx=8)

        buttons_row = ttk.Frame(frame)
        buttons_row.pack(anchor="w", pady=(15, 0))

        ttk.Button(
            buttons_row, text="Сохранить настройки", command=self._save_settings
        ).pack(side="left")

        ttk.Button(
            buttons_row, text="Сбросить настройки", command=self._reset_settings
        ).pack(side="left", padx=8)

        ttk.Label(
            frame,
            text="Без задания обеих папок запуск сканирования невозможен.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(15, 0))

        return frame

    # =================================================================
    #  ОБЩИЕ ХЕЛПЕРЫ UI
    # =================================================================
    def _make_tab(self, notebook: ttk.Notebook, title: str, columns: tuple) -> ttk.Treeview:
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=title)

        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings")

        for key, col_title, width in columns:
            tree.heading(key, text=col_title)
            tree.column(
                key, width=width,
                stretch=(key in ("path", "name", "content", "data", "full_path", "root_path")),
            )

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="we")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        return tree

    def _switch_section(self, key: str) -> None:
        for frame in self.sections.values():
            frame.pack_forget()
        self.sections[key].pack(fill="both", expand=True)

        for name, btn in self.nav_buttons.items():
            btn.configure(style="NavActive.TButton" if name == key else "Nav.TButton")

        if key == "db":
            try:
                self._refresh_db()
            except Exception:
                pass

    def _pick_dir(self, var: StringVar) -> None:
        chosen = filedialog.askdirectory(initialdir=var.get() or None)
        if chosen:
            var.set(chosen)

    # =================================================================
    #  ТЕМЫ
    # =================================================================
    def _apply_theme(self) -> None:
        theme = THEMES["dark" if self.dark_theme.get() else "light"]
        s = self._style

        s.theme_use("clam")
        s.configure(".", background=theme["bg"], foreground=theme["fg"])
        s.configure("TFrame", background=theme["bg"])
        s.configure("TLabel", background=theme["bg"], foreground=theme["fg"])
        s.configure("TButton", background=theme["panel"], foreground=theme["fg"])
        s.configure("TEntry", fieldbackground=theme["panel"], foreground=theme["fg"])
        s.configure("TNotebook", background=theme["bg"])
        s.configure("TNotebook.Tab", background=theme["panel"], foreground=theme["fg"])
        s.map(
            "TNotebook.Tab",
            background=[("selected", theme["accent"])],
            foreground=[("selected", theme["accent_fg"])],
        )
        s.configure(
            "Treeview",
            background=theme["tree_bg"],
            fieldbackground=theme["tree_bg"],
            foreground=theme["tree_fg"],
        )
        s.configure("Treeview.Heading", background=theme["panel"], foreground=theme["fg"])
        s.configure("TProgressbar", background=theme["accent"])

        s.configure("Sidebar.TFrame", background=theme["sidebar"])
        s.configure(
            "Sidebar.TLabel",
            background=theme["sidebar"],
            foreground=theme["sidebar_fg"],
        )
        s.configure(
            "Nav.TButton",
            background=theme["sidebar"],
            foreground=theme["sidebar_fg"],
            anchor="w",
            padding=(10, 8),
            relief="flat",
        )
        s.configure(
            "NavActive.TButton",
            background=theme["accent"],
            foreground=theme["accent_fg"],
            anchor="w",
            padding=(10, 8),
            relief="flat",
        )
        s.configure(
            "Report.TButton",
            background=theme["panel"],
            foreground=theme["fg"],
            padding=(8, 6),
        )
        s.configure("Footer.TFrame", background=theme["footer"])
        s.configure(
            "Footer.TLabel",
            background=theme["footer"],
            foreground=theme["footer_fg"],
        )
        s.configure(
            "Hint.TLabel",
            background=theme["bg"],
            foreground=theme["footer_fg"],
        )

        self.root.configure(bg=theme["bg"])

    def _on_theme_toggle(self) -> None:
        self.settings["theme"] = "dark" if self.dark_theme.get() else "light"
        save_settings(self.settings)
        self._apply_theme()

    # =================================================================
    #  СТАРТ / ЗАВЕРШЕНИЕ
    # =================================================================
    def _on_startup(self) -> None:
        try:
            self._refresh_db()
            self.status.set("Готов к сканированию.")
        except Exception as e:
            self.status.set("БД недоступна")
            messagebox.showwarning(
                "База данных недоступна",
                f"Не удалось прочитать БД при запуске:\n"
                f"{type(e).__name__}: {e}\n\nDSN: {DB_DSN}",
            )

    # =================================================================
    #  ДЕЙСТВИЯ СКАНИРОВАНИЯ
    # =================================================================
    def _choose_folder(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.folder.get() or None)
        if chosen:
            self.folder.set(chosen)

    def _clear_scan(self) -> None:
        for tree in (self.cur_files, self.cur_scans, self.cur_paths, self.cur_json):
            for row in tree.get_children():
                tree.delete(row)
        self.records.clear()
        self.last_scan_report = None
        self.last_changes_report = None
        self.scan_report_btn.configure(text="Отчёт сканирования: —", state="disabled")
        self.changes_report_btn.configure(text="Отчёт изменений: —", state="disabled")
        self.status.set("Очищено")

    def _start_scan(self) -> None:
        if self.process is not None:
            return

        scan_dir_str = self.scan_reports_dir.get().strip()
        changes_dir_str = self.changes_reports_dir.get().strip()

        if not scan_dir_str or not changes_dir_str:
            messagebox.showerror(
                "Ошибка",
                "Перед запуском сканирования задайте обе папки для отчётов:\n"
                "— папку отчётов сканирования;\n"
                "— папку отчётов изменений.\n\n"
                "Откройте раздел «Настройки».",
            )
            return

        try:
            Path(scan_dir_str).mkdir(parents=True, exist_ok=True)
            Path(changes_dir_str).mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror(
                "Ошибка",
                f"Не удалось создать папки для отчётов:\n{type(e).__name__}: {e}",
            )
            return

        folder = self.folder.get().strip()
        if not folder:
            messagebox.showerror(
                "Ошибка",
                "Не выбрана папка для сканирования.\nНажмите «Обзор…» и выберите папку.",
            )
            return
        if not Path(folder).exists():
            messagebox.showerror("Ошибка", f"Папка не найдена:\n{folder}")
            return

        self.scan_root = folder
        try:
            self.previous_content = self._load_previous_content(folder)
        except Exception as e:
            self.previous_content = {}
            print(f"[WARN] Не удалось прочитать предыдущее содержимое: {e}")

        self._clear_scan()
        self.progress.start(12)
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.status.set(f"Сканирую: {folder}")

        cmd = [sys.executable, "-u", str(MAIN_SCRIPT), folder]
        env = {"ENABLE_DB_SYNC_OVERRIDE": "1"}

        self.process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env={**os.environ, **env},
        )
        threading.Thread(target=self._reader, args=(self.process,), daemon=True).start()

    def _load_previous_content(self, root: str) -> dict[str, str]:
        root_str = str(Path(root).resolve())
        result: dict[str, str] = {}
        with psycopg2.connect(DB_DSN) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT p.full_path, f.name, f.content
                      FROM files f
                      JOIN paths p ON p.id = f.path_id
                     WHERE (p.full_path = %s OR p.full_path LIKE %s)
                       AND f.status <> 'deleted'
                    """,
                    (root_str, root_str.rstrip("\\/") + "\\%"),
                )
                for full_path, name, content in cur.fetchall():
                    rec_path = str(Path(full_path) / name)
                    result[rec_path] = content or ""
        return result

    def _stop_scan(self) -> None:
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self.status.set("Останавливается…")

    def _reader(self, proc: subprocess.Popen) -> None:
        assert proc.stdout is not None
        for line in proc.stdout:
            line = line.rstrip()
            if line.startswith("[СТАРТ]") or line.startswith("[СНИМОК]"):
                self.log_queue.put(line)
        proc.wait()
        self.log_queue.put(None)

    def _poll_log(self) -> None:
        try:
            while True:
                item = self.log_queue.get_nowait()
                if item is None:
                    self._on_finished()
                    continue
                if isinstance(item, str):
                    self.status.set(item)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log)

    def _on_finished(self) -> None:
        self.progress.stop()
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.process = None

        try:
            self._refresh_db()
        except Exception:
            pass

        try:
            self._load_scan_panel()
        except Exception as e:
            messagebox.showerror(
                "Ошибка чтения БД",
                f"Не удалось прочитать текущий скан:\n{type(e).__name__}: {e}\n\nDSN: {DB_DSN}",
            )
            self.status.set("Готово (ошибка чтения БД)")
            return

        if self.records:
            self._write_reports()
            self.status.set(f"Готово. Записей: {len(self.records)}")
        else:
            self.status.set("Готово (в files ничего не найдено)")

    # =================================================================
    #  ЗАГРУЗКА ПАНЕЛИ ТЕКУЩЕГО СКАНА
    # =================================================================
    def _load_scan_panel(self) -> None:
        for tree in (self.cur_files, self.cur_scans, self.cur_paths, self.cur_json):
            for row in tree.get_children():
                tree.delete(row)
        self.records.clear()

        with psycopg2.connect(DB_DSN) as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT
                        f.id,
                        p.full_path || '\\' || f.name AS full_path,
                        f.name,
                        f.extension,
                        f.size_bytes,
                        f.md5_hash,
                        f.created_at,
                        f.modified_at,
                        f.scanned_at,
                        f.status,
                        f.change_count,
                        f.content
                    FROM files f
                    JOIN paths p ON p.id = f.path_id
                    WHERE f.last_scan_id = (SELECT MAX(id) FROM scans)
                    ORDER BY f.id ASC
                """)
                cur_files_rows = cur.fetchall()

                cur.execute("""
                    SELECT id, root_path, started_at, finished_at,
                           files_total, files_new, files_unchanged,
                           files_modified, files_deleted
                    FROM scans
                    WHERE id = (SELECT MAX(id) FROM scans)
                    ORDER BY id ASC
                """)
                cur_scans_rows = cur.fetchall()

                cur.execute("""
                    SELECT DISTINCT p.id, p.parent_id, p.name, p.full_path
                    FROM paths p
                    WHERE p.id IN (
                        SELECT f.path_id FROM files f
                        WHERE f.last_scan_id = (SELECT MAX(id) FROM scans)
                    )
                    ORDER BY p.id ASC
                """)
                cur_paths_rows = cur.fetchall()

                cur.execute("""
                    SELECT j.id, j.file_id, j.updated_at, j.data::text
                    FROM json j
                    JOIN files f ON f.id = j.file_id
                    WHERE f.last_scan_id = (SELECT MAX(id) FROM scans)
                    ORDER BY j.id ASC
                """)
                cur_json_rows = cur.fetchall()

        self._fill_files(self.cur_files, cur_files_rows)
        self._fill_scans(self.cur_scans, cur_scans_rows)
        self._fill_paths(self.cur_paths, cur_paths_rows)
        self._fill_json(self.cur_json, cur_json_rows)

        for r in cur_files_rows:
            self.records.append({
                "id":           r[0],
                "path":         r[1],
                "name":         r[2],
                "extension":    r[3] or "",
                "size_bytes":   r[4],
                "md5_hash":     r[5] or "",
                "created_at":   r[6].strftime("%Y-%m-%d %H:%M:%S") if r[6] else "",
                "modified_at":  r[7].strftime("%Y-%m-%d %H:%M:%S") if r[7] else "",
                "scanned_at":   r[8].strftime("%Y-%m-%d %H:%M:%S") if r[8] else "",
                "status":       r[9] or "",
                "change_count": r[10] if r[10] is not None else 0,
                "content":      r[11] or "",
            })

    # =================================================================
    #  ЗАГРУЗКА БД
    # =================================================================
    def _refresh_db(self) -> None:
        try:
            self._load_db_panel()
            total = len(self.all_files.get_children())
            self.status.set(f"БД обновлена. files: {total}")
        except Exception as e:
            self.status.set(f"Ошибка БД: {type(e).__name__}")
            raise

    def _load_db_panel(self) -> None:
        for tree in (self.all_files, self.all_scans, self.all_paths, self.all_json):
            for row in tree.get_children():
                tree.delete(row)

        with psycopg2.connect(DB_DSN) as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT
                        f.id,
                        p.full_path || '\\' || f.name AS full_path,
                        f.name,
                        f.extension,
                        f.size_bytes,
                        f.md5_hash,
                        f.created_at,
                        f.modified_at,
                        f.scanned_at,
                        f.status,
                        f.change_count,
                        f.content
                    FROM files f
                    JOIN paths p ON p.id = f.path_id
                    ORDER BY f.id ASC
                """)
                all_files_rows = cur.fetchall()

                cur.execute("""
                    SELECT id, root_path, started_at, finished_at,
                           files_total, files_new, files_unchanged,
                           files_modified, files_deleted
                    FROM scans
                    ORDER BY id ASC
                """)
                all_scans_rows = cur.fetchall()

                cur.execute("""
                    SELECT id, parent_id, name, full_path
                    FROM paths
                    ORDER BY id ASC
                """)
                all_paths_rows = cur.fetchall()

                cur.execute("""
                    SELECT id, file_id, updated_at, data::text
                    FROM json
                    ORDER BY id ASC
                """)
                all_json_rows = cur.fetchall()

        self._fill_files(self.all_files, all_files_rows)
        self._fill_scans(self.all_scans, all_scans_rows)
        self._fill_paths(self.all_paths, all_paths_rows)
        self._fill_json(self.all_json, all_json_rows)

    # =================================================================
    #  ЗАПОЛНЕНИЕ ТАБЛИЦ
    # =================================================================
    @staticmethod
    def _fill_files(tree, rows) -> None:
        for r in rows:
            rec = {
                "id":           r[0],
                "path":         r[1],
                "name":         r[2],
                "extension":    r[3] or "",
                "size_bytes":   r[4],
                "md5_hash":     r[5] or "",
                "created_at":   r[6].strftime("%Y-%m-%d %H:%M:%S") if r[6] else "",
                "modified_at":  r[7].strftime("%Y-%m-%d %H:%M:%S") if r[7] else "",
                "scanned_at":   r[8].strftime("%Y-%m-%d %H:%M:%S") if r[8] else "",
                "status":       r[9] or "",
                "change_count": r[10] if r[10] is not None else 0,
                "content":      (r[11] or "")[:300],
            }
            tree.insert("", END, values=tuple(rec.get(c[0], "") for c in FILES_COLUMNS))

    @staticmethod
    def _fill_scans(tree, rows) -> None:
        for r in rows:
            row = {
                "id":              r[0],
                "root_path":       r[1] or "",
                "started_at":      r[2].strftime("%Y-%m-%d %H:%M:%S") if r[2] else "",
                "finished_at":     r[3].strftime("%Y-%m-%d %H:%M:%S") if r[3] else "",
                "files_total":     r[4],
                "files_new":       r[5],
                "files_unchanged": r[6],
                "files_modified":  r[7],
                "files_deleted":   r[8],
            }
            tree.insert("", END, values=tuple(row.get(c[0], "") for c in SCANS_COLUMNS))

    @staticmethod
    def _fill_paths(tree, rows) -> None:
        for r in rows:
            row = {
                "id":        r[0],
                "parent_id": r[1] if r[1] is not None else "",
                "name":      r[2] or "",
                "full_path": r[3] or "",
            }
            tree.insert("", END, values=tuple(row.get(c[0], "") for c in PATHS_COLUMNS))

    @staticmethod
    def _fill_json(tree, rows) -> None:
        for r in rows:
            row = {
                "id":         r[0],
                "file_id":    r[1],
                "updated_at": r[2].strftime("%Y-%m-%d %H:%M:%S") if r[2] else "",
                "data":       (r[3] or "")[:1000],
            }
            tree.insert("", END, values=tuple(row.get(c[0], "") for c in JSON_COLUMNS))

    # =================================================================
    #  ОТЧЁТЫ
    # =

    # =================================================================
    #  ОТЧЁТЫ
    # =================================================================
    def _write_reports(self) -> None:
        self._write_scan_report()
        self._write_changes_report()

    def _write_scan_report(self) -> None:
        out_dir_str = self.scan_reports_dir.get().strip()
        if not out_dir_str:
            self.last_scan_report = None
            self.scan_report_btn.configure(
                text="Отчёт сканирования: папка не задана", state="disabled"
            )
            return

        try:
            out_dir = Path(out_dir_str)
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            self.last_scan_report = None
            self.scan_report_btn.configure(
                text=f"Отчёт сканирования: ошибка ({e})", state="disabled"
            )
            return

        now = datetime.now()
        filename = f"Отчёт сканирования от {now.strftime('%Y-%m-%d %H-%M-%S')}.txt"
        path = out_dir / filename

        summary = {k: 0 for k in STATUS_RU}
        for r in self.records:
            summary[r["status"]] = summary.get(r["status"], 0) + 1

        lines = []
        lines.append("=" * 78)
        lines.append("ОТЧЁТ О СКАНИРОВАНИИ ФАЙЛОВОГО ХРАНИЛИЩА")
        lines.append("=" * 78)
        lines.append(f"Дата и время:  {now.strftime('%d.%m.%Y %H:%M:%S')}")
        lines.append(f"Папка:         {self.folder.get()}")
        lines.append("")
        lines.append(
            f"Всего: {len(self.records)}  |  "
            f"новых: {summary['new']}  |  "
            f"без изменений: {summary['unchanged']}  |  "
            f"изменено: {summary['modified']}  |  "
            f"удалено: {summary['deleted']}"
        )
        lines.append("")
        lines.append("-" * 78)

        for i, r in enumerate(self.records, 1):
            lines.append(f"#{i}  [{r['status']}]  {r['name']}  (id={r['id']})")
            lines.append(f"  Путь:           {r['path']}")
            lines.append(f"  Расширение:     {r['extension']}")
            lines.append(f"  Размер:         {r['size_bytes']} байт")
            lines.append(f"  MD5:            {r['md5_hash']}")
            lines.append(f"  Создан:         {r['created_at']}")
            lines.append(f"  Изменён:        {r['modified_at']}")
            lines.append(f"  Просканирован:  {r['scanned_at']}")
            lines.append(f"  Изменений:      {r['change_count']}")
            lines.append("-" * 78)

        try:
            path.write_text("\n".join(lines), encoding="utf-8")
            self.last_scan_report = path
            self.scan_report_btn.configure(
                text=f"Отчёт сканирования: {path.name}", state="normal"
            )
        except OSError as e:
            self.last_scan_report = None
            self.scan_report_btn.configure(
                text=f"Отчёт сканирования: ошибка ({e})", state="disabled"
            )

    def _write_changes_report(self) -> None:
        """
        Отчёт ТОЛЬКО об изменённых файлах (status == 'modified').
        Новые и удалённые сюда НЕ попадают.
        Для каждого изменённого файла выводится только то, что изменилось:
            было:
            <старая строка>
            стало:
            <новая строка>
        Если изменённых нет — в файле строка «Файлы не изменялись».
        """
        out_dir_str = self.changes_reports_dir.get().strip()
        if not out_dir_str:
            self.last_changes_report = None
            self.changes_report_btn.configure(
                text="Отчёт изменений: папка не задана", state="disabled"
            )
            return

        try:
            out_dir = Path(out_dir_str)
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            self.last_changes_report = None
            self.changes_report_btn.configure(
                text=f"Отчёт изменений: ошибка ({e})", state="disabled"
            )
            return

        now = datetime.now()
        filename = f"Отчёт изменений от {now.strftime('%Y-%m-%d %H-%M-%S')}.txt"
        path = out_dir / filename

        modified_items = [r for r in self.records if r["status"] == "modified"]

        lines = []
        lines.append("=" * 78)
        lines.append("ОТЧЁТ ОБ ИЗМЕНЕНИЯХ ФАЙЛОВ")
        lines.append("(сравнение с предыдущим состоянием в базе данных)")
        lines.append("=" * 78)
        lines.append(f"Дата и время:  {now.strftime('%d.%m.%Y %H:%M:%S')}")
        lines.append(f"Папка:         {self.folder.get()}")
        lines.append("")

        if not modified_items:
            lines.append("Файлы не изменялись.")
        else:
            lines.append(f"Изменённых файлов: {len(modified_items)}")
            lines.append("")

            for i, r in enumerate(modified_items, 1):
                lines.append("=" * 78)
                lines.append(f"#{i}  {r['name']}  (id={r['id']})")
                lines.append("=" * 78)
                lines.append(f"  Путь:           {r['path']}")
                lines.append(f"  Расширение:     {r['extension']}")
                lines.append(f"  Размер:         {r['size_bytes']} байт")
                lines.append(f"  MD5:            {r['md5_hash']}")
                lines.append(f"  Изменён:        {r['modified_at']}")
                lines.append(f"  Всего изменений: {r['change_count']}")
                lines.append("")

                old_content = self.previous_content.get(r["path"], "")
                new_content = r.get("content", "") or ""

                if not old_content and not new_content:
                    lines.append(
                        "  (содержимое не извлекалось — сравнение по MD5 показало изменение)"
                    )
                else:
                    old_lines = old_content.splitlines()
                    new_lines = new_content.splitlines()

                    sm = difflib.SequenceMatcher(None, old_lines, new_lines)
                    changes_found = False

                    for tag, i1, i2, j1, j2 in sm.get_opcodes():
                        if tag == "equal":
                            continue  # неизменённые строки не выводим
                        changes_found = True

                        old_block = old_lines[i1:i2]
                        new_block = new_lines[j1:j2]

                        if old_block:
                            lines.append("  было:")
                            for ln in old_block:
                                lines.append(f"  {ln}")
                        if new_block:
                            lines.append("\n  стало:")
                            for ln in new_block:
                                lines.append(f"  {ln}")
                        lines.append("")  # пустая строка между блоками изменений

                    if not changes_found:
                        lines.append(
                            "  (текст не изменился, но изменились метаданные/бинарное содержимое)"
                        )
                lines.append("")

        try:
            path.write_text("\n".join(lines), encoding="utf-8")
            self.last_changes_report = path
            label = (
                f"Отчёт изменений: {path.name} ({len(modified_items)})"
                if modified_items
                else f"Отчёт изменений: {path.name} (изменений нет)"
            )
            self.changes_report_btn.configure(text=label, state="normal")
        except OSError as e:
            self.last_changes_report = None
            self.changes_report_btn.configure(
                text=f"Отчёт изменений: ошибка ({e})", state="disabled"
            )

    def _open_scan_report(self) -> None:
        self._open_file(self.last_scan_report)

    def _open_changes_report(self) -> None:
        self._open_file(self.last_changes_report)

    @staticmethod
    def _open_file(path: Path | None) -> None:
        if path is None or not path.exists():
            messagebox.showwarning("Файл не найден", "Отчёт ещё не создан или был перемещён.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror(
                "Ошибка открытия",
                f"Не удалось открыть файл:\n{path}\n\n{type(e).__name__}: {e}",
            )

    # =================================================================
    #  СОХРАНЕНИЕ / СБРОС НАСТРОЕК
    # =================================================================
    def _save_settings(self) -> None:
        self.settings["reports_scan_dir"] = self.scan_reports_dir.get().strip()
        self.settings["reports_changes_dir"] = self.changes_reports_dir.get().strip()
        self.settings["theme"] = "dark" if self.dark_theme.get() else "light"
        save_settings(self.settings)
        messagebox.showinfo("Настройки", "Настройки сохранены.")
        self.status.set("Настройки сохранены.")

    def _reset_settings(self) -> None:
        """
        Сбрасывает ТОЛЬКО папки для отчётов (делает их пустыми).
        Тема (светлая/тёмная) НЕ сбрасывается.
        """
        if not messagebox.askyesno(
            "Сброс настроек",
            "Очистить папки для отчётов?\n\n"
            "Тема оформления останется без изменений.",
        ):
            return

        self.scan_reports_dir.set("")
        self.changes_reports_dir.set("")

        self.settings["reports_scan_dir"] = ""
        self.settings["reports_changes_dir"] = ""
        # theme намеренно не трогаем
        save_settings(self.settings)

        self.status.set("Папки отчётов сброшены (тема сохранена).")
        messagebox.showinfo(
            "Сброс настроек",
            "Папки для отчётов очищены.\nТема оформления сохранена.",
        )


def main() -> None:
    root = Tk()
    try:
        root.tk.call("tk", "scaling", 1.3)
    except Exception:
        pass
    ScannerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

