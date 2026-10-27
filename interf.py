# gui.py
from __future__ import annotations  # позволяет писать типы "как строки", чтобы не было ошибок с порядком классов

import os  # работа с переменными окружения (например, передать env дочернему процессу)
import queue  # очередь: складываем сообщения из фонового потока, забираем в главном
import subprocess  # запуск main.py как отдельного процесса
import sys  # доступ к sys.executable - пути к текущему Python
import threading  # отдельный поток, чтобы чтение вывода не вешало окно
from datetime import datetime  # берём текущие дату и время для имени отчёта
from pathlib import Path  # удобная работа с путями к файлам и папкам
from tkinter import Tk, StringVar, END  # главное окно, строковая переменная, конец списка
from tkinter import ttk, filedialog, messagebox  # виджеты, диалог выбора папки, всплывающие окна

import psycopg2  # драйвер PostgreSQL - через него читаем таблицы для показа

APP_TITLE = "Сканер файлов"  # заголовок главного окна
MAIN_SCRIPT = Path(__file__).with_name("main.py")  # путь к main.py - он лежит в той же папке, что и gui.py

# Тот же DSN, что в main.py. Можно переопределить через переменную окружения DB_DSN.
DB_DSN = os.environ.get(  # читаем переменную окружения DB_DSN
    "DB_DSN",  # имя переменной
    "host=localhost port=5432 dbname=file_storage user=postgres password=12345",  # значение по умолчанию: локальная база
)

# Папка для TXT-отчётов (создаётся автоматически, если её нет)
REPORTS_DIR = Path(r"C:\Users\katana_17\OneDrive\Desktop\ХУЙНЯ")

# ---- Описание колонок для таблицы files: (ключ, русская подпись, ширина в пикселях) ----
FILES_COLUMNS = (
    ("id",           "id",              70),  # id записи в таблице files
    ("path",         "Путь",           300),  # полный путь: папка + имя
    ("name",         "Имя",            170),  # имя файла
    ("extension",    "Расширение",      90),  # расширение (.txt, .pdf, ...)
    ("size_bytes",   "Размер (байт)",  110),  # размер в байтах
    ("md5_hash",     "MD5",            250),  # контрольная сумма MD5
    ("created_at",   "Создан",         150),  # дата создания файла
    ("modified_at",  "Изменён",        150),  # дата последнего изменения
    ("scanned_at",   "Просканирован",  150),  # когда программа последний раз видела файл
    ("status",       "Статус",         110),  # new / unchanged / modified / deleted
    ("change_count", "Изменений",       90),  # сколько раз файл менялся
    ("content",      "Содержимое",     300),  # первые 300 символов извлечённого текста
)

# ---- Описание колонок для таблицы scans ----
SCANS_COLUMNS = (
    ("id",               "id",               60),  # id сканирования
    ("root_path",        "Корень скана",    320),  # какую папку сканировали
    ("started_at",       "Начат",           150),  # когда начали
    ("finished_at",      "Завершён",        150),  # когда закончили
    ("files_total",      "Всего",            80),  # сколько файлов всего
    ("files_new",        "Новых",            80),  # сколько новых
    ("files_unchanged",  "Без изм.",         90),  # сколько не менялось
    ("files_modified",   "Изменено",         90),  # сколько изменилось
    ("files_deleted",    "Удалено",          90),  # сколько удалено
)

# ---- Описание колонок для таблицы paths ----
PATHS_COLUMNS = (
    ("id",         "id",          60),  # id папки
    ("parent_id",  "parent_id",   90),  # id родительской папки (у корня пусто)
    ("name",       "Имя",        200),  # имя папки
    ("full_path",  "Полный путь", 500),  # полный путь к папке
)

# ---- Описание колонок для таблицы json ----
JSON_COLUMNS = (
    ("id",          "id",          60),  # id записи
    ("file_id",     "file_id",     80),  # ссылка на files.id
    ("updated_at",  "Обновлён",   150),  # когда обновили JSON
    ("data",        "data",      1000),  # сам JSON (первые 1000 символов)
)

# Русские подписи статусов (используются в TXT-отчёте)
STATUS_RU = {
    "new": "НОВЫЙ",              # файл встречен впервые
    "unchanged": "БЕЗ ИЗМЕНЕНИЙ",  # хэш не изменился
    "modified": "ИЗМЕНЁН",        # хэш другой
    "deleted": "УДАЛЁН",          # файла больше нет на диске
}


class ScannerGUI:  # главный класс интерфейса
    def __init__(self, root: Tk) -> None:  # конструктор: получает главное окно
        self.root = root  # запоминаем окно
        self.root.title(APP_TITLE)  # ставим заголовок окна
        self.root.geometry("1600x800")  # размер окна: 1600 на 800

        self.folder = StringVar(value="")  # путь к папке - ПУСТОЙ по умолчанию
        self.process: subprocess.Popen | None = None  # тут хранится запущенный main.py (или None)
        self.log_queue: queue.Queue[str | None] = queue.Queue()  # очередь сообщений из фонового потока
        self.records: list[dict] = []  # записи текущего скана - нужны для отчёта

        self._build_ui()  # строим интерфейс
        self._poll_log()  # запускаем цикл чтения очереди

        # ---- При запуске заполняем ТОЛЬКО правую панель (вся БД) ----
        self.root.after(50, self._load_right_on_start)  # через 50 мс вызвать загрузку правой панели

    # ---------- ПОСТРОЕНИЕ ИНТЕРФЕЙСА ----------
    def _build_ui(self) -> None:  # собираем все элементы окна
        # --- верхняя панель: путь и кнопка выбора папки ---
        top = ttk.Frame(self.root, padding=10)  # рамка с отступами
        top.pack(fill="x")  # растягиваем по горизонтали

        ttk.Label(top, text="Папка:").grid(row=0, column=0, sticky="w")  # подпись "Папка:"
        ttk.Entry(top, textvariable=self.folder, width=70).grid(  # поле ввода пути (пустое по умолчанию)
            row=0, column=1, padx=5, sticky="we"  # справа от подписи, тянется по ширине
        )
        ttk.Button(top, text="Обзор…", command=self._choose_folder).grid(row=0, column=2)  # кнопка выбора папки
        top.columnconfigure(1, weight=1)  # поле ввода занимает всё свободное место

        # --- панель кнопок ---
        buttons = ttk.Frame(self.root, padding=(10, 0))  # рамка для кнопок
        buttons.pack(fill="x")  # растягиваем по горизонтали

        self.btn_start = ttk.Button(buttons, text="▶ Запустить скан", command=self._start_scan)  # кнопка "Запустить скан"
        self.btn_start.pack(side="left")  # прижимаем влево

        self.btn_stop = ttk.Button(  # кнопка "Остановить"
            buttons, text="■ Остановить", command=self._stop_scan, state="disabled"  # пока недоступна
        )
        self.btn_stop.pack(side="left", padx=6)  # влево, с отступом

        # Очистка — только левой панели
        ttk.Button(buttons, text="Очистить", command=self._clear_left).pack(side="left", padx=6)  # кнопка "Очистить"

        # Обновление — только правой панели
        ttk.Button(  # кнопка "Обновить БД"
            buttons, text="⟳ Обновить БД", command=self._refresh_right  # вызывает _refresh_right
        ).pack(side="left", padx=6)  # влево, с отступом

        self.status = StringVar(value="Готов к работе")  # строка статуса (справа)
        ttk.Label(buttons, textvariable=self.status).pack(side="right")  # прижимаем вправо

        # --- прогресс-бар (неопределённый: точный прогресс узнать сложно) ---
        self.progress = ttk.Progressbar(self.root, mode="indeterminate")  # полоса-анимация
        self.progress.pack(fill="x", padx=10, pady=(6, 0))  # растягиваем по ширине

        # --- основная область: слева текущий скан, справа вся БД ---
        paned = ttk.Panedwindow(self.root, orient="horizontal")  # разделитель: две панели рядом
        paned.pack(fill="both", expand=True, padx=10, pady=10)  # растягиваем на всё свободное место

        # Левая панель: 4 вкладки с данными текущего скана
        left_frame = ttk.Frame(paned)  # рамка левой панели
        ttk.Label(left_frame, text="ТЕКУЩИЙ СКАН", font=("", 10, "bold")).pack(anchor="w")  # заголовок
        self.notebook_current = ttk.Notebook(left_frame)  # блокнот с вкладками (левая панель)
        self.notebook_current.pack(fill="both", expand=True)  # растягиваем

        # Правая панель: 4 вкладки со всеми данными из БД
        right_frame = ttk.Frame(paned)  # рамка правой панели
        ttk.Label(right_frame, text="ВСЯ БАЗА ДАННЫХ", font=("", 10, "bold")).pack(anchor="w")  # заголовок
        self.notebook_all = ttk.Notebook(right_frame)  # блокнот с вкладками (правая панель)
        self.notebook_all.pack(fill="both", expand=True)  # растягиваем

        paned.add(left_frame, weight=1)  # добавляем левую панель, обе равной ширины
        paned.add(right_frame, weight=1)  # добавляем правую панель

        # --- создаём по 4 таблицы в каждой панели ---
        # Слева - текущий скан
        self.cur_files = self._make_tab(self.notebook_current, "files", FILES_COLUMNS)  # вкладка files слева
        self.cur_scans = self._make_tab(self.notebook_current, "scans", SCANS_COLUMNS)  # вкладка scans слева
        self.cur_paths = self._make_tab(self.notebook_current, "paths", PATHS_COLUMNS)  # вкладка paths слева
        self.cur_json  = self._make_tab(self.notebook_current, "json",  JSON_COLUMNS)   # вкладка json слева

        # Справа - вся база
        self.all_files = self._make_tab(self.notebook_all, "files", FILES_COLUMNS)  # вкладка files справа
        self.all_scans = self._make_tab(self.notebook_all, "scans", SCANS_COLUMNS)  # вкладка scans справа
        self.all_paths = self._make_tab(self.notebook_all, "paths", PATHS_COLUMNS)  # вкладка paths справа
        self.all_json  = self._make_tab(self.notebook_all, "json",  JSON_COLUMNS)   # вкладка json справа

    def _make_tab(self, notebook: ttk.Notebook, title: str, columns: tuple) -> ttk.Treeview:  # создаёт вкладку с таблицей
        """Создаёт вкладку с таблицей и возвращает Treeview."""
        frame = ttk.Frame(notebook)  # новая рамка - это будет вкладка
        notebook.add(frame, text=title)  # добавляем рамку как вкладку с названием title

        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings")  # сама таблица; колонки из списка

        for key, col_title, width in columns:  # перебираем все колонки
            tree.heading(key, text=col_title)  # русская подпись колонки
            tree.column(  # настраиваем колонку
                key, width=width,  # ширина
                stretch=(key in ("path", "name", "content", "data", "full_path", "root_path")),  # эти растягиваются, остальные фиксированы
            )

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)  # вертикальный скроллбар
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)  # горизонтальный скроллбар
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)  # связываем скроллбары с таблицей

        tree.grid(row=0, column=0, sticky="nsew")  # таблица занимает основную ячейку
        vsb.grid(row=0, column=1, sticky="ns")  # вертикальный скролл справа
        hsb.grid(row=1, column=0, sticky="we")  # горизонтальный скролл снизу
        frame.rowconfigure(0, weight=1)  # таблица растягивается по высоте
        frame.columnconfigure(0, weight=1)  # и по ширине

        return tree  # возвращаем таблицу - пригодится снаружи

    # ---------- ДЕЙСТВИЯ ----------
    def _choose_folder(self) -> None:  # открыть диалог выбора папки
        """Диалог выбора папки."""
        chosen = filedialog.askdirectory(initialdir=self.folder.get() or None)  # открываем выбор папки
        if chosen:  # если пользователь что-то выбрал
            self.folder.set(chosen)  # подставляем путь в поле ввода

    def _clear_left(self) -> None:  # очистить только левую панель
        """Очистить ТОЛЬКО левую панель (текущий скан). Правую не трогаем."""
        for tree in (self.cur_files, self.cur_scans, self.cur_paths, self.cur_json):  # проходим по 4 левым таблицам
            for row in tree.get_children():  # берём все строки
                tree.delete(row)  # и удаляем их
        self.records.clear()  # чистим список записей для отчёта
        self.status.set("Левая панель очищена")  # пишем статус

    def _refresh_right(self) -> None:  # обновить только правую панель
        """Обновить ТОЛЬКО правую панель (вся БД). Левую не трогаем."""
        try:  # пытаемся
            self._load_right_panel()  # перечитываем правую панель из БД
            total = len(self.all_files.get_children())  # считаем, сколько строк в правой files
            self.status.set(f"Правая панель обновлена из БД. files: {total}")  # пишем статус
        except Exception as e:  # если что-то сломалось
            messagebox.showerror(  # показываем ошибку
                "Ошибка обновления БД",  # заголовок
                f"{type(e).__name__}: {e}\n\nDSN: {DB_DSN}",  # текст ошибки и DSN
            )

    def _start_scan(self) -> None:  # запустить сканирование
        """Запустить main.py в отдельном процессе с папкой из поля ввода."""
        if self.process is not None:  # если уже что-то запущено
            return  # ничего не делаем

        folder = self.folder.get().strip()  # берём путь из поля, убираем пробелы по краям

        # ---- ПРОВЕРКА: пользователь нажал "Запустить скан", но путь не выбран ----
        if not folder:  # если поле пустое
            messagebox.showerror(  # показываем ошибку ТОЛЬКО в этот момент
                "Ошибка",
                "Не выбрана папка для сканирования.\nНажмите «Обзор…» и выберите папку.",
            )
            return  # выходим, ничего не запускаем

        if not Path(folder).exists():  # если папки по указанному пути нет
            messagebox.showerror("Ошибка", f"Папка не найдена:\n{folder}")  # тоже ошибка
            return  # выходим

        self._clear_left()  # чистим левую панель (правую НЕ трогаем)
        self.progress.start(12)  # запускаем анимацию прогресса
        self.btn_start.configure(state="disabled")  # кнопку "Запустить" отключаем
        self.btn_stop.configure(state="normal")  # кнопку "Остановить" включаем
        self.status.set(f"Сканирую: {folder}")  # пишем, что сканируем

        cmd = [sys.executable, "-u", str(MAIN_SCRIPT), folder]  # команда: python -u main.py "папка"
        env = {"ENABLE_DB_SYNC_OVERRIDE": "1"}  # всегда пишем в БД

        self.process = subprocess.Popen(  # запускаем main.py отдельным процессом
            cmd,  # команда
            stdout=subprocess.PIPE,  # забираем stdout
            stderr=subprocess.STDOUT,  # stderr тоже в stdout
            text=True,  # работаем как с текстом, а не байтами
            encoding="utf-8",  # кодировка вывода
            errors="replace",  # битые байты не ломают чтение
            bufsize=1,  # построчная буферизация
            env={**os.environ, **env},  # окружение: текущее + наш флаг
        )
        threading.Thread(target=self._reader, args=(self.process,), daemon=True).start()  # читаем вывод в отдельном потоке

    def _stop_scan(self) -> None:  # остановить сканирование
        """Остановить процесс сканирования."""
        if self.process and self.process.poll() is None:  # если процесс жив
            self.process.terminate()  # просим его завершиться
            self.status.set("Останавливается…")  # пишем статус

    def _reader(self, proc: subprocess.Popen) -> None:  # фоновый поток: читает вывод процесса
        """Фоновый поток: читает stdout процесса и складывает в очередь."""
        assert proc.stdout is not None  # гарантируем, что stdout есть
        for line in proc.stdout:  # читаем построчно
            line = line.rstrip()  # убираем \n в конце
            if line.startswith("[СТАРТ]") or line.startswith("[СНИМОК]"):  # нас интересуют только служебные строки
                self.log_queue.put(line)  # складываем в очередь
        proc.wait()  # ждём завершения процесса
        self.log_queue.put(None)  # отправляем None как сигнал "всё, закончили"

    # ---------- ОБРАБОТКА ОЧЕРЕДИ В UI ----------
    def _poll_log(self) -> None:  # каждые 100 мс проверяем очередь
        """Периодически (каждые 100 мс) проверяем очередь и обновляем UI."""
        try:  # пытаемся
            while True:  # в цикле, пока очередь не опустеет
                item = self.log_queue.get_nowait()  # берём сообщение без ожидания
                if item is None:  # если это сигнал завершения
                    self._on_finished()  # обрабатываем завершение процесса
                    continue  # и идём дальше
                if isinstance(item, str):  # если это строка
                    self.status.set(item)  # показываем её в статусе
        except queue.Empty:  # если очередь пуста
            pass  # ничего не делаем
        self.root.after(100, self._poll_log)  # перезапускаем себя через 100 мс

    def _load_right_on_start(self) -> None:  # при запуске программы
        """При запуске: заполняем ТОЛЬКО правую панель (вся БД)."""
        try:  # пытаемся
            self._load_right_panel()  # читаем правую панель из БД
            self.status.set("Правая панель загружена из БД. Готов к сканированию.")  # пишем статус
        except Exception as e:  # если БД недоступна
            self.status.set("БД недоступна — правая панель пуста")  # пишем в статус
            messagebox.showwarning(  # показываем предупреждение
                "База данных недоступна",  # заголовок
                f"Не удалось прочитать БД при запуске:\n"  # текст
                f"{type(e).__name__}: {e}\n\n"  # тип ошибки и её текст
                f"DSN: {DB_DSN}\n\n"  # DSN
                f"Окно откроется, правая панель будет пустой.",  # что будет дальше
            )

    def _on_finished(self) -> None:  # после завершения main.py
        """Вызывается после завершения процесса main.py: заполняем ЛЕВУЮ панель."""
        self.progress.stop()  # останавливаем анимацию прогресса
        self.btn_start.configure(state="normal")  # включаем кнопку "Запустить"
        self.btn_stop.configure(state="disabled")  # отключаем кнопку "Остановить"
        self.process = None  # обнуляем ссылку на процесс

        try:  # пытаемся
            self._load_left_panel()  # читаем левую панель из БД (текущий скан)
        except Exception as e:  # если ошибка
            messagebox.showerror(  # показываем её
                "Ошибка чтения БД",  # заголовок
                f"Не удалось прочитать текущий скан:\n{type(e).__name__}: {e}\n\nDSN: {DB_DSN}",  # текст
            )
            self.status.set("Готово (ошибка чтения БД)")  # статус
            return  # выходим

        if self.records:  # если есть записи
            report = self._write_report()  # создаём TXT-отчёт
            self.status.set(f"Готово. Записей: {len(self.records)}. Отчёт: {report}")  # пишем статус
        else:  # если записей нет
            self.status.set("Готово (в files ничего не найдено)")  # пишем статус

    # ---------- ЧТЕНИЕ ИЗ POSTGRESQL ----------
    @staticmethod
    def _fill_files(tree, rows):  # заполняет таблицу files строками из БД
        for r in rows:  # перебираем строки результата SQL
            rec = {  # собираем словарь для одной строки
                "id":           r[0],  # id
                "path":         r[1],  # полный путь
                "name":         r[2],  # имя файла
                "extension":    r[3] or "",  # расширение (или пусто)
                "size_bytes":   r[4],  # размер
                "md5_hash":     r[5] or "",  # md5 (или пусто)
                "created_at":   r[6].strftime("%Y-%m-%d %H:%M:%S") if r[6] else "",  # дата создания текстом
                "modified_at":  r[7].strftime("%Y-%m-%d %H:%M:%S") if r[7] else "",  # дата изменения текстом
                "scanned_at":   r[8].strftime("%Y-%m-%d %H:%M:%S") if r[8] else "",  # момент скана текстом
                "status":       r[9] or "",  # статус
                "change_count": r[10] if r[10] is not None else 0,  # счётчик изменений
                "content":      (r[11] or "")[:300],  # первые 300 символов содержимого
            }
            tree.insert("", END, values=tuple(rec.get(c[0], "") for c in FILES_COLUMNS))  # добавляем строку в таблицу

    @staticmethod
    def _fill_scans(tree, rows):  # заполняет таблицу scans
        for r in rows:  # перебираем строки
            row = {  # собираем словарь
                "id":              r[0],  # id скана
                "root_path":       r[1] or "",  # корень
                "started_at":      r[2].strftime("%Y-%m-%d %H:%M:%S") if r[2] else "",  # начало
                "finished_at":     r[3].strftime("%Y-%m-%d %H:%M:%S") if r[3] else "",  # конец
                "files_total":     r[4],  # всего файлов
                "files_new":       r[5],  # новых
                "files_unchanged": r[6],  # без изменений
                "files_modified":  r[7],  # изменённых
                "files_deleted":   r[8],  # удалённых
            }
            tree.insert("", END, values=tuple(row.get(c[0], "") for c in SCANS_COLUMNS))  # добавляем строку

    @staticmethod
    def _fill_paths(tree, rows):  # заполняет таблицу paths
        for r in rows:  # перебираем строки
            row = {  # собираем словарь
                "id":        r[0],  # id папки
                "parent_id": r[1] if r[1] is not None else "",  # родитель (или пусто)
                "name":      r[2] or "",  # имя
                "full_path": r[3] or "",  # полный путь
            }
            tree.insert("", END, values=tuple(row.get(c[0], "") for c in PATHS_COLUMNS))  # добавляем строку

    @staticmethod
    def _fill_json(tree, rows):  # заполняет таблицу json
        for r in rows:  # перебираем строки
            row = {  # собираем словарь
                "id":         r[0],  # id
                "file_id":    r[1],  # file_id
                "updated_at": r[2].strftime("%Y-%m-%d %H:%M:%S") if r[2] else "",  # когда обновили
                "data":       (r[3] or "")[:1000],  # JSON (первые 1000 символов)
            }
            tree.insert("", END, values=tuple(row.get(c[0], "") for c in JSON_COLUMNS))  # добавляем строку

    def _load_left_panel(self) -> None:  # заполняет левую панель (текущий скан)
        """Заполняет ЛЕВУЮ панель: текущий скан (последний scans.id)."""
        for tree in (self.cur_files, self.cur_scans, self.cur_paths, self.cur_json):  # проходим по 4 левым таблицам
            for row in tree.get_children():  # берём все строки
                tree.delete(row)  # и удаляем
        self.records.clear()  # чистим список записей для отчёта

        with psycopg2.connect(DB_DSN) as conn:  # открываем соединение с БД (закроется само)
            with conn.cursor() as cur:  # открываем курсор (закроется сам)
                # ВАЖНО: комментарии питона внутри SQL-запросов ЗАПРЕЩЕНЫ (будет ошибка базы).
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
                cur_files_rows = cur.fetchall()  # забираем все строки

                cur.execute("""
                    SELECT id, root_path, started_at, finished_at,
                           files_total, files_new, files_unchanged,
                           files_modified, files_deleted
                    FROM scans
                    WHERE id = (SELECT MAX(id) FROM scans)
                    ORDER BY id ASC
                """)
                cur_scans_rows = cur.fetchall()  # забираем все строки

                cur.execute("""
                    SELECT DISTINCT p.id, p.parent_id, p.name, p.full_path
                    FROM paths p
                    WHERE p.id IN (
                        SELECT f.path_id FROM files f
                        WHERE f.last_scan_id = (SELECT MAX(id) FROM scans)
                    )
                    ORDER BY p.id ASC
                """)
                cur_paths_rows = cur.fetchall()  # забираем все строки

                cur.execute("""
                    SELECT j.id, j.file_id, j.updated_at, j.data::text
                    FROM json j
                    JOIN files f ON f.id = j.file_id
                    WHERE f.last_scan_id = (SELECT MAX(id) FROM scans)
                    ORDER BY j.id ASC
                """)
                cur_json_rows = cur.fetchall()  # забираем все строки

        self._fill_files(self.cur_files, cur_files_rows)  # заполняем files
        self._fill_scans(self.cur_scans, cur_scans_rows)  # заполняем scans
        self._fill_paths(self.cur_paths, cur_paths_rows)  # заполняем paths
        self._fill_json(self.cur_json, cur_json_rows)  # заполняем json

        for r in cur_files_rows:  # для каждой строки files
            self.records.append({  # добавляем в список записей (для отчёта)
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
            })

    def _load_right_panel(self) -> None:  # заполняет правую панель (вся БД)
        """Заполняет ПРАВУЮ панель: вся БД, без лимитов."""
        for tree in (self.all_files, self.all_scans, self.all_paths, self.all_json):  # проходим по 4 правым таблицам
            for row in tree.get_children():  # берём все строки
                tree.delete(row)  # и удаляем

        with psycopg2.connect(DB_DSN) as conn:  # открываем соединение
            with conn.cursor() as cur:  # открываем курсор
                # ВАЖНО: комментарии питона внутри SQL-запросов ЗАПРЕЩЕНЫ (будет ошибка базы).
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
                        f.content                    FROM files f
                    JOIN paths p ON p.id = f.path_id
                    ORDER BY f.id ASC
                """)
                all_files_rows = cur.fetchall()  # забираем все строки

                cur.execute("""
                    SELECT id, root_path, started_at, finished_at,
                           files_total, files_new, files_unchanged,
                           files_modified, files_deleted
                    FROM scans
                    ORDER BY id ASC
                """)
                all_scans_rows = cur.fetchall()  # забираем все строки

                cur.execute("""
                    SELECT id, parent_id, name, full_path
                    FROM paths
                    ORDER BY id ASC
                """)
                all_paths_rows = cur.fetchall()  # забираем все строки

                cur.execute("""
                    SELECT id, file_id, updated_at, data::text
                    FROM json
                    ORDER BY id ASC
                """)
                all_json_rows = cur.fetchall()  # забираем все строки

        self._fill_files(self.all_files, all_files_rows)  # заполняем files
        self._fill_scans(self.all_scans, all_scans_rows)  # заполняем scans
        self._fill_paths(self.all_paths, all_paths_rows)  # заполняем paths
        self._fill_json(self.all_json, all_json_rows)  # заполняем json

    # ---------- TXT-ОТЧЁТ ----------
    def _write_report(self) -> Path:  # создаёт TXT-отчёт
        """Создаёт TXT-отчёт о текущем скане в REPORTS_DIR."""
        now = datetime.now()  # текущая дата и время
        filename = (  # собираем имя файла
            f"Сканирование файлового хранилища от "
            f"{now.strftime('%Y-%m-%d %H-%M-%S')}.txt"  # дата и время, двоеточия заменены на "-"
        )

        try:  # пытаемся писать в REPORTS_DIR
            REPORTS_DIR.mkdir(parents=True, exist_ok=True)  # создаём папку, если её нет
            report_path = REPORTS_DIR / filename  # полный путь к отчёту
        except OSError:  # если не получилось
            report_path = Path(self.folder.get()) / filename  # пишем в папку скана
            try:  # пытаемся создать папку
                report_path.parent.mkdir(parents=True, exist_ok=True)
            except OSError:  # если и тут не вышло
                report_path = Path.cwd() / filename  # пишем в текущую папку

        summary = {k: 0 for k in STATUS_RU}  # счётчики статусов, изначально нули
        for r in self.records:  # перебираем записи
            summary[r["status"]] = summary.get(r["status"], 0) + 1  # увеличиваем счётчик нужного статуса

        lines: list[str] = []  # сюда собираем строки отчёта
        lines.append("=" * 78)  # разделитель
        lines.append("ОТЧЁТ О СКАНИРОВАНИИ ФАЙЛОВОГО ХРАНИЛИЩА")  # заголовок
        lines.append("=" * 78)  # разделитель
        lines.append(f"Дата и время:  {now.strftime('%d.%m.%Y %H:%M:%S')}")  # когда сделали отчёт
        lines.append(f"Папка:         {self.folder.get()}")  # какую папку сканировали
        lines.append("")  # пустая строка
        lines.append(  # строка сводки
            f"Всего: {len(self.records)}  |  "
            f"новых: {summary['new']}  |  "
            f"без изменений: {summary['unchanged']}  |  "
            f"изменено: {summary['modified']}  |  "
            f"удалено: {summary['deleted']}"
        )
        lines.append("")  # пустая строка
        lines.append("-" * 78)  # разделитель

        for i, r in enumerate(self.records, 1):  # перебираем записи с номерами с 1
            lines.append(f"#{i}  [{r['status']}]  {r['name']}  (id={r['id']})")  # номер, статус, имя, id
            lines.append(f"  Путь:           {r['path']}")  # путь
            lines.append(f"  Расширение:     {r['extension']}")  # расширение
            lines.append(f"  Размер:         {r['size_bytes']} байт")  # размер
            lines.append(f"  MD5:            {r['md5_hash']}")  # md5
            lines.append(f"  Создан:         {r['created_at']}")  # создан
            lines.append(f"  Изменён:        {r['modified_at']}")  # изменён
            lines.append(f"  Просканирован:  {r['scanned_at']}")  # просканирован
            lines.append(f"  Изменений:      {r['change_count']}")  # сколько раз менялся
            lines.append("-" * 78)  # разделитель между записями

        report_path.write_text("\n".join(lines), encoding="utf-8")  # пишем всё в файл в кодировке utf-8
        return report_path  # возвращаем путь к файлу


def main() -> None:  # точка входа
    root = Tk()  # создаём главное окно
    try:  # пытаемся
        root.tk.call("tk", "scaling", 1.3)  # чуть крупнее шрифт на HiDPI-мониторах
    except Exception:  # если не получилось
        pass  # ничего страшного
    ScannerGUI(root)  # создаём GUI
    root.mainloop()  # запускаем главный цикл окна


if __name__ == "__main__":  # если файл запущен напрямую
    main()  # запускаем main