-- таблица сканирований, одна строка = один запуск сканера
CREATE TABLE IF NOT EXISTS scans (
    id               SERIAL PRIMARY KEY,
    root_path        TEXT NOT NULL,                      -- какую папку сканили
    started_at       TIMESTAMPTZ NOT NULL DEFAULT now(), -- когда начали
    finished_at      TIMESTAMPTZ,                        -- когда закончили, пока идет скан будет пусто
    files_total      INTEGER NOT NULL DEFAULT 0,         -- всего файлов
    files_new        INTEGER NOT NULL DEFAULT 0,         -- новых
    files_unchanged  INTEGER NOT NULL DEFAULT 0,         -- без изменений
    files_modified   INTEGER NOT NULL DEFAULT 0,         -- измененных
    files_deleted    INTEGER NOT NULL DEFAULT 0          -- удаленных
);

-- таблица путей (папки)
CREATE TABLE IF NOT EXISTS paths (
    id         BIGSERIAL PRIMARY KEY,
    parent_id  BIGINT REFERENCES paths(id) ON DELETE CASCADE, -- родительская папка, у корня пусто
    name       VARCHAR(255) NOT NULL,                         -- имя папки
    full_path  TEXT NOT NULL UNIQUE                           -- полный путь, не должен повторяться
);

-- индекс по parent_id
-- нужен чтобы быстро находить все вложенные папки внутри какой-то папки
-- без него база при каждом таком запросе читала бы всю таблицу paths
CREATE INDEX IF NOT EXISTS idx_paths_parent ON paths(parent_id);

-- таблица файлов
CREATE TABLE IF NOT EXISTS files (
    id             BIGSERIAL PRIMARY KEY,
    path_id        BIGINT NOT NULL REFERENCES paths(id) ON DELETE CASCADE, -- в какой папке лежит
    name           VARCHAR(255) NOT NULL,
    extension      VARCHAR(32),
    size_bytes     BIGINT NOT NULL,
    md5_hash       CHAR(32),
    sha256_hash    CHAR(64) NOT NULL,
    created_at     TIMESTAMPTZ NOT NULL,                      -- дата создания файла
    modified_at    TIMESTAMPTZ NOT NULL,                      -- дата изменения файла
    scanned_at     TIMESTAMPTZ NOT NULL DEFAULT now(),        -- когда просканили
    content        TEXT,                                      -- текст из файла, если получилось достать
    status         VARCHAR(16) NOT NULL
                   CHECK (status IN ('new', 'unchanged', 'modified', 'deleted')),
    first_scan_id  INTEGER REFERENCES scans(id),              -- в каком скане нашли в первый раз
    last_scan_id   INTEGER REFERENCES scans(id),              -- в каком скане обновили последний раз
    UNIQUE (path_id, name)                                    -- в одной папке два файла с одним именем нельзя
);

-- индексы для таблицы files
-- они не обязательны, скан и без них работает, просто запросы на большом
-- количестве файлов будут быстрее. минус: запись в таблицу чуть медленнее

-- по статусу
-- для запросов типа "покажи все измененные" или "все удаленные" (where status = 'modified')
CREATE INDEX IF NOT EXISTS idx_files_status ON files(status);

-- по хэшу
-- чтобы быстро искать дубликаты, у одинаковых файлов хэш одинаковый
-- (select sha256_hash, count(*) from files group by sha256_hash having count(*) > 1)
CREATE INDEX IF NOT EXISTS idx_files_hash   ON files(sha256_hash);

-- по последнему скану
-- чтобы быстро доставать файлы из конкретного сканирования
-- (where last_scan_id = 5)
CREATE INDEX IF NOT EXISTS idx_files_scan   ON files(last_scan_id);
