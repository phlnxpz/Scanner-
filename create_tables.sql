-- ============================================================
--  База данных проекта «3var» (сканер файлов).
--  Скрипт создаёт базу данных и все таблицы. СУБД: PostgreSQL.
--
--  Как выполнить в pgAdmin 4:
--    ЧАСТЬ 1 - подключитесь к служебной базе «postgres»,
--              откройте Query Tool и выполните первую часть.
--    ЧАСТЬ 2 - подключитесь к базе file_storage
--              (правой кнопкой по базе -> Query Tool)
--              и выполните вторую часть.
--
--  Чтобы пересоздать базу с нуля, сначала выполните в базе postgres:
--      DROP DATABASE IF EXISTS file_storage WITH (FORCE);
-- ============================================================

-- ============================================================
-- ЧАСТЬ 1. Создание базы данных
-- ============================================================
CREATE DATABASE file_storage;

-- ============================================================
-- ЧАСТЬ 2. Таблицы и индексы
-- ============================================================
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
    created_at     TIMESTAMPTZ NOT NULL,                      -- дата создания файла
    modified_at    TIMESTAMPTZ NOT NULL,                      -- дата изменения файла
    scanned_at     TIMESTAMPTZ NOT NULL DEFAULT now(),        -- когда просканили
    content        TEXT,                                      -- текст из файла, если получилось достать
    change_count   INTEGER NOT NULL DEFAULT 0,                -- число изменений modified_at
    status         VARCHAR(16) NOT NULL
                   CHECK (status IN ('new', 'unchanged', 'modified', 'deleted')),
    first_scan_id  INTEGER REFERENCES scans(id),              -- в каком скане нашли в первый раз
    last_scan_id   INTEGER REFERENCES scans(id),              -- в каком скане обновили последний раз
    UNIQUE (path_id, name)                                    -- в одной папке два файла с одним именем нельзя
);

-- Таблица JSON: одна строка на каждый файл, полный объект метаданных в data.
CREATE TABLE IF NOT EXISTS json (
    id          BIGSERIAL PRIMARY KEY,
    file_id     BIGINT NOT NULL UNIQUE REFERENCES files(id) ON DELETE CASCADE,
    data        JSONB NOT NULL,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Миграция существующей базы
ALTER TABLE files ADD COLUMN IF NOT EXISTS change_count INTEGER NOT NULL DEFAULT 0;

-- Удаляем вторичные индексы; оставляем idx_paths_parent.
DROP INDEX IF EXISTS idx_files_status;
DROP INDEX IF EXISTS idx_files_hash;
DROP INDEX IF EXISTS idx_files_scan;

-- ============================================================
-- ЧАСТЬ 3. Проверка (можно выполнить после создания)
-- ============================================================
-- список таблиц
SELECT table_name FROM information_schema.tables
WHERE table_schema = 'public' ORDER BY table_name;

-- структура таблицы files
SELECT column_name, data_type FROM information_schema.columns
WHERE table_schema = 'public' AND table_name = 'files' ORDER BY ordinal_position;
