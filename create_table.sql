-- create_table.sql
-- Скрипт создания таблицы files в PostgreSQL.
-- Порядок: создать базу file_storage, подключиться к ней, выполнить этот скрипт.

-- CREATE DATABASE file_storage;   -- выполняется отдельно, до подключения к базе

CREATE TABLE files (
    path        TEXT          PRIMARY KEY,
    name        VARCHAR(255)  NOT NULL,
    extension   VARCHAR(32),
    size_bytes  BIGINT        NOT NULL,
    sha256_hash CHAR(64)      NOT NULL,
    mime_type   VARCHAR(128),
    created_at  TIMESTAMPTZ   NOT NULL,
    modified_at TIMESTAMPTZ   NOT NULL,
    scanned_at  TIMESTAMPTZ   NOT NULL DEFAULT now(),
    content     TEXT,
    status      VARCHAR(16)   NOT NULL
        CHECK (status IN ('new', 'unchanged', 'modified', 'deleted'))
);
