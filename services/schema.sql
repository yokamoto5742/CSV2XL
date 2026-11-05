-- 医療文書担当一覧のスキーマ(C# 版でも同じ SQL を使う)
-- 日付は 'YYYY-MM-DD'、日時は 'YYYY-MM-DDThh:mm:ss' の文字列で保存する

CREATE TABLE IF NOT EXISTS documents (
    id             INTEGER PRIMARY KEY,
    received_date  TEXT    NOT NULL,  -- 預り日
    patient_id     INTEGER NOT NULL,  -- 患者ID
    patient_name   TEXT    NOT NULL,  -- 患者名
    document_name  TEXT    NOT NULL,  -- 文書名
    department     TEXT    NOT NULL,  -- 診療科
    doctor_name    TEXT    NOT NULL,  -- 医師名
    staff_name     TEXT,              -- 担当者名
    requested_date TEXT,              -- 医師依頼日
    note           TEXT,              -- 備考
    version        INTEGER NOT NULL DEFAULT 1,  -- 楽観ロック用。更新のたびに1増やす
    deleted        INTEGER NOT NULL DEFAULT 0,  -- 1 なら削除済み(CSV の重複チェックに使うため行は残す)
    created_at     TEXT    NOT NULL,
    updated_at     TEXT    NOT NULL
);

-- 同じ内容の行が実データに存在するため、A〜F 列の組み合わせは UNIQUE にしない
CREATE INDEX IF NOT EXISTS ix_documents_key
    ON documents (received_date, patient_id, document_name);
CREATE INDEX IF NOT EXISTS ix_documents_patient_id
    ON documents (patient_id);

-- 全端末で共有するリスト(category: staff / exclude_doc / exclude_doctor)
CREATE TABLE IF NOT EXISTS list_items (
    id         INTEGER PRIMARY KEY,
    category   TEXT    NOT NULL,
    name       TEXT    NOT NULL,
    sort_order INTEGER NOT NULL,
    UNIQUE (category, name)
);

PRAGMA user_version = 2;
