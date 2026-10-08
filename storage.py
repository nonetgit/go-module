import sqlite3
import json
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / "go_module.db"

# СХЕМА

SCHEMA = """
CREATE TABLE IF NOT EXISTS employees (
    employee_id TEXT PRIMARY KEY,
    fio         TEXT,
    department  TEXT,
    position    TEXT,
    face_size   INTEGER,
    status      TEXT
);

CREATE TABLE IF NOT EXISTS siz_cards (
    siz_id           TEXT PRIMARY KEY,
    name             TEXT,
    party_no         TEXT,
    quantity         INTEGER,
    storage          TEXT,
    manufacture_date TEXT,
    expiry_date      TEXT,
    status           TEXT
);

CREATE TABLE IF NOT EXISTS nfgo_members (
    employee_id TEXT PRIMARY KEY,
    fio         TEXT,
    department  TEXT,
    formation   TEXT,
    role        TEXT,
    status      TEXT
);

CREATE TABLE IF NOT EXISTS training (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id TEXT,
    fio         TEXT,
    topic       TEXT,
    hours       INTEGER,
    result      TEXT,
    date        TEXT,
    UNIQUE(employee_id, topic, date)
);

CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT
);
"""


def _connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = _connect()
    conn.executescript(SCHEMA)
    conn.commit()
    conn.close()

# СОХРАНЕНИЕ (upsert — не создаёт дубликатов)

def save_integrated_data(data: dict):
    conn = _connect()
    cur = conn.cursor()

    # Сотрудники
    cur.executemany("""
        INSERT INTO employees (employee_id, fio, department, position,
                               face_size, status)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(employee_id) DO UPDATE SET
            fio = excluded.fio,
            department = excluded.department,
            position = excluded.position,
            face_size = excluded.face_size,
            status = excluded.status
    """, [(e["employee_id"], e["fio"], e["department"], e["position"],
           e["face_size"], e["status"]) for e in data["employees"]])

    # СИЗ
    cur.executemany("""
        INSERT INTO siz_cards (siz_id, name, party_no, quantity, storage,
                               manufacture_date, expiry_date, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(siz_id) DO UPDATE SET
            name = excluded.name,
            party_no = excluded.party_no,
            quantity = excluded.quantity,
            storage = excluded.storage,
            manufacture_date = excluded.manufacture_date,
            expiry_date = excluded.expiry_date,
            status = excluded.status
    """, [(s["siz_id"], s["name"], s["party_no"], s["quantity"],
           s["storage"], s["manufacture_date"], s["expiry_date"], s["status"])
          for s in data["siz_cards"]])

    # НФГО
    cur.executemany("""
        INSERT INTO nfgo_members (employee_id, fio, department,
                                  formation, role, status)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(employee_id) DO UPDATE SET
            fio = excluded.fio,
            department = excluded.department,
            formation = excluded.formation,
            role = excluded.role,
            status = excluded.status
    """, [(m["employee_id"], m["fio"], m["department"],
           m["formation"], m["role"], m["status"]) for m in data["nfgo_members"]])

    # Обучение
    for t in data["training"]:
        cur.execute("""
            INSERT OR IGNORE INTO training
                (employee_id, fio, topic, hours, result, date)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (t["employee_id"], t.get("fio", ""), t["topic"],
              t["hours"], t["result"], t["date"]))

    # Мета: отметка, что интеграция была
    cur.execute("""
        INSERT INTO meta (key, value) VALUES ('integration_done', '1')
        ON CONFLICT(key) DO UPDATE SET value = '1'
    """)

    conn.commit()
    conn.close()

# ЧТЕНИЕ

def load_data() -> dict | None:
    conn = _connect()
    cur = conn.cursor()

    # Проверяем метку интеграции
    cur.execute("SELECT value FROM meta WHERE key = 'integration_done'")
    row = cur.fetchone()
    if not row or row["value"] != "1":
        conn.close()
        return None

    employees = [dict(r) for r in cur.execute(
        "SELECT employee_id, fio, department, position, face_size, status "
        "FROM employees ORDER BY employee_id").fetchall()]

    siz_cards = [dict(r) for r in cur.execute(
        "SELECT siz_id, name, party_no, quantity, storage, "
        "manufacture_date, expiry_date, status "
        "FROM siz_cards ORDER BY siz_id").fetchall()]

    nfgo_members = [dict(r) for r in cur.execute(
        "SELECT employee_id, fio, department, formation, role, status "
        "FROM nfgo_members ORDER BY formation, employee_id").fetchall()]

    training = [dict(r) for r in cur.execute(
        "SELECT employee_id, fio, topic, hours, result, date "
        "FROM training ORDER BY date DESC, employee_id").fetchall()]

    conn.close()

    if not employees and not siz_cards:
        return None

    return {
        "employees":    employees,
        "siz_cards":    siz_cards,
        "nfgo_members": nfgo_members,
        "training":     training,
    }

# ОЧИСТКА

def clear_all():
    conn = _connect()
    cur = conn.cursor()
    for table in ("employees", "siz_cards", "nfgo_members", "training", "meta"):
        cur.execute(f"DELETE FROM {table}")
    conn.commit()
    conn.close()


def has_data() -> bool:
    return load_data() is not None

# ИНИЦИАЛИЗАЦИЯ ПРИ СТАРТЕ

init_db()