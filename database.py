# -*- coding: utf-8 -*-
"""机房与设备持久化存储（本地 SQLite / Turso 在线 SQLite）。"""
from __future__ import annotations

import json
import os
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import libsql_client

from data_loader import ENV_INSPECTION_ITEMS

DEFAULT_DB_PATH = Path(__file__).resolve().parent / "yanyuan_inspection.db"
SEEDS_DIR = Path(__file__).resolve().parent / "seeds"

# 社区简称 → 种子数据文件（首次启动且该社区无数据时自动导入）
COMMUNITY_SEED_FILES: dict[str, str] = {
    "谷": "gu_rooms.json",
}


def _load_secrets_into_env() -> None:
    """从 Streamlit secrets 注入数据库连接（若环境变量未设置）。"""
    try:
        import streamlit as st

        db = st.secrets.get("database", {})
        if not os.getenv("YANYUAN_INSPECTION_DATABASE_URL") and db.get("url"):
            os.environ["YANYUAN_INSPECTION_DATABASE_URL"] = str(db["url"]).strip()
        if not os.getenv("YANYUAN_INSPECTION_AUTH_TOKEN") and db.get("auth_token"):
            os.environ["YANYUAN_INSPECTION_AUTH_TOKEN"] = str(db["auth_token"]).strip()
    except Exception:
        pass


def resolve_db_config() -> tuple[str, str]:
    """返回 (database_url, auth_token)。本地为 file: 路径，在线为 libsql://。"""
    _load_secrets_into_env()
    url = os.getenv("YANYUAN_INSPECTION_DATABASE_URL", "").strip()
    token = os.getenv("YANYUAN_INSPECTION_AUTH_TOKEN", "").strip()
    if url:
        return url, token
    db_path = Path(os.getenv("YANYUAN_INSPECTION_DB", DEFAULT_DB_PATH)).expanduser().resolve()
    return f"file:{db_path.as_posix()}", ""


def is_online_db() -> bool:
    url, _ = resolve_db_config()
    return url.startswith("libsql://")


def get_db_info() -> dict[str, str | bool]:
    url, _ = resolve_db_config()
    online = url.startswith("libsql://")
    if online:
        host = url.replace("libsql://", "").split("?")[0].rstrip("/")
        return {
            "online": True,
            "label": f"在线 SQLite（Turso · {host}）",
            "mode": "online",
        }
    path = url.removeprefix("file:")
    return {
        "online": False,
        "label": f"本地 SQLite（{path}）",
        "mode": "local",
    }


@contextmanager
def _connect():
    url, token = resolve_db_config()
    kwargs: dict = {}
    if token:
        kwargs["auth_token"] = token
    client = libsql_client.create_client_sync(url, **kwargs)
    try:
        client.execute("PRAGMA foreign_keys = ON")
        yield client
    finally:
        client.close()


def _rows_to_dicts(result) -> list[dict]:
    if not result.rows:
        return []
    return [dict(zip(result.columns, row)) for row in result.rows]


def _row_to_entry(room_row: dict, devices: list[dict]) -> dict:
    meta = json.loads(room_row.get("meta_json") or "{}")
    env_items = json.loads(room_row.get("env_check_items_json") or "[]")
    if not env_items:
        env_items = list(ENV_INSPECTION_ITEMS)

    device_list = []
    for d in devices:
        check_items = json.loads(d.get("check_items_json") or "[]")
        device_list.append(
            {
                "设备类型": d["device_category"],
                "数量": int(d["quantity"]),
                "check_items": check_items,
            }
        )

    return {
        "id": int(room_row["id"]),
        "社区分类": room_row["community_category"],
        "机房名称": room_row["room_name"],
        "机房类型": room_row["room_type"],
        "meta": meta,
        "devices": device_list,
        "env_check_items": env_items,
        "custom": bool(room_row["is_custom"]),
    }


def init_db() -> None:
    schema_statements = [
        """
        CREATE TABLE IF NOT EXISTS rooms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            community_category TEXT NOT NULL,
            room_name TEXT NOT NULL,
            room_type TEXT NOT NULL,
            meta_json TEXT NOT NULL DEFAULT '{}',
            env_check_items_json TEXT NOT NULL DEFAULT '[]',
            is_custom INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(community_category, room_name)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS room_devices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            room_id INTEGER NOT NULL,
            device_category TEXT NOT NULL,
            quantity INTEGER NOT NULL DEFAULT 1,
            check_items_json TEXT NOT NULL DEFAULT '[]',
            FOREIGN KEY (room_id) REFERENCES rooms(id) ON DELETE CASCADE,
            UNIQUE(room_id, device_category)
        )
        """,
        "CREATE INDEX IF NOT EXISTS idx_rooms_community ON rooms(community_category)",
    ]
    with _connect() as conn:
        for sql in schema_statements:
            conn.execute(sql)


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def list_community_categories() -> list[str]:
    init_db()
    with _connect() as conn:
        result = conn.execute(
            "SELECT DISTINCT community_category FROM rooms ORDER BY community_category"
        )
    return [row["community_category"] for row in _rows_to_dicts(result)]


def load_entries(community_category: str | None = None) -> list[dict]:
    init_db()
    with _connect() as conn:
        if community_category:
            rooms_result = conn.execute(
                "SELECT * FROM rooms WHERE community_category = ? ORDER BY id",
                [community_category],
            )
        else:
            rooms_result = conn.execute("SELECT * FROM rooms ORDER BY community_category, id")

        entries: list[dict] = []
        for room in _rows_to_dicts(rooms_result):
            devices_result = conn.execute(
                "SELECT * FROM room_devices WHERE room_id = ? ORDER BY id",
                [room["id"]],
            )
            devices = _rows_to_dicts(devices_result)
            entries.append(_row_to_entry(room, devices))
        return entries


def insert_room(entry: dict) -> int:
    init_db()
    now = _now()
    meta_json = json.dumps(entry.get("meta") or {}, ensure_ascii=False)
    env_json = json.dumps(entry.get("env_check_items") or list(ENV_INSPECTION_ITEMS), ensure_ascii=False)
    with _connect() as conn:
        result = conn.execute(
            """
            INSERT INTO rooms (
                community_category, room_name, room_type, meta_json,
                env_check_items_json, is_custom, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                entry["社区分类"],
                entry["机房名称"],
                entry["机房类型"],
                meta_json,
                env_json,
                1 if entry.get("custom") else 0,
                now,
                now,
            ],
        )
        return int(result.last_insert_rowid)


def delete_room(room_id: int) -> None:
    init_db()
    with _connect() as conn:
        conn.execute("DELETE FROM rooms WHERE id = ?", [room_id])


def clear_rooms(community_category: str | None = None) -> None:
    init_db()
    with _connect() as conn:
        if community_category:
            conn.execute("DELETE FROM rooms WHERE community_category = ?", [community_category])
        else:
            conn.execute("DELETE FROM rooms")


def save_devices(room_id: int, devices: list[dict]) -> None:
    init_db()
    now = _now()
    with _connect() as conn:
        conn.execute("DELETE FROM room_devices WHERE room_id = ?", [room_id])
        for device in devices:
            check_json = json.dumps(device.get("check_items") or [], ensure_ascii=False)
            conn.execute(
                """
                INSERT INTO room_devices (room_id, device_category, quantity, check_items_json)
                VALUES (?, ?, ?, ?)
                """,
                [
                    room_id,
                    device["设备类型"],
                    int(device["数量"]),
                    check_json,
                ],
            )
        conn.execute("UPDATE rooms SET updated_at = ? WHERE id = ?", [now, room_id])


def room_exists(community_category: str, room_name: str) -> bool:
    init_db()
    with _connect() as conn:
        result = conn.execute(
            "SELECT 1 FROM rooms WHERE community_category = ? AND room_name = ?",
            [community_category, room_name],
        )
    return bool(result.rows)


def count_rooms(community_category: str) -> int:
    init_db()
    with _connect() as conn:
        result = conn.execute(
            "SELECT COUNT(*) AS cnt FROM rooms WHERE community_category = ?",
            [community_category],
        )
    rows = _rows_to_dicts(result)
    return int(rows[0]["cnt"]) if rows else 0


def seed_community_if_empty(community_category: str) -> int:
    """若社区尚无数据，则从 seeds/ 导入预置机房。返回新增条数。"""
    seed_file = COMMUNITY_SEED_FILES.get(community_category)
    if not seed_file or count_rooms(community_category) > 0:
        return 0

    path = SEEDS_DIR / seed_file
    if not path.exists():
        return 0

    with open(path, encoding="utf-8") as f:
        entries = json.load(f)

    inserted = 0
    for entry in entries:
        if room_exists(community_category, entry["机房名称"]):
            continue
        entry = {**entry, "社区分类": community_category}
        room_id = insert_room(entry)
        devices = entry.get("devices") or []
        if devices:
            save_devices(room_id, devices)
        inserted += 1
    return inserted


def seed_all_communities_if_empty() -> dict[str, int]:
    """为所有配置了种子文件的社区执行空库导入。"""
    return {
        community: seed_community_if_empty(community)
        for community in COMMUNITY_SEED_FILES
    }
