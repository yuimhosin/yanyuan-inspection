# -*- coding: utf-8 -*-
import json
import re
from pathlib import Path

import pandas as pd

META_FIELDS = [
    "巡检专业",
    "机房类型",
    "机房简称",
    "所属分期",
    "所属楼栋（空间）",
    "巡检频次/日",
    "环境巡检",
]
# 仅用于读取燕园巡检.xlsx，不在界面与导出中展示

EXCEL_PATH = Path(__file__).resolve().parent.parent / "燕园巡检.xlsx"
INSPECTION_ITEMS_PATH = Path(__file__).resolve().parent / "inspection_items.json"

ENV_INSPECTION_ITEMS = ["灯光", "温度", "湿度", "标识", "安全隐患", "卫生"]
ENV_INSPECTION_LABEL = "环境巡检"

# 检查项中与空间/环境巡检重复、不应出现在设备行里的条目
ENV_LIKE_DEVICE_ITEMS = {
    "机房卫生",
    "通风照明",
    "照明通风",
    "照明及卫生",
    "设备运行声音",
}

# 空间性设备：检查项被更具体设备覆盖，或仅为机房环境类描述
REMOVED_SPATIAL_DEVICE_CATEGORIES = {
    "楼高区",
    "楼中区",
    "楼低区",
    "制冷站房",
    "生活热水机房",
    "燃气热水机房",
}


def resolve_device_category(category: str) -> str:
    return DEVICE_CATEGORY_ALIASES.get(category, category)

# 按机房类型排除不合理设备
ROOM_TYPE_DEVICE_EXCLUDES: dict[str, set[str]] = {
    "配电室": {"冷水机组"},
}

# 原始设备类别 → 展示/检查项键名
DEVICE_CATEGORY_ALIASES: dict[str, str] = {
    "电梯机房": "电梯机房设备",
}

# 作业指导书附表补充的机房类型与设备（不依赖燕园巡检.xlsx）
MANUAL_ROOM_DEVICE_MAP: dict[str, list[str]] = {
    "热力站": ["一次水", "热交换器", "循环泵"],
    "换热站": ["一次水", "热交换器", "循环泵"],
    "柴发机房": ["柴油发电机", "发电机组"],
    "泳池机房": [
        "过滤沙缸",
        "循环过滤泵",
        "按摩泵",
        "热泵水循环泵",
        "加药泵",
        "通风系统",
        "排水系统",
        "加臭氧系统",
        "水箱补水系统",
    ],
    "强电竖井": ["配电柜"],
    "集水井": ["集水坑", "潜污泵", "配电箱", "单向阀"],
}

_ZONE_SUPPLY_ITEMS = {
    "供水压力Mpa",
    "回水压力Mpa",
    "供水温度℃",
    "回水温度℃",
}


def extract_device_category(name: str) -> str:
    """从设备全名提取去重后的类别（去掉编号与括号内编码）。"""
    s = str(name).strip()
    s = re.sub(r"^\d+#", "", s)
    s = re.sub(r"（[^）]*）", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r"^\d+号", "", s)
    if re.match(r"^\d+路$", s):
        return "路"
    s = re.sub(r"\d+$", "", s)
    return s.strip() or str(name).strip()


def _normalize_room_type(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def _infer_room_type(room_name: str, raw_type: str) -> str:
    if raw_type:
        return raw_type
    if "热水" in room_name or "换热" in room_name:
        return "热水机房"
    if "配电" in room_name:
        return "配电室"
    if "空调" in room_name:
        return "空调机房"
    if "电梯" in room_name:
        return "电梯机房"
    if "制冷" in room_name:
        return "制冷站"
    if "热力" in room_name:
        return "热力站"
    if "换热" in room_name:
        return "换热站"
    if "柴发" in room_name or "柴油发电" in room_name:
        return "柴发机房"
    if "泳池" in room_name:
        return "泳池机房"
    if "竖井" in room_name:
        return "强电竖井"
    if "集水井" in room_name or "集水坑" in room_name:
        return "集水井"
    if "给水" in room_name or "水泵" in room_name:
        return "给水机房"
    return ""


def load_inspection_data(path: str | Path = EXCEL_PATH) -> list[dict]:
    df = pd.read_excel(path)
    param_cols = [c for c in df.columns if c.startswith("Unnamed")]

    rooms: list[dict] = []
    current: dict | None = None

    for _, row in df.iterrows():
        if pd.notna(row.get("机房名称")):
            if current:
                rooms.append(current)
            room_name = str(row["机房名称"]).strip()
            raw_type = _normalize_room_type(row.get("机房类型"))
            room_type = _infer_room_type(room_name, raw_type)
            meta = {k: row[k] for k in META_FIELDS}
            meta["机房类型"] = room_type or raw_type
            current = {
                "机房名称": room_name,
                "meta": meta,
                "devices": [],
            }
        if current and pd.notna(row.get("设备参数巡检")):
            params = [str(row[c]).strip() for c in param_cols if pd.notna(row[c])]
            device_name = str(row["设备参数巡检"]).strip()
            current["devices"].append(
                {
                    "name": device_name,
                    "category": extract_device_category(device_name),
                    "params": params,
                }
            )

    if current:
        rooms.append(current)

    return rooms


def build_mappings(rooms: list[dict]) -> tuple[dict[str, str], dict[str, list[str]], dict[str, dict]]:
    room_name_to_type: dict[str, str] = {}
    type_to_categories: dict[str, set[str]] = {}
    room_catalog: dict[str, dict] = {}

    for room in rooms:
        name = room["机房名称"]
        room_type = _normalize_room_type(room["meta"].get("机房类型"))
        room_type = _infer_room_type(name, room_type)
        if not room_type:
            continue

        room_name_to_type[name] = room_type
        room_catalog[name] = {
            "机房名称": name,
            "机房类型": room_type,
            "meta": {**room["meta"], "机房类型": room_type},
        }

        bucket = type_to_categories.setdefault(room_type, set())
        for device in room["devices"]:
            bucket.add(device["category"])

    sorted_mapping = {k: sorted(v) for k, v in sorted(type_to_categories.items())}
    return room_name_to_type, sorted_mapping, room_catalog


def merge_manual_room_types(type_to_categories: dict[str, list[str]]) -> dict[str, list[str]]:
    merged = {room_type: list(categories) for room_type, categories in type_to_categories.items()}
    for room_type, devices in MANUAL_ROOM_DEVICE_MAP.items():
        merged[room_type] = sorted(devices)
    return merged


def _is_spatial_duplicate_category(
    category: str,
    items_map: dict[str, list[str]],
) -> bool:
    if category in REMOVED_SPATIAL_DEVICE_CATEGORIES:
        return True

    items = items_map.get(category, [])
    if not items:
        return False

    # 楼X区：仅供回水压力温度，与「楼X区生活热水」重复
    if re.match(r"^楼[高中低]区$", category) and set(items) <= _ZONE_SUPPLY_ITEMS:
        return True

    # 仅剩环境/机房类检查项的空间描述
    env_only = ENV_LIKE_DEVICE_ITEMS | set(ENV_INSPECTION_ITEMS)
    if all(item in env_only for item in items):
        return True

    return False


def apply_device_category_rules(
    type_to_categories: dict[str, list[str]],
    items_map: dict[str, list[str]],
) -> dict[str, list[str]]:
    result: dict[str, list[str]] = {}
    for room_type, categories in type_to_categories.items():
        excludes = ROOM_TYPE_DEVICE_EXCLUDES.get(room_type, set())
        kept = []
        for cat in categories:
            if cat in excludes:
                continue
            if _is_spatial_duplicate_category(cat, items_map):
                continue
            display_cat = resolve_device_category(cat)
            if display_cat not in kept:
                kept.append(display_cat)
        result[room_type] = sorted(kept)
    return result


def get_valid_device_categories(room_type: str, type_to_categories: dict[str, list[str]]) -> list[str]:
    return list(type_to_categories.get(room_type, []))


def is_valid_device_for_room(room_type: str, device_category: str, type_to_categories: dict[str, list[str]]) -> bool:
    return device_category in type_to_categories.get(room_type, [])


def format_meta_value(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip() or "—"


def build_empty_meta(room_type: str) -> dict:
    return {field: None for field in META_FIELDS} | {"机房类型": room_type}


def resolve_room_meta(
    room_name: str,
    room_type: str,
    room_catalog: dict[str, dict],
) -> dict:
    """优先使用已有机房的基础信息，机房类型以用户选择为准。"""
    if room_name in room_catalog:
        meta = dict(room_catalog[room_name]["meta"])
        meta["机房类型"] = room_type
        return meta
    return build_empty_meta(room_type)


def ensure_env_check_items(entry: dict) -> list[str]:
    if not entry.get("env_check_items"):
        entry["env_check_items"] = list(ENV_INSPECTION_ITEMS)
    return entry["env_check_items"]


def load_inspection_items(path: str | Path = INSPECTION_ITEMS_PATH) -> dict[str, list[str]]:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def get_check_items(device_type: str, items_map: dict[str, list[str]]) -> list[str]:
    key = resolve_device_category(device_type)
    items = list(items_map.get(key, items_map.get(device_type, [])))
    skip = set(ENV_INSPECTION_ITEMS) | ENV_LIKE_DEVICE_ITEMS
    return [item for item in items if item not in skip]


def ensure_device_check_items(
    device: dict,
    items_map: dict[str, list[str]],
) -> list[str]:
    if not device.get("check_items"):
        device["check_items"] = get_check_items(device.get("设备类型", ""), items_map)
    else:
        device["check_items"] = [
            item
            for item in device["check_items"]
            if item not in ENV_INSPECTION_ITEMS and item not in ENV_LIKE_DEVICE_ITEMS
        ]
    return device["check_items"]


def build_summary_rows(
    entries: list[dict],
    items_map: dict[str, list[str]] | None = None,
) -> list[dict]:
    rows: list[dict] = []
    max_check_cols = 0

    for entry in entries:
        env_items = ensure_env_check_items(entry)
        max_check_cols = max(max_check_cols, len(env_items))
        rows.append(
            {
                "机房名称": entry["机房名称"],
                "机房类型": entry["机房类型"],
                "设备类型": ENV_INSPECTION_LABEL,
                "数量": 1,
                "check_items": env_items,
            }
        )

        devices = entry.get("devices") or []
        for device in devices:
            check_items = (
                ensure_device_check_items(device, items_map)
                if items_map
                else device.get("check_items", [])
            )
            max_check_cols = max(max_check_cols, len(check_items))
            rows.append(
                {
                    "机房名称": entry["机房名称"],
                    "机房类型": entry["机房类型"],
                    "设备类型": device["设备类型"],
                    "数量": device["数量"],
                    "check_items": check_items,
                }
            )

    output: list[dict] = []
    for row in rows:
        flat = {
            "机房名称": row["机房名称"],
            "机房类型": row["机房类型"],
            "设备类型": row["设备类型"],
            "数量": row["数量"],
        }
        for i in range(1, max_check_cols + 1):
            flat[f"需巡检项目{i}"] = (
                row["check_items"][i - 1] if i <= len(row["check_items"]) else ""
            )
        output.append(flat)
    return output
