# -*- coding: utf-8 -*-
"""SOP 表（sop.xls）与巡检系统机房/设备匹配，并导入数据库。"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from data_loader import (
    ENV_INSPECTION_ITEMS,
    ENV_INSPECTION_LABEL,
    _infer_room_type,
    apply_device_category_rules,
    get_all_device_categories,
    get_check_items,
    load_inspection_items,
    load_master_mappings,
    merge_manual_room_types,
    resolve_room_meta,
)
from database import init_db, insert_room, load_entries, room_exists, save_devices

DEFAULT_SOP_PATH = Path(r"d:\浏览器下载地址\sop.xls")
DEFAULT_COMMUNITY = "谷"

# 社区前缀 → 社区简称（可扩展其他园区）
COMMUNITY_PREFIX_MAP = {
    "大清谷": "谷",
    "燕园": "燕",
}

# SOP 机房名关键词 → 系统机房类型（优先于通用推断）
SOP_ROOM_KEYWORD_RULES: list[tuple[str, str]] = [
    ("高配室", "配电室"),
    ("强电井", "强电竖井"),
    ("热交互", "换热站"),
    ("板换机房", "换热站"),
    ("空调板换", "空调机房"),
    ("消防泵房", "给水机房"),
    ("冷却塔", "制冷站"),
    ("空气源热泵", "热水机房"),
    ("负压机房", "风机房"),
    ("排烟", "风机房"),
    ("排风", "风机房"),
    ("补风", "风机房"),
    ("进风", "风机房"),
    ("柴油发电", "柴发机房"),
    ("泳池", "泳池机房"),
    ("集水井", "集水井"),
    ("弱电井", "弱电井"),
    ("制冷", "制冷站"),
    ("空调", "空调机房"),
    ("电梯", "电梯机房"),
    ("给水", "给水机房"),
    ("污水泵", "给水机房"),
    ("锅炉", "热力站"),
]

# 检查项文本归一化替换（SOP 用语 → 系统用语）
ITEM_ALIASES = {
    "环境卫生": "卫生",
    "照明灯": "灯光",
    "照明及卫生": "卫生",
    "通风照明": "灯光",
    "供水压力mpa": "供水压力Mpa",
    "回水压力mpa": "回水压力Mpa",
    "供水温度℃": "供水温度℃",
    "回水温度℃": "回水温度℃",
    "进水温度(℃)": "蒸发器进水温",
    "出水温度(℃)": "蒸发器出水温",
    "进水温度℃": "蒸发器进水温",
    "出水温度℃": "蒸发器出水温",
    "蒸发器压力(kpa)": "蒸发器制冷剂压力",
    "冷凝器压力(kpa)": "冷凝器制冷剂压力",
    "供油压力(kpa)": "油压差(高油压)",
    "油箱温度(℃)": "油缸内油温",
    "饱和蒸发温度(℃)": "蒸发器趋近温度",
    "饱和冷凝温度(℃)": "冷凝器趋近温度",
    "压缩机排气温度(℃)": "压缩机马达线圈温度",
}

# 过于通用、不能单独用于判定设备类型的检查项
GENERIC_CHECK_ITEMS = {
    "运行状态",
    "出口压力",
    "进口压力",
    "压力mpa",
    "是否使用",
    "阀门状态",
    "运行状态是否在运行",
}


@dataclass
class SopRoomRecord:
    sop_name: str
    community: str
    room_name: str
    check_items: list[str] = field(default_factory=list)


@dataclass
class MatchResult:
    sop_name: str
    community: str
    room_name: str
    room_type: str
    devices: list[dict]
    env_items: list[str]
    skipped_reason: str | None = None


def _normalize_item(text: str) -> str:
    s = str(text).strip().lower()
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", "", s)
    return ITEM_ALIASES.get(s, s)


def _meaningful_items(items: set[str]) -> set[str]:
    return {item for item in items if item not in GENERIC_CHECK_ITEMS}


def parse_community_and_room(sop_room_name: str) -> tuple[str | None, str]:
    """从 SOP 标准分类解析社区简称与机房名称。"""
    name = str(sop_room_name).strip()
    if name.endswith("巡检"):
        name = name[:-2]

    for prefix, community in COMMUNITY_PREFIX_MAP.items():
        if name.startswith(prefix):
            room_name = name[len(prefix) :].strip()
            return community, room_name or name

    return None, name


def infer_sop_room_type(room_name: str) -> str:
    """根据 SOP 机房名推断系统机房类型。"""
    for keyword, room_type in SOP_ROOM_KEYWORD_RULES:
        if keyword in room_name:
            return room_type
    return _infer_room_type(room_name, "")


def load_known_room_types() -> set[str]:
    _, type_to_categories, _ = load_master_mappings(None)
    merged = merge_manual_room_types(type_to_categories)
    return set(merged.keys())


SOP_TEMPLATE_COLUMNS = ["标准分类", "标准名称", "检查内容", "操作类型"]

SOP_TEMPLATE_EXAMPLE_ROWS = [
    {
        "标准分类": "大清谷空调机房巡检",
        "标准名称": "空调运行与阀门",
        "检查内容": "供水压力MPa",
        "操作类型": "判断",
    },
    {
        "标准分类": "大清谷空调机房巡检",
        "标准名称": "空调运行与阀门",
        "检查内容": "回水压力MPa",
        "操作类型": "判断",
    },
    {
        "标准分类": "大清谷空调机房巡检",
        "标准名称": "空调运行与阀门",
        "检查内容": "阀门状态",
        "操作类型": "判断",
    },
    {
        "标准分类": "大清谷强电井巡检",
        "标准名称": "环境与标识",
        "检查内容": "环境卫生",
        "操作类型": "判断",
    },
    {
        "标准分类": "大清谷强电井巡检",
        "标准名称": "环境与标识",
        "检查内容": "照明灯",
        "操作类型": "判断",
    },
    {
        "标准分类": "燕园1号楼B1层1号空调机房巡检",
        "标准名称": "示例：燕园前缀",
        "检查内容": "供水温度℃",
        "操作类型": "判断",
    },
]


def get_sop_template_dataframe() -> pd.DataFrame:
    return pd.DataFrame(SOP_TEMPLATE_EXAMPLE_ROWS, columns=SOP_TEMPLATE_COLUMNS)


def build_sop_template_bytes() -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        get_sop_template_dataframe().to_excel(writer, index=False, sheet_name="SOP示例")
    return buffer.getvalue()


def parse_sop_dataframe(df: pd.DataFrame) -> list[SopRoomRecord]:
    """将 SOP DataFrame 解析为机房记录列表。"""
    if df.shape[1] < 3:
        raise ValueError(
            "SOP 表列数不足，至少需要 4 列：标准分类、标准名称、检查内容、操作类型"
        )

    room_col, content_col = df.columns[0], df.columns[2]
    grouped: dict[str, list[str]] = {}

    for _, row in df.iterrows():
        sop_name = row.get(room_col)
        if pd.isna(sop_name):
            continue
        sop_name = str(sop_name).strip()
        if not sop_name or sop_name.startswith("杭州"):
            continue

        content = row.get(content_col)
        if pd.isna(content):
            continue
        grouped.setdefault(sop_name, []).append(str(content).strip())

    records: list[SopRoomRecord] = []
    for sop_name, items in grouped.items():
        community, room_name = parse_community_and_room(sop_name)
        if not community:
            continue
        seen: set[str] = set()
        unique_items: list[str] = []
        for item in items:
            if item not in seen:
                seen.add(item)
                unique_items.append(item)
        records.append(
            SopRoomRecord(
                sop_name=sop_name,
                community=community,
                room_name=room_name,
                check_items=unique_items,
            )
        )
    return records


def read_sop_excel(
    source: str | Path | bytes | io.BytesIO,
    *,
    filename: str | None = None,
) -> list[SopRoomRecord]:
    """读取 SOP 表格（路径或上传字节），按标准分类聚合检查项。"""
    if isinstance(source, (str, Path)):
        df = pd.read_excel(source)
    else:
        buffer = io.BytesIO(source) if isinstance(source, bytes) else source
        name = (filename or "").lower()
        engine = "xlrd" if name.endswith(".xls") else None
        df = pd.read_excel(buffer, engine=engine)
    return parse_sop_dataframe(df)


def _split_env_items(sop_items: list[str]) -> tuple[list[str], list[str]]:
    """将 SOP 检查项拆分为环境项与疑似设备参数项。"""
    env_norm = {_normalize_item(x) for x in ENV_INSPECTION_ITEMS}
    env_like = {
        _normalize_item(x)
        for x in [
            "环境卫生",
            "照明灯",
            "通风照明",
            "照明及卫生",
            "机房卫生",
            "安全隐患",
            "标识",
            "温度",
            "湿度",
        ]
    }

    env_items: list[str] = []
    device_items: list[str] = []
    for item in sop_items:
        norm = _normalize_item(item)
        if norm in env_norm or norm in env_like or any(k in item for k in ("卫生", "照明", "标识", "安全", "温度", "湿度")):
            mapped = ITEM_ALIASES.get(norm, item)
            if mapped in ENV_INSPECTION_ITEMS and mapped not in env_items:
                env_items.append(mapped)
            elif item not in env_items and any(k in item for k in ("卫生", "照明", "标识", "安全")):
                # 保留原始描述供展示，但映射到最接近的环境项
                for env in ENV_INSPECTION_ITEMS:
                    if env not in env_items and (
                        (env == "卫生" and "卫生" in item)
                        or (env == "灯光" and "照明" in item)
                        or (env == "标识" in item and env == "标识")
                        or (env == "安全隐患" and "安全" in item)
                    ):
                        env_items.append(env)
        else:
            device_items.append(item)

    if not env_items:
        env_items = list(ENV_INSPECTION_ITEMS)
    return env_items, device_items


def _score_device_category(
    sop_norm: set[str],
    category: str,
    inspection_items: dict[str, list[str]],
) -> tuple[float, int, list[str]]:
    check_items = get_check_items(category, inspection_items)
    if not check_items:
        return 0.0, 0, []

    sop_meaningful = _meaningful_items(sop_norm)
    cat_norm = {_normalize_item(x) for x in check_items}
    cat_meaningful = _meaningful_items(cat_norm)
    if not sop_meaningful or not cat_meaningful:
        return 0.0, 0, check_items

    exact = sop_meaningful & cat_meaningful
    exact_count = len(exact)

    fuzzy_count = 0
    for sop_item in sop_meaningful - exact:
        if len(sop_item) < 5:
            continue
        if any(sop_item in cat_item or cat_item in sop_item for cat_item in cat_meaningful):
            fuzzy_count += 1

    score = (exact_count + fuzzy_count * 0.35) / max(len(cat_meaningful), 1)
    return score, exact_count, check_items


def _keyword_devices_for_room(
    room_type: str,
    sop_device_items: list[str],
    inspection_items: dict[str, list[str]],
) -> list[dict]:
    """按关键词补充高置信度设备（主要用于配电室参数项）。"""
    if room_type != "配电室" or not sop_device_items:
        return []

    blob = " ".join(sop_device_items).lower()
    keyword_rules: list[tuple[str, tuple[str, ...]]] = [
        ("变压器", ("变压器", "输出a项", "输出b项", "输出c项")),
        ("进线", ("高压输入", "高压输出", "进线", "馈线")),
        ("路", ("有功", "无功", "功率因数")),
        ("母线", ("母线",)),
        ("直流", ("直流", "蓄电池", "电池")),
        ("高压馈出线电流", ("211", "212", "213", "221", "222", "223", "馈线电流")),
    ]

    result: list[dict] = []
    seen: set[str] = set()
    for category, keywords in keyword_rules:
        if category in seen:
            continue
        if not any(k in blob for k in keywords):
            continue
        check_items = get_check_items(category, inspection_items)
        if not check_items:
            continue
        seen.add(category)
        result.append({"设备类型": category, "数量": 1, "check_items": check_items})
    return result


def match_devices_for_room(
    sop_device_items: list[str],
    room_type: str,
    inspection_items: dict[str, list[str]],
    type_to_categories: dict[str, list[str]],
) -> list[dict]:
    """将 SOP 设备类检查项匹配到系统设备类型（仅保留高置信度结果）。"""
    if not sop_device_items:
        return []

    sop_norm = {_normalize_item(x) for x in sop_device_items}
    if not _meaningful_items(sop_norm):
        return _keyword_devices_for_room(room_type, sop_device_items, inspection_items)

    type_candidates = list(type_to_categories.get(room_type, []))
    all_categories = get_all_device_categories(inspection_items)
    search_order = type_candidates if type_candidates else all_categories

    scored: list[tuple[float, int, str, list[str]]] = []
    for category in search_order:
        score, exact_count, check_items = _score_device_category(
            sop_norm, category, inspection_items
        )
        if exact_count >= 2 or (exact_count >= 1 and score >= 0.2):
            scored.append((score, exact_count, category, check_items))

    scored.sort(key=lambda x: (x[1], x[0]), reverse=True)

    result: list[dict] = []
    seen: set[str] = set()
    for _score, _exact, category, check_items in scored[:3]:
        if category in seen:
            continue
        seen.add(category)
        result.append(
            {
                "设备类型": category,
                "数量": 1,
                "check_items": check_items,
            }
        )

    for device in _keyword_devices_for_room(room_type, sop_device_items, inspection_items):
        if device["设备类型"] not in seen:
            seen.add(device["设备类型"])
            result.append(device)
    return result


def match_sop_records(
    records: list[SopRoomRecord],
    inspection_items: dict[str, list[str]] | None = None,
    known_room_types: set[str] | None = None,
    type_to_categories: dict[str, list[str]] | None = None,
) -> tuple[list[MatchResult], list[MatchResult]]:
    """匹配 SOP 记录，返回 (成功列表, 跳过列表)。"""
    items = inspection_items or load_inspection_items()
    valid_types = known_room_types or load_known_room_types()
    if type_to_categories is None:
        _, raw_types, _ = load_master_mappings(None)
        type_to_categories = apply_device_category_rules(
            merge_manual_room_types(raw_types), items
        )

    matched: list[MatchResult] = []
    skipped: list[MatchResult] = []

    for record in records:
        room_type = infer_sop_room_type(record.room_name)
        if not room_type or room_type not in valid_types:
            skipped.append(
                MatchResult(
                    sop_name=record.sop_name,
                    community=record.community,
                    room_name=record.room_name,
                    room_type=room_type or "",
                    devices=[],
                    env_items=[],
                    skipped_reason=f"机房类型无法匹配: {record.room_name}",
                )
            )
            continue

        env_items, device_items = _split_env_items(record.check_items)
        devices = match_devices_for_room(
            device_items, room_type, items, type_to_categories
        )
        matched.append(
            MatchResult(
                sop_name=record.sop_name,
                community=record.community,
                room_name=record.room_name,
                room_type=room_type,
                devices=devices,
                env_items=env_items,
            )
        )

    return matched, skipped


def import_matches_to_db(
    matches: list[MatchResult],
    room_catalog: dict | None = None,
    *,
    skip_existing: bool = True,
) -> dict[str, int]:
    """将匹配结果写入数据库。"""
    init_db()
    _, _, catalog = load_master_mappings(None)
    catalog = room_catalog or catalog

    stats = {"inserted": 0, "skipped_existing": 0, "devices_saved": 0}

    for match in matches:
        if skip_existing and room_exists(match.community, match.room_name):
            stats["skipped_existing"] += 1
            continue

        meta = resolve_room_meta(match.room_name, match.room_type, catalog)
        entry = {
            "社区分类": match.community,
            "机房名称": match.room_name,
            "机房类型": match.room_type,
            "meta": meta,
            "devices": match.devices,
            "env_check_items": match.env_items,
            "custom": match.room_name not in catalog,
        }
        room_id = insert_room(entry)
        if match.devices:
            save_devices(room_id, match.devices)
            stats["devices_saved"] += len(match.devices)
        stats["inserted"] += 1

    return stats


def sync_sop_from_upload(
    file_content: bytes,
    filename: str,
    community_filter: str | None = None,
    *,
    skip_existing: bool = True,
) -> dict:
    """上传 SOP 表格并同步到数据库，返回供界面展示的结果摘要。"""
    try:
        records = read_sop_excel(file_content, filename=filename)
    except Exception as exc:
        return {
            "ok": False,
            "error": f"表格读取失败：{exc}",
            "filename": filename,
        }

    all_count = len(records)
    if community_filter:
        records = [r for r in records if r.community == community_filter]

    if not records:
        prefixes = "、".join(COMMUNITY_PREFIX_MAP.keys())
        return {
            "ok": False,
            "error": (
                f"未识别到可导入机房。请确认「标准分类」含社区前缀（{prefixes}），"
                f"且与所选社区一致。"
            ),
            "filename": filename,
            "total_sop_rooms": all_count,
        }

    matched, skipped = match_sop_records(records)
    import_stats = import_matches_to_db(matched, skip_existing=skip_existing)

    return {
        "ok": True,
        "filename": filename,
        "total_sop_rooms": len(records),
        "matched": len(matched),
        "skipped": len(skipped),
        "import": import_stats,
        "matched_rooms": [
            {
                "社区": m.community,
                "机房名称": m.room_name,
                "机房类型": m.room_type,
                "设备": [d["设备类型"] for d in m.devices],
            }
            for m in matched
        ],
        "skipped_rooms": [
            {"sop": s.sop_name, "原因": s.skipped_reason} for s in skipped
        ],
    }


def run_import(
    sop_path: str | Path = DEFAULT_SOP_PATH,
    community_filter: str | None = DEFAULT_COMMUNITY,
    *,
    dry_run: bool = False,
) -> dict:
    """读取 SOP、匹配并导入数据库。"""
    records = read_sop_excel(sop_path)
    if community_filter:
        records = [r for r in records if r.community == community_filter]

    matched, skipped = match_sop_records(records)
    stats = {
        "sop_path": str(sop_path),
        "total_sop_rooms": len(records),
        "matched": len(matched),
        "skipped": len(skipped),
        "matched_rooms": [
            {
                "社区": m.community,
                "机房名称": m.room_name,
                "机房类型": m.room_type,
                "设备数": len(m.devices),
                "设备": [d["设备类型"] for d in m.devices],
            }
            for m in matched
        ],
        "skipped_rooms": [
            {"sop": s.sop_name, "原因": s.skipped_reason}
            for s in skipped
        ],
    }

    if not dry_run:
        import_stats = import_matches_to_db(matched)
        stats["import"] = import_stats
        stats["db_total"] = len(load_entries(community_filter))

    return stats


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="SOP 表匹配并导入巡检数据库")
    parser.add_argument(
        "--sop",
        default=str(DEFAULT_SOP_PATH),
        help="SOP Excel 路径（默认: d:\\浏览器下载地址\\sop.xls）",
    )
    parser.add_argument(
        "--community",
        default=DEFAULT_COMMUNITY,
        help="只处理指定社区简称（默认: 谷）",
    )
    parser.add_argument(
        "--all-communities",
        action="store_true",
        help="处理 SOP 中所有可识别社区",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅输出匹配结果，不写入数据库",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="以 JSON 格式输出结果",
    )
    args = parser.parse_args(argv)

    community = None if args.all_communities else args.community
    result = run_import(args.sop, community, dry_run=args.dry_run)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"SOP 文件: {result['sop_path']}")
        print(f"机房总数: {result['total_sop_rooms']}  匹配: {result['matched']}  跳过: {result['skipped']}")
        if result["matched_rooms"]:
            print("\n已匹配机房:")
            for row in result["matched_rooms"]:
                devices = "、".join(row["设备"]) if row["设备"] else f"仅{ENV_INSPECTION_LABEL}"
                print(f"  [{row['社区']}] {row['机房名称']} ({row['机房类型']}) → {devices}")
        if result["skipped_rooms"]:
            print("\n未匹配（已忽略）:")
            for row in result["skipped_rooms"]:
                print(f"  {row['sop']} — {row['原因']}")
        if not args.dry_run and "import" in result:
            imp = result["import"]
            print(
                f"\n入库: 新增 {imp['inserted']} 条，"
                f"已存在跳过 {imp['skipped_existing']} 条，"
                f"设备 {imp['devices_saved']} 个"
            )

    return 0


if __name__ == "__main__":
    sys.exit(main())
