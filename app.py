# -*- coding: utf-8 -*-

"""巡检上报表 — Streamlit"""

from pathlib import Path



import pandas as pd

import streamlit as st



from data_loader import (

    ENV_INSPECTION_ITEMS,

    ENV_INSPECTION_LABEL,

    apply_device_category_rules,

    build_summary_rows,

    ensure_device_check_items,

    ensure_env_check_items,

    get_check_items,

    get_community_categories,

    get_valid_device_categories,

    is_valid_device_for_room,

    load_inspection_items,

    load_master_mappings,

    merge_manual_room_types,

    resolve_excel_path,

    resolve_room_meta,

)

from sop_matcher import (
    COMMUNITY_PREFIX_MAP,
    build_sop_template_bytes,
    get_sop_template_dataframe,
    sync_sop_from_upload,
)

from database import (
    clear_rooms,
    delete_room,
    get_db_info,
    init_db,
    insert_room,
    load_entries,
    room_exists,
    save_devices,
    seed_all_communities_if_empty,
)



st.set_page_config(page_title="巡检上报", page_icon="📋", layout="wide")



st.markdown(

    """

    <style>

    .report-card {

        background: #f8fafc;

        border: 1px solid #dbe3ef;

        border-radius: 10px;

        padding: 14px 16px;

        margin-bottom: 12px;

    }

    .report-title { font-weight: 700; color: #0f172a; margin-bottom: 4px; }

    .report-sub { color: #64748b; font-size: 0.9rem; }

    .check-tag {

        display: inline-block;

        background: #ecfdf5;

        color: #047857;

        border-radius: 6px;

        padding: 3px 10px;

        margin: 2px 6px 2px 0;

        font-size: 0.84rem;

        border: 1px solid #a7f3d0;

    }

    .check-title { color: #374151; font-size: 0.9rem; margin: 8px 0 4px; }

    .sop-hint-box {

        background: #fffbeb;

        border: 1px solid #fcd34d;

        border-radius: 10px;

        padding: 12px 14px;

        margin: 8px 0 12px;

    }

    .sop-hint-title { font-weight: 700; color: #92400e; margin-bottom: 6px; }

    .sop-hint-text { color: #78350f; font-size: 0.9rem; line-height: 1.55; }

    .sop-table-wrap {

        border: 1px solid #dbe3ef;

        border-radius: 8px;

        overflow: hidden;

        margin-top: 8px;

    }

    .sop-table-wrap table {

        width: 100%;

        border-collapse: collapse;

        font-size: 0.86rem;

    }

    .sop-table-wrap th {

        background: #eff6ff;

        color: #1e3a8a;

        text-align: left;

        padding: 8px 10px;

        border-bottom: 1px solid #dbe3ef;

    }

    .sop-table-wrap td {

        padding: 7px 10px;

        border-bottom: 1px solid #eef2f7;

        color: #334155;

    }

    .sop-table-wrap tr:last-child td { border-bottom: none; }

    </style>

    """,

    unsafe_allow_html=True,

)





def _init_state() -> None:

    init_db()
    seed_all_communities_if_empty()

    if "report_entries" not in st.session_state:

        st.session_state.report_entries = []

    if "active_room_idx" not in st.session_state:

        st.session_state.active_room_idx = 0

    if "reset_room_form" not in st.session_state:

        st.session_state.reset_room_form = False

    if "focus_latest_room" not in st.session_state:

        st.session_state.focus_latest_room = False

    if "_prev_community" not in st.session_state:

        st.session_state._prev_community = None





def _reload_entries_for_community(community: str) -> None:

    st.session_state.report_entries = load_entries(community)

    st.session_state._prev_community = community

    st.session_state.active_room_idx = 0





def _sync_active_room_select(entry_count: int) -> None:

    """在 selectbox 渲染前同步选中项，确保新添加机房后自动切到最新一条。"""

    if entry_count <= 0:

        return



    latest = entry_count - 1

    if st.session_state.focus_latest_room:

        st.session_state.active_room_select = latest

        st.session_state.active_room_idx = latest

        st.session_state.focus_latest_room = False

        return



    current = st.session_state.get("active_room_select", 0)

    if current >= entry_count:

        st.session_state.active_room_select = latest

        st.session_state.active_room_idx = latest





def _apply_form_pending_actions() -> None:

    """在控件渲染前处理表单状态，避免 widget 实例化后修改 session_state 报错。"""

    if st.session_state.reset_room_form:

        st.session_state.room_name_input = ""

        st.session_state.reset_room_form = False





@st.cache_data(show_spinner="正在加载基础数据…")

def _load_master_data(excel_path: str | None):

    _, type_to_categories, room_catalog = load_master_mappings(excel_path)

    inspection_items = load_inspection_items()

    type_to_categories = merge_manual_room_types(type_to_categories)

    type_to_categories = apply_device_category_rules(type_to_categories, inspection_items)

    return type_to_categories, room_catalog, inspection_items





def _find_entry_index(room_name: str) -> int | None:

    for i, entry in enumerate(st.session_state.report_entries):

        if entry["机房名称"] == room_name:

            return i

    return None





def _persist_devices(entry: dict) -> None:

    if entry.get("id"):

        save_devices(entry["id"], entry.get("devices") or [])





def _render_check_items(items: list[str], title: str = "需巡检项目") -> None:

    if not items:

        st.caption("暂无对应作业指导书检查项目。")

        return

    tags = "".join(f'<span class="check-tag">{item}</span>' for item in items)

    st.markdown(

        f'<div class="check-title">{title}（{len(items)} 项）</div>{tags}',

        unsafe_allow_html=True,

    )





def _render_summary_table(inspection_items: dict[str, list[str]]) -> None:

    rows = build_summary_rows(st.session_state.report_entries, inspection_items)

    if rows:

        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    else:

        st.info("暂无上报记录，请先添加机房与设备。")





def _export_dataframe(export_rows: list[dict]) -> pd.DataFrame:

    if not export_rows:

        return pd.DataFrame()



    preferred = ["社区分类", "机房名称", "机房类型", "设备类型", "数量"]

    check_cols = sorted(

        [c for c in export_rows[0] if c.startswith("需巡检项目")],

        key=lambda x: int(x.replace("需巡检项目", "") or 0),

    )

    columns = preferred + check_cols

    return pd.DataFrame(export_rows).reindex(columns=columns)





def _to_excel_bytes(df: pd.DataFrame) -> bytes:

    from io import BytesIO



    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        df.to_excel(writer, index=False, sheet_name="巡检上报")

    return buffer.getvalue()





def _render_sop_template_hint() -> None:

    prefixes = "、".join(f"「{p}」" for p in COMMUNITY_PREFIX_MAP)

    example_html = get_sop_template_dataframe().to_html(index=False, border=0, escape=False)

    st.markdown(

        f"""

        <div class="sop-hint-box">

            <div class="sop-hint-title">SOP 表格格式说明</div>

            <div class="sop-hint-text">

                上传 .xls / .xlsx，系统会自动匹配机房与设备并写入当前社区数据库。<br>

                · 第 1 列 <b>标准分类</b>：须含社区前缀（{prefixes}）+ 机房名 + 「巡检」<br>

                · 第 2 列 <b>标准名称</b>：检查项分组说明（可留空）<br>

                · 第 3 列 <b>检查内容</b>：具体巡检参数或环境项<br>

                · 第 4 列 <b>操作类型</b>：一般为「判断」<br>

                · 同一机房占多行；无法匹配的机房/设备会自动跳过

            </div>

            <div class="sop-table-wrap">{example_html}</div>

        </div>

        """,

        unsafe_allow_html=True,

    )





def _render_sop_upload_sync(community: str) -> None:

    with st.expander("📤 SOP 表格上传与自动同步", expanded=False):

        _render_sop_template_hint()

        dl_col, up_col = st.columns([1, 2])

        with dl_col:

            st.download_button(

                "下载示例模板 (.xlsx)",

                data=build_sop_template_bytes(),

                file_name="SOP表格示例.xlsx",

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                use_container_width=True,

            )

            st.caption("可按此格式填写后上传。")

        with up_col:

            only_current = st.checkbox(

                f"仅同步当前社区「{community}」",

                value=True,

                help="关闭后将导入表格中所有已识别社区的数据",

            )

            uploaded = st.file_uploader(

                "选择 SOP 表格",

                type=["xls", "xlsx"],

                key="sop_file_uploader",

                help="上传后立即自动匹配并写入数据库",

            )

        if not uploaded:

            return

        file_key = f"{uploaded.name}:{uploaded.size}"

        if st.session_state.get("_last_sop_sync_key") == file_key:

            if summary := st.session_state.get("_last_sop_sync_summary"):

                st.info(summary)

            return

        with st.spinner("正在匹配并同步…"):

            result = sync_sop_from_upload(

                uploaded.getvalue(),

                uploaded.name,

                community_filter=community if only_current else None,

            )

        st.session_state._last_sop_sync_key = file_key

        if not result.get("ok"):

            st.error(result.get("error", "同步失败"))

            st.session_state._last_sop_sync_summary = None

            return

        imp = result["import"]

        summary = (

            f"「{uploaded.name}」同步完成：匹配 {result['matched']} 个机房，"

            f"新增 {imp['inserted']} 条，已存在跳过 {imp['skipped_existing']} 条，"

            f"设备 {imp['devices_saved']} 个；未匹配 {result['skipped']} 条已忽略。"

        )

        st.session_state._last_sop_sync_summary = summary

        st.success(summary)

        if result["matched_rooms"]:

            preview = pd.DataFrame(

                [

                    {

                        "社区": r["社区"],

                        "机房名称": r["机房名称"],

                        "机房类型": r["机房类型"],

                        "设备": "、".join(r["设备"]) if r["设备"] else f"仅{ENV_INSPECTION_LABEL}",

                    }

                    for r in result["matched_rooms"][:20]

                ]

            )

            st.dataframe(preview, use_container_width=True, hide_index=True)

            if len(result["matched_rooms"]) > 20:

                st.caption(f"仅展示前 20 条，共 {len(result['matched_rooms'])} 条。")

        if result["skipped_rooms"]:

            with st.expander(f"未匹配项（{len(result['skipped_rooms'])}）", expanded=False):

                for row in result["skipped_rooms"]:

                    st.caption(f"{row['sop']} — {row['原因']}")

        _reload_entries_for_community(community)

        st.rerun()





def main() -> None:

    _init_state()



    excel_path = resolve_excel_path()
    excel_path_str = str(excel_path) if excel_path else None

    type_to_categories, room_catalog, inspection_items = _load_master_data(excel_path_str)



    db_info = get_db_info()

    st.title("巡检上报表")

    st.caption(
        f"按社区管理机房；任意机房类型均可添加全部设备类型。"
        f" 数据存储：{db_info['label']}。"
    )



    community_options = get_community_categories()

    community = st.selectbox(

        "社区分类",

        options=community_options,

        key="community_select",

        help="选择所属养老社区（园区简称），决定数据库存储分区",

    )



    if st.session_state._prev_community != community:

        _reload_entries_for_community(community)



    _render_sop_upload_sync(community)



    room_types = sorted(type_to_categories.keys())



    # ── 第一步：录入机房 ──

    st.subheader("1. 录入机房")



    _apply_form_pending_actions()



    col_name, col_type, col_btn = st.columns([3, 2, 1])

    with col_name:

        room_name = st.text_input(

            "机房名称",

            placeholder="可自定义输入，如：5号楼B1层新配电室",

            key="room_name_input",

        )

    with col_type:

        room_type = st.selectbox(

            "机房类型",

            options=room_types,

            key="room_type_input",

            help="自行选择；决定环境巡检与汇总分类",

        )

    with col_btn:

        st.write("")

        st.write("")

        add_room = st.button("添加机房", type="primary", use_container_width=True)



    if room_name in room_catalog:

        suggested = room_catalog[room_name]["机房类型"]

        if suggested != room_type:

            st.caption(f"提示：基础数据中「{room_name}」的类型为「{suggested}」，当前已选手动类型「{room_type}」。")

        else:

            st.caption("已匹配基础数据，将自动带入该机房的其他基础信息。")



    if add_room:

        room_name = room_name.strip()

        if not room_name:

            st.warning("请填写机房名称。")

        elif not room_type:

            st.warning("请选择机房类型。")

        elif _find_entry_index(room_name) is not None or room_exists(community, room_name):

            st.warning(f"「{community}」下「{room_name}」已存在，请勿重复添加。")

        else:

            meta = resolve_room_meta(room_name, room_type, room_catalog)

            entry = {

                "社区分类": community,

                "机房名称": room_name,

                "机房类型": room_type,

                "meta": meta,

                "devices": [],

                "env_check_items": list(ENV_INSPECTION_ITEMS),

                "custom": room_name not in room_catalog,

            }

            entry["id"] = insert_room(entry)

            st.session_state.report_entries.append(entry)

            st.session_state.focus_latest_room = True

            st.session_state.reset_room_form = True

            st.success(f"已添加并保存：{community} · {room_name}（{room_type}）")

            st.rerun()



    if not st.session_state.report_entries:

        st.divider()

        st.subheader("上报汇总")

        _render_summary_table(inspection_items)

        return



    st.divider()



    # ── 第二步：为已添加机房录入设备 ──

    st.subheader("2. 添加设备类型与数量")



    room_labels = [

        f"{e['机房名称']}（{e['机房类型']}）" for e in st.session_state.report_entries

    ]

    _sync_active_room_select(len(room_labels))

    active_idx = st.selectbox(

        "当前操作的机房",

        range(len(room_labels)),

        format_func=lambda i: room_labels[i],

        key="active_room_select",

    )

    st.session_state.active_room_idx = active_idx

    entry = st.session_state.report_entries[active_idx]

    room_type = entry["机房类型"]



    st.markdown(

        f'<div class="report-card"><div class="report-title">{entry["机房名称"]}</div>'

        f'<div class="report-sub">社区分类：{entry.get("社区分类", community)} · 机房类型：{room_type}</div></div>',

        unsafe_allow_html=True,

    )

    if entry.get("custom"):

        st.caption("自定义机房。")



    st.markdown(f"**{ENV_INSPECTION_LABEL}（必检）**")

    _render_check_items(ensure_env_check_items(entry), title=ENV_INSPECTION_LABEL)



    st.markdown("**设备巡检**")

    valid_categories = get_valid_device_categories(room_type, type_to_categories, inspection_items)

    dcol1, dcol2, dcol3 = st.columns([3, 1, 1])

    with dcol1:

        device_category = st.selectbox(

            "设备类型",

            options=[""] + valid_categories,

            format_func=lambda x: "请选择设备类型" if x == "" else x,

            key=f"device_cat_{active_idx}",

            help="可选全部已配置检查项目的设备类型",

        )

        if device_category:

            preview_items = get_check_items(device_category, inspection_items)

            _render_check_items(preview_items)

    with dcol2:

        quantity = st.number_input(

            "数量",

            min_value=1,

            max_value=999,

            value=1,

            step=1,

            key=f"device_qty_{active_idx}",

        )

    with dcol3:

        st.write("")

        st.write("")

        add_device = st.button("添加设备", use_container_width=True)



    if add_device:

        if not device_category:

            st.warning("请选择设备类型。")

        elif not is_valid_device_for_room(
            room_type, device_category, type_to_categories, inspection_items
        ):

            st.error(f"「{device_category}」不是有效设备类型。")

        else:

            existing = next(

                (d for d in entry["devices"] if d["设备类型"] == device_category),

                None,

            )

            if existing:

                existing["数量"] += int(quantity)

                st.success(f"已累加 {device_category} 数量 +{quantity}")

            else:

                entry["devices"].append(

                    {

                        "设备类型": device_category,

                        "数量": int(quantity),

                        "check_items": get_check_items(device_category, inspection_items),

                    }

                )

                st.success(f"已添加 {device_category} × {quantity}")

            _persist_devices(entry)

            st.rerun()



    # 当前机房已添加的设备

    if entry["devices"]:

        st.markdown("**本机房已添加设备**")

        for device in entry["devices"]:

            items = ensure_device_check_items(device, inspection_items)

            st.markdown(f"**{device['设备类型']}** × {device['数量']}")

            _render_check_items(items)



        if st.button("清空本机房设备", key=f"clear_devices_{active_idx}"):

            entry["devices"] = []

            _persist_devices(entry)

            st.rerun()



    st.divider()



    # ── 已添加机房列表 & 汇总 ──

    st.subheader("3. 本次上报汇总")



    action_cols = st.columns([1, 1, 4])

    with action_cols[0]:

        if st.button("删除当前机房", type="secondary"):

            if entry.get("id"):

                delete_room(entry["id"])

            st.session_state.report_entries.pop(active_idx)

            st.session_state.active_room_idx = max(0, active_idx - 1)

            st.rerun()

    with action_cols[1]:

        if st.button("清空本社区上报"):

            clear_rooms(community)

            st.session_state.report_entries = []

            st.rerun()



    for i, e in enumerate(st.session_state.report_entries):

        env_text = f"{ENV_INSPECTION_LABEL}（必检）"

        device_text = (

            "、".join(f"{d['设备类型']}×{d['数量']}" for d in e["devices"])

            if e["devices"]

            else "（未添加设备）"

        )

        st.markdown(

            f"**{i + 1}. {e['机房名称']}** · {e.get('社区分类', community)} · {e['机房类型']}  \n"

            f"环境：{env_text}  \n"

            f"设备：{device_text}"

        )



    st.divider()

    _render_summary_table(inspection_items)



    export_rows = build_summary_rows(st.session_state.report_entries, inspection_items)

    if export_rows:

        export_df = _export_dataframe(export_rows)

        col_csv, col_xlsx = st.columns(2)

        with col_csv:

            csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")

            st.download_button(

                "导出 CSV",

                data=csv_bytes,

                file_name=f"巡检上报_{community}.csv",

                mime="text/csv",

                use_container_width=True,

            )

        with col_xlsx:

            st.download_button(

                "导出 Excel",

                data=_to_excel_bytes(export_df),

                file_name=f"巡检上报_{community}.xlsx",

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                use_container_width=True,

            )





if __name__ == "__main__":

    main()


