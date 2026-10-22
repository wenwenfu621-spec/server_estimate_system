"""
app.py - CAD 自動報價與尺寸辨識系統（板金件）
Version: v3.0.0_20261023
Description: 上傳 3D 模型 → 報價預覽表（可編輯材質、料厚、表面處理、利潤）→ 輸出 Excel / Word。

版本紀錄
--------
v3.0.0_20261023：
    1. 移除「客戶與報價表頭資訊」區塊。
    2. 辨識後先顯示報價預覽表（與 Excel 同欄位），可選材質、修改料厚、勾選
       表面處理（自動帶入預設加收金額，可逐格修改）、填寫利潤 %，確認後才輸出。
    3. 材質單價表：SECC/SGCC 參考中鋼盤價公告、SUS 參考冷軋 304 亞洲價、AL 參考
       LME 鋁價，皆標示出處，可手動修改（見 pricing.py）。
    4. 組立件（2 個以上實體）在預覽表與上方提醒中標示。
    5. 可分批上傳：每次辨識的檔案接在預覽表後面，最後一次輸出同一份 Excel / Word；
       同檔名不重複加入，預覽表可刪除列。
    6. 四視圖圖卡不在網頁顯示，只放在 Word 報告；Excel 改輸出 .xlsx 並在
       「示意圖」欄貼等角圖。
    7. 啟動時清理逾時暫存檔；session_state key 一律加 cadq_ 前綴。
v2.12.0_20260831：
    標題縮小、淡化。
v2.11.0_20260830：
    標題改為 position: fixed，對齊 Streamlit header 或入口網站問候語列。
v2.10.0_20260830：
    主畫面收進 render()，可獨立執行，也可被「Max 工具箱」入口網站當模組載入。
    st.set_page_config() 與 inject_custom_elements() 只在獨立執行時呼叫。
"""

import os
import io
import sys
import base64
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(_MODULE_DIR)
from cad_parser import parse_cad_with_screenshot, is_valid_image, temp_path, cleanup_stale_temp_files
from report_export import generate_excel_report, generate_word_report, compute_row
from pricing import fetch_market_prices, MATERIALS

APP_VERSION = "v3.0.0_20261023"

# 表面處理：(勾選欄, 金額欄, 預設價 key)
TREATMENT_COLS = [("噴漆", "噴漆費", "paint"), ("陽極", "陽極費", "anodize"), ("電鍍", "電鍍費", "plating")]
DEFAULT_FEES = {"paint": 0.0, "anodize": 0.0, "plating": 0.0}

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def inject_custom_elements():
    """注入左下角版本號標籤與中央個人識別頭像徽章"""
    avatar_candidates = ["avatar.jpg", "avatar.jpeg", "avatar.png", "avatar.JPG", "avatar.PNG", "Avatar.png"]
    img_base64 = ""
    mime_type = "image/png"

    for af in avatar_candidates:
        af_path = os.path.join(_MODULE_DIR, af)
        if os.path.exists(af_path):
            with open(af_path, "rb") as img_f:
                img_base64 = base64.b64encode(img_f.read()).decode("utf-8")
                mime_type = "image/jpeg" if af.lower().endswith((".jpg", ".jpeg")) else "image/png"
            break

    avatar_html = f'<img src="data:{mime_type};base64,{img_base64}" style="width: 36px; height: 36px; border-radius: 50%; object-fit: cover; margin-right: 8px; border: 1.5px solid #ccc; background-color: #fff;">' if img_base64 else ""

    custom_css = f"""
    <style>
    .version-badge-left {{
        position: fixed; bottom: 16px; left: 16px;
        background-color: rgba(240, 242, 246, 0.9); padding: 4px 12px;
        border-radius: 12px; font-family: monospace; font-size: 0.8rem;
        color: #555555; border: 1px solid #d0d0d0; z-index: 999999; pointer-events: none;
    }}
    .custom-footer-max {{
        position: fixed; bottom: 16px; left: 50%; transform: translateX(-50%);
        display: flex; align-items: center; background-color: rgba(255, 255, 255, 0.95);
        padding: 4px 14px; border-radius: 20px; box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.15);
        z-index: 999999; pointer-events: none;
    }}
    .custom-footer-text {{
        font-family: 'Comic Sans MS', cursive, sans-serif; font-weight: bold;
        font-style: italic; font-size: 0.95rem; color: #333333; white-space: nowrap;
    }}
    </style>
    <div class="version-badge-left">Version: {APP_VERSION}</div>
    <div class="custom-footer-max">{avatar_html}<span class="custom-footer-text">Design by Max</span></div>
    """
    st.markdown(custom_css, unsafe_allow_html=True)


@st.cache_data(ttl=6 * 3600, show_spinner="正在抓取材質牌價…")
def get_market_prices():
    return fetch_market_prices()


# ---------------------------------------------------------------- session state

def _init_state():
    defaults = {
        "cadq_uploader_ver": 0,
        "cadq_results": [],         # 成功解析的結果，順序與預覽表各列一致
        "cadq_messages": [],        # 最近一次辨識的結果 / 錯誤 / 略過訊息
        "cadq_rows": None,          # 預覽表 DataFrame
        "cadq_fee_custom": [],      # 每列 {金額欄: 是否為手動修改}
        "cadq_editor_ver": 0,
        "cadq_temp_files": [],
        "cadq_mat_table": None,     # 材質單價表 DataFrame
        "cadq_mat_ver": 0,
        "cadq_fees": dict(DEFAULT_FEES),
        "cadq_profit": 0.0,
        "cadq_excel": None,
        "cadq_word": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    if "cadq_temp_cleaned" not in st.session_state:
        cleanup_stale_temp_files()
        st.session_state.cadq_temp_cleaned = True


def _cleanup_temp():
    for fpath in st.session_state.get("cadq_temp_files") or []:
        if fpath and os.path.exists(fpath):
            try:
                os.remove(fpath)
            except Exception:
                pass
    st.session_state.cadq_temp_files = []


def _reset_session():
    _cleanup_temp()
    st.session_state.cadq_results = []
    st.session_state.cadq_messages = []
    st.session_state.cadq_rows = None
    st.session_state.cadq_fee_custom = []
    st.session_state.cadq_excel = None
    st.session_state.cadq_word = None
    st.session_state.cadq_uploader_ver += 1


def _invalidate_outputs():
    st.session_state.cadq_excel = None
    st.session_state.cadq_word = None


# ---------------------------------------------------------------- 材質單價表

def _build_mat_table(market) -> pd.DataFrame:
    prices = market["prices"]
    return pd.DataFrame(
        [{"材質": m, "密度(g/cm³)": prices[m]["density"], "單價(元/kg)": prices[m]["price"], "來源": prices[m]["source"]}
         for m in MATERIALS]
    )


def _material_lookup() -> dict:
    table = st.session_state.cadq_mat_table
    if table is None:
        return {}
    return {row["材質"]: (row["密度(g/cm³)"], row["單價(元/kg)"]) for _, row in table.iterrows()}


def _on_mat_edit(editor_key: str):
    table = st.session_state.cadq_mat_table.copy()
    for idx, change in st.session_state[editor_key]["edited_rows"].items():
        for col, val in change.items():
            table.at[int(idx), col] = val
    st.session_state.cadq_mat_table = table
    st.session_state.cadq_mat_ver += 1
    _recompute_rows()


def _refetch_prices():
    get_market_prices.clear()
    st.session_state.cadq_mat_table = None


# ---------------------------------------------------------------- 預覽表

def _thumbnail_uri(path) -> str:
    if not is_valid_image(path):
        return None
    with Image.open(path) as im:
        im = im.convert("RGB")
        im.thumbnail((160, 120))
        buf = io.BytesIO()
        im.save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def _row_notes(res) -> str:
    notes = []
    if res.get("is_assembly"):
        notes.append(f"⚠ 組立件（{res.get('solid_count')} 個實體），尺寸與料厚以整體計算")
    if res.get("thickness_suspicious"):
        notes.append("⚠ " + (res.get("thickness_note") or "請確認料厚"))
    if res.get("image_error"):
        notes.append("示意圖產生失敗")
    return "；".join(notes)


def _new_rows(results) -> pd.DataFrame:
    records = []
    for res in results:
        rec = {
            "序項": 0,
            "示意圖": _thumbnail_uri(res.get("iso_image_path")),
            "品名": res["file_name"],
            "材質": None,
            "料厚": res.get("thickness"),
            "尺寸": res["dimensions_str"],
            "重量(kg)": None,
        }
        for flag, fee, _ in TREATMENT_COLS:
            rec[flag] = False
            rec[fee] = None
        rec.update({"金額": None, "含利潤": None, "備註": "", "提醒": _row_notes(res)})
        records.append(rec)
    df = pd.DataFrame(records)
    for col in ["料厚", "重量(kg)", "金額", "含利潤"] + [fee for _, fee, _ in TREATMENT_COLS]:
        df[col] = df[col].astype("float64")
    # 必須是文字型別：整欄都是 None 的 object 欄會被 Streamlit 當成「空值型別」，下拉選單無法寫入
    df["材質"] = df["材質"].astype("string")
    return df


def _append_rows(results):
    """把新辨識的檔案接在預覽表後面（已填的內容保留）"""
    new_df = _new_rows(results)
    old_df = st.session_state.cadq_rows
    df = new_df if old_df is None or old_df.empty else pd.concat([old_df, new_df], ignore_index=True)
    df["序項"] = range(1, len(df) + 1)
    st.session_state.cadq_rows = df
    st.session_state.cadq_results = st.session_state.cadq_results + list(results)
    st.session_state.cadq_fee_custom += [{fee: False for _, fee, _ in TREATMENT_COLS} for _ in results]


def _delete_rows(indices):
    removed = set(indices)
    keep = [i for i in range(len(st.session_state.cadq_rows)) if i not in removed]
    for i in removed:
        for key in ("image_path", "iso_image_path"):
            path = st.session_state.cadq_results[i].get(key)
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass
    df = st.session_state.cadq_rows.iloc[keep].reset_index(drop=True)
    df["序項"] = range(1, len(df) + 1)
    st.session_state.cadq_rows = df
    st.session_state.cadq_results = [st.session_state.cadq_results[i] for i in keep]
    st.session_state.cadq_fee_custom = [st.session_state.cadq_fee_custom[i] for i in keep]


def _row_item(i: int, row, lookup: dict) -> dict:
    """把預覽表的一列組成 report_export 使用的 item"""
    res = st.session_state.cadq_results[i]
    material = row["材質"] if isinstance(row["材質"], str) and row["材質"] else None
    density, price = lookup.get(material, (None, None)) if material else (None, None)
    item = {
        "file_name": row["品名"],
        "material": material,
        "thickness": row["料厚"],
        "dimensions_str": row["尺寸"],
        "length": res["length"],
        "width": res["width"],
        "unit": res.get("unit", "mm"),
        "used_mode": res.get("used_mode"),
        "density": density,
        "unit_price": price,
        "remark": row["備註"] or "",
        "iso_image_path": res.get("iso_image_path"),
        "image_path": res.get("image_path"),
        "image_error": res.get("image_error"),
    }
    for flag, fee, key in TREATMENT_COLS:
        item[key] = row[fee] if row[flag] else None
    return item


def _recompute_rows():
    df = st.session_state.cadq_rows
    if df is None or df.empty:
        return
    lookup = _material_lookup()
    profit = st.session_state.cadq_profit or 0.0
    for i, row in df.iterrows():
        calc = compute_row(_row_item(i, row, lookup), profit)
        df.at[i, "重量(kg)"] = calc["weight"]
        df.at[i, "金額"] = calc["amount"]
        df.at[i, "含利潤"] = calc["with_profit"]
    _invalidate_outputs()


def _on_rows_edit(editor_key: str):
    df = st.session_state.cadq_rows
    fees = st.session_state.cadq_fees
    custom = st.session_state.cadq_fee_custom
    flag_cols = {flag: (fee, key) for flag, fee, key in TREATMENT_COLS}
    fee_cols = {fee: flag for flag, fee, _ in TREATMENT_COLS}

    state = st.session_state[editor_key]
    for idx, change in state["edited_rows"].items():
        i = int(idx)
        # 先處理勾選，再處理金額：同一次編輯若兩者都改，以手填金額為準
        for col, val in change.items():
            if col in flag_cols:
                fee_col, key = flag_cols[col]
                df.at[i, col] = bool(val)
                df.at[i, fee_col] = fees.get(key, 0.0) if val else None
                custom[i][fee_col] = False
        for col, val in change.items():
            if col in fee_cols:
                if val is None or pd.isna(val):
                    df.at[i, fee_cols[col]] = False
                    df.at[i, col] = None
                    custom[i][col] = False
                else:
                    df.at[i, fee_cols[col]] = True
                    df.at[i, col] = float(val)
                    custom[i][col] = True
            elif col not in flag_cols:
                df.at[i, col] = val

    # 編輯的列號以刪除前為準，所以先套用編輯再刪列
    if state.get("deleted_rows"):
        _delete_rows(sorted(int(i) for i in state["deleted_rows"]))

    st.session_state.cadq_editor_ver += 1
    _recompute_rows()


def _on_fee_default_change(key: str, fee_col: str, flag_col: str):
    """預設加收金額變更時，同步更新已勾選、且沒有手動改過金額的列"""
    new_value = st.session_state[f"cadq_fee_input_{key}"]
    st.session_state.cadq_fees[key] = new_value
    df = st.session_state.cadq_rows
    if df is not None:
        for i in range(len(df)):
            if df.at[i, flag_col] and not st.session_state.cadq_fee_custom[i][fee_col]:
                df.at[i, fee_col] = new_value
        st.session_state.cadq_editor_ver += 1
    _recompute_rows()


def _on_profit_change():
    st.session_state.cadq_profit = st.session_state.cadq_profit_input
    _recompute_rows()


def _apply_bulk_material():
    """批次設定材質：套用到全部列，或只套用到還沒選材質的列"""
    df = st.session_state.cadq_rows
    material = st.session_state.cadq_bulk_material
    if df is None or not material:
        return
    only_empty = st.session_state.cadq_bulk_scope == "只套用到未選材質的列"
    for i in range(len(df)):
        if only_empty and isinstance(df.at[i, "材質"], str) and df.at[i, "材質"]:
            continue
        df.at[i, "材質"] = material
    st.session_state.cadq_editor_ver += 1
    _recompute_rows()


# ---------------------------------------------------------------- 畫面

def _render_title():
    st.markdown(
        """
        <div id="cad-title-spacer"></div>
        <div id="cad-title-bar" style="
            position: fixed; top: 0; left: 0; width: 100%; z-index: 999980;
            text-align: center; margin: 0; padding: 0.6rem 0;
            background: rgba(255, 255, 255, 0.55);
            backdrop-filter: blur(6px); -webkit-backdrop-filter: blur(6px);
        ">
            <h1 style="margin: 0; color: rgba(148, 163, 184, 0.45); font-weight: 400; font-size: 15px;">CAD 自動報價與尺寸辨識系統</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )
    components.html(
        """
        <script>
        (function () {
            function getDoc() { try { return window.parent.document; } catch (e) { return null; } }
            function align(doc) {
                var bar = doc.getElementById("cad-title-bar");
                var spacer = doc.getElementById("cad-title-spacer");
                if (!bar) return;
                var header = doc.querySelector('header[data-testid="stHeader"]');
                var top = 0, left = 0, width = "100%";
                if (header) {
                    var hRect = header.getBoundingClientRect();
                    top = hRect.bottom;
                    left = hRect.left;
                    width = hRect.width + "px";
                }
                var greetingBar = doc.getElementById("portal-greeting-bar");
                if (greetingBar) top += greetingBar.getBoundingClientRect().height;
                bar.style.top = top + "px";
                bar.style.left = left + "px";
                bar.style.width = width;
                if (spacer) spacer.style.height = bar.getBoundingClientRect().height + "px";
            }
            var doc = getDoc();
            if (!doc) return;
            if (!doc.__cadTitlePinAttached) {
                doc.__cadTitlePinAttached = true;
                doc.defaultView.setInterval(function () { align(doc); }, 1000);
            }
            align(doc);
        })();
        </script>
        """,
        height=0,
    )


def _render_bbox_explainer():
    with st.expander("📐 加工素材計算採用下列兩種方式之最小值", expanded=False):
        col_obb, col_aabb = st.columns(2)

        with col_obb:
            st.markdown(
                """
                <svg viewBox="0 0 220 155" xmlns="http://www.w3.org/2000/svg" style="width:100%; height:auto;">
                    <rect x="10" y="10" width="200" height="135" rx="10" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1"/>
                    <g transform="translate(110,60) rotate(-18)">
                        <rect x="-70" y="-11" width="140" height="22" rx="6" fill="#64748b"/>
                        <rect x="-84" y="-24" width="168" height="48" rx="4" fill="none" stroke="#3b82f6" stroke-width="2.5" stroke-dasharray="6,4"/>
                    </g>
                    <text x="110" y="132" text-anchor="middle" font-size="13" fill="#334155" font-family="sans-serif" font-weight="bold">最小素材包容盒 (OBB)</text>
                </svg>
                """,
                unsafe_allow_html=True
            )
            st.markdown("**【OBB 最小素材包容盒】**")
            st.markdown(
                "- 貼合零件幾何方向旋轉量測\n"
                "- 素材精確省料\n"
                "- 適用：CNC 實體備料估價"
            )

        with col_aabb:
            st.markdown(
                """
                <svg viewBox="0 0 220 155" xmlns="http://www.w3.org/2000/svg" style="width:100%; height:auto;">
                    <rect x="10" y="10" width="200" height="135" rx="10" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1"/>
                    <g transform="translate(110,60) rotate(-18)">
                        <rect x="-70" y="-11" width="140" height="22" rx="6" fill="#64748b"/>
                    </g>
                    <rect x="42" y="25" width="136" height="70" fill="none" stroke="#ef4444" stroke-width="2.5" stroke-dasharray="6,4"/>
                    <text x="110" y="132" text-anchor="middle" font-size="13" fill="#334155" font-family="sans-serif" font-weight="bold">標準投影外框 (AABB)</text>
                </svg>
                """,
                unsafe_allow_html=True
            )
            st.markdown("**【AABB 標準投影外框】**")
            st.markdown(
                "- 沿世界座標軸 (X/Y/Z) 垂直外包\n"
                "- 包含傾斜投影落差\n"
                "- 適用：外箱體積、正交零件估價"
            )


def _run_parsing(uploaded_files):
    """辨識這一批檔案並接到預覽表後面；與清單中同檔名的檔案略過"""
    _invalidate_outputs()
    existing = {r["file_name"] for r in st.session_state.cadq_results}
    messages, batch, todo = [], [], []
    for up_file in uploaded_files:
        if up_file.name in existing:
            messages.append(("warning", f"⏭️ {up_file.name} 已在清單中，略過不重複加入。"))
        else:
            existing.add(up_file.name)
            todo.append(up_file)

    progress_bar = st.progress(0)
    status_text = st.empty()
    for idx, up_file in enumerate(todo):
        fname = up_file.name
        status_text.text(f"正在處理 3D 檔案 ({idx + 1}/{len(todo)}): {fname} ...")
        tmp = temp_path(os.path.splitext(fname)[1])
        with open(tmp, "wb") as f:
            f.write(up_file.getvalue())

        res = parse_cad_with_screenshot(tmp)
        res["file_name"] = fname
        try:
            os.remove(tmp)
        except Exception:
            pass
        if res.get("status") == "success":
            batch.append(res)
            for key in ("image_path", "iso_image_path"):
                if res.get(key):
                    st.session_state.cadq_temp_files.append(res[key])
        else:
            messages.append(("error", f"❌ {fname} 解析失敗：{res.get('error_message')}"))
        progress_bar.progress((idx + 1) / max(1, len(todo)))

    if batch:
        _append_rows(batch)
        messages.insert(0, ("success", f"🎉 已加入 {len(batch)} 個檔案，清單共 {len(st.session_state.cadq_results)} 個。"
                                       "可繼續上傳下一批，最後再一起輸出。"))
    st.session_state.cadq_messages = messages
    st.session_state.cadq_editor_ver += 1
    st.session_state.cadq_uploader_ver += 1   # 清空上傳框，方便上傳下一批
    _recompute_rows()
    st.rerun()


def _render_price_settings():
    if st.session_state.cadq_mat_table is None:
        market = get_market_prices()
        st.session_state.cadq_mat_table = _build_mat_table(market)
        st.session_state.cadq_mat_ver += 1
    market = get_market_prices()

    with st.expander("💰 材質單價與表面處理設定", expanded=True):
        st.markdown("**材質單價表**（可直接修改密度與單價，修改後預覽表即時重算）")
        mat_key = f"cadq_mat_editor_{st.session_state.cadq_mat_ver}"
        st.data_editor(
            st.session_state.cadq_mat_table,
            key=mat_key,
            on_change=_on_mat_edit,
            args=(mat_key,),
            hide_index=True,
            disabled=["材質", "來源"],
            column_config={
                "材質": st.column_config.TextColumn(width="small"),
                "密度(g/cm³)": st.column_config.NumberColumn(min_value=0.0, step=0.01, format="%.2f", width="small"),
                "單價(元/kg)": st.column_config.NumberColumn(min_value=0.0, step=0.1, format="%.2f", width="small"),
                "來源": st.column_config.TextColumn(width="large"),
            },
        )
        csc = market.get("csc")
        caption = f"抓取日期：{market['fetched']}；匯率 USD→TWD {market['fx']:.2f}（{market['fx_source']}）"
        if csc:
            caption += f"；中鋼公告：[{csc['title']}]({csc['url']})（{csc['date']}）"
        st.caption(caption)
        st.button("🔄 重新抓取牌價（會覆蓋手動修改的單價）", on_click=_refetch_prices)

        st.markdown("**表面處理預設加收金額**（勾選後自動帶入，預覽表中每格仍可個別修改）")
        cols = st.columns(len(TREATMENT_COLS) + 1)
        for col, (flag, fee, key) in zip(cols, TREATMENT_COLS):
            with col:
                st.number_input(
                    f"{flag}（元）", min_value=0.0, step=10.0,
                    value=float(st.session_state.cadq_fees.get(key, 0.0)),
                    key=f"cadq_fee_input_{key}",
                    on_change=_on_fee_default_change, args=(key, fee, flag),
                )
        with cols[-1]:
            st.number_input(
                "利潤（%）", min_value=0.0, step=1.0,
                value=float(st.session_state.cadq_profit),
                key="cadq_profit_input", on_change=_on_profit_change,
            )


def _render_preview():
    for level, text in st.session_state.cadq_messages:
        getattr(st, level)(text)

    assemblies = [r for r in st.session_state.cadq_results if r.get("is_assembly")]
    if assemblies:
        names = "、".join(f"{r['file_name']}（{r['solid_count']} 個實體）" for r in assemblies)
        st.warning(
            f"⚠️ 下列檔案為**組立件**：{names}。\n\n"
            "組立件的尺寸、料厚與重量是以整個組立計算（可能包含螺帽、鉚釘等五金），請確認後再報價。"
        )

    df = st.session_state.cadq_rows
    if df is None or df.empty:
        return

    st.subheader(f"📋 報價預覽表（共 {len(df)} 個檔案）")

    c_mat, c_scope, c_btn = st.columns([1, 2, 1], vertical_alignment="bottom")
    with c_mat:
        st.selectbox("批次設定材質", MATERIALS, key="cadq_bulk_material")
    with c_scope:
        st.radio("套用範圍", ["套用到全部列", "只套用到未選材質的列"], key="cadq_bulk_scope", horizontal=True)
    with c_btn:
        st.button("套用材質", on_click=_apply_bulk_material, width="stretch")

    st.caption("材質、料厚、表面處理、備註可直接編輯：**雙擊**儲存格（或點一下後按 Enter）開啟下拉選單或輸入框。"
               "金額 = 重量 × 單價 + 表面處理費，含利潤 = 金額 × (1 + 利潤%)。"
               "要刪除某列：勾選該列最左側方框，再按表格右上角的垃圾桶。")

    editor_key = f"cadq_rows_editor_{st.session_state.cadq_editor_ver}"
    column_config = {
        "序項": st.column_config.NumberColumn(width="small", pinned=True),
        "示意圖": st.column_config.ImageColumn(width="small"),
        "品名": st.column_config.TextColumn(width="medium"),
        "材質": st.column_config.SelectboxColumn(options=MATERIALS, width="small"),
        "料厚": st.column_config.NumberColumn(min_value=0.0, step=0.01, format="%.2f", width="small",
                                            help="自動量測，可手動修改（mm）"),
        "尺寸": st.column_config.TextColumn(width="small"),
        "重量(kg)": st.column_config.NumberColumn(format="%.3f", width="small"),
        "金額": st.column_config.NumberColumn(format="%.0f", width="small"),
        "含利潤": st.column_config.NumberColumn(format="%.0f", width="small"),
        "備註": st.column_config.TextColumn(width="medium"),
        "提醒": st.column_config.TextColumn(width="large"),
    }
    for flag, fee, _ in TREATMENT_COLS:
        column_config[flag] = st.column_config.CheckboxColumn(width="small")
        column_config[fee] = st.column_config.NumberColumn(min_value=0.0, step=1.0, format="%.0f", width="small")

    st.data_editor(
        df,
        key=editor_key,
        on_change=_on_rows_edit,
        args=(editor_key,),
        hide_index=True,
        num_rows="delete",
        row_height=64,
        placeholder="",
        disabled=["序項", "示意圖", "品名", "尺寸", "重量(kg)", "金額", "含利潤", "提醒"],
        column_config=column_config,
    )

    total_amount = df["金額"].fillna(0).sum()
    total_profit = df["含利潤"].fillna(0).sum()
    c1, c2 = st.columns(2)
    c1.metric("金額合計", f"{total_amount:,.0f} 元")
    c2.metric(f"含利潤合計（{st.session_state.cadq_profit:g}%）", f"{total_profit:,.0f} 元")


def _generate_outputs():
    df = st.session_state.cadq_rows
    lookup = _material_lookup()
    items = [_row_item(i, row, lookup) for i, row in df.iterrows()]
    profit = st.session_state.cadq_profit or 0.0

    excel_path = temp_path(".xlsx")
    st.session_state.cadq_temp_files.append(excel_path)
    generate_excel_report(items, excel_path, profit_pct=profit)
    with open(excel_path, "rb") as f:
        st.session_state.cadq_excel = f.read()

    try:
        word_path = temp_path(".docx")
        st.session_state.cadq_temp_files.append(word_path)
        generate_word_report(items, word_path, profit_pct=profit)
        with open(word_path, "rb") as f:
            st.session_state.cadq_word = f.read()
    except Exception as e_w:
        st.session_state.cadq_word = None
        st.warning(f"⚠️ Word 產生警告: {str(e_w)}")


def _render_outputs():
    st.markdown("---")
    if st.session_state.cadq_rows is not None and not st.session_state.cadq_rows.empty:
        if st.button("📤 確認並產生 Excel / Word", type="primary", width="stretch"):
            with st.spinner("正在產生報表…"):
                _generate_outputs()

    today_str = datetime.now().strftime("%Y%m%d")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.session_state.cadq_excel:
            st.download_button("📊 下載 Excel 報價單", st.session_state.cadq_excel,
                               f"板金報價_{today_str}.xlsx", mime=XLSX_MIME, width="stretch")
        else:
            st.button("📊 Excel 尚未產生", disabled=True, width="stretch")
    with col2:
        if st.session_state.cadq_word:
            st.download_button("📝 下載圖文報告 (.docx)", st.session_state.cadq_word,
                               f"板金報價_{today_str}.docx", mime=DOCX_MIME, width="stretch")
        else:
            st.button("📝 Word 尚未產生", disabled=True, width="stretch")
    with col3:
        st.button("🔄 清空清單重新開始", on_click=_reset_session, width="stretch")
    if st.session_state.cadq_rows is not None and not st.session_state.cadq_excel:
        st.caption("預覽表修改後需重新按「確認並產生」，下載的檔案才會是最新內容。")


def render() -> None:
    """渲染 CAD 自動報價與尺寸辨識系統的完整功能介面。

    獨立執行（見檔案最下面的 __main__ 區塊）與被入口網站當成模組載入時
    共用這支函式；st.set_page_config() 與版面裝飾不放在這裡。
    """
    _init_state()
    _render_title()
    st.write("上傳 `.step`, `.stp`, `.igs`, `.iges` 板金 3D 模型檔，自動計算素材尺寸與料厚，"
             "在預覽表確認材質、表面處理與利潤後，輸出 Excel 報價單與 Word 圖文報告。")
    _render_bbox_explainer()

    uploaded_files = st.file_uploader(
        "上傳 3D CAD 圖檔 (可多選)",
        type=["step", "stp", "igs", "iges"],
        accept_multiple_files=True,
        key=f"cadq_uploader_{st.session_state.cadq_uploader_ver}"
    )
    if uploaded_files:
        st.write(f"已選擇 **{len(uploaded_files)}** 個 3D 檔案。")
        label = "➕ 辨識並加入清單" if st.session_state.cadq_results else "🎯 開始 3D 幾何辨識"
        if st.button(label, type="primary", width="stretch"):
            _run_parsing(uploaded_files)

    if st.session_state.cadq_results or st.session_state.cadq_messages:
        _render_price_settings()
        _render_preview()
        _render_outputs()


if __name__ == "__main__":
    st.set_page_config(page_title="CAD 自動報價與尺寸辨識系統", page_icon="⚙️", layout="wide")
    inject_custom_elements()
    render()
