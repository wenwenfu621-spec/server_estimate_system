"""
app.py - CAD 自動報價與尺寸辨識系統
Version: v2.12.0_20260831
Description: 已移除 2D Mylar 辨識分頁，系統啟動後直接顯示 3D CAD 辨識功能。
             合併原 pages/1_3D_CAD.py 內容，改為單頁應用（無多頁側邊導覽）。

版本紀錄
--------
v2.12.0_20260831：
    標題再縮小、再淡化：字體從原本的 h1 預設大小（約 44px）縮到 15px
    （約 1/3），顏色從 #94a3b8 改成 rgba(148, 163, 184, 0.45)（同色調再
    加上透明度），讓標題更不搶眼、隱約可見即可，跟畫面底圖融合。
v2.11.0_20260830：
    標題從小灰字副標題（0.85rem）改成跟矽膠模成本估算統一規格的 <h1>
    （字體大小維持瀏覽器/Streamlit 預設 h1 大小，跟矽膠模那支程式一致，
    顏色沿用原本的淡灰色 #94a3b8、font-weight 400）；並改成
    position: fixed 固定在視窗最上方，往下捲動頁面也會一直留在原地。用
    components.html 注入的 JS 讀取 Streamlit 自己的頂端 header 位置/寬度
    來對齊，若這支程式是被「Max 工具箱」入口網站載入、畫面上有
    id="portal-greeting-bar" 的問候語列，標題列會自動疊在它正下方；獨立
    執行（沒有問候語列）時則貼齊 header 正下方。用一個 spacer 元素撐開
    原本標題的高度，避免標題變成 fixed 後，底下的內容跟著往上跳一段。
v2.10.0_20260830：
    把整支程式的主畫面邏輯收進 render() 函式，讓這支 app.py 除了可以像以前
    一樣用 `streamlit run app.py` 獨立執行，也可以被「Max 工具箱」入口網站
    當成模組載入、在同一個 Streamlit process 裡呼叫 render() 顯示在中央工作
    區。st.set_page_config() 與 inject_custom_elements()（版本標籤/署名）
    只在獨立執行時呼叫一次，被入口網站載入時由入口網站統一處理，避免
    st.set_page_config 被呼叫兩次而噴錯、以及版本標籤重複顯示。
    同時把 has_xlsm 判斷 template.xlsm 是否存在的邏輯，改成以本檔案所在
    目錄（__file__）為準的絕對路徑，不再依賴執行時的 working directory
    （被入口網站以不同目錄啟動時，原本的相對路徑會找不到範本檔）。
"""

import streamlit as st
import streamlit.components.v1 as components
import os
import sys
import base64
import tempfile
from datetime import datetime

_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(_MODULE_DIR)
from cad_parser import parse_cad_with_screenshot, generate_excel_report, generate_word_report, is_valid_image


def inject_custom_elements():
    """注入左下角版本號標籤與中央個人識別頭像徽章"""
    app_dir = os.path.dirname(os.path.abspath(__file__))
    avatar_candidates = ["avatar.jpg", "avatar.jpeg", "avatar.png", "avatar.JPG", "avatar.PNG", "Avatar.png"]
    img_base64 = ""
    mime_type = "image/png"

    for af in avatar_candidates:
        af_path = os.path.join(app_dir, af)
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
    <div class="version-badge-left">Version: v2.9.0_20260825</div>
    <div class="custom-footer-max">{avatar_html}<span class="custom-footer-text">Design by Max</span></div>
    """
    st.markdown(custom_css, unsafe_allow_html=True)


def cleanup_3d_temp():
    if "temp_files_3d" in st.session_state and st.session_state.temp_files_3d:
        for fpath in st.session_state.temp_files_3d:
            if fpath and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass
        st.session_state.temp_files_3d = []


def reset_3d_session():
    cleanup_3d_temp()
    st.session_state.parsed_3d_results = None
    st.session_state.excel_3d_bytes = None
    st.session_state.word_3d_bytes = None
    st.session_state.uploader_3d_key += 1


def render() -> None:
    """渲染 CAD 自動報價與尺寸辨識系統的完整功能介面。

    獨立執行（見檔案最下面的 __main__ 區塊）與被入口網站當成模組載入時
    共用這支函式；st.set_page_config() 與版面裝飾不放在這裡，各自的呼叫端
    只需要負責呼叫一次。
    """
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
    st.write("上傳 `.step`, `.stp`, `.igs`, `.iges` 3D 模型檔，自動計算 OBB/AABB 包容盒與第三角法三視圖。")

    if "uploader_3d_key" not in st.session_state:
        st.session_state.uploader_3d_key = 0
    if "parsed_3d_results" not in st.session_state:
        st.session_state.parsed_3d_results = None
    if "excel_3d_bytes" not in st.session_state:
        st.session_state.excel_3d_bytes = None
    if "word_3d_bytes" not in st.session_state:
        st.session_state.word_3d_bytes = None
    if "temp_files_3d" not in st.session_state:
        st.session_state.temp_files_3d = []

    with st.expander("📝 客戶與報價表頭資訊填寫", expanded=True):
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            customer_input = st.text_input("客戶名稱", key=f"cust_3d_{st.session_state.uploader_3d_key}")
            phone_input = st.text_input("聯絡電話", key=f"phone_3d_{st.session_state.uploader_3d_key}")
        with col_c2:
            contact_input = st.text_input("聯絡人", key=f"contact_3d_{st.session_state.uploader_3d_key}")
            fax_input = st.text_input("傳真", key=f"fax_3d_{st.session_state.uploader_3d_key}")

    with st.expander("📐 加工素材計算採用下列兩種方式之最小值", expanded=True):
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

    uploaded_3d_files = st.file_uploader(
        "上傳 3D CAD 圖檔 (可多選)",
        type=["step", "stp", "igs", "iges"],
        accept_multiple_files=True,
        key=f"uploader_3d_{st.session_state.uploader_3d_key}"
    )

    if uploaded_3d_files:
        st.write(f"已選擇 **{len(uploaded_3d_files)}** 個 3D 檔案。")

        if st.button("🎯 開始 3D 幾何與工程視圖辨識", type="primary", width="stretch"):
            cleanup_3d_temp()
            parsed_results = []
            progress_bar = st.progress(0)
            status_text = st.empty()

            for idx, up_file in enumerate(uploaded_3d_files):
                fname = up_file.name
                ext = os.path.splitext(fname)[1]
                status_text.text(f"正在處理 3D 檔案 ({idx+1}/{len(uploaded_3d_files)}): {fname} ...")

                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(up_file.getvalue())
                    tmp_path = tmp.name
                    st.session_state.temp_files_3d.append(tmp_path)

                res = parse_cad_with_screenshot(tmp_path)
                res["file_name"] = fname
                parsed_results.append(res)

                if res.get("image_path"):
                    st.session_state.temp_files_3d.append(res["image_path"])

                progress_bar.progress((idx + 1) / len(uploaded_3d_files))

            status_text.success("🎉 3D 檔案全部辨識完成！")

            header_info = {"customer": customer_input, "contact": contact_input, "phone": phone_input, "fax": fax_input}

            # --- 處理 Excel 副檔名與 MIME ---
            has_xlsm = os.path.exists(os.path.join(_MODULE_DIR, "template.xlsm")) or os.path.exists(os.path.join(_MODULE_DIR, "Template.xlsm"))
            export_ext = ".xlsm" if has_xlsm else ".xlsx"

            excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=export_ext)
            excel_path = excel_tmp.name
            excel_tmp.close()
            st.session_state.temp_files_3d.append(excel_path)

            generate_excel_report(parsed_results, excel_path, header_info=header_info)
            with open(excel_path, "rb") as f:
                excel_bytes = f.read()

            word_bytes = None
            try:
                word_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                word_path = word_tmp.name
                word_tmp.close()
                st.session_state.temp_files_3d.append(word_path)

                generate_word_report(parsed_results, word_path, header_info=header_info)
                with open(word_path, "rb") as f_w:
                    word_bytes = f_w.read()
            except Exception as e_w:
                st.warning(f"⚠️ Word 產生警告: {str(e_w)}")

            st.session_state.parsed_3d_results = parsed_results
            st.session_state.excel_3d_bytes = excel_bytes
            st.session_state.word_3d_bytes = word_bytes

    if st.session_state.parsed_3d_results:
        st.subheader("📋 3D 辨識結果預覽")
        for res in st.session_state.parsed_3d_results:
            if res.get("status") == "success":
                dims_esc = res['dimensions_str'].replace("*", "\\*")
                mode = res.get("used_mode", "OBB")
                with st.expander(f"📄 {res['file_name']} - 【尺寸：{dims_esc} mm ({mode})】"):
                    col1, col2 = st.columns([2, 3])
                    with col1:
                        st.write(f"**檔名**：{res['file_name']}")
                        st.write(f"**尺寸**：{dims_esc} mm ({mode})")
                        st.write(f"**演算法**：{mode} 模式")
                        st.write(f"**拆解數值**：長 {res.get('length')} / 寬 {res.get('width')} / 高 {res.get('height')}")
                    with col2:
                        if is_valid_image(res.get("image_path")):
                            st.image(res["image_path"], caption="CAD 工程三視圖圖卡 (第三角法)", use_container_width=True)
                        else:
                            st.warning("⚠️ 預覽圖產生失敗")
            else:
                st.error(f"❌ {res['file_name']} 解析失敗：{res.get('error_message')}")

        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        today_str = datetime.now().strftime("%Y%m%d")

        # --- 輸出按鈕參數 ---
        has_xlsm = os.path.exists(os.path.join(_MODULE_DIR, "template.xlsm")) or os.path.exists(os.path.join(_MODULE_DIR, "Template.xlsm"))
        download_ext = ".xlsm" if has_xlsm else ".xlsx"
        download_mime = "application/vnd.ms-excel.sheet.macroEnabled.12" if has_xlsm else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        with col1:
            st.download_button(
                "📊 下載 Excel 報價單",
                st.session_state.excel_3d_bytes,
                f"3D_Report_{today_str}{download_ext}",
                mime=download_mime,
                width="stretch"
            )
        with col2:
            if st.session_state.word_3d_bytes:
                st.download_button("📝 下載圖文報告 (.docx)", st.session_state.word_3d_bytes, f"3D_Report_{today_str}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", width="stretch")
            else:
                st.button("📝 Word 不可用", disabled=True, width="stretch")
        with col3:
            st.button("🔄 重置頁面", on_click=reset_3d_session, width="stretch")


if __name__ == "__main__":
    st.set_page_config(page_title="CAD 自動報價與尺寸辨識系統", page_icon="⚙️", layout="centered")
    inject_custom_elements()
    render()
