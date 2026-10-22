"""
cad_parser.py - 3D CAD 解析與報表產出模組
Version: v2.9.5_20261022
Description: 專責 3D STEP/IGES 管道，與 2D DXF 管道完全分離。
             修正 Excel 匯出邏輯：嚴格對齊 template.xlsm 與 xlsm 副檔名與 MIME 類型。

版本紀錄
--------
v2.9.5_20261022：
    1. 修正旋轉零件尺寸多 1mm：Bnd_OBB 回傳值帶 +2e-7 的容差雜訊（100 會變
       100.0000002），原本直接 math.ceil 會進位成 101。實測理論 100x50x10 的零件
       旋轉 15/30/45 度會輸出 101*51*11。新增 ceil_dim()，進位前先扣 1e-6。
    2. 修正四視圖圖卡中文標籤變方框：原本只用 DejaVuSans（沒有 CJK 字形）。新增
       load_label_font()，依序找 Noto CJK / 文鼎 / 文泉驛 / Windows 微軟正黑體 /
       macOS 蘋方，都找不到才退回 DejaVuSans。伺服器需安裝 CJK 字型套件
       （例如 fonts-noto-cjk；portal 的 Dockerfile 已有 fonts-arphic-uming）。
v2.9.4_20260904：
    parse_cad_with_screenshot() 高度（三軸中最短的一軸）小於 1mm 時，
    原本一律用 math.ceil() 無條件進位成 1mm，導致薄件（例如 0.2~0.9mm
    的墊片/薄板）的高度數字失真。改成排序時保留原始浮點數（sorted_raw），
    只有高度 < 1 時才跳過無條件進位、直接顯示四捨五入到小數點後兩位的
    實際量測值；長度、寬度維持原本無條件進位邏輯不變。OBB/AABB 兩種模式
    的體積比較邏輯（用於決定採用哪種包容盒）不受影響，因為兩種寫法算出
    來的體積本來就一樣（進位後三數相乘，跟先排序再進位是同一個值）。
v2.9.3_20260830：
    cairosvg 的 import 保護原本只接 ImportError，但 cairosvg 底層
    （cairocffi）在 Windows 找不到原生 Cairo 函式庫時拋的是 OSError，
    接不住會讓整支 cad_parser.py 在 import 階段直接崩潰，而不是照原本
    設計優雅退化成 HAS_CAIROSVG=False。改成接 Exception。在整合到「Max
    工具箱」入口網站、實際跑起來測試時發現（這台機器沒裝 Cairo 原生
    函式庫），屬於既有程式的環境相容性問題，與本次整合的改動無關。
v2.9.2_20260830：
    generate_excel_report() 尋找 template.xlsm/xlsx 的邏輯，原本用
    os.path.exists("template.xlsm") 相對於「執行時的 working directory」找檔，
    當這支程式被入口網站（Portal）當成模組載入、以不同目錄啟動時會找不到範本、
    誤判成「沒有範本」而用空白活頁簿輸出。改成以本檔案自身所在目錄
    （__file__）為準組出絕對路徑，不再受呼叫端 cwd 影響。
"""

import os
import glob
import math
import uuid
import tempfile
from datetime import datetime
from typing import Dict, Any, List, Optional
import cadquery as cq
from PIL import Image, ImageDraw, ImageFont
import openpyxl
from openpyxl.styles import Font

# 引入 docx 模組
try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# 引入 CairoSVG 向量轉譯模組
# 注意：cairosvg 底層透過 cairocffi 動態載入系統原生 Cairo 函式庫，Windows
# 沒裝 GTK/Cairo 執行環境時會拋 OSError（不是 ImportError），原本只接
# ImportError 接不住，會讓整支 cad_parser.py 在 import 階段直接崩潰，
# 而不是照原本設計優雅退化成「沒有這個功能」。改成接 Exception。
try:
    import cairosvg
    HAS_CAIROSVG = True
except Exception:
    HAS_CAIROSVG = False

# 引入 OpenCASCADE 原生 Bnd_OBB 模組以進行精確幾何最小包容盒計算
try:
    from OCP.Bnd import Bnd_OBB
    from OCP.BRepBndLib import BRepBndLib
    HAS_OCP_OBB = True
except ImportError:
    HAS_OCP_OBB = False


def safe_str(val: Any) -> str:
    """安全轉為字串並去除首尾空白"""
    if val is None:
        return ""
    try:
        return str(val).strip()
    except Exception:
        return ""


def set_cell_value_safe(ws, row: int, col: int, value: Any, font: Optional[Font] = None):
    """安全填寫儲存格：若遇到 MergedCell，寫入其最左上角的 Master Cell"""
    try:
        cell = ws.cell(row=row, column=col)
        if type(cell).__name__ == 'MergedCell':
            for rng in ws.merged_cells.ranges:
                if cell.coordinate in rng:
                    master_cell = ws.cell(row=rng.min_row, column=rng.min_col)
                    master_cell.value = value
                    if font:
                        master_cell.font = font
                    return
        cell.value = value
        if font:
            cell.font = font
    except Exception:
        pass


def is_valid_image(path: Optional[str]) -> bool:
    """驗證圖片檔是否存在、大小大於 0、且可被 Pillow 正確讀取"""
    if not path or not os.path.exists(path):
        return False
    if os.path.getsize(path) <= 0:
        return False
    try:
        with Image.open(path) as img:
            img.verify()
        return True
    except Exception:
        return False


def calculate_obb_dimensions(model: cq.Workplane) -> List[float]:
    """使用 OpenCASCADE 原生 Bnd_OBB 精確計算 3D 實體的最小素材包容盒 (OBB)"""
    if HAS_OCP_OBB:
        try:
            occ_shape = model.val().wrapped
            obb = Bnd_OBB()
            BRepBndLib.AddOBB_s(occ_shape, obb, True, True, True)
            dims = [obb.XHSize() * 2.0, obb.YHSize() * 2.0, obb.ZHSize() * 2.0]
            if all(d > 0 for d in dims):
                return dims
        except Exception:
            pass

    bbox = model.val().BoundingBox()
    return [bbox.xlen, bbox.ylen, bbox.zlen]


_DIM_EPS = 1e-6

_CJK_FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "C:/Windows/Fonts/msjh.ttc",
    "C:/Windows/Fonts/mingliu.ttc",
    "/System/Library/Fonts/PingFang.ttc",
]
_CJK_FONT_GLOBS = [
    "/usr/share/fonts/**/NotoSansCJK*-Regular.*",
    "/usr/share/fonts/**/uming.*",
    "/usr/share/fonts/**/wqy-*.ttc",
]


def ceil_dim(value: float) -> int:
    """無條件進位，但先扣掉量測容差：OBB 回傳值帶 +2e-7 的雜訊（例如 100 會變 100.0000002），
    直接 math.ceil 會讓剛好是整數的尺寸多進 1mm。"""
    return math.ceil(value - _DIM_EPS)


def load_label_font(size: int):
    """載入能顯示中文的字型；找不到時才退回 DejaVuSans / Pillow 內建字型（中文會變方框）。"""
    paths = [p for p in _CJK_FONT_CANDIDATES if os.path.exists(p)]
    if not paths:
        for pattern in _CJK_FONT_GLOBS:
            paths.extend(glob.glob(pattern, recursive=True))
    for path in paths:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    try:
        return ImageFont.truetype("DejaVuSans.ttf", size)
    except Exception:
        return ImageFont.load_default()


def render_single_view(
    model: cq.Workplane, 
    proj_dir: tuple, 
    out_png_path: str, 
    out_w: int = 800, 
    out_h: int = 800
) -> bool:
    """將 3D 模型以指定的投影方向導出為 SVG，並使用 cairosvg 轉為指定大小之 PNG"""
    temp_dir = tempfile.gettempdir()
    unique_id = str(uuid.uuid4())[:8]
    svg_path = os.path.join(temp_dir, f"view_{unique_id}.svg")

    try:
        cq.exporters.export(model, svg_path, opt={"projectionDir": proj_dir, "showHidden": False})
        if not os.path.exists(svg_path) or os.path.getsize(svg_path) <= 0:
            return False

        if HAS_CAIROSVG:
            cairosvg.svg2png(url=svg_path, write_to=out_png_path, output_width=out_w, output_height=out_h)
            return is_valid_image(out_png_path)
        return False
    except Exception:
        return False
    finally:
        if os.path.exists(svg_path):
            try:
                os.remove(svg_path)
            except Exception:
                pass


def auto_crop_image_to_bbox(input_path: str, margin_px: int = 12) -> Optional[Image.Image]:
    """精確裁切 PNG 圖像中非背景區域與異常延伸線，並附加安全像素 Margin"""
    if not is_valid_image(input_path):
        return None
    try:
        with Image.open(input_path) as img:
            img_rgba = img.convert("RGBA")
            w, h = img_rgba.size
            mask = Image.new("L", (w, h), 0)
            pixels_rgba = img_rgba.load()
            pixels_mask = mask.load()

            for y in range(h):
                for x in range(w):
                    r, g, b, a = pixels_rgba[x, y]
                    if a > 10 and (r < 240 or g < 240 or b < 240):
                        pixels_mask[x, y] = 255

            bbox = mask.getbbox()
            if not bbox:
                return img_rgba.copy()

            l, u, r, d = bbox
            nl = max(0, l - margin_px)
            nu = max(0, u - margin_px)
            nr = min(w, r + margin_px)
            nd = min(h, d + margin_px)
            return img_rgba.crop((nl, nu, nr, nd))
    except Exception:
        return None


def generate_engineering_4views_card(model: cq.Workplane, final_png_path: str) -> bool:
    """
    實作共同比例 (Common Scale Factor) 與第三角法對齊 (Projection Alignment) 的工程圖卡合成引擎。
    """
    bbox_3d = model.val().BoundingBox()
    x_len = max(0.1, bbox_3d.xlen)
    y_len = max(0.1, bbox_3d.ylen)
    z_len = max(0.1, bbox_3d.zlen)

    max_span = max(x_len, y_len, z_len)
    cell_draw_area = 580.0
    scale_ratio = cell_draw_area / max_span
    
    top_w = max(50, int(x_len * scale_ratio))
    top_h = max(50, int(y_len * scale_ratio))
    
    front_w = max(50, int(x_len * scale_ratio))
    front_h = max(50, int(z_len * scale_ratio))
    
    right_w = max(50, int(y_len * scale_ratio))
    right_h = max(50, int(z_len * scale_ratio))

    temp_dir = tempfile.gettempdir()
    uid = str(uuid.uuid4())[:8]
    p_top = os.path.join(temp_dir, f"top_{uid}.png")
    p_front = os.path.join(temp_dir, f"front_{uid}.png")
    p_right = os.path.join(temp_dir, f"right_{uid}.png")
    p_iso = os.path.join(temp_dir, f"iso_{uid}.png")

    try:
        ok_t = render_single_view(model, (0, 0, 1), p_top, top_w + 120, top_h + 120)
        ok_f = render_single_view(model, (0, -1, 0), p_front, front_w + 120, front_h + 120)
        ok_r = render_single_view(model, (1, 0, 0), p_right, right_w + 120, right_h + 120)
        ok_i = render_single_view(model, (1, 1, 1), p_iso, 800, 800)

        if not (ok_t and ok_f and ok_r and ok_i):
            return False

        img_top = auto_crop_image_to_bbox(p_top, margin_px=8)
        img_front = auto_crop_image_to_bbox(p_front, margin_px=8)
        img_right = auto_crop_image_to_bbox(p_right, margin_px=8)
        img_iso = auto_crop_image_to_bbox(p_iso, margin_px=14)

        if not (img_top and img_front and img_right and img_iso):
            return False

        canvas_w, canvas_h = 1600, 1200
        canvas = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 255))

        cx_left, cx_right = 420, 1180
        cy_top, cy_bottom = 300, 900

        fw, fh = img_front.size
        front_x = cx_left - (fw // 2)
        front_y = cy_bottom - (fh // 2)
        canvas.paste(img_front, (front_x, front_y), img_front)

        tw, th = img_top.size
        top_x = front_x + (fw // 2) - (tw // 2)
        top_y = cy_top - (th // 2)
        canvas.paste(img_top, (top_x, top_y), img_top)

        rw, rh = img_right.size
        right_x = cx_right - (rw // 2)
        right_y = front_y + (fh // 2) - (rh // 2)
        canvas.paste(img_right, (right_x, right_y), img_right)

        iw, ih = img_iso.size
        iso_x = cx_right - (iw // 2)
        iso_y = cy_top - (ih // 2)
        canvas.paste(img_iso, (iso_x, iso_y), img_iso)

        draw = ImageDraw.Draw(canvas)
        font = load_label_font(22)
        font_small = load_label_font(16)

        draw.text((30, 25), "PROJECTION: THIRD ANGLE", fill=(100, 100, 100, 255), font=font_small)

        draw.text((cx_left - 60, cy_top + 240), "平面圖 TOP", fill=(60, 60, 60, 255), font=font)
        draw.text((cx_left - 65, cy_bottom + 240), "正面圖 FRONT", fill=(60, 60, 60, 255), font=font)
        draw.text((cx_right - 65, cy_bottom + 240), "右視圖 RIGHT", fill=(60, 60, 60, 255), font=font)
        draw.text((cx_right - 80, cy_top + 240), "等角圖 ISOMETRIC", fill=(60, 60, 60, 255), font=font)

        draw.line([(800, 80), (800, 1120)], fill=(230, 230, 230, 255), width=2)
        draw.line([(80, 600), (1520, 600)], fill=(230, 230, 230, 255), width=2)

        canvas.convert("RGB").save(final_png_path, "PNG")
        return is_valid_image(final_png_path)

    except Exception as e_card:
        print(f"[ENGINEERING CARD ERROR] {str(e_card)}")
        return False
    finally:
        for p in [p_top, p_front, p_right, p_iso]:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass


def parse_cad_with_screenshot(file_path: str) -> Dict[str, Any]:
    """讀取 3D CAD 檔案，進行 OBB/AABB 比對並產出工程三視圖圖卡"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到檔案：{file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ('.step', '.stp', '.igs', '.iges'):
        raise ValueError(f"不支援的格式 '{ext}'")

    try:
        if ext in ('.step', '.stp'):
            model = cq.importers.importStep(file_path)
        else:
            try:
                model = cq.importers.importShape(file_path)
            except Exception:
                from OCP.IGESControl import IGESControl_Reader
                reader = IGESControl_Reader()
                reader.ReadFile(file_path)
                reader.TransferRoots()
                occ_shape = reader.Shape()
                model = cq.Workplane("XY").newObject([cq.Shape.cast(occ_shape)])

        raw_dims_obb = calculate_obb_dimensions(model)
        sorted_raw_obb = sorted(raw_dims_obb, reverse=True)
        sorted_obb = [ceil_dim(v) for v in sorted_raw_obb]
        vol_obb = sorted_obb[0] * sorted_obb[1] * sorted_obb[2]

        bbox = model.val().BoundingBox()
        sorted_raw_aabb = sorted([bbox.xlen, bbox.ylen, bbox.zlen], reverse=True)
        sorted_aabb = [ceil_dim(v) for v in sorted_raw_aabb]
        vol_aabb = sorted_aabb[0] * sorted_aabb[1] * sorted_aabb[2]

        if vol_obb <= vol_aabb:
            sorted_raw, sorted_ceiled = sorted_raw_obb, sorted_obb
            used_mode = "OBB"
        else:
            sorted_raw, sorted_ceiled = sorted_raw_aabb, sorted_aabb
            used_mode = "AABB"

        length_val, width_val = sorted_ceiled[0], sorted_ceiled[1]
        # 高度小於 1mm 時直接顯示實際量測值，不做無條件進位（避免 0.x mm 的薄件被灌成 1mm）
        height_raw = sorted_raw[2]
        height_val = round(height_raw, 2) if height_raw < 1 else sorted_ceiled[2]

        dimensions_str = f"{length_val}*{width_val}*{height_val}"
    except Exception as e:
        return {
            "status": "error",
            "file_type": "3D",
            "file_name": os.path.basename(file_path),
            "error_message": f"CAD 幾何解析失敗: {str(e)}"
        }

    img_path = None
    image_error = None
    unique_id = str(uuid.uuid4())[:8]
    temp_dir = tempfile.gettempdir()
    final_card_png = os.path.join(temp_dir, f"cad_card_{unique_id}.png")

    try:
        success = generate_engineering_4views_card(model, final_card_png)
        if success and is_valid_image(final_card_png):
            img_path = final_card_png
        else:
            raise ValueError("工程三視圖圖卡合成失敗。")
    except Exception as e_img:
        image_error = f"{type(e_img).__name__}: {str(e_img)}"
        print(f"[CAD IMAGE ERROR] {os.path.basename(file_path)}: {image_error}")
        img_path = None

    return {
        "status": "success",
        "file_type": "3D",
        "file_name": os.path.basename(file_path),
        "dimensions_str": dimensions_str,
        "length": length_val,
        "width": width_val,
        "height": height_val,
        "unit": "mm",
        "used_mode": used_mode,
        "image_path": img_path,
        "image_error": image_error
    }


_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))


def generate_excel_report(
    parsed_results: List[Dict[str, Any]],
    output_excel_path: str,
    header_info: Optional[Dict[str, Any]] = None
):
    """Excel 報表匯出：精確對齊 template.xlsm/xlsx 格式與 keep_vba 參數"""
    template_candidates = ["template.xlsm", "Template.xlsm", "template.xlsx", "Template.xlsx"]
    template_file = next(
        (os.path.join(_MODULE_DIR, tf) for tf in template_candidates if os.path.exists(os.path.join(_MODULE_DIR, tf))),
        None,
    )
    is_xlsm = template_file and template_file.lower().endswith('.xlsm')

    if template_file:
        wb = openpyxl.load_workbook(template_file, data_only=False, keep_vba=is_xlsm)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "報價單"

    kai_font_regular = Font(name="標楷體", size=11, bold=False)
    kai_font_bold = Font(name="標楷體", size=11, bold=True)

    if header_info and isinstance(header_info, dict):
        set_cell_value_safe(ws, 4, 2, f"客戶名稱 : {safe_str(header_info.get('customer'))}", font=kai_font_regular)
        set_cell_value_safe(ws, 5, 2, f"聯 絡 人 : {safe_str(header_info.get('contact'))}", font=kai_font_regular)
        set_cell_value_safe(ws, 6, 2, f"聯絡電話 : {safe_str(header_info.get('phone'))}", font=kai_font_regular)
        set_cell_value_safe(ws, 7, 2, f"傳    真 : {safe_str(header_info.get('fax'))}", font=kai_font_regular)

    set_cell_value_safe(ws, 7, 15, datetime.now().strftime("%Y/%m/%d"), font=kai_font_regular)

    start_row = 10
    max_b_len, max_p_len = 12, 16

    for idx, item in enumerate(parsed_results):
        row_num = start_row + idx
        ws.cell(row=row_num, column=1, value=idx + 1).font = kai_font_regular
        
        file_name = item.get("file_name", "")
        ws.cell(row=row_num, column=2, value=file_name).font = kai_font_regular
        max_b_len = max(max_b_len, len(str(file_name)))
        
        dims_str = item.get("dimensions_str", "")
        ws.cell(row=row_num, column=16, value=dims_str).font = kai_font_bold
        max_p_len = max(max_p_len, len(str(dims_str)))
        
        if "length" in item:
            ws.cell(row=row_num, column=17, value=item["length"]).font = kai_font_bold
        if "width" in item:
            ws.cell(row=row_num, column=18, value=item["width"]).font = kai_font_bold
        if "height" in item:
            ws.cell(row=row_num, column=19, value=item["height"]).font = kai_font_bold
        elif "thickness" in item and item["thickness"]:
            ws.cell(row=row_num, column=19, value=item["thickness"]).font = kai_font_bold
            
        ws.cell(row=row_num, column=20, value=item.get("unit", "mm")).font = kai_font_regular

        ws.cell(row=row_num, column=7, value=f"=E{row_num}*F{row_num}").font = kai_font_regular
        ws.cell(row=row_num, column=10, value=f"=H{row_num}*I{row_num}").font = kai_font_regular
        ws.cell(row=row_num, column=13, value=f"=K{row_num}*L{row_num}").font = kai_font_regular

    ws.column_dimensions['B'].width = max(ws.column_dimensions['B'].width or 0, max_b_len + 6)
    ws.column_dimensions['P'].width = max(ws.column_dimensions['P'].width or 0, max_p_len + 6)

    total_row = 101
    for r in range(10, ws.max_row + 1):
        if "合計" in str(ws.cell(row=r, column=1).value or "") or "合計" in str(ws.cell(row=r, column=2).value or ""):
            total_row = r
            break

    data_end_row = total_row - 1
    set_cell_value_safe(ws, total_row, 7, f"=SUM(G10:G{data_end_row})", font=kai_font_bold)
    set_cell_value_safe(ws, total_row, 10, f"=SUM(J10:J{data_end_row})", font=kai_font_bold)
    set_cell_value_safe(ws, total_row, 13, f"=SUM(M10:M{data_end_row})", font=kai_font_bold)
    set_cell_value_safe(ws, total_row, 15, f"=G{total_row}+J{total_row}+M{total_row}", font=kai_font_bold)

    wb.save(output_excel_path)


def generate_word_report(
    parsed_results: List[Dict[str, Any]], 
    output_word_path: str,
    header_info: Optional[Dict[str, Any]] = None
):
    """
    建立 Word 圖文報價單 (.docx)，實作 keep_with_next 原子區塊防跨頁保護 (Atomic CAD Block)。
    支援 3D 機構件與 2D Mylar 模切材料雙軌報表產出。
    """
    if not HAS_DOCX:
        raise ModuleNotFoundError("系統缺少 python-docx 套件，無法產生 Word 報表。")

    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("⚙️ CAD 與 2D 模切材料報價明細單")
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.name = '標楷體'

    today_str = datetime.now().strftime("%Y年%m月%d日")
    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sub_run = sub_p.add_run(f"報價日期：{today_str}")
    sub_run.font.size = Pt(10)
    sub_run.font.name = '標楷體'

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    if header_info and isinstance(header_info, dict):
        cust = safe_str(header_info.get("customer")) or "未填寫"
        contact = safe_str(header_info.get("contact")) or "未填寫"
        phone = safe_str(header_info.get("phone")) or "未填寫"
        fax = safe_str(header_info.get("fax")) or "未填寫"

        info_table = doc.add_table(rows=2, cols=2)
        info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        info_table.autofit = False

        cells = info_table.rows[0].cells
        cells[0].text = f"客戶名稱：{cust}"
        cells[1].text = f"聯 絡 人：{contact}"

        cells_2 = info_table.rows[1].cells
        cells_2[0].text = f"聯絡電話：{phone}"
        cells_2[1].text = f"傳    真：{fax}"

        for row in info_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = '標楷體'
                        run.font.size = Pt(11)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    for idx, item in enumerate(parsed_results):
        if item.get("status") != "success":
            continue

        file_name = item.get("file_name", "")
        dims_str = item.get("dimensions_str", "")
        ftype = item.get("file_type", "3D")
        img_path = item.get("image_path")
        img_err = item.get("image_error")

        p_item = doc.add_paragraph()
        p_item.paragraph_format.keep_with_next = True
        r_item = p_item.add_run(f"【項目 {idx+1}】 檔名：{file_name} ({'2D 模切' if ftype == '2D' else '3D 機構件'})")
        r_item.font.bold = True
        r_item.font.size = Pt(12)
        r_item.font.name = '標楷體'

        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.left_indent = Inches(0.2)
        p_desc.paragraph_format.keep_with_next = True
        
        if ftype == '2D':
            mat_type = item.get("material_type", "Mylar")
            gross_area = item.get("gross_area", "待確認")
            net_area = item.get("net_area", "待確認")
            thickness = item.get("thickness", "待確認")
            
            r_d1 = p_desc.add_run(f"• 材料類型：{mat_type} (厚度: {thickness} mm)\n")
            r_d1.font.bold = True
            r_d1.font.name = '標楷體'
            
            r_d2 = p_desc.add_run(f"• 外形尺寸 (長×寬×厚)：{dims_str} mm\n• 外形總面積 (Gross Area)：{gross_area} mm²\n• 淨材料面積 (Net Area)：{net_area}\n")
            r_d2.font.name = '標楷體'
        else:
            used_mode = item.get("used_mode", "OBB")
            length = item.get("length", 0)
            width = item.get("width", 0)
            height = item.get("height", 0)
            unit = item.get("unit", "mm")
            
            r_d1 = p_desc.add_run(f"• 採納素材尺寸 (長*寬*高)：{dims_str} {unit} ({used_mode} 模式)\n")
            r_d1.font.bold = True
            r_d1.font.name = '標楷體'
            
            r_d2 = p_desc.add_run(f"• 尺寸拆解數值：長 {length} {unit} / 寬 {width} {unit} / 高 {height} {unit}\n")
            r_d2.font.name = '標楷體'

        if is_valid_image(img_path):
            try:
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if idx < len(parsed_results) - 1:
                    p_img.paragraph_format.keep_with_next = True
                    
                run_img = p_img.add_run()
                run_img.add_picture(img_path, width=Inches(5.8))
            except Exception as e_word_img:
                p_err = doc.add_paragraph(f"   [ 無法產生視圖預覽 (Word插入例外: {str(e_word_img)}) ]")
                p_err.runs[0].font.color.rgb = RGBColor(128, 128, 128)
        else:
            err_msg = f" ({img_err})" if img_err else ""
            p_none = doc.add_paragraph(f"   [ 無法產生視圖預覽{err_msg} ]")
            p_none.runs[0].font.color.rgb = RGBColor(128, 128, 128)

        doc.add_paragraph().paragraph_format.space_after = Pt(8)

    doc.save(output_word_path)
