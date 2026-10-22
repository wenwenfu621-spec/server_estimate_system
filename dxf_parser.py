"""
dxf_parser.py - 2D DXF / Mylar 模切檔案解析與預覽模組
Version: v2.7.0_20260821
Description: 獨立的 2D 評估管道，使用 ezdxf 解析 DXF 輪廓、單位與尺寸，
             提供高品質 2D 輪廓預覽圖，並嚴格隔離 3D CAD 邏輯。
"""

import os
import math
import uuid
import tempfile
from typing import Dict, Any, List, Optional
from PIL import Image, ImageDraw

try:
    import ezdxf
    HAS_EZDXF = True
except ImportError:
    HAS_EZDXF = False


def is_valid_image(path: Optional[str]) -> bool:
    """驗證圖片是否有效"""
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


def parse_dxf_2d(file_path: str, user_thickness: Optional[float] = None, user_unit_choice: Optional[str] = None) -> Dict[str, Any]:
    """
    解析 2D DXF 檔案，取得外形尺寸 (Length x Width)、單位、預覽圖及基本幾何特徵。
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到檔案：{file_path}")

    if not HAS_EZDXF:
        return {
            "status": "error",
            "file_name": os.path.basename(file_path),
            "error_message": "系統缺少 ezdxf 套件，無法解析 2D DXF 檔案。"
        }

    try:
        doc = ezdxf.readfile(file_path)
    except Exception as e:
        return {
            "status": "error",
            "file_name": os.path.basename(file_path),
            "error_message": f"DXF 檔案結構損毀或無法讀取: {str(e)}"
        }

    # 1. 單位識別
    # $INSUNITS: 1=Inches, 4=Mm, 5=Cm
    insunits = doc.header.get("$INSUNITS", 0)
    unit_str = "mm"
    unit_warning = None

    if insunits == 1:
        unit_str = "inch"
    elif insunits == 4:
        unit_str = "mm"
    elif insunits == 5:
        unit_str = "cm"
    else:
        if user_unit_choice:
            unit_str = user_unit_choice
        else:
            unit_str = "unitless"
            unit_warning = "⚠️ 此 DXF 未定義圖面單位，預設以 mm 計算，請確認單位是否正確。"

    # 2. 幾何實體萃取 (過濾圖層與無關文字)
    msp = doc.modelspace()
    min_x, min_y = float('inf'), float('inf')
    max_x, max_y = float('-inf'), float('-inf')
    entity_count = 0

    ignored_layers = {"DIM", "DIMENSIONS", "TEXT", "BORDER", "TITLE", "CENTER", "CONSTRUCTION"}

    for entity in msp:
        layer = str(entity.dxf.layer).upper()
        if any(ig in layer for ig in ignored_layers):
            continue

        dxftype = entity.dxftype()
        try:
            if dxftype == 'LINE':
                s = entity.dxf.start
                e = entity.dxf.end
                min_x = min(min_x, s.x, e.x)
                max_x = max(max_x, s.x, e.x)
                min_y = min(min_y, s.y, e.y)
                max_y = max(max_y, s.y, e.y)
                entity_count += 1
            elif dxftype in ('CIRCLE', 'ARC'):
                center = entity.dxf.center
                radius = entity.dxf.radius
                min_x = min(min_x, center.x - radius)
                max_x = max(max_x, center.x + radius)
                min_y = min(min_y, center.y - radius)
                max_y = max(max_y, center.y + radius)
                entity_count += 1
            elif dxftype in ('LWPOLYLINE', 'POLYLINE'):
                for vertex in entity.get_points():
                    min_x = min(min_x, vertex[0])
                    max_x = max(max_x, vertex[0])
                    min_y = min(min_y, vertex[1])
                    max_y = max(max_y, vertex[1])
                entity_count += 1
        except Exception:
            pass

    if entity_count == 0 or min_x == float('inf'):
        return {
            "status": "error",
            "file_name": os.path.basename(file_path),
            "error_message": "DXF 圖面中未找到有效的可解析幾何圖元。"
        }

    raw_width = max_x - min_x
    raw_height = max_y - min_y

    # 單位換算至 mm
    scale_to_mm = 1.0
    if unit_str == "inch":
        scale_to_mm = 25.4
    elif unit_str == "cm":
        scale_to_mm = 10.0

    length_val = round(max(raw_width, raw_height) * scale_to_mm, 2)
    width_val = round(min(raw_width, raw_height) * scale_to_mm, 2)
    thickness_val = user_thickness if user_thickness and user_thickness > 0 else None

    dims_str = f"{length_val}*{width_val}"
    if thickness_val:
        dims_str += f"*{thickness_val}"

    # 3. 產出 2D 輪廓預覽圖 (Pillow 繪製線條)
    img_path = None
    try:
        canvas_size = 800
        img = Image.new("RGB", (canvas_size, canvas_size), (255, 255, 255))
        draw = ImageDraw.Draw(img)

        # 計算縮放與置中投影
        span_x = max(raw_width, 1e-6)
        span_y = max(raw_height, 1e-6)
        scale = (canvas_size * 0.8) / max(span_x, span_y)
        offset_x = (canvas_size - span_x * scale) / 2 - min_x * scale
        offset_y = (canvas_size - span_y * scale) / 2 - min_y * scale

        def transform_pt(x, y):
            return (int(x * scale + offset_x), int(canvas_size - (y * scale + offset_y)))

        for entity in msp:
            layer = str(entity.dxf.layer).upper()
            if any(ig in layer for ig in ignored_layers):
                continue
            dxftype = entity.dxftype()
            try:
                if dxftype == 'LINE':
                    p1 = transform_pt(entity.dxf.start.x, entity.dxf.start.y)
                    p2 = transform_pt(entity.dxf.end.x, entity.dxf.end.y)
                    draw.line([p1, p2], fill=(30, 41, 59), width=2)
                elif dxftype == 'CIRCLE':
                    c = transform_pt(entity.dxf.center.x, entity.dxf.center.y)
                    r = int(entity.dxf.radius * scale)
                    draw.ellipse([c[0]-r, c[1]-r, c[0]+r, c[1]+r], outline=(30, 41, 59), width=2)
                elif dxftype in ('LWPOLYLINE', 'POLYLINE'):
                    pts = [transform_pt(pt[0], pt[1]) for pt in entity.get_points()]
                    if len(pts) > 1:
                        draw.line(pts, fill=(30, 41, 59), width=2)
            except Exception:
                pass

        temp_dir = tempfile.gettempdir()
        unique_id = str(uuid.uuid4())[:8]
        preview_path = os.path.join(temp_dir, f"dxf_prev_{unique_id}.png")
        img.save(preview_path, "PNG")

        if is_valid_image(preview_path):
            img_path = preview_path
    except Exception:
        img_path = None

    return {
        "status": "success",
        "file_type": "2D",
        "file_name": os.path.basename(file_path),
        "dimensions_str": dims_str,
        "length": length_val,
        "width": width_val,
        "thickness": thickness_val,
        "unit": "mm",
        "unit_warning": unit_warning,
        "gross_area": round(length_val * width_val, 2),
        "net_area": "待確認",
        "hole_count": "待確認",
        "image_path": img_path,
        "image_error": None
    }