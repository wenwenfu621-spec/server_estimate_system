"""
cad_parser.py - 3D CAD 幾何解析與視圖產生模組
Version: v3.0.0_20261023
Description: 讀取 STEP/IGES，計算 OBB/AABB 素材尺寸、板金料厚，並產生
             第三角法四視圖圖卡與單張等角示意圖。報表輸出移到 report_export.py。

版本紀錄
--------
v3.0.0_20261023：
    1. 新增 measure_sheet_thickness()：以「成對平行平面 / 同軸圓柱面」的距離
       判斷板金料厚，並用體積 ÷ (表面積 ÷ 2) 交叉檢查，不像板金件時標示 suspicious。
    2. 四視圖產生時另存等角圖（iso_image_path），供 Excel「示意圖」欄與預覽表使用。
       解析結果新增 solid_count / is_assembly，2 個以上實體視為組立件。
    3. 四視圖改用 HLRBRep_PolyAlgo（多邊形近似隱藏線）取得投影線、直接以 PIL 繪製：
       組立件從每個視角數十秒降到 1 秒內，且不再需要 cairosvg / 系統 Cairo 函式庫。
    4. 暫存檔統一加上 TEMP_PREFIX 前綴，新增 cleanup_stale_temp_files() 清理逾時暫存。
    5. Excel / Word 輸出搬到 report_export.py；移除 2D 殘留與客戶表頭資訊。
v2.9.5_20261022：
    修正旋轉零件尺寸多 1mm（ceil_dim 扣容差）與四視圖中文標籤變方框（load_label_font）。
v2.9.4_20260904：
    高度 < 1mm 時不無條件進位，改顯示實際量測值（四捨五入到小數點後兩位）。
v2.9.3_20260830：
    cairosvg 的 import 保護改接 Exception（Windows 缺 Cairo 時拋 OSError）。
v2.9.2_20260830：
    範本路徑改以 __file__ 所在目錄為準，不受呼叫端 cwd 影響。
"""

import os
import glob
import math
import time
import uuid
import tempfile
from typing import Dict, Any, List, Optional

import numpy as np
import cadquery as cq
from PIL import Image, ImageDraw, ImageFont

# 引入 OpenCASCADE 原生模組：OBB 最小包容盒、曲面型別判斷、面與面距離
try:
    from OCP.Bnd import Bnd_OBB
    from OCP.BRepBndLib import BRepBndLib
    from OCP.BRepAdaptor import BRepAdaptor_Surface
    from OCP.BRepExtrema import BRepExtrema_DistShapeShape
    from OCP.BRepClass import BRepClass_FaceClassifier
    from OCP.TopAbs import TopAbs_IN, TopAbs_ON
    from OCP.gp import gp_Pnt as _gp_Pnt
    HAS_OCP_OBB = True
except ImportError:
    HAS_OCP_OBB = False

# 視圖投影（多邊形隱藏線）
from OCP.BRepMesh import BRepMesh_IncrementalMesh
from OCP.HLRBRep import HLRBRep_PolyAlgo, HLRBRep_PolyHLRToShape
from OCP.HLRAlgo import HLRAlgo_Projector
from OCP.gp import gp_Ax2, gp_Pnt, gp_Dir
from OCP.BRepAdaptor import BRepAdaptor_Curve
from OCP.GeomAbs import GeomAbs_Line
from OCP.TopExp import TopExp_Explorer
from OCP.TopAbs import TopAbs_EDGE
from OCP.TopoDS import TopoDS


# 本程式產生的暫存檔都以此開頭，方便 cleanup_stale_temp_files() 辨識
TEMP_PREFIX = "cadq_"


def temp_path(suffix: str) -> str:
    """在系統暫存目錄組出一個帶 TEMP_PREFIX 的唯一檔名"""
    return os.path.join(tempfile.gettempdir(), f"{TEMP_PREFIX}{uuid.uuid4().hex[:12]}{suffix}")


def cleanup_stale_temp_files(max_age_hours: float = 6.0) -> int:
    """刪除超過 max_age_hours 的本程式暫存檔（使用者直接關分頁時留下的），回傳刪除數量"""
    cutoff = time.time() - max_age_hours * 3600
    removed = 0
    for path in glob.glob(os.path.join(tempfile.gettempdir(), f"{TEMP_PREFIX}*")):
        try:
            if os.path.getmtime(path) < cutoff:
                os.remove(path)
                removed += 1
        except Exception:
            pass
    return removed


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


def _cylinder_of(face) -> Optional[tuple]:
    """回傳圓柱面的 (半徑, 軸心點 xyz, 軸方向 xyz)；非圓柱面回傳 None"""
    try:
        cyl = BRepAdaptor_Surface(face.wrapped).Cylinder()
        ax = cyl.Axis()
        loc, d = ax.Location(), ax.Direction()
        return cyl.Radius(), (loc.X(), loc.Y(), loc.Z()), (d.X(), d.Y(), d.Z())
    except Exception:
        return None


def _point_in_face(face, xyz) -> bool:
    """投影點是否落在面的範圍內（含邊界）"""
    try:
        state = BRepClass_FaceClassifier(face.wrapped, _gp_Pnt(*map(float, xyz)), 1e-4).State()
        return state in (TopAbs_IN, TopAbs_ON)
    except Exception:
        return False


def _faces_overlap(f1, c1, n1, f2, c2, gap: float, extrema_budget: list) -> bool:
    """兩個平行對向平面的投影是否重疊（中間是實體而非錯開的兩片）。
    先用便宜的「中心點投影是否落在對面面內」判斷；判斷不出來（例如面中心剛好在孔上）
    才用 BRepExtrema 算面與面最短距離，且限制呼叫次數（大面一次可達數十毫秒）。"""
    if _point_in_face(f1, c2 + n1 * gap) or _point_in_face(f2, c1 - n1 * gap):
        return True
    if extrema_budget[0] <= 0:
        return False
    extrema_budget[0] -= 1
    try:
        dist = BRepExtrema_DistShapeShape(f1.wrapped, f2.wrapped)
        return dist.IsDone() and abs(dist.Value() - gap) <= 1e-3
    except Exception:
        return False


def measure_sheet_thickness(model: cq.Workplane) -> Dict[str, Any]:
    """量測板金件料厚。

    板金件的每個平面背後幾乎都有一個「距離固定、法向量相反、中間是實體」的平面，
    兩者距離就是料厚；折彎處則是同軸圓柱面的半徑差。收集所有這類配對的距離，
    以較小那面的面積加權，取權重最大的距離當料厚。

    另用「體積 ÷ (表面積 ÷ 2)」粗估一次：板金件兩者應接近；實心塊會差很多，
    這時 suspicious=True，交給使用者在預覽表確認。
    """
    result = {"thickness": None, "suspicious": True, "note": "無法判斷料厚"}
    if not HAS_OCP_OBB:
        return result
    try:
        shape = model.val()
        faces = shape.Faces()
        face_areas = [f.Area() for f in faces]
        total_area = sum(face_areas) or 1.0
        # 只有封閉實體的體積才有意義；純殼面（Shell）時略過體積交叉檢查
        volume = sum(s.Volume() for s in shape.Solids())
    except Exception:
        return result

    planes, cylinders = [], []
    for f, area in zip(faces, face_areas):
        try:
            gtype = f.geomType()
            if gtype == "PLANE":
                planes.append((f, f.normalAt(), f.Center(), area))
            elif gtype == "CYLINDER":
                cyl = _cylinder_of(f)
                if cyl:
                    cylinders.append((cyl, area))
        except Exception:
            continue

    weights: Dict[float, float] = {}

    # 平面配對：法向量相反、對面平面在本面「後方」（中間是實體），且兩面投影確實重疊
    # 以 numpy 一次算完所有平面兩兩組合（組立件可達數千個面，Python 雙迴圈太慢）
    if len(planes) >= 2:
        normals = np.array([[n.x, n.y, n.z] for _, n, _, _ in planes])
        centers = np.array([[c.x, c.y, c.z] for _, _, c, _ in planes])
        areas = np.array([a for _, _, _, a in planes])
        dots = normals @ normals.T
        # gap[i, j] = -(c_j - c_i)·n_i
        gaps = -(centers @ normals.T).T + np.sum(centers * normals, axis=1)[:, None]
        mask = np.triu(dots < -0.999, k=1) & (gaps > 1e-4)
        ii, jj = np.nonzero(mask)
        weights_ij = np.minimum(areas[ii], areas[jj])
        order = np.argsort(-weights_ij)[:_MAX_PAIR_CHECKS]
        extrema_budget = [_MAX_EXTREMA_CHECKS]
        for k in order:
            i, j = ii[k], jj[k]
            gap = float(gaps[i, j])
            if not _faces_overlap(planes[i][0], centers[i], normals[i], planes[j][0], centers[j], gap, extrema_budget):
                continue
            key = round(gap, 2)
            weights[key] = weights.get(key, 0.0) + float(weights_ij[k])

    # 折彎處：同軸圓柱面的半徑差（同樣用 numpy 一次比對所有組合）
    if len(cylinders) >= 2:
        radii = np.array([r for (r, _, _), _ in cylinders])
        locs = np.array([p for (_, p, _), _ in cylinders])
        dirs = np.array([d for (_, _, d), _ in cylinders])
        c_areas = np.array([a for _, a in cylinders])
        parallel = np.abs(np.abs(dirs @ dirs.T) - 1.0) <= 1e-4
        # 軸心偏移量扣掉沿軸方向的分量，剩下的就是兩軸線間距
        offsets = locs[None, :, :] - locs[:, None, :]
        along = np.einsum("ijk,ik->ij", offsets, dirs)
        perp = np.linalg.norm(offsets - along[:, :, None] * dirs[:, None, :], axis=2)
        gaps = np.abs(radii[:, None] - radii[None, :])
        mask = np.triu(parallel & (perp <= 1e-3) & (gaps > 1e-4), k=1)
        for i, j in zip(*np.nonzero(mask)):
            key = round(float(gaps[i, j]), 2)
            weights[key] = weights.get(key, 0.0) + float(min(c_areas[i], c_areas[j]))

    if not weights:
        return result

    thickness, best_weight = max(weights.items(), key=lambda kv: kv[1])
    coverage = 2.0 * best_weight / total_area
    suspicious = coverage < 0.5
    if volume > 0:
        ratio = (2.0 * volume / total_area) / thickness
        suspicious = suspicious or not (0.7 <= ratio <= 1.15)

    result.update({
        "thickness": thickness,
        "suspicious": bool(suspicious),
        "note": "疑似非板金件，請確認料厚" if suspicious else "",
    })
    return result


_DIM_EPS = 1e-6
_MAX_PAIR_CHECKS = 400
_MAX_EXTREMA_CHECKS = 40

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


# 視角：(視線方向 = 指向觀察者, 圖面 X 軸方向)。第三角法，Z 軸朝上。
VIEW_AXES = {
    "top": ((0, 0, 1), (1, 0, 0)),
    "front": ((0, -1, 0), (1, 0, 0)),
    "right": ((1, 0, 0), (0, 1, 0)),
    "iso": ((1, 1, 1), (-1, 1, 0)),
}
_LINE_COLOR = (40, 40, 40, 255)
# 圖卡每個象限可放視圖的範圍（下方要留給標籤文字）
_VIEW_BOX_W, _VIEW_BOX_H = 680.0, 470.0
_SUPERSAMPLE = 2


def _mesh_shape(model: cq.Workplane):
    """多邊形隱藏線演算法需要先把實體三角化；誤差取最大邊長的 0.2%"""
    shape = model.val().wrapped
    bbox = model.val().BoundingBox()
    span = max(bbox.xlen, bbox.ylen, bbox.zlen, 1e-3)
    BRepMesh_IncrementalMesh(shape, span * 0.002, False, 0.35, True)
    return shape


def project_view_lines(shape, view: str) -> List[np.ndarray]:
    """以 HLRBRep_PolyAlgo（多邊形近似隱藏線）取得指定視角的可見線，回傳 2D 折線清單。

    比 CadQuery SVG 匯出用的精確 HLR（HLRBRep_Algo）快一到兩個數量級，
    組立件從數十秒降到 1 秒內；且直接畫成 PNG，不需要 Cairo。
    """
    direction, x_dir = VIEW_AXES[view]
    projector = HLRAlgo_Projector(gp_Ax2(gp_Pnt(0, 0, 0), gp_Dir(*direction), gp_Dir(*x_dir)))
    algo = HLRBRep_PolyAlgo()
    algo.Load(shape)
    algo.Projector(projector)
    algo.Update()
    to_shape = HLRBRep_PolyHLRToShape()
    to_shape.Update(algo)

    lines = []
    for compound in (to_shape.VCompound(), to_shape.OutLineVCompound(), to_shape.Rg1LineVCompound()):
        if compound.IsNull():
            continue
        explorer = TopExp_Explorer(compound, TopAbs_EDGE)
        while explorer.More():
            curve = BRepAdaptor_Curve(TopoDS.Edge_s(explorer.Current()))
            first, last = curve.FirstParameter(), curve.LastParameter()
            if curve.GetType() == GeomAbs_Line:
                params = (first, last)
            else:
                params = np.linspace(first, last, 24)
            pts = [curve.Value(p) for p in params]
            lines.append(np.array([(p.X(), p.Y()) for p in pts]))
            explorer.Next()
    return lines


def _lines_extent(lines: List[np.ndarray]):
    if not lines:
        return 0.0, 0.0, 1.0, 1.0
    allpts = np.vstack(lines)
    return allpts[:, 0].min(), allpts[:, 1].min(), allpts[:, 0].max(), allpts[:, 1].max()


def draw_view_image(lines: List[np.ndarray], px_per_mm: float, margin_px: int = 10) -> Image.Image:
    """把投影折線畫成透明背景的 RGBA 圖（先放大 2 倍畫再縮小，線條較平滑）"""
    min_x, min_y, max_x, max_y = _lines_extent(lines)
    s = px_per_mm * _SUPERSAMPLE
    m = margin_px * _SUPERSAMPLE
    w = max(1, int(math.ceil((max_x - min_x) * s)) + 2 * m)
    h = max(1, int(math.ceil((max_y - min_y) * s)) + 2 * m)
    img = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    width = max(1, round(1.2 * _SUPERSAMPLE))
    for line in lines:
        xy = [((x - min_x) * s + m, (max_y - y) * s + m) for x, y in line]
        if len(xy) >= 2:
            draw.line(xy, fill=_LINE_COLOR, width=width, joint="curve")
    return img.resize((max(1, w // _SUPERSAMPLE), max(1, h // _SUPERSAMPLE)), Image.LANCZOS)


def _fit_scale(lines: List[np.ndarray], box_w: float, box_h: float) -> float:
    min_x, min_y, max_x, max_y = _lines_extent(lines)
    return min(box_w / max(max_x - min_x, 1e-3), box_h / max(max_y - min_y, 1e-3))


def _save_on_white(img_rgba: Image.Image, out_path: str) -> bool:
    """把透明背景的 RGBA 圖貼到白底後存成 PNG"""
    bg = Image.new("RGB", img_rgba.size, (255, 255, 255))
    bg.paste(img_rgba, (0, 0), img_rgba)
    bg.save(out_path, "PNG")
    return is_valid_image(out_path)


def generate_engineering_4views_card(
    model: cq.Workplane,
    final_png_path: str,
    iso_png_path: Optional[str] = None
) -> Dict[str, bool]:
    """
    實作共同比例 (Common Scale Factor) 與第三角法對齊 (Projection Alignment) 的工程圖卡合成引擎。
    若指定 iso_png_path，另把等角圖單獨存成一張（不受其他視圖成敗影響）。
    回傳 {"card": 圖卡是否成功, "iso": 等角圖是否成功}。
    """
    status = {"card": False, "iso": False}
    try:
        shape = _mesh_shape(model)
        lines = {view: project_view_lines(shape, view) for view in VIEW_AXES}
    except Exception as e_proj:
        print(f"[VIEW PROJECTION ERROR] {e_proj}")
        return status

    if not lines["iso"]:
        return status
    img_iso = draw_view_image(lines["iso"], _fit_scale(lines["iso"], _VIEW_BOX_W, _VIEW_BOX_H), margin_px=14)
    if iso_png_path:
        status["iso"] = _save_on_white(img_iso, iso_png_path)

    try:
        # 三個正投影共用同一比例，維持彼此尺寸關係；比例取「每個視圖都放得進象限」的最大值
        # 平面圖 寬 X 高 Y、正面圖 寬 X 高 Z、右視圖 寬 Y 高 Z
        bbox_3d = model.val().BoundingBox()
        x_len, y_len, z_len = (max(v, 0.1) for v in (bbox_3d.xlen, bbox_3d.ylen, bbox_3d.zlen))
        scale_ratio = min(_VIEW_BOX_W / max(x_len, y_len), _VIEW_BOX_H / max(y_len, z_len))
        img_top = draw_view_image(lines["top"], scale_ratio, margin_px=8)
        img_front = draw_view_image(lines["front"], scale_ratio, margin_px=8)
        img_right = draw_view_image(lines["right"], scale_ratio, margin_px=8)

        canvas_w, canvas_h = 1600, 1200
        canvas = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 255))

        cx_left, cx_right = 420, 1180
        cy_top, cy_bottom = 305, 885

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

        draw.text((cx_left - 60, cy_top + 248), "平面圖 TOP", fill=(60, 60, 60, 255), font=font)
        draw.text((cx_left - 65, cy_bottom + 248), "正面圖 FRONT", fill=(60, 60, 60, 255), font=font)
        draw.text((cx_right - 65, cy_bottom + 248), "右視圖 RIGHT", fill=(60, 60, 60, 255), font=font)
        draw.text((cx_right - 80, cy_top + 248), "等角圖 ISOMETRIC", fill=(60, 60, 60, 255), font=font)

        draw.line([(800, 80), (800, 1120)], fill=(230, 230, 230, 255), width=2)
        draw.line([(80, 600), (1520, 600)], fill=(230, 230, 230, 255), width=2)

        canvas.convert("RGB").save(final_png_path, "PNG")
        status["card"] = is_valid_image(final_png_path)
        return status

    except Exception as e_card:
        print(f"[ENGINEERING CARD ERROR] {str(e_card)}")
        return status


def load_cad_model(file_path: str) -> cq.Workplane:
    """讀入 STEP / IGES 檔案成 CadQuery Workplane"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.step', '.stp'):
        return cq.importers.importStep(file_path)
    try:
        return cq.importers.importShape(file_path)
    except Exception:
        from OCP.IGESControl import IGESControl_Reader
        reader = IGESControl_Reader()
        reader.ReadFile(file_path)
        reader.TransferRoots()
        occ_shape = reader.Shape()
        return cq.Workplane("XY").newObject([cq.Shape.cast(occ_shape)])


def parse_cad_with_screenshot(file_path: str) -> Dict[str, Any]:
    """讀取 3D CAD 檔案：OBB/AABB 素材尺寸、板金料厚、四視圖圖卡與等角示意圖"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到檔案：{file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ('.step', '.stp', '.igs', '.iges'):
        raise ValueError(f"不支援的格式 '{ext}'")

    try:
        model = load_cad_model(file_path)

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
        thickness_info = measure_sheet_thickness(model)
        # 2 個以上實體 = 組立件（例如板金 + 壓鉚螺帽），尺寸與料厚是以整體計算
        solid_count = len(model.val().Solids())
    except Exception as e:
        return {
            "status": "error",
            "file_type": "3D",
            "file_name": os.path.basename(file_path),
            "error_message": f"CAD 幾何解析失敗: {str(e)}"
        }

    card_png = temp_path(".png")
    iso_png = temp_path(".png")
    image_error = None
    try:
        views = generate_engineering_4views_card(model, card_png, iso_png)
        if not views["card"]:
            image_error = "工程四視圖圖卡合成失敗"
    except Exception as e_img:
        views = {"card": False, "iso": False}
        image_error = f"{type(e_img).__name__}: {str(e_img)}"
    if image_error:
        print(f"[CAD IMAGE ERROR] {os.path.basename(file_path)}: {image_error}")

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
        "thickness": thickness_info["thickness"],
        "thickness_suspicious": thickness_info["suspicious"],
        "thickness_note": thickness_info["note"],
        "solid_count": solid_count,
        "is_assembly": solid_count > 1,
        "image_path": card_png if views["card"] else None,
        "iso_image_path": iso_png if views["iso"] else None,
        "image_error": image_error
    }
