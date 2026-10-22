"""
pricing.py - 材質牌價抓取模組
Version: v3.0.0_20261023
Description: 讀取 material_prices.json 基準表，並嘗試從公開網站抓最新行情：

    SECC / SGCC：中鋼新聞中心「月盤盤價」公告（中鋼只公布調幅，不公布絕對價）
    SUS          ：經濟部基本金屬供需監控平台「冷軋不銹鋼板捲 304 亞洲價」（美元/公噸）
    AL           ：Westmetall 轉載之 LME 鋁現貨結算價（美元/公噸）
    匯率         ：open.er-api.com 美元兌新台幣

任何一個來源抓不到都不影響報價：回傳 None，由呼叫端退回基準表價格。
"""

import os
import re
import json
from datetime import date
from typing import Dict, Any, Optional

import requests

_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
PRICE_TABLE_PATH = os.path.join(_MODULE_DIR, "material_prices.json")

MATERIALS = ["SECC", "SGCC", "SUS", "AL", "Other"]

_HEADERS = {"User-Agent": "Mozilla/5.0 (CAD-Quote-System)"}
_TIMEOUT = 10

CSC_NEWS_URL = "https://www.csc.com.tw/CS/CSC/NC/nc.aspx"
CSC_CONTENT_BASE = "https://www.csc.com.tw/CS/CSC/NC/neli/"
SUS_URL = "https://chart.metaltrade.tw/sts/ndomestic/2/"
AL_URL = "https://www.westmetall.com/en/markdaten.php?action=table&field=LME_Al_cash"
FX_URL = "https://open.er-api.com/v6/latest/USD"


def load_price_table() -> Dict[str, Any]:
    """讀取 material_prices.json；檔案缺漏時給一份最小可用的預設值"""
    try:
        with open(PRICE_TABLE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {
            "fx_usd_twd": 32.0,
            "csc_basis": "",
            "materials": {m: {"density": 7.85, "price": 0.0, "premium": 0.0, "label": m} for m in MATERIALS},
        }


def _get_text(url: str) -> str:
    resp = requests.get(url, headers=_HEADERS, timeout=_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def _plain(html: str) -> str:
    html = re.sub(r"<(script|style)\b.*?</\1>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))


def _number(text: str) -> float:
    return float(text.replace(",", ""))


def fetch_fx_usd_twd() -> Optional[float]:
    try:
        resp = requests.get(FX_URL, headers=_HEADERS, timeout=_TIMEOUT)
        resp.raise_for_status()
        rate = resp.json().get("rates", {}).get("TWD")
        return float(rate) if rate else None
    except Exception:
        return None


def fetch_csc_adjustment() -> Optional[Dict[str, Any]]:
    """抓中鋼最新一期「盤價」公告，回傳標題、日期與電鍍鋅 / 熱浸鍍鋅的調幅（元/公噸）"""
    try:
        listing = _get_text(CSC_NEWS_URL)
        link = None
        for m in re.finditer(r'href=["\']([^"\']*neli_content\.aspx\?newid=\w+)["\'](.*?)</a>', listing, re.S):
            # 標題可能在連結內或緊接在後面，往後多看一段文字
            tail = _plain(m.group(2) + listing[m.end():m.end() + 800])
            if "盤價" in tail:
                link = m.group(1)
                break
        if not link:
            return None
        url = link if link.startswith("http") else CSC_CONTENT_BASE + link.split("/")[-1]
        text = _plain(_get_text(url))

        title_m = re.search(r"(中鋼[^ ]{0,20}盤價)", text)
        date_m = re.search(r"發佈日期[:：]\s*(\d{4}/\d{1,2}/\d{1,2})", text)

        def adj(keyword: str) -> Optional[float]:
            m = re.search(keyword + r"[^+\-－\d]{0,30}([+\-－]?\s?\d[\d,]*)", text)
            if m:
                return _number(m.group(1).replace("－", "-").replace(" ", ""))
            if re.search(keyword + r"[^。]{0,30}(持平|不變)", text):
                return 0.0
            return None

        return {
            "title": title_m.group(1) if title_m else "中鋼盤價公告",
            "date": date_m.group(1) if date_m else "",
            "url": url,
            "SECC": adj("電鍍鋅鋼捲"),
            "SGCC": adj("熱浸鍍鋅鋼捲"),
        }
    except Exception:
        return None


def fetch_sus304_usd() -> Optional[Dict[str, Any]]:
    """冷軋不銹鋼板捲 304 亞洲月均價（美元/公噸）"""
    try:
        text = _plain(_get_text(SUS_URL))
        m = re.search(r"(\d{4})年(\d{1,2})月平均價格為\s*([\d,\.]+)\s*美元/公噸", text)
        if not m:
            return None
        return {"usd_per_ton": _number(m.group(3)), "date": f"{m.group(1)}/{int(m.group(2)):02d}"}
    except Exception:
        return None


def fetch_lme_aluminum_usd() -> Optional[Dict[str, Any]]:
    """LME 鋁現貨結算價（美元/公噸），取表格第一列（最新交易日）"""
    try:
        html = _get_text(AL_URL)
        for row in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
            cells = [c.strip() for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", row, re.S)]
            cells = [_plain(c).strip() for c in cells]
            if len(cells) >= 2 and re.match(r"\d{1,2}\. \w+ \d{4}", cells[0]) and re.match(r"[\d,\.]+$", cells[1]):
                return {"usd_per_ton": _number(cells[1]), "date": cells[0]}
        return None
    except Exception:
        return None


def fetch_market_prices() -> Dict[str, Any]:
    """組合所有線上行情，並以基準表為底算出各材質的建議單價（元/kg）。

    回傳：
        {
          "prices": {材質: {"price": 元/kg, "density": g/cm³, "source": 說明文字}},
          "csc": 中鋼公告 dict 或 None,
          "fx": 使用的匯率, "fx_source": 說明,
          "fetched": 今天日期,
        }
    """
    table = load_price_table()
    base = table.get("materials", {})
    fx = fetch_fx_usd_twd()
    fx_source = "open.er-api.com" if fx else "基準表備用匯率"
    fx = fx or float(table.get("fx_usd_twd", 32.0))

    prices: Dict[str, Dict[str, Any]] = {}
    for mat in MATERIALS:
        info = base.get(mat, {})
        prices[mat] = {
            "price": float(info.get("price", 0.0)),
            "density": float(info.get("density", 7.85)),
            "source": f"基準表（{table.get('updated', '')}）",
        }

    csc = fetch_csc_adjustment()
    if csc:
        for mat in ("SECC", "SGCC"):
            delta = csc.get(mat)
            if delta is None:
                continue
            same_basis = csc.get("title") == table.get("csc_basis")
            sign = "+" if delta >= 0 else ""
            note = "（已含於基準價）" if same_basis else "（尚未套用至基準價）"
            prices[mat]["source"] += f"；{csc['title']} {sign}{delta:,.0f} 元/噸{note}"

    def from_usd(mat: str, data: Optional[Dict[str, Any]], label: str):
        if not data:
            prices[mat]["source"] += f"；{label}抓取失敗，使用基準價"
            return
        premium = float(base.get(mat, {}).get("premium", 0.0))
        twd_per_kg = data["usd_per_ton"] * fx / 1000.0 + premium
        prices[mat]["price"] = round(twd_per_kg, 2)
        extra = f" + 溢價 {premium:g}" if premium else ""
        prices[mat]["source"] = (
            f"{label} {data['date']}：US${data['usd_per_ton']:,.1f}/噸 × 匯率 {fx:.2f}{extra}"
        )

    from_usd("SUS", fetch_sus304_usd(), "冷軋 304 亞洲價（經濟部金屬供需監控平台 / Fastmarkets）")
    from_usd("AL", fetch_lme_aluminum_usd(), "LME 鋁現貨結算價（Westmetall）")

    return {
        "prices": prices,
        "csc": csc,
        "fx": fx,
        "fx_source": fx_source,
        "fetched": date.today().isoformat(),
    }
