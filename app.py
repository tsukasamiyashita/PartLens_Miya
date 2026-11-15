# -*- coding: utf-8 -*-
"""
PartLens_Miya-v1.0.0
Gemini API マルチモーダル解析 & ローカルエンジン搭載
2D図面 (PDF / CAD DXF) から部品表 (BOM) を作成・編集し Excel (.xlsx) へ出力するデスクトップアプリ
Windows 11 ライトモード完全固定対応
"""

import os
import io
import sys
import re
import csv
import json
import math
import copy
import uuid
import base64
import ctypes
import tempfile
import datetime
import threading
import traceback
import urllib.request
import urllib.error
import urllib.parse
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Tuple, Optional, Any

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
import ezdxf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage

try:
    from google import genai
    from google.genai import types as genai_types
    HAS_GOOGLE_GENAI = True
except Exception:
    HAS_GOOGLE_GENAI = False


# ==============================================================================
# 定数・ライトモード固定パレット・Windows API設定
# ==============================================================================

APP_NAME = "PartLens_Miya-v1.0.0"
APP_TITLE = "PartLens_Miya-v1.0.0 - AI搭載 2D図面(PDF/CAD) 部品表(BOM)作成システム"
CONFIG_FILE_PATH = os.path.join(os.path.expanduser("~"), ".partlens_miya_config.json")


class LightPalette:
    BG_MAIN = "#F1F5F9"
    BG_PANEL = "#FFFFFF"
    BG_SUB = "#F8FAFC"
    BG_HEADER = "#1E3A8A"
    BG_CANVAS = "#E2E8F0"
    BG_ROW_EVEN = "#FFFFFF"
    BG_ROW_ODD = "#F8FAFC"
    BG_SELECTED = "#DBEAFE"
    FG_MAIN = "#0F172A"
    FG_SUB = "#475569"
    FG_MUTED = "#64748B"
    FG_HEADER = "#FFFFFF"
    FG_ACCENT = "#1D4ED8"
    BTN_PRIMARY_BG = "#2563EB"
    BTN_PRIMARY_FG = "#FFFFFF"
    BTN_AI_BG = "#7C3AED"
    BTN_AI_FG = "#FFFFFF"
    BTN_SUCCESS_BG = "#059669"
    BTN_SUCCESS_FG = "#FFFFFF"
    BTN_WARN_BG = "#D97706"
    BTN_WARN_FG = "#FFFFFF"
    BTN_DEFAULT_BG = "#FFFFFF"
    BTN_DEFAULT_FG = "#1E293B"
    BORDER = "#CBD5E1"
    HIGHLIGHT_BOX = "#F59E0B"
    SELECT_BOX = "#EF4444"


def get_resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def enforce_windows_light_mode(window: tk.Misc) -> None:
    if sys.platform == "win32":
        try:
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()
            val = ctypes.c_int(0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(val), ctypes.sizeof(val))
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 19, ctypes.byref(val), ctypes.sizeof(val))
        except Exception:
            pass


def enable_high_dpi_awareness() -> None:
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


# ==============================================================================
# JIS規格材質マスター・表面処理・標準部品辞書
# ==============================================================================

MATERIAL_MASTER: Dict[str, Dict[str, Any]] = {
    "SS400": {"density": 7.85, "treat": "四三酸化鉄皮膜"},
    "S45C": {"density": 7.85, "treat": "四三酸化鉄皮膜"},
    "S50C": {"density": 7.85, "treat": "四三酸化鉄皮膜"},
    "SCM435": {"density": 7.85, "treat": "調質・四三酸化鉄皮膜"},
    "SCM440": {"density": 7.85, "treat": "調質・四三酸化鉄皮膜"},
    "SKD11": {"density": 7.85, "treat": "焼入れ焼戻し"},
    "SKH51": {"density": 8.15, "treat": "焼入れ焼戻し"},
    "SPCC": {"density": 7.85, "treat": "ユニクロめっき"},
    "SPHC": {"density": 7.85, "treat": "塗装"},
    "SECC": {"density": 7.85, "treat": "生地"},
    "SGP": {"density": 7.85, "treat": "塗装"},
    "STKM13A": {"density": 7.85, "treat": "四三酸化鉄皮膜"},
    "STKR400": {"density": 7.85, "treat": "塗装"},
    "FC200": {"density": 7.20, "treat": "塗装"},
    "FCD450": {"density": 7.10, "treat": "塗装"},
    "SUJ2": {"density": 7.83, "treat": "焼入れ研磨"},
    "SUS304": {"density": 7.93, "treat": "生地(酸洗)"},
    "SUS303": {"density": 7.93, "treat": "生地"},
    "SUS316": {"density": 7.98, "treat": "生地"},
    "SUS430": {"density": 7.70, "treat": "生地"},
    "SUS440C": {"density": 7.75, "treat": "焼入れ"},
    "A5052": {"density": 2.68, "treat": "白アルマイト"},
    "A5056": {"density": 2.64, "treat": "白アルマイト"},
    "A6061": {"density": 2.70, "treat": "白アルマイト"},
    "A6063": {"density": 2.70, "treat": "アルマイト"},
    "A2017": {"density": 2.79, "treat": "白アルマイト"},
    "A7075": {"density": 2.80, "treat": "硬質アルマイト"},
    "ADC12": {"density": 2.70, "treat": "生地"},
    "C3604": {"density": 8.50, "treat": "生地"},
    "C2801": {"density": 8.40, "treat": "生地"},
    "C1100": {"density": 8.89, "treat": "生地"},
    "CAC406": {"density": 8.80, "treat": "生地"},
    "POM": {"density": 1.41, "treat": "生地"},
    "MCナイロン": {"density": 1.16, "treat": "生地"},
    "PEEK": {"density": 1.30, "treat": "生地"},
    "PTFE": {"density": 2.16, "treat": "生地"},
    "UHMW-PE": {"density": 0.94, "treat": "生地"},
    "PC": {"density": 1.20, "treat": "透明"},
    "ABS": {"density": 1.05, "treat": "生地"},
    "PMMA": {"density": 1.19, "treat": "透明"},
    "PVC": {"density": 1.40, "treat": "生地"},
    "ウレタン": {"density": 1.20, "treat": "生地"},
    "NBR": {"density": 1.20, "treat": "生地"},
}

SURFACE_TREATMENTS: List[str] = [
    "生地", "四三酸化鉄皮膜(黒染め)", "無電解ニッケルめっき", "硬質クロムめっき",
    "三価クロメートめっき", "ユニクロめっき", "白アルマイト", "黒アルマイト",
    "硬質アルマイト", "焼入れ研磨", "高周波焼入れ", "浸炭焼入れ", "調質(QT)",
    "メラミン焼付塗装", "粉体塗装", "ウレタン塗装", "酸洗・不動態化処理", "電解研磨",
]

CATEGORIES: List[str] = ["加工品", "購入品", "標準部品", "製缶・溶接品", "板金品", "支給品", "組立品"]
UNITS: List[str] = ["個", "本", "枚", "式", "セット", "組", "m", "kg"]

STANDARD_PART_KEYWORDS: List[str] = [
    "ボルト", "キャップボルト", "ナット", "ワッシャ", "座金", "ねじ", "ピン", "平行キー",
    "止め輪", "ベアリング", "軸受", "カムフォロア", "リニアガイド", "ボールねじ", "リニアブッシュ",
    "Oリング", "オイルシール", "シリンダ", "モータ", "減速機", "ベルト", "プーリ", "スプロケット",
    "チェーン", "ギヤ", "カップリング", "センサ", "スイッチ", "キャスター", "アジャスタ",
]

# 2026年11月時点 Gemini API 公式モデル体系に基づく最適モデル・カタログ定義
DEFAULT_OPTIMAL_GEMINI_MODEL: str = "gemini-3.8-flash"

DEPRECATED_GEMINI_MODEL_REPLACEMENTS: Dict[str, str] = {
    "gemini-2.0-flash": "gemini-3.8-flash",
    "gemini-2.0-flash-lite": "gemini-3.5-flash-lite",
    "gemini-2.0-pro-exp": "gemini-3.8-flash",
    "gemini-2.5-pro": "gemini-3.8-flash",
    "gemini-3.7-flash": "gemini-3.8-flash",
}

GEMINI_MODEL_CATALOG: List[Dict[str, Any]] = [
    {
        "id": "gemini-3.8-flash",
        "display_name": "Gemini 3.8 Flash",
        "generation": "Gemini 3.8",
        "tier_badge": "★図面解析 最推奨",
        "free_tier": "無料枠あり",
        "speed": "高速",
        "accuracy": "非常に高い",
        "pricing": "無料枠 / 低単価 ($0.75~$1.50)",
        "supports_media_resolution": True,
        "is_gemini4": False,
        "recommended_rank": 1,
        "recommended": True,
        "source": "標準カタログ",
        "description": (
            "2D図面・部品表(BOM)解析における総合最適モデル。高解像度画像読取(MEDIA_RESOLUTION_HIGH)と"
            "位置座標(box_2d)・JSON構造化出力に優れ、細かな寸法値や表題欄の読み取り精度と応答速度・コストを高い次元で両立します。"
        ),
    },
    {
        "id": "gemini-4-hydrogen",
        "display_name": "Gemini 4 Hydrogen",
        "generation": "Gemini 4",
        "tier_badge": "Gemini 4 標準・推奨",
        "free_tier": "無料枠あり",
        "speed": "高速",
        "accuracy": "非常に高い",
        "pricing": "無料枠 / 標準 ($1.50/$7.50)",
        "supports_media_resolution": True,
        "is_gemini4": True,
        "recommended_rank": 2,
        "recommended": True,
        "source": "標準カタログ",
        "description": (
            "最新Gemini 4世代の汎用バランスモデル（無料枠から利用可能）。"
            "複雑な加工品仕様の解釈、JIS材質正規化、購入品メーカー推論を伴うBOM作成に強みを発揮します。"
        ),
    },
    {
        "id": "gemini-4-argon",
        "display_name": "Gemini 4 Argon",
        "generation": "Gemini 4",
        "tier_badge": "Gemini 4 高度推論",
        "free_tier": "有料プラン",
        "speed": "標準",
        "accuracy": "最高クラス",
        "pricing": "有料 ($2.00/$10.00)",
        "supports_media_resolution": True,
        "is_gemini4": True,
        "recommended_rank": 3,
        "recommended": True,
        "source": "標準カタログ",
        "description": (
            "部品点数が非常に多い大規模組立図や、多数ページの複雑な図面セットの統合推論に適した高効率フロンティアモデル。"
            "AI Pro / AI Ultra 等の上位APIプラン環境で高い解析精度を発揮します。"
        ),
    },
    {
        "id": "gemini-4-carbon",
        "display_name": "Gemini 4 Carbon",
        "generation": "Gemini 4",
        "tier_badge": "Gemini 4 最高精度",
        "free_tier": "有料(Ultra等)",
        "speed": "じっくり推論",
        "accuracy": "最高峰",
        "pricing": "有料 ($2.00~$4.00)",
        "supports_media_resolution": True,
        "is_gemini4": True,
        "recommended_rank": 4,
        "recommended": False,
        "source": "標準カタログ",
        "description": (
            "Google最上位の深層推論フロンティアモデル。文字が潰れた古いスキャン紙図面や手書き注記が多い難読図面を"
            "徹底精査したい場合に最適です（対応有料プランが必要です）。"
        ),
    },
    {
        "id": "gemini-3.6-flash",
        "display_name": "Gemini 3.6 Flash",
        "generation": "Gemini 3.6",
        "tier_badge": "安定版 (GA)",
        "free_tier": "無料枠/有料",
        "speed": "高速",
        "accuracy": "高い",
        "pricing": "低単価 ($0.75/$3.75)",
        "supports_media_resolution": True,
        "is_gemini4": False,
        "recommended_rank": 5,
        "recommended": False,
        "source": "標準カタログ",
        "description": (
            "一般提供(GA)済みの安定版Flashモデル。提供終了予定のない安定した運用環境や、"
            "旧Gemini 2.0系からの確実な置き換え先として安心してお使いいただけます。"
        ),
    },
    {
        "id": "gemini-3.5-flash-lite",
        "display_name": "Gemini 3.5 Flash-Lite",
        "generation": "Gemini 3.5",
        "tier_badge": "超高速・省コスト (GA)",
        "free_tier": "無料枠/最安",
        "speed": "超高速",
        "accuracy": "標準〜高い",
        "pricing": "最安帯 ($0.30/$2.50)",
        "supports_media_resolution": True,
        "is_gemini4": False,
        "recommended_rank": 6,
        "recommended": False,
        "source": "標準カタログ",
        "description": (
            "超高速・低コストの軽量GAモデル。数十枚の図面フォルダを一括高速スキャンする場合や、"
            "文字がはっきりした単品図の連続処理で待機時間とAPIコストを抑えたい場合に最適です。"
        ),
    },
    {
        "id": "gemini-3.1-pro-preview",
        "display_name": "Gemini 3.1 Pro (Preview)",
        "generation": "Gemini 3.1",
        "tier_badge": "Pro プレビュー",
        "free_tier": "有料/制限",
        "speed": "標準",
        "accuracy": "非常に高い",
        "pricing": "Pro料金帯",
        "supports_media_resolution": True,
        "is_gemini4": False,
        "recommended_rank": 7,
        "recommended": False,
        "source": "標準カタログ",
        "description": (
            "Gemini 3.1世代のPro系プレビューモデル。図面内の長文注記や部品関連性の精査・比較検証向けです。"
        ),
    },
    {
        "id": "gemini-2.5-flash",
        "display_name": "Gemini 2.5 Flash (旧世代)",
        "generation": "Gemini 2.5",
        "tier_badge": "旧世代・互換用",
        "free_tier": "制限あり",
        "speed": "高速",
        "accuracy": "標準",
        "pricing": "旧世代料金",
        "supports_media_resolution": False,
        "is_gemini4": False,
        "recommended_rank": 80,
        "recommended": False,
        "source": "標準カタログ",
        "description": (
            "旧2.5世代の互換維持用モデル（段階的移行対象）。新規利用ではより高精度・最新対応の "
            "gemini-3.8-flash または gemini-4-hydrogen の利用を推奨します。"
        ),
    },
]

GEMINI_MODELS: List[str] = [m["id"] for m in GEMINI_MODEL_CATALOG]


def strip_gemini_model_prefix(model_id: str) -> str:
    raw = str(model_id or "").strip()
    if raw.startswith("models/"):
        raw = raw[len("models/"):]
    return raw.strip()


def migrate_deprecated_model_id(model_id: str) -> str:
    clean = strip_gemini_model_prefix(model_id)
    if not clean:
        return DEFAULT_OPTIMAL_GEMINI_MODEL
    if clean in DEPRECATED_GEMINI_MODEL_REPLACEMENTS:
        return DEPRECATED_GEMINI_MODEL_REPLACEMENTS[clean]
    return clean


def get_gemini_model_catalog_map() -> Dict[str, Dict[str, Any]]:
    return {m["id"]: copy.deepcopy(m) for m in GEMINI_MODEL_CATALOG}


def get_gemini_model_entry(
    model_id: str,
    catalog_entries: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    clean = strip_gemini_model_prefix(model_id)
    source_list = catalog_entries if catalog_entries else GEMINI_MODEL_CATALOG
    for m in source_list:
        if m.get("id") == clean:
            return copy.deepcopy(m)
    for m in GEMINI_MODEL_CATALOG:
        if m.get("id") == clean:
            return copy.deepcopy(m)
    is_g4 = clean.startswith("gemini-4")
    supports_mr = is_g4 or clean.startswith("gemini-3")
    return {
        "id": clean or DEFAULT_OPTIMAL_GEMINI_MODEL,
        "display_name": clean or DEFAULT_OPTIMAL_GEMINI_MODEL,
        "generation": "Gemini 4" if is_g4 else ("Gemini 3" if clean.startswith("gemini-3") else "カスタム/API"),
        "tier_badge": "カスタムモデル" if clean else "★図面解析 最推奨",
        "free_tier": "契約プラン依存",
        "speed": "APIモデル参照",
        "accuracy": "APIモデル参照",
        "pricing": "Google AI料金表準拠",
        "supports_media_resolution": supports_mr,
        "is_gemini4": is_g4,
        "recommended_rank": 50,
        "recommended": False,
        "source": "手動指定/API",
        "description": "ユーザーが指定したモデルID、またはAPI一覧から取得したモデルです。",
    }


def select_optimal_gemini_model(
    models: Optional[List[Dict[str, Any]]] = None,
    priority: str = "balanced",
) -> str:
    """
    利用可能なモデル一覧（未指定時は標準カタログ）から、
    2D図面・部品表(BOM)解析に最適なモデルIDを優先方針に応じて自動選定する。
    priority:
      - "balanced": 精度・速度・コスト総合最適 (図面解析 標準推奨)
      - "accuracy": 難読・複雑図面向け 最高精度優先
      - "speed":    大量一括解析向け 高速・省コスト優先
    """
    pool = models if models else GEMINI_MODEL_CATALOG
    available_ids = {strip_gemini_model_prefix(m.get("id", "")) for m in pool if m.get("id")}
    if not available_ids:
        return DEFAULT_OPTIMAL_GEMINI_MODEL

    priority_orders: Dict[str, List[str]] = {
        "balanced": [
            "gemini-3.8-flash",
            "gemini-4-hydrogen",
            "gemini-3.6-flash",
            "gemini-4-argon",
            "gemini-3.5-flash-lite",
            "gemini-4-carbon",
            "gemini-3.1-pro-preview",
            "gemini-2.5-flash",
        ],
        "accuracy": [
            "gemini-4-carbon",
            "gemini-4-argon",
            "gemini-4-hydrogen",
            "gemini-3.8-flash",
            "gemini-3.1-pro-preview",
            "gemini-3.6-flash",
            "gemini-3.5-flash-lite",
        ],
        "speed": [
            "gemini-3.5-flash-lite",
            "gemini-3.8-flash",
            "gemini-3.6-flash",
            "gemini-4-hydrogen",
            "gemini-4-argon",
            "gemini-2.5-flash",
        ],
    }
    order = priority_orders.get(priority, priority_orders["balanced"])
    for cand in order:
        if cand in available_ids:
            return cand

    # 未知の最新世代モデル（例: gemini-4.* / gemini-3.*）のみが存在する場合はランク順で先頭を返す
    sorted_pool = sorted(
        [m for m in pool if m.get("id")],
        key=lambda x: (int(x.get("recommended_rank", 99)), str(x.get("id", ""))),
    )
    return strip_gemini_model_prefix(sorted_pool[0]["id"]) if sorted_pool else DEFAULT_OPTIMAL_GEMINI_MODEL


def format_model_choice_label(entry: Dict[str, Any]) -> str:
    mid = strip_gemini_model_prefix(str(entry.get("id", "")))
    badge = str(entry.get("tier_badge", "")).strip()
    return f"{mid}  [{badge}]" if badge else mid


def extract_model_id_from_choice(text: str) -> str:
    raw = str(text or "").strip()
    if "  [" in raw:
        raw = raw.split("  [", 1)[0].strip()
    elif " |" in raw:
        raw = raw.split(" |", 1)[0].strip()
    return strip_gemini_model_prefix(raw)


# ==============================================================================
# データモデル層
# ==============================================================================

@dataclass
class PartItem:
    uid: str = field(default_factory=lambda: str(uuid.uuid4()))
    item_no: str = "1"
    part_number: str = ""
    part_name: str = ""
    specification: str = ""
    material: str = ""
    treatment: str = ""
    quantity: float = 1.0
    unit: str = "個"
    unit_weight: float = 0.0
    category: str = "加工品"
    manufacturer: str = ""
    unit_price: float = 0.0
    source_file: str = ""
    source_page: int = 1
    remarks: str = ""
    bbox: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)

    @property
    def total_weight(self) -> float:
        return round(self.quantity * self.unit_weight, 4)

    @property
    def amount(self) -> float:
        return round(self.quantity * self.unit_price, 2)

    def formatted_qty(self) -> str:
        return str(int(self.quantity)) if float(self.quantity).is_integer() else f"{self.quantity:.2f}".rstrip("0").rstrip(".")

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["bbox"] = list(self.bbox)
        return d

    @staticmethod
    def _safe_float(val: Any, default: float = 0.0) -> float:
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val or "").replace(",", "").strip()
        m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
        if m:
            try:
                return float(m.group(0))
            except Exception:
                pass
        return default

    @staticmethod
    def _safe_int(val: Any, default: int = 1) -> int:
        if isinstance(val, int):
            return val
        s = str(val or "").strip()
        m = re.search(r"\d+", s)
        if m:
            try:
                return int(m.group(0))
            except Exception:
                pass
        return default

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PartItem":
        b = data.get("bbox", [0.0, 0.0, 0.0, 0.0])
        bbox_t = (
            (cls._safe_float(b[0]), cls._safe_float(b[1]), cls._safe_float(b[2]), cls._safe_float(b[3]))
            if isinstance(b, (list, tuple)) and len(b) == 4
            else (0.0, 0.0, 0.0, 0.0)
        )
        return cls(
            uid=str(data.get("uid", str(uuid.uuid4())) or str(uuid.uuid4())),
            item_no=str(data.get("item_no", "1") or "1"),
            part_number=str(data.get("part_number", "") or ""),
            part_name=str(data.get("part_name", "") or ""),
            specification=str(data.get("specification", "") or ""),
            material=str(data.get("material", "") or ""),
            treatment=str(data.get("treatment", "") or ""),
            quantity=max(cls._safe_float(data.get("quantity", 1.0), 1.0), 0.0),
            unit=str(data.get("unit", "個") or "個"),
            unit_weight=max(cls._safe_float(data.get("unit_weight", 0.0), 0.0), 0.0),
            category=str(data.get("category", "加工品") or "加工品"),
            manufacturer=str(data.get("manufacturer", "") or ""),
            unit_price=max(cls._safe_float(data.get("unit_price", 0.0), 0.0), 0.0),
            source_file=str(data.get("source_file", "") or ""),
            source_page=max(cls._safe_int(data.get("source_page", 1), 1), 1),
            remarks=str(data.get("remarks", "") or ""),
            bbox=bbox_t,
        )


@dataclass
class TextToken:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    page_no: int = 1

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) * 0.5

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) * 0.5


@dataclass
class DrawingPageData:
    page_no: int
    width: float
    height: float
    preview_image: Optional[Image.Image] = None
    tokens: List[TextToken] = field(default_factory=list)
    title_meta: Dict[str, str] = field(default_factory=dict)


@dataclass
class DrawingDocument:
    file_path: str
    file_name: str
    file_type: str
    pages: List[DrawingPageData] = field(default_factory=list)
    extracted_count: int = 0
    drawing_title: str = ""
    drawing_number: str = ""


@dataclass
class GeminiConfig:
    api_key: str = ""
    model_name: str = DEFAULT_OPTIMAL_GEMINI_MODEL
    auto_use_on_load: bool = True
    use_for_area_select: bool = True
    replace_existing_on_ai: bool = True
    custom_instruction: str = ""
    save_api_key: bool = True
    cached_models: List[Dict[str, Any]] = field(default_factory=list)

    def load(self) -> None:
        # 1. 環境変数 / .env の確認
        env_key = os.environ.get("GEMINI_API_KEY", "").strip()
        if not env_key and os.path.exists(".env"):
            try:
                with open(".env", "r", encoding="utf-8") as f:
                    for line in f:
                        if line.strip().startswith("GEMINI_API_KEY="):
                            env_key = line.strip().split("=", 1)[1].strip().strip("\"'")
            except Exception:
                pass
        if env_key:
            self.api_key = env_key

        # 2. 設定ファイルの読み込み
        if os.path.exists(CONFIG_FILE_PATH):
            try:
                with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                    d = json.load(f)
                if d.get("api_key"):
                    self.api_key = str(d["api_key"])
                raw_model = str(d.get("model_name", self.model_name))
                self.model_name = migrate_deprecated_model_id(raw_model)
                self.auto_use_on_load = bool(d.get("auto_use_on_load", True))
                self.use_for_area_select = bool(d.get("use_for_area_select", True))
                self.replace_existing_on_ai = bool(d.get("replace_existing_on_ai", True))
                self.custom_instruction = str(d.get("custom_instruction", ""))
                self.save_api_key = bool(d.get("save_api_key", True))
                raw_cache = d.get("cached_models", [])
                if isinstance(raw_cache, list):
                    self.cached_models = [m for m in raw_cache if isinstance(m, dict) and m.get("id")]
            except Exception:
                pass

    def save(self) -> None:
        try:
            self.model_name = strip_gemini_model_prefix(self.model_name) or DEFAULT_OPTIMAL_GEMINI_MODEL
            data = {
                "api_key": self.api_key if self.save_api_key else "",
                "model_name": self.model_name,
                "auto_use_on_load": self.auto_use_on_load,
                "use_for_area_select": self.use_for_area_select,
                "replace_existing_on_ai": self.replace_existing_on_ai,
                "custom_instruction": self.custom_instruction,
                "save_api_key": self.save_api_key,
                "cached_models": self.cached_models[:40],
            }
            with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


# ==============================================================================
# 重量計算・スマート部品ルール解析エンジン
# ==============================================================================

class SmartPartAnalyzer:
    @staticmethod
    def normalize_text(text: str) -> str:
        if not text:
            return ""
        zen = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ－（）×φΦ"
        han = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-()xφφ"
        cleaned = text.translate(str.maketrans(zen, han))
        return re.sub(r"\s+", " ", cleaned).strip()

    @classmethod
    def detect_material(cls, text: str) -> Tuple[str, str]:
        norm = cls.normalize_text(text)
        upper = norm.upper()
        for mat_key in sorted(MATERIAL_MASTER.keys(), key=len, reverse=True):
            if re.search(re.escape(mat_key), upper, re.IGNORECASE):
                return mat_key, norm
        alias_map = {
            "ジュラコン": "POM", "ポリアセタール": "POM", "テフロン": "PTFE",
            "アクリル": "PMMA", "塩ビ": "PVC", "真鍮": "C3604", "黄銅": "C3604",
            "黒皮": "SS400", "ミガキ": "S45C", "ステンレス": "SUS304", "アルミ": "A5052",
        }
        for alias, mat_key in alias_map.items():
            if alias in norm:
                return mat_key, norm
        return "", norm

    @classmethod
    def detect_treatment(cls, text: str) -> str:
        norm = cls.normalize_text(text)
        rules = [
            ("四三酸化鉄皮膜", "四三酸化鉄皮膜(黒染め)"), ("黒染", "四三酸化鉄皮膜(黒染め)"),
            ("無電解ニッケル", "無電解ニッケルめっき"), ("カニゼン", "無電解ニッケルめっき"),
            ("硬質クロム", "硬質クロムめっき"), ("三価クロメート", "三価クロメートめっき"),
            ("ユニクロ", "ユニクロめっき"), ("硬質アルマイト", "硬質アルマイト"),
            ("黒アルマイト", "黒アルマイト"), ("白アルマイト", "白アルマイト"),
            ("アルマイト", "白アルマイト"), ("高周波焼入", "高周波焼入れ"),
            ("浸炭焼入", "浸炭焼入れ"), ("焼入", "焼入れ研磨"),
            ("メラミン", "メラミン焼付塗装"), ("粉体塗装", "粉体塗装"),
            ("塗装", "ウレタン塗装"), ("酸洗", "酸洗・不動態化処理"), ("電解研磨", "電解研磨"),
        ]
        for kw, canonical in rules:
            if kw in norm:
                return canonical
        return ""

    @classmethod
    def detect_category(cls, part_name: str, spec: str, material: str, manufacturer: str) -> str:
        combined = f"{part_name} {spec} {manufacturer}".upper()
        for kw in STANDARD_PART_KEYWORDS:
            if kw.upper() in combined:
                if any(b in combined for b in ["ボルト", "ナット", "座金", "ワッシャ", "ピン", "キー", "止め輪", "ねじ"]):
                    return "標準部品"
                return "購入品"
        if re.search(r"\bM\d+(\.\d+)?\s*[xX×]\s*\d+", combined):
            return "標準部品"
        if manufacturer.strip():
            return "購入品"
        if any(w in combined for w in ["溶接", "フレーム", "架台", "製缶"]):
            return "製缶・溶接品"
        if any(w in combined for w in ["カバー", "ブラケット", "ステー", "パネル"]) and material in ("SPCC", "SECC", "SUS304", "A5052"):
            if re.search(r"\bt\s*[0-5](\.\d+)?\b", combined, re.IGNORECASE):
                return "板金品"
        return "加工品"

    @classmethod
    def estimate_weight_from_spec(cls, spec: str, material: str) -> float:
        if not spec:
            return 0.0
        norm = cls.normalize_text(spec)
        density = 7.85
        if material in MATERIAL_MASTER:
            density = float(MATERIAL_MASTER[material]["density"])
        else:
            mat, _ = cls.detect_material(material)
            if mat in MATERIAL_MASTER:
                density = float(MATERIAL_MASTER[mat]["density"])

        # 1. 角パイプ (A x B x tT x L または □A x B x t x L)
        m_sq_pipe = re.search(
            r"(?:□|角)?\s*(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)\s*[xX]\s*t?\s*(\d+(?:\.\d+)?)\s*[xX]\s*[L]?\s*(\d+(?:\.\d+)?)",
            norm,
            re.IGNORECASE,
        )
        if m_sq_pipe:
            a_out = float(m_sq_pipe.group(1))
            b_out = float(m_sq_pipe.group(2))
            t_w = float(m_sq_pipe.group(3))
            length = float(m_sq_pipe.group(4))
            if a_out > t_w * 2.0 and b_out > t_w * 2.0 and t_w <= 40.0:
                in_a = max(a_out - 2.0 * t_w, 0.0)
                in_b = max(b_out - 2.0 * t_w, 0.0)
                vol = (a_out * b_out - in_a * in_b) * length
                return round(vol * density / 1_000_000.0, 3)

        # 2. 丸パイプ φD x tT x L
        m_pipe = re.search(r"[φD]\s*(\d+(?:\.\d+)?)\s*[xX]\s*t?\s*(\d+(?:\.\d+)?)\s*[xX]\s*[L]?\s*(\d+(?:\.\d+)?)", norm, re.IGNORECASE)
        if m_pipe:
            d_out, t_w, length = float(m_pipe.group(1)), float(m_pipe.group(2)), float(m_pipe.group(3))
            if d_out > t_w * 2.0 and t_w < 30.0:
                r_o, r_i = d_out * 0.5, d_out * 0.5 - t_w
                return round(math.pi * (r_o * r_o - r_i * r_i) * length * density / 1_000_000.0, 3)

        # 3. 六角棒 HEX H x L / 六角 H x L
        m_hex = re.search(r"(?:HEX|六角|対辺)\s*(\d+(?:\.\d+)?)\s*[xX\-]\s*[L]?\s*(\d+(?:\.\d+)?)", norm, re.IGNORECASE)
        if m_hex:
            h_flat, length = float(m_hex.group(1)), float(m_hex.group(2))
            vol = (math.sqrt(3.0) * 0.5) * (h_flat * h_flat) * length
            return round(vol * density / 1_000_000.0, 3)

        # 4. プレート t x W x L
        m_plate = re.search(r"[tT]?\s*(\d+(?:\.\d+)?)\s*[xX]\s*[tT]?\s*(\d+(?:\.\d+)?)\s*[xX]\s*[tT]?\s*[L]?\s*(\d+(?:\.\d+)?)", norm)
        if m_plate:
            vol = float(m_plate.group(1)) * float(m_plate.group(2)) * float(m_plate.group(3))
            return round(vol * density / 1_000_000.0, 3)

        # 5. 丸棒 φD x L
        m_cyl = re.search(r"[φD]\s*(\d+(?:\.\d+)?)\s*[xX\-]\s*[L]?\s*(\d+(?:\.\d+)?)", norm, re.IGNORECASE)
        if m_cyl:
            r = float(m_cyl.group(1)) * 0.5
            return round(math.pi * (r * r) * float(m_cyl.group(2)) * density / 1_000_000.0, 3)

        # 6. ボルト M径 x 長さ
        m_bolt = re.search(r"\bM(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)", norm, re.IGNORECASE)
        if m_bolt:
            d, length = float(m_bolt.group(1)), float(m_bolt.group(2))
            r = d * 0.5
            return round(math.pi * (r * r) * (length + d * 1.1) * density / 1_000_000.0, 4)

        return 0.0

    @classmethod
    def classify_table_headers(cls, header_cells: List[str]) -> Dict[int, str]:
        col_map: Dict[int, str] = {}
        rules = [
            ("item_no", ["NO", "NO.", "項番", "番号", "符号", "ITEM", "POS"]),
            ("part_number", ["図番", "部品番号", "品番", "DRAWING", "DWG", "PART NO", "CODE"]),
            ("part_name", ["品名", "名称", "部品名", "部品名称", "NAME", "DESCRIPTION"]),
            ("specification", ["型式", "寸法", "仕様", "サイズ", "規格", "材料寸法", "SIZE", "SPEC", "MODEL"]),
            ("material", ["材質", "材料", "MATERIAL", "MAT"]),
            ("treatment", ["表面処理", "熱処理", "処理", "塗装", "FINISH", "TREATMENT"]),
            ("quantity", ["数量", "個数", "員数", "数", "QTY", "Q'TY", "QUANTITY"]),
            ("unit", ["単位", "UNIT"]),
            ("unit_weight", ["単重", "重量", "質量", "WEIGHT", "MASS"]),
            ("manufacturer", ["メーカー", "メーカ", "手配先", "仕入先", "MAKER", "VENDOR"]),
            ("unit_price", ["単価", "価格", "PRICE"]),
            ("remarks", ["備考", "記事", "摘要", "注記", "REMARK", "REMARKS", "NOTE"]),
        ]
        for idx, raw_cell in enumerate(header_cells):
            cell = cls.normalize_text(raw_cell).upper()
            if not cell:
                continue
            for field_name, keywords in rules:
                if field_name in col_map.values():
                    continue
                if any(kw == cell or (len(kw) >= 2 and kw in cell) for kw in keywords):
                    col_map[idx] = field_name
                    break
        return col_map

    @classmethod
    def parse_unstructured_row(
        cls, cells: List[str], default_idx: int, source_file: str, page_no: int, bbox: Tuple[float, float, float, float]
    ) -> Optional[PartItem]:
        cleaned = [cls.normalize_text(c) for c in cells if cls.normalize_text(c)]
        if not cleaned:
            return None
        joined = " ".join(cleaned)
        if sum(1 for w in ["品名", "材質", "数量", "備考", "図番", "名称"] if w in joined) >= 2:
            return None
        if any(w == joined for w in ["部品番号", "部品名称", "第三角法", "尺度", "SCALE", "PROJECTION"]):
            return None

        item = PartItem(item_no=str(default_idx), source_file=source_file, source_page=page_no, bbox=bbox)
        remaining: List[str] = []

        for idx, token in enumerate(cleaned):
            if idx == 0 and re.fullmatch(r"\d{1,3}", token):
                item.item_no = token
                continue
            mat, _ = cls.detect_material(token)
            if mat and not item.material and len(token) <= 18:
                item.material = mat
                treat = cls.detect_treatment(token)
                if treat and not item.treatment:
                    item.treatment = treat
                continue
            treat = cls.detect_treatment(token)
            if treat and not item.treatment and len(token) <= 20:
                item.treatment = treat
                continue
            m_qu = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(個|本|枚|式|組|セット|m|kg)", token)
            if m_qu:
                item.quantity, item.unit = float(m_qu.group(1)), m_qu.group(2)
                continue
            if idx >= 1 and re.fullmatch(r"\d{1,3}", token):
                val = int(token)
                if 1 <= val <= 500 and item.quantity == 1.0:
                    item.quantity = float(val)
                    continue
            m_wt = re.fullmatch(r"(\d+(?:\.\d+)?)\s*kg", token, re.IGNORECASE)
            if m_wt:
                item.unit_weight = float(m_wt.group(1))
                continue
            if re.search(r"([tT]\d+|\d+\s*[xX×]\s*\d+|[φD]\d+|\bM\d+\s*[xX×]\s*\d+)", token) and not item.specification:
                item.specification = token
                continue
            if not re.search(r"[ぁ-んァ-ン一-龥]", token) and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-_/\.\(\)]{2,24}", token):
                if not item.part_number:
                    item.part_number = token
                    continue
                elif not item.specification:
                    item.specification = token
                    continue
            remaining.append(token)

        for tp in remaining:
            if any(mk in tp for mk in ["ミスミ", "MISUMI", "THK", "SMC", "CKD", "NSK", "NTN", "IKO", "オリエンタル", "キーエンス", "オムロン", "NBK"]) and not item.manufacturer:
                item.manufacturer = tp
            elif not item.part_name:
                item.part_name = tp
            elif not item.specification:
                item.specification = tp
            elif not item.remarks:
                item.remarks = tp
            else:
                item.remarks = f"{item.remarks} {tp}"

        if not item.part_name and item.part_number:
            item.part_name = item.part_number
        if not item.part_name and not item.part_number and not item.specification:
            return None
        if item.material and not item.treatment and item.material in MATERIAL_MASTER:
            item.treatment = MATERIAL_MASTER[item.material]["treat"]
        if item.unit_weight == 0.0 and item.specification:
            item.unit_weight = cls.estimate_weight_from_spec(item.specification, item.material)
        item.category = cls.detect_category(item.part_name, item.specification, item.material, item.manufacturer)
        return item


# ==============================================================================
# Gemini API マルチモーダル図面解析エンジン
# ==============================================================================

class GeminiVisionAnalyzer:
    """
    Gemini API (google-genai SDK および REST API フォールバック) を用いて
    2D図面の画像＋ベクターテキストから高精度に部品表(BOM)を構造化抽出するエンジン
    """

    SYSTEM_PROMPT = """あなたは熟練した機械設計・製造エンジニアであり、2D図面（PDF/CAD組立図・部品図）の解析エキスパートです。
与えられた2D図面画像および図面内の抽出テキスト情報から、部品表（BOM: Bill of Materials）のすべての部品行と表題欄情報を漏れなく正確に抽出し、指定されたJSONフォーマットのみで出力してください。

【厳守する解析ルール】
1. 組立図に「部品表（部品欄）」がある場合は、すべての行（下から上に番号が振られている場合も上から下の場合も含む）を1行も漏らさず抽出してください。
2. 図面内のバルーン（丸数字や引き出し線）や注記に部品情報がある場合はそれも補完してください。
3. 単品図（1枚の図面に1つの加工部品のみが描かれている図面）で部品表がない場合は、表題欄の「図面名称」「図番」「材質」「処理」「数量」および投影図内の外形寸法（例: 板厚t12, 幅100, 長さ250 -> t12x100x250、丸棒φ30, 長さ150 -> φ30x150）を読み取って1件の部品データとして構築してください。
4. 材質は可能な限りJIS規格記号（SS400, S45C, SUS304, SUS303, A5052, A6063, SCM435,SKD11, FC200, MCナイロン, POM 等）に正規化してください。
5. specification（型式・材料寸法）には、加工品なら材料寸法（例: t10x180x550, φ30x420）、購入品や標準部品なら型番・サイズ（例: M8x25, UCFL206, 6005ZZ）を記載してください。
6. unit_weight（単重 kg）が図面に記載されていればその数値を、記載がなくても寸法と材質から計算可能であれば概算単重(kg)を数値で入れてください。
7. category（調達区分）は「加工品」「購入品」「標準部品」「板金品」「製缶・溶接品」のいずれかを適切に判定してください。
8. box_2d には、各部品行（または該当部品）が図面画像内のどこにあるかを 0〜1000 の正規化座標 [ymin, xmin, ymax, xmax]（整数4つ）で出力してください。不明な場合は [0, 0, 0, 0] としてください。
9. 出力は必ず純粋なJSONオブジェクトのみとし、前後の説明文は一切含めないでください。"""

    EXCLUDED_MODEL_KEYWORDS: Tuple[str, ...] = (
        "embedding", "aqa", "imagen", "veo", "lyria", "-tts",
        "native-audio", "live", "robotics", "gemma", "learnlm",
        "decision", "nano-banana",
    )

    @classmethod
    def _is_candidate_drawing_model(cls, model_id: str, supported_methods: List[str]) -> bool:
        mid = strip_gemini_model_prefix(model_id).lower()
        if not mid or "gemini" not in mid:
            return False
        if mid in DEPRECATED_GEMINI_MODEL_REPLACEMENTS:
            return False
        if any(bad in mid for bad in cls.EXCLUDED_MODEL_KEYWORDS):
            return False
        if supported_methods:
            lower_methods = {str(m).lower() for m in supported_methods}
            if "generatecontent" not in lower_methods:
                return False
        return True

    @classmethod
    def parse_models_api_payload(cls, resp_json: Dict[str, Any]) -> List[Dict[str, Any]]:
        results: List[Dict[str, Any]] = []
        raw_list = resp_json.get("models", [])
        if not isinstance(raw_list, list):
            return results
        for m in raw_list:
            if not isinstance(m, dict):
                continue
            mid = strip_gemini_model_prefix(str(m.get("name", "") or m.get("baseModelId", "")))
            methods = m.get("supportedGenerationMethods") or m.get("supportedActions") or []
            methods_list = [str(x) for x in methods] if isinstance(methods, list) else []
            if not cls._is_candidate_drawing_model(mid, methods_list):
                continue
            results.append({
                "id": mid,
                "display_name": str(m.get("displayName", "") or mid).strip(),
                "description": str(m.get("description", "") or "").strip(),
                "input_token_limit": int(m.get("inputTokenLimit", 0) or 0),
                "output_token_limit": int(m.get("outputTokenLimit", 0) or 0),
                "supported_methods": methods_list,
            })
        return results

    @staticmethod
    def _infer_dynamic_rank_and_badge(model_id: str) -> Tuple[int, str, str, str, str]:
        low = model_id.lower()
        if "gemini-4" in low:
            if "carbon" in low:
                return 10, "Gemini 4 最高精度", "じっくり推論", "最高峰", "Gemini 4"
            if "argon" in low:
                return 11, "Gemini 4 高度推論", "標準", "最高クラス", "Gemini 4"
            if "hydrogen" in low:
                return 12, "Gemini 4 バランス", "高速", "非常に高い", "Gemini 4"
            return 14, "Gemini 4 (API最新)", "標準〜高速", "非常に高い", "Gemini 4"
        if "gemini-3.8" in low or "gemini-3.9" in low:
            return 15, "Gemini 3.8+ (API)", "高速", "非常に高い", "Gemini 3.8+"
        if "gemini-3" in low:
            if "flash-lite" in low:
                return 25, "Flash-Lite (軽量)", "超高速", "高い", "Gemini 3"
            if "pro" in low:
                return 18, "Pro (API)", "標準", "非常に高い", "Gemini 3"
            return 20, "Flash (API)", "高速", "高い", "Gemini 3"
        if "flash-latest" in low:
            return 22, "最新Flashエイリアス", "高速", "高い", "Latest"
        if "pro-latest" in low:
            return 24, "最新Proエイリアス", "標準", "非常に高い", "Latest"
        if "gemini-2.5" in low:
            return 80, "旧世代 (互換用)", "高速", "標準", "Gemini 2.5"
        return 60, "API取得モデル", "標準", "標準〜高い", "Gemini API"

    @classmethod
    def merge_catalog_with_api_models(cls, api_models: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        組み込みの最新カタログと API から取得した利用可能モデル一覧を統合し、
        2D図面・BOM解析に最適な順へソートして返す。
        """
        catalog_map = get_gemini_model_catalog_map()
        merged_by_id: Dict[str, Dict[str, Any]] = {}

        for raw in api_models:
            mid = strip_gemini_model_prefix(str(raw.get("id", "")))
            if not mid or mid in DEPRECATED_GEMINI_MODEL_REPLACEMENTS:
                continue
            if mid in catalog_map:
                entry = copy.deepcopy(catalog_map[mid])
                entry["source"] = "API利用可能"
                if raw.get("input_token_limit"):
                    entry["input_token_limit"] = raw["input_token_limit"]
                if raw.get("output_token_limit"):
                    entry["output_token_limit"] = raw["output_token_limit"]
                merged_by_id[mid] = entry
            else:
                rank, badge, speed, acc, gen = cls._infer_dynamic_rank_and_badge(mid)
                api_desc = str(raw.get("description", "") or "").strip()
                if not api_desc:
                    api_desc = "Gemini API から動的に取得された生成対応モデルです。"
                merged_by_id[mid] = {
                    "id": mid,
                    "display_name": str(raw.get("display_name", "") or mid),
                    "generation": gen,
                    "tier_badge": badge,
                    "free_tier": "APIキー権限に準拠",
                    "speed": speed,
                    "accuracy": acc,
                    "pricing": "Google AI料金表参照",
                    "supports_media_resolution": mid.startswith("gemini-3") or mid.startswith("gemini-4") or "latest" in mid,
                    "is_gemini4": mid.startswith("gemini-4"),
                    "recommended_rank": rank,
                    "recommended": rank <= 20,
                    "source": "API利用可能",
                    "description": api_desc,
                    "input_token_limit": int(raw.get("input_token_limit", 0) or 0),
                    "output_token_limit": int(raw.get("output_token_limit", 0) or 0),
                }

        # API未収載でも標準カタログの推奨項目を末尾に保持（参考表示用）
        for cat_entry in GEMINI_MODEL_CATALOG:
            cid = cat_entry["id"]
            if cid not in merged_by_id:
                fallback_entry = copy.deepcopy(cat_entry)
                fallback_entry["source"] = "カタログ (参考)"
                fallback_entry["recommended_rank"] = int(cat_entry.get("recommended_rank", 50)) + 100
                merged_by_id[cid] = fallback_entry

        sorted_entries = sorted(
            merged_by_id.values(),
            key=lambda x: (
                0 if x.get("source") == "API利用可能" else 1,
                int(x.get("recommended_rank", 99)),
                str(x.get("id", "")),
            ),
        )
        return sorted_entries

    @classmethod
    def fetch_available_models(cls, api_key: str) -> List[Dict[str, Any]]:
        """
        Gemini API (google-genai SDK または REST API v1beta/models) から
        現在のAPIキーで利用可能な最新モデル一覧を取得し、最適度順に整理して返す。
        """
        clean_key = api_key.strip()
        if not clean_key:
            raise ValueError("APIキーが入力されていません。")

        collected: List[Dict[str, Any]] = []

        # 1. google-genai SDK で一覧取得を試行
        if HAS_GOOGLE_GENAI:
            try:
                client = genai.Client(api_key=clean_key)
                for m in client.models.list():
                    mid = strip_gemini_model_prefix(getattr(m, "name", "") or getattr(m, "base_model_id", ""))
                    actions = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", None) or []
                    actions_list = [str(a) for a in actions] if actions else []
                    if not cls._is_candidate_drawing_model(mid, actions_list):
                        continue
                    collected.append({
                        "id": mid,
                        "display_name": str(getattr(m, "display_name", "") or mid).strip(),
                        "description": str(getattr(m, "description", "") or "").strip(),
                        "input_token_limit": int(getattr(m, "input_token_limit", 0) or 0),
                        "output_token_limit": int(getattr(m, "output_token_limit", 0) or 0),
                        "supported_methods": actions_list,
                    })
                if collected:
                    return cls.merge_catalog_with_api_models(collected)
            except Exception:
                collected = []

        # 2. REST API (v1beta/models) でページネーション取得
        encoded_key = urllib.parse.quote(clean_key, safe="")
        page_token = ""
        for _ in range(5):
            url = f"https://generativelanguage.googleapis.com/v1beta/models?key={encoded_key}&pageSize=100"
            if page_token:
                url += f"&pageToken={urllib.parse.quote(page_token, safe='')}"
            req = urllib.request.Request(url, method="GET")
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    resp_json = json.loads(resp.read().decode("utf-8"))
                collected.extend(cls.parse_models_api_payload(resp_json))
                page_token = str(resp_json.get("nextPageToken", "") or "").strip()
                if not page_token:
                    break
            except urllib.error.HTTPError as he:
                detail = ""
                try:
                    detail = he.read().decode("utf-8")
                except Exception:
                    pass
                raise RuntimeError(f"モデル一覧の取得に失敗しました (HTTP {he.code}): {detail[:250]}")
            except Exception as e:
                raise RuntimeError(f"モデル一覧の取得中に通信エラーが発生しました: {e}")

        if not collected:
            raise RuntimeError("利用可能な Gemini 生成モデルを取得できませんでした。APIキーの権限をご確認ください。")

        return cls.merge_catalog_with_api_models(collected)

    @classmethod
    def test_connection(cls, config: GeminiConfig) -> Tuple[bool, str]:
        if not config.api_key.strip():
            return False, "APIキーが入力されていません。"
        prompt = '次のJSONのみを返してください: {"status": "ok", "message": "接続成功"}'
        try:
            res_text = cls._call_gemini_api(config, prompt, image=None)
            if "ok" in res_text.lower() or "接続" in res_text:
                return True, f"Gemini API ({config.model_name}) への接続に成功しました！"
            return True, f"応答を受信しました: {res_text[:80]}"
        except Exception as e:
            return False, f"接続エラー: {e}"

    @classmethod
    def analyze_drawing_page(
        cls,
        config: GeminiConfig,
        page_data: DrawingPageData,
        source_file: str,
        crop_rect: Optional[Tuple[float, float, float, float]] = None,
    ) -> Tuple[List[PartItem], Dict[str, str]]:
        """
        ページ全体または選択矩形範囲の画像＋テキストトークンをGemini APIへ送信し、
        抽出された PartItem リストと表題欄メタデータ辞書を返す。
        """
        if not config.api_key.strip():
            raise ValueError("Gemini APIキーが設定されていません。上部の「Gemini AI設定」から設定してください。")
        if page_data.preview_image is None:
            raise ValueError("解析対象の図面プレビュー画像が存在しません。")

        img_to_send = page_data.preview_image
        iw, ih = img_to_send.size
        offset_x, offset_y = 0.0, 0.0
        region_w, region_h = page_data.width, page_data.height

        # 範囲選択時はその領域を高解像度クロップして送信
        if crop_rect is not None:
            px0, py0, px1, py1 = crop_rect
            offset_x, offset_y = px0, py0
            region_w = max(px1 - px0, 1.0)
            region_h = max(py1 - py0, 1.0)
            cx0 = int(max(0.0, (px0 / max(page_data.width, 1.0)) * iw))
            cy0 = int(max(0.0, (py0 / max(page_data.height, 1.0)) * ih))
            cx1 = int(min(float(iw), (px1 / max(page_data.width, 1.0)) * iw))
            cy1 = int(min(float(ih), (py1 / max(page_data.height, 1.0)) * ih))
            if cx1 - cx0 > 10 and cy1 - cy0 > 10:
                img_to_send = img_to_send.crop((cx0, cy0, cx1, cy1))

        # ベクターテキストトークンの補助コンテキスト化（文字認識のクロスチェック用）
        relevant_tokens = page_data.tokens
        if crop_rect is not None:
            px0, py0, px1, py1 = crop_rect
            relevant_tokens = [
                t for t in page_data.tokens
                if t.cx >= px0 and t.cx <= px1 and t.cy >= py0 and t.cy <= py1
            ]

        sorted_tokens = sorted(relevant_tokens, key=lambda t: (round(t.cy / 8.0), t.x0))
        token_lines: List[str] = []
        for t in sorted_tokens[:350]:
            rel_x = int(((t.cx - offset_x) / max(region_w, 1.0)) * 1000)
            rel_y = int(((t.cy - offset_y) / max(region_h, 1.0)) * 1000)
            token_lines.append(f"[{rel_y},{rel_x}] {t.text}")
        vector_context = "\n".join(token_lines) if token_lines else "(ベクター文字なし - 画像から直接OCR解析してください)"

        extra_inst = f"\n【ユーザー追加指示】\n{config.custom_instruction}\n" if config.custom_instruction.strip() else ""
        mode_desc = "選択された矩形範囲（部品欄・注記エリア）" if crop_rect else f"2D図面全体 (ファイル名: {source_file}, ページ: {page_data.page_no})"

        user_prompt = f"""{cls.SYSTEM_PROMPT}
{extra_inst}
【解析対象】
{mode_desc}

【図面ファイルから抽出されたベクター文字と正規化座標 [Y(0-1000), X(0-1000)] ヒント】
{vector_context}

【出力JSONスキーマ】
以下のキーを持つJSONオブジェクトを出力してください:
{{
  "drawing_title": "図面名称（不明なら空文字）",
  "drawing_number": "図面番号（不明なら空文字）",
  "parts": [
    {{
      "item_no": "1",
      "part_number": "図番または型番",
      "part_name": "品名・部品名称",
      "specification": "型式・材料寸法・仕様",
      "material": "材質記号",
      "treatment": "表面処理・熱処理",
      "quantity": 1,
      "unit": "個",
      "unit_weight": 0.0,
      "category": "加工品",
      "manufacturer": "メーカー名",
      "remarks": "備考",
      "box_2d": [ymin, xmin, ymax, xmax]
    }}
  ]
}}"""

        raw_json_str = cls._call_gemini_api(config, user_prompt, image=img_to_send)
        parsed = cls._extract_json_dict(raw_json_str)

        meta = {
            "drawing_title": str(parsed.get("drawing_title", "") or "").strip(),
            "drawing_number": str(parsed.get("drawing_number", "") or "").strip(),
        }

        parts_list: List[PartItem] = []
        raw_parts = parsed.get("parts", [])
        if isinstance(raw_parts, list):
            for idx, rp in enumerate(raw_parts, start=1):
                if not isinstance(rp, dict):
                    continue
                item = cls._dict_to_part_item(
                    rp, idx, source_file, page_data.page_no, offset_x, offset_y, region_w, region_h
                )
                if item:
                    parts_list.append(item)

        return parts_list, meta

    @classmethod
    def refine_bom_parts(cls, config: GeminiConfig, current_parts: List[PartItem]) -> List[PartItem]:
        """既存の部品表リストをGemini APIで一括クレンジング・規格正規化・単重/区分/メーカー推論する"""
        if not config.api_key.strip():
            raise ValueError("Gemini APIキーが設定されていません。")
        if not current_parts:
            return []

        simplified_input = []
        for p in current_parts:
            simplified_input.append({
                "uid": p.uid,
                "item_no": p.item_no,
                "part_number": p.part_number,
                "part_name": p.part_name,
                "specification": p.specification,
                "material": p.material,
                "treatment": p.treatment,
                "quantity": p.quantity,
                "unit": p.unit,
                "unit_weight": p.unit_weight,
                "category": p.category,
                "manufacturer": p.manufacturer,
                "remarks": p.remarks,
            })

        prompt = f"""あなたは機械設計・購買管理のエキスパートAIです。
以下の部品表(BOM)データ配列を精査し、各行のデータをより正確かつ標準的な表記に補正・補完してJSONで返してください。

【補正ルール】
1. uid は入力と同じ値を必ずそのまま保持してください。
2. 材質(material)に表記ゆれがあればJIS規格記号（SS400, S45C, SUS304, A5052, SCM435, POM, MCナイロン等）に統一してください。
3. 表面処理(treatment)が空欄で、一般的な加工部品かつ材質から標準処理が明らかな場合は推奨処理を補完してください。
4. 型式・寸法(specification)や図番(part_number)に市販品の型番（例: UCFL206, 6005ZZ, CDQ2B32, M8x25等）が含まれる場合、品名(part_name)、メーカー名(manufacturer)、調達区分(category: 購入品/標準部品)を正確に補完してください。
5. 単重(unit_weight)が 0.0 で、寸法と材質から1個あたりの質量(kg)が計算・推定できる場合は、正確な単重(kg)を計算して数値で設定してください。

【入力部品リストJSON】
{json.dumps(simplified_input, ensure_ascii=False)}

【出力JSON形式】
{{"parts": [ ...補正後のオブジェクト配列... ]}}"""

        raw_res = cls._call_gemini_api(config, prompt, image=None)
        parsed = cls._extract_json_dict(raw_res)
        refined_list = parsed.get("parts", [])
        if not isinstance(refined_list, list):
            return current_parts

        refined_map = {str(r.get("uid", "")): r for r in refined_list if isinstance(r, dict)}
        updated_parts: List[PartItem] = []

        for orig in current_parts:
            cp = copy.deepcopy(orig)
            if cp.uid in refined_map:
                r = refined_map[cp.uid]
                cp.part_number = str(r.get("part_number", cp.part_number) or cp.part_number).strip()
                cp.part_name = str(r.get("part_name", cp.part_name) or cp.part_name).strip()
                cp.specification = str(r.get("specification", cp.specification) or cp.specification).strip()
                cp.material = str(r.get("material", cp.material) or cp.material).strip()
                cp.treatment = str(r.get("treatment", cp.treatment) or cp.treatment).strip()
                cp.unit = str(r.get("unit", cp.unit) or cp.unit).strip()
                cp.category = str(r.get("category", cp.category) or cp.category).strip()
                cp.manufacturer = str(r.get("manufacturer", cp.manufacturer) or cp.manufacturer).strip()
                cp.remarks = str(r.get("remarks", cp.remarks) or cp.remarks).strip()
                try:
                    w = float(r.get("unit_weight", cp.unit_weight) or 0.0)
                    if w > 0:
                        cp.unit_weight = round(w, 4)
                except Exception:
                    pass
            if cp.unit_weight == 0.0 and cp.specification:
                cp.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(cp.specification, cp.material)
            updated_parts.append(cp)

        return updated_parts

    @classmethod
    def _dict_to_part_item(
        cls,
        d: Dict[str, Any],
        default_idx: int,
        source_file: str,
        page_no: int,
        offset_x: float,
        offset_y: float,
        region_w: float,
        region_h: float,
    ) -> Optional[PartItem]:
        p_num = str(d.get("part_number", "") or "").strip()
        p_name = str(d.get("part_name", "") or "").strip()
        spec = str(d.get("specification", "") or "").strip()
        if not p_name and not p_num and not spec:
            return None

        item_no = str(d.get("item_no", default_idx) or default_idx).strip()
        mat = str(d.get("material", "") or "").strip()
        treat = str(d.get("treatment", "") or "").strip()
        unit = str(d.get("unit", "個") or "個").strip()
        cat = str(d.get("category", "") or "").strip()
        maker = str(d.get("manufacturer", "") or "").strip()
        remarks = str(d.get("remarks", "") or "").strip()

        try:
            qty_raw = str(d.get("quantity", 1) or 1)
            m_q = re.search(r"\d+(?:\.\d+)?", qty_raw)
            qty = float(m_q.group(0)) if m_q else 1.0
        except Exception:
            qty = 1.0

        try:
            wt_raw = str(d.get("unit_weight", 0.0) or 0.0)
            m_w = re.search(r"\d+(?:\.\d+)?", wt_raw)
            u_weight = float(m_w.group(0)) if m_w else 0.0
        except Exception:
            u_weight = 0.0

        # box_2d [ymin, xmin, ymax, xmax] (0-1000) -> ページ実座標へ変換
        bbox = (0.0, 0.0, 0.0, 0.0)
        b2d = d.get("box_2d")
        if isinstance(b2d, (list, tuple)) and len(b2d) == 4:
            try:
                ymin, xmin, ymax, xmax = float(b2d[0]), float(b2d[1]), float(b2d[2]), float(b2d[3])
                if ymax > ymin and xmax > xmin:
                    bx0 = offset_x + (xmin / 1000.0) * region_w
                    by0 = offset_y + (ymin / 1000.0) * region_h
                    bx1 = offset_x + (xmax / 1000.0) * region_w
                    by1 = offset_y + (ymax / 1000.0) * region_h
                    bbox = (bx0, by0, bx1, by1)
            except Exception:
                pass

        if u_weight == 0.0 and spec:
            u_weight = SmartPartAnalyzer.estimate_weight_from_spec(spec, mat)
        if not cat or cat not in CATEGORIES:
            cat = SmartPartAnalyzer.detect_category(p_name, spec, mat, maker)

        return PartItem(
            item_no=item_no,
            part_number=p_num,
            part_name=p_name or p_num,
            specification=spec,
            material=mat,
            treatment=treat,
            quantity=max(qty, 0.0),
            unit=unit,
            unit_weight=max(u_weight, 0.0),
            category=cat,
            manufacturer=maker,
            source_file=source_file,
            source_page=page_no,
            remarks=remarks,
            bbox=bbox,
        )

    @staticmethod
    def _supports_media_resolution(model_name: str) -> bool:
        low = model_name.lower()
        return low.startswith("gemini-3") or low.startswith("gemini-4") or "latest" in low

    @staticmethod
    def _is_gemini4_model(model_name: str) -> bool:
        return model_name.lower().startswith("gemini-4")

    @classmethod
    def _call_gemini_api(cls, config: GeminiConfig, prompt: str, image: Optional[Image.Image] = None) -> str:
        """
        google-genai SDK が利用可能ならSDK経由で呼び出し、
        失敗または未導入環境では標準ライブラリ urllib.request による REST API で確実に呼び出す。
        Gemini 3 / 4 系モデルでは画像送信時に MEDIA_RESOLUTION_HIGH を適用し図面読取精度を高める。
        """
        api_key = config.api_key.strip()
        model_name = strip_gemini_model_prefix(config.model_name.strip()) or DEFAULT_OPTIMAL_GEMINI_MODEL
        use_mr = (image is not None) and cls._supports_media_resolution(model_name)
        is_g4 = cls._is_gemini4_model(model_name)

        # 1. google-genai SDK 経由の試行
        if HAS_GOOGLE_GENAI:
            for try_mr in ([True, False] if use_mr else [False]):
                try:
                    client = genai.Client(api_key=api_key)
                    contents: List[Any] = []
                    if image is not None:
                        contents.append(image)
                    contents.append(prompt)

                    if try_mr:
                        mr_enum = getattr(genai_types, "MediaResolution", None)
                        mr_val = mr_enum.MEDIA_RESOLUTION_HIGH if (mr_enum is not None and hasattr(mr_enum, "MEDIA_RESOLUTION_HIGH")) else "MEDIA_RESOLUTION_HIGH"
                        if is_g4:
                            gen_cfg = genai_types.GenerateContentConfig(
                                response_mime_type="application/json",
                                media_resolution=mr_val,
                            )
                        else:
                            gen_cfg = genai_types.GenerateContentConfig(
                                temperature=0.1,
                                response_mime_type="application/json",
                                media_resolution=mr_val,
                            )
                    else:
                        if is_g4:
                            gen_cfg = genai_types.GenerateContentConfig(
                                response_mime_type="application/json",
                            )
                        else:
                            gen_cfg = genai_types.GenerateContentConfig(
                                temperature=0.1,
                                response_mime_type="application/json",
                            )
                    response = client.models.generate_content(
                        model=model_name,
                        contents=contents,
                        config=gen_cfg,
                    )
                    if response and response.text:
                        return response.text
                except Exception:
                    # 非対応パラメータ時は基本設定で再試行、さらにダメならREST APIへフォールバック
                    continue

        # 2. 標準 urllib.request による Gemini REST API 直接呼び出し
        encoded_key = urllib.parse.quote(api_key, safe="")
        encoded_model = urllib.parse.quote(model_name, safe="-._")
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{encoded_model}:generateContent?key={encoded_key}"
        parts_payload: List[Dict[str, Any]] = []

        if image is not None:
            buf = io.BytesIO()
            image.save(buf, format="PNG")
            b64_data = base64.b64encode(buf.getvalue()).decode("ascii")
            parts_payload.append({
                "inline_data": {
                    "mime_type": "image/png",
                    "data": b64_data,
                }
            })

        parts_payload.append({"text": prompt})

        for try_mr in ([True, False] if use_mr else [False]):
            gen_config: Dict[str, Any] = {
                "responseMimeType": "application/json",
            }
            if not is_g4:
                gen_config["temperature"] = 0.1
            if try_mr:
                gen_config["mediaResolution"] = "MEDIA_RESOLUTION_HIGH"

            payload = {
                "contents": [{"parts": parts_payload}],
                "generationConfig": gen_config,
            }

            req_bytes = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(
                url,
                data=req_bytes,
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            try:
                with urllib.request.urlopen(req, timeout=120) as resp:
                    resp_body = resp.read().decode("utf-8")
                    resp_json = json.loads(resp_body)
                    candidates = resp_json.get("candidates", [])
                    if candidates:
                        c_parts = candidates[0].get("content", {}).get("parts", [])
                        texts = [p.get("text", "") for p in c_parts if "text" in p]
                        return "\n".join(texts)
                    raise ValueError(f"Gemini APIから有効な候補が返されませんでした: {resp_body[:200]}")
            except urllib.error.HTTPError as he:
                err_detail = ""
                try:
                    err_detail = he.read().decode("utf-8")
                except Exception:
                    pass
                # mediaResolution 非対応モデルで 400 が返った場合は基本パラメータでリトライ
                if he.code == 400 and try_mr:
                    continue
                raise RuntimeError(f"Gemini API HTTPエラー ({he.code}): {err_detail[:300]}")
            except Exception as e:
                raise RuntimeError(f"Gemini API 通信エラー: {e}")

        raise RuntimeError("Gemini API の呼び出しに失敗しました。")

    @staticmethod
    def _extract_json_dict(raw_text: str) -> Dict[str, Any]:
        cleaned = raw_text.strip()
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s*```$", "", cleaned)
        try:
            obj = json.loads(cleaned)
            if isinstance(obj, dict):
                return obj
            if isinstance(obj, list):
                return {"parts": obj}
        except Exception:
            pass
        m = re.search(r"\{[\s\S]*\}", cleaned)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
        raise ValueError("AI応答からJSONデータを解析できませんでした。")


# ==============================================================================
# ローカル 2D図面解析エンジン層 (PDF & CAD DXF)
# ==============================================================================

SUPPORTED_IMAGE_EXTENSIONS: Tuple[str, ...] = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")
SUPPORTED_DRAWING_EXTENSIONS: Tuple[str, ...] = (".pdf", ".dxf") + SUPPORTED_IMAGE_EXTENSIONS


class DrawingParserEngine:
    @classmethod
    def parse_file(cls, file_path: str) -> Tuple[DrawingDocument, List[PartItem]]:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".pdf":
            return cls.parse_pdf(file_path)
        elif ext == ".dxf":
            return cls.parse_dxf(file_path)
        elif ext in SUPPORTED_IMAGE_EXTENSIONS:
            return cls.parse_image(file_path)
        raise ValueError(f"未対応のファイル形式です: {ext} (対応: PDF, DXF, PNG, JPG, BMP, TIFF, WEBP)")

    @classmethod
    def parse_image(cls, file_path: str) -> Tuple[DrawingDocument, List[PartItem]]:
        file_name = os.path.basename(file_path)
        doc_model = DrawingDocument(file_path=file_path, file_name=file_name, file_type="IMAGE")
        with Image.open(file_path) as src_img:
            rgb_img = src_img.convert("RGB")
        iw, ih = rgb_img.size
        page_data = DrawingPageData(
            page_no=1,
            width=float(max(iw, 1)),
            height=float(max(ih, 1)),
            preview_image=rgb_img,
            tokens=[],
            title_meta={"drawing_title": os.path.splitext(file_name)[0], "drawing_number": "", "material": ""},
        )
        doc_model.drawing_title = os.path.splitext(file_name)[0]
        doc_model.pages.append(page_data)
        doc_model.extracted_count = 0
        return doc_model, []

    @classmethod
    def parse_pdf(cls, file_path: str) -> Tuple[DrawingDocument, List[PartItem]]:
        file_name = os.path.basename(file_path)
        doc_model = DrawingDocument(file_path=file_path, file_name=file_name, file_type="PDF")
        all_parts: List[PartItem] = []

        pdf_doc = fitz.open(file_path)
        try:
            for page_idx in range(len(pdf_doc)):
                page = pdf_doc[page_idx]
                page_no = page_idx + 1
                width, height = float(page.rect.width), float(page.rect.height)

                zoom = min(max(1700.0 / max(width, height, 1.0), 1.3), 3.0)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                tokens = [
                    TextToken(str(w[4]).strip(), float(w[0]), float(w[1]), float(w[2]), float(w[3]), page_no)
                    for w in page.get_text("words") if str(w[4]).strip()
                ]
                page_data = DrawingPageData(page_no=page_no, width=width, height=height, preview_image=img, tokens=tokens)

                page_parts: List[PartItem] = []
                try:
                    for tab in page.find_tables():
                        extracted = tab.extract()
                        if extracted and len(extracted) >= 2:
                            tb_bbox = tuple(float(v) for v in tab.bbox)
                            page_parts.extend(cls._extract_parts_from_2d_cells(extracted, file_name, page_no, tb_bbox))
                except Exception:
                    pass

                if not page_parts and tokens:
                    page_parts.extend(cls.extract_parts_from_tokens(tokens, file_name, page_no, (0.0, 0.0, width, height)))

                title_meta = cls._detect_title_block_meta(tokens, width, height)
                page_data.title_meta = title_meta
                if title_meta.get("drawing_title") and not doc_model.drawing_title:
                    doc_model.drawing_title = title_meta["drawing_title"]
                if title_meta.get("drawing_number") and not doc_model.drawing_number:
                    doc_model.drawing_number = title_meta["drawing_number"]

                if not page_parts and tokens:
                    single_part = cls._extract_single_part_fallback(tokens, title_meta, file_name, page_no, width, height)
                    if single_part:
                        page_parts.append(single_part)

                doc_model.pages.append(page_data)
                all_parts.extend(page_parts)
        finally:
            pdf_doc.close()

        doc_model.extracted_count = len(all_parts)
        return doc_model, all_parts

    @classmethod
    def parse_dxf(cls, file_path: str) -> Tuple[DrawingDocument, List[PartItem]]:
        file_name = os.path.basename(file_path)
        doc_model = DrawingDocument(file_path=file_path, file_name=file_name, file_type="DXF")
        all_parts: List[PartItem] = []

        try:
            dxf_doc = ezdxf.readfile(file_path)
        except UnicodeDecodeError:
            dxf_doc = ezdxf.readfile(file_path, encoding="cp932")

        msp = dxf_doc.modelspace()
        raw_texts: List[Tuple[str, float, float, float]] = []
        lines_data: List[Tuple[float, float, float, float, str]] = []
        circles_data: List[Tuple[float, float, float]] = []
        attrib_parts: List[PartItem] = []
        min_x, min_y, max_x, max_y = float("inf"), float("inf"), float("-inf"), float("-inf")

        def update_bounds(x: float, y: float) -> None:
            nonlocal min_x, min_y, max_x, max_y
            if math.isfinite(x) and math.isfinite(y):
                min_x, min_y, max_x, max_y = min(min_x, x), min(min_y, y), max(max_x, x), max(max_y, y)

        def collect_entity(ent: Any, ox: float = 0.0, oy: float = 0.0) -> None:
            et = ent.dxftype()
            try:
                if et == "LINE":
                    x0, y0 = float(ent.dxf.start.x) + ox, float(ent.dxf.start.y) + oy
                    x1, y1 = float(ent.dxf.end.x) + ox, float(ent.dxf.end.y) + oy
                    lines_data.append((x0, y0, x1, y1, "line"))
                    update_bounds(x0, y0)
                    update_bounds(x1, y1)
                elif et == "LWPOLYLINE":
                    pts = list(ent.get_points("xy"))
                    for i in range(len(pts) - 1):
                        x0, y0, x1, y1 = float(pts[i][0]) + ox, float(pts[i][1]) + oy, float(pts[i + 1][0]) + ox, float(pts[i + 1][1]) + oy
                        lines_data.append((x0, y0, x1, y1, "poly"))
                        update_bounds(x0, y0)
                        update_bounds(x1, y1)
                    if ent.closed and len(pts) >= 2:
                        lines_data.append((float(pts[-1][0]) + ox, float(pts[-1][1]) + oy, float(pts[0][0]) + ox, float(pts[0][1]) + oy, "poly"))
                elif et in ("CIRCLE", "ARC"):
                    cx, cy, r = float(ent.dxf.center.x) + ox, float(ent.dxf.center.y) + oy, float(ent.dxf.radius)
                    circles_data.append((cx, cy, r))
                    update_bounds(cx - r, cy - r)
                    update_bounds(cx + r, cy + r)
                elif et == "TEXT":
                    txt = str(ent.dxf.text).strip()
                    if txt:
                        tx, ty = float(ent.dxf.insert.x) + ox, float(ent.dxf.insert.y) + oy
                        th = float(getattr(ent.dxf, "height", 3.5))
                        raw_texts.append((txt, tx, ty, th))
                        update_bounds(tx, ty)
                elif et == "MTEXT":
                    txt = ent.plain_text().strip()
                    if txt:
                        tx, ty = float(ent.dxf.insert.x) + ox, float(ent.dxf.insert.y) + oy
                        th = float(getattr(ent.dxf, "char_height", 3.5))
                        for l_i, line_s in enumerate(txt.splitlines()):
                            if line_s.strip():
                                raw_texts.append((line_s.strip(), tx, ty - l_i * th * 1.4, th))
                                update_bounds(tx, ty)
                elif et == "INSERT":
                    ix, iy = float(ent.dxf.insert.x) + ox, float(ent.dxf.insert.y) + oy
                    update_bounds(ix, iy)
                    if hasattr(ent, "attribs") and ent.attribs:
                        ad: Dict[str, str] = {}
                        for att in ent.attribs:
                            val = str(att.dxf.text).strip()
                            if val:
                                ad[str(att.dxf.tag).upper().strip()] = val
                                raw_texts.append((val, float(att.dxf.insert.x), float(att.dxf.insert.y), 3.5))
                        p_attr = cls._part_from_dxf_attribs(ad, file_name, len(attrib_parts) + 1)
                        if p_attr:
                            attrib_parts.append(p_attr)
                    bname = ent.dxf.name
                    if bname in dxf_doc.blocks:
                        for sub_e in dxf_doc.blocks[bname]:
                            if sub_e.dxftype() in ("LINE", "LWPOLYLINE", "CIRCLE", "TEXT", "MTEXT"):
                                collect_entity(sub_e, ix, iy)
            except Exception:
                pass

        for e in msp:
            collect_entity(e)

        if not math.isfinite(min_x) or max_x <= min_x:
            min_x, min_y, max_x, max_y = 0.0, 0.0, 420.0, 297.0
        mx, my = max((max_x - min_x) * 0.04, 10.0), max((max_y - min_y) * 0.04, 10.0)
        min_x, min_y, max_x, max_y = min_x - mx, min_y - my, max_x + mx, max_y + my
        world_w, world_h = max(max_x - min_x, 1.0), max(max_y - min_y, 1.0)

        tokens: List[TextToken] = []
        for txt, tx, ty, th in raw_texts:
            sx0, sy1 = tx - min_x, max_y - ty
            sy0 = max(0.0, sy1 - max(th * 1.2, 4.0))
            sx1 = sx0 + max(len(txt) * th * 0.85, 8.0)
            tokens.append(TextToken(txt, sx0, sy0, sx1, sy1, 1))

        preview_img = cls._render_dxf_preview(world_w, world_h, min_x, max_y, lines_data, circles_data, tokens)
        page_data = DrawingPageData(page_no=1, width=world_w, height=world_h, preview_image=preview_img, tokens=tokens)

        all_parts.extend(attrib_parts)
        existing = {(p.part_number, p.part_name) for p in all_parts}
        for cp in cls.extract_parts_from_tokens(tokens, file_name, 1, (0.0, 0.0, world_w, world_h)):
            if (cp.part_number, cp.part_name) not in existing:
                all_parts.append(cp)

        title_meta = cls._detect_title_block_meta(tokens, world_w, world_h)
        page_data.title_meta = title_meta
        doc_model.drawing_title = title_meta.get("drawing_title", "")
        doc_model.drawing_number = title_meta.get("drawing_number", "")
        if not all_parts and tokens:
            single_part = cls._extract_single_part_fallback(tokens, title_meta, file_name, 1, world_w, world_h)
            if single_part:
                all_parts.append(single_part)
        doc_model.pages.append(page_data)
        doc_model.extracted_count = len(all_parts)
        return doc_model, all_parts

    @classmethod
    def _render_dxf_preview(
        cls, world_w: float, world_h: float, min_x: float, max_y: float,
        lines_data: List[Tuple[float, float, float, float, str]],
        circles_data: List[Tuple[float, float, float]],
        tokens: List[TextToken],
    ) -> Image.Image:
        target_w = 1600
        scale = target_w / max(world_w, 1.0)
        target_h = int(max(min(world_h * scale, 2200), 600))
        scale = min(target_w / max(world_w, 1.0), target_h / max(world_h, 1.0))

        img = Image.new("RGB", (target_w, target_h), (255, 255, 255))
        draw = ImageDraw.Draw(img)
        font_sm = None
        for fp in ["C:/Windows/Fonts/meiryo.ttc", "C:/Windows/Fonts/YuGothM.ttc", "C:/Windows/Fonts/msgothic.ttc"]:
            if os.path.exists(fp):
                try:
                    font_sm = ImageFont.truetype(fp, 13)
                    break
                except Exception:
                    pass

        for x0, y0, x1, y1, kind in lines_data:
            c = (30, 58, 138) if kind == "poly" else (15, 23, 42)
            draw.line([(int((x0 - min_x) * scale), int((max_y - y0) * scale)), (int((x1 - min_x) * scale), int((max_y - y1) * scale))], fill=c, width=2)

        for cx, cy, r in circles_data:
            px0, py0 = int((cx - r - min_x) * scale), int((max_y - (cy + r)) * scale)
            px1, py1 = int((cx + r - min_x) * scale), int((max_y - (cy - r)) * scale)
            if px1 > px0 and py1 > py0:
                draw.ellipse([px0, py0, px1, py1], outline=(37, 99, 235), width=2)

        for tk_item in tokens:
            draw.text((int(tk_item.x0 * scale), int(tk_item.y0 * scale)), tk_item.text, fill=(15, 23, 42), font=font_sm)
        return img

    @classmethod
    def _part_from_dxf_attribs(cls, attr_dict: Dict[str, str], source_file: str, idx: int) -> Optional[PartItem]:
        if not attr_dict:
            return None
        item = PartItem(item_no=str(idx), source_file=source_file, source_page=1)
        for tag, val in attr_dict.items():
            if any(k in tag for k in ["DWG", "PART_NO", "CODE", "図番", "品番"]):
                item.part_number = val
            elif any(k in tag for k in ["NAME", "DESC", "TITLE", "品名", "名称"]):
                item.part_name = val
            elif any(k in tag for k in ["SPEC", "SIZE", "DIM", "型式", "寸法"]):
                item.specification = val
            elif any(k in tag for k in ["MAT", "材質", "材料"]):
                item.material = val
            elif any(k in tag for k in ["FINISH", "TREAT", "表面処理"]):
                item.treatment = val
            elif any(k in tag for k in ["QTY", "COUNT", "数量"]):
                m = re.search(r"\d+(?:\.\d+)?", val)
                if m:
                    item.quantity = float(m.group(0))
        if not item.part_name and not item.part_number:
            return None
        if item.unit_weight == 0.0 and item.specification:
            item.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(item.specification, item.material)
        item.category = SmartPartAnalyzer.detect_category(item.part_name, item.specification, item.material, item.manufacturer)
        return item

    @classmethod
    def _extract_parts_from_2d_cells(
        cls, rows: List[List[Any]], source_file: str, page_no: int, table_bbox: Tuple[float, float, float, float]
    ) -> List[PartItem]:
        cleaned_rows = [[SmartPartAnalyzer.normalize_text(str(c) if c is not None else "") for c in r] for r in rows]
        cleaned_rows = [r for r in cleaned_rows if any(r)]
        if len(cleaned_rows) < 2:
            return []

        header_idx, col_map, bottom_up = -1, {}, False
        for idx, r in enumerate(cleaned_rows):
            cmap = SmartPartAnalyzer.classify_table_headers(r)
            if len(cmap) >= 2 and ("part_name" in cmap.values() or "part_number" in cmap.values()):
                header_idx, col_map = idx, cmap
                bottom_up = (idx == len(cleaned_rows) - 1)
                break

        results: List[PartItem] = []
        bx0, by0, bx1, by1 = table_bbox
        row_h = (by1 - by0) / max(len(cleaned_rows), 1)

        if header_idx != -1:
            indices = range(header_idx - 1, -1, -1) if bottom_up else range(header_idx + 1, len(cleaned_rows))
            for seq_idx, r_idx in enumerate(indices, start=1):
                rb = (bx0, by0 + r_idx * row_h, bx1, by0 + (r_idx + 1) * row_h)
                item = cls._map_row_with_col_map(cleaned_rows[r_idx], col_map, seq_idx, source_file, page_no, rb)
                if item:
                    results.append(item)
        else:
            for seq_idx, r in enumerate(cleaned_rows, start=1):
                rb = (bx0, by0 + (seq_idx - 1) * row_h, bx1, by0 + seq_idx * row_h)
                item = SmartPartAnalyzer.parse_unstructured_row(r, seq_idx, source_file, page_no, rb)
                if item:
                    results.append(item)
        return results

    @classmethod
    def _map_row_with_col_map(
        cls, row_cells: List[str], col_map: Dict[int, str], default_no: int,
        source_file: str, page_no: int, bbox: Tuple[float, float, float, float]
    ) -> Optional[PartItem]:
        if not any(row_cells):
            return None
        joined = " ".join(row_cells)
        if any(w in joined for w in ["部品表", "PARTS LIST", "BILL OF MATERIALS", "第三角法"]):
            return None

        item = PartItem(item_no=str(default_no), source_file=source_file, source_page=page_no, bbox=bbox)
        unmapped: List[str] = []

        for c_idx, val in enumerate(row_cells):
            val = val.strip()
            if not val:
                continue
            fn = col_map.get(c_idx)
            if fn == "item_no":
                item.item_no = val
            elif fn == "part_number":
                item.part_number = val
            elif fn == "part_name":
                item.part_name = val
            elif fn == "specification":
                item.specification = val
            elif fn == "material":
                mat, _ = SmartPartAnalyzer.detect_material(val)
                item.material = mat if mat else val
            elif fn == "treatment":
                item.treatment = val
            elif fn == "quantity":
                m = re.search(r"\d+(?:\.\d+)?", val)
                if m:
                    item.quantity = float(m.group(0))
            elif fn == "unit":
                item.unit = val
            elif fn == "unit_weight":
                m = re.search(r"\d+(?:\.\d+)?", val)
                if m:
                    item.unit_weight = float(m.group(0))
            elif fn == "manufacturer":
                item.manufacturer = val
            elif fn == "remarks":
                item.remarks = val
            else:
                unmapped.append(val)

        for ex in unmapped:
            if not item.material:
                mat, _ = SmartPartAnalyzer.detect_material(ex)
                if mat:
                    item.material = mat
                    continue
            if not item.remarks:
                item.remarks = ex

        if not item.part_name and not item.part_number:
            return None
        if item.material and not item.treatment and item.material in MATERIAL_MASTER:
            item.treatment = MATERIAL_MASTER[item.material]["treat"]
        if item.unit_weight == 0.0 and item.specification:
            item.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(item.specification, item.material)
        item.category = SmartPartAnalyzer.detect_category(item.part_name, item.specification, item.material, item.manufacturer)
        return item

    @classmethod
    def extract_parts_from_tokens(
        cls, tokens: List[TextToken], source_file: str, page_no: int,
        area_rect: Tuple[float, float, float, float], force_all_rows: bool = False
    ) -> List[PartItem]:
        ax0, ay0, ax1, ay1 = area_rect
        filtered = [t for t in tokens if ax0 <= t.cx <= ax1 and ay0 <= t.cy <= ay1]
        if not filtered:
            return []

        sorted_by_y = sorted(filtered, key=lambda t: t.cy)
        avg_h = sum(max(t.y1 - t.y0, 4.0) for t in sorted_by_y) / len(sorted_by_y)
        y_tol = max(avg_h * 0.65, 3.5)

        clusters: List[List[TextToken]] = []
        for tok in sorted_by_y:
            if not clusters or abs(tok.cy - sum(t.cy for t in clusters[-1]) / len(clusters[-1])) > y_tol:
                clusters.append([tok])
            else:
                clusters[-1].append(tok)

        structured: List[Tuple[List[str], List[float], Tuple[float, float, float, float]]] = []
        for cl in clusters:
            cl_s = sorted(cl, key=lambda t: t.x0)
            cells, cell_xs = [], []
            cur_t, cur_x0, prev_x1 = cl_s[0].text, cl_s[0].x0, cl_s[0].x1
            for tok in cl_s[1:]:
                if tok.x0 - prev_x1 <= max(avg_h * 0.8, 5.0):
                    cur_t = f"{cur_t} {tok.text}"
                    prev_x1 = max(prev_x1, tok.x1)
                else:
                    cells.append(cur_t)
                    cell_xs.append(cur_x0)
                    cur_t, cur_x0, prev_x1 = tok.text, tok.x0, tok.x1
            cells.append(cur_t)
            cell_xs.append(cur_x0)
            structured.append((cells, cell_xs, (min(t.x0 for t in cl_s), min(t.y0 for t in cl_s), max(t.x1 for t in cl_s), max(t.y1 for t in cl_s))))

        h_idx, h_map, h_xs, bottom_up = -1, {}, [], False
        for r_idx, (cells, c_xs, _) in enumerate(structured):
            cmap = SmartPartAnalyzer.classify_table_headers(cells)
            if len(cmap) >= 2 and ("part_name" in cmap.values() or "part_number" in cmap.values()):
                h_idx, h_map, h_xs = r_idx, cmap, c_xs
                bottom_up = (r_idx >= len(structured) - 2 and r_idx >= 2)
                break

        extracted: List[PartItem] = []
        if h_idx != -1 and h_xs:
            indices = range(h_idx - 1, -1, -1) if bottom_up else range(h_idx + 1, len(structured))
            for seq_no, idx in enumerate(indices, start=1):
                cells, c_xs, r_bbox = structured[idx]
                if not force_all_rows and max(c_xs) < min(h_xs) - 60.0:
                    continue
                aligned = [""] * len(h_xs)
                for cs, cx in zip(cells, c_xs):
                    best_c = min(range(len(h_xs)), key=lambda hi: abs(cx - h_xs[hi]))
                    aligned[best_c] = f"{aligned[best_c]} {cs}".strip()
                item = cls._map_row_with_col_map(aligned, h_map, seq_no, source_file, page_no, r_bbox)
                if item:
                    extracted.append(item)
        else:
            seq_no = 1
            for cells, _, r_bbox in structured:
                if not force_all_rows and len(cells) < 2:
                    continue
                item = SmartPartAnalyzer.parse_unstructured_row(cells, seq_no, source_file, page_no, r_bbox)
                if item and (force_all_rows or item.material or item.part_number or len(cells) >= 3):
                    extracted.append(item)
                    seq_no += 1
        return extracted

    @classmethod
    def _detect_title_block_meta(cls, tokens: List[TextToken], width: float, height: float) -> Dict[str, str]:
        meta = {"drawing_title": "", "drawing_number": "", "material": "", "treatment": ""}
        tb_tokens = [t for t in tokens if t.cx >= width * 0.45 and t.cy >= height * 0.65] or tokens
        for i, tok in enumerate(tb_tokens):
            txt = SmartPartAnalyzer.normalize_text(tok.text)
            mat, _ = SmartPartAnalyzer.detect_material(txt)
            if mat and not meta["material"]:
                meta["material"] = mat
            treat = SmartPartAnalyzer.detect_treatment(txt)
            if treat and not meta["treatment"]:
                meta["treatment"] = treat
            if re.fullmatch(r"[A-Z]{2,5}-\d{3,6}(?:-\d{1,3})?", txt) and not meta["drawing_number"]:
                meta["drawing_number"] = txt
            if txt in ("図面名称", "名称", "品名", "TITLE") and i + 1 < len(tb_tokens):
                cand = SmartPartAnalyzer.normalize_text(tb_tokens[i + 1].text)
                if len(cand) >= 2 and cand not in ("図番", "尺度", "材質", "数量"):
                    meta["drawing_title"] = cand
        return meta

    @classmethod
    def _extract_single_part_fallback(
        cls,
        tokens: List[TextToken],
        title_meta: Dict[str, str],
        source_file: str,
        page_no: int,
        width: float,
        height: float,
    ) -> Optional[PartItem]:
        """部品表テーブルを持たない単品図（1部品1図面）から表題欄と材料寸法を読み取って1件の部品を構築する"""
        p_name = title_meta.get("drawing_title", "").strip()
        p_num = title_meta.get("drawing_number", "").strip()
        mat = title_meta.get("material", "").strip()
        treat = title_meta.get("treatment", "").strip()
        spec = ""
        bbox = (width * 0.45, height * 0.65, width * 0.98, height * 0.98)

        for tok in tokens:
            norm = SmartPartAnalyzer.normalize_text(tok.text)
            if not mat:
                m_cand, _ = SmartPartAnalyzer.detect_material(norm)
                if m_cand:
                    mat = m_cand
            if not treat:
                t_cand = SmartPartAnalyzer.detect_treatment(norm)
                if t_cand:
                    treat = t_cand
            if not spec and re.search(r"([tT]\d+(?:\.\d+)?\s*[xX]\s*\d+|[φD]\d+(?:\.\d+)?\s*[xX]\s*\d+)", norm):
                spec = norm
                bbox = (tok.x0, tok.y0, tok.x1, tok.y1)

        if not (p_name or p_num or (mat and spec)):
            return None
        if not p_name:
            p_name = os.path.splitext(source_file)[0]
        if mat and not treat and mat in MATERIAL_MASTER:
            treat = MATERIAL_MASTER[mat]["treat"]
        u_wt = SmartPartAnalyzer.estimate_weight_from_spec(spec, mat) if spec else 0.0
        cat = SmartPartAnalyzer.detect_category(p_name, spec, mat, "")
        return PartItem(
            item_no="1",
            part_number=p_num,
            part_name=p_name,
            specification=spec,
            material=mat,
            treatment=treat,
            quantity=1.0,
            unit="個",
            unit_weight=u_wt,
            category=cat,
            source_file=source_file,
            source_page=page_no,
            remarks="単品図自動抽出",
            bbox=bbox,
        )


# ==============================================================================
# サンプル2D図面 (PDF / DXF) 自動生成エンジン
# ==============================================================================

class SampleDrawingGenerator:
    SAMPLE_BOM_ROWS_PDF = [
        ["No.", "図番 / 品番", "部品名称", "型式・材料寸法", "材質", "表面処理", "数量", "単位", "単重(kg)", "備考"],
        ["1", "CV-1001-01", "メインベースフレーム", "t12x350x600", "SS400", "メラミン焼付塗装", "1", "個", "19.78", "製缶溶接構造"],
        ["2", "CV-1001-02", "ドライブシャフト", "φ30x420", "S45C", "無電解ニッケルめっき", "1", "本", "2.33", "キー溝加工あり"],
        ["3", "CV-1001-03", "従動シャフト", "φ25x400", "S45C", "無電解ニッケルめっき", "1", "本", "1.54", ""],
        ["4", "CV-1001-04", "サイドプレート(L)", "t10x180x550", "A5052", "白アルマイト", "1", "枚", "2.65", "マシニング加工"],
        ["5", "CV-1001-05", "サイドプレート(R)", "t10x180x550", "A5052", "白アルマイト", "1", "枚", "2.65", "マシニング加工"],
        ["6", "CV-1001-06", "ガイドレール", "t8x40x500", "MCナイロン", "生地", "2", "本", "0.19", "摺動ガイド"],
        ["7", "CV-1001-07", "センサブラケット", "t2.0x60x120", "SUS304", "生地(酸洗)", "2", "個", "0.11", "板金曲げ"],
        ["8", "CV-1001-08", "ディスタンスカラー", "φ35xt5x45", "SUS303", "生地", "4", "個", "0.17", "旋盤加工"],
        ["9", "UCFL206", "ひしフランジ形軸受ユニット", "φ30用 UCFL206", "FC200", "塗装", "2", "個", "0.90", "NTN製"],
        ["10", "6005ZZ", "深溝玉軸受(ベアリング)", "内径25x外径47x幅12", "SUJ2", "焼入れ", "2", "個", "0.08", "NSK製"],
        ["11", "BLM5120-15S", "ブラシレスモータ&ギヤヘッド", "120W 減速比1:15", "-", "-", "1", "式", "3.20", "オリエンタルモーター"],
        ["12", "CB-M8-25", "六角穴付ボルト", "M8x25", "SCM435", "四三酸化鉄皮膜", "16", "本", "0.015", "標準部品"],
    ]

    SAMPLE_BOM_ROWS_DXF = [
        ["1", "BRK-200-01", "マウントベース", "t16x200x280", "S45C", "四三酸化鉄皮膜", "1", "7.03", "フライス6面加工"],
        ["2", "BRK-200-02", "クランプアーム", "t15x80x160", "SCM435", "無電解ニッケルめっき", "2", "1.51", "調質後加工"],
        ["3", "BRK-200-03", "ヒンジピン", "φ16x110", "SUS304", "焼入れ研磨", "2", "0.18", "h7公差"],
        ["4", "BRK-200-04", "樹脂パッド", "t10x50x80", "POM", "生地", "2", "0.056", "ワーク当たり面"],
        ["5", "CDQ2B32-25DZ", "薄形エアシリンダ", "φ32 ストローク25", "A6063", "アルマイト", "1", "0.38", "SMC製"],
        ["6", "CB-M10-30", "六角穴付ボルト", "M10x30", "SCM435", "四三酸化鉄皮膜", "4", "0.028", "標準部品"],
    ]

    @classmethod
    def create_sample_pdf(cls, output_path: str) -> str:
        doc = fitz.open()
        page = doc.new_page(width=1190, height=842)
        shape = page.new_shape()

        shape.draw_rect(fitz.Rect(20, 20, 1170, 822))
        shape.finish(color=(0.05, 0.1, 0.25), width=2.0)
        shape.draw_rect(fitz.Rect(28, 28, 1162, 814))
        shape.finish(color=(0.05, 0.1, 0.25), width=1.0)

        # 2D組立図（正面図・平面図）
        for r in [(90, 440, 560, 470), (110, 470, 145, 680), (505, 470, 540, 680), (100, 320, 550, 440), (100, 110, 550, 250), (135, 70, 175, 290), (475, 95, 515, 265), (115, 45, 220, 70)]:
            shape.draw_rect(fitz.Rect(*r))
        shape.finish(color=(0.1, 0.2, 0.45), width=1.5)

        for cx, cy, rad in [(155, 380, 42), (155, 380, 16), (495, 380, 38), (495, 380, 14)]:
            shape.draw_circle(fitz.Point(cx, cy), rad)
        shape.finish(color=(0.05, 0.35, 0.75), width=1.8)

        # 部品表グリッド
        tx0, ty0, tx1, row_h = 595, 45, 1152, 26
        col_widths = [32, 82, 115, 92, 55, 88, 32, 30, 31]
        num_rows = len(cls.SAMPLE_BOM_ROWS_PDF)
        ty1 = ty0 + num_rows * row_h

        shape.draw_rect(fitz.Rect(tx0, ty0, tx1, ty0 + row_h))
        shape.finish(color=(0.1, 0.2, 0.4), fill=(0.88, 0.93, 0.98), width=1.2)

        for r_i in range(num_rows + 1):
            shape.draw_line(fitz.Point(tx0, ty0 + r_i * row_h), fitz.Point(tx1, ty0 + r_i * row_h))
        cur_x = tx0
        col_xs = [cur_x]
        for cw in col_widths[:-1]:
            cur_x += cw
            col_xs.append(cur_x)
            shape.draw_line(fitz.Point(cur_x, ty0), fitz.Point(cur_x, ty1))
        col_xs.append(tx1)
        shape.draw_rect(fitz.Rect(tx0, ty0, tx1, ty1))
        shape.draw_rect(fitz.Rect(720, 700, 1152, 806))
        shape.finish(color=(0.1, 0.15, 0.25), width=1.0)
        shape.commit()

        fn = "japan"
        page.insert_text(fitz.Point(240, 298), "【平面図 (TOP VIEW)】", fontname=fn, fontsize=11, color=(0.1, 0.2, 0.4))
        page.insert_text(fitz.Point(210, 725), "【正面図 (FRONT VIEW) - 全長 600mm × 全高 420mm】", fontname=fn, fontsize=11, color=(0.1, 0.2, 0.4))

        for r_idx, r_vals in enumerate(cls.SAMPLE_BOM_ROWS_PDF):
            y_t = ty0 + r_idx * row_h + 17
            disp = [r_vals[0], r_vals[1], r_vals[2], r_vals[3], r_vals[4], r_vals[5], r_vals[6], r_vals[8], r_vals[9]]
            for c_idx, cell_s in enumerate(disp):
                page.insert_text(fitz.Point(col_xs[c_idx] + 4, y_t), cell_s, fontname=fn, fontsize=8.2, color=(0.05, 0.08, 0.15))

        page.insert_text(fitz.Point(735, 725), "図面名称", fontname=fn, fontsize=10)
        page.insert_text(fitz.Point(850, 725), "ベルトコンベア駆動ユニット組立図", fontname=fn, fontsize=12)
        page.insert_text(fitz.Point(735, 760), "図面番号", fontname=fn, fontsize=10)
        page.insert_text(fitz.Point(850, 760), "CV-1001-ASM", fontname=fn, fontsize=11)
        doc.save(output_path)
        doc.close()
        return output_path

    @classmethod
    def create_sample_dxf(cls, output_path: str) -> str:
        doc = ezdxf.new("R2010")
        msp = doc.modelspace()
        msp.add_lwpolyline([(0, 0), (420, 0), (420, 297), (0, 297), (0, 0)], close=True)
        msp.add_lwpolyline([(5, 5), (415, 5), (415, 292), (5, 292), (5, 5)], close=True)
        msp.add_lwpolyline([(30, 60), (190, 60), (190, 85), (30, 85), (30, 60)], close=True)
        msp.add_lwpolyline([(60, 85), (95, 85), (95, 210), (60, 210), (60, 85)], close=True)
        msp.add_lwpolyline([(125, 85), (160, 85), (160, 210), (125, 210), (125, 85)], close=True)
        msp.add_circle((77.5, 175), 12)
        msp.add_circle((142.5, 175), 12)

        headers = ["No", "図番", "品名", "寸法", "材質", "表面処理", "数量", "単重", "備考"]
        col_xs = [208, 220, 255, 295, 330, 355, 385, 396, 405]
        top_y = 278.0
        for c_x, h_txt in zip(col_xs, headers):
            t_e = msp.add_text(h_txt, dxfattribs={"height": 3.0})
            t_e.dxf.insert = (c_x, top_y - 3)
        for r_idx, r_vals in enumerate(cls.SAMPLE_BOM_ROWS_DXF, start=1):
            ry = top_y - 5 - r_idx * 10.0
            for c_x, val_s in zip(col_xs, r_vals):
                t_e = msp.add_text(val_s, dxfattribs={"height": 2.8})
                t_e.dxf.insert = (c_x, ry + 2)
        t1 = msp.add_text("図面名称", dxfattribs={"height": 3.2})
        t1.dxf.insert = (265, 32)
        t2 = msp.add_text("クランプブラケットユニット組図", dxfattribs={"height": 4.0})
        t2.dxf.insert = (300, 32)
        t3 = msp.add_text("BRK-200-ASM", dxfattribs={"height": 3.5})
        t3.dxf.insert = (300, 15)
        doc.saveas(output_path)
        return output_path


# ==============================================================================
# Excel部品表 (BOM) 出力エンジン
# ==============================================================================

@dataclass
class ExcelExportConfig:
    template_type: str = "standard"
    split_by_drawing: bool = True
    merge_duplicates: bool = False
    include_preview_sheet: bool = True
    project_name: str = ""
    drawing_no: str = ""
    author_name: str = ""


class ExcelBOMExporter:
    @classmethod
    def export_to_excel(
        cls, output_path: str, parts: List[PartItem], documents: List[DrawingDocument], config: ExcelExportConfig
    ) -> None:
        wb = openpyxl.Workbook()
        ws_main = wb.active
        ws_main.title = "部品表(BOM)"

        target_parts = cls.merge_duplicate_parts(copy.deepcopy(parts)) if config.merge_duplicates else copy.deepcopy(parts)
        cls._populate_bom_worksheet(ws_main, target_parts, config, "総合部品表 (ALL PARTS)")

        # フォーマット種別に応じた専用手配シートの自動作成
        if config.template_type == "purchasing":
            purch_parts = [copy.deepcopy(p) for p in target_parts if p.category in ("購入品", "標準部品", "支給品")]
            if purch_parts:
                for idx_p, pp in enumerate(purch_parts, start=1):
                    pp.item_no = str(idx_p)
                ws_p = wb.create_sheet(title="購買・標準品手配表")
                cls._populate_bom_worksheet(ws_p, purch_parts, config, "購入品・標準部品 集計リスト")
        elif config.template_type == "manufacturing":
            mfg_parts = [copy.deepcopy(p) for p in target_parts if p.category in ("加工品", "板金品", "製缶・溶接品", "組立品")]
            if mfg_parts:
                for idx_m, mp in enumerate(mfg_parts, start=1):
                    mp.item_no = str(idx_m)
                ws_m = wb.create_sheet(title="加工・製作手配表")
                cls._populate_bom_worksheet(ws_m, mfg_parts, config, "加工品・板金・製缶 製作リスト")

        if config.split_by_drawing and len(documents) >= 1:
            grouped: Dict[str, List[PartItem]] = {}
            for p in parts:
                grouped.setdefault(p.source_file or "手入力部品", []).append(p)
            if len(grouped) > 1:
                for idx, (src_name, sub_parts) in enumerate(grouped.items(), start=1):
                    safe = re.sub(r"[\\/*?:\[\]]", "_", os.path.splitext(src_name)[0])[:22]
                    base_title = f"{idx}_{safe}"
                    sheet_title = base_title[:31]
                    suffix = 2
                    while sheet_title in wb.sheetnames:
                        sheet_title = f"{base_title[:27]}_{suffix}"
                        suffix += 1
                    ws_sub = wb.create_sheet(title=sheet_title)
                    cls._populate_bom_worksheet(ws_sub, copy.deepcopy(sub_parts), config, f"図面別部品表: {src_name}")

        temp_imgs: List[str] = []
        if config.include_preview_sheet and documents:
            ws_prev = wb.create_sheet(title="図面プレビュー")
            ws_prev["A1"] = "収録2D図面プレビュー一覧"
            ws_prev["A1"].font = Font(name="Meiryo UI", size=14, bold=True, color="1E3A8A")
            cur_row = 3
            for doc in documents:
                for pg in doc.pages:
                    if pg.preview_image is not None:
                        ws_prev.cell(row=cur_row, column=1, value=f"■ {doc.file_name} (Page {pg.page_no}) {doc.drawing_title}").font = Font(name="Meiryo UI", size=11, bold=True)
                        cur_row += 1
                        try:
                            thumb = pg.preview_image.copy()
                            thumb.thumbnail((900, 600), Image.Resampling.LANCZOS)
                            fd, tp = tempfile.mkstemp(suffix=".png")
                            os.close(fd)
                            thumb.save(tp, format="PNG")
                            temp_imgs.append(tp)
                            ws_prev.add_image(OpenpyxlImage(tp), f"A{cur_row}")
                            cur_row += int(thumb.height / 19) + 3
                        except Exception:
                            cur_row += 2
        try:
            wb.save(output_path)
        finally:
            for tp in temp_imgs:
                if os.path.exists(tp):
                    try:
                        os.remove(tp)
                    except Exception:
                        pass

    @classmethod
    def merge_duplicate_parts(cls, parts: List[PartItem]) -> List[PartItem]:
        merged: Dict[Tuple[str, str, str, str], PartItem] = {}
        for p in parts:
            key = (p.part_number.strip().upper(), p.part_name.strip(), p.specification.strip().upper(), p.material.strip().upper())
            if key in merged:
                ex = merged[key]
                ex.quantity += p.quantity
                if p.source_file and p.source_file not in ex.source_file:
                    ex.source_file = f"{ex.source_file}, {p.source_file}"
            else:
                merged[key] = copy.deepcopy(p)
        res = list(merged.values())
        for idx, item in enumerate(res, start=1):
            item.item_no = str(idx)
        return res

    @classmethod
    def _populate_bom_worksheet(cls, ws: Any, parts: List[PartItem], config: ExcelExportConfig, subtitle: str) -> None:
        ff = "Meiryo UI"
        titles = {
            "standard": "部 品 表 (BILL OF MATERIALS)",
            "purchasing": "購買・手配部品リスト (PURCHASING BOM)",
            "manufacturing": "加工・製造部品リスト (MANUFACTURING BOM)",
        }
        thin = Side(style="thin", color="CBD5E1")
        med = Side(style="medium", color="1E3A8A")
        dbl = Side(style="double", color="1E3A8A")
        b_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
        b_head = Border(left=thin, right=thin, top=med, bottom=med)
        b_tot = Border(left=thin, right=thin, top=med, bottom=dbl)

        total_cols = 15
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
        tc = ws.cell(row=1, column=1, value=titles.get(config.template_type, titles["standard"]))
        tc.font = Font(name=ff, size=15, bold=True, color="FFFFFF")
        tc.fill = PatternFill("solid", fgColor="1E3A8A")
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        today_s = datetime.date.today().strftime("%Y-%m-%d")
        metas = [
            (2, 1, "案件・製品名", 2, 2, config.project_name or "未設定", 4),
            (2, 5, "総合図番", 2, 6, config.drawing_no or "-", 8),
            (2, 9, "作成者", 2, 10, config.author_name or "-", 11),
            (2, 12, "出力日", 2, 13, today_s, 15),
            (3, 1, "シート区分", 3, 2, subtitle, 8),
            (3, 9, "総部品行数", 3, 10, f"{len(parts)} 件", 11),
            (3, 12, "システム", 3, 13, APP_NAME, 15),
        ]
        for r_i in (2, 3):
            ws.row_dimensions[r_i].height = 20
            for c_i in range(1, total_cols + 1):
                ws.cell(row=r_i, column=c_i).border = b_cell
        for lr, lc, lt, vr, vc, vt, vce in metas:
            lc_cell = ws.cell(row=lr, column=lc, value=lt)
            lc_cell.font = Font(name=ff, size=9, bold=True, color="1E293B")
            lc_cell.fill = PatternFill("solid", fgColor="E2E8F0")
            lc_cell.alignment = Alignment(horizontal="center", vertical="center")
            if vce > vc:
                ws.merge_cells(start_row=vr, start_column=vc, end_row=vr, end_column=vce)
            vc_cell = ws.cell(row=vr, column=vc, value=vt)
            vc_cell.font = Font(name=ff, size=10, color="0F172A")
            vc_cell.alignment = Alignment(horizontal="left", vertical="center")

        headers = [
            ("No.", 6, "center"), ("図番 / 品番", 18, "left"), ("品名 / 部品名称", 24, "left"),
            ("型式・寸法・仕様", 22, "left"), ("材質", 13, "center"), ("表面処理・熱処理", 18, "left"),
            ("数量", 9, "right"), ("単位", 7, "center"), ("単重(kg)", 11, "right"),
            ("総重量(kg)", 12, "right"), ("調達区分", 12, "center"), ("メーカー", 16, "left"),
            ("単価(円)", 12, "right"), ("金額(円)", 14, "right"), ("抽出元図面 / 備考", 26, "left"),
        ]
        h_row = 5
        ws.row_dimensions[h_row].height = 25
        for c_idx, (ht, cw, _) in enumerate(headers, start=1):
            c = ws.cell(row=h_row, column=c_idx, value=ht)
            c.font = Font(name=ff, size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1E40AF")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = b_head
            ws.column_dimensions[get_column_letter(c_idx)].width = cw

        s_row = 6
        for idx, p in enumerate(parts):
            r = s_row + idx
            ws.row_dimensions[r].height = 22
            rf = PatternFill("solid", fgColor="FFFFFF" if idx % 2 == 0 else "F8FAFC")
            cf = PatternFill("solid", fgColor="EFF6FF")
            rem = f"[{p.source_file}] {p.remarks}".strip() if p.source_file else p.remarks
            vals = [
                int(p.item_no) if p.item_no.isdigit() else p.item_no,
                p.part_number, p.part_name, p.specification, p.material, p.treatment,
                p.quantity, p.unit, p.unit_weight, f"=G{r}*I{r}",
                p.category, p.manufacturer, p.unit_price, f"=G{r}*M{r}", rem,
            ]
            for c_idx, val in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c_idx, value=val)
                cell.font = Font(name=ff, size=10, bold=(c_idx in (3, 10, 14)), color="0F172A")
                cell.fill = cf if c_idx in (10, 14) else rf
                cell.border = b_cell
                cell.alignment = Alignment(horizontal=headers[c_idx - 1][2], vertical="center")
                if c_idx == 7:
                    cell.number_format = "#,##0.##"
                elif c_idx in (9, 10):
                    cell.number_format = "0.000"
                elif c_idx in (13, 14):
                    cell.number_format = "#,##0"

        e_row = max(s_row + len(parts) - 1, s_row)
        t_row = s_row + len(parts)
        ws.row_dimensions[t_row].height = 25
        for c_idx in range(1, total_cols + 1):
            c = ws.cell(row=t_row, column=c_idx, value="")
            c.font = Font(name=ff, size=10, bold=True, color="1E3A8A")
            c.fill = PatternFill("solid", fgColor="DBEAFE")
            c.border = b_tot
            c.alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=t_row, column=3, value="合 計 (TOTAL)").alignment = Alignment(horizontal="center", vertical="center")
        if parts:
            ws.cell(row=t_row, column=7, value=f"=SUM(G{s_row}:G{e_row})").number_format = "#,##0.##"
            ws.cell(row=t_row, column=10, value=f"=SUM(J{s_row}:J{e_row})").number_format = "0.000"
            ws.cell(row=t_row, column=14, value=f"=SUM(N{s_row}:N{e_row})").number_format = "#,##0"

        ws.auto_filter.ref = f"A{h_row}:O{t_row}"
        ws.freeze_panes = f"A{s_row}"
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4


# ==============================================================================
# ダイアログ群 (Gemini AI設定 / 重量計算機 / Excel出力設定 / 使い方ガイド)
# ==============================================================================

class GeminiSettingsDialog(tk.Toplevel):
    """Gemini APIキー、最新モデル一覧選択・最適モデル自動選定、AI解析オプション設定ダイアログ"""

    def __init__(self, parent: tk.Misc, config: GeminiConfig, on_save_callback: Any):
        super().__init__(parent)
        self.title(f"Gemini AI 高精度解析・モデル選択設定 - {APP_NAME}")
        self.geometry("820x700")
        self.minsize(720, 610)
        self.configure(bg=LightPalette.BG_MAIN)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        enforce_windows_light_mode(self)

        self.config_obj = config
        self.on_save_callback = on_save_callback
        self.model_entries: List[Dict[str, Any]] = self._build_initial_model_entries()
        self._is_fetching = False
        self._is_testing = False
        self._syncing_selection = False
        self._build_ui()
        self._populate_model_tree()
        self._sync_tree_and_detail_from_model_var()

    def _build_initial_model_entries(self) -> List[Dict[str, Any]]:
        if self.config_obj.cached_models:
            return GeminiVisionAnalyzer.merge_catalog_with_api_models(self.config_obj.cached_models)
        return copy.deepcopy(GEMINI_MODEL_CATALOG)

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=LightPalette.BTN_AI_BG, padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text="Gemini API マルチモーダル図面解析設定（最新モデル選択対応）",
            bg=LightPalette.BTN_AI_BG,
            fg="#FFFFFF",
            font=("Meiryo UI", 11, "bold"),
        ).pack(anchor="w")

        body = tk.Frame(self, bg=LightPalette.BG_PANEL, padx=18, pady=12)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        # 1. APIキー
        tk.Label(body, text="1. Gemini APIキー (Google AI Studio):", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9, "bold")).pack(anchor="w")
        key_row = tk.Frame(body, bg=LightPalette.BG_PANEL)
        key_row.pack(fill=tk.X, pady=(4, 4))

        self.api_key_var = tk.StringVar(value=self.config_obj.api_key)
        self.key_entry = ttk.Entry(key_row, textvariable=self.api_key_var, show="●", width=52)
        self.key_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.show_key = False
        tk.Button(
            key_row, text="表示/隠す", command=self._toggle_key_vis,
            bg=LightPalette.BTN_DEFAULT_BG, fg=LightPalette.FG_MAIN, relief=tk.SOLID, bd=1, padx=8, font=("Meiryo UI", 8)
        ).pack(side=tk.LEFT, padx=(6, 0))

        self.save_key_var = tk.BooleanVar(value=self.config_obj.save_api_key)
        ttk.Checkbutton(body, text="APIキーをPC内の設定ファイルに保存する（次回から自動読込）", variable=self.save_key_var).pack(anchor="w", pady=(0, 8))

        # 2. モデル選択セクション（最新一覧＋最適選択＋ライブ取得）
        tk.Label(
            body,
            text="2. 使用する Gemini モデル（一覧クリックで選択 / APIからリアルタイム更新可）:",
            bg=LightPalette.BG_PANEL,
            fg=LightPalette.FG_MAIN,
            font=("Meiryo UI", 9, "bold"),
        ).pack(anchor="w")

        model_ctrl = tk.Frame(body, bg=LightPalette.BG_PANEL)
        model_ctrl.pack(fill=tk.X, pady=(4, 6))

        tk.Label(model_ctrl, text="選択モデルID:", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_SUB, font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT)
        self.model_var = tk.StringVar(value=strip_gemini_model_prefix(self.config_obj.model_name) or DEFAULT_OPTIMAL_GEMINI_MODEL)
        self.model_combo = ttk.Combobox(
            model_ctrl,
            textvariable=self.model_var,
            values=[m["id"] for m in self.model_entries],
            width=26,
        )
        self.model_combo.pack(side=tk.LEFT, padx=(4, 8))
        self.model_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_tree_and_detail_from_model_var())
        self.model_var.trace_add("write", lambda *_: self._on_model_var_edited())

        tk.Button(
            model_ctrl, text="★最適モデル(バランス)", command=lambda: self._apply_optimal_model("balanced"),
            bg="#EDE9FE", fg="#5B21B6", relief=tk.SOLID, bd=1, padx=7, pady=2, font=("Meiryo UI", 8, "bold"), cursor="hand2"
        ).pack(side=tk.LEFT, padx=2)
        tk.Button(
            model_ctrl, text="精度優先", command=lambda: self._apply_optimal_model("accuracy"),
            bg="#EFF6FF", fg="#1D4ED8", relief=tk.SOLID, bd=1, padx=7, pady=2, font=("Meiryo UI", 8), cursor="hand2"
        ).pack(side=tk.LEFT, padx=2)
        tk.Button(
            model_ctrl, text="高速・省コスト", command=lambda: self._apply_optimal_model("speed"),
            bg="#ECFDF5", fg="#047857", relief=tk.SOLID, bd=1, padx=7, pady=2, font=("Meiryo UI", 8), cursor="hand2"
        ).pack(side=tk.LEFT, padx=2)

        self.btn_fetch_models = tk.Button(
            model_ctrl,
            text="APIから最新モデル一覧を取得",
            command=self._fetch_models_async,
            bg="#2563EB",
            fg="#FFFFFF",
            relief=tk.FLAT,
            padx=9,
            pady=3,
            font=("Meiryo UI", 8, "bold"),
            cursor="hand2",
        )
        self.btn_fetch_models.pack(side=tk.RIGHT)

        # モデル一覧 Treeview
        tree_wrap = tk.Frame(body, bg=LightPalette.BG_PANEL, bd=1, relief=tk.SOLID)
        tree_wrap.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        cols = ("tier", "model_id", "display_name", "speed", "accuracy", "pricing", "source")
        self.model_tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", height=7, selectmode="browse")
        col_specs = [
            ("tier", "推奨 / 区分", 130, "w"),
            ("model_id", "モデルID", 160, "w"),
            ("display_name", "モデル名", 135, "w"),
            ("speed", "速度", 78, "center"),
            ("accuracy", "図面精度", 82, "center"),
            ("pricing", "料金・無料枠目安", 160, "w"),
            ("source", "取得元", 92, "center"),
        ]
        for cid, heading, w, anchor in col_specs:
            self.model_tree.heading(cid, text=heading)
            self.model_tree.column(cid, width=w, minwidth=65, anchor=anchor)

        tree_sb = ttk.Scrollbar(tree_wrap, orient=tk.VERTICAL, command=self.model_tree.yview)
        self.model_tree.configure(yscrollcommand=tree_sb.set)
        self.model_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.model_tree.tag_configure("recommended", background="#F5F3FF", foreground="#0F172A")
        self.model_tree.tag_configure("normal", background="#FFFFFF", foreground="#0F172A")
        self.model_tree.tag_configure("legacy", background="#F8FAFC", foreground="#64748B")

        self.model_tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.model_tree.bind("<Double-1>", self._on_tree_double_click)

        # 選択中モデルの詳細ボックス
        detail_box = tk.Frame(body, bg=LightPalette.BG_SUB, bd=1, relief=tk.SOLID, padx=10, pady=6)
        detail_box.pack(fill=tk.X, pady=(0, 8))
        self.model_detail_title_lbl = tk.Label(
            detail_box,
            text="",
            bg=LightPalette.BG_SUB,
            fg=LightPalette.FG_ACCENT,
            font=("Meiryo UI", 8, "bold"),
            anchor="w",
        )
        self.model_detail_title_lbl.pack(fill=tk.X)
        self.model_detail_desc_lbl = tk.Label(
            detail_box,
            text="",
            bg=LightPalette.BG_SUB,
            fg=LightPalette.FG_MAIN,
            font=("Meiryo UI", 8),
            anchor="w",
            justify="left",
            wraplength=750,
        )
        self.model_detail_desc_lbl.pack(fill=tk.X, pady=(2, 0))

        # 3. 動作オプション
        self.auto_load_var = tk.BooleanVar(value=self.config_obj.auto_use_on_load)
        ttk.Checkbutton(body, text="図面ファイルを開いたときに自動で Gemini AI 解析を実行する", variable=self.auto_load_var).pack(anchor="w", pady=1)

        self.area_ai_var = tk.BooleanVar(value=self.config_obj.use_for_area_select)
        ttk.Checkbutton(body, text="マウス範囲選択抽出時に選択エリアの画像を Gemini AI で高精度解析する", variable=self.area_ai_var).pack(anchor="w", pady=1)

        self.replace_var = tk.BooleanVar(value=self.config_obj.replace_existing_on_ai)
        ttk.Checkbutton(body, text="AI図面解析実行時、対象ページの既存抽出行をAI結果で置き換える", variable=self.replace_var).pack(anchor="w", pady=1)

        # 4. カスタム指示
        tk.Label(body, text="3. AIへの追加指示（任意: 例「ボルト・ナットは除外」「備考に図面番号を入れる」など）:", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 8, "bold")).pack(anchor="w", pady=(6, 2))
        self.custom_var = tk.StringVar(value=self.config_obj.custom_instruction)
        ttk.Entry(body, textvariable=self.custom_var, width=72).pack(fill=tk.X, pady=(0, 6))

        # 接続テスト・一覧取得ステータス
        self.test_status_lbl = tk.Label(body, text="", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_ACCENT, font=("Meiryo UI", 8, "bold"), anchor="w")
        self.test_status_lbl.pack(fill=tk.X, pady=2)

        btn_bar = tk.Frame(body, bg=LightPalette.BG_PANEL)
        btn_bar.pack(fill=tk.X, pady=(6, 0))

        self.btn_test_conn = tk.Button(
            btn_bar, text="接続テスト", command=self._run_test,
            bg="#EDE9FE", fg="#5B21B6", relief=tk.SOLID, bd=1, padx=12, pady=5, font=("Meiryo UI", 9, "bold"), cursor="hand2"
        )
        self.btn_test_conn.pack(side=tk.LEFT)

        tk.Button(
            btn_bar, text="設定を保存", command=self._save,
            bg=LightPalette.BTN_PRIMARY_BG, fg=LightPalette.BTN_PRIMARY_FG, relief=tk.FLAT, padx=18, pady=5, font=("Meiryo UI", 9, "bold"), cursor="hand2"
        ).pack(side=tk.RIGHT, padx=(6, 0))

        tk.Button(
            btn_bar, text="キャンセル", command=self.destroy,
            bg=LightPalette.BTN_DEFAULT_BG, fg=LightPalette.BTN_DEFAULT_FG, relief=tk.SOLID, bd=1, padx=12, pady=5, font=("Meiryo UI", 9)
        ).pack(side=tk.RIGHT)

    def _populate_model_tree(self) -> None:
        self.model_tree.delete(*self.model_tree.get_children())
        combo_ids: List[str] = []
        for idx, m in enumerate(self.model_entries):
            mid = str(m.get("id", ""))
            if not mid:
                continue
            combo_ids.append(mid)
            rank = int(m.get("recommended_rank", 50))
            if m.get("recommended") or rank <= 12:
                tag = "recommended"
            elif rank >= 80:
                tag = "legacy"
            else:
                tag = "normal"

            iid = f"m_{idx}"
            self.model_tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    m.get("tier_badge", ""),
                    mid,
                    m.get("display_name", mid),
                    m.get("speed", ""),
                    m.get("accuracy", ""),
                    m.get("pricing", ""),
                    m.get("source", "標準カタログ"),
                ),
                tags=(tag,),
            )
        if combo_ids:
            self.model_combo.configure(values=combo_ids)

    def _sync_tree_and_detail_from_model_var(self) -> None:
        target_id = strip_gemini_model_prefix(self.model_var.get())
        self._update_model_detail_panel(target_id)
        if self._syncing_selection:
            return
        self._syncing_selection = True
        try:
            for iid in self.model_tree.get_children():
                vals = self.model_tree.item(iid, "values")
                if vals and len(vals) > 1 and vals[1] == target_id:
                    self.model_tree.selection_set(iid)
                    self.model_tree.see(iid)
                    break
        finally:
            self._syncing_selection = False

    def _on_model_var_edited(self) -> None:
        if self._syncing_selection:
            return
        self._sync_tree_and_detail_from_model_var()

    def _on_tree_select(self, _event: Any = None) -> None:
        if self._syncing_selection:
            return
        sel = self.model_tree.selection()
        if not sel:
            return
        vals = self.model_tree.item(sel[0], "values")
        if vals and len(vals) > 1:
            mid = str(vals[1]).strip()
            self._syncing_selection = True
            try:
                self.model_var.set(mid)
                self._update_model_detail_panel(mid)
            finally:
                self._syncing_selection = False

    def _on_tree_double_click(self, _event: Any = None) -> None:
        self._on_tree_select()
        mid = strip_gemini_model_prefix(self.model_var.get())
        if mid:
            self.test_status_lbl.configure(
                text=f"モデル「{mid}」を選択しました。「設定を保存」で確定します。",
                fg=LightPalette.FG_ACCENT,
            )

    def _update_model_detail_panel(self, model_id: str) -> None:
        entry = get_gemini_model_entry(model_id, self.model_entries)
        mid = entry.get("id", model_id)
        badge = entry.get("tier_badge", "")
        free_info = entry.get("free_tier", "")
        gen = entry.get("generation", "")
        in_tok = int(entry.get("input_token_limit", 0) or 0)
        tok_note = f" / 最大入力: {in_tok:,} tokens" if in_tok > 0 else ""
        mr_note = " / 高精細図面モード(MEDIA_RESOLUTION_HIGH)対応" if entry.get("supports_media_resolution") else ""
        self.model_detail_title_lbl.configure(
            text=f"選択モデル詳細: {mid} ({gen} / {badge} / 無料枠: {free_info}{tok_note}{mr_note})"
        )
        self.model_detail_desc_lbl.configure(text=str(entry.get("description", "")))

    def _apply_optimal_model(self, priority: str = "balanced") -> None:
        best_id = select_optimal_gemini_model(self.model_entries, priority=priority)
        self.model_var.set(best_id)
        self._sync_tree_and_detail_from_model_var()
        label_map = {
            "balanced": "図面解析バランス最適",
            "accuracy": "高精度優先最適",
            "speed": "高速・省コスト最適",
        }
        self.test_status_lbl.configure(
            text=f"「{label_map.get(priority, '最適')}」基準でモデル「{best_id}」を自動選択しました。",
            fg="#059669",
        )

    def _fetch_models_async(self) -> None:
        if self._is_fetching:
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showinfo("APIキー未入力", "APIから最新モデル一覧を取得するには、先に上の欄へ Gemini APIキーを入力してください。", parent=self)
            return

        self._is_fetching = True
        self.btn_fetch_models.configure(state=tk.DISABLED, text="取得中...")
        self.test_status_lbl.configure(text="Gemini API から現在利用可能な最新モデル一覧を取得しています...", fg=LightPalette.FG_SUB)
        self.update_idletasks()

        def worker() -> None:
            try:
                entries = GeminiVisionAnalyzer.fetch_available_models(api_key)
                self.after(0, lambda e=entries: self._on_fetch_models_done(e, None))
            except Exception as exc:
                msg = str(exc)
                self.after(0, lambda m=msg: self._on_fetch_models_done(None, m))

        threading.Thread(target=worker, daemon=True).start()

    def _on_fetch_models_done(self, entries: Optional[List[Dict[str, Any]]], err: Optional[str]) -> None:
        if not self.winfo_exists():
            return
        self._is_fetching = False
        self.btn_fetch_models.configure(state=tk.NORMAL, text="APIから最新モデル一覧を取得")
        if err or not entries:
            self.test_status_lbl.configure(text=f"モデル一覧取得エラー: {err or '不明なエラー'}", fg="#DC2626")
            return

        self.model_entries = entries
        live_count = sum(1 for m in entries if m.get("source") == "API利用可能")
        self._populate_model_tree()
        best_id = select_optimal_gemini_model(entries, priority="balanced")
        current_id = strip_gemini_model_prefix(self.model_var.get())
        live_ids = {m["id"] for m in entries if m.get("source") == "API利用可能"}
        if not current_id or (live_ids and current_id not in live_ids):
            self.model_var.set(best_id)
        self._sync_tree_and_detail_from_model_var()
        self.test_status_lbl.configure(
            text=f"APIから最新モデル {live_count}件を取得し一覧を更新しました（現在のおすすめ最適モデル: {best_id}）。",
            fg="#059669",
        )

    def _toggle_key_vis(self) -> None:
        self.show_key = not self.show_key
        self.key_entry.configure(show="" if self.show_key else "●")

    def _run_test(self) -> None:
        if self._is_testing:
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            self.test_status_lbl.configure(text="APIキーが入力されていません。", fg="#DC2626")
            return
        model_id = strip_gemini_model_prefix(self.model_var.get()) or DEFAULT_OPTIMAL_GEMINI_MODEL
        tmp_cfg = GeminiConfig(api_key=api_key, model_name=model_id)
        self._is_testing = True
        self.btn_test_conn.configure(state=tk.DISABLED, text="テスト中...")
        self.test_status_lbl.configure(text=f"Gemini API ({model_id}) へ接続テスト中...", fg=LightPalette.FG_SUB)
        self.update_idletasks()

        def worker() -> None:
            ok, msg = GeminiVisionAnalyzer.test_connection(tmp_cfg)
            self.after(0, lambda o=ok, m=msg: self._on_test_done(o, m))

        threading.Thread(target=worker, daemon=True).start()

    def _on_test_done(self, ok: bool, msg: str) -> None:
        if not self.winfo_exists():
            return
        self._is_testing = False
        self.btn_test_conn.configure(state=tk.NORMAL, text="接続テスト")
        self.test_status_lbl.configure(text=msg, fg="#059669" if ok else "#DC2626")

    def _save(self) -> None:
        self.config_obj.api_key = self.api_key_var.get().strip()
        self.config_obj.model_name = strip_gemini_model_prefix(self.model_var.get()) or DEFAULT_OPTIMAL_GEMINI_MODEL
        self.config_obj.auto_use_on_load = self.auto_load_var.get()
        self.config_obj.use_for_area_select = self.area_ai_var.get()
        self.config_obj.replace_existing_on_ai = self.replace_var.get()
        self.config_obj.custom_instruction = self.custom_var.get().strip()
        self.config_obj.save_api_key = self.save_key_var.get()
        live_only = [m for m in self.model_entries if m.get("source") == "API利用可能"]
        if live_only:
            self.config_obj.cached_models = live_only
        self.config_obj.save()
        if self.on_save_callback:
            self.on_save_callback()
        self.destroy()


class WeightCalculatorDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, initial_material: str, on_apply_callback: Any):
        super().__init__(parent)
        self.title(f"部品重量計算ツール - {APP_NAME}")
        self.geometry("500x450")
        self.configure(bg=LightPalette.BG_MAIN)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        enforce_windows_light_mode(self)

        self.on_apply_callback = on_apply_callback
        self.calculated_weight = 0.0
        self.generated_spec = ""
        self._build_ui(initial_material)
        self._recalculate()

    def _build_ui(self, initial_material: str) -> None:
        header = tk.Frame(self, bg=LightPalette.BG_HEADER, padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="形状・材質からの単重(kg)自動計算機", bg=LightPalette.BG_HEADER, fg="#FFFFFF", font=("Meiryo UI", 11, "bold")).pack(anchor="w")

        body = tk.Frame(self, bg=LightPalette.BG_PANEL, padx=20, pady=14)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        tk.Label(body, text="1. 部品の形状:", bg=LightPalette.BG_PANEL, font=("Meiryo UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.shape_var = tk.StringVar(value="プレート / 角材 (t × W × L)")
        shapes = ["プレート / 角材 (t × W × L)", "丸棒 / 円柱 (φD × L)", "丸パイプ (外径φD × 肉厚t × L)", "角パイプ (辺A × 辺B × 肉厚t × L)", "六角棒 (対辺H × L)"]
        cb_s = ttk.Combobox(body, textvariable=self.shape_var, values=shapes, state="readonly", width=32)
        cb_s.grid(row=0, column=1, sticky="ew", pady=5, padx=6)
        cb_s.bind("<<ComboboxSelected>>", lambda _e: self._on_shape_change())

        tk.Label(body, text="2. 材質 (JIS):", bg=LightPalette.BG_PANEL, font=("Meiryo UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        mat_def = initial_material if initial_material in MATERIAL_MASTER else "SS400"
        self.mat_var = tk.StringVar(value=mat_def)
        cb_m = ttk.Combobox(body, textvariable=self.mat_var, values=list(MATERIAL_MASTER.keys()), width=32)
        cb_m.grid(row=1, column=1, sticky="ew", pady=5, padx=6)
        cb_m.bind("<<ComboboxSelected>>", lambda _e: self._on_mat_change())

        tk.Label(body, text="3. 比重 (g/cm³):", bg=LightPalette.BG_PANEL).grid(row=2, column=0, sticky="w", pady=5)
        self.density_var = tk.StringVar(value=str(MATERIAL_MASTER.get(mat_def, {}).get("density", 7.85)))
        self.density_var.trace_add("write", lambda *_: self._recalculate())
        ttk.Entry(body, textvariable=self.density_var, width=16).grid(row=2, column=1, sticky="w", pady=5, padx=6)

        self.dim_labels: List[tk.Label] = []
        self.dim_vars: List[tk.StringVar] = []
        self.dim_entries: List[ttk.Entry] = []
        for i, dv in enumerate(["10", "100", "200", "3.2"]):
            lbl = tk.Label(body, text=f"寸法 {i + 1} (mm):", bg=LightPalette.BG_PANEL)
            lbl.grid(row=3 + i, column=0, sticky="w", pady=4)
            v = tk.StringVar(value=dv)
            v.trace_add("write", lambda *_: self._recalculate())
            ent = ttk.Entry(body, textvariable=v, width=18)
            ent.grid(row=3 + i, column=1, sticky="w", pady=4, padx=6)
            self.dim_labels.append(lbl)
            self.dim_vars.append(v)
            self.dim_entries.append(ent)

        res_f = tk.Frame(body, bg=LightPalette.BG_SELECTED, padx=12, pady=8)
        res_f.grid(row=7, column=0, columnspan=2, sticky="ew", pady=12)
        self.result_lbl = tk.Label(res_f, text="計算単重: 0.000 kg", bg=LightPalette.BG_SELECTED, fg=LightPalette.FG_ACCENT, font=("Meiryo UI", 12, "bold"))
        self.result_lbl.pack(anchor="w")
        self.spec_lbl = tk.Label(res_f, text="", bg=LightPalette.BG_SELECTED, fg=LightPalette.FG_SUB, font=("Meiryo UI", 9))
        self.spec_lbl.pack(anchor="w")

        btn_b = tk.Frame(body, bg=LightPalette.BG_PANEL)
        btn_b.grid(row=8, column=0, columnspan=2, sticky="e")
        tk.Button(btn_b, text="選択中の部品行へ反映", bg=LightPalette.BTN_PRIMARY_BG, fg="#FFFFFF", relief=tk.FLAT, padx=14, pady=5, font=("Meiryo UI", 9, "bold"), command=self._apply).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_b, text="閉じる", bg=LightPalette.BTN_DEFAULT_BG, relief=tk.SOLID, bd=1, padx=12, pady=5, command=self.destroy).pack(side=tk.LEFT)
        self._on_shape_change()

    def _on_mat_change(self) -> None:
        m = self.mat_var.get()
        if m in MATERIAL_MASTER:
            self.density_var.set(str(MATERIAL_MASTER[m]["density"]))
        self._recalculate()

    def _on_shape_change(self) -> None:
        s = self.shape_var.get()
        if "プレート" in s:
            cfgs = [("板厚 t (mm):", True), ("幅 W (mm):", True), ("長さ L (mm):", True), ("-", False)]
        elif "丸棒" in s:
            cfgs = [("外径 φD (mm):", True), ("長さ L (mm):", True), ("-", False), ("-", False)]
        elif "丸パイプ" in s:
            cfgs = [("外径 φD (mm):", True), ("肉厚 t (mm):", True), ("長さ L (mm):", True), ("-", False)]
        elif "角パイプ" in s:
            cfgs = [("辺 A (mm):", True), ("辺 B (mm):", True), ("肉厚 t (mm):", True), ("長さ L (mm):", True)]
        else:
            cfgs = [("六角対辺 H (mm):", True), ("長さ L (mm):", True), ("-", False), ("-", False)]
        for i, (lt, act) in enumerate(cfgs):
            self.dim_labels[i].configure(text=lt)
            self.dim_entries[i].configure(state="normal" if act else "disabled")
        self._recalculate()

    def _recalculate(self) -> None:
        try:
            density = float(self.density_var.get())
            vals = [float(v.get() or 0.0) for v in self.dim_vars]
        except ValueError:
            return
        s = self.shape_var.get()
        if "プレート" in s:
            vol, spec = vals[0] * vals[1] * vals[2], f"t{vals[0]:g}x{vals[1]:g}x{vals[2]:g}"
        elif "丸棒" in s:
            r = vals[0] * 0.5
            vol, spec = math.pi * (r * r) * vals[1], f"φ{vals[0]:g}x{vals[1]:g}"
        elif "丸パイプ" in s:
            ro, ri = vals[0] * 0.5, max(vals[0] * 0.5 - vals[1], 0.0)
            vol, spec = math.pi * (ro * ro - ri * ri) * vals[2], f"φ{vals[0]:g}xt{vals[1]:g}x{vals[2]:g}"
        elif "角パイプ" in s:
            ia, ib = max(vals[0] - 2.0 * vals[2], 0.0), max(vals[1] - 2.0 * vals[2], 0.0)
            vol, spec = (vals[0] * vals[1] - ia * ib) * vals[3], f"{vals[0]:g}x{vals[1]:g}xt{vals[2]:g}x{vals[3]:g}"
        else:
            vol, spec = (math.sqrt(3.0) * 0.5) * (vals[0] * vals[0]) * vals[1], f"HEX{vals[0]:g}x{vals[1]:g}"
        self.calculated_weight = round(max(vol, 0.0) * density / 1_000_000.0, 3)
        self.generated_spec = spec
        self.result_lbl.configure(text=f"計算単重: {self.calculated_weight:.3f} kg")
        self.spec_lbl.configure(text=f"寸法表記: {spec} (材質: {self.mat_var.get()})")

    def _apply(self) -> None:
        if self.on_apply_callback:
            self.on_apply_callback(self.calculated_weight, self.generated_spec, self.mat_var.get())
        self.destroy()


class ExcelSettingsDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, config: ExcelExportConfig, on_save_callback: Any):
        super().__init__(parent)
        self.title(f"Excel出力設定 - {APP_NAME}")
        self.geometry("460x340")
        self.configure(bg=LightPalette.BG_MAIN)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        enforce_windows_light_mode(self)
        self.config_obj = config
        self.on_save_callback = on_save_callback

        header = tk.Frame(self, bg=LightPalette.BG_HEADER, padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="Excel部品表 (.xlsx) 出力設定", bg=LightPalette.BG_HEADER, fg="#FFFFFF", font=("Meiryo UI", 11, "bold")).pack(anchor="w")

        body = tk.Frame(self, bg=LightPalette.BG_PANEL, padx=20, pady=14)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        self.tpl_var = tk.StringVar(value=config.template_type)
        for lt, vt in [("標準部品表フォーマット（全項目・推奨）", "standard"), ("購買・見積手配用フォーマット", "purchasing"), ("製造・加工手配用フォーマット", "manufacturing")]:
            ttk.Radiobutton(body, text=lt, variable=self.tpl_var, value=vt).pack(anchor="w", pady=2)

        ttk.Separator(body, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        self.split_var = tk.BooleanVar(value=config.split_by_drawing)
        ttk.Checkbutton(body, text="複数図面がある場合に「図面別シート」も同時作成する", variable=self.split_var).pack(anchor="w", pady=2)
        self.merge_var = tk.BooleanVar(value=config.merge_duplicates)
        ttk.Checkbutton(body, text="出力時に同一品番・品名を名寄せ（数量合算）して集計する", variable=self.merge_var).pack(anchor="w", pady=2)
        self.prev_var = tk.BooleanVar(value=config.include_preview_sheet)
        ttk.Checkbutton(body, text="別シートに「2D図面プレビュー画像」を添付する", variable=self.prev_var).pack(anchor="w", pady=2)

        btn_b = tk.Frame(body, bg=LightPalette.BG_PANEL)
        btn_b.pack(fill=tk.X, pady=(14, 0))
        tk.Button(btn_b, text="設定を保存", bg=LightPalette.BTN_PRIMARY_BG, fg="#FFFFFF", relief=tk.FLAT, padx=16, pady=5, font=("Meiryo UI", 9, "bold"), command=self._save).pack(side=tk.RIGHT, padx=4)
        tk.Button(btn_b, text="キャンセル", bg=LightPalette.BTN_DEFAULT_BG, relief=tk.SOLID, bd=1, padx=12, pady=5, command=self.destroy).pack(side=tk.RIGHT)

    def _save(self) -> None:
        self.config_obj.template_type = self.tpl_var.get()
        self.config_obj.split_by_drawing = self.split_var.get()
        self.config_obj.merge_duplicates = self.merge_var.get()
        self.config_obj.include_preview_sheet = self.prev_var.get()
        if self.on_save_callback:
            self.on_save_callback()
        self.destroy()


class ReadmeViewerDialog(tk.Toplevel):
    """目次ジャンプ・キーワード検索・リッチテキスト整形付きの取扱説明書ビューア"""

    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        self.title(f"使い方ガイド (取扱説明書) - {APP_NAME}")
        self.geometry("880x680")
        self.minsize(720, 520)
        self.configure(bg=LightPalette.BG_MAIN)
        self.transient(parent)
        enforce_windows_light_mode(self)

        self.section_marks: List[Tuple[str, str]] = []
        self._search_pos = "1.0"

        header = tk.Frame(self, bg=LightPalette.BG_HEADER, padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text=f"{APP_NAME} 使い方ガイド（取扱説明書）",
            bg=LightPalette.BG_HEADER,
            fg="#FFFFFF",
            font=("Meiryo UI", 12, "bold"),
        ).pack(side=tk.LEFT)

        search_box = tk.Frame(header, bg=LightPalette.BG_HEADER)
        search_box.pack(side=tk.RIGHT)
        tk.Label(search_box, text="ガイド内検索:", bg=LightPalette.BG_HEADER, fg="#E2E8F0", font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar(value="")
        s_ent = ttk.Entry(search_box, textvariable=self.search_var, width=20)
        s_ent.pack(side=tk.LEFT, padx=2)
        s_ent.bind("<Return>", lambda _e: self._search_next())
        tk.Button(
            search_box, text="次を検索", command=self._search_next,
            bg="#2563EB", fg="#FFFFFF", relief=tk.FLAT, padx=8, pady=2, font=("Meiryo UI", 8, "bold"), cursor="hand2"
        ).pack(side=tk.LEFT, padx=2)
        tk.Button(
            search_box, text="クリア", command=self._clear_search,
            bg="#FFFFFF", fg="#1E293B", relief=tk.FLAT, padx=6, pady=2, font=("Meiryo UI", 8), cursor="hand2"
        ).pack(side=tk.LEFT, padx=2)

        # 目次ジャンプバー
        self.nav_bar = tk.Frame(self, bg="#E2E8F0", padx=12, pady=6)
        self.nav_bar.pack(fill=tk.X)
        tk.Label(self.nav_bar, text="目次ジャンプ:", bg="#E2E8F0", fg=LightPalette.FG_ACCENT, font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))

        tf = tk.Frame(self, bg=LightPalette.BG_PANEL, padx=10, pady=10, bd=1, relief=tk.SOLID)
        tf.pack(fill=tk.BOTH, expand=True, padx=12, pady=(8, 10))
        sb = ttk.Scrollbar(tf, orient=tk.VERTICAL)
        self.txt = tk.Text(
            tf,
            bg=LightPalette.BG_PANEL,
            fg=LightPalette.FG_MAIN,
            font=("Meiryo UI", 10),
            wrap=tk.WORD,
            yscrollcommand=sb.set,
            relief=tk.FLAT,
            padx=14,
            pady=10,
            spacing1=2,
            spacing2=2,
            spacing3=4,
        )
        sb.configure(command=self.txt.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._configure_text_tags()

        rp = get_resource_path("README.md")
        content = "README.md が見つかりませんでした。"
        if os.path.exists(rp):
            try:
                with open(rp, "r", encoding="utf-8") as f:
                    content = f.read()
            except Exception as e:
                content = str(e)

        self._render_formatted_markdown(content)
        self._build_nav_buttons()
        self.txt.configure(state=tk.DISABLED)

    def _configure_text_tags(self) -> None:
        self.txt.tag_configure("h1", font=("Meiryo UI", 14, "bold"), foreground="#1E3A8A", background="#EFF6FF", spacing1=8, spacing3=8, lmargin1=6, lmargin2=6)
        self.txt.tag_configure("h2", font=("Meiryo UI", 11, "bold"), foreground="#1D4ED8", background="#F8FAFC", spacing1=12, spacing3=6, lmargin1=4, lmargin2=4)
        self.txt.tag_configure("h3", font=("Meiryo UI", 10, "bold"), foreground="#5B21B6", spacing1=8, spacing3=4, lmargin1=8, lmargin2=8)
        self.txt.tag_configure("bullet", font=("Meiryo UI", 10), foreground="#0F172A", lmargin1=18, lmargin2=32, spacing1=2, spacing3=2)
        self.txt.tag_configure("sub_bullet", font=("Meiryo UI", 9), foreground="#334155", lmargin1=36, lmargin2=50, spacing1=1, spacing3=2)
        self.txt.tag_configure("step", font=("Meiryo UI", 10), foreground="#0F172A", lmargin1=14, lmargin2=28, spacing1=3, spacing3=3)
        self.txt.tag_configure("code", font=("Consolas", 10, "bold"), foreground="#1E40AF", background="#DBEAFE")
        self.txt.tag_configure("hr", font=("Meiryo UI", 6), foreground="#CBD5E1", spacing1=4, spacing3=4)
        self.txt.tag_configure("search_hit", background="#FEF08A", foreground="#0F172A")

    def _insert_line_with_inline_code(self, line_text: str, base_tag: str = "") -> None:
        parts = re.split(r"(`[^`]+`)", line_text)
        for part in parts:
            if not part:
                continue
            if part.startswith("`") and part.endswith("`") and len(part) >= 2:
                code_str = f" {part[1:-1]} "
                tags = (base_tag, "code") if base_tag else ("code",)
                self.txt.insert(tk.END, code_str, tags)
            else:
                tags = (base_tag,) if base_tag else ()
                self.txt.insert(tk.END, part, tags)
        self.txt.insert(tk.END, "\n", (base_tag,) if base_tag else ())

    def _render_formatted_markdown(self, raw_md: str) -> None:
        for line in raw_md.splitlines():
            stripped = line.strip()
            if stripped.startswith("# "):
                title_t = stripped[2:].strip()
                self.txt.insert(tk.END, f"  {title_t}\n", ("h1",))
            elif stripped.startswith("## "):
                sec_t = stripped[3:].strip()
                mark_name = f"sec_{len(self.section_marks)}"
                idx_pos = self.txt.index(tk.END)
                self.txt.mark_set(mark_name, idx_pos)
                self.txt.mark_gravity(mark_name, tk.LEFT)
                short_lbl = re.sub(r"^\d+\.\s*", "", sec_t)[:14]
                self.section_marks.append((short_lbl, mark_name))
                self.txt.insert(tk.END, f"■ {sec_t}\n", ("h2",))
            elif stripped.startswith("### "):
                sub_t = stripped[4:].strip()
                self.txt.insert(tk.END, f"◆ {sub_t}\n", ("h3",))
            elif stripped == "---":
                self.txt.insert(tk.END, "─" * 76 + "\n", ("hr",))
            elif line.startswith("   - ") or line.startswith("     - ") or line.startswith("  - "):
                self._insert_line_with_inline_code("・ " + stripped[2:].strip(), "sub_bullet")
            elif stripped.startswith("- "):
                self._insert_line_with_inline_code("● " + stripped[2:].strip(), "bullet")
            elif re.match(r"^\d+\.\s+", stripped):
                self._insert_line_with_inline_code(stripped, "step")
            else:
                self._insert_line_with_inline_code(line, "")

    def _build_nav_buttons(self) -> None:
        for lbl, mark in self.section_marks:
            tk.Button(
                self.nav_bar,
                text=lbl,
                command=lambda m=mark: self.txt.see(m),
                bg="#FFFFFF",
                fg="#1E3A8A",
                relief=tk.SOLID,
                bd=1,
                padx=7,
                pady=2,
                font=("Meiryo UI", 8, "bold"),
                cursor="hand2",
            ).pack(side=tk.LEFT, padx=2)

    def _search_next(self) -> None:
        kw = self.search_var.get().strip()
        self.txt.tag_remove("search_hit", "1.0", tk.END)
        if not kw:
            return
        pos = self.txt.search(kw, self._search_pos, stopindex=tk.END, nocase=True)
        if not pos:
            pos = self.txt.search(kw, "1.0", stopindex=tk.END, nocase=True)
        if pos:
            end_pos = f"{pos}+{len(kw)}c"
            self.txt.tag_add("search_hit", pos, end_pos)
            self.txt.see(pos)
            self._search_pos = end_pos

    def _clear_search(self) -> None:
        self.search_var.set("")
        self._search_pos = "1.0"
        self.txt.tag_remove("search_hit", "1.0", tk.END)


# ==============================================================================
# メインアプリケーション GUI (PartLensApp)
# ==============================================================================

class PartLensApp:
    TREE_COLUMNS = (
        "item_no", "part_number", "part_name", "specification", "material", "treatment",
        "quantity", "unit", "unit_weight", "total_weight", "category", "manufacturer",
        "unit_price", "amount", "source_file", "remarks",
    )

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1500x900")
        self.root.minsize(1200, 720)
        self.root.configure(bg=LightPalette.BG_MAIN)

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        self._setup_light_mode_styles()
        enforce_windows_light_mode(self.root)

        self.gemini_config = GeminiConfig()
        self.gemini_config.load()
        self.is_ai_busy = False

        self.documents: List[DrawingDocument] = []
        self.current_doc_idx = -1
        self.current_page_idx = 0
        self.parts: List[PartItem] = []
        self.selected_part_uid: Optional[str] = None

        self.undo_stack: List[List[PartItem]] = []
        self.redo_stack: List[List[PartItem]] = []

        self.zoom_scale = 1.0
        self.canvas_offset_x = 0.0
        self.canvas_offset_y = 0.0
        self.interaction_mode = "pan"
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.selection_rect_id: Optional[int] = None
        self.highlight_bbox: Optional[Tuple[float, float, float, float]] = None
        self._tk_preview_img: Optional[ImageTk.PhotoImage] = None
        self._inline_editor: Optional[tk.Widget] = None
        self._sort_col_desc: Dict[str, bool] = {}

        self.excel_config = ExcelExportConfig(project_name="新規機械設計プロジェクト", drawing_no="ASM-0001")

        self._build_header()
        self._build_main_workspace()
        self._build_statusbar()
        self._bind_shortcuts()
        self._refresh_quick_model_combo()
        self._update_ai_badge()

    def _setup_light_mode_styles(self) -> None:
        self.root.option_add("*Background", LightPalette.BG_PANEL)
        self.root.option_add("*Foreground", LightPalette.FG_MAIN)
        self.root.option_add("*Entry.Background", "#FFFFFF")
        self.root.option_add("*Entry.Foreground", LightPalette.FG_MAIN)
        self.root.option_add("*Listbox.Background", "#FFFFFF")
        self.root.option_add("*Listbox.Foreground", LightPalette.FG_MAIN)
        self.root.option_add("*TCombobox*Listbox.background", "#FFFFFF")
        self.root.option_add("*TCombobox*Listbox.foreground", LightPalette.FG_MAIN)

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", background=LightPalette.BG_PANEL, foreground=LightPalette.FG_MAIN, font=("Meiryo UI", 9))
        style.configure("TFrame", background=LightPalette.BG_PANEL)
        style.configure("TLabel", background=LightPalette.BG_PANEL, foreground=LightPalette.FG_MAIN, font=("Meiryo UI", 9))
        style.configure("TLabelframe", background=LightPalette.BG_PANEL, bordercolor=LightPalette.BORDER)
        style.configure("TLabelframe.Label", background=LightPalette.BG_PANEL, foreground=LightPalette.FG_ACCENT, font=("Meiryo UI", 9, "bold"))
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground=LightPalette.FG_MAIN, bordercolor=LightPalette.BORDER, padding=4)
        style.configure("TCombobox", fieldbackground="#FFFFFF", background=LightPalette.BG_SUB, foreground=LightPalette.FG_MAIN, padding=4)
        style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")], foreground=[("readonly", LightPalette.FG_MAIN)])
        style.configure("TCheckbutton", background=LightPalette.BG_PANEL, foreground=LightPalette.FG_MAIN, indicatorbackground="#FFFFFF")
        style.configure("TRadiobutton", background=LightPalette.BG_PANEL, foreground=LightPalette.FG_MAIN, indicatorbackground="#FFFFFF")
        style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground=LightPalette.FG_MAIN, rowheight=26, font=("Meiryo UI", 9))
        style.configure("Treeview.Heading", background="#E2E8F0", foreground="#0F172A", font=("Meiryo UI", 9, "bold"), padding=(4, 5))
        style.map("Treeview", background=[("selected", LightPalette.BG_SELECTED)], foreground=[("selected", LightPalette.FG_ACCENT)])

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=LightPalette.BG_HEADER, padx=14, pady=9)
        header.pack(fill=tk.X)

        title_box = tk.Frame(header, bg=LightPalette.BG_HEADER)
        title_box.pack(side=tk.LEFT)
        tk.Label(title_box, text="PartLens Miya", bg=LightPalette.BG_HEADER, fg="#FFFFFF", font=("Meiryo UI", 14, "bold")).pack(side=tk.LEFT)
        self.ai_badge_lbl = tk.Label(title_box, text="Gemini AI: 未設定", bg="#312E81", fg="#E0E7FF", font=("Meiryo UI", 8, "bold"), padx=8, pady=2)
        self.ai_badge_lbl.pack(side=tk.LEFT, padx=10)

        btn_box = tk.Frame(header, bg=LightPalette.BG_HEADER)
        btn_box.pack(side=tk.RIGHT)
        self._hdr_btn(btn_box, "図面ファイルを開く (PDF/DXF)", self.action_open_files, "#2563EB", "#FFFFFF")
        self._hdr_btn(btn_box, "フォルダ一括読込", self.action_open_folder, "#3B82F6", "#FFFFFF")
        self._hdr_btn(btn_box, "サンプル図面で試す", self.action_load_sample_drawings, "#059669", "#FFFFFF")
        self._hdr_btn(btn_box, "Gemini AI設定", self.action_open_gemini_settings, "#7C3AED", "#FFFFFF")
        self._hdr_btn(btn_box, "プロジェクト保存", self.action_save_project, "#FFFFFF", "#1E293B")
        self._hdr_btn(btn_box, "プロジェクト読込", self.action_load_project, "#FFFFFF", "#1E293B")
        self._hdr_btn(btn_box, "使い方ガイド", self.action_show_readme, "#FEF3C7", "#92400E")

    def _hdr_btn(self, parent: tk.Misc, text: str, cmd: Any, bg: str, fg: str) -> tk.Button:
        b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg, activebackground=bg, activeforeground=fg, relief=tk.FLAT, padx=10, pady=4, font=("Meiryo UI", 9, "bold"), cursor="hand2")
        b.pack(side=tk.LEFT, padx=3)
        return b

    def _update_ai_badge(self) -> None:
        cur_model = strip_gemini_model_prefix(self.gemini_config.model_name) or DEFAULT_OPTIMAL_GEMINI_MODEL
        if self.gemini_config.api_key.strip():
            self.ai_badge_lbl.configure(text=f"Gemini AI: 有効 ({cur_model})", bg="#5B21B6", fg="#FFFFFF")
        else:
            self.ai_badge_lbl.configure(text=f"Gemini AI: APIキー未設定 (選択モデル: {cur_model})", bg="#334155", fg="#E2E8F0")

    def _get_current_model_entries(self) -> List[Dict[str, Any]]:
        if self.gemini_config.cached_models:
            return GeminiVisionAnalyzer.merge_catalog_with_api_models(self.gemini_config.cached_models)
        return copy.deepcopy(GEMINI_MODEL_CATALOG)

    def _refresh_quick_model_combo(self) -> None:
        if not hasattr(self, "quick_model_combo"):
            return
        entries = self._get_current_model_entries()
        cur_id = strip_gemini_model_prefix(self.gemini_config.model_name) or DEFAULT_OPTIMAL_GEMINI_MODEL
        choices: List[str] = []
        selected_label = ""
        for e in entries:
            lbl = format_model_choice_label(e)
            choices.append(lbl)
            if e.get("id") == cur_id:
                selected_label = lbl
        if not selected_label:
            custom_entry = get_gemini_model_entry(cur_id, entries)
            selected_label = format_model_choice_label(custom_entry)
            choices.insert(0, selected_label)
        self.quick_model_combo.configure(values=choices)
        self.quick_model_var.set(selected_label)

    def _on_quick_model_selected(self, _event: Any = None) -> None:
        chosen_id = extract_model_id_from_choice(self.quick_model_var.get())
        if not chosen_id:
            return
        self.gemini_config.model_name = chosen_id
        self.gemini_config.save()
        self._update_ai_badge()
        entry = get_gemini_model_entry(chosen_id, self._get_current_model_entries())
        self.set_status(f"使用Geminiモデルを「{chosen_id} ({entry.get('tier_badge', '')})」に切り替えました。")

    def action_select_optimal_model(self) -> None:
        entries = self._get_current_model_entries()
        best_id = select_optimal_gemini_model(entries, priority="balanced")
        self.gemini_config.model_name = best_id
        self.gemini_config.save()
        self._refresh_quick_model_combo()
        self._update_ai_badge()
        self.set_status(f"2D図面解析の推奨最適モデル「{best_id}」を設定しました。")

    def _on_gemini_settings_saved(self) -> None:
        self._refresh_quick_model_combo()
        self._update_ai_badge()
        self.set_status(f"Gemini AI設定を更新しました（使用モデル: {self.gemini_config.model_name}）。")

    def _build_main_workspace(self) -> None:
        self.paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        left_pane = tk.Frame(self.paned, bg=LightPalette.BG_PANEL, bd=1, relief=tk.SOLID)
        self.paned.add(left_pane, weight=4)
        self._build_left_preview_pane(left_pane)

        right_pane = tk.Frame(self.paned, bg=LightPalette.BG_PANEL, bd=1, relief=tk.SOLID)
        self.paned.add(right_pane, weight=6)
        self._build_right_bom_pane(right_pane)

    def _build_left_preview_pane(self, parent: tk.Frame) -> None:
        doc_frame = tk.Frame(parent, bg=LightPalette.BG_SUB, padx=8, pady=6)
        doc_frame.pack(fill=tk.X)

        top_bar = tk.Frame(doc_frame, bg=LightPalette.BG_SUB)
        top_bar.pack(fill=tk.X)
        tk.Label(top_bar, text="■ 読み込み済み 2D図面一覧", bg=LightPalette.BG_SUB, fg=LightPalette.FG_ACCENT, font=("Meiryo UI", 9, "bold")).pack(side=tk.LEFT)

        tk.Button(top_bar, text="一覧クリア", command=self.action_clear_documents, bg="#FFFFFF", relief=tk.SOLID, bd=1, padx=6, font=("Meiryo UI", 8)).pack(side=tk.RIGHT, padx=2)
        tk.Button(top_bar, text="全図面をAI一括解析", command=self.action_ai_analyze_all_docs, bg=LightPalette.BTN_AI_BG, fg="#FFFFFF", relief=tk.FLAT, padx=8, font=("Meiryo UI", 8, "bold"), cursor="hand2").pack(side=tk.RIGHT, padx=3)
        tk.Button(top_bar, text="AI図面解析 (Gemini)", command=self.action_ai_analyze_current_page, bg="#6D28D9", fg="#FFFFFF", relief=tk.FLAT, padx=8, font=("Meiryo UI", 8, "bold"), cursor="hand2").pack(side=tk.RIGHT, padx=3)

        list_c = tk.Frame(doc_frame, bg=LightPalette.BG_PANEL)
        list_c.pack(fill=tk.X, pady=(4, 0))
        self.doc_listbox = tk.Listbox(list_c, height=4, bg="#FFFFFF", fg=LightPalette.FG_MAIN, selectbackground=LightPalette.BG_SELECTED, selectforeground=LightPalette.FG_ACCENT, font=("Meiryo UI", 9), relief=tk.SOLID, bd=1, exportselection=False)
        sb = ttk.Scrollbar(list_c, orient=tk.VERTICAL, command=self.doc_listbox.yview)
        self.doc_listbox.configure(yscrollcommand=sb.set)
        self.doc_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.doc_listbox.bind("<<ListboxSelect>>", self._on_select_document)

        ai_bar = tk.Frame(parent, bg="#F5F3FF", padx=8, pady=4, bd=1, relief=tk.SOLID)
        ai_bar.pack(fill=tk.X, padx=0, pady=(0, 2))
        tk.Label(ai_bar, text="Geminiモデル:", bg="#F5F3FF", fg="#5B21B6", font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT)
        self.quick_model_var = tk.StringVar()
        self.quick_model_combo = ttk.Combobox(ai_bar, textvariable=self.quick_model_var, state="readonly", width=34)
        self.quick_model_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 4))
        self.quick_model_combo.bind("<<ComboboxSelected>>", self._on_quick_model_selected)
        tk.Button(
            ai_bar, text="★最適選択", command=self.action_select_optimal_model,
            bg="#EDE9FE", fg="#5B21B6", relief=tk.SOLID, bd=1, padx=6, pady=1, font=("Meiryo UI", 8, "bold"), cursor="hand2"
        ).pack(side=tk.LEFT, padx=2)
        tk.Button(
            ai_bar, text="一覧/設定", command=self.action_open_gemini_settings,
            bg="#FFFFFF", fg="#1E293B", relief=tk.SOLID, bd=1, padx=6, pady=1, font=("Meiryo UI", 8), cursor="hand2"
        ).pack(side=tk.LEFT, padx=(2, 0))

        ctrl_bar = tk.Frame(parent, bg=LightPalette.BG_PANEL, padx=8, pady=5)
        ctrl_bar.pack(fill=tk.X)
        tk.Button(ctrl_bar, text="◀", width=2, command=lambda: self._change_page(-1), bg="#FFFFFF", relief=tk.SOLID, bd=1).pack(side=tk.LEFT)
        self.page_info_lbl = tk.Label(ctrl_bar, text="Page: 0 / 0", width=11, font=("Meiryo UI", 8, "bold"))
        self.page_info_lbl.pack(side=tk.LEFT, padx=2)
        tk.Button(ctrl_bar, text="▶", width=2, command=lambda: self._change_page(1), bg="#FFFFFF", relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(ctrl_bar, text="－", width=2, command=lambda: self._zoom_by(0.8), bg="#FFFFFF", relief=tk.SOLID, bd=1).pack(side=tk.LEFT)
        self.zoom_lbl = tk.Label(ctrl_bar, text="100%", width=6, font=("Meiryo UI", 8))
        self.zoom_lbl.pack(side=tk.LEFT)
        tk.Button(ctrl_bar, text="＋", width=2, command=lambda: self._zoom_by(1.25), bg="#FFFFFF", relief=tk.SOLID, bd=1).pack(side=tk.LEFT)
        tk.Button(ctrl_bar, text="全体フィット", command=self._fit_canvas_view, bg="#FFFFFF", relief=tk.SOLID, bd=1, padx=6, font=("Meiryo UI", 8)).pack(side=tk.LEFT, padx=4)

        self.btn_mode_select = tk.Button(ctrl_bar, text="範囲選択して部品抽出: OFF", command=self._toggle_interaction_mode, bg="#FFFFFF", fg=LightPalette.FG_ACCENT, relief=tk.SOLID, bd=1, padx=8, pady=2, font=("Meiryo UI", 8, "bold"), cursor="hand2")
        self.btn_mode_select.pack(side=tk.RIGHT)

        tk.Label(
            parent,
            text="ヒント: 「AI図面解析 (Gemini)」で画像＋文字のハイブリッドAI解析を実行。範囲選択ONで囲むと選択箇所をピンポイントAI抽出します。",
            bg="#F3E8FF", fg="#5B21B6", font=("Meiryo UI", 8), anchor="w", padx=8, pady=4
        ).pack(fill=tk.X)

        self.preview_canvas = tk.Canvas(parent, bg=LightPalette.BG_CANVAS, highlightthickness=0, cursor="fleur")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas.bind("<Configure>", lambda _e: self._redraw_canvas())
        self.preview_canvas.bind("<ButtonPress-1>", self._on_canvas_mouse_down)
        self.preview_canvas.bind("<B1-Motion>", self._on_canvas_mouse_move)
        self.preview_canvas.bind("<ButtonRelease-1>", self._on_canvas_mouse_up)
        self.preview_canvas.bind("<Button-3>", self._on_canvas_right_click)
        self.preview_canvas.bind("<MouseWheel>", lambda e: self._zoom_by(1.15 if e.delta > 0 else 0.85))

    def _build_right_bom_pane(self, parent: tk.Frame) -> None:
        meta_bar = tk.Frame(parent, bg=LightPalette.BG_SUB, padx=10, pady=8)
        meta_bar.pack(fill=tk.X)
        tk.Label(meta_bar, text="案件名/製品名:", bg=LightPalette.BG_SUB, font=("Meiryo UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        self.proj_name_var = tk.StringVar(value=self.excel_config.project_name)
        ttk.Entry(meta_bar, textvariable=self.proj_name_var, width=24).grid(row=0, column=1, padx=(4, 10))

        tk.Label(meta_bar, text="総合図番:", bg=LightPalette.BG_SUB, font=("Meiryo UI", 9, "bold")).grid(row=0, column=2, sticky="w")
        self.dwg_no_var = tk.StringVar(value=self.excel_config.drawing_no)
        ttk.Entry(meta_bar, textvariable=self.dwg_no_var, width=15).grid(row=0, column=3, padx=(4, 10))

        tk.Label(meta_bar, text="作成者:", bg=LightPalette.BG_SUB, font=("Meiryo UI", 9, "bold")).grid(row=0, column=4, sticky="w")
        self.author_var = tk.StringVar(value="")
        ttk.Entry(meta_bar, textvariable=self.author_var, width=12).grid(row=0, column=5, padx=(4, 10))

        tk.Label(meta_bar, text="検索:", bg=LightPalette.BG_SUB, fg=LightPalette.FG_ACCENT, font=("Meiryo UI", 9, "bold")).grid(row=0, column=6, sticky="w")
        self.filter_var = tk.StringVar(value="")
        self.filter_var.trace_add("write", lambda *_: self._refresh_treeview())
        ttk.Entry(meta_bar, textvariable=self.filter_var, width=16).grid(row=0, column=7, padx=4)

        tool_bar = tk.Frame(parent, bg=LightPalette.BG_PANEL, padx=8, pady=5)
        tool_bar.pack(fill=tk.X)
        self._tool_btn(tool_bar, "＋ 行追加", self.action_add_row, "#1D4ED8", True)
        self._tool_btn(tool_bar, "行複製", self.action_duplicate_row)
        self._tool_btn(tool_bar, "－ 行削除", self.action_delete_rows, "#DC2626")
        self._tool_btn(tool_bar, "↑", lambda: self.action_move_row(-1))
        self._tool_btn(tool_bar, "↓", lambda: self.action_move_row(1))
        self._tool_btn(tool_bar, "No.振直", self.action_renumber_rows)
        self._tool_btn(tool_bar, "重複名寄せ統合", self.action_merge_duplicates, "#047857", True)
        self._tool_btn(tool_bar, "寸法から重量計算", self.action_auto_calc_weights, "#1D4ED8")
        self._tool_btn(tool_bar, "重量計算機", self.action_open_weight_calculator, "#7C3AED", True)
        self._tool_btn(tool_bar, "AIデータ補正・補完", self.action_ai_refine_bom, "#6D28D9", True)
        self._tool_btn(tool_bar, "Excel貼付コピー", self.action_copy_selected_to_clipboard, "#0F766E")
        self._tool_btn(tool_bar, "Undo", self.action_undo)
        self._tool_btn(tool_bar, "Redo", self.action_redo)
        self._tool_btn(tool_bar, "全クリア", self.action_clear_all_parts, "#64748B")

        tc = tk.Frame(parent, bg=LightPalette.BG_PANEL, padx=8)
        tc.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tc, columns=self.TREE_COLUMNS, show="headings", selectmode="extended")
        ys = ttk.Scrollbar(tc, orient=tk.VERTICAL, command=self.tree.yview)
        xs = ttk.Scrollbar(tc, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side=tk.RIGHT, fill=tk.Y)
        xs.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        col_defs = [
            ("item_no", "No.", 44, "center"), ("part_number", "図番 / 品番", 115, "w"),
            ("part_name", "品名 / 部品名称", 150, "w"), ("specification", "型式・寸法・仕様", 145, "w"),
            ("material", "材質", 82, "center"), ("treatment", "表面処理・熱処理", 125, "w"),
            ("quantity", "数量", 55, "e"), ("unit", "単位", 46, "center"),
            ("unit_weight", "単重(kg)", 70, "e"), ("total_weight", "総重量(kg)", 76, "e"),
            ("category", "調達区分", 84, "center"), ("manufacturer", "メーカー", 100, "w"),
            ("unit_price", "単価(円)", 75, "e"), ("amount", "金額(円)", 85, "e"),
            ("source_file", "抽出元図面", 110, "w"), ("remarks", "備考", 130, "w"),
        ]
        for cid, h_txt, w_px, anc in col_defs:
            self.tree.heading(cid, text=h_txt, command=lambda c=cid: self._sort_tree_by_column(c))
            self.tree.column(cid, width=w_px, minwidth=38, anchor=anc)
        self.tree.tag_configure("even", background=LightPalette.BG_ROW_EVEN)
        self.tree.tag_configure("odd", background=LightPalette.BG_ROW_ODD)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self._init_context_menus()

        form_f = ttk.LabelFrame(parent, text=" 選択行の詳細編集 (表のセルをダブルクリックしても直接編集できます) ", padding=(10, 5))
        form_f.pack(fill=tk.X, padx=8, pady=5)
        self._build_detail_form(form_f)

        footer = tk.Frame(parent, bg=LightPalette.BG_SUB, padx=12, pady=9, bd=1, relief=tk.SOLID)
        footer.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.summary_lbl = tk.Label(footer, text="総部品行数: 0 件 | 総個数: 0 | 総重量: 0.000 kg | 合計金額: ¥0", bg=LightPalette.BG_SUB, font=("Meiryo UI", 10, "bold"))
        self.summary_lbl.pack(side=tk.LEFT)

        tk.Button(footer, text="Excel部品表を出力 (.xlsx)", command=self.action_export_excel, bg=LightPalette.BTN_SUCCESS_BG, fg="#FFFFFF", relief=tk.FLAT, padx=18, pady=6, font=("Meiryo UI", 10, "bold"), cursor="hand2").pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(footer, text="Excel出力設定...", command=self.action_open_excel_settings, bg="#FFFFFF", fg=LightPalette.FG_ACCENT, relief=tk.SOLID, bd=1, padx=10, pady=5, font=("Meiryo UI", 9, "bold")).pack(side=tk.RIGHT, padx=4)
        tk.Button(footer, text="CSV出力 (.csv)", command=self.action_export_csv, bg="#FFFFFF", relief=tk.SOLID, bd=1, padx=10, pady=5, font=("Meiryo UI", 9)).pack(side=tk.RIGHT, padx=4)

    def _tool_btn(self, parent: tk.Misc, text: str, cmd: Any, fg: str = LightPalette.FG_MAIN, bold: bool = False) -> tk.Button:
        b = tk.Button(parent, text=text, command=cmd, bg="#FFFFFF", fg=fg, relief=tk.SOLID, bd=1, padx=6, pady=3, font=("Meiryo UI", 8, "bold") if bold else ("Meiryo UI", 8), cursor="hand2")
        b.pack(side=tk.LEFT, padx=2)
        return b

    def _build_detail_form(self, parent: ttk.LabelFrame) -> None:
        keys = ["item_no", "part_number", "part_name", "specification", "material", "treatment", "quantity", "unit", "unit_weight", "category", "manufacturer", "unit_price", "remarks"]
        self.form_vars: Dict[str, tk.StringVar] = {k: tk.StringVar(value="") for k in keys}

        r0 = tk.Frame(parent, bg=LightPalette.BG_PANEL)
        r0.pack(fill=tk.X, pady=2)
        self._f_ent(r0, "No:", "item_no", 5)
        self._f_ent(r0, "図番/品番:", "part_number", 15)
        self._f_ent(r0, "品名:", "part_name", 20)
        self._f_ent(r0, "型式・寸法:", "specification", 18)
        self._f_cmb(r0, "材質:", "material", list(MATERIAL_MASTER.keys()), 10)
        self._f_cmb(r0, "表面処理:", "treatment", SURFACE_TREATMENTS, 15)

        r1 = tk.Frame(parent, bg=LightPalette.BG_PANEL)
        r1.pack(fill=tk.X, pady=2)
        self._f_ent(r1, "数量:", "quantity", 6)
        self._f_cmb(r1, "単位:", "unit", UNITS, 5)
        self._f_ent(r1, "単重(kg):", "unit_weight", 8)
        self._f_cmb(r1, "区分:", "category", CATEGORIES, 9)
        self._f_ent(r1, "メーカー:", "manufacturer", 12)
        self._f_ent(r1, "単価(円):", "unit_price", 8)
        self._f_ent(r1, "備考:", "remarks", 15)
        tk.Button(r1, text="選択行を更新", command=self.action_apply_detail_form, bg=LightPalette.BTN_PRIMARY_BG, fg="#FFFFFF", relief=tk.FLAT, padx=10, pady=3, font=("Meiryo UI", 8, "bold")).pack(side=tk.RIGHT, padx=4)

    def _f_ent(self, row_f: tk.Frame, lbl: str, key: str, w: int) -> None:
        tk.Label(row_f, text=lbl, bg=LightPalette.BG_PANEL, fg=LightPalette.FG_SUB, font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT, padx=(3, 1))
        e = ttk.Entry(row_f, textvariable=self.form_vars[key], width=w)
        e.pack(side=tk.LEFT, padx=(0, 5))
        e.bind("<Return>", lambda _e: self.action_apply_detail_form())

    def _f_cmb(self, row_f: tk.Frame, lbl: str, key: str, vals: List[str], w: int) -> None:
        tk.Label(row_f, text=lbl, bg=LightPalette.BG_PANEL, fg=LightPalette.FG_SUB, font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT, padx=(3, 1))
        c = ttk.Combobox(row_f, textvariable=self.form_vars[key], values=vals, width=w)
        c.pack(side=tk.LEFT, padx=(0, 5))
        c.bind("<Return>", lambda _e: self.action_apply_detail_form())

    def _build_statusbar(self) -> None:
        sf = tk.Frame(self.root, bg="#E2E8F0", padx=12, pady=4)
        sf.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_var = tk.StringVar(value="準備完了: 「図面ファイルを開く」または「サンプル図面で試す」をクリックしてください。高精度解析には「Gemini AI設定」をご利用ください。")
        tk.Label(sf, textvariable=self.status_var, bg="#E2E8F0", fg=LightPalette.FG_MAIN, font=("Meiryo UI", 8), anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(sf, text=f"{APP_NAME} (Windows 11 Light Mode)", bg="#E2E8F0", fg=LightPalette.FG_MUTED, font=("Meiryo UI", 8)).pack(side=tk.RIGHT)

    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-o>", lambda _e: self.action_open_files())
        self.root.bind("<Control-s>", lambda _e: self.action_export_excel())
        self.root.bind("<Control-z>", lambda _e: self.action_undo())
        self.root.bind("<Control-y>", lambda _e: self.action_redo())
        self.root.bind("<Control-c>", lambda e: self.action_copy_selected_to_clipboard() if isinstance(e.widget, ttk.Treeview) else None)
        self.root.bind("<F1>", lambda _e: self.action_show_readme())
        self.root.bind("<Delete>", lambda e: self.action_delete_rows() if isinstance(e.widget, ttk.Treeview) else None)

    def _init_context_menus(self) -> None:
        self.tree_menu = tk.Menu(self.root, tearoff=0, bg="#FFFFFF", fg=LightPalette.FG_MAIN, activebackground=LightPalette.BG_SELECTED, activeforeground=LightPalette.FG_ACCENT, font=("Meiryo UI", 9))
        self.tree_menu.add_command(label="＋ 新規行を追加", command=self.action_add_row)
        self.tree_menu.add_command(label="選択行を複製", command=self.action_duplicate_row)
        self.tree_menu.add_command(label="－ 選択行を削除 (Delete)", command=self.action_delete_rows)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="選択行をクリップボードへコピー (Excel貼付用 / Ctrl+C)", command=self.action_copy_selected_to_clipboard)
        self.tree_menu.add_command(label="重量計算機を開く", command=self.action_open_weight_calculator)
        self.tree_menu.add_command(label="寸法から重量を一括自動計算", command=self.action_auto_calc_weights)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="AIデータ補正・補完を実行", command=self.action_ai_refine_bom)

        self.canvas_menu = tk.Menu(self.root, tearoff=0, bg="#FFFFFF", fg=LightPalette.FG_MAIN, activebackground=LightPalette.BG_SELECTED, activeforeground=LightPalette.FG_ACCENT, font=("Meiryo UI", 9))
        self.canvas_menu.add_command(label="全体フィット表示", command=self._fit_canvas_view)
        self.canvas_menu.add_command(label="拡大 (＋)", command=lambda: self._zoom_by(1.25))
        self.canvas_menu.add_command(label="縮小 (－)", command=lambda: self._zoom_by(0.8))
        self.canvas_menu.add_separator()
        self.canvas_menu.add_command(label="範囲選択抽出モード ON/OFF 切替", command=self._toggle_interaction_mode)
        self.canvas_menu.add_command(label="この図面ページを Gemini AI で解析", command=self.action_ai_analyze_current_page)

    def _on_tree_right_click(self, event: tk.Event) -> None:
        row_id = self.tree.identify_row(event.y)
        if row_id and row_id not in self.tree.selection():
            self.tree.selection_set(row_id)
            self._on_tree_select(event)
        try:
            self.tree_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.tree_menu.grab_release()

    def _on_canvas_right_click(self, event: tk.Event) -> None:
        try:
            self.canvas_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.canvas_menu.grab_release()

    def action_copy_selected_to_clipboard(self) -> None:
        if not self.parts:
            self.set_status("コピーする部品データがありません。")
            return
        sel = self.tree.selection()
        if sel:
            sel_set = set(sel)
            targets = [p for p in self.parts if p.uid in sel_set]
        else:
            targets = list(self.parts)
        if not targets:
            return
        lines = [
            "\t".join(["No.", "図番/品番", "品名", "型式・寸法", "材質", "表面処理", "数量", "単位", "単重(kg)", "総重量(kg)", "調達区分", "メーカー", "単価(円)", "金額(円)", "抽出元図面", "備考"])
        ]
        for p in targets:
            lines.append("\t".join([
                str(p.item_no), str(p.part_number), str(p.part_name), str(p.specification),
                str(p.material), str(p.treatment), p.formatted_qty(), str(p.unit),
                f"{p.unit_weight:.3f}", f"{p.total_weight:.3f}", str(p.category),
                str(p.manufacturer), str(int(p.unit_price)), str(int(p.amount)),
                str(p.source_file), str(p.remarks),
            ]))
        tsv_text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(tsv_text)
        self.set_status(f"部品データ {len(targets)} 件をクリップボードにコピーしました（Excel上でそのまま Ctrl+V で貼り付け可能です）。")

    def _push_undo_state(self) -> None:
        self.undo_stack.append(copy.deepcopy(self.parts))
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def action_undo(self) -> None:
        if self.undo_stack:
            self.redo_stack.append(copy.deepcopy(self.parts))
            self.parts = self.undo_stack.pop()
            self._refresh_treeview()
            self.set_status("直前の操作を元に戻しました (Undo)。")

    def action_redo(self) -> None:
        if self.redo_stack:
            self.undo_stack.append(copy.deepcopy(self.parts))
            self.parts = self.redo_stack.pop()
            self._refresh_treeview()
            self.set_status("操作をやり直しました (Redo)。")

    def set_status(self, msg: str) -> None:
        self.status_var.set(msg)
        self.root.update_idletasks()

    # --------------------------------------------------------------------------
    # Gemini AI 高精度図面解析アクション（非同期スレッド実行）
    # --------------------------------------------------------------------------
    def action_open_gemini_settings(self) -> None:
        dlg = GeminiSettingsDialog(self.root, self.gemini_config, self._on_gemini_settings_saved)
        self.root.wait_window(dlg)

    def _ensure_api_key(self) -> bool:
        if not self.gemini_config.api_key.strip():
            if messagebox.askyesno("Gemini APIキー未設定", "Gemini AI解析を使用するにはAPIキーの設定が必要です。\n今すぐ「Gemini AI設定」画面を開きますか？"):
                self.action_open_gemini_settings()
            return bool(self.gemini_config.api_key.strip())
        return True

    def action_ai_analyze_current_page(self) -> None:
        if self.is_ai_busy:
            messagebox.showinfo("AI解析中", "現在Gemini AIによる解析処理を実行中です。完了までお待ちください。")
            return
        page_data = self._get_current_page()
        if not page_data or self.current_doc_idx < 0:
            messagebox.showinfo("図面未選択", "先に図面ファイル (PDF / DXF) を読み込んでください。")
            return
        if not self._ensure_api_key():
            return

        doc = self.documents[self.current_doc_idx]
        self._run_ai_page_analysis_async([(doc, page_data)], crop_rect=None)

    def action_ai_analyze_all_docs(self) -> None:
        if self.is_ai_busy:
            return
        if not self.documents:
            messagebox.showinfo("図面未読込", "先に図面ファイルを読み込んでください。")
            return
        if not self._ensure_api_key():
            return

        targets: List[Tuple[DrawingDocument, DrawingPageData]] = []
        for doc in self.documents:
            for pg in doc.pages:
                targets.append((doc, pg))
        self._run_ai_page_analysis_async(targets, crop_rect=None)

    def _run_ai_page_analysis_async(
        self,
        targets: List[Tuple[DrawingDocument, DrawingPageData]],
        crop_rect: Optional[Tuple[float, float, float, float]] = None,
    ) -> None:
        self.is_ai_busy = True
        self.set_status(f"Gemini AI ({self.gemini_config.model_name}) で図面を高精度解析中... しばらくお待ちください。")

        def worker() -> None:
            results: List[Tuple[DrawingDocument, DrawingPageData, List[PartItem], Dict[str, str]]] = []
            err_msg = ""
            try:
                for doc, pg in targets:
                    parts_out, meta_out = GeminiVisionAnalyzer.analyze_drawing_page(
                        self.gemini_config, pg, doc.file_name, crop_rect=crop_rect
                    )
                    results.append((doc, pg, parts_out, meta_out))
            except Exception as e:
                err_msg = str(e)

            def on_done() -> None:
                if not self.root.winfo_exists():
                    return
                self.is_ai_busy = False
                if err_msg:
                    self.set_status(f"Gemini AI解析エラー: {err_msg}")
                    messagebox.showerror("Gemini AI解析エラー", f"AI解析中にエラーが発生しました:\n{err_msg}")
                    return

                self._push_undo_state()
                total_added = 0
                for doc, pg, parts_out, meta_out in results:
                    if meta_out.get("drawing_title"):
                        doc.drawing_title = meta_out["drawing_title"]
                        if self.proj_name_var.get() in ("", "新規機械設計プロジェクト"):
                            self.proj_name_var.set(doc.drawing_title)
                    if meta_out.get("drawing_number"):
                        doc.drawing_number = meta_out["drawing_number"]
                        if self.dwg_no_var.get() in ("", "ASM-0001"):
                            self.dwg_no_var.set(doc.drawing_number)

                    if crop_rect is None and self.gemini_config.replace_existing_on_ai:
                        self.parts = [
                            p for p in self.parts
                            if not (p.source_file == doc.file_name and p.source_page == pg.page_no)
                        ]
                    self.parts.extend(parts_out)
                    doc.extracted_count = sum(1 for p in self.parts if p.source_file == doc.file_name)
                    total_added += len(parts_out)

                self._refresh_document_listbox()
                self.action_renumber_rows(push_undo=False)
                if crop_rect is not None:
                    self.highlight_bbox = crop_rect
                self._redraw_canvas()
                self.set_status(f"Gemini AI解析完了！ {total_added} 件の部品データを高精度抽出しました。")

            self.root.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()

    def action_ai_refine_bom(self) -> None:
        if self.is_ai_busy:
            return
        if not self.parts:
            messagebox.showinfo("データなし", "補正する部品表データがありません。")
            return
        if not self._ensure_api_key():
            return

        self.is_ai_busy = True
        self.set_status(f"Gemini AI ({self.gemini_config.model_name}) で部品表データの規格正規化・単重/メーカー補完を実行中...")
        snapshot = copy.deepcopy(self.parts)

        def worker() -> None:
            err_msg = ""
            refined: List[PartItem] = []
            try:
                refined = GeminiVisionAnalyzer.refine_bom_parts(self.gemini_config, snapshot)
            except Exception as e:
                err_msg = str(e)

            def on_done() -> None:
                if not self.root.winfo_exists():
                    return
                self.is_ai_busy = False
                if err_msg:
                    self.set_status(f"AI補正エラー: {err_msg}")
                    messagebox.showerror("AIデータ補正エラー", f"AI補正中にエラーが発生しました:\n{err_msg}")
                    return
                self._push_undo_state()
                self.parts = refined
                self._refresh_treeview()
                self.set_status("Gemini AIによる部品表データの正規化・単重/メーカー/区分補完が完了しました！")

            self.root.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()

    # --------------------------------------------------------------------------
    # ファイル読込・サンプル生成
    # --------------------------------------------------------------------------
    def action_open_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="2D図面ファイル (PDF / DXF / 画像) を選択",
            filetypes=[
                ("対応2D図面すべて (*.pdf;*.dxf;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp)", "*.pdf *.dxf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                ("PDF図面 (*.pdf)", "*.pdf"),
                ("CAD DXF図面 (*.dxf)", "*.dxf"),
                ("画像図面 (*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp)", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
            ],
        )
        if paths:
            self._load_drawing_files(list(paths))

    def action_open_folder(self) -> None:
        folder = filedialog.askdirectory(title="図面フォルダを選択")
        if not folder:
            return
        files = [
            os.path.join(folder, f)
            for f in sorted(os.listdir(folder))
            if os.path.splitext(f)[1].lower() in SUPPORTED_DRAWING_EXTENSIONS
        ]
        if not files:
            messagebox.showinfo("フォルダ読込", "指定フォルダ内に対象の図面ファイル (PDF / DXF / 画像) が見つかりませんでした。")
            return
        self._load_drawing_files(files)

    def action_load_sample_drawings(self) -> None:
        self.set_status("サンプル2D図面 (PDF / DXF) を生成しています...")
        try:
            s_dir = os.path.join(tempfile.gettempdir(), "PartLens_Sample_Drawings")
            os.makedirs(s_dir, exist_ok=True)
            p_pdf = SampleDrawingGenerator.create_sample_pdf(os.path.join(s_dir, "Sample_Conveyor_Drive_Unit_CV1001.pdf"))
            p_dxf = SampleDrawingGenerator.create_sample_dxf(os.path.join(s_dir, "Sample_Clamp_Bracket_BRK200.dxf"))
            self._load_drawing_files([p_pdf, p_dxf])
        except Exception as e:
            messagebox.showerror("サンプル生成エラー", str(e))

    def _load_drawing_files(self, file_paths: List[str]) -> None:
        self._push_undo_state()
        new_docs: List[DrawingDocument] = []
        added_parts = 0
        errors: List[str] = []

        for path in file_paths:
            self.set_status(f"図面を読み込み中: {os.path.basename(path)} ...")
            try:
                doc_model, extracted = DrawingParserEngine.parse_file(path)
                self.documents.append(doc_model)
                new_docs.append(doc_model)
                self.parts.extend(extracted)
                added_parts += len(extracted)
                if doc_model.drawing_title and self.proj_name_var.get() in ("", "新規機械設計プロジェクト"):
                    self.proj_name_var.set(doc_model.drawing_title)
                if doc_model.drawing_number and self.dwg_no_var.get() in ("", "ASM-0001"):
                    self.dwg_no_var.set(doc_model.drawing_number)
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {e}")

        self._refresh_document_listbox()
        if self.documents and self.current_doc_idx == -1:
            self.current_doc_idx = len(self.documents) - len(new_docs)
            self.current_page_idx = 0
            self.doc_listbox.selection_clear(0, tk.END)
            self.doc_listbox.selection_set(self.current_doc_idx)

        self.action_renumber_rows(push_undo=False)
        self._fit_canvas_view()
        if errors:
            messagebox.showwarning("読込エラー", "\n".join(errors))

        # APIキー設定済み かつ 自動AI解析がONなら引き続いてGemini高精度解析を実行
        if new_docs and self.gemini_config.api_key.strip() and self.gemini_config.auto_use_on_load:
            targets = [(d, pg) for d in new_docs for pg in d.pages]
            self._run_ai_page_analysis_async(targets, crop_rect=None)
        else:
            self.set_status(f"図面 {len(new_docs)} 件を読込完了 ({added_parts} 件抽出)。さらに精度を高めるには「AI図面解析 (Gemini)」を押してください。")

    def action_clear_documents(self) -> None:
        if self.documents and messagebox.askyesno("確認", "読み込み済みの図面リストをクリアしますか？"):
            self.documents.clear()
            self.current_doc_idx = -1
            self.current_page_idx = 0
            self.highlight_bbox = None
            self._refresh_document_listbox()
            self._redraw_canvas()

    # --------------------------------------------------------------------------
    # プレビューキャンバス操作・矩形範囲選択抽出
    # --------------------------------------------------------------------------
    def _refresh_document_listbox(self) -> None:
        self.doc_listbox.delete(0, tk.END)
        for doc in self.documents:
            lbl = f"[{doc.file_type}] {doc.file_name} ({len(doc.pages)}P / 抽出:{doc.extracted_count}件)"
            if doc.drawing_title:
                lbl += f" - {doc.drawing_title}"
            self.doc_listbox.insert(tk.END, lbl)

    def _on_select_document(self, _e: tk.Event) -> None:
        sel = self.doc_listbox.curselection()
        if sel and 0 <= int(sel[0]) < len(self.documents):
            self.current_doc_idx = int(sel[0])
            self.current_page_idx = 0
            self.highlight_bbox = None
            self._fit_canvas_view()

    def _get_current_page(self) -> Optional[DrawingPageData]:
        if 0 <= self.current_doc_idx < len(self.documents):
            doc = self.documents[self.current_doc_idx]
            if 0 <= self.current_page_idx < len(doc.pages):
                return doc.pages[self.current_page_idx]
        return None

    def _change_page(self, delta: int) -> None:
        if 0 <= self.current_doc_idx < len(self.documents):
            doc = self.documents[self.current_doc_idx]
            if 0 <= self.current_page_idx + delta < len(doc.pages):
                self.current_page_idx += delta
                self.highlight_bbox = None
                self._fit_canvas_view()

    def _zoom_by(self, factor: float) -> None:
        self.zoom_scale = min(max(self.zoom_scale * factor, 0.15), 5.0)
        self._redraw_canvas()

    def _fit_canvas_view(self) -> None:
        pg = self._get_current_page()
        if not pg or not pg.preview_image:
            self._redraw_canvas()
            return
        cw, ch = max(self.preview_canvas.winfo_width(), 400), max(self.preview_canvas.winfo_height(), 300)
        iw, ih = pg.preview_image.size
        self.zoom_scale = min((cw - 24) / max(iw, 1), (ch - 24) / max(ih, 1), 2.0)
        self.canvas_offset_x = (cw - iw * self.zoom_scale) * 0.5
        self.canvas_offset_y = (ch - ih * self.zoom_scale) * 0.5
        self._redraw_canvas()

    def _toggle_interaction_mode(self) -> None:
        if self.interaction_mode == "pan":
            self.interaction_mode = "select_area"
            self.btn_mode_select.configure(text="範囲選択して部品抽出: ON (ドラッグで囲む)", bg=LightPalette.BTN_WARN_BG, fg="#FFFFFF")
            self.preview_canvas.configure(cursor="crosshair")
            self.set_status("範囲選択モードON: 図面上の部品欄や注記をマウス左ドラッグで囲むと、その範囲から部品を抽出します。")
        else:
            self.interaction_mode = "pan"
            self.btn_mode_select.configure(text="範囲選択して部品抽出: OFF", bg="#FFFFFF", fg=LightPalette.FG_ACCENT)
            self.preview_canvas.configure(cursor="fleur")

    def _redraw_canvas(self) -> None:
        self.preview_canvas.delete("all")
        pg = self._get_current_page()
        if not pg or not pg.preview_image:
            self.page_info_lbl.configure(text="Page: 0 / 0")
            self.zoom_lbl.configure(text="100%")
            cw, ch = max(self.preview_canvas.winfo_width(), 400), max(self.preview_canvas.winfo_height(), 300)
            self.preview_canvas.create_text(
                cw // 2, ch // 2,
                text="ここに2D図面 (PDF / DXF) のプレビューが表示されます\n\n「図面ファイルを開く」または「サンプル図面で試す」をクリックしてください",
                fill=LightPalette.FG_MUTED, font=("Meiryo UI", 10, "bold"), justify=tk.CENTER,
            )
            return

        doc = self.documents[self.current_doc_idx]
        self.page_info_lbl.configure(text=f"Page: {self.current_page_idx + 1} / {len(doc.pages)}")
        self.zoom_lbl.configure(text=f"{int(self.zoom_scale * 100)}%")

        iw, ih = pg.preview_image.size
        dw, dh = max(int(iw * self.zoom_scale), 10), max(int(ih * self.zoom_scale), 10)
        self._tk_preview_img = ImageTk.PhotoImage(pg.preview_image.resize((dw, dh), Image.Resampling.BILINEAR))

        ox, oy = int(self.canvas_offset_x), int(self.canvas_offset_y)
        self.preview_canvas.create_image(ox, oy, anchor=tk.NW, image=self._tk_preview_img)
        self.preview_canvas.create_rectangle(ox, oy, ox + dw, oy + dh, outline="#94A3B8", width=1)

        if self.highlight_bbox and any(self.highlight_bbox) and pg.width > 0 and pg.height > 0:
            bx0, by0, bx1, by1 = self.highlight_bbox
            sx0, sy0 = ox + (bx0 / pg.width) * dw, oy + (by0 / pg.height) * dh
            sx1, sy1 = ox + (bx1 / pg.width) * dw, oy + (by1 / pg.height) * dh
            self.preview_canvas.create_rectangle(sx0 - 4, sy0 - 4, sx1 + 4, sy1 + 4, outline=LightPalette.HIGHLIGHT_BOX, width=3)

    def _on_canvas_mouse_down(self, event: tk.Event) -> None:
        self.drag_start_x, self.drag_start_y = event.x, event.y
        if self.interaction_mode == "select_area":
            if self.selection_rect_id:
                self.preview_canvas.delete(self.selection_rect_id)
            self.selection_rect_id = self.preview_canvas.create_rectangle(event.x, event.y, event.x, event.y, outline=LightPalette.SELECT_BOX, width=2, dash=(5, 3))

    def _on_canvas_mouse_move(self, event: tk.Event) -> None:
        if self.interaction_mode == "pan":
            self.canvas_offset_x += event.x - self.drag_start_x
            self.canvas_offset_y += event.y - self.drag_start_y
            self.drag_start_x, self.drag_start_y = event.x, event.y
            self._redraw_canvas()
        elif self.interaction_mode == "select_area" and self.selection_rect_id:
            self.preview_canvas.coords(self.selection_rect_id, self.drag_start_x, self.drag_start_y, event.x, event.y)

    def _on_canvas_mouse_up(self, event: tk.Event) -> None:
        if self.interaction_mode != "select_area":
            return
        if self.selection_rect_id:
            self.preview_canvas.delete(self.selection_rect_id)
            self.selection_rect_id = None

        pg = self._get_current_page()
        if not pg or not pg.preview_image:
            return
        x0_s, y0_s = min(self.drag_start_x, event.x), min(self.drag_start_y, event.y)
        x1_s, y1_s = max(self.drag_start_x, event.x), max(self.drag_start_y, event.y)
        if (x1_s - x0_s) < 12 or (y1_s - y0_s) < 10:
            return

        iw, ih = pg.preview_image.size
        dw, dh = max(iw * self.zoom_scale, 1.0), max(ih * self.zoom_scale, 1.0)
        px0 = max(0.0, ((x0_s - self.canvas_offset_x) / dw) * pg.width)
        py0 = max(0.0, ((y0_s - self.canvas_offset_y) / dh) * pg.height)
        px1 = min(pg.width, ((x1_s - self.canvas_offset_x) / dw) * pg.width)
        py1 = min(pg.height, ((y1_s - self.canvas_offset_y) / dh) * pg.height)

        doc = self.documents[self.current_doc_idx]

        # Gemini APIキーが設定されており範囲AI抽出が有効なら、選択範囲のクロップ画像＋テキストでGemini解析を実行
        if self.gemini_config.api_key.strip() and self.gemini_config.use_for_area_select:
            self._run_ai_page_analysis_async([(doc, pg)], crop_rect=(px0, py0, px1, py1))
            return

        extracted = DrawingParserEngine.extract_parts_from_tokens(pg.tokens, doc.file_name, pg.page_no, (px0, py0, px1, py1), force_all_rows=True)
        if not extracted:
            messagebox.showinfo("範囲選択抽出", "選択範囲内にテキストが見つかりませんでした。\nスキャン図面の場合は「Gemini AI設定」でAPIキーを設定すると画像認識で読み取れます。")
            return
        self._push_undo_state()
        self.parts.extend(extracted)
        doc.extracted_count += len(extracted)
        self._refresh_document_listbox()
        self.action_renumber_rows(push_undo=False)
        self.highlight_bbox = (px0, py0, px1, py1)
        self._redraw_canvas()
        self.set_status(f"選択範囲から {len(extracted)} 件の部品データを抽出し追加しました。")

    # --------------------------------------------------------------------------
    # Treeview 表示・編集・行操作
    # --------------------------------------------------------------------------
    def _recalculate_doc_counts(self) -> None:
        if not self.documents:
            return
        changed = False
        for doc in self.documents:
            cnt = sum(1 for p in self.parts if p.source_file == doc.file_name)
            if doc.extracted_count != cnt:
                doc.extracted_count = cnt
                changed = True
        if changed:
            self._refresh_document_listbox()
            if 0 <= self.current_doc_idx < len(self.documents):
                self.doc_listbox.selection_clear(0, tk.END)
                self.doc_listbox.selection_set(self.current_doc_idx)

    def _refresh_treeview(self) -> None:
        self._close_inline_editor()
        self._recalculate_doc_counts()
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        kw = self.filter_var.get().strip().lower()
        vis_i = 0
        for p in self.parts:
            if kw and kw not in f"{p.part_number} {p.part_name} {p.specification} {p.material} {p.treatment} {p.manufacturer} {p.remarks}".lower():
                continue
            tag = "even" if vis_i % 2 == 0 else "odd"
            vis_i += 1
            vals = (
                p.item_no, p.part_number, p.part_name, p.specification, p.material, p.treatment,
                p.formatted_qty(), p.unit, f"{p.unit_weight:.3f}", f"{p.total_weight:.3f}",
                p.category, p.manufacturer, f"{int(p.unit_price):,}", f"{int(p.amount):,}", p.source_file, p.remarks,
            )
            self.tree.insert("", tk.END, iid=p.uid, values=vals, tags=(tag,))
        self._update_summary_bar()

    def _update_summary_bar(self) -> None:
        tq = sum(p.quantity for p in self.parts)
        tw = sum(p.total_weight for p in self.parts)
        ta = sum(p.amount for p in self.parts)
        qs = str(int(tq)) if float(tq).is_integer() else f"{tq:.2f}"
        self.summary_lbl.configure(text=f"総部品行数: {len(self.parts)} 件   |   総個数: {qs}   |   総重量: {tw:.3f} kg   |   合計金額: ¥{int(ta):,}")

    def _find_part_by_uid(self, uid: str) -> Optional[PartItem]:
        for p in self.parts:
            if p.uid == uid:
                return p
        return None

    def _on_tree_select(self, _e: tk.Event) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        part = self._find_part_by_uid(sel[0])
        if not part:
            return
        self.selected_part_uid = part.uid
        for k in ["item_no", "part_number", "part_name", "specification", "material", "treatment", "unit", "category", "manufacturer", "remarks"]:
            self.form_vars[k].set(str(getattr(part, k, "")))
        self.form_vars["quantity"].set(part.formatted_qty())
        self.form_vars["unit_weight"].set(f"{part.unit_weight:.3f}")
        self.form_vars["unit_price"].set(str(int(part.unit_price)) if float(part.unit_price).is_integer() else str(part.unit_price))

        if part.source_file:
            for d_i, doc in enumerate(self.documents):
                if doc.file_name == part.source_file:
                    self.current_doc_idx = d_i
                    self.current_page_idx = max(min(part.source_page - 1, len(doc.pages) - 1), 0)
                    break
        self.highlight_bbox = part.bbox
        self._redraw_canvas()

    def action_apply_detail_form(self) -> None:
        if not self.selected_part_uid:
            return
        part = self._find_part_by_uid(self.selected_part_uid)
        if not part:
            return
        self._push_undo_state()
        for k in ["item_no", "part_number", "part_name", "specification", "material", "treatment", "unit", "category", "manufacturer", "remarks"]:
            setattr(part, k, self.form_vars[k].get().strip())
        for nk in ["quantity", "unit_weight", "unit_price"]:
            try:
                setattr(part, nk, max(float(self.form_vars[nk].get().replace(",", "")), 0.0))
            except ValueError:
                pass
        if part.unit_weight == 0.0 and part.specification:
            part.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(part.specification, part.material)
        self._refresh_treeview()
        if self.tree.exists(part.uid):
            self.tree.selection_set(part.uid)
        self.set_status(f"No.{part.item_no} 「{part.part_name}」を更新しました。")

    def _close_inline_editor(self) -> None:
        if self._inline_editor is not None:
            try:
                self._inline_editor.destroy()
            except Exception:
                pass
            self._inline_editor = None

    def _on_tree_double_click(self, event: tk.Event) -> None:
        self._close_inline_editor()
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_idx = int(col_id.replace("#", "")) - 1
        if not (0 <= col_idx < len(self.TREE_COLUMNS)):
            return
        fn = self.TREE_COLUMNS[col_idx]
        if fn == "total_weight":
            fn, col_id = "unit_weight", f"#{self.TREE_COLUMNS.index('unit_weight') + 1}"
        elif fn == "amount":
            fn, col_id = "unit_price", f"#{self.TREE_COLUMNS.index('unit_price') + 1}"

        bbox = self.tree.bbox(row_id, col_id)
        part = self._find_part_by_uid(row_id)
        if not bbox or not part:
            return
        bx, by, bw, bh = bbox
        var = tk.StringVar(value=part.formatted_qty() if fn == "quantity" else str(getattr(part, fn, "")))

        if fn == "material":
            ed: tk.Widget = ttk.Combobox(self.tree, textvariable=var, values=list(MATERIAL_MASTER.keys()))
        elif fn == "treatment":
            ed = ttk.Combobox(self.tree, textvariable=var, values=SURFACE_TREATMENTS)
        elif fn == "category":
            ed = ttk.Combobox(self.tree, textvariable=var, values=CATEGORIES)
        elif fn == "unit":
            ed = ttk.Combobox(self.tree, textvariable=var, values=UNITS)
        else:
            ed = tk.Entry(self.tree, textvariable=var, bg="#FEF9C3", fg=LightPalette.FG_MAIN, relief=tk.SOLID, bd=1, font=("Meiryo UI", 9))

        ed.place(x=bx, y=by, width=max(bw, 70), height=bh)
        ed.focus_set()
        if isinstance(ed, tk.Entry):
            ed.select_range(0, tk.END)
        self._inline_editor = ed

        def commit(_e: Optional[tk.Event] = None) -> None:
            if self._inline_editor is None:
                return
            nv = var.get().strip()
            self._close_inline_editor()
            self._push_undo_state()
            if fn in ("quantity", "unit_weight", "unit_price"):
                try:
                    setattr(part, fn, max(float(nv.replace(",", "")), 0.0))
                except ValueError:
                    pass
            else:
                setattr(part, fn, nv)
            if fn in ("specification", "material") and part.unit_weight == 0.0 and part.specification:
                part.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(part.specification, part.material)
            self._refresh_treeview()
            if self.tree.exists(part.uid):
                self.tree.selection_set(part.uid)
                self._on_tree_select(event)

        ed.bind("<Return>", commit)
        if isinstance(ed, ttk.Combobox):
            ed.bind("<<ComboboxSelected>>", commit)
        else:
            ed.bind("<FocusOut>", commit)
        ed.bind("<Escape>", lambda _e: self._close_inline_editor())

    def _sort_tree_by_column(self, col_name: str) -> None:
        if not self.parts:
            return
        self._push_undo_state()
        desc = self._sort_col_desc.get(col_name, False)
        self._sort_col_desc[col_name] = not desc

        def sort_key(p: PartItem) -> Any:
            if col_name == "item_no":
                return int(p.item_no) if str(p.item_no).isdigit() else 999999
            if col_name in ("quantity", "unit_weight", "total_weight", "unit_price", "amount"):
                return float(getattr(p, col_name, 0.0) or 0.0)
            return str(getattr(p, col_name, "") or "").lower()

        self.parts.sort(key=sort_key, reverse=desc)
        self._refresh_treeview()
        order_str = "降順 (▼)" if desc else "昇順 (▲)"
        self.set_status(f"列「{col_name}」で {order_str} に並び替えました。")

    def action_add_row(self) -> None:
        self._push_undo_state()
        src = self.documents[self.current_doc_idx].file_name if 0 <= self.current_doc_idx < len(self.documents) else ""
        ni = PartItem(item_no=str(len(self.parts) + 1), part_number=f"P-{len(self.parts) + 1:03d}", part_name="新規部品", material="SS400", treatment="四三酸化鉄皮膜(黒染め)", source_file=src)
        self.parts.append(ni)
        self._refresh_treeview()
        self.tree.selection_set(ni.uid)
        self.tree.see(ni.uid)
        self.selected_part_uid = ni.uid

    def action_duplicate_row(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        part = self._find_part_by_uid(sel[0])
        if part:
            self._push_undo_state()
            dup = copy.deepcopy(part)
            dup.uid = str(uuid.uuid4())
            self.parts.insert(self.parts.index(part) + 1, dup)
            self.action_renumber_rows(push_undo=False)

    def action_delete_rows(self) -> None:
        sel = self.tree.selection()
        if sel:
            self._push_undo_state()
            del_set = set(sel)
            self.parts = [p for p in self.parts if p.uid not in del_set]
            self.selected_part_uid = None
            self._refresh_treeview()

    def action_move_row(self, direction: int) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        part = self._find_part_by_uid(sel[0])
        if part:
            idx = self.parts.index(part)
            n_idx = idx + direction
            if 0 <= n_idx < len(self.parts):
                self._push_undo_state()
                self.parts[idx], self.parts[n_idx] = self.parts[n_idx], self.parts[idx]
                self._refresh_treeview()
                self.tree.selection_set(part.uid)

    def action_renumber_rows(self, push_undo: bool = True) -> None:
        if push_undo and self.parts:
            self._push_undo_state()
        for i, p in enumerate(self.parts, start=1):
            p.item_no = str(i)
        self._refresh_treeview()

    def action_merge_duplicates(self) -> None:
        if len(self.parts) >= 2:
            self._push_undo_state()
            self.parts = ExcelBOMExporter.merge_duplicate_parts(self.parts)
            self._refresh_treeview()
            self.set_status("重複部品の名寄せ統合を実行しました。")

    def action_auto_calc_weights(self) -> None:
        if not self.parts:
            return
        self._push_undo_state()
        cnt = 0
        for p in self.parts:
            w = SmartPartAnalyzer.estimate_weight_from_spec(p.specification, p.material)
            if w > 0:
                p.unit_weight = w
                cnt += 1
        self._refresh_treeview()
        self.set_status(f"寸法と材質から {cnt} 件の単重(kg)を自動計算しました。")

    def action_open_weight_calculator(self) -> None:
        init_m = "SS400"
        if self.selected_part_uid:
            p = self._find_part_by_uid(self.selected_part_uid)
            if p and p.material:
                init_m = p.material

        def on_apply(w_kg: float, spec_s: str, mat_s: str) -> None:
            if self.selected_part_uid:
                p = self._find_part_by_uid(self.selected_part_uid)
                if p:
                    self._push_undo_state()
                    p.unit_weight = w_kg
                    if not p.specification:
                        p.specification = spec_s
                    if not p.material:
                        p.material = mat_s
                    self._refresh_treeview()
        WeightCalculatorDialog(self.root, init_m, on_apply)

    def action_clear_all_parts(self) -> None:
        if self.parts and messagebox.askyesno("全クリア", "部品表データをすべてクリアしますか？"):
            self._push_undo_state()
            self.parts.clear()
            self.selected_part_uid = None
            self._refresh_treeview()

    # --------------------------------------------------------------------------
    # Excel / CSV 出力・プロジェクト保存/読込
    # --------------------------------------------------------------------------
    def action_open_excel_settings(self) -> None:
        ExcelSettingsDialog(self.root, self.excel_config, lambda: self.set_status("Excel出力設定を更新しました。"))

    def action_export_excel(self) -> None:
        if not self.parts:
            messagebox.showwarning("データなし", "出力する部品データがありません。")
            return
        self.excel_config.project_name = self.proj_name_var.get().strip() or "機械設計部品表"
        self.excel_config.drawing_no = self.dwg_no_var.get().strip()
        self.excel_config.author_name = self.author_var.get().strip()

        def_name = re.sub(r"[\\/*?:\"<>|]", "_", f"部品表_{self.excel_config.drawing_no or 'BOM'}_{datetime.date.today().strftime('%Y%m%d')}.xlsx")
        out_path = filedialog.asksaveasfilename(title="Excel部品表 (.xlsx) の保存", defaultextension=".xlsx", initialfile=def_name, filetypes=[("Excel ブック (*.xlsx)", "*.xlsx")])
        if not out_path:
            return
        try:
            ExcelBOMExporter.export_to_excel(out_path, self.parts, self.documents, self.excel_config)
            self.set_status(f"Excel部品表を出力しました: {out_path}")
            if messagebox.askyesno("出力完了", f"Excel部品表の出力が完了しました！\n\n{out_path}\n\n今すぐファイルを開きますか？"):
                try:
                    os.startfile(out_path)
                except Exception:
                    pass
        except PermissionError:
            messagebox.showerror(
                "ファイル使用中エラー",
                f"保存先のExcelファイルが他のソフト（Excelなど）で開かれています。\n\n{out_path}\n\nファイルを閉じてから、もう一度出力してください。",
            )
        except Exception as e:
            messagebox.showerror("Excel出力エラー", f"保存中にエラーが発生しました:\n{e}")

    def action_export_csv(self) -> None:
        if not self.parts:
            messagebox.showwarning("データなし", "出力する部品データがありません。")
            return
        out_path = filedialog.asksaveasfilename(title="CSV保存", defaultextension=".csv", initialfile="部品表_BOM.csv", filetypes=[("CSV (*.csv)", "*.csv")])
        if not out_path:
            return
        try:
            with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["No.", "図番/品番", "品名", "型式・寸法", "材質", "表面処理", "数量", "単位", "単重(kg)", "総重量(kg)", "調達区分", "メーカー", "単価", "金額", "抽出元図面", "備考"])
                for p in self.parts:
                    w.writerow([p.item_no, p.part_number, p.part_name, p.specification, p.material, p.treatment, p.formatted_qty(), p.unit, p.unit_weight, p.total_weight, p.category, p.manufacturer, p.unit_price, p.amount, p.source_file, p.remarks])
            messagebox.showinfo("CSV出力完了", f"CSVを保存しました:\n{out_path}")
        except PermissionError:
            messagebox.showerror(
                "ファイル使用中エラー",
                f"保存先のCSVファイルが他のソフト（Excelなど）で開かれています。\n\n{out_path}\n\nファイルを閉じてから、もう一度保存してください。",
            )
        except Exception as e:
            messagebox.showerror("CSVエラー", str(e))

    def action_save_project(self) -> None:
        out_path = filedialog.asksaveasfilename(title="プロジェクト保存", defaultextension=".plens.json", initialfile="Project.plens.json", filetypes=[("PartLens Project (*.plens.json)", "*.plens.json")])
        if not out_path:
            return
        data = {
            "app_name": APP_NAME,
            "project_name": self.proj_name_var.get(),
            "drawing_no": self.dwg_no_var.get(),
            "author_name": self.author_var.get(),
            "parts": [p.to_dict() for p in self.parts],
            "source_files": [d.file_path for d in self.documents if os.path.exists(d.file_path)],
        }
        try:
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存完了", f"プロジェクトを保存しました:\n{out_path}")
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))

    def action_load_project(self) -> None:
        in_path = filedialog.askopenfilename(title="プロジェクト読込", filetypes=[("PartLens Project (*.plens.json;*.json)", "*.plens.json *.json")])
        if not in_path:
            return
        try:
            with open(in_path, "r", encoding="utf-8") as f:
                d = json.load(f)
            self._push_undo_state()
            self.proj_name_var.set(str(d.get("project_name", "")))
            self.dwg_no_var.set(str(d.get("drawing_no", "")))
            self.author_var.set(str(d.get("author_name", "")))
            self.parts = [PartItem.from_dict(x) for x in d.get("parts", [])]
            self.documents.clear()
            for fp in d.get("source_files", []):
                if os.path.exists(fp):
                    try:
                        dm, _ = DrawingParserEngine.parse_file(fp)
                        self.documents.append(dm)
                    except Exception:
                        pass
            self._refresh_document_listbox()
            if self.documents:
                self.current_doc_idx = 0
                self._fit_canvas_view()
            self._refresh_treeview()
        except Exception as e:
            messagebox.showerror("読込エラー", str(e))

    def action_show_readme(self) -> None:
        ReadmeViewerDialog(self.root)


def main() -> None:
    enable_high_dpi_awareness()
    root = tk.Tk()
    PartLensApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
