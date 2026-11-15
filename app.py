# -*- coding: utf-8 -*-
"""
PartLens_Miya-v1.0.0
2D図面 (PDF / CAD DXF) から部品表 (BOM) を作成・編集し Excel (.xlsx) へ出力するデスクトップアプリケーション
Windows 11 ライトモード完全固定対応
"""

import os
import sys
import re
import csv
import json
import math
import copy
import uuid
import ctypes
import tempfile
import datetime
import traceback
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


# ==============================================================================
# アプリケーション基本定数・カラーパレット（ライトモード完全固定）
# ==============================================================================

APP_NAME = "PartLens_Miya-v1.0.0"
APP_TITLE = "PartLens_Miya-v1.0.0 - 2D図面(PDF/CAD) 部品表(BOM)作成システム"
APP_VERSION = "1.0.0"


class LightPalette:
    """システムがダークモードであっても常にライトモードで描画するためのカラー定義"""
    BG_MAIN = "#F1F5F9"         # 全体背景（明るいスレートグレー）
    BG_PANEL = "#FFFFFF"        # パネル・カード背景（純白）
    BG_SUB = "#F8FAFC"          # サブパネル・ツールバー背景
    BG_HEADER = "#1E3A8A"       # ヘッダーバー背景（ディープロイヤルブルー）
    BG_CANVAS = "#E2E8F0"       # プレビューキャンバス背景
    BG_ROW_EVEN = "#FFFFFF"     # テーブル偶数行
    BG_ROW_ODD = "#F8FAFC"      # テーブル奇数行
    BG_SELECTED = "#DBEAFE"     # 選択行背景（ライトブルー）

    FG_MAIN = "#0F172A"         # メイン文字色（濃紺ブラック）
    FG_SUB = "#475569"          # サブ文字色（スレートグレー）
    FG_MUTED = "#64748B"        # 補足文字色
    FG_HEADER = "#FFFFFF"       # ヘッダー文字色（白）
    FG_ACCENT = "#1D4ED8"       # アクセント文字色

    BTN_PRIMARY_BG = "#2563EB"  # プライマリボタン背景
    BTN_PRIMARY_ACTIVE = "#1D4ED8"
    BTN_PRIMARY_FG = "#FFFFFF"

    BTN_SUCCESS_BG = "#059669"  # Excel出力ボタン等のグリーン
    BTN_SUCCESS_ACTIVE = "#047857"
    BTN_SUCCESS_FG = "#FFFFFF"

    BTN_WARN_BG = "#D97706"     # 範囲選択などのアクセントオレンジ
    BTN_WARN_ACTIVE = "#B45309"
    BTN_WARN_FG = "#FFFFFF"

    BTN_DEFAULT_BG = "#FFFFFF"  # 標準ボタン背景
    BTN_DEFAULT_ACTIVE = "#E2E8F0"
    BTN_DEFAULT_FG = "#1E293B"

    BORDER = "#CBD5E1"          # 境界線色
    BORDER_FOCUS = "#3B82F6"    # フォーカス枠色
    HIGHLIGHT_BOX = "#F59E0B"   # 図面ハイライト枠色（アンバー）
    SELECT_BOX = "#EF4444"      # マウス範囲選択枠色（レッド）


def get_resource_path(relative_path: str) -> str:
    """PyInstallerバンドル時(_MEIPASS)および通常実行時のリソースパスを解決する"""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def enforce_windows_light_mode(window: tk.Misc) -> None:
    """Windows 11/10 のシステム設定に関わらずウィンドウをライトモードに固定する"""
    if sys.platform == "win32":
        try:
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            if not hwnd:
                hwnd = window.winfo_id()
            # DWMWA_USE_IMMERSIVE_DARK_MODE = 20 -> 0 (FALSE: ライトモード強制)
            value = ctypes.c_int(0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 20, ctypes.byref(value), ctypes.sizeof(value)
            )
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 19, ctypes.byref(value), ctypes.sizeof(value)
            )
        except Exception:
            pass


def enable_high_dpi_awareness() -> None:
    """Windows 11 高DPI環境での文字ぼやけを防止する"""
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


# ==============================================================================
# JIS規格材質辞書・表面処理・標準部品マスターデータ
# ==============================================================================

MATERIAL_MASTER: Dict[str, Dict[str, Any]] = {
    "SS400": {"category": "一般構造用圧延鋼材", "density": 7.85, "default_treat": "四三酸化鉄皮膜"},
    "S45C": {"category": "機械構造用炭素鋼", "density": 7.85, "default_treat": "四三酸化鉄皮膜"},
    "S50C": {"category": "機械構造用炭素鋼", "density": 7.85, "default_treat": "四三酸化鉄皮膜"},
    "S35C": {"category": "機械構造用炭素鋼", "density": 7.85, "default_treat": "四三酸化鉄皮膜"},
    "SCM435": {"category": "クロムモリブデン鋼", "density": 7.85, "default_treat": "調質・四三酸化鉄皮膜"},
    "SCM440": {"category": "クロムモリブデン鋼", "density": 7.85, "default_treat": "調質・四三酸化鉄皮膜"},
    "SCM415": {"category": "クロムモリブデン鋼", "density": 7.85, "default_treat": "浸炭焼入れ研磨"},
    "SKD11": {"category": "合金工具鋼", "density": 7.85, "default_treat": "焼入れ焼戻し(HRC58-60)"},
    "SKH51": {"category": "高速度工具鋼(ハイス)", "density": 8.15, "default_treat": "焼入れ焼戻し"},
    "SK85": {"category": "炭素工具鋼(旧SK5)", "density": 7.85, "default_treat": "焼入れ"},
    "SPCC": {"category": "冷間圧延鋼板", "density": 7.85, "default_treat": "ユニクロめっき"},
    "SPHC": {"category": "熱間圧延軟鋼板", "density": 7.85, "default_treat": "塗装"},
    "SECC": {"category": "電気亜鉛めっき鋼板", "density": 7.85, "default_treat": "生地"},
    "SGP": {"category": "配管用炭素鋼鋼管", "density": 7.85, "default_treat": "塗装"},
    "STKM13A": {"category": "機械構造用炭素鋼鋼管", "density": 7.85, "default_treat": "四三酸化鉄皮膜"},
    "STKR400": {"category": "一般構造用角形鋼管", "density": 7.85, "default_treat": "塗装"},
    "FC200": {"category": "ねずみ鋳鉄", "density": 7.20, "default_treat": "塗装"},
    "FC250": {"category": "ねずみ鋳鉄", "density": 7.25, "default_treat": "塗装"},
    "FCD450": {"category": "球状黒鉛鋳鉄", "density": 7.10, "default_treat": "塗装"},
    "SUS304": {"category": "オーステナイト系ステンレス", "density": 7.93, "default_treat": "生地(酸洗)"},
    "SUS303": {"category": "快削ステンレス", "density": 7.93, "default_treat": "生地"},
    "SUS316": {"category": "耐食ステンレス", "density": 7.98, "default_treat": "生地"},
    "SUS316L": {"category": "極低炭素ステンレス", "density": 7.98, "default_treat": "生地"},
    "SUS430": {"category": "フェライト系ステンレス", "density": 7.70, "default_treat": "生地"},
    "SUS440C": {"category": "マルテンサイト系ステンレス", "density": 7.75, "default_treat": "焼入れ"},
    "SUS630": {"category": "析出硬化系ステンレス", "density": 7.78, "default_treat": "H900熱処理"},
    "A5052": {"category": "アルミニウム合金", "density": 2.68, "default_treat": "白アルマイト"},
    "A5056": {"category": "アルミニウム合金(丸棒等)", "density": 2.64, "default_treat": "白アルマイト"},
    "A6061": {"category": "構造用アルミニウム合金", "density": 2.70, "default_treat": "白アルマイト"},
    "A6063": {"category": "押出用アルミニウム合金", "density": 2.70, "default_treat": "アルマイト"},
    "A2017": {"category": "ジュラルミン", "density": 2.79, "default_treat": "白アルマイト"},
    "A7075": {"category": "超々ジュラルミン", "density": 2.80, "default_treat": "硬質アルマイト"},
    "ADC12": {"category": "アルミダイカスト", "density": 2.70, "default_treat": "生地"},
    "C3604": {"category": "快削黄銅(真鍮)", "density": 8.50, "default_treat": "生地"},
    "C2801": {"category": "黄銅板", "density": 8.40, "default_treat": "生地"},
    "C1100": {"category": "タフピッチ銅", "density": 8.89, "default_treat": "生地"},
    "C1020": {"category": "無酸素銅", "density": 8.94, "default_treat": "生地"},
    "CAC406": {"category": "青銅鋳物(BC6)", "density": 8.80, "default_treat": "生地"},
    "POM": {"category": "ポリアセタール(ジュラコン)", "density": 1.41, "default_treat": "生地"},
    "MCナイロン": {"category": "モノマーキャストナイロン", "density": 1.16, "default_treat": "生地"},
    "66ナイロン": {"category": "ポリアミド樹脂", "density": 1.14, "default_treat": "生地"},
    "PEEK": {"category": "ポリエーテルエーテルケトン", "density": 1.30, "default_treat": "生地"},
    "PTFE": {"category": "四フッ化エチレン(テフロン)", "density": 2.16, "default_treat": "生地"},
    "UHMW-PE": {"category": "超高分子量ポリエチレン", "density": 0.94, "default_treat": "生地"},
    "PC": {"category": "ポリカーボネート", "density": 1.20, "default_treat": "透明"},
    "ABS": {"category": "ABS樹脂", "density": 1.05, "default_treat": "生地"},
    "PMMA": {"category": "アクリル樹脂", "density": 1.19, "default_treat": "透明"},
    "PVC": {"category": "硬質塩化ビニル", "density": 1.40, "default_treat": "生地"},
    "ウレタン": {"category": "ポリウレタン(PU)", "density": 1.20, "default_treat": "生地"},
    "NBR": {"category": "ニトリルゴム", "density": 1.20, "default_treat": "生地"},
    "CRゴム": {"category": "クロロプレンゴム", "density": 1.25, "default_treat": "生地"},
    "シリコンゴム": {"category": "シリコーンゴム", "density": 1.18, "default_treat": "生地"},
}

SURFACE_TREATMENTS: List[str] = [
    "生地",
    "四三酸化鉄皮膜(黒染め)",
    "無電解ニッケルめっき",
    "硬質クロムめっき",
    "三価クロメートめっき",
    "ユニクロめっき",
    "溶融亜鉛めっき",
    "白アルマイト",
    "黒アルマイト",
    "硬質アルマイト",
    "焼入れ研磨",
    "高周波焼入れ",
    "浸炭焼入れ",
    "窒化処理",
    "調質(QT)",
    "メラミン焼付塗装",
    "粉体塗装",
    "ウレタン塗装",
    "錆止め塗装",
    "酸洗・不動態化処理",
    "電解研磨",
    "バフ研磨",
]

CATEGORIES: List[str] = [
    "加工品",
    "購入品",
    "標準部品",
    "製缶・溶接品",
    "板金品",
    "支給品",
    "組立品",
]

UNITS: List[str] = ["個", "本", "枚", "式", "セット", "組", "m", "kg"]

STANDARD_PART_KEYWORDS: List[str] = [
    "ボルト", "六角穴付", "キャップボルト", "ナット", "ワッシャ", "座金", "ねじ", "ビス",
    "ピン", "平行ピン", "ノックピン", "割ピン", "スプリングピン",
    "平行キー", "止め輪", "C形止め輪", "Eリング",
    "ベアリング", "軸受", "玉軸受", "カムフォロア", "ピローブロック", "ベアリングホルダ",
    "リニアガイド", "ボールねじ", "リニアブッシュ", "シャフトホルダ",
    "Oリング", "オイルシール", "パッキン", "ガスケット",
    "シリンダ", "エアシリンダ", "モータ", "ステッピングモータ", "サーボモータ", "減速機",
    "タイミングベルト", "プーリ", "スプロケット", "チェーン", "歯車", "ギヤ", "カップリング",
    "センサ", "フォトセンサ", "近接スイッチ", "スイッチ", "ケーブルベア",
    "キャスター", "アジャスタ", "取っ手", "ハンドル", "ヒンジ", "蝶番", "マグネット",
]


# ==============================================================================
# データモデル層
# ==============================================================================

@dataclass
class PartItem:
    """部品表(BOM)の1行を表すデータクラス"""
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
        if float(self.quantity).is_integer():
            return str(int(self.quantity))
        return f"{self.quantity:.2f}".rstrip("0").rstrip(".")

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["bbox"] = list(self.bbox)
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PartItem":
        bbox_val = data.get("bbox", [0.0, 0.0, 0.0, 0.0])
        if isinstance(bbox_val, (list, tuple)) and len(bbox_val) == 4:
            bbox_tuple = (float(bbox_val[0]), float(bbox_val[1]), float(bbox_val[2]), float(bbox_val[3]))
        else:
            bbox_tuple = (0.0, 0.0, 0.0, 0.0)
        return cls(
            uid=str(data.get("uid", str(uuid.uuid4()))),
            item_no=str(data.get("item_no", "1")),
            part_number=str(data.get("part_number", "")),
            part_name=str(data.get("part_name", "")),
            specification=str(data.get("specification", "")),
            material=str(data.get("material", "")),
            treatment=str(data.get("treatment", "")),
            quantity=float(data.get("quantity", 1.0)),
            unit=str(data.get("unit", "個")),
            unit_weight=float(data.get("unit_weight", 0.0)),
            category=str(data.get("category", "加工品")),
            manufacturer=str(data.get("manufacturer", "")),
            unit_price=float(data.get("unit_price", 0.0)),
            source_file=str(data.get("source_file", "")),
            source_page=int(data.get("source_page", 1)),
            remarks=str(data.get("remarks", "")),
            bbox=bbox_tuple,
        )


@dataclass
class TextToken:
    """図面内の位置座標付きテキスト要素"""
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
    """図面の1ページ（または1レイアウト）の解析データとプレビュー画像"""
    page_no: int
    width: float
    height: float
    preview_image: Optional[Image.Image] = None
    tokens: List[TextToken] = field(default_factory=list)
    raw_tables: List[List[List[str]]] = field(default_factory=list)
    title_meta: Dict[str, str] = field(default_factory=dict)


@dataclass
class DrawingDocument:
    """読み込まれた1つの図面ファイル(PDF/DXF)"""
    file_path: str
    file_name: str
    file_type: str  # "PDF" or "DXF"
    pages: List[DrawingPageData] = field(default_factory=list)
    extracted_count: int = 0
    drawing_title: str = ""
    drawing_number: str = ""


# ==============================================================================
# 重量計算・スマート部品テキスト解析エンジン
# ==============================================================================

class SmartPartAnalyzer:
    """図面テキストから品番・品名・寸法・材質・表面処理・数量・単重を高精度に判定するエンジン"""

    @staticmethod
    def normalize_text(text: str) -> str:
        if not text:
            return ""
        # 全角英数字・記号の一部を半角に統一
        zen = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ－（）×φΦ"
        han = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-()xφφ"
        trans = str.maketrans(zen, han)
        cleaned = text.translate(trans)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        return cleaned

    @classmethod
    def detect_material(cls, text: str) -> Tuple[str, str]:
        """文字列から材質と残りの文字列を抽出する"""
        norm = cls.normalize_text(text)
        upper = norm.upper()
        # 長いキー名から優先マッチング (SUS316L -> SUS316 等)
        sorted_keys = sorted(MATERIAL_MASTER.keys(), key=len, reverse=True)
        for mat_key in sorted_keys:
            pattern = re.compile(re.escape(mat_key), re.IGNORECASE)
            if pattern.search(upper):
                return mat_key, norm
        # よくある別名表記
        alias_map = {
            "ジュラコン": "POM",
            "ポリアセタール": "POM",
            "テフロン": "PTFE",
            "アクリル": "PMMA",
            "塩ビ": "PVC",
            "真鍮": "C3604",
            "黄銅": "C3604",
            "黒皮": "SS400",
            "ミガキ": "S45C",
            "ステンレス": "SUS304",
            "アルミ": "A5052",
        }
        for alias, mat_key in alias_map.items():
            if alias in norm:
                return mat_key, norm
        return "", norm

    @classmethod
    def detect_treatment(cls, text: str) -> str:
        """文字列から表面処理・熱処理を検出する"""
        norm = cls.normalize_text(text)
        treat_keywords = [
            ("四三酸化鉄皮膜", "四三酸化鉄皮膜(黒染め)"),
            ("黒染め", "四三酸化鉄皮膜(黒染め)"),
            ("黒染", "四三酸化鉄皮膜(黒染め)"),
            ("無電解ニッケル", "無電解ニッケルめっき"),
            ("カニゼン", "無電解ニッケルめっき"),
            ("硬質クロム", "硬質クロムめっき"),
            ("三価クロメート", "三価クロメートめっき"),
            ("ユニクロ", "ユニクロめっき"),
            ("硬質アルマイト", "硬質アルマイト"),
            ("黒アルマイト", "黒アルマイト"),
            ("白アルマイト", "白アルマイト"),
            ("アルマイト", "白アルマイト"),
            ("高周波焼入", "高周波焼入れ"),
            ("浸炭焼入", "浸炭焼入れ"),
            ("焼入研磨", "焼入れ研磨"),
            ("焼入れ", "焼入れ研磨"),
            ("窒化", "窒化処理"),
            ("メラミン塗装", "メラミン焼付塗装"),
            ("粉体塗装", "粉体塗装"),
            ("塗装", "ウレタン塗装"),
            ("酸洗", "酸洗・不動態化処理"),
            ("電解研磨", "電解研磨"),
            ("バフ", "バフ研磨"),
        ]
        for kw, canonical in treat_keywords:
            if kw in norm:
                return canonical
        return ""

    @classmethod
    def detect_category(cls, part_name: str, spec: str, material: str, manufacturer: str) -> str:
        """部品名や仕様・メーカーから調達区分（加工品 / 標準部品 / 購入品 / 板金品 / 製缶・溶接品）を自動推定"""
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
        if any(w in combined for w in ["溶接", "フレーム", "架台", "ベースフレーム", "製缶"]):
            return "製缶・溶接品"
        if any(w in combined for w in ["カバー", "ブラケット", "ステー", "パネル", "シャーシ"]) and material in ("SPCC", "SECC", "SUS304", "A5052"):
            if re.search(r"\bt\s*[0-5](\.\d+)?\b", combined, re.IGNORECASE):
                return "板金品"
        return "加工品"

    @classmethod
    def estimate_weight_from_spec(cls, spec: str, material: str) -> float:
        """
        寸法文字列（例: t10x100x200, 100x50x12, φ30x150, φ50xt3x200 等）と材質比重から単重(kg)を自動推算する
        1 cm3 = 1000 mm3, 単重(kg) = 体積(mm3) * 比重(g/cm3) / 1,000,000
        """
        if not spec:
            return 0.0
        norm = cls.normalize_text(spec)
        density = 7.85  # デフォルト鉄比重
        if material in MATERIAL_MASTER:
            density = float(MATERIAL_MASTER[material]["density"])
        else:
            detected_mat, _ = cls.detect_material(material)
            if detected_mat in MATERIAL_MASTER:
                density = float(MATERIAL_MASTER[detected_mat]["density"])

        # パターン1: 丸パイプ φD x tT x L (例: φ50xt3x200, φ60.5x3.2x500)
        m_pipe = re.search(
            r"[φD]\s*(\d+(?:\.\d+)?)\s*[xX]\s*t?\s*(\d+(?:\.\d+)?)\s*[xX]\s*[L]?\s*(\d+(?:\.\d+)?)",
            norm, re.IGNORECASE
        )
        if m_pipe:
            d_out = float(m_pipe.group(1))
            t_wall = float(m_pipe.group(2))
            length = float(m_pipe.group(3))
            if d_out > t_wall * 2.0 and t_wall < 30.0:
                r_out = d_out * 0.5
                r_in = r_out - t_wall
                vol_mm3 = math.pi * (r_out * r_out - r_in * r_in) * length
                return round(vol_mm3 * density / 1_000_000.0, 3)

        # パターン2: 直方体・プレート (例: t12x100x250, 100x200xt10, 50x80x150)
        m_plate = re.search(
            r"[tT]?\s*(\d+(?:\.\d+)?)\s*[xX]\s*[tT]?\s*(\d+(?:\.\d+)?)\s*[xX]\s*[tT]?\s*[L]?\s*(\d+(?:\.\d+)?)",
            norm
        )
        if m_plate:
            a = float(m_plate.group(1))
            b = float(m_plate.group(2))
            c = float(m_plate.group(3))
            vol_mm3 = a * b * c
            return round(vol_mm3 * density / 1_000_000.0, 3)

        # パターン3: 丸棒・円柱 φD x L (例: φ30x150, φ25xL80, D20x100)
        m_cyl = re.search(
            r"[φD]\s*(\d+(?:\.\d+)?)\s*[xX\-]\s*[L]?\s*(\d+(?:\.\d+)?)",
            norm, re.IGNORECASE
        )
        if m_cyl:
            d_val = float(m_cyl.group(1))
            length = float(m_cyl.group(2))
            radius = d_val * 0.5
            vol_mm3 = math.pi * (radius * radius) * length
            return round(vol_mm3 * density / 1_000_000.0, 3)

        # パターン4: 六角ボルト M径x長さ の概算重量
        m_bolt = re.search(r"\bM(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)", norm, re.IGNORECASE)
        if m_bolt:
            d_val = float(m_bolt.group(1))
            length = float(m_bolt.group(2))
            radius = d_val * 0.5
            # 軸部＋頭部（軸径の約1.2倍径×0.7倍高）の概算
            vol_mm3 = math.pi * (radius * radius) * (length + d_val * 1.1)
            return round(vol_mm3 * density / 1_000_000.0, 4)

        return 0.0

    @classmethod
    def classify_table_headers(cls, header_cells: List[str]) -> Dict[int, str]:
        """テーブルヘッダー行の各列がどのフィールドに対応するかを判定する"""
        col_map: Dict[int, str] = {}
        header_rules = [
            ("item_no", ["NO", "NO.", "項番", "番号", "符号", "品番NO", "ITEM", "POS", "SYM"]),
            ("part_number", ["図番", "部品番号", "品番", "DRAWING", "DWG", "PART NO", "CODE", "部品コード"]),
            ("part_name", ["品名", "名称", "部品名", "部品名称", "NAME", "DESCRIPTION", "TITLE"]),
            ("specification", ["型式", "寸法", "仕様", "サイズ", "規格", "材料寸法", "SIZE", "SPEC", "MODEL", "TYPE"]),
            ("material", ["材質", "材料", "MATERIAL", "MAT"]),
            ("treatment", ["表面処理", "熱処理", "処理", "塗装", "FINISH", "TREATMENT", "COATING"]),
            ("quantity", ["数量", "個数", "員数", "数", "QTY", "Q'TY", "QUANTITY"]),
            ("unit", ["単位", "UNIT"]),
            ("unit_weight", ["単重", "重量", "質量", "WEIGHT", "MASS", "WT"]),
            ("manufacturer", ["メーカー", "メーカ", "製造元", "手配先", "仕入先", "MAKER", "VENDOR", "BRAND"]),
            ("unit_price", ["単価", "価格", "PRICE", "COST"]),
            ("remarks", ["備考", "記事", "摘要", "注記", "REMARK", "REMARKS", "NOTE", "NOTES"]),
        ]

        for idx, raw_cell in enumerate(header_cells):
            cell = cls.normalize_text(raw_cell).upper()
            if not cell:
                continue
            for field_name, keywords in header_rules:
                if field_name in col_map.values():
                    continue
                for kw in keywords:
                    if kw == cell or (len(kw) >= 2 and kw in cell):
                        col_map[idx] = field_name
                        break
                if idx in col_map:
                    break
        return col_map

    @classmethod
    def parse_unstructured_row(
        cls, cells: List[str], default_idx: int, source_file: str, page_no: int, bbox: Tuple[float, float, float, float]
    ) -> Optional[PartItem]:
        """ヘッダーが不明な行や範囲選択されたテキスト行からスマートにPartItemを構築する"""
        cleaned = [cls.normalize_text(c) for c in cells if cls.normalize_text(c)]
        if not cleaned:
            return None

        joined = " ".join(cleaned)
        # ヘッダー行そのものや図面枠の定型句は除外
        ignore_words = ["部品番号", "部品名称", "第三角法", "尺度", "SCALE", "PROJECTION", "設計", "検図", "承認", "作成日"]
        hit_header = sum(1 for w in ["品名", "材質", "数量", "備考", "図番", "名称"] if w in joined)
        if hit_header >= 2:
            return None
        if any(w == joined for w in ignore_words):
            return None

        item = PartItem(
            item_no=str(default_idx),
            source_file=source_file,
            source_page=page_no,
            bbox=bbox,
        )

        remaining: List[str] = []
        for idx, token in enumerate(cleaned):
            # 1. 先頭セルが純粋な小整数なら項番
            if idx == 0 and re.fullmatch(r"\d{1,3}", token):
                item.item_no = token
                continue

            # 2. 材質判定
            mat, _ = cls.detect_material(token)
            if mat and not item.material and len(token) <= 18:
                item.material = mat
                treat = cls.detect_treatment(token)
                if treat and not item.treatment:
                    item.treatment = treat
                continue

            # 3. 表面処理判定
            treat = cls.detect_treatment(token)
            if treat and not item.treatment and len(token) <= 20:
                item.treatment = treat
                continue

            # 4. 数量判定 (例: "2", "4個", "1式", "x4", "QTY:2")
            m_qty_unit = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(個|本|枚|式|組|セット|m|kg)", token)
            if m_qty_unit:
                item.quantity = float(m_qty_unit.group(1))
                item.unit = m_qty_unit.group(2)
                continue

            if idx >= 1 and re.fullmatch(r"\d{1,3}", token):
                val = int(token)
                if 1 <= val <= 500 and item.quantity == 1.0:
                    item.quantity = float(val)
                    continue

            # 5. 重量表記 (例: "1.25kg", "0.8 kg")
            m_wt = re.fullmatch(r"(\d+(?:\.\d+)?)\s*kg", token, re.IGNORECASE)
            if m_wt:
                item.unit_weight = float(m_wt.group(1))
                continue

            # 6. 寸法・仕様表記 (例: t10x100x200, φ30x150, M8x25)
            if re.search(r"([tT]\d+|\d+\s*[xX×]\s*\d+|[φD]\d+|\bM\d+\s*[xX×]\s*\d+)", token):
                if not item.specification:
                    item.specification = token
                    continue

            # 7. 図番・型番パターン (英字+数字+ハイフン、かつ日本語を含まない)
            if not re.search(r"[ぁ-んァ-ン一-龥]", token) and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-_/\.\(\)]{2,24}", token):
                if not item.part_number:
                    item.part_number = token
                    continue
                elif not item.specification:
                    item.specification = token
                    continue

            remaining.append(token)

        # 残りのトークンを品名・備考・メーカーへ割り当て
        for text_piece in remaining:
            if any(mk in text_piece for mk in ["ミスミ", "MISUMI", "THK", "SMC", "CKD", "NSK", "NTN", "IKO", "オリエンタル", "キーエンス", "オムロン", "鍋屋", "NBK", "イマオ", "椿本"]):
                if not item.manufacturer:
                    item.manufacturer = text_piece
                    continue
            if not item.part_name:
                item.part_name = text_piece
            elif not item.specification:
                item.specification = text_piece
            elif not item.remarks:
                item.remarks = text_piece
            else:
                item.remarks = f"{item.remarks} {text_piece}"

        # 品名が空で図番や仕様だけある場合の補完
        if not item.part_name and item.part_number:
            if re.match(r"^M\d+", item.part_number, re.IGNORECASE):
                item.specification = item.part_number
                item.part_name = "六角穴付ボルト"
                item.part_number = ""
            else:
                item.part_name = item.part_number

        if not item.part_name and not item.part_number and not item.specification:
            return None

        # 材質に応じたデフォルト表面処理と重量自動推算
        if item.material and not item.treatment:
            if item.material in MATERIAL_MASTER:
                item.treatment = MATERIAL_MASTER[item.material]["default_treat"]

        if item.unit_weight == 0.0 and item.specification:
            item.unit_weight = cls.estimate_weight_from_spec(item.specification, item.material)

        item.category = cls.detect_category(item.part_name, item.specification, item.material, item.manufacturer)
        return item


# ==============================================================================
# 2D図面解析エンジン層 (PDF & CAD DXF)
# ==============================================================================

class DrawingParserEngine:
    """PDFおよびDXF図面からページプレビュー・テキスト座標・部品リストを抽出するコアエンジン"""

    @classmethod
    def parse_file(cls, file_path: str) -> Tuple[DrawingDocument, List[PartItem]]:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".pdf":
            return cls.parse_pdf(file_path)
        elif ext == ".dxf":
            return cls.parse_dxf(file_path)
        else:
            raise ValueError(f"未対応のファイル形式です: {ext} (対応形式: .pdf, .dxf)")

    # --------------------------------------------------------------------------
    # PDF 図面解析
    # --------------------------------------------------------------------------
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
                rect = page.rect
                width, height = float(rect.width), float(rect.height)

                # プレビュー画像のレンダリング (長辺が約1600pxになるようにスケール調整)
                max_dim = max(width, height, 1.0)
                zoom = min(max(1600.0 / max_dim, 1.2), 3.0)
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # テキストワード群の座標取得
                tokens: List[TextToken] = []
                words = page.get_text("words")
                for w in words:
                    # w = (x0, y0, x1, y1, "word", block_no, line_no, word_no)
                    txt = str(w[4]).strip()
                    if txt:
                        tokens.append(
                            TextToken(
                                text=txt,
                                x0=float(w[0]),
                                y0=float(w[1]),
                                x1=float(w[2]),
                                y1=float(w[3]),
                                page_no=page_no,
                            )
                        )

                page_data = DrawingPageData(
                    page_no=page_no,
                    width=width,
                    height=height,
                    preview_image=img,
                    tokens=tokens,
                )

                # 1. PyMuPDFのテーブル構造検出を試行
                page_parts: List[PartItem] = []
                try:
                    tabs = page.find_tables()
                    for tab in tabs:
                        extracted = tab.extract()
                        if extracted and len(extracted) >= 2:
                            page_data.raw_tables.append(extracted)
                            tab_bbox = tuple(float(v) for v in tab.bbox)
                            parts_from_tab = cls._extract_parts_from_2d_cells(
                                extracted, file_name, page_no, tab_bbox
                            )
                            page_parts.extend(parts_from_tab)
                except Exception:
                    pass

                # 2. テーブル検出で部品が取れなかった場合（または追加で）空間座標クラスタリング解析を実行
                if not page_parts and tokens:
                    clustered_parts = cls.extract_parts_from_tokens(
                        tokens, file_name, page_no, (0.0, 0.0, width, height)
                    )
                    page_parts.extend(clustered_parts)

                # 表題欄メタデータの推定
                title_meta = cls._detect_title_block_meta(tokens, width, height)
                page_data.title_meta = title_meta
                if title_meta.get("drawing_title") and not doc_model.drawing_title:
                    doc_model.drawing_title = title_meta["drawing_title"]
                if title_meta.get("drawing_number") and not doc_model.drawing_number:
                    doc_model.drawing_number = title_meta["drawing_number"]

                # 単品図（部品表がなく表題欄のみある図面）の場合のフォールバック生成
                if not page_parts and title_meta.get("drawing_title"):
                    single_item = PartItem(
                        item_no="1",
                        part_number=title_meta.get("drawing_number", os.path.splitext(file_name)[0]),
                        part_name=title_meta.get("drawing_title", ""),
                        material=title_meta.get("material", "SS400"),
                        treatment=title_meta.get("treatment", ""),
                        quantity=1.0,
                        unit="個",
                        source_file=file_name,
                        source_page=page_no,
                        bbox=(width * 0.55, height * 0.75, width * 0.98, height * 0.98),
                    )
                    page_parts.append(single_item)

                doc_model.pages.append(page_data)
                all_parts.extend(page_parts)
        finally:
            pdf_doc.close()

        doc_model.extracted_count = len(all_parts)
        return doc_model, all_parts

    # --------------------------------------------------------------------------
    # DXF (2D CAD) 図面解析＆ベクタープレビュー生成
    # --------------------------------------------------------------------------
    @classmethod
    def parse_dxf(cls, file_path: str) -> Tuple[DrawingDocument, List[PartItem]]:
        file_name = os.path.basename(file_path)
        doc_model = DrawingDocument(file_path=file_path, file_name=file_name, file_type="DXF")
        all_parts: List[PartItem] = []

        # DXF読み込み（エンコーディング配慮）
        try:
            dxf_doc = ezdxf.readfile(file_path)
        except UnicodeDecodeError:
            dxf_doc = ezdxf.readfile(file_path, encoding="cp932")

        msp = dxf_doc.modelspace()

        # 1. 幾何エンティティとテキストを走査してバウンディングボックスを算出
        raw_texts: List[Tuple[str, float, float, float]] = []  # (text, x, y, height)
        lines_data: List[Tuple[float, float, float, float, str]] = []  # (x0, y0, x1, y1, color_hint)
        circles_data: List[Tuple[float, float, float]] = []  # (cx, cy, r)
        attrib_parts: List[PartItem] = []

        min_x, min_y = float("inf"), float("inf")
        max_x, max_y = float("-inf"), float("-inf")

        def update_bounds(x: float, y: float) -> None:
            nonlocal min_x, min_y, max_x, max_y
            if math.isfinite(x) and math.isfinite(y):
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)

        def collect_entity(entity: Any, offset_x: float = 0.0, offset_y: float = 0.0) -> None:
            etype = entity.dxftype()
            try:
                if etype == "LINE":
                    s = entity.dxf.start
                    e = entity.dxf.end
                    x0, y0 = float(s.x) + offset_x, float(s.y) + offset_y
                    x1, y1 = float(e.x) + offset_x, float(e.y) + offset_y
                    lines_data.append((x0, y0, x1, y1, "line"))
                    update_bounds(x0, y0)
                    update_bounds(x1, y1)

                elif etype == "LWPOLYLINE":
                    pts = list(entity.get_points("xy"))
                    if len(pts) >= 2:
                        for i in range(len(pts) - 1):
                            x0, y0 = float(pts[i][0]) + offset_x, float(pts[i][1]) + offset_y
                            x1, y1 = float(pts[i + 1][0]) + offset_x, float(pts[i + 1][1]) + offset_y
                            lines_data.append((x0, y0, x1, y1, "poly"))
                            update_bounds(x0, y0)
                            update_bounds(x1, y1)
                        if entity.closed:
                            x0, y0 = float(pts[-1][0]) + offset_x, float(pts[-1][1]) + offset_y
                            x1, y1 = float(pts[0][0]) + offset_x, float(pts[0][1]) + offset_y
                            lines_data.append((x0, y0, x1, y1, "poly"))

                elif etype == "CIRCLE":
                    c = entity.dxf.center
                    r = float(entity.dxf.radius)
                    cx, cy = float(c.x) + offset_x, float(c.y) + offset_y
                    circles_data.append((cx, cy, r))
                    update_bounds(cx - r, cy - r)
                    update_bounds(cx + r, cy + r)

                elif etype == "ARC":
                    c = entity.dxf.center
                    r = float(entity.dxf.radius)
                    cx, cy = float(c.x) + offset_x, float(c.y) + offset_y
                    circles_data.append((cx, cy, r))
                    update_bounds(cx - r, cy - r)
                    update_bounds(cx + r, cy + r)

                elif etype == "TEXT":
                    txt = str(entity.dxf.text).strip()
                    if txt:
                        ins = entity.dxf.insert
                        tx, ty = float(ins.x) + offset_x, float(ins.y) + offset_y
                        th = float(getattr(entity.dxf, "height", 3.5))
                        raw_texts.append((txt, tx, ty, th))
                        update_bounds(tx, ty)
                        update_bounds(tx + len(txt) * th * 0.8, ty + th)

                elif etype == "MTEXT":
                    txt = entity.plain_text().strip()
                    if txt:
                        ins = entity.dxf.insert
                        tx, ty = float(ins.x) + offset_x, float(ins.y) + offset_y
                        th = float(getattr(entity.dxf, "char_height", 3.5))
                        for line_idx, line_str in enumerate(txt.splitlines()):
                            line_str = line_str.strip()
                            if line_str:
                                ly = ty - line_idx * (th * 1.4)
                                raw_texts.append((line_str, tx, ly, th))
                                update_bounds(tx, ly)

                elif etype == "INSERT":
                    ins = entity.dxf.insert
                    ix, iy = float(ins.x) + offset_x, float(ins.y) + offset_y
                    update_bounds(ix, iy)
                    # ブロック属性(ATTRIB)の抽出
                    if hasattr(entity, "attribs") and len(entity.attribs) > 0:
                        attr_dict: Dict[str, str] = {}
                        for att in entity.attribs:
                            tag = str(att.dxf.tag).upper().strip()
                            val = str(att.dxf.text).strip()
                            if val:
                                attr_dict[tag] = val
                                a_ins = att.dxf.insert
                                ax, ay = float(a_ins.x), float(a_ins.y)
                                raw_texts.append((val, ax, ay, 3.5))
                        part_from_attr = cls._part_from_dxf_attribs(attr_dict, file_name, len(attrib_parts) + 1)
                        if part_from_attr:
                            attrib_parts.append(part_from_attr)

                    # ブロック内図形の展開（1階層）
                    bname = entity.dxf.name
                    if bname in dxf_doc.blocks:
                        blk = dxf_doc.blocks[bname]
                        for sub_e in blk:
                            if sub_e.dxftype() in ("LINE", "LWPOLYLINE", "CIRCLE", "TEXT", "MTEXT"):
                                collect_entity(sub_e, ix, iy)
            except Exception:
                pass

        for ent in msp:
            collect_entity(ent)

        if not math.isfinite(min_x) or not math.isfinite(max_x) or max_x <= min_x:
            min_x, min_y, max_x, max_y = 0.0, 0.0, 420.0, 297.0

        margin_x = max((max_x - min_x) * 0.04, 10.0)
        margin_y = max((max_y - min_y) * 0.04, 10.0)
        min_x -= margin_x
        min_y -= margin_y
        max_x += margin_x
        max_y += margin_y

        world_w = max(max_x - min_x, 1.0)
        world_h = max(max_y - min_y, 1.0)

        # DXFはY軸が上向きなので、画面座標系(左上原点、Y軸下向き)に変換
        tokens: List[TextToken] = []
        for txt, tx, ty, th in raw_texts:
            sx0 = tx - min_x
            sy1 = max_y - ty
            sy0 = max(0.0, sy1 - max(th * 1.2, 4.0))
            sx1 = sx0 + max(len(txt) * th * 0.85, 8.0)
            tokens.append(
                TextToken(
                    text=txt,
                    x0=sx0,
                    y0=sy0,
                    x1=sx1,
                    y1=sy1,
                    page_no=1,
                )
            )

        # DXFの2Dプレビュー画像を生成
        preview_img = cls._render_dxf_preview(
            world_w, world_h, min_x, max_y, lines_data, circles_data, tokens
        )

        page_data = DrawingPageData(
            page_no=1,
            width=world_w,
            height=world_h,
            preview_image=preview_img,
            tokens=tokens,
        )

        # 部品抽出: 1. ブロック属性 -> 2. 空間座標クラスタリング
        if attrib_parts:
            all_parts.extend(attrib_parts)

        clustered_parts = cls.extract_parts_from_tokens(
            tokens, file_name, 1, (0.0, 0.0, world_w, world_h)
        )
        # 重複を避けてマージ
        existing_keys = {(p.part_number, p.part_name) for p in all_parts}
        for cp in clustered_parts:
            if (cp.part_number, cp.part_name) not in existing_keys:
                all_parts.append(cp)

        # 連番振り直し
        for idx, p in enumerate(all_parts, start=1):
            if not p.item_no or p.item_no == "0":
                p.item_no = str(idx)

        title_meta = cls._detect_title_block_meta(tokens, world_w, world_h)
        page_data.title_meta = title_meta
        doc_model.drawing_title = title_meta.get("drawing_title", "")
        doc_model.drawing_number = title_meta.get("drawing_number", "")
        doc_model.pages.append(page_data)
        doc_model.extracted_count = len(all_parts)

        return doc_model, all_parts

    @classmethod
    def _render_dxf_preview(
        cls,
        world_w: float,
        world_h: float,
        min_x: float,
        max_y: float,
        lines_data: List[Tuple[float, float, float, float, str]],
        circles_data: List[Tuple[float, float, float]],
        tokens: List[TextToken],
    ) -> Image.Image:
        """DXFの線分・円・テキストをPillow画像に美しくベクターレンダリングする"""
        target_w = 1600
        scale = target_w / max(world_w, 1.0)
        target_h = int(max(min(world_h * scale, 2200), 600))
        scale = min(target_w / max(world_w, 1.0), target_h / max(world_h, 1.0))

        img = Image.new("RGB", (target_w, target_h), (255, 255, 255))
        draw = ImageDraw.Draw(img)

        # フォント準備 (Windows 11 標準日本語フォント)
        font_sm = None
        for font_path in [
            "C:/Windows/Fonts/meiryo.ttc",
            "C:/Windows/Fonts/YuGothM.ttc",
            "C:/Windows/Fonts/msgothic.ttc",
        ]:
            if os.path.exists(font_path):
                try:
                    font_sm = ImageFont.truetype(font_path, 12)
                    break
                except Exception:
                    pass

        # 線分描画
        for x0, y0, x1, y1, kind in lines_data:
            px0 = int((x0 - min_x) * scale)
            py0 = int((max_y - y0) * scale)
            px1 = int((x1 - min_x) * scale)
            py1 = int((max_y - y1) * scale)
            color = (30, 58, 138) if kind == "poly" else (15, 23, 42)
            draw.line([(px0, py0), (px1, py1)], fill=color, width=2)

        # 円描画
        for cx, cy, r in circles_data:
            px0 = int((cx - r - min_x) * scale)
            py0 = int((max_y - (cy + r)) * scale)
            px1 = int((cx + r - min_x) * scale)
            py1 = int((max_y - (cy - r)) * scale)
            if px1 > px0 and py1 > py0:
                draw.ellipse([px0, py0, px1, py1], outline=(37, 99, 235), width=2)

        # テキスト描画
        for tk_item in tokens:
            px = int(tk_item.x0 * scale)
            py = int(tk_item.y0 * scale)
            draw.text((px, py), tk_item.text, fill=(15, 23, 42), font=font_sm)

        return img

    @classmethod
    def _part_from_dxf_attribs(
        cls, attr_dict: Dict[str, str], source_file: str, idx: int
    ) -> Optional[PartItem]:
        """DXFブロック属性辞書からPartItemを生成する"""
        if not attr_dict:
            return None
        item = PartItem(item_no=str(idx), source_file=source_file, source_page=1)
        for tag, val in attr_dict.items():
            t = tag.upper()
            if any(k in t for k in ["NO", "POS", "ITEM", "番号", "符号"]):
                if re.fullmatch(r"\d+", val):
                    item.item_no = val
                elif not item.part_number:
                    item.part_number = val
            elif any(k in t for k in ["DWG", "PART_NO", "CODE", "図番", "品番"]):
                item.part_number = val
            elif any(k in t for k in ["NAME", "DESC", "TITLE", "品名", "名称"]):
                item.part_name = val
            elif any(k in t for k in ["SPEC", "SIZE", "DIM", "MODEL", "型式", "寸法", "仕様"]):
                item.specification = val
            elif any(k in t for k in ["MAT", "材質", "材料"]):
                item.material = val
            elif any(k in t for k in ["FINISH", "TREAT", "表面処理", "処理"]):
                item.treatment = val
            elif any(k in t for k in ["QTY", "COUNT", "数量", "個数"]):
                m = re.search(r"\d+(?:\.\d+)?", val)
                if m:
                    item.quantity = float(m.group(0))
            elif any(k in t for k in ["WEIGHT", "MASS", "単重", "重量"]):
                m = re.search(r"\d+(?:\.\d+)?", val)
                if m:
                    item.unit_weight = float(m.group(0))
            elif any(k in t for k in ["MAKER", "VENDOR", "メーカー", "メーカ"]):
                item.manufacturer = val
            elif any(k in t for k in ["REMARK", "NOTE", "備考"]):
                item.remarks = val

        if not item.part_name and not item.part_number:
            return None
        if item.unit_weight == 0.0 and item.specification:
            item.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(item.specification, item.material)
        item.category = SmartPartAnalyzer.detect_category(
            item.part_name, item.specification, item.material, item.manufacturer
        )
        return item

    # --------------------------------------------------------------------------
    # 2Dテーブルおよび空間座標クラスタリングによる部品表抽出
    # --------------------------------------------------------------------------
    @classmethod
    def _extract_parts_from_2d_cells(
        cls,
        rows: List[List[Any]],
        source_file: str,
        page_no: int,
        table_bbox: Tuple[float, float, float, float],
    ) -> List[PartItem]:
        """2次元セル配列からヘッダー位置を特定してPartItemリストを抽出する"""
        cleaned_rows: List[List[str]] = []
        for r in rows:
            cleaned_r = [SmartPartAnalyzer.normalize_text(str(c) if c is not None else "") for c in r]
            if any(cleaned_r):
                cleaned_rows.append(cleaned_r)

        if len(cleaned_rows) < 2:
            return []

        # ヘッダー行を探す（上から順、または下から上に積まれる部品表のため最下行もチェック）
        header_idx = -1
        col_map: Dict[int, str] = {}
        bottom_up = False

        for idx, r in enumerate(cleaned_rows):
            cmap = SmartPartAnalyzer.classify_table_headers(r)
            if len(cmap) >= 2 and ("part_name" in cmap.values() or "part_number" in cmap.values() or "material" in cmap.values()):
                header_idx = idx
                col_map = cmap
                if idx == len(cleaned_rows) - 1:
                    bottom_up = True
                break

        results: List[PartItem] = []
        total_rows = max(len(cleaned_rows), 1)
        bx0, by0, bx1, by1 = table_bbox
        row_h = (by1 - by0) / total_rows

        if header_idx != -1:
            data_indices = (
                range(header_idx - 1, -1, -1) if bottom_up else range(header_idx + 1, len(cleaned_rows))
            )
            for seq_idx, r_idx in enumerate(data_indices, start=1):
                r = cleaned_rows[r_idx]
                row_bbox = (bx0, by0 + r_idx * row_h, bx1, by0 + (r_idx + 1) * row_h)
                item = cls._map_row_with_col_map(r, col_map, seq_idx, source_file, page_no, row_bbox)
                if item:
                    results.append(item)
        else:
            for seq_idx, r in enumerate(cleaned_rows, start=1):
                row_bbox = (bx0, by0 + (seq_idx - 1) * row_h, bx1, by0 + seq_idx * row_h)
                item = SmartPartAnalyzer.parse_unstructured_row(r, seq_idx, source_file, page_no, row_bbox)
                if item:
                    results.append(item)

        return results

    @classmethod
    def _map_row_with_col_map(
        cls,
        row_cells: List[str],
        col_map: Dict[int, str],
        default_no: int,
        source_file: str,
        page_no: int,
        bbox: Tuple[float, float, float, float],
    ) -> Optional[PartItem]:
        """ヘッダーマッピングに従って1行のセル群をPartItemに変換する"""
        if not any(row_cells):
            return None

        # 別のヘッダーや集計行はスキップ
        joined = " ".join(row_cells)
        if any(w in joined for w in ["部品表", "PARTS LIST", "BILL OF MATERIALS", "第三角法", "PROJECTION"]):
            return None

        item = PartItem(
            item_no=str(default_no),
            source_file=source_file,
            source_page=page_no,
            bbox=bbox,
        )

        unmapped_texts: List[str] = []
        for c_idx, val in enumerate(row_cells):
            val = val.strip()
            if not val:
                continue
            field_name = col_map.get(c_idx)
            if field_name == "item_no":
                item.item_no = val
            elif field_name == "part_number":
                item.part_number = val
            elif field_name == "part_name":
                item.part_name = val
            elif field_name == "specification":
                item.specification = val
            elif field_name == "material":
                mat, _ = SmartPartAnalyzer.detect_material(val)
                item.material = mat if mat else val
                treat = SmartPartAnalyzer.detect_treatment(val)
                if treat and not item.treatment:
                    item.treatment = treat
            elif field_name == "treatment":
                item.treatment = val
            elif field_name == "quantity":
                m = re.search(r"\d+(?:\.\d+)?", val)
                if m:
                    item.quantity = float(m.group(0))
                u = re.search(r"(個|本|枚|式|組|セット|m|kg)", val)
                if u:
                    item.unit = u.group(1)
            elif field_name == "unit":
                item.unit = val
            elif field_name == "unit_weight":
                m = re.search(r"\d+(?:\.\d+)?", val)
                if m:
                    item.unit_weight = float(m.group(0))
            elif field_name == "manufacturer":
                item.manufacturer = val
            elif field_name == "unit_price":
                m = re.search(r"[\d,]+(?:\.\d+)?", val)
                if m:
                    item.unit_price = float(m.group(0).replace(",", ""))
            elif field_name == "remarks":
                item.remarks = val
            else:
                unmapped_texts.append(val)

        # 未マッピング列に重要情報があれば補完
        for extra in unmapped_texts:
            if not item.material:
                mat, _ = SmartPartAnalyzer.detect_material(extra)
                if mat:
                    item.material = mat
                    continue
            if not item.treatment:
                treat = SmartPartAnalyzer.detect_treatment(extra)
                if treat:
                    item.treatment = treat
                    continue
            if not item.specification and re.search(r"([tT]\d+|\d+[xX×]\d+|[φD]\d+|\bM\d+)", extra):
                item.specification = extra
                continue
            if not item.remarks:
                item.remarks = extra

        if not item.part_name and not item.part_number:
            return None

        if item.material and not item.treatment and item.material in MATERIAL_MASTER:
            item.treatment = MATERIAL_MASTER[item.material]["default_treat"]

        if item.unit_weight == 0.0 and item.specification:
            item.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(item.specification, item.material)

        item.category = SmartPartAnalyzer.detect_category(
            item.part_name, item.specification, item.material, item.manufacturer
        )
        return item

    @classmethod
    def extract_parts_from_tokens(
        cls,
        tokens: List[TextToken],
        source_file: str,
        page_no: int,
        area_rect: Tuple[float, float, float, float],
        force_all_rows: bool = False,
    ) -> List[PartItem]:
        """
        指定矩形領域内のTextToken群をY座標で行クラスタリングし、X座標で列整列して部品リストを抽出する。
        force_all_rows=True のとき（マウス範囲選択時）はヘッダーがなくても範囲内の全有効行を部品変換する。
        """
        ax0, ay0, ax1, ay1 = area_rect
        filtered = [
            t for t in tokens
            if t.cx >= ax0 and t.cx <= ax1 and t.cy >= ay0 and t.cy <= ay1
        ]
        if not filtered:
            return []

        # Y座標でソートし、近接するトークンを同一行（RowCluster）にまとめる
        sorted_by_y = sorted(filtered, key=lambda t: t.cy)
        avg_h = sum(max(t.y1 - t.y0, 4.0) for t in sorted_by_y) / len(sorted_by_y)
        y_tol = max(avg_h * 0.65, 3.5)

        row_clusters: List[List[TextToken]] = []
        for tok in sorted_by_y:
            if not row_clusters:
                row_clusters.append([tok])
            else:
                prev_cy = sum(t.cy for t in row_clusters[-1]) / len(row_clusters[-1])
                if abs(tok.cy - prev_cy) <= y_tol:
                    row_clusters[-1].append(tok)
                else:
                    row_clusters.append([tok])

        # 各行内でX座標順にソートし、近すぎる単語は結合、離れている単語は別セルとする
        structured_rows: List[Tuple[List[str], List[float], Tuple[float, float, float, float]]] = []
        for cluster in row_clusters:
            cluster_sorted = sorted(cluster, key=lambda t: t.x0)
            cells: List[str] = []
            cell_xs: List[float] = []
            cur_text = cluster_sorted[0].text
            cur_x0 = cluster_sorted[0].x0
            prev_x1 = cluster_sorted[0].x1

            for tok in cluster_sorted[1:]:
                gap = tok.x0 - prev_x1
                if gap <= max(avg_h * 0.8, 5.0):
                    cur_text = f"{cur_text} {tok.text}"
                    prev_x1 = max(prev_x1, tok.x1)
                else:
                    cells.append(cur_text)
                    cell_xs.append(cur_x0)
                    cur_text = tok.text
                    cur_x0 = tok.x0
                    prev_x1 = tok.x1
            cells.append(cur_text)
            cell_xs.append(cur_x0)

            rx0 = min(t.x0 for t in cluster_sorted)
            ry0 = min(t.y0 for t in cluster_sorted)
            rx1 = max(t.x1 for t in cluster_sorted)
            ry1 = max(t.y1 for t in cluster_sorted)
            structured_rows.append((cells, cell_xs, (rx0, ry0, rx1, ry1)))

        # ヘッダー行を探す
        header_row_idx = -1
        header_col_map: Dict[int, str] = {}
        header_xs: List[float] = []
        bottom_up = False

        for r_idx, (cells, c_xs, _) in enumerate(structured_rows):
            cmap = SmartPartAnalyzer.classify_table_headers(cells)
            if len(cmap) >= 2 and ("part_name" in cmap.values() or "part_number" in cmap.values()):
                header_row_idx = r_idx
                header_col_map = cmap
                header_xs = c_xs
                # 下側にデータ行が少なく上側にデータ行が多い場合は下から上へ積む部品欄と判定
                if r_idx >= len(structured_rows) - 2 and r_idx >= 2:
                    bottom_up = True
                break

        extracted_parts: List[PartItem] = []

        if header_row_idx != -1 and header_xs:
            indices = (
                range(header_row_idx - 1, -1, -1)
                if bottom_up
                else range(header_row_idx + 1, len(structured_rows))
            )
            for seq_no, idx in enumerate(indices, start=1):
                cells, c_xs, r_bbox = structured_rows[idx]
                # ヘッダーとのX座標範囲が大きく外れている行（図面左側の注記など）は除外
                if not force_all_rows:
                    hx_min = min(header_xs) - 60.0
                    if max(c_xs) < hx_min:
                        continue
                # 各セルのX座標を最も近いヘッダー列に割り当てる
                aligned_cells = [""] * len(header_xs)
                for cell_str, cx in zip(cells, c_xs):
                    best_col = min(range(len(header_xs)), key=lambda hi: abs(cx - header_xs[hi]))
                    if aligned_cells[best_col]:
                        aligned_cells[best_col] = f"{aligned_cells[best_col]} {cell_str}"
                    else:
                        aligned_cells[best_col] = cell_str

                item = cls._map_row_with_col_map(
                    aligned_cells, header_col_map, seq_no, source_file, page_no, r_bbox
                )
                if item:
                    extracted_parts.append(item)
        else:
            # ヘッダーが見つからない場合（範囲選択時または材質・品番を含む行の自動検出）
            seq_no = 1
            for cells, _, r_bbox in structured_rows:
                if not force_all_rows and len(cells) < 2:
                    continue
                item = SmartPartAnalyzer.parse_unstructured_row(
                    cells, seq_no, source_file, page_no, r_bbox
                )
                if item:
                    # 全体自動スキャン時は、材質または図番または2列以上の有効情報がある行のみ採用
                    if force_all_rows or item.material or item.part_number or len(cells) >= 3:
                        extracted_parts.append(item)
                        seq_no += 1

        return extracted_parts

    @classmethod
    def _detect_title_block_meta(
        cls, tokens: List[TextToken], width: float, height: float
    ) -> Dict[str, str]:
        """図面右下付近の表題欄から図面名称・図番・材質を推定する"""
        meta: Dict[str, str] = {"drawing_title": "", "drawing_number": "", "material": ""}
        # 右下領域 (X > 45%, Y > 65%) のトークンを調べる
        tb_tokens = [t for t in tokens if t.cx >= width * 0.45 and t.cy >= height * 0.65]
        if not tb_tokens:
            tb_tokens = tokens

        for i, tok in enumerate(tb_tokens):
            txt = SmartPartAnalyzer.normalize_text(tok.text)
            mat, _ = SmartPartAnalyzer.detect_material(txt)
            if mat and not meta["material"]:
                meta["material"] = mat
            if re.fullmatch(r"[A-Z]{2,5}-\d{3,6}(?:-\d{1,3})?", txt) and not meta["drawing_number"]:
                meta["drawing_number"] = txt
            if txt in ("図面名称", "名称", "品名", "TITLE", "NAME") and i + 1 < len(tb_tokens):
                cand = SmartPartAnalyzer.normalize_text(tb_tokens[i + 1].text)
                if len(cand) >= 2 and cand not in ("図番", "尺度", "材質", "数量"):
                    meta["drawing_title"] = cand

        return meta


# ==============================================================================
# サンプル2D図面 (PDF / DXF) 自動生成エンジン
# ==============================================================================

class SampleDrawingGenerator:
    """動作確認用のリアルな機械設計2D図面（組立図PDF＆ブラケット組図DXF）をその場で生成する"""

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
        ["13", "CB-M6-16", "六角穴付ボルト", "M6x16", "SUS304", "生地", "8", "本", "0.006", "標準部品"],
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
        """A3横サイズ相当(1190 x 842 pt)のベクター2D機械組立図PDFを生成する"""
        doc = fitz.open()
        page = doc.new_page(width=1190, height=842)
        shape = page.new_shape()

        # 1. 図面外枠・内枠の描画
        shape.draw_rect(fitz.Rect(20, 20, 1170, 822))
        shape.finish(color=(0.05, 0.1, 0.25), width=2.0)
        shape.draw_rect(fitz.Rect(28, 28, 1162, 814))
        shape.finish(color=(0.05, 0.1, 0.25), width=1.0)

        # 2. 左側に2Dコンベア駆動ユニット組立図（正面図・平面図・バルーン）を描画
        # ベースフレーム
        shape.draw_rect(fitz.Rect(90, 440, 560, 470))
        shape.draw_rect(fitz.Rect(110, 470, 145, 680))
        shape.draw_rect(fitz.Rect(505, 470, 540, 680))
        shape.draw_rect(fitz.Rect(95, 680, 160, 695))
        shape.draw_rect(fitz.Rect(490, 680, 555, 695))
        # サイドプレート＆コンベアフレーム
        shape.draw_rect(fitz.Rect(100, 320, 550, 440))
        shape.finish(color=(0.1, 0.2, 0.45), width=1.6)

        # ドライブシャフト・プーリ（左右の円）
        shape.draw_circle(fitz.Point(155, 380), 42)
        shape.draw_circle(fitz.Point(155, 380), 16)
        shape.draw_circle(fitz.Point(495, 380), 38)
        shape.draw_circle(fitz.Point(495, 380), 14)
        shape.finish(color=(0.05, 0.35, 0.75), width=1.8)

        # 平面図（上段）
        shape.draw_rect(fitz.Rect(100, 110, 550, 250))
        shape.draw_rect(fitz.Rect(100, 110, 550, 125))
        shape.draw_rect(fitz.Rect(100, 235, 550, 250))
        shape.draw_rect(fitz.Rect(135, 70, 175, 290))
        shape.draw_rect(fitz.Rect(475, 95, 515, 265))
        # モータユニット
        shape.draw_rect(fitz.Rect(115, 45, 220, 70))
        shape.finish(color=(0.1, 0.2, 0.45), width=1.5)

        # 中心線（一点鎖線風）
        shape.draw_line(fitz.Point(75, 380), fitz.Point(575, 380))
        shape.draw_line(fitz.Point(155, 310), fitz.Point(155, 450))
        shape.draw_line(fitz.Point(495, 310), fitz.Point(495, 450))
        shape.draw_line(fitz.Point(75, 180), fitz.Point(575, 180))
        shape.finish(color=(0.8, 0.2, 0.2), width=0.7, dashes="[8 3 2 3] 0")

        # バルーン（部品番号引き出し線と丸囲み）
        balloons = [
            (260, 455, 220, 530, "1"),
            (155, 364, 80, 295, "2"),
            (495, 366, 575, 295, "3"),
            (320, 320, 350, 275, "4"),
            (320, 110, 360, 65, "6"),
            (155, 110, 85, 65, "9"),
            (190, 55, 270, 55, "11"),
        ]
        for bx0, by0, bx1, by1, b_num in balloons:
            shape.draw_line(fitz.Point(bx0, by0), fitz.Point(bx1, by1))
            shape.draw_circle(fitz.Point(bx1, by1), 13)
        shape.finish(color=(0.15, 0.4, 0.2), width=1.1)

        # 3. 右上に部品表テーブル (X: 600〜1150, Y: 45〜465) を描画
        tx0, ty0, tx1 = 595, 45, 1152
        col_widths = [32, 82, 115, 92, 55, 88, 32, 30, 31]  # 合計557
        row_h = 26
        num_rows = len(cls.SAMPLE_BOM_ROWS_PDF)
        ty1 = ty0 + num_rows * row_h

        # ヘッダー背景塗りつぶし
        shape.draw_rect(fitz.Rect(tx0, ty0, tx1, ty0 + row_h))
        shape.finish(color=(0.1, 0.2, 0.4), fill=(0.88, 0.93, 0.98), width=1.2)

        # 横罫線
        for r_i in range(num_rows + 1):
            y_pos = ty0 + r_i * row_h
            shape.draw_line(fitz.Point(tx0, y_pos), fitz.Point(tx1, y_pos))
        # 縦罫線
        cur_x = tx0
        col_x_positions = [cur_x]
        for cw in col_widths[:-1]:
            cur_x += cw
            col_x_positions.append(cur_x)
            shape.draw_line(fitz.Point(cur_x, ty0), fitz.Point(cur_x, ty1))
        col_x_positions.append(tx1)
        shape.draw_rect(fitz.Rect(tx0, ty0, tx1, ty1))
        shape.finish(color=(0.1, 0.15, 0.25), width=1.0)

        # 4. 右下に表題欄 (Title Block) を描画
        tb_x0, tb_y0, tb_x1, tb_y1 = 720, 700, 1152, 806
        shape.draw_rect(fitz.Rect(tb_x0, tb_y0, tb_x1, tb_y1))
        shape.draw_line(fitz.Point(tb_x0, 735), fitz.Point(tb_x1, 735))
        shape.draw_line(fitz.Point(tb_x0, 770), fitz.Point(tb_x1, 770))
        shape.draw_line(fitz.Point(840, tb_y0), fitz.Point(840, tb_y1))
        shape.finish(color=(0.1, 0.15, 0.25), width=1.2)

        shape.commit()

        # 日本語テキスト書き込み (PyMuPDF内蔵CMap/CJKフォント 'japan' を使用)
        fontname = "japan"
        # バルーン番号
        for _, _, bx1, by1, b_num in balloons:
            page.insert_text(fitz.Point(bx1 - 4, by1 + 4), b_num, fontname=fontname, fontsize=10, color=(0.1, 0.3, 0.1))

        # 図面見出し・寸法テキスト
        page.insert_text(fitz.Point(250, 298), "【平面図 (TOP VIEW)】", fontname=fontname, fontsize=11, color=(0.1, 0.2, 0.4))
        page.insert_text(fitz.Point(250, 725), "【正面図 (FRONT VIEW) - 全長 600mm × 全高 420mm】", fontname=fontname, fontsize=11, color=(0.1, 0.2, 0.4))

        # 部品表セル文字列の書き込み
        for r_idx, row_vals in enumerate(cls.SAMPLE_BOM_ROWS_PDF):
            y_text = ty0 + r_idx * row_h + 17
            # 10列あるうちの先頭9列（または備考を結合）
            display_vals = [
                row_vals[0],
                row_vals[1],
                row_vals[2],
                row_vals[3],
                row_vals[4],
                row_vals[5],
                row_vals[6],
                row_vals[8],
                row_vals[9],
            ]
            for c_idx, cell_str in enumerate(display_vals):
                x_text = col_x_positions[c_idx] + 4
                page.insert_text(
                    fitz.Point(x_text, y_text),
                    cell_str,
                    fontname=fontname,
                    fontsize=8.2 if r_idx > 0 else 8.5,
                    color=(0.05, 0.08, 0.15),
                )

        # 表題欄テキスト
        page.insert_text(fitz.Point(730, 722), "図面名称", fontname=fontname, fontsize=10, color=(0.2, 0.2, 0.3))
        page.insert_text(fitz.Point(855, 722), "ベルトコンベア駆動ユニット組立図", fontname=fontname, fontsize=12, color=(0.05, 0.1, 0.25))
        page.insert_text(fitz.Point(730, 757), "図面番号", fontname=fontname, fontsize=10, color=(0.2, 0.2, 0.3))
        page.insert_text(fitz.Point(855, 757), "CV-1001-ASM", fontname=fontname, fontsize=11, color=(0.05, 0.1, 0.25))
        page.insert_text(fitz.Point(730, 792), "尺度 / 設計", fontname=fontname, fontsize=10, color=(0.2, 0.2, 0.3))
        page.insert_text(fitz.Point(855, 792), "SCALE 1:5   設計: PartLens技術部", fontname=fontname, fontsize=10, color=(0.05, 0.1, 0.25))

        doc.save(output_path)
        doc.close()
        return output_path

    @classmethod
    def create_sample_dxf(cls, output_path: str) -> str:
        """2D CAD (DXF) のサンプル図面（クランプブラケット組図＋部品欄）を生成する"""
        doc = ezdxf.new("R2010")
        msp = doc.modelspace()

        # 外枠 (0,0) - (420, 297) A3相当
        msp.add_lwpolyline([(0, 0), (420, 0), (420, 297), (0, 297), (0, 0)], close=True)
        msp.add_lwpolyline([(5, 5), (415, 5), (415, 292), (5, 292), (5, 5)], close=True)

        # 左側にブラケット形状を描画
        msp.add_lwpolyline([(30, 60), (190, 60), (190, 85), (30, 85), (30, 60)], close=True)
        msp.add_lwpolyline([(60, 85), (95, 85), (95, 210), (60, 210), (60, 85)], close=True)
        msp.add_lwpolyline([(125, 85), (160, 85), (160, 210), (125, 210), (125, 85)], close=True)
        msp.add_circle((77.5, 175), 12)
        msp.add_circle((142.5, 175), 12)
        msp.add_circle((50, 72.5), 6)
        msp.add_circle((170, 72.5), 6)

        # 右上に部品表 (X: 205〜410, Y: 280から下へ)
        headers = ["No", "図番", "品名", "寸法", "材質", "表面処理", "数量", "単重", "備考"]
        col_xs = [208, 220, 255, 295, 330, 355, 385, 396, 405]
        top_y = 278.0
        row_step = 10.0

        # ヘッダー行
        msp.add_line((205, top_y + 4), (412, top_y + 4))
        for c_x, h_txt in zip(col_xs, headers):
            t_ent = msp.add_text(h_txt, dxfattribs={"height": 3.0})
            t_ent.dxf.insert = (c_x, top_y - 3)
        msp.add_line((205, top_y - 5), (412, top_y - 5))

        # データ行
        for r_idx, r_vals in enumerate(cls.SAMPLE_BOM_ROWS_DXF, start=1):
            ry = top_y - 5 - r_idx * row_step
            for c_x, val_str in zip(col_xs, r_vals):
                t_ent = msp.add_text(val_str, dxfattribs={"height": 2.8})
                t_ent.dxf.insert = (c_x, ry + 2)
            msp.add_line((205, ry), (412, ry))

        # 表題欄
        msp.add_lwpolyline([(260, 5), (415, 5), (415, 45), (260, 45), (260, 5)], close=True)
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
    """Excel出力時のカスタマイズ設定"""
    template_type: str = "standard"       # "standard", "purchasing", "manufacturing"
    split_by_drawing: bool = True         # 複数図面時に図面別シートも作成するか
    merge_duplicates: bool = False        # 出力時に同一部品を名寄せ合算するか
    include_preview_sheet: bool = True    # 図面プレビュー画像シートを添付するか
    project_name: str = ""
    drawing_no: str = ""
    author_name: str = ""
    company_name: str = ""


class ExcelBOMExporter:
    """openpyxlを用いて実務品質のExcel部品表(.xlsx)を生成するエクスポーター"""

    @classmethod
    def export_to_excel(
        cls,
        output_path: str,
        parts: List[PartItem],
        documents: List[DrawingDocument],
        config: ExcelExportConfig,
    ) -> None:
        wb = openpyxl.Workbook()
        # デフォルトシートをメインの部品表として使用
        ws_main = wb.active
        ws_main.title = "部品表(BOM)"

        target_parts = copy.deepcopy(parts)
        if config.merge_duplicates:
            target_parts = cls.merge_duplicate_parts(target_parts)

        cls._populate_bom_worksheet(
            ws_main,
            target_parts,
            config,
            sheet_subtitle="総合部品表 (ALL PARTS)",
        )

        # 複数図面がある場合の図面別シート出力
        if config.split_by_drawing and len(documents) >= 1:
            grouped: Dict[str, List[PartItem]] = {}
            for p in parts:
                key = p.source_file if p.source_file else "手入力部品"
                grouped.setdefault(key, []).append(p)

            if len(grouped) > 1:
                for idx, (src_name, sub_parts) in enumerate(grouped.items(), start=1):
                    safe_title = re.sub(r"[\\/*?:\[\]]", "_", os.path.splitext(src_name)[0])[:24]
                    sheet_name = f"{idx}_{safe_title}"
                    ws_sub = wb.create_sheet(title=sheet_name)
                    cls._populate_bom_worksheet(
                        ws_sub,
                        sub_parts,
                        config,
                        sheet_subtitle=f"図面別部品表: {src_name}",
                    )

        # 図面プレビュー画像シートの添付
        temp_img_paths: List[str] = []
        if config.include_preview_sheet and documents:
            ws_prev = wb.create_sheet(title="図面プレビュー")
            ws_prev.views.sheetView[0].showGridLines = True
            ws_prev["A1"] = "収録2D図面プレビュー一覧"
            ws_prev["A1"].font = Font(name="Meiryo UI", size=14, bold=True, color="1E3A8A")

            cur_row = 3
            for doc in documents:
                for pg in doc.pages:
                    if pg.preview_image is not None:
                        ws_prev.cell(
                            row=cur_row, column=1,
                            value=f"■ ファイル: {doc.file_name} (Page {pg.page_no})  {doc.drawing_title}"
                        ).font = Font(name="Meiryo UI", size=11, bold=True, color="0F172A")
                        cur_row += 1

                        try:
                            thumb = pg.preview_image.copy()
                            thumb.thumbnail((920, 620), Image.Resampling.LANCZOS)
                            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
                            os.close(tmp_fd)
                            thumb.save(tmp_path, format="PNG")
                            temp_img_paths.append(tmp_path)

                            xl_img = OpenpyxlImage(tmp_path)
                            ws_prev.add_image(xl_img, f"A{cur_row}")
                            rows_occupied = int(thumb.height / 19) + 3
                            cur_row += rows_occupied
                        except Exception:
                            cur_row += 2

        try:
            wb.save(output_path)
        finally:
            for tp in temp_img_paths:
                if os.path.exists(tp):
                    try:
                        os.remove(tp)
                    except Exception:
                        pass

    @classmethod
    def merge_duplicate_parts(cls, parts: List[PartItem]) -> List[PartItem]:
        """同一図番・品名・仕様・材質の部品を名寄せし数量を合算する"""
        merged_map: Dict[Tuple[str, str, str, str], PartItem] = {}
        for p in parts:
            key = (
                p.part_number.strip().upper(),
                p.part_name.strip(),
                p.specification.strip().upper(),
                p.material.strip().upper(),
            )
            if key in merged_map:
                existing = merged_map[key]
                existing.quantity += p.quantity
                if p.source_file and p.source_file not in existing.source_file:
                    existing.source_file = f"{existing.source_file}, {p.source_file}"
                if p.remarks and p.remarks not in existing.remarks:
                    existing.remarks = f"{existing.remarks} / {p.remarks}".strip(" /")
            else:
                merged_map[key] = copy.deepcopy(p)

        result = list(merged_map.values())
        for idx, item in enumerate(result, start=1):
            item.item_no = str(idx)
        return result

    @classmethod
    def _populate_bom_worksheet(
        cls,
        ws: Any,
        parts: List[PartItem],
        config: ExcelExportConfig,
        sheet_subtitle: str,
    ) -> None:
        """ワークシートに美しいヘッダー・罫線・数式・書式付きの部品表を書き込む"""
        font_family = "Meiryo UI"

        # テンプレート別のタイトル定義
        template_titles = {
            "standard": "部 品 表 (BILL OF MATERIALS)",
            "purchasing": "購買・手配部品リスト (PURCHASING BOM)",
            "manufacturing": "加工・製造部品リスト (MANUFACTURING BOM)",
        }
        main_title = template_titles.get(config.template_type, "部 品 表 (BILL OF MATERIALS)")

        # スタイル定義
        title_font = Font(name=font_family, size=16, bold=True, color="FFFFFF")
        title_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")

        meta_label_font = Font(name=font_family, size=9, bold=True, color="1E293B")
        meta_label_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        meta_val_font = Font(name=font_family, size=10, bold=False, color="0F172A")

        header_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")

        data_font = Font(name=font_family, size=10, color="0F172A")
        data_font_bold = Font(name=font_family, size=10, bold=True, color="0F172A")
        even_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        odd_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        calc_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

        total_font = Font(name=font_family, size=10, bold=True, color="1E3A8A")
        total_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")

        thin_side = Side(style="thin", color="CBD5E1")
        med_side = Side(style="medium", color="1E3A8A")
        double_side = Side(style="double", color="1E3A8A")

        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        header_border = Border(left=thin_side, right=thin_side, top=med_side, bottom=med_side)
        total_border = Border(left=thin_side, right=thin_side, top=med_side, bottom=double_side)

        # 1. タイトルバナー (1行目)
        total_cols = 15
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
        title_cell = ws.cell(row=1, column=1, value=main_title)
        title_cell.font = title_font
        title_cell.fill = title_fill
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32

        # 2. プロジェクト・図面メタデータ欄 (2〜3行目)
        today_str = datetime.date.today().strftime("%Y-%m-%d")
        meta_items = [
            (2, 1, "案件・製品名", 2, 2, config.project_name or "未設定", 4),
            (2, 5, "総合図番", 2, 6, config.drawing_no or "-", 8),
            (2, 9, "作成者", 2, 10, config.author_name or "-", 11),
            (2, 12, "出力日", 2, 13, today_str, 15),
            (3, 1, "シート区分", 3, 2, sheet_subtitle, 8),
            (3, 9, "総部品行数", 3, 10, f"{len(parts)} 件", 11),
            (3, 12, "作成システム", 3, 13, APP_NAME, 15),
        ]
        for r_idx in (2, 3):
            ws.row_dimensions[r_idx].height = 20
            for c_i in range(1, total_cols + 1):
                ws.cell(row=r_idx, column=c_i).border = cell_border

        for lr, lc, l_text, vr, vc, v_text, vc_end in meta_items:
            l_cell = ws.cell(row=lr, column=lc, value=l_text)
            l_cell.font = meta_label_font
            l_cell.fill = meta_label_fill
            l_cell.alignment = Alignment(horizontal="center", vertical="center")

            if vc_end > vc:
                ws.merge_cells(start_row=vr, start_column=vc, end_row=vr, end_column=vc_end)
            v_cell = ws.cell(row=vr, column=vc, value=v_text)
            v_cell.font = meta_val_font
            v_cell.alignment = Alignment(horizontal="left", vertical="center")

        # 4行目は空行スペーサー
        ws.row_dimensions[4].height = 8

        # 5行目: テーブル列ヘッダー
        headers = [
            ("No.", 6, "center"),
            ("図番 / 品番", 18, "left"),
            ("品名 / 部品名称", 24, "left"),
            ("型式・寸法・仕様", 22, "left"),
            ("材質", 13, "center"),
            ("表面処理・熱処理", 18, "left"),
            ("数量", 9, "right"),
            ("単位", 7, "center"),
            ("単重(kg)", 11, "right"),
            ("総重量(kg)", 12, "right"),
            ("調達区分", 12, "center"),
            ("メーカー / 手配先", 16, "left"),
            ("単価(円)", 12, "right"),
            ("金額(円)", 14, "right"),
            ("抽出元図面 / 備考", 26, "left"),
        ]

        header_row = 5
        ws.row_dimensions[header_row].height = 26
        for col_idx, (h_text, col_w, _) in enumerate(headers, start=1):
            c = ws.cell(row=header_row, column=col_idx, value=h_text)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = header_border
            ws.column_dimensions[get_column_letter(col_idx)].width = col_w

        # 6行目以降: データ行
        start_data_row = 6
        for idx, p in enumerate(parts):
            r = start_data_row + idx
            ws.row_dimensions[r].height = 22
            row_fill = even_fill if idx % 2 == 0 else odd_fill

            remark_combined = p.remarks
            if p.source_file:
                remark_combined = f"[{p.source_file}] {p.remarks}".strip()

            # Excel数式: J列(総重量) = G列(数量) * I列(単重), N列(金額) = G列(数量) * M列(単価)
            weight_formula = f"=G{r}*I{r}"
            amount_formula = f"=G{r}*M{r}"

            row_values = [
                int(p.item_no) if p.item_no.isdigit() else p.item_no,
                p.part_number,
                p.part_name,
                p.specification,
                p.material,
                p.treatment,
                p.quantity,
                p.unit,
                p.unit_weight,
                weight_formula,
                p.category,
                p.manufacturer,
                p.unit_price,
                amount_formula,
                remark_combined,
            ]

            for col_idx, val in enumerate(row_values, start=1):
                cell = ws.cell(row=r, column=col_idx, value=val)
                cell.font = data_font_bold if col_idx in (3, 10, 14) else data_font
                cell.fill = calc_fill if col_idx in (10, 14) else row_fill
                cell.border = cell_border
                align_h = headers[col_idx - 1][2]
                cell.alignment = Alignment(horizontal=align_h, vertical="center")

                # 数値フォーマット
                if col_idx == 7:
                    cell.number_format = "#,##0.##"
                elif col_idx in (9, 10):
                    cell.number_format = "0.000"
                elif col_idx in (13, 14):
                    cell.number_format = "#,##0"

        # 末尾: 合計(SUM)集計行
        end_data_row = max(start_data_row + len(parts) - 1, start_data_row)
        total_row = start_data_row + len(parts)
        ws.row_dimensions[total_row].height = 25

        for col_idx in range(1, total_cols + 1):
            c = ws.cell(row=total_row, column=col_idx, value="")
            c.font = total_font
            c.fill = total_fill
            c.border = total_border
            c.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=total_row, column=3, value="合 計 (TOTAL)").alignment = Alignment(
            horizontal="center", vertical="center"
        )
        if parts:
            c_qty = ws.cell(row=total_row, column=7, value=f"=SUM(G{start_data_row}:G{end_data_row})")
            c_qty.number_format = "#,##0.##"

            c_wt = ws.cell(row=total_row, column=10, value=f"=SUM(J{start_data_row}:J{end_data_row})")
            c_wt.number_format = "0.000"

            c_amt = ws.cell(row=total_row, column=14, value=f"=SUM(N{start_data_row}:N{end_data_row})")
            c_amt.number_format = "#,##0"

        # テンプレートに応じた不要列の非表示最適化
        if config.template_type == "purchasing":
            ws.column_dimensions["F"].hidden = True
            ws.column_dimensions["I"].hidden = True
        elif config.template_type == "manufacturing":
            ws.column_dimensions["M"].hidden = True
            ws.column_dimensions["N"].hidden = True

        # オートフィルタ・ウィンドウ枠固定・A4横1ページ幅印刷設定
        ws.auto_filter.ref = f"A{header_row}:O{total_row}"
        ws.freeze_panes = f"A{start_data_row}"
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0


# ==============================================================================
# サブダイアログ群 (重量計算機 / Excel出力設定 / 使い方ガイド)
# ==============================================================================

class WeightCalculatorDialog(tk.Toplevel):
    """形状(板材・丸棒・丸パイプ・角パイプ・六角棒)と材質比重から単重(kg)を計算するダイアログ"""

    def __init__(self, parent: tk.Misc, initial_material: str, on_apply_callback: Any):
        super().__init__(parent)
        self.title("部品重量計算ツール - PartLens_Miya-v1.0.0")
        self.geometry("520x470")
        self.configure(bg=LightPalette.BG_MAIN)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        enforce_windows_light_mode(self)

        self.on_apply_callback = on_apply_callback
        self.calculated_weight: float = 0.0
        self.generated_spec: str = ""

        self._build_ui(initial_material)
        self._recalculate()

    def _build_ui(self, initial_material: str) -> None:
        header = tk.Frame(self, bg=LightPalette.BG_HEADER, padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text="形状・材質からの単重(kg)自動計算機",
            bg=LightPalette.BG_HEADER,
            fg=LightPalette.FG_HEADER,
            font=("Meiryo UI", 11, "bold"),
        ).pack(anchor="w")

        body = tk.Frame(self, bg=LightPalette.BG_PANEL, padx=20, pady=16)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        # 1. 形状選択
        tk.Label(body, text="1. 部品の形状:", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", pady=6
        )
        self.shape_var = tk.StringVar(value="プレート / 角材 (t × W × L)")
        shapes = [
            "プレート / 角材 (t × W × L)",
            "丸棒 / 円柱 (φD × L)",
            "丸パイプ (外径φD × 肉厚t × L)",
            "角パイプ (辺A × 辺B × 肉厚t × L)",
            "六角棒 (対辺H × L)",
        ]
        self.shape_combo = ttk.Combobox(body, textvariable=self.shape_var, values=shapes, state="readonly", width=34)
        self.shape_combo.grid(row=0, column=1, sticky="ew", pady=6, padx=6)
        self.shape_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_shape_change())

        # 2. 材質選択
        tk.Label(body, text="2. 材質 (JIS規格):", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9, "bold")).grid(
            row=1, column=0, sticky="w", pady=6
        )
        mat_default = initial_material if initial_material in MATERIAL_MASTER else "SS400"
        self.mat_var = tk.StringVar(value=mat_default)
        self.mat_combo = ttk.Combobox(
            body, textvariable=self.mat_var, values=list(MATERIAL_MASTER.keys()), width=34
        )
        self.mat_combo.grid(row=1, column=1, sticky="ew", pady=6, padx=6)
        self.mat_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_mat_change())

        # 3. 比重
        tk.Label(body, text="3. 比重 (g/cm³):", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9)).grid(
            row=2, column=0, sticky="w", pady=6
        )
        init_density = MATERIAL_MASTER.get(mat_default, {}).get("density", 7.85)
        self.density_var = tk.StringVar(value=str(init_density))
        self.density_entry = ttk.Entry(body, textvariable=self.density_var, width=18)
        self.density_entry.grid(row=2, column=1, sticky="w", pady=6, padx=6)
        self.density_var.trace_add("write", lambda *_: self._recalculate())

        # 4. 寸法パラメータ入力欄 (最大4個)
        self.dim_labels: List[tk.Label] = []
        self.dim_vars: List[tk.StringVar] = []
        self.dim_entries: List[ttk.Entry] = []
        default_vals = ["10", "100", "200", "3.2"]

        for i in range(4):
            lbl = tk.Label(body, text=f"寸法 {i + 1} (mm):", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9))
            lbl.grid(row=3 + i, column=0, sticky="w", pady=5)
            v = tk.StringVar(value=default_vals[i])
            v.trace_add("write", lambda *_: self._recalculate())
            ent = ttk.Entry(body, textvariable=v, width=20)
            ent.grid(row=3 + i, column=1, sticky="w", pady=5, padx=6)
            self.dim_labels.append(lbl)
            self.dim_vars.append(v)
            self.dim_entries.append(ent)

        # 計算結果表示カード
        res_frame = tk.Frame(body, bg=LightPalette.BG_SELECTED, padx=14, pady=10, highlightthickness=1, highlightbackground=LightPalette.BORDER_FOCUS)
        res_frame.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(14, 10))

        self.result_lbl = tk.Label(
            res_frame,
            text="計算単重: 0.000 kg",
            bg=LightPalette.BG_SELECTED,
            fg=LightPalette.FG_ACCENT,
            font=("Meiryo UI", 13, "bold"),
        )
        self.result_lbl.pack(anchor="w")

        self.spec_lbl = tk.Label(
            res_frame,
            text="寸法表記: t10x100x200",
            bg=LightPalette.BG_SELECTED,
            fg=LightPalette.FG_SUB,
            font=("Meiryo UI", 9),
        )
        self.spec_lbl.pack(anchor="w", pady=(2, 0))

        # ボタンバー
        btn_bar = tk.Frame(body, bg=LightPalette.BG_PANEL)
        btn_bar.grid(row=8, column=0, columnspan=2, sticky="e", pady=(8, 0))

        tk.Button(
            btn_bar,
            text="選択中の部品行へ反映",
            bg=LightPalette.BTN_PRIMARY_BG,
            fg=LightPalette.BTN_PRIMARY_FG,
            activebackground=LightPalette.BTN_PRIMARY_ACTIVE,
            activeforeground=LightPalette.BTN_PRIMARY_FG,
            relief=tk.FLAT,
            padx=16,
            pady=6,
            font=("Meiryo UI", 9, "bold"),
            cursor="hand2",
            command=self._apply_and_close,
        ).pack(side=tk.LEFT, padx=6)

        tk.Button(
            btn_bar,
            text="閉じる",
            bg=LightPalette.BTN_DEFAULT_BG,
            fg=LightPalette.BTN_DEFAULT_FG,
            relief=tk.SOLID,
            bd=1,
            padx=14,
            pady=6,
            font=("Meiryo UI", 9),
            command=self.destroy,
        ).pack(side=tk.LEFT)

        self._on_shape_change()

    def _on_mat_change(self) -> None:
        m = self.mat_var.get()
        if m in MATERIAL_MASTER:
            self.density_var.set(str(MATERIAL_MASTER[m]["density"]))
        self._recalculate()

    def _on_shape_change(self) -> None:
        shape = self.shape_var.get()
        if "プレート" in shape:
            configs = [("板厚 t (mm):", True), ("幅 W (mm):", True), ("長さ L (mm):", True), ("-", False)]
        elif "丸棒" in shape:
            configs = [("外径 φD (mm):", True), ("長さ L (mm):", True), ("-", False), ("-", False)]
        elif "丸パイプ" in shape:
            configs = [("外径 φD (mm):", True), ("肉厚 t (mm):", True), ("長さ L (mm):", True), ("-", False)]
        elif "角パイプ" in shape:
            configs = [("辺 A (mm):", True), ("辺 B (mm):", True), ("肉厚 t (mm):", True), ("長さ L (mm):", True)]
        else:
            configs = [("六角対辺 H (mm):", True), ("長さ L (mm):", True), ("-", False), ("-", False)]

        for i, (label_text, active) in enumerate(configs):
            self.dim_labels[i].configure(text=label_text)
            if active:
                self.dim_entries[i].configure(state="normal")
            else:
                self.dim_entries[i].configure(state="disabled")
        self._recalculate()

    def _recalculate(self) -> None:
        try:
            density = float(self.density_var.get())
            vals = [float(v.get() or 0.0) for v in self.dim_vars]
        except ValueError:
            return

        shape = self.shape_var.get()
        vol_mm3 = 0.0
        spec_str = ""

        if "プレート" in shape:
            t, w, length = vals[0], vals[1], vals[2]
            vol_mm3 = t * w * length
            spec_str = f"t{t:g}x{w:g}x{length:g}"
        elif "丸棒" in shape:
            d, length = vals[0], vals[1]
            r = d * 0.5
            vol_mm3 = math.pi * (r * r) * length
            spec_str = f"φ{d:g}x{length:g}"
        elif "丸パイプ" in shape:
            d, t, length = vals[0], vals[1], vals[2]
            r_out = d * 0.5
            r_in = max(r_out - t, 0.0)
            vol_mm3 = math.pi * (r_out * r_out - r_in * r_in) * length
            spec_str = f"φ{d:g}xt{t:g}x{length:g}"
        elif "角パイプ" in shape:
            a, b, t, length = vals[0], vals[1], vals[2], vals[3]
            inner_a = max(a - 2.0 * t, 0.0)
            inner_b = max(b - 2.0 * t, 0.0)
            vol_mm3 = (a * b - inner_a * inner_b) * length
            spec_str = f"{a:g}x{b:g}xt{t:g}x{length:g}"
        else:
            h, length = vals[0], vals[1]
            # 正六角形面積 = (sqrt(3)/2) * H^2
            vol_mm3 = (math.sqrt(3.0) * 0.5) * (h * h) * length
            spec_str = f"HEX{h:g}x{length:g}"

        self.calculated_weight = round(max(vol_mm3, 0.0) * density / 1_000_000.0, 3)
        self.generated_spec = spec_str
        self.result_lbl.configure(text=f"計算単重: {self.calculated_weight:.3f} kg")
        self.spec_lbl.configure(text=f"寸法表記: {spec_str}  (材質: {self.mat_var.get()})")

    def _apply_and_close(self) -> None:
        if self.on_apply_callback:
            self.on_apply_callback(self.calculated_weight, self.generated_spec, self.mat_var.get())
        self.destroy()


class ExcelSettingsDialog(tk.Toplevel):
    """Excel出力フォーマットやシート構成を設定するダイアログ"""

    def __init__(self, parent: tk.Misc, config: ExcelExportConfig, on_save_callback: Any):
        super().__init__(parent)
        self.title("Excel出力設定 - PartLens_Miya-v1.0.0")
        self.geometry("480x360")
        self.configure(bg=LightPalette.BG_MAIN)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        enforce_windows_light_mode(self)

        self.config_obj = config
        self.on_save_callback = on_save_callback
        self._build_ui()

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=LightPalette.BG_HEADER, padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text="Excel部品表 (.xlsx) 出力オプション設定",
            bg=LightPalette.BG_HEADER,
            fg=LightPalette.FG_HEADER,
            font=("Meiryo UI", 11, "bold"),
        ).pack(anchor="w")

        body = tk.Frame(self, bg=LightPalette.BG_PANEL, padx=20, pady=16)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        tk.Label(
            body, text="出力テンプレート形式:", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9, "bold")
        ).pack(anchor="w", pady=(0, 6))

        self.tpl_var = tk.StringVar(value=self.config_obj.template_type)
        tpl_options = [
            ("標準部品表フォーマット（全項目・推奨）", "standard"),
            ("購買・見積手配用フォーマット（メーカー・金額重視）", "purchasing"),
            ("製造・加工手配用フォーマット（材質・寸法・重量重視）", "manufacturing"),
        ]
        for label_t, val_t in tpl_options:
            ttk.Radiobutton(body, text=label_t, variable=self.tpl_var, value=val_t).pack(anchor="w", padx=12, pady=2)

        ttk.Separator(body, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=12)

        self.split_var = tk.BooleanVar(value=self.config_obj.split_by_drawing)
        ttk.Checkbutton(
            body, text="複数図面がある場合に「図面別シート」も同時作成する", variable=self.split_var
        ).pack(anchor="w", pady=3)

        self.merge_var = tk.BooleanVar(value=self.config_obj.merge_duplicates)
        ttk.Checkbutton(
            body, text="出力時に同一品番・品名を名寄せ（数量合算）して集計する", variable=self.merge_var
        ).pack(anchor="w", pady=3)

        self.prev_var = tk.BooleanVar(value=self.config_obj.include_preview_sheet)
        ttk.Checkbutton(
            body, text="別シートに「2D図面プレビュー画像」を添付する", variable=self.prev_var
        ).pack(anchor="w", pady=3)

        btn_bar = tk.Frame(body, bg=LightPalette.BG_PANEL)
        btn_bar.pack(fill=tk.X, pady=(18, 0))

        tk.Button(
            btn_bar,
            text="設定を保存",
            bg=LightPalette.BTN_PRIMARY_BG,
            fg=LightPalette.BTN_PRIMARY_FG,
            relief=tk.FLAT,
            padx=18,
            pady=6,
            font=("Meiryo UI", 9, "bold"),
            cursor="hand2",
            command=self._save,
        ).pack(side=tk.RIGHT, padx=(6, 0))

        tk.Button(
            btn_bar,
            text="キャンセル",
            bg=LightPalette.BTN_DEFAULT_BG,
            fg=LightPalette.BTN_DEFAULT_FG,
            relief=tk.SOLID,
            bd=1,
            padx=14,
            pady=6,
            font=("Meiryo UI", 9),
            command=self.destroy,
        ).pack(side=tk.RIGHT)

    def _save(self) -> None:
        self.config_obj.template_type = self.tpl_var.get()
        self.config_obj.split_by_drawing = self.split_var.get()
        self.config_obj.merge_duplicates = self.merge_var.get()
        self.config_obj.include_preview_sheet = self.prev_var.get()
        if self.on_save_callback:
            self.on_save_callback()
        self.destroy()


class ReadmeViewerDialog(tk.Toplevel):
    """同梱のREADME.md（取扱説明書）をアプリ内でわかりやすく表示するダイアログ"""

    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        self.title(f"使い方ガイド (取扱説明書) - {APP_NAME}")
        self.geometry("760x620")
        self.configure(bg=LightPalette.BG_MAIN)
        self.transient(parent)
        enforce_windows_light_mode(self)

        header = tk.Frame(self, bg=LightPalette.BG_HEADER, padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text=f"{APP_NAME} 取扱説明書・操作ガイド",
            bg=LightPalette.BG_HEADER,
            fg=LightPalette.FG_HEADER,
            font=("Meiryo UI", 11, "bold"),
        ).pack(side=tk.LEFT)

        text_frame = tk.Frame(self, bg=LightPalette.BG_PANEL, padx=12, pady=12)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL)
        txt_widget = tk.Text(
            text_frame,
            bg=LightPalette.BG_PANEL,
            fg=LightPalette.FG_MAIN,
            font=("Meiryo UI", 10),
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            relief=tk.FLAT,
            padx=10,
            pady=8,
        )
        scrollbar.configure(command=txt_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        txt_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        readme_path = get_resource_path("README.md")
        content = "取扱説明書ファイル (README.md) が見つかりませんでした。"
        if os.path.exists(readme_path):
            try:
                with open(readme_path, "r", encoding="utf-8") as f:
                    content = f.read()
            except Exception as e:
                content = f"読み込みエラー: {e}"

        txt_widget.insert("1.0", content)
        txt_widget.configure(state=tk.DISABLED)


# ==============================================================================
# メインアプリケーション GUI (PartLensApp)
# ==============================================================================

class PartLensApp:
    """PartLens_Miya-v1.0.0 メインデスクトップアプリケーションクラス"""

    TREE_COLUMNS = (
        "item_no",
        "part_number",
        "part_name",
        "specification",
        "material",
        "treatment",
        "quantity",
        "unit",
        "unit_weight",
        "total_weight",
        "category",
        "manufacturer",
        "unit_price",
        "amount",
        "source_file",
        "remarks",
    )

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1480x900")
        self.root.minsize(1180, 720)
        self.root.configure(bg=LightPalette.BG_MAIN)

        # アイコン設定
        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        # ライトモード完全固定スタイルの適用
        self._setup_light_mode_styles()
        enforce_windows_light_mode(self.root)

        # アプリ状態データ
        self.documents: List[DrawingDocument] = []
        self.current_doc_idx: int = -1
        self.current_page_idx: int = 0
        self.parts: List[PartItem] = []
        self.selected_part_uid: Optional[str] = None

        # Undo / Redo 履歴スタック
        self.undo_stack: List[List[PartItem]] = []
        self.redo_stack: List[List[PartItem]] = []

        # プレビューキャンバス状態
        self.zoom_scale: float = 1.0
        self.canvas_offset_x: float = 0.0
        self.canvas_offset_y: float = 0.0
        self.interaction_mode: str = "pan"  # "pan" or "select_area"
        self.drag_start_x: int = 0
        self.drag_start_y: int = 0
        self.selection_rect_id: Optional[int] = None
        self.highlight_bbox: Optional[Tuple[float, float, float, float]] = None
        self._tk_preview_img: Optional[ImageTk.PhotoImage] = None
        self._inline_editor: Optional[tk.Widget] = None

        # Excel出力設定
        self.excel_config = ExcelExportConfig(
            project_name="新規機械設計プロジェクト",
            drawing_no="ASM-0001",
            author_name="",
        )

        # UI構築
        self._build_header()
        self._build_main_workspace()
        self._build_statusbar()
        self._bind_shortcuts()

    # --------------------------------------------------------------------------
    # ライトモード固定テーマ設定
    # --------------------------------------------------------------------------
    def _setup_light_mode_styles(self) -> None:
        """システムがダークモードでもアプリ内の全ウィジェットをライトモードに固定する"""
        self.root.option_add("*Background", LightPalette.BG_PANEL)
        self.root.option_add("*Foreground", LightPalette.FG_MAIN)
        self.root.option_add("*Entry.Background", "#FFFFFF")
        self.root.option_add("*Entry.Foreground", LightPalette.FG_MAIN)
        self.root.option_add("*Listbox.Background", "#FFFFFF")
        self.root.option_add("*Listbox.Foreground", LightPalette.FG_MAIN)
        self.root.option_add("*Listbox.selectBackground", LightPalette.BG_SELECTED)
        self.root.option_add("*Listbox.selectForeground", LightPalette.FG_ACCENT)
        self.root.option_add("*TCombobox*Listbox.background", "#FFFFFF")
        self.root.option_add("*TCombobox*Listbox.foreground", LightPalette.FG_MAIN)
        self.root.option_add("*TCombobox*Listbox.selectBackground", LightPalette.BG_SELECTED)
        self.root.option_add("*TCombobox*Listbox.selectForeground", LightPalette.FG_ACCENT)

        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure(".", background=LightPalette.BG_PANEL, foreground=LightPalette.FG_MAIN, font=("Meiryo UI", 9))
        style.configure("TFrame", background=LightPalette.BG_PANEL)
        style.configure("Main.TFrame", background=LightPalette.BG_MAIN)
        style.configure("Sub.TFrame", background=LightPalette.BG_SUB)
        style.configure("TLabel", background=LightPalette.BG_PANEL, foreground=LightPalette.FG_MAIN, font=("Meiryo UI", 9))
        style.configure("Sub.TLabel", background=LightPalette.BG_SUB, foreground=LightPalette.FG_SUB, font=("Meiryo UI", 9))
        style.configure("TLabelframe", background=LightPalette.BG_PANEL, bordercolor=LightPalette.BORDER)
        style.configure("TLabelframe.Label", background=LightPalette.BG_PANEL, foreground=LightPalette.FG_ACCENT, font=("Meiryo UI", 9, "bold"))

        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            foreground=LightPalette.FG_MAIN,
            bordercolor=LightPalette.BORDER,
            insertcolor=LightPalette.FG_MAIN,
            padding=4,
        )
        style.configure(
            "TCombobox",
            fieldbackground="#FFFFFF",
            background=LightPalette.BG_SUB,
            foreground=LightPalette.FG_MAIN,
            arrowcolor=LightPalette.FG_MAIN,
            bordercolor=LightPalette.BORDER,
            padding=4,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#FFFFFF"), ("disabled", LightPalette.BG_SUB)],
            foreground=[("readonly", LightPalette.FG_MAIN), ("disabled", LightPalette.FG_MUTED)],
            selectbackground=[("readonly", "#FFFFFF")],
            selectforeground=[("readonly", LightPalette.FG_MAIN)],
        )

        style.configure(
            "TCheckbutton",
            background=LightPalette.BG_PANEL,
            foreground=LightPalette.FG_MAIN,
            indicatorbackground="#FFFFFF",
            indicatorforeground=LightPalette.BTN_PRIMARY_BG,
        )
        style.map("TCheckbutton", background=[("active", LightPalette.BG_PANEL)])

        style.configure(
            "TRadiobutton",
            background=LightPalette.BG_PANEL,
            foreground=LightPalette.FG_MAIN,
            indicatorbackground="#FFFFFF",
            indicatorforeground=LightPalette.BTN_PRIMARY_BG,
        )
        style.map("TRadiobutton", background=[("active", LightPalette.BG_PANEL)])

        style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=LightPalette.FG_MAIN,
            rowheight=26,
            bordercolor=LightPalette.BORDER,
            font=("Meiryo UI", 9),
        )
        style.configure(
            "Treeview.Heading",
            background="#E2E8F0",
            foreground="#0F172A",
            relief="flat",
            font=("Meiryo UI", 9, "bold"),
            padding=(4, 5),
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#CBD5E1")],
        )
        style.map(
            "Treeview",
            background=[("selected", LightPalette.BG_SELECTED)],
            foreground=[("selected", LightPalette.FG_ACCENT)],
        )

        style.configure("TPanedwindow", background=LightPalette.BG_MAIN)
        style.configure("TScrollbar", background=LightPalette.BG_SUB, troughcolor=LightPalette.BG_MAIN, arrowcolor=LightPalette.FG_SUB)

    # --------------------------------------------------------------------------
    # UI構築: 上部ヘッダー
    # --------------------------------------------------------------------------
    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=LightPalette.BG_HEADER, padx=16, pady=10)
        header.pack(fill=tk.X)

        # 左側タイトル
        title_box = tk.Frame(header, bg=LightPalette.BG_HEADER)
        title_box.pack(side=tk.LEFT)

        tk.Label(
            title_box,
            text="PartLens Miya",
            bg=LightPalette.BG_HEADER,
            fg="#FFFFFF",
            font=("Meiryo UI", 15, "bold"),
        ).pack(side=tk.LEFT)

        tk.Label(
            title_box,
            text="v1.0.0   |   2D図面 (PDF / CAD DXF) 部品表(BOM) 自動抽出・Excel作成システム",
            bg=LightPalette.BG_HEADER,
            fg="#BFDBFE",
            font=("Meiryo UI", 9),
            padx=10,
        ).pack(side=tk.LEFT, pady=(4, 0))

        # 右側アクションボタン群
        btn_box = tk.Frame(header, bg=LightPalette.BG_HEADER)
        btn_box.pack(side=tk.RIGHT)

        self._create_header_btn(btn_box, "図面ファイルを開く (PDF/DXF)", self.action_open_files, bg="#2563EB", fg="#FFFFFF")
        self._create_header_btn(btn_box, "フォルダ一括読込", self.action_open_folder, bg="#3B82F6", fg="#FFFFFF")
        self._create_header_btn(btn_box, "サンプル図面で試す", self.action_load_sample_drawings, bg="#059669", fg="#FFFFFF")
        self._create_header_btn(btn_box, "プロジェクト保存", self.action_save_project, bg="#FFFFFF", fg="#1E293B")
        self._create_header_btn(btn_box, "プロジェクト読込", self.action_load_project, bg="#FFFFFF", fg="#1E293B")
        self._create_header_btn(btn_box, "使い方ガイド", self.action_show_readme, bg="#FEF3C7", fg="#92400E")

    def _create_header_btn(self, parent: tk.Misc, text: str, command: Any, bg: str, fg: str) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=bg,
            activeforeground=fg,
            relief=tk.FLAT,
            padx=12,
            pady=5,
            font=("Meiryo UI", 9, "bold"),
            cursor="hand2",
        )
        btn.pack(side=tk.LEFT, padx=4)
        return btn

    # --------------------------------------------------------------------------
    # UI構築: 左右分割メインワークスペース
    # --------------------------------------------------------------------------
    def _build_main_workspace(self) -> None:
        self.paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        # 左ペイン: 2D図面リスト＆プレビューキャンバス
        left_pane = tk.Frame(self.paned, bg=LightPalette.BG_PANEL, bd=1, relief=tk.SOLID)
        self.paned.add(left_pane, weight=4)
        self._build_left_preview_pane(left_pane)

        # 右ペイン: 部品表エディタ＆詳細フォーム＆Excel出力
        right_pane = tk.Frame(self.paned, bg=LightPalette.BG_PANEL, bd=1, relief=tk.SOLID)
        self.paned.add(right_pane, weight=6)
        self._build_right_bom_pane(right_pane)

    def _build_left_preview_pane(self, parent: tk.Frame) -> None:
        # 1. 読み込み図面リストセクション
        doc_frame = tk.Frame(parent, bg=LightPalette.BG_SUB, padx=8, pady=6)
        doc_frame.pack(fill=tk.X)

        top_lbl_bar = tk.Frame(doc_frame, bg=LightPalette.BG_SUB)
        top_lbl_bar.pack(fill=tk.X)
        tk.Label(
            top_lbl_bar,
            text="■ 読み込み済み 2D図面一覧 (PDF / DXF)",
            bg=LightPalette.BG_SUB,
            fg=LightPalette.FG_ACCENT,
            font=("Meiryo UI", 9, "bold"),
        ).pack(side=tk.LEFT)

        tk.Button(
            top_lbl_bar,
            text="一覧クリア",
            command=self.action_clear_documents,
            bg=LightPalette.BTN_DEFAULT_BG,
            fg=LightPalette.FG_SUB,
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=1,
            font=("Meiryo UI", 8),
        ).pack(side=tk.RIGHT, padx=2)

        tk.Button(
            top_lbl_bar,
            text="現在図面を再解析",
            command=self.action_reparse_current_page,
            bg=LightPalette.BTN_DEFAULT_BG,
            fg=LightPalette.FG_ACCENT,
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=1,
            font=("Meiryo UI", 8, "bold"),
        ).pack(side=tk.RIGHT, padx=4)

        list_container = tk.Frame(doc_frame, bg=LightPalette.BG_PANEL)
        list_container.pack(fill=tk.X, pady=(4, 0))

        self.doc_listbox = tk.Listbox(
            list_container,
            height=4,
            bg="#FFFFFF",
            fg=LightPalette.FG_MAIN,
            selectbackground=LightPalette.BG_SELECTED,
            selectforeground=LightPalette.FG_ACCENT,
            font=("Meiryo UI", 9),
            relief=tk.SOLID,
            bd=1,
            exportselection=False,
        )
        doc_scroll = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.doc_listbox.yview)
        self.doc_listbox.configure(yscrollcommand=doc_scroll.set)
        self.doc_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        doc_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.doc_listbox.bind("<<ListboxSelect>>", self._on_select_document)

        # 2. プレビュー操作ツールバー
        ctrl_bar = tk.Frame(parent, bg=LightPalette.BG_PANEL, padx=8, pady=6)
        ctrl_bar.pack(fill=tk.X)

        # ページ切替
        self.btn_prev_page = tk.Button(
            ctrl_bar, text="◀", width=2, command=lambda: self._change_page(-1),
            bg=LightPalette.BTN_DEFAULT_BG, fg=LightPalette.FG_MAIN, relief=tk.SOLID, bd=1
        )
        self.btn_prev_page.pack(side=tk.LEFT)

        self.page_info_lbl = tk.Label(ctrl_bar, text="Page: 0 / 0", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, width=11, font=("Meiryo UI", 8, "bold"))
        self.page_info_lbl.pack(side=tk.LEFT, padx=2)

        self.btn_next_page = tk.Button(
            ctrl_bar, text="▶", width=2, command=lambda: self._change_page(1),
            bg=LightPalette.BTN_DEFAULT_BG, fg=LightPalette.FG_MAIN, relief=tk.SOLID, bd=1
        )
        self.btn_next_page.pack(side=tk.LEFT, padx=(0, 8))

        # ズーム操作
        tk.Button(
            ctrl_bar, text="－", width=2, command=lambda: self._zoom_by(0.8),
            bg=LightPalette.BTN_DEFAULT_BG, fg=LightPalette.FG_MAIN, relief=tk.SOLID, bd=1
        ).pack(side=tk.LEFT)
        self.zoom_lbl = tk.Label(ctrl_bar, text="100%", bg=LightPalette.BG_PANEL, fg=LightPalette.FG_MAIN, width=6, font=("Meiryo UI", 8))
        self.zoom_lbl.pack(side=tk.LEFT)
        tk.Button(
            ctrl_bar, text="＋", width=2, command=lambda: self._zoom_by(1.25),
            bg=LightPalette.BTN_DEFAULT_BG, fg=LightPalette.FG_MAIN, relief=tk.SOLID, bd=1
        ).pack(side=tk.LEFT)
        tk.Button(
            ctrl_bar, text="全体フィット", command=self._fit_canvas_view,
            bg=LightPalette.BTN_DEFAULT_BG, fg=LightPalette.FG_MAIN, relief=tk.SOLID, bd=1, padx=6, font=("Meiryo UI", 8)
        ).pack(side=tk.LEFT, padx=4)

        # モード切替ボタン（パン移動 vs 矩形範囲選択して部品抽出）
        self.btn_mode_select = tk.Button(
            ctrl_bar,
            text="範囲選択して部品抽出: OFF",
            command=self._toggle_interaction_mode,
            bg=LightPalette.BTN_DEFAULT_BG,
            fg=LightPalette.FG_ACCENT,
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=2,
            font=("Meiryo UI", 8, "bold"),
            cursor="hand2",
        )
        self.btn_mode_select.pack(side=tk.RIGHT)

        # ガイドバー
        self.preview_hint_lbl = tk.Label(
            parent,
            text="操作ヒント: マウスホイールで拡大縮小、左ドラッグで図面移動。「範囲選択して部品抽出」をONにすると、囲んだ範囲から部品を読み取ります。",
            bg="#EFF6FF",
            fg="#1E40AF",
            font=("Meiryo UI", 8),
            anchor="w",
            padx=8,
            pady=4,
        )
        self.preview_hint_lbl.pack(fill=tk.X)

        # 3. 2D図面プレビューキャンバス
        canvas_container = tk.Frame(parent, bg=LightPalette.BG_CANVAS)
        canvas_container.pack(fill=tk.BOTH, expand=True)

        self.preview_canvas = tk.Canvas(
            canvas_container,
            bg=LightPalette.BG_CANVAS,
            highlightthickness=0,
            cursor="fleur",
        )
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)

        # キャンバスイベント
        self.preview_canvas.bind("<Configure>", lambda _e: self._redraw_canvas())
        self.preview_canvas.bind("<ButtonPress-1>", self._on_canvas_mouse_down)
        self.preview_canvas.bind("<B1-Motion>", self._on_canvas_mouse_move)
        self.preview_canvas.bind("<ButtonRelease-1>", self._on_canvas_mouse_up)
        self.preview_canvas.bind("<MouseWheel>", self._on_canvas_mouse_wheel)

    def _build_right_bom_pane(self, parent: tk.Frame) -> None:
        # 1. プロジェクト・図面ヘッダー情報欄
        meta_bar = tk.Frame(parent, bg=LightPalette.BG_SUB, padx=10, pady=8)
        meta_bar.pack(fill=tk.X)

        tk.Label(meta_bar, text="案件名/製品名:", bg=LightPalette.BG_SUB, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        self.proj_name_var = tk.StringVar(value=self.excel_config.project_name)
        ttk.Entry(meta_bar, textvariable=self.proj_name_var, width=26).grid(row=0, column=1, padx=(4, 12), sticky="w")

        tk.Label(meta_bar, text="総合図番:", bg=LightPalette.BG_SUB, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9, "bold")).grid(row=0, column=2, sticky="w")
        self.dwg_no_var = tk.StringVar(value=self.excel_config.drawing_no)
        ttk.Entry(meta_bar, textvariable=self.dwg_no_var, width=16).grid(row=0, column=3, padx=(4, 12), sticky="w")

        tk.Label(meta_bar, text="作成者:", bg=LightPalette.BG_SUB, fg=LightPalette.FG_MAIN, font=("Meiryo UI", 9, "bold")).grid(row=0, column=4, sticky="w")
        self.author_var = tk.StringVar(value=self.excel_config.author_name)
        ttk.Entry(meta_bar, textvariable=self.author_var, width=14).grid(row=0, column=5, padx=(4, 12), sticky="w")

        tk.Label(meta_bar, text="検索フィルタ:", bg=LightPalette.BG_SUB, fg=LightPalette.FG_ACCENT, font=("Meiryo UI", 9, "bold")).grid(row=0, column=6, sticky="w")
        self.filter_var = tk.StringVar(value="")
        self.filter_var.trace_add("write", lambda *_: self._refresh_treeview())
        ttk.Entry(meta_bar, textvariable=self.filter_var, width=18).grid(row=0, column=7, padx=(4, 4), sticky="w")

        # 2. 部品表編集ツールバー
        tool_bar = tk.Frame(parent, bg=LightPalette.BG_PANEL, padx=8, pady=6)
        tool_bar.pack(fill=tk.X)

        self._create_tool_btn(tool_bar, "＋ 行追加", self.action_add_row, fg="#1D4ED8", bold=True)
        self._create_tool_btn(tool_bar, "行複製", self.action_duplicate_row)
        self._create_tool_btn(tool_bar, "－ 行削除", self.action_delete_rows, fg="#DC2626")
        self._create_tool_btn(tool_bar, "↑ 上へ", lambda: self.action_move_row(-1))
        self._create_tool_btn(tool_bar, "↓ 下へ", lambda: self.action_move_row(1))
        self._create_tool_btn(tool_bar, "No.連番振直", self.action_renumber_rows)
        self._create_tool_btn(tool_bar, "重複部品の名寄せ統合", self.action_merge_duplicates, fg="#047857", bold=True)
        self._create_tool_btn(tool_bar, "寸法から重量自動計算", self.action_auto_calc_weights, fg="#1D4ED8")
        self._create_tool_btn(tool_bar, "重量計算機...", self.action_open_weight_calculator, fg="#7C3AED", bold=True)
        self._create_tool_btn(tool_bar, "元に戻す(Undo)", self.action_undo)
        self._create_tool_btn(tool_bar, "やり直す(Redo)", self.action_redo)
        self._create_tool_btn(tool_bar, "全クリア", self.action_clear_all_parts, fg="#64748B")

        # 3. 部品表 (Treeview)
        table_container = tk.Frame(parent, bg=LightPalette.BG_PANEL, padx=8)
        table_container.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(
            table_container,
            columns=self.TREE_COLUMNS,
            show="headings",
            selectmode="extended",
        )
        y_scroll = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_container, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        col_defs = [
            ("item_no", "No.", 46, "center"),
            ("part_number", "図番 / 品番", 115, "w"),
            ("part_name", "品名 / 部品名称", 155, "w"),
            ("specification", "型式・寸法・仕様", 145, "w"),
            ("material", "材質", 85, "center"),
            ("treatment", "表面処理・熱処理", 125, "w"),
            ("quantity", "数量", 58, "e"),
            ("unit", "単位", 48, "center"),
            ("unit_weight", "単重(kg)", 72, "e"),
            ("total_weight", "総重量(kg)", 78, "e"),
            ("category", "調達区分", 85, "center"),
            ("manufacturer", "メーカー", 105, "w"),
            ("unit_price", "単価(円)", 78, "e"),
            ("amount", "金額(円)", 88, "e"),
            ("source_file", "抽出元図面", 115, "w"),
            ("remarks", "備考", 140, "w"),
        ]
        for cid, heading_txt, width_px, anchor_pos in col_defs:
            self.tree.heading(cid, text=heading_txt, command=lambda c=cid: self._sort_tree_by_column(c))
            self.tree.column(cid, width=width_px, minwidth=40, anchor=anchor_pos)

        self.tree.tag_configure("even", background=LightPalette.BG_ROW_EVEN, foreground=LightPalette.FG_MAIN)
        self.tree.tag_configure("odd", background=LightPalette.BG_ROW_ODD, foreground=LightPalette.FG_MAIN)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # 4. 選択行の詳細編集フォーム
        form_frame = ttk.LabelFrame(parent, text=" 選択行の詳細編集 (ダブルクリックでも表内で直接編集できます) ", padding=(10, 6))
        form_frame.pack(fill=tk.X, padx=8, pady=6)
        self._build_detail_form(form_frame)

        # 5. 下部集計サマリー＆Excel出力フッター
        footer = tk.Frame(parent, bg=LightPalette.BG_SUB, padx=12, pady=10, bd=1, relief=tk.SOLID)
        footer.pack(fill=tk.X, padx=8, pady=(0, 8))

        self.summary_lbl = tk.Label(
            footer,
            text="総部品行数: 0 件   |   総個数: 0   |   総重量: 0.000 kg   |   合計金額: ¥0",
            bg=LightPalette.BG_SUB,
            fg=LightPalette.FG_MAIN,
            font=("Meiryo UI", 10, "bold"),
        )
        self.summary_lbl.pack(side=tk.LEFT)

        # 出力ボタン群
        tk.Button(
            footer,
            text="Excel部品表を出力 (.xlsx)",
            command=self.action_export_excel,
            bg=LightPalette.BTN_SUCCESS_BG,
            fg=LightPalette.BTN_SUCCESS_FG,
            activebackground=LightPalette.BTN_SUCCESS_ACTIVE,
            activeforeground=LightPalette.BTN_SUCCESS_FG,
            relief=tk.FLAT,
            padx=20,
            pady=7,
            font=("Meiryo UI", 10, "bold"),
            cursor="hand2",
        ).pack(side=tk.RIGHT, padx=(6, 0))

        tk.Button(
            footer,
            text="Excel出力設定...",
            command=self.action_open_excel_settings,
            bg=LightPalette.BTN_DEFAULT_BG,
            fg=LightPalette.FG_ACCENT,
            relief=tk.SOLID,
            bd=1,
            padx=12,
            pady=6,
            font=("Meiryo UI", 9, "bold"),
            cursor="hand2",
        ).pack(side=tk.RIGHT, padx=4)

        tk.Button(
            footer,
            text="CSV出力 (.csv)",
            command=self.action_export_csv,
            bg=LightPalette.BTN_DEFAULT_BG,
            fg=LightPalette.FG_MAIN,
            relief=tk.SOLID,
            bd=1,
            padx=12,
            pady=6,
            font=("Meiryo UI", 9),
            cursor="hand2",
        ).pack(side=tk.RIGHT, padx=4)

    def _create_tool_btn(
        self, parent: tk.Misc, text: str, command: Any, fg: str = LightPalette.FG_MAIN, bold: bool = False
    ) -> tk.Button:
        font_spec = ("Meiryo UI", 8, "bold") if bold else ("Meiryo UI", 8)
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=LightPalette.BTN_DEFAULT_BG,
            fg=fg,
            activebackground=LightPalette.BTN_DEFAULT_ACTIVE,
            relief=tk.SOLID,
            bd=1,
            padx=7,
            pady=3,
            font=font_spec,
            cursor="hand2",
        )
        btn.pack(side=tk.LEFT, padx=2)
        return btn

    def _build_detail_form(self, parent: ttk.LabelFrame) -> None:
        self.form_vars: Dict[str, tk.StringVar] = {
            k: tk.StringVar(value="")
            for k in [
                "item_no", "part_number", "part_name", "specification",
                "material", "treatment", "quantity", "unit",
                "unit_weight", "category", "manufacturer", "unit_price", "remarks"
            ]
        }

        # 1段目
        r0 = tk.Frame(parent, bg=LightPalette.BG_PANEL)
        r0.pack(fill=tk.X, pady=2)

        self._add_form_field(r0, "No:", "item_no", width=5)
        self._add_form_field(r0, "図番/品番:", "part_number", width=16)
        self._add_form_field(r0, "品名:", "part_name", width=22)
        self._add_form_field(r0, "型式・寸法:", "specification", width=20)
        self._add_form_combo(r0, "材質:", "material", list(MATERIAL_MASTER.keys()), width=11)
        self._add_form_combo(r0, "表面処理:", "treatment", SURFACE_TREATMENTS, width=16)

        # 2段目
        r1 = tk.Frame(parent, bg=LightPalette.BG_PANEL)
        r1.pack(fill=tk.X, pady=3)

        self._add_form_field(r1, "数量:", "quantity", width=7)
        self._add_form_combo(r1, "単位:", "unit", UNITS, width=6)
        self._add_form_field(r1, "単重(kg):", "unit_weight", width=9)
        self._add_form_combo(r1, "区分:", "category", CATEGORIES, width=10)
        self._add_form_field(r1, "メーカー:", "manufacturer", width=13)
        self._add_form_field(r1, "単価(円):", "unit_price", width=9)
        self._add_form_field(r1, "備考:", "remarks", width=16)

        tk.Button(
            r1,
            text="選択行を更新 (Enter)",
            command=self.action_apply_detail_form,
            bg=LightPalette.BTN_PRIMARY_BG,
            fg=LightPalette.BTN_PRIMARY_FG,
            activebackground=LightPalette.BTN_PRIMARY_ACTIVE,
            activeforeground=LightPalette.BTN_PRIMARY_FG,
            relief=tk.FLAT,
            padx=12,
            pady=3,
            font=("Meiryo UI", 8, "bold"),
            cursor="hand2",
        ).pack(side=tk.RIGHT, padx=4)

    def _add_form_field(self, row_frame: tk.Frame, label_txt: str, key: str, width: int) -> None:
        tk.Label(row_frame, text=label_txt, bg=LightPalette.BG_PANEL, fg=LightPalette.FG_SUB, font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT, padx=(4, 2))
        ent = ttk.Entry(row_frame, textvariable=self.form_vars[key], width=width)
        ent.pack(side=tk.LEFT, padx=(0, 6))
        ent.bind("<Return>", lambda _e: self.action_apply_detail_form())

    def _add_form_combo(self, row_frame: tk.Frame, label_txt: str, key: str, values: List[str], width: int) -> None:
        tk.Label(row_frame, text=label_txt, bg=LightPalette.BG_PANEL, fg=LightPalette.FG_SUB, font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT, padx=(4, 2))
        cb = ttk.Combobox(row_frame, textvariable=self.form_vars[key], values=values, width=width)
        cb.pack(side=tk.LEFT, padx=(0, 6))
        cb.bind("<Return>", lambda _e: self.action_apply_detail_form())

    def _build_statusbar(self) -> None:
        status_frame = tk.Frame(self.root, bg="#E2E8F0", padx=12, pady=4)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_var = tk.StringVar(
            value="準備完了: 「図面ファイルを開く (PDF/DXF)」または「サンプル図面で試す」をクリックしてください。"
        )
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg="#E2E8F0",
            fg=LightPalette.FG_MAIN,
            font=("Meiryo UI", 8),
            anchor="w",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            status_frame,
            text=f"{APP_NAME} (Windows 11 Light Mode)",
            bg="#E2E8F0",
            fg=LightPalette.FG_MUTED,
            font=("Meiryo UI", 8),
        ).pack(side=tk.RIGHT)

    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-o>", lambda _e: self.action_open_files())
        self.root.bind("<Control-s>", lambda _e: self.action_export_excel())
        self.root.bind("<Control-z>", lambda _e: self.action_undo())
        self.root.bind("<Control-y>", lambda _e: self.action_redo())
        self.root.bind("<Delete>", self._on_delete_key)

    def _on_delete_key(self, event: tk.Event) -> None:
        if isinstance(event.widget, ttk.Treeview):
            self.action_delete_rows()

    # --------------------------------------------------------------------------
    # Undo / Redo 履歴管理
    # --------------------------------------------------------------------------
    def _push_undo_state(self) -> None:
        self.undo_stack.append(copy.deepcopy(self.parts))
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def action_undo(self) -> None:
        if not self.undo_stack:
            self.set_status("これ以上元に戻す履歴はありません。")
            return
        self.redo_stack.append(copy.deepcopy(self.parts))
        self.parts = self.undo_stack.pop()
        self._refresh_treeview()
        self.set_status("直前の操作を元に戻しました (Undo)。")

    def action_redo(self) -> None:
        if not self.redo_stack:
            self.set_status("やり直す履歴はありません。")
            return
        self.undo_stack.append(copy.deepcopy(self.parts))
        self.parts = self.redo_stack.pop()
        self._refresh_treeview()
        self.set_status("操作をやり直しました (Redo)。")

    def set_status(self, message: str) -> None:
        self.status_var.set(message)
        self.root.update_idletasks()

    # --------------------------------------------------------------------------
    # ファイル読込・サンプル図面生成アクション
    # --------------------------------------------------------------------------
    def action_open_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="2D図面ファイル (PDF / DXF) を選択",
            filetypes=[
                ("2D図面ファイル (*.pdf;*.dxf)", "*.pdf *.dxf"),
                ("PDF図面ファイル (*.pdf)", "*.pdf"),
                ("CAD DXF図面ファイル (*.dxf)", "*.dxf"),
                ("すべてのファイル (*.*)", "*.*"),
            ],
        )
        if paths:
            self._load_drawing_files(list(paths))

    def action_open_folder(self) -> None:
        folder = filedialog.askdirectory(title="図面ファイルが入ったフォルダを選択")
        if not folder:
            return
        found_files: List[str] = []
        for entry in sorted(os.listdir(folder)):
            ext = os.path.splitext(entry)[1].lower()
            if ext in (".pdf", ".dxf"):
                found_files.append(os.path.join(folder, entry))
        if not found_files:
            messagebox.showinfo("フォルダ読込", "指定されたフォルダ内に PDF (.pdf) または DXF (.dxf) ファイルが見つかりませんでした。")
            return
        self._load_drawing_files(found_files)

    def action_load_sample_drawings(self) -> None:
        """サンプル2D図面(PDF組立図 & DXFブラケット図)を自動作成して読み込む"""
        self.set_status("サンプル2D図面 (PDF / DXF) を生成しています...")
        try:
            sample_dir = os.path.join(tempfile.gettempdir(), "PartLens_Sample_Drawings")
            os.makedirs(sample_dir, exist_ok=True)
            pdf_path = os.path.join(sample_dir, "Sample_Conveyor_Drive_Unit_CV1001.pdf")
            dxf_path = os.path.join(sample_dir, "Sample_Clamp_Bracket_BRK200.dxf")

            SampleDrawingGenerator.create_sample_pdf(pdf_path)
            SampleDrawingGenerator.create_sample_dxf(dxf_path)

            self._load_drawing_files([pdf_path, dxf_path])
            self.set_status("サンプル図面 (PDF 1件, DXF 1件) を生成・解析し、部品表を自動作成しました！")
        except Exception as e:
            messagebox.showerror("サンプル生成エラー", f"サンプル図面の作成中にエラーが発生しました:\n{e}")

    def _load_drawing_files(self, file_paths: List[str]) -> None:
        self._push_undo_state()
        added_docs = 0
        added_parts = 0
        errors: List[str] = []

        for path in file_paths:
            self.set_status(f"図面を解析中: {os.path.basename(path)} ...")
            try:
                doc_model, extracted = DrawingParserEngine.parse_file(path)
                self.documents.append(doc_model)
                self.parts.extend(extracted)
                added_docs += 1
                added_parts += len(extracted)

                # 最初の図面から案件名・図番が取れたらヘッダーに自動セット
                if doc_model.drawing_title and self.proj_name_var.get() in ("", "新規機械設計プロジェクト"):
                    self.proj_name_var.set(doc_model.drawing_title)
                if doc_model.drawing_number and self.dwg_no_var.get() in ("", "ASM-0001"):
                    self.dwg_no_var.set(doc_model.drawing_number)
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {e}")

        self._refresh_document_listbox()
        if self.documents and self.current_doc_idx == -1:
            self.current_doc_idx = len(self.documents) - added_docs
            self.current_page_idx = 0
            self.doc_listbox.selection_clear(0, tk.END)
            self.doc_listbox.selection_set(self.current_doc_idx)

        self.action_renumber_rows(push_undo=False)
        self._fit_canvas_view()

        if errors:
            messagebox.showwarning("一部ファイルの読み込みエラー", "\n".join(errors))

        self.set_status(f"図面 {added_docs} ファイルを読込完了。合計 {added_parts} 件の部品データを自動抽出しました。")

    def action_clear_documents(self) -> None:
        if not self.documents:
            return
        if messagebox.askyesno("確認", "読み込み済みの図面リストをクリアしますか？\n（右側の部品表データは保持されます）"):
            self.documents.clear()
            self.current_doc_idx = -1
            self.current_page_idx = 0
            self.highlight_bbox = None
            self._refresh_document_listbox()
            self._redraw_canvas()
            self.set_status("図面リストをクリアしました。")

    def action_reparse_current_page(self) -> None:
        page_data = self._get_current_page()
        if not page_data or self.current_doc_idx < 0:
            messagebox.showinfo("再解析", "先に図面ファイルを読み込んでください。")
            return
        doc = self.documents[self.current_doc_idx]
        self._push_undo_state()
        new_parts = DrawingParserEngine.extract_parts_from_tokens(
            page_data.tokens,
            doc.file_name,
            page_data.page_no,
            (0.0, 0.0, page_data.width, page_data.height),
            force_all_rows=False,
        )
        if not new_parts:
            messagebox.showinfo("再解析", "追加で検出された部品行はありませんでした。\n範囲選択モードで直接部品欄を囲んで抽出することも可能です。")
            return
        self.parts.extend(new_parts)
        self.action_renumber_rows(push_undo=False)
        self.set_status(f"{doc.file_name} (Page {page_data.page_no}) から {len(new_parts)} 件の部品を追加抽出しました。")

    # --------------------------------------------------------------------------
    # 図面プレビュー＆インタラクティブ矩形範囲選択抽出
    # --------------------------------------------------------------------------
    def _refresh_document_listbox(self) -> None:
        self.doc_listbox.delete(0, tk.END)
        for idx, doc in enumerate(self.documents, start=1):
            label = f"[{doc.file_type}] {doc.file_name}  ({len(doc.pages)}P / 抽出:{doc.extracted_count}件)"
            if doc.drawing_title:
                label += f" - {doc.drawing_title}"
            self.doc_listbox.insert(tk.END, label)

    def _on_select_document(self, _event: tk.Event) -> None:
        sel = self.doc_listbox.curselection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.documents):
            self.current_doc_idx = idx
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
            new_idx = self.current_page_idx + delta
            if 0 <= new_idx < len(doc.pages):
                self.current_page_idx = new_idx
                self.highlight_bbox = None
                self._fit_canvas_view()

    def _zoom_by(self, factor: float) -> None:
        self.zoom_scale = min(max(self.zoom_scale * factor, 0.15), 5.0)
        self._redraw_canvas()

    def _fit_canvas_view(self) -> None:
        page_data = self._get_current_page()
        if not page_data or not page_data.preview_image:
            self._redraw_canvas()
            return
        cw = max(self.preview_canvas.winfo_width(), 400)
        ch = max(self.preview_canvas.winfo_height(), 300)
        iw, ih = page_data.preview_image.size
        scale_x = (cw - 24) / max(iw, 1)
        scale_y = (ch - 24) / max(ih, 1)
        self.zoom_scale = min(scale_x, scale_y, 2.0)
        self.canvas_offset_x = (cw - iw * self.zoom_scale) * 0.5
        self.canvas_offset_y = (ch - ih * self.zoom_scale) * 0.5
        self._redraw_canvas()

    def _toggle_interaction_mode(self) -> None:
        if self.interaction_mode == "pan":
            self.interaction_mode = "select_area"
            self.btn_mode_select.configure(
                text="範囲選択して部品抽出: ON (ドラッグで囲む)",
                bg=LightPalette.BTN_WARN_BG,
                fg=LightPalette.BTN_WARN_FG,
            )
            self.preview_canvas.configure(cursor="crosshair")
            self.set_status("範囲選択抽出モードON: プレビュー図面上の部品欄やテキストをマウス左ドラッグで四角く囲んでください。")
        else:
            self.interaction_mode = "pan"
            self.btn_mode_select.configure(
                text="範囲選択して部品抽出: OFF",
                bg=LightPalette.BTN_DEFAULT_BG,
                fg=LightPalette.FG_ACCENT,
            )
            self.preview_canvas.configure(cursor="fleur")
            self.set_status("移動(パン)モード: プレビュー図面をドラッグして移動できます。")

    def _redraw_canvas(self) -> None:
        self.preview_canvas.delete("all")
        page_data = self._get_current_page()

        if not page_data or not page_data.preview_image:
            self.page_info_lbl.configure(text="Page: 0 / 0")
            self.zoom_lbl.configure(text="100%")
            cw = max(self.preview_canvas.winfo_width(), 400)
            ch = max(self.preview_canvas.winfo_height(), 300)
            self.preview_canvas.create_text(
                cw // 2,
                ch // 2,
                text="ここに2D図面 (PDF / DXF) のプレビューが表示されます\n\n上部の「図面ファイルを開く」または「サンプル図面で試す」をクリックしてください",
                fill=LightPalette.FG_MUTED,
                font=("Meiryo UI", 10, "bold"),
                justify=tk.CENTER,
            )
            return

        doc = self.documents[self.current_doc_idx]
        self.page_info_lbl.configure(text=f"Page: {self.current_page_idx + 1} / {len(doc.pages)}")
        self.zoom_lbl.configure(text=f"{int(self.zoom_scale * 100)}%")

        iw, ih = page_data.preview_image.size
        disp_w = max(int(iw * self.zoom_scale), 10)
        disp_h = max(int(ih * self.zoom_scale), 10)

        resized = page_data.preview_image.resize((disp_w, disp_h), Image.Resampling.BILINEAR)
        self._tk_preview_img = ImageTk.PhotoImage(resized)

        ox = int(self.canvas_offset_x)
        oy = int(self.canvas_offset_y)
        self.preview_canvas.create_image(ox, oy, anchor=tk.NW, image=self._tk_preview_img)
        self.preview_canvas.create_rectangle(ox, oy, ox + disp_w, oy + disp_h, outline="#94A3B8", width=1)

        # 選択部品行の座標ハイライト枠を描画
        if self.highlight_bbox and any(self.highlight_bbox):
            bx0, by0, bx1, by1 = self.highlight_bbox
            if page_data.width > 0 and page_data.height > 0:
                sx0 = ox + (bx0 / page_data.width) * disp_w
                sy0 = oy + (by0 / page_data.height) * disp_h
                sx1 = ox + (bx1 / page_data.width) * disp_w
                sy1 = oy + (by1 / page_data.height) * disp_h
                self.preview_canvas.create_rectangle(
                    sx0 - 4, sy0 - 4, sx1 + 4, sy1 + 4,
                    outline=LightPalette.HIGHLIGHT_BOX,
                    width=3,
                )

    def _on_canvas_mouse_down(self, event: tk.Event) -> None:
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        if self.interaction_mode == "select_area":
            if self.selection_rect_id:
                self.preview_canvas.delete(self.selection_rect_id)
            self.selection_rect_id = self.preview_canvas.create_rectangle(
                event.x, event.y, event.x, event.y,
                outline=LightPalette.SELECT_BOX,
                width=2,
                dash=(5, 3),
            )

    def _on_canvas_mouse_move(self, event: tk.Event) -> None:
        if self.interaction_mode == "pan":
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y
            self.canvas_offset_x += dx
            self.canvas_offset_y += dy
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self._redraw_canvas()
        elif self.interaction_mode == "select_area" and self.selection_rect_id:
            self.preview_canvas.coords(
                self.selection_rect_id,
                self.drag_start_x,
                self.drag_start_y,
                event.x,
                event.y,
            )

    def _on_canvas_mouse_up(self, event: tk.Event) -> None:
        if self.interaction_mode != "select_area":
            return
        if self.selection_rect_id:
            self.preview_canvas.delete(self.selection_rect_id)
            self.selection_rect_id = None

        page_data = self._get_current_page()
        if not page_data or not page_data.preview_image:
            return

        x0_scr = min(self.drag_start_x, event.x)
        y0_scr = min(self.drag_start_y, event.y)
        x1_scr = max(self.drag_start_x, event.x)
        y1_scr = max(self.drag_start_y, event.y)

        if (x1_scr - x0_scr) < 12 or (y1_scr - y0_scr) < 10:
            return

        iw, ih = page_data.preview_image.size
        disp_w = max(iw * self.zoom_scale, 1.0)
        disp_h = max(ih * self.zoom_scale, 1.0)

        # スクリーン座標 -> ページ実座標へ逆変換
        rel_x0 = (x0_scr - self.canvas_offset_x) / disp_w
        rel_y0 = (y0_scr - self.canvas_offset_y) / disp_h
        rel_x1 = (x1_scr - self.canvas_offset_x) / disp_w
        rel_y1 = (y1_scr - self.canvas_offset_y) / disp_h

        px0 = max(0.0, rel_x0 * page_data.width)
        py0 = max(0.0, rel_y0 * page_data.height)
        px1 = min(page_data.width, rel_x1 * page_data.width)
        py1 = min(page_data.height, rel_y1 * page_data.height)

        doc = self.documents[self.current_doc_idx]
        extracted = DrawingParserEngine.extract_parts_from_tokens(
            page_data.tokens,
            doc.file_name,
            page_data.page_no,
            (px0, py0, px1, py1),
            force_all_rows=True,
        )

        if not extracted:
            messagebox.showinfo(
                "範囲選択抽出",
                "選択された範囲内にテキストデータが検出されませんでした。\n文字を含む箇所を広めに囲んで再度お試しください。",
            )
            return

        self._push_undo_state()
        self.parts.extend(extracted)
        doc.extracted_count += len(extracted)
        self._refresh_document_listbox()
        self.action_renumber_rows(push_undo=False)
        self.highlight_bbox = (px0, py0, px1, py1)
        self._redraw_canvas()
        self.set_status(f"選択範囲から {len(extracted)} 件の部品データを抽出し、部品表へ追加しました！")

    def _on_canvas_mouse_wheel(self, event: tk.Event) -> None:
        factor = 1.15 if event.delta > 0 else 0.85
        self._zoom_by(factor)

    # --------------------------------------------------------------------------
    # 部品表 (Treeview) の表示・インライン編集・操作アクション
    # --------------------------------------------------------------------------
    def _refresh_treeview(self) -> None:
        self._close_inline_editor()
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)

        keyword = self.filter_var.get().strip().lower()
        visible_idx = 0

        for p in self.parts:
            if keyword:
                haystack = f"{p.part_number} {p.part_name} {p.specification} {p.material} {p.treatment} {p.manufacturer} {p.remarks}".lower()
                if keyword not in haystack:
                    continue

            tag = "even" if visible_idx % 2 == 0 else "odd"
            visible_idx += 1

            values = (
                p.item_no,
                p.part_number,
                p.part_name,
                p.specification,
                p.material,
                p.treatment,
                p.formatted_qty(),
                p.unit,
                f"{p.unit_weight:.3f}" if p.unit_weight > 0 else "0.000",
                f"{p.total_weight:.3f}" if p.total_weight > 0 else "0.000",
                p.category,
                p.manufacturer,
                f"{int(p.unit_price):,}" if p.unit_price > 0 else "0",
                f"{int(p.amount):,}" if p.amount > 0 else "0",
                p.source_file,
                p.remarks,
            )
            self.tree.insert("", tk.END, iid=p.uid, values=values, tags=(tag,))

        self._update_summary_bar()

    def _update_summary_bar(self) -> None:
        total_rows = len(self.parts)
        total_qty = sum(p.quantity for p in self.parts)
        total_wt = sum(p.total_weight for p in self.parts)
        total_amt = sum(p.amount for p in self.parts)

        qty_str = str(int(total_qty)) if float(total_qty).is_integer() else f"{total_qty:.2f}"
        self.summary_lbl.configure(
            text=f"総部品行数: {total_rows} 件   |   総個数: {qty_str}   |   総重量: {total_wt:.3f} kg   |   合計金額: ¥{int(total_amt):,}"
        )

    def _find_part_by_uid(self, uid: str) -> Optional[PartItem]:
        for p in self.parts:
            if p.uid == uid:
                return p
        return None

    def _on_tree_select(self, _event: tk.Event) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        uid = sel[0]
        part = self._find_part_by_uid(uid)
        if not part:
            return

        self.selected_part_uid = uid
        self.form_vars["item_no"].set(part.item_no)
        self.form_vars["part_number"].set(part.part_number)
        self.form_vars["part_name"].set(part.part_name)
        self.form_vars["specification"].set(part.specification)
        self.form_vars["material"].set(part.material)
        self.form_vars["treatment"].set(part.treatment)
        self.form_vars["quantity"].set(part.formatted_qty())
        self.form_vars["unit"].set(part.unit)
        self.form_vars["unit_weight"].set(f"{part.unit_weight:.3f}")
        self.form_vars["category"].set(part.category)
        self.form_vars["manufacturer"].set(part.manufacturer)
        self.form_vars["unit_price"].set(str(int(part.unit_price)) if float(part.unit_price).is_integer() else str(part.unit_price))
        self.form_vars["remarks"].set(part.remarks)

        # 抽出元図面があればプレビューを切り替えて座標ハイライト表示
        if part.source_file:
            for d_idx, doc in enumerate(self.documents):
                if doc.file_name == part.source_file:
                    if self.current_doc_idx != d_idx or self.current_page_idx != max(part.source_page - 1, 0):
                        self.current_doc_idx = d_idx
                        self.current_page_idx = max(min(part.source_page - 1, len(doc.pages) - 1), 0)
                    break
        self.highlight_bbox = part.bbox
        self._redraw_canvas()

    def action_apply_detail_form(self) -> None:
        if not self.selected_part_uid:
            messagebox.showinfo("行未選択", "編集する部品行を一覧から選択してください。")
            return
        part = self._find_part_by_uid(self.selected_part_uid)
        if not part:
            return

        self._push_undo_state()
        part.item_no = self.form_vars["item_no"].get().strip() or part.item_no
        part.part_number = self.form_vars["part_number"].get().strip()
        part.part_name = self.form_vars["part_name"].get().strip()
        part.specification = self.form_vars["specification"].get().strip()
        part.material = self.form_vars["material"].get().strip()
        part.treatment = self.form_vars["treatment"].get().strip()
        part.unit = self.form_vars["unit"].get().strip() or "個"
        part.category = self.form_vars["category"].get().strip() or "加工品"
        part.manufacturer = self.form_vars["manufacturer"].get().strip()
        part.remarks = self.form_vars["remarks"].get().strip()

        try:
            part.quantity = max(float(self.form_vars["quantity"].get().replace(",", "")), 0.0)
        except ValueError:
            pass
        try:
            part.unit_weight = max(float(self.form_vars["unit_weight"].get().replace(",", "")), 0.0)
        except ValueError:
            pass
        try:
            part.unit_price = max(float(self.form_vars["unit_price"].get().replace(",", "")), 0.0)
        except ValueError:
            pass

        # 単重が0で寸法がある場合は自動推算
        if part.unit_weight == 0.0 and part.specification:
            part.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(part.specification, part.material)

        self._refresh_treeview()
        if self.tree.exists(part.uid):
            self.tree.selection_set(part.uid)
        self.set_status(f"No.{part.item_no} 「{part.part_name}」の部品情報を更新しました。")

    # --------------------------------------------------------------------------
    # Treeview ダブルクリック・セル直接インライン編集
    # --------------------------------------------------------------------------
    def _close_inline_editor(self) -> None:
        if self._inline_editor is not None:
            try:
                self._inline_editor.destroy()
            except Exception:
                pass
            self._inline_editor = None

    def _on_tree_double_click(self, event: tk.Event) -> None:
        self._close_inline_editor()
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        col_idx = int(col_id.replace("#", "")) - 1
        if col_idx < 0 or col_idx >= len(self.TREE_COLUMNS):
            return
        field_name = self.TREE_COLUMNS[col_idx]

        # 自動計算列(total_weight, amount)がダブルクリックされた場合は単重または単価列を編集対象にする
        if field_name == "total_weight":
            field_name = "unit_weight"
            col_idx = self.TREE_COLUMNS.index("unit_weight")
            col_id = f"#{col_idx + 1}"
        elif field_name == "amount":
            field_name = "unit_price"
            col_idx = self.TREE_COLUMNS.index("unit_price")
            col_id = f"#{col_idx + 1}"

        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        bx, by, bw, bh = bbox

        part = self._find_part_by_uid(row_id)
        if not part:
            return

        cur_val = str(getattr(part, field_name, ""))
        if field_name == "quantity":
            cur_val = part.formatted_qty()

        var = tk.StringVar(value=cur_val)

        if field_name == "material":
            editor: tk.Widget = ttk.Combobox(self.tree, textvariable=var, values=list(MATERIAL_MASTER.keys()))
        elif field_name == "treatment":
            editor = ttk.Combobox(self.tree, textvariable=var, values=SURFACE_TREATMENTS)
        elif field_name == "category":
            editor = ttk.Combobox(self.tree, textvariable=var, values=CATEGORIES)
        elif field_name == "unit":
            editor = ttk.Combobox(self.tree, textvariable=var, values=UNITS)
        else:
            editor = tk.Entry(
                self.tree,
                textvariable=var,
                bg="#FEF9C3",
                fg=LightPalette.FG_MAIN,
                relief=tk.SOLID,
                bd=1,
                font=("Meiryo UI", 9),
            )

        editor.place(x=bx, y=by, width=max(bw, 70), height=bh)
        editor.focus_set()
        if isinstance(editor, tk.Entry):
            editor.select_range(0, tk.END)

        self._inline_editor = editor

        def commit_edit(_e: Optional[tk.Event] = None) -> None:
            if self._inline_editor is None:
                return
            new_val = var.get().strip()
            self._close_inline_editor()
            self._push_undo_state()
            if field_name in ("quantity", "unit_weight", "unit_price"):
                try:
                    num = max(float(new_val.replace(",", "")), 0.0)
                    setattr(part, field_name, num)
                except ValueError:
                    pass
            else:
                setattr(part, field_name, new_val)

            if field_name in ("specification", "material") and part.unit_weight == 0.0 and part.specification:
                part.unit_weight = SmartPartAnalyzer.estimate_weight_from_spec(part.specification, part.material)

            self._refresh_treeview()
            if self.tree.exists(part.uid):
                self.tree.selection_set(part.uid)
                self._on_tree_select(event)

        editor.bind("<Return>", commit_edit)
        editor.bind("<FocusOut>", commit_edit)
        editor.bind("<Escape>", lambda _e: self._close_inline_editor())

    def _sort_tree_by_column(self, col_name: str) -> None:
        if not self.parts:
            return
        self._push_undo_state()

        def sort_key(p: PartItem) -> Any:
            if col_name == "item_no":
                return int(p.item_no) if p.item_no.isdigit() else 999999
            if col_name in ("quantity", "unit_weight", "total_weight", "unit_price", "amount"):
                return float(getattr(p, col_name, 0.0))
            return str(getattr(p, col_name, "")).lower()

        self.parts.sort(key=sort_key)
        self._refresh_treeview()
        self.set_status(f"列「{col_name}」で並び替えました。")

    # --------------------------------------------------------------------------
    # 行追加・削除・移動・名寄せ・重量一括計算アクション
    # --------------------------------------------------------------------------
    def action_add_row(self) -> None:
        self._push_undo_state()
        next_no = str(len(self.parts) + 1)
        src_name = ""
        if 0 <= self.current_doc_idx < len(self.documents):
            src_name = self.documents[self.current_doc_idx].file_name

        new_item = PartItem(
            item_no=next_no,
            part_number=f"P-{len(self.parts) + 1:03d}",
            part_name="新規部品",
            specification="",
            material="SS400",
            treatment="四三酸化鉄皮膜(黒染め)",
            quantity=1.0,
            unit="個",
            category="加工品",
            source_file=src_name,
        )
        self.parts.append(new_item)
        self._refresh_treeview()
        if self.tree.exists(new_item.uid):
            self.tree.selection_set(new_item.uid)
            self.tree.see(new_item.uid)
            self.selected_part_uid = new_item.uid
        self.set_status("新しい部品行を追加しました。")

    def action_duplicate_row(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("複製", "複製する行を選択してください。")
            return
        part = self._find_part_by_uid(sel[0])
        if not part:
            return
        self._push_undo_state()
        dup = copy.deepcopy(part)
        dup.uid = str(uuid.uuid4())
        idx = self.parts.index(part)
        self.parts.insert(idx + 1, dup)
        self.action_renumber_rows(push_undo=False)
        if self.tree.exists(dup.uid):
            self.tree.selection_set(dup.uid)
        self.set_status(f"「{part.part_name}」の行を複製しました。")

    def action_delete_rows(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        self._push_undo_state()
        del_uids = set(sel)
        self.parts = [p for p in self.parts if p.uid not in del_uids]
        self.selected_part_uid = None
        self._refresh_treeview()
        self.set_status(f"{len(del_uids)} 件の部品行を削除しました。")

    def action_move_row(self, direction: int) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        uid = sel[0]
        part = self._find_part_by_uid(uid)
        if not part:
            return
        idx = self.parts.index(part)
        new_idx = idx + direction
        if 0 <= new_idx < len(self.parts):
            self._push_undo_state()
            self.parts[idx], self.parts[new_idx] = self.parts[new_idx], self.parts[idx]
            self._refresh_treeview()
            self.tree.selection_set(uid)
            self.tree.see(uid)

    def action_renumber_rows(self, push_undo: bool = True) -> None:
        if not self.parts:
            self._refresh_treeview()
            return
        if push_undo:
            self._push_undo_state()
        for idx, p in enumerate(self.parts, start=1):
            p.item_no = str(idx)
        self._refresh_treeview()
        if push_undo:
            self.set_status("全行の項番(No.)を1から連番に振り直しました。")

    def action_merge_duplicates(self) -> None:
        if len(self.parts) < 2:
            return
        before_count = len(self.parts)
        self._push_undo_state()
        self.parts = ExcelBOMExporter.merge_duplicate_parts(self.parts)
        after_count = len(self.parts)
        self._refresh_treeview()
        self.set_status(f"重複部品の名寄せ統合を実行しました ({before_count} 行 → {after_count} 行)。")

    def action_auto_calc_weights(self) -> None:
        if not self.parts:
            return
        self._push_undo_state()
        updated = 0
        for p in self.parts:
            if p.specification:
                w = SmartPartAnalyzer.estimate_weight_from_spec(p.specification, p.material)
                if w > 0:
                    p.unit_weight = w
                    updated += 1
        self._refresh_treeview()
        self.set_status(f"型式・寸法と材質から {updated} 件の部品の単重(kg)を自動計算しました。")

    def action_open_weight_calculator(self) -> None:
        init_mat = "SS400"
        if self.selected_part_uid:
            part = self._find_part_by_uid(self.selected_part_uid)
            if part and part.material:
                init_mat = part.material

        def on_apply(weight_kg: float, spec_str: str, mat_str: str) -> None:
            if self.selected_part_uid:
                part = self._find_part_by_uid(self.selected_part_uid)
                if part:
                    self._push_undo_state()
                    part.unit_weight = weight_kg
                    if not part.specification:
                        part.specification = spec_str
                    if not part.material:
                        part.material = mat_str
                    self._refresh_treeview()
                    self.tree.selection_set(part.uid)
                    self.set_status(f"計算結果 ({weight_kg:.3f} kg) を選択中の部品行へ反映しました。")
            else:
                self.form_vars["unit_weight"].set(f"{weight_kg:.3f}")
                self.form_vars["specification"].set(spec_str)
                self.form_vars["material"].set(mat_str)

        WeightCalculatorDialog(self.root, init_mat, on_apply)

    def action_clear_all_parts(self) -> None:
        if not self.parts:
            return
        if messagebox.askyesno("全クリア確認", "現在の部品表データをすべてクリアしますか？\n（元に戻すボタンで復元可能です）"):
            self._push_undo_state()
            self.parts.clear()
            self.selected_part_uid = None
            self._refresh_treeview()
            self.set_status("部品表をクリアしました。")

    # --------------------------------------------------------------------------
    # Excel (.xlsx) / CSV 出力・プロジェクト保存/読込アクション
    # --------------------------------------------------------------------------
    def action_open_excel_settings(self) -> None:
        self.excel_config.project_name = self.proj_name_var.get().strip()
        self.excel_config.drawing_no = self.dwg_no_var.get().strip()
        self.excel_config.author_name = self.author_var.get().strip()
        ExcelSettingsDialog(
            self.root,
            self.excel_config,
            lambda: self.set_status("Excel出力設定を更新しました。"),
        )

    def action_export_excel(self) -> None:
        if not self.parts:
            messagebox.showwarning("部品データなし", "出力する部品データがありません。\n先に図面を読み込むか行を追加してください。")
            return

        self.excel_config.project_name = self.proj_name_var.get().strip() or "機械設計部品表"
        self.excel_config.drawing_no = self.dwg_no_var.get().strip()
        self.excel_config.author_name = self.author_var.get().strip()

        default_name = f"部品表_{self.excel_config.drawing_no or 'BOM'}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
        default_name = re.sub(r"[\\/*?:\"<>|]", "_", default_name)

        out_path = filedialog.asksaveasfilename(
            title="Excel部品表 (.xlsx) の保存先を指定",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel ブック (*.xlsx)", "*.xlsx")],
        )
        if not out_path:
            return

        self.set_status(f"Excel部品表を作成中: {os.path.basename(out_path)} ...")
        try:
            ExcelBOMExporter.export_to_excel(out_path, self.parts, self.documents, self.excel_config)
            self.set_status(f"Excel部品表を保存しました: {out_path}")
            if messagebox.askyesno("Excel出力完了", f"Excel部品表の出力が完了しました！\n\n保存先:\n{out_path}\n\n今すぐこのExcelファイルを開きますか？"):
                try:
                    os.startfile(out_path)
                except Exception:
                    pass
        except PermissionError:
            messagebox.showerror("保存エラー", "出力先のExcelファイルが別のプログラムで開かれています。\nExcelを閉じてから再度出力してください。")
        except Exception as e:
            messagebox.showerror("Excel出力エラー", f"Excelファイルの出力中にエラーが発生しました:\n{e}\n{traceback.format_exc()}")

    def action_export_csv(self) -> None:
        if not self.parts:
            messagebox.showwarning("部品データなし", "出力する部品データがありません。")
            return

        default_name = f"部品表_{self.dwg_no_var.get().strip() or 'BOM'}.csv"
        out_path = filedialog.asksaveasfilename(
            title="CSV部品表 (.csv) の保存先を指定",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSVファイル (Excel対応 UTF-8 BOM付) (*.csv)", "*.csv")],
        )
        if not out_path:
            return

        try:
            with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "No.", "図番/品番", "品名", "型式・寸法・仕様", "材質", "表面処理",
                    "数量", "単位", "単重(kg)", "総重量(kg)", "調達区分", "メーカー",
                    "単価(円)", "金額(円)", "抽出元図面", "備考"
                ])
                for p in self.parts:
                    writer.writerow([
                        p.item_no, p.part_number, p.part_name, p.specification,
                        p.material, p.treatment, p.formatted_qty(), p.unit,
                        p.unit_weight, p.total_weight, p.category, p.manufacturer,
                        p.unit_price, p.amount, p.source_file, p.remarks
                    ])
            self.set_status(f"CSVファイルを保存しました: {out_path}")
            messagebox.showinfo("CSV出力完了", f"CSVファイルを出力しました:\n{out_path}")
        except Exception as e:
            messagebox.showerror("CSV出力エラー", f"CSV保存中にエラーが発生しました:\n{e}")

    def action_save_project(self) -> None:
        out_path = filedialog.asksaveasfilename(
            title="プロジェクトファイルの保存",
            defaultextension=".plens.json",
            initialfile=f"{self.dwg_no_var.get().strip() or 'Project'}.plens.json",
            filetypes=[("PartLens プロジェクト (*.plens.json;*.json)", "*.plens.json *.json")],
        )
        if not out_path:
            return

        data = {
            "app_name": APP_NAME,
            "saved_at": datetime.datetime.now().isoformat(),
            "project_name": self.proj_name_var.get(),
            "drawing_no": self.dwg_no_var.get(),
            "author_name": self.author_var.get(),
            "parts": [p.to_dict() for p in self.parts],
            "source_files": [d.file_path for d in self.documents if os.path.exists(d.file_path)],
        }
        try:
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.set_status(f"プロジェクトを保存しました: {out_path}")
            messagebox.showinfo("保存完了", f"作業状態を保存しました:\n{out_path}")
        except Exception as e:
            messagebox.showerror("保存エラー", f"プロジェクト保存中にエラーが発生しました:\n{e}")

    def action_load_project(self) -> None:
        in_path = filedialog.askopenfilename(
            title="プロジェクトファイルを開く",
            filetypes=[("PartLens プロジェクト (*.plens.json;*.json)", "*.plens.json *.json")],
        )
        if not in_path:
            return

        try:
            with open(in_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            self._push_undo_state()
            self.proj_name_var.set(str(data.get("project_name", "")))
            self.dwg_no_var.set(str(data.get("drawing_no", "")))
            self.author_var.set(str(data.get("author_name", "")))

            loaded_parts = [PartItem.from_dict(d) for d in data.get("parts", [])]
            self.parts = loaded_parts

            # 元図面ファイルが存在すればプレビュー用に再読込
            self.documents.clear()
            for fpath in data.get("source_files", []):
                if os.path.exists(fpath):
                    try:
                        doc_model, _ = DrawingParserEngine.parse_file(fpath)
                        self.documents.append(doc_model)
                    except Exception:
                        pass

            self._refresh_document_listbox()
            if self.documents:
                self.current_doc_idx = 0
                self.current_page_idx = 0
                self._fit_canvas_view()
            self._refresh_treeview()
            self.set_status(f"プロジェクトを読み込みました ({len(self.parts)} 件の部品データ)。")
        except Exception as e:
            messagebox.showerror("読込エラー", f"プロジェクトの読み込みに失敗しました:\n{e}")

    def action_show_readme(self) -> None:
        ReadmeViewerDialog(self.root)


# ==============================================================================
# エントリーポイント
# ==============================================================================

def main() -> None:
    enable_high_dpi_awareness()
    root = tk.Tk()
    app = PartLensApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
