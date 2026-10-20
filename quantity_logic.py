"""建具数量集計ロジック Ver.2-A

OCR座標付き検出を、ページ種別・建具マスター・重複除去・除外領域に基づき
階別／ページ別に集計するための独立モジュール。
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from enum import Enum
from math import hypot
from typing import Iterable, Optional
import re


class DrawingType(str, Enum):
    DOOR_WINDOW_SCHEDULE = "建具表"
    FLOOR_PLAN = "平面図"
    KEY_PLAN = "建具キープラン"
    ELEVATION = "立面図"
    SECTION = "断面図"
    DETAIL = "詳細図"
    OTHER = "その他"
    UNKNOWN = "不明"


class CountPolicy(str, Enum):
    COUNT = "数量集計に使用"
    REFERENCE = "照合のみ"
    EXCLUDE = "集計対象外"
    REVIEW = "要確認"


class QuantityStatus(str, Enum):
    CANDIDATE = "数量確認候補"
    REVIEW = "要目視確認"
    UNCONFIRMED = "数量未確定"
    CONFIRMED = "確認済み"


@dataclass(frozen=True)
class BoundingBox:
    x: float
    y: float
    width: float
    height: float

    @property
    def cx(self) -> float:
        return self.x + self.width / 2

    @property
    def cy(self) -> float:
        return self.y + self.height / 2

    @property
    def area(self) -> float:
        return max(self.width, 0) * max(self.height, 0)

    def iou(self, other: "BoundingBox") -> float:
        x1, y1 = max(self.x, other.x), max(self.y, other.y)
        x2 = min(self.x + self.width, other.x + other.width)
        y2 = min(self.y + self.height, other.y + other.height)
        inter = max(0, x2 - x1) * max(0, y2 - y1)
        union = self.area + other.area - inter
        return inter / union if union else 0.0


@dataclass
class PageMetadata:
    file_name: str
    page: int
    width: int
    height: int
    title_text: str = ""
    drawing_no: str = ""
    floor: str = ""
    revision: str = ""
    drawing_type: DrawingType = DrawingType.UNKNOWN
    drawing_type_confidence: float = 0.0
    count_policy: CountPolicy = CountPolicy.REVIEW
    is_latest_revision: Optional[bool] = None


@dataclass
class JoineryMaster:
    symbol: str
    name: str = ""
    width: Optional[int] = None
    height: Optional[int] = None
    opening_type: str = ""
    source_page: Optional[int] = None


@dataclass
class SymbolDetection:
    detection_id: str
    raw_text: str
    normalized_symbol: str
    file_name: str
    page: int
    bbox: BoundingBox
    ocr_confidence: float
    drawing_type: DrawingType
    floor: str = ""
    drawing_no: str = ""
    matched_master_symbol: str = ""
    match_distance: Optional[int] = None
    auto_corrected: bool = False
    excluded: bool = False
    exclusion_reason: str = ""
    duplicate_of: str = ""
    manually_confirmed: bool = False

    @property
    def effective_symbol(self) -> str:
        return self.matched_master_symbol or self.normalized_symbol


@dataclass
class QuantityRecord:
    symbol: str
    floor: str
    file_name: str
    drawing_no: str
    page: int
    count: int
    detection_ids: list[str] = field(default_factory=list)
    mean_confidence: float = 0.0
    status: QuantityStatus = QuantityStatus.UNCONFIRMED
    reasons: list[str] = field(default_factory=list)


@dataclass
class QuantitySummary:
    symbol: str
    floor_counts: dict[str, int]
    total: int
    status: QuantityStatus
    reasons: list[str]
    source_records: list[QuantityRecord]


DRAWING_RULES = [
    (DrawingType.DOOR_WINDOW_SCHEDULE, ["建具表", "アルミ建具表", "鋼製建具表", "建具リスト"]),
    (DrawingType.KEY_PLAN, ["建具キープラン", "建具配置図"]),
    (DrawingType.FLOOR_PLAN, ["平面図", "基準階平面図", "平面詳細図"]),
    (DrawingType.ELEVATION, ["立面図", "東立面", "西立面", "南立面", "北立面"]),
    (DrawingType.SECTION, ["断面図", "矩計図"]),
    (DrawingType.DETAIL, ["建具詳細図", "開口部詳細図", "サッシ詳細図", "納まり図"]),
]


def normalize_symbol(text: str) -> str:
    text = text.upper().translate(str.maketrans({"Ｏ": "0", "Ｉ": "1", "Ｌ": "1"}))
    text = re.sub(r"\s+", "", text)
    text = text.replace("－", "-").replace("_", "-")
    # 接頭辞の後ろにある数字部分だけ、OCRで混同しやすい O/I/L を補正
    m = re.fullmatch(r"([A-Z]{1,3})-?([0-9OIL]{1,3})", text)
    if m:
        suffix = m.group(2).translate(str.maketrans({"O": "0", "I": "1", "L": "1"}))
        return f"{m.group(1)}-{suffix}"
    text = re.sub(r"([A-Z]{1,3})(\d{1,3})$", r"\1-\2", text)
    return text


def classify_drawing(title_text: str) -> tuple[DrawingType, float]:
    compact = re.sub(r"\s+", "", title_text)
    for dtype, words in DRAWING_RULES:
        matches = [w for w in words if w in compact]
        if matches:
            exact_bonus = 0.15 if compact in words else 0.0
            return dtype, min(1.0, 0.8 + exact_bonus)
    return DrawingType.UNKNOWN, 0.0


def default_count_policy(drawing_type: DrawingType) -> CountPolicy:
    if drawing_type == DrawingType.FLOOR_PLAN:
        return CountPolicy.COUNT
    if drawing_type == DrawingType.KEY_PLAN:
        return CountPolicy.REVIEW
    if drawing_type in {DrawingType.DOOR_WINDOW_SCHEDULE, DrawingType.ELEVATION,
                        DrawingType.SECTION, DrawingType.DETAIL}:
        return CountPolicy.REFERENCE
    return CountPolicy.EXCLUDE


def levenshtein(a: str, b: str) -> int:
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(cur[-1] + 1, prev[j] + 1, prev[j-1] + (ca != cb)))
        prev = cur
    return prev[-1]


def match_to_master(raw_symbol: str, masters: Iterable[JoineryMaster], max_distance: int = 1):
    symbol = normalize_symbol(raw_symbol)
    candidates = []
    for master in masters:
        target = normalize_symbol(master.symbol)
        d = levenshtein(symbol, target)
        if d <= max_distance:
            candidates.append((d, target))
    candidates.sort()
    if not candidates:
        return "", None, False
    best_distance = candidates[0][0]
    best = [c for c in candidates if c[0] == best_distance]
    if len(best) != 1:
        return "", best_distance, False
    return best[0][1], best_distance, best_distance > 0


def inside_excluded_region(det: SymbolDetection, page: PageMetadata,
                           right_margin_ratio: float = 0.15,
                           bottom_margin_ratio: float = 0.10) -> Optional[str]:
    # 初期ルール。画面側でページごとに上書き可能にする。
    if det.bbox.cx >= page.width * (1 - right_margin_ratio):
        return "右端の表題欄・凡例候補"
    if det.bbox.cy >= page.height * (1 - bottom_margin_ratio):
        return "下端の図面タイトル・改訂欄候補"
    return None


def are_near_duplicates(a: SymbolDetection, b: SymbolDetection,
                        page_width: int, page_height: int,
                        x_ratio: float = 0.01, y_ratio: float = 0.007,
                        iou_threshold: float = 0.50) -> bool:
    if a.page != b.page or a.file_name != b.file_name:
        return False
    if a.effective_symbol != b.effective_symbol:
        return False
    if a.bbox.iou(b.bbox) >= iou_threshold:
        return True
    return abs(a.bbox.cx - b.bbox.cx) <= page_width * x_ratio and \
           abs(a.bbox.cy - b.bbox.cy) <= page_height * y_ratio


def prepare_detections(detections: list[SymbolDetection], pages: dict[tuple[str, int], PageMetadata],
                       masters: list[JoineryMaster], min_confidence: float = 45.0) -> list[SymbolDetection]:
    """マスター照合、除外領域判定、近接重複除去を行う。入力オブジェクトを更新する。"""
    for det in detections:
        det.normalized_symbol = normalize_symbol(det.raw_text)
        matched, distance, corrected = match_to_master(det.normalized_symbol, masters)
        det.matched_master_symbol = matched
        det.match_distance = distance
        det.auto_corrected = corrected
        page = pages[(det.file_name, det.page)]
        if page.count_policy != CountPolicy.COUNT:
            det.excluded = True
            det.exclusion_reason = f"{page.drawing_type.value}: {page.count_policy.value}"
        elif page.is_latest_revision is False:
            det.excluded = True
            det.exclusion_reason = "旧版図面"
        elif det.ocr_confidence < min_confidence:
            det.excluded = True
            det.exclusion_reason = "OCR信頼度不足"
        elif not matched:
            det.excluded = True
            det.exclusion_reason = "建具マスターに一致なし、または候補複数"
        else:
            region_reason = inside_excluded_region(det, page)
            if region_reason:
                det.excluded = True
                det.exclusion_reason = region_reason

    # 信頼度が高いものを代表として重複除去
    for key in sorted(pages):
        page = pages[key]
        page_dets = sorted(
            [d for d in detections if (d.file_name, d.page) == key and not d.excluded],
            key=lambda d: d.ocr_confidence, reverse=True
        )
        kept = []
        for det in page_dets:
            duplicate = next((k for k in kept if are_near_duplicates(det, k, page.width, page.height)), None)
            if duplicate:
                det.excluded = True
                det.exclusion_reason = "近接重複"
                det.duplicate_of = duplicate.detection_id
            else:
                kept.append(det)
    return detections


def aggregate_quantities(detections: list[SymbolDetection], pages: dict[tuple[str, int], PageMetadata],
                         masters: list[JoineryMaster]) -> tuple[list[QuantityRecord], list[QuantitySummary]]:
    master_symbols = {normalize_symbol(m.symbol) for m in masters}
    grouped: dict[tuple[str, str, str, str, int], list[SymbolDetection]] = {}
    for det in detections:
        if det.excluded:
            continue
        symbol = det.effective_symbol
        if symbol not in master_symbols:
            continue
        page = pages[(det.file_name, det.page)]
        floor = det.floor or page.floor or "階不明"
        key = (symbol, floor, det.file_name, det.drawing_no or page.drawing_no, det.page)
        grouped.setdefault(key, []).append(det)

    records = []
    for (symbol, floor, file_name, drawing_no, page_no), matches in grouped.items():
        mean_conf = sum(d.ocr_confidence for d in matches) / len(matches)
        reasons = []
        status = QuantityStatus.CANDIDATE
        if floor == "階不明":
            status = QuantityStatus.REVIEW; reasons.append("階情報不明")
        if any(d.auto_corrected for d in matches):
            status = QuantityStatus.REVIEW; reasons.append("OCR符号を自動補正")
        if mean_conf < 70:
            status = QuantityStatus.REVIEW; reasons.append("平均OCR信頼度70未満")
        records.append(QuantityRecord(
            symbol=symbol, floor=floor, file_name=file_name, drawing_no=drawing_no,
            page=page_no, count=len(matches), detection_ids=[d.detection_id for d in matches],
            mean_confidence=round(mean_conf, 1), status=status, reasons=reasons
        ))

    summaries = []
    for symbol in sorted(master_symbols):
        source = [r for r in records if r.symbol == symbol]
        floor_counts: dict[str, int] = {}
        for rec in source:
            floor_counts[rec.floor] = floor_counts.get(rec.floor, 0) + rec.count
        reasons = []
        if not source:
            status = QuantityStatus.UNCONFIRMED; reasons.append("数量集計対象の平面図で検出なし")
        elif any(r.status in {QuantityStatus.REVIEW, QuantityStatus.UNCONFIRMED} for r in source):
            status = QuantityStatus.REVIEW; reasons.extend(sorted({x for r in source for x in r.reasons}))
        else:
            status = QuantityStatus.CANDIDATE
        summaries.append(QuantitySummary(
            symbol=symbol, floor_counts=floor_counts, total=sum(floor_counts.values()),
            status=status, reasons=reasons, source_records=source
        ))
    return records, summaries


def detections_to_rows(detections: list[SymbolDetection]) -> list[dict]:
    rows=[]
    for d in detections:
        x=asdict(d); x["bbox"]=asdict(d.bbox); x["drawing_type"]=d.drawing_type.value
        rows.append(x)
    return rows


def summaries_to_rows(summaries: list[QuantitySummary]) -> list[dict]:
    return [{
        "建具符号": s.symbol,
        "階別数量": " / ".join(f"{k}:{v}" for k,v in sorted(s.floor_counts.items())),
        "合計数量": s.total,
        "数量判定": s.status.value,
        "判定理由": " / ".join(s.reasons),
        "根拠ページ": " / ".join(f"{r.drawing_no or r.file_name} p.{r.page}:{r.count}" for r in s.source_records),
    } for s in summaries]
