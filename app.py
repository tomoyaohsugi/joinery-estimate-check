import io
import re
from datetime import datetime

import cv2
import numpy as np
import pandas as pd
import pytesseract
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pdf2image import convert_from_bytes
from PIL import Image, ImageDraw
from pypdf import PdfReader
from pytesseract import Output

from quantity_logic import (
    BoundingBox, CountPolicy, DrawingType, JoineryMaster, PageMetadata,
    SymbolDetection, aggregate_quantities, classify_drawing,
    default_count_policy, detections_to_rows, normalize_symbol,
    prepare_detections, summaries_to_rows,
)

st.set_page_config(page_title="建具数量集計 Ver.2-A", page_icon="🏗️", layout="wide")

SYMBOL_RX = re.compile(r"^(?:AW|AD|SD|WD|SW|AG|SG|FW|FD|WW|PG|PW|D|W)[\-－_ ]?[0-9OIL]{1,3}$", re.I)
TITLE_WORDS = ["建具表", "建具キープラン", "建具配置図", "平面図", "立面図", "断面図", "矩計図", "建具詳細図", "サッシ詳細図", "納まり図"]


def preprocess(img: Image.Image) -> Image.Image:
    arr = np.array(img.convert("RGB"))
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    gray = cv2.resize(gray, None, fx=1.35, fy=1.35, interpolation=cv2.INTER_CUBIC)
    gray = cv2.bilateralFilter(gray, 5, 40, 40)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 31, 11)
    return Image.fromarray(binary)


def ocr_page(img: Image.Image):
    processed = preprocess(img)
    data = pytesseract.image_to_data(
        processed, lang="jpn+eng",
        config="--oem 3 --psm 11 -c preserve_interword_spaces=1",
        output_type=Output.DICT,
    )
    scale = 1.35
    tokens = []
    for i, text in enumerate(data["text"]):
        text = str(text).strip()
        try:
            conf = float(data["conf"][i])
        except Exception:
            conf = -1
        if not text or conf < 15:
            continue
        tokens.append({
            "text": text,
            "conf": conf,
            "x": float(data["left"][i]) / scale,
            "y": float(data["top"][i]) / scale,
            "w": float(data["width"][i]) / scale,
            "h": float(data["height"][i]) / scale,
        })
    text = " ".join(t["text"] for t in tokens)
    return tokens, text


def combine_symbol_tokens(tokens):
    """単一トークンと、横方向に分割された AW - 01 等を候補化する。"""
    candidates = []
    for i, t in enumerate(tokens):
        compact = re.sub(r"\s+", "", t["text"])
        if SYMBOL_RX.fullmatch(compact):
            candidates.append((i, i, compact))
        for length in (2, 3):
            if i + length > len(tokens):
                continue
            block = tokens[i:i+length]
            # 同じ行、高さの2倍以内の隙間に限定
            if max(x["y"] for x in block) - min(x["y"] for x in block) > max(x["h"] for x in block) * 0.8:
                continue
            ordered = sorted(block, key=lambda x: x["x"])
            gaps = [ordered[j+1]["x"] - (ordered[j]["x"] + ordered[j]["w"]) for j in range(len(ordered)-1)]
            if gaps and max(gaps) > max(x["h"] for x in block) * 2.5:
                continue
            joined = "".join(x["text"] for x in ordered)
            joined = re.sub(r"\s+", "", joined)
            if SYMBOL_RX.fullmatch(joined):
                candidates.append((i, i+length-1, joined))
    # 同じ範囲・同じ正規化値を排除
    unique = {}
    for a, b, raw in candidates:
        unique[(a, b, normalize_symbol(raw))] = (a, b, raw)
    return list(unique.values())


def detect_floor(text):
    m = re.search(r"([BＲR]?\d{1,2})\s*階", text, re.I)
    if m:
        return m.group(1).upper() + "階"
    m = re.search(r"(地上|地下)?\s*([一二三四五六七八九十0-9]+)\s*階", text)
    return (m.group(1) or "") + m.group(2) + "階" if m else ""


def detect_title(text):
    compact = re.sub(r"\s+", "", text)
    hits = [w for w in TITLE_WORDS if w in compact]
    floor = detect_floor(compact)
    return (floor + (hits[0] if hits else "")) if floor or hits else compact[:80]


def make_page_data(file_name, page_no, img, tokens, text):
    title = detect_title(text)
    dtype, confidence = classify_drawing(title)
    policy = default_count_policy(dtype)
    meta = PageMetadata(
        file_name=file_name, page=page_no, width=img.width, height=img.height,
        title_text=title, drawing_no="", floor=detect_floor(text),
        drawing_type=dtype, drawing_type_confidence=confidence,
        count_policy=policy, is_latest_revision=True,
    )
    detections = []
    for n, (a, b, raw) in enumerate(combine_symbol_tokens(tokens), 1):
        block = tokens[a:b+1]
        x1 = min(t["x"] for t in block); y1 = min(t["y"] for t in block)
        x2 = max(t["x"] + t["w"] for t in block); y2 = max(t["y"] + t["h"] for t in block)
        confidence_value = sum(t["conf"] for t in block) / len(block)
        detections.append(SymbolDetection(
            detection_id=f"{file_name}-p{page_no}-d{n}", raw_text=raw,
            normalized_symbol=normalize_symbol(raw), file_name=file_name, page=page_no,
            bbox=BoundingBox(x1, y1, x2-x1, y2-y1),
            ocr_confidence=round(confidence_value, 1), drawing_type=dtype,
            floor=meta.floor, drawing_no=meta.drawing_no,
        ))
    return meta, detections


def render_overlay(img, detections, selected_ids=None):
    selected_ids = selected_ids or set()
    out = img.copy().convert("RGB")
    draw = ImageDraw.Draw(out)
    for d in detections:
        color = "#00A000" if d.detection_id in selected_ids else "#E00000"
        b = d.bbox
        draw.rectangle((b.x, b.y, b.x+b.width, b.y+b.height), outline=color, width=4)
        draw.text((b.x, max(0, b.y-18)), d.effective_symbol, fill=color)
    return out


def export_excel(summary_rows, detection_rows, page_rows, project):
    wb = Workbook(); ws = wb.active; ws.title = "数量集計"
    datasets = [
        (ws, pd.DataFrame(summary_rows)),
        (wb.create_sheet("検出位置"), pd.DataFrame(detection_rows)),
        (wb.create_sheet("ページ設定"), pd.DataFrame(page_rows)),
    ]
    for sh, df in datasets:
        if df.empty:
            sh["A1"] = "該当なし"; continue
        for c, h in enumerate(df.columns, 1):
            cell = sh.cell(1, c, h); cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="17365D")
        for r, row in enumerate(df.itertuples(index=False), 2):
            for c, value in enumerate(row, 1):
                sh.cell(r, c, str(value) if isinstance(value, (dict, list)) else value)
                sh.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
        for c in range(1, sh.max_column + 1):
            sh.column_dimensions[get_column_letter(c)].width = 22
        sh.freeze_panes = "A2"
    info = wb.create_sheet("案件情報")
    info.append(["案件名", project]); info.append(["出力日時", datetime.now().strftime("%Y-%m-%d %H:%M")])
    info.append(["注意", "数量はOCR検出候補です。図面上の検出位置を確認して確定してください。"]) 
    bio = io.BytesIO(); wb.save(bio); return bio.getvalue()


st.title("建具数量集計・見積前チェック Ver.2-A")
st.caption("平面図だけを数量集計対象にし、建具表・立面図・詳細図は数量へ加算しません。")
with st.sidebar:
    project = st.text_input("案件名", placeholder="例：○○新築工事")
    dpi = st.select_slider("OCR解像度（DPI）", [200, 250, 300], 250)
    min_conf = st.slider("最低OCR信頼度", 20, 90, 45)
    st.warning("数量は検出位置を確認してから使用してください。")

files = st.file_uploader("建築図面PDF", type=["pdf"], accept_multiple_files=True)
if st.button("OCR解析", type="primary", disabled=not files, use_container_width=True):
    pages = {}; detections = []; images = {}; progress = st.progress(0)
    total_pages = sum(len(PdfReader(io.BytesIO(f.getvalue())).pages) for f in files)
    done = 0
    for f in files:
        raw = f.getvalue(); reader = PdfReader(io.BytesIO(raw))
        for page_no in range(1, len(reader.pages)+1):
            img = convert_from_bytes(raw, dpi=dpi, first_page=page_no, last_page=page_no, fmt="png", thread_count=1)[0]
            tokens, text = ocr_page(img)
            meta, dets = make_page_data(f.name, page_no, img, tokens, text)
            pages[(f.name, page_no)] = meta; detections.extend(dets); images[(f.name, page_no)] = img
            done += 1; progress.progress(done/total_pages, text=f"{done}/{total_pages}ページ")
    st.session_state.update(pages=pages, raw_detections=detections, page_images=images,
                            project=project or "未設定案件")
    st.session_state.pop("quantity_result", None)

if "pages" in st.session_state:
    st.subheader("1. ページ種別・数量集計対象を確認")
    page_rows = []
    for key, p in st.session_state.pages.items():
        page_rows.append({
            "key": f"{key[0]}::{key[1]}", "ファイル": key[0], "ページ": key[1],
            "図面タイトル候補": p.title_text, "図面種別": p.drawing_type.value,
            "階": p.floor, "集計方針": p.count_policy.value,
        })
    edited = st.data_editor(pd.DataFrame(page_rows), hide_index=True, use_container_width=True,
        disabled=["key", "ファイル", "ページ"],
        column_config={
            "図面種別": st.column_config.SelectboxColumn(options=[x.value for x in DrawingType]),
            "集計方針": st.column_config.SelectboxColumn(options=[x.value for x in CountPolicy]),
        }, key="page_editor")
    for _, row in edited.iterrows():
        file_name, page_str = row["key"].rsplit("::", 1); key = (file_name, int(page_str)); p = st.session_state.pages[key]
        p.title_text = row["図面タイトル候補"]; p.drawing_type = DrawingType(row["図面種別"])
        p.floor = row["階"]; p.count_policy = CountPolicy(row["集計方針"])
    st.subheader("2. 建具マスター")
    detected_symbols = sorted({normalize_symbol(d.raw_text) for d in st.session_state.raw_detections})
    master_df = st.data_editor(pd.DataFrame({"建具符号": detected_symbols, "採用": [True]*len(detected_symbols)}),
                               hide_index=True, use_container_width=True, num_rows="dynamic", key="master_editor")
    masters = [JoineryMaster(str(r["建具符号"])) for _, r in master_df.iterrows() if r.get("採用") and str(r.get("建具符号", "")).strip()]

    if st.button("数量候補を集計", type="primary", disabled=not masters):
        # OCR再解析なしでページ設定とマスターを反映
        detections = st.session_state.raw_detections
        for d in detections:
            p = st.session_state.pages[(d.file_name, d.page)]
            d.drawing_type = p.drawing_type; d.floor = p.floor; d.drawing_no = p.drawing_no
            d.excluded = False; d.exclusion_reason = ""; d.duplicate_of = ""
        prepare_detections(detections, st.session_state.pages, masters, min_confidence=float(min_conf))
        records, summaries = aggregate_quantities(detections, st.session_state.pages, masters)
        st.session_state.quantity_result = (detections, records, summaries)

if "quantity_result" in st.session_state:
    detections, records, summaries = st.session_state.quantity_result
    summary_rows = summaries_to_rows(summaries)
    st.subheader("3. 階別・ページ別数量候補")
    st.dataframe(pd.DataFrame(summary_rows), hide_index=True, use_container_width=True)

    st.subheader("4. 検出位置を確認")
    count_pages = [k for k, p in st.session_state.pages.items() if p.count_policy == CountPolicy.COUNT]
    if count_pages:
        selected_page = st.selectbox("確認するページ", count_pages, format_func=lambda k: f"{k[0]} p.{k[1]} {st.session_state.pages[k].floor}")
        page_dets = [d for d in detections if (d.file_name, d.page) == selected_page]
        valid_ids = {d.detection_id for d in page_dets if not d.excluded}
        overlay = render_overlay(st.session_state.page_images[selected_page], page_dets, valid_ids)
        st.image(overlay, caption="緑＝集計対象、赤＝除外候補", use_container_width=True)
        review_df = pd.DataFrame([{
            "採用": not d.excluded, "検出ID": d.detection_id, "建具符号": d.effective_symbol,
            "OCR信頼度": d.ocr_confidence, "除外理由": d.exclusion_reason,
            "X": round(d.bbox.x), "Y": round(d.bbox.y),
        } for d in page_dets])
        if not review_df.empty:
            corrected = st.data_editor(review_df, hide_index=True, use_container_width=True,
                                       disabled=["検出ID", "OCR信頼度", "除外理由", "X", "Y"], key=f"review-{selected_page}")
            if st.button("このページの修正を反映"):
                adopted = set(corrected.loc[corrected["採用"], "検出ID"])
                symbol_updates = dict(zip(corrected["検出ID"], corrected["建具符号"]))
                for d in page_dets:
                    d.excluded = d.detection_id not in adopted
                    d.exclusion_reason = "手動除外" if d.excluded else ""
                    d.matched_master_symbol = normalize_symbol(str(symbol_updates[d.detection_id]))
                    d.manually_confirmed = True
                _, summaries = aggregate_quantities(detections, st.session_state.pages,
                    [JoineryMaster(str(x)) for x in pd.DataFrame(summary_rows)["建具符号"]])
                st.session_state.quantity_result = (detections, records, summaries)
                st.success("修正を反映しました。")
    else:
        st.warning("数量集計に使用するページがありません。ページ設定で平面図を選択してください。")

    st.subheader("5. Excel出力")
    page_export = [{"ファイル": p.file_name, "ページ": p.page, "図面種別": p.drawing_type.value,
                    "階": p.floor, "集計方針": p.count_policy.value} for p in st.session_state.pages.values()]
    detection_rows = detections_to_rows(detections)
    st.download_button("数量集計Excelをダウンロード",
        export_excel(summaries_to_rows(st.session_state.quantity_result[2]), detection_rows, page_export, st.session_state.project),
        file_name=f"{st.session_state.project}_建具数量集計.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
