import io
import re
import hashlib
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="建具工事 見積前チェック", page_icon="🏗️", layout="wide")

STATUS_ORDER = {"見積前確認必須": 0, "要確認": 1, "見積準備OK": 2}
SYMBOL_RE = re.compile(r"(?<![A-Z0-9])((?:AW|AD|SD|WD|SW|AG|SG|FW|FD|WW|PG|PW|D|W)[-－_ ]?\d{1,3})(?!\d)", re.I)
SIZE_RES = [
    re.compile(r"(?:W|Ｗ)\s*[:=]?\s*(\d{2,5})\s*[×xX＊*]\s*(?:H|Ｈ)?\s*[:=]?\s*(\d{2,5})", re.I),
    re.compile(r"(\d{2,5})\s*[×xX＊*]\s*(\d{2,5})"),
    re.compile(r"(?:W|Ｗ)\s*[:=]?\s*(\d{2,5}).{0,15}?(?:H|Ｈ)\s*[:=]?\s*(\d{2,5})", re.I),
]
QTY_RE = re.compile(r"(?:数量|箇所|ヶ所|個数|台数|枚数|QTY|Qty)\s*[:：=]?\s*(\d{1,3})", re.I)

KEYWORDS = {
    "開閉形式": ["FIX", "引違", "片引", "両引", "片開", "両開", "縦すべり", "横すべり", "内倒", "上げ下げ", "排煙"],
    "材質": ["アルミ", "スチール", "ステンレス", "木製", "樹脂"],
    "ガラス": ["Low-E", "複層", "合わせガラス", "強化ガラス", "網入", "型板", "防火ガラス", "透明ガラス", "フロート"],
    "防火・性能": ["特定防火設備", "防火設備", "耐風圧", "気密", "水密", "遮音", "断熱"],
    "付属品": ["網戸", "面格子", "手すり", "フック棒", "オペレーター", "ドアクローザー", "フロアヒンジ", "戸当り", "シリンダー", "錠", "ガラリ", "水切り", "額縁"],
    "工事関連": ["撤去", "処分", "カバー工法", "はつり", "シール", "コーキング", "養生", "搬入", "揚重", "足場", "溶接", "アンカー", "自動ドア", "排煙装置"],
}

REQUIRED_CRITICAL = ["W", "H", "数量", "開閉形式"]
REQUIRED_CAUTION = ["ガラス", "防火・性能", "色", "付属品"]


def clean_symbol(s: str) -> str:
    return re.sub(r"[-－_ ]", "-", s.upper()).strip("-")


def extract_pdf_pages(uploaded_file):
    data = uploaded_file.getvalue()
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for idx, page in enumerate(reader.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        pages.append({"page": idx, "text": text, "chars": len(text.strip())})
    return pages, hashlib.sha256(data).hexdigest()[:12]


def nearest_size(context):
    for rx in SIZE_RES:
        m = rx.search(context)
        if m:
            return int(m.group(1)), int(m.group(2)), m.group(0)
    return None, None, ""


def first_keywords(context, words):
    found = []
    lower = context.lower()
    for word in words:
        if word.lower() in lower:
            found.append(word)
    return "、".join(dict.fromkeys(found))


def infer_color(context):
    colors = ["ブラック", "ホワイト", "シルバー", "ブロンズ", "ステンカラー", "グレー", "木目", "黒", "白"]
    return first_keywords(context, colors)


def analyze_pages(file_name, pages):
    rows = []
    symbol_occurrences = {}
    for page in pages:
        text = page["text"]
        for match in SYMBOL_RE.finditer(text):
            symbol = clean_symbol(match.group(1))
            start = max(0, match.start() - 100)
            end = min(len(text), match.end() + 260)
            context = re.sub(r"\s+", " ", text[start:end]).strip()
            w, h, size_raw = nearest_size(context)
            qty_match = QTY_RE.search(context)
            qty = int(qty_match.group(1)) if qty_match else None
            row = {
                "建具符号": symbol,
                "建具名称": "",
                "建具種別": "",
                "材質": first_keywords(context, KEYWORDS["材質"]),
                "開閉形式": first_keywords(context, KEYWORDS["開閉形式"]),
                "W": w,
                "H": h,
                "数量": qty,
                "設置階": "",
                "設置場所": "",
                "建具構成": first_keywords(context, ["連窓", "段窓", "方立", "無目", "ランマ", "袖", "ガラリ", "パネル"]),
                "ガラス": first_keywords(context, KEYWORDS["ガラス"]),
                "防火・性能": first_keywords(context, KEYWORDS["防火・性能"]),
                "色": infer_color(context),
                "付属品": first_keywords(context, KEYWORDS["付属品"]),
                "取付・躯体条件": first_keywords(context, ["RC", "ALC", "S造", "木造", "壁厚", "開口", "サッシ芯", "腰高", "FL", "シール", "アンカー"]),
                "工事関連": first_keywords(context, KEYWORDS["工事関連"]),
                "参照ファイル": file_name,
                "参照ページ": page["page"],
                "根拠抜粋": context[:320],
                "寸法原文": size_raw,
                "AI読取信頼度": "高" if w and h else "中" if context else "低",
            }
            rows.append(row)
            symbol_occurrences.setdefault(symbol, []).append(row)

    if not rows:
        return pd.DataFrame(), []

    # Consolidate duplicate observations by symbol, retaining evidence and flagging conflicts.
    consolidated, conflicts = [], []
    for symbol, obs in symbol_occurrences.items():
        def uniq(field):
            vals = [str(x[field]).strip() for x in obs if x.get(field) not in (None, "", "nan")]
            return list(dict.fromkeys(vals))
        base = obs[0].copy()
        pages_seen = uniq("参照ページ")
        base["参照ページ"] = ", ".join(pages_seen)
        base["根拠抜粋"] = " | ".join(uniq("根拠抜粋")[:3])
        for field in ["W", "H", "数量", "開閉形式", "材質", "ガラス", "防火・性能", "色", "付属品", "建具構成", "取付・躯体条件", "工事関連"]:
            vals = uniq(field)
            if len(vals) == 1:
                base[field] = int(vals[0]) if field in ["W", "H", "数量"] and vals[0].isdigit() else vals[0]
            elif len(vals) > 1:
                base[field] = " / ".join(vals)
                conflicts.append({"建具符号": symbol, "項目": field, "差異内容": " / ".join(vals), "参照ページ": base["参照ページ"]})
        consolidated.append(base)
    return pd.DataFrame(consolidated), conflicts


def judge_row(row, conflict_symbols):
    missing_critical = [x for x in REQUIRED_CRITICAL if pd.isna(row.get(x)) or str(row.get(x)).strip() in ("", "nan")]
    missing_caution = [x for x in REQUIRED_CAUTION if pd.isna(row.get(x)) or str(row.get(x)).strip() in ("", "nan")]
    reasons = []
    if row["建具符号"] in conflict_symbols:
        reasons.append("図面内で情報差異あり")
    if missing_critical:
        reasons.append("重要情報不足: " + "、".join(missing_critical))
    if missing_caution:
        reasons.append("確認推奨: " + "、".join(missing_caution))
    if row["建具符号"] in conflict_symbols or missing_critical:
        status = "見積前確認必須"
    elif missing_caution:
        status = "要確認"
    else:
        status = "見積準備OK"
    return status, " / ".join(reasons)


def make_excel(df, conflicts, project):
    wb = Workbook()
    ws = wb.active; ws.title = "建具拾い出し一覧"
    ws2 = wb.create_sheet("見積前要確認")
    ws3 = wb.create_sheet("図面不整合")
    ws4 = wb.create_sheet("案件情報")
    navy = "17365D"; blue = "D9EAF7"; red = "F4CCCC"; yellow = "FFF2CC"; green = "E2F0D9"
    thin = Side(style="thin", color="B7B7B7")
    headers = list(df.columns)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h); cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor=navy); cell.alignment = Alignment(horizontal="center"); cell.border = Border(bottom=thin)
    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(r_idx, c_idx, value); cell.alignment = Alignment(vertical="top", wrap_text=True)
        status = df.iloc[r_idx-2]["判定"]
        fill = red if status == "見積前確認必須" else yellow if status == "要確認" else green
        ws.cell(r_idx, headers.index("判定")+1).fill = PatternFill("solid", fgColor=fill)
    for i, h in enumerate(headers, 1):
        max_len = max([len(str(h))] + [len(str(ws.cell(r, i).value or "")) for r in range(2, min(ws.max_row, 100)+1)])
        ws.column_dimensions[get_column_letter(i)].width = min(max(max_len + 2, 10), 40)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions; ws.sheet_view.showGridLines = False

    issues = df[df["判定"] != "見積準備OK"]
    issue_cols = ["建具符号", "判定", "要確認理由", "W", "H", "数量", "参照ファイル", "参照ページ", "根拠抜粋"]
    for c,h in enumerate(issue_cols,1):
        cell=ws2.cell(1,c,h); cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor=navy)
    for r_idx, (_, row) in enumerate(issues.iterrows(),2):
        for c_idx,h in enumerate(issue_cols,1): ws2.cell(r_idx,c_idx,row.get(h,"")); ws2.cell(r_idx,c_idx).alignment=Alignment(wrap_text=True,vertical="top")
    ws2.sheet_view.showGridLines=False; ws2.freeze_panes="A2"

    conflict_cols=["建具符号","項目","差異内容","参照ページ"]
    for c,h in enumerate(conflict_cols,1):
        cell=ws3.cell(1,c,h); cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor=navy)
    for r_idx,row in enumerate(conflicts,2):
        for c_idx,h in enumerate(conflict_cols,1): ws3.cell(r_idx,c_idx,row.get(h,""))
    ws3.sheet_view.showGridLines=False

    info = [("案件名", project), ("出力日時", datetime.now().strftime("%Y-%m-%d %H:%M")), ("注意", "本結果は見積前の確認支援用です。製品適合、製作寸法、施工可否は図面原本・メーカー資料・設計者確認を優先してください。")]
    for r,(k,v) in enumerate(info,1): ws4.cell(r,1,k).font=Font(bold=True); ws4.cell(r,2,v); ws4.cell(r,2).alignment=Alignment(wrap_text=True)
    ws4.column_dimensions["A"].width=18; ws4.column_dimensions["B"].width=90; ws4.sheet_view.showGridLines=False
    bio=io.BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()


st.title("🏗️ 建具工事 見積前チェック・図面解析システム")
st.caption("建築図面PDFから建具情報を拾い、見積前の不足情報・図面内差異・確認事項を整理するVer.1 MVP")

with st.sidebar:
    st.header("案件情報")
    project = st.text_input("案件名", placeholder="例：○○新築工事")
    st.info("入力は建築図面PDFだけです。読めない値を推測で確定せず、要確認として出力します。")

files = st.file_uploader("建築図面PDFをアップロード", type=["pdf"], accept_multiple_files=True)
run = st.button("解析開始", type="primary", use_container_width=True, disabled=not files)

if run:
    all_df=[]; all_conflicts=[]; scan_warnings=[]
    with st.spinner("建築図面を解析しています..."):
        for f in files:
            pages, file_hash = extract_pdf_pages(f)
            low_text_pages=[p["page"] for p in pages if p["chars"] < 20]
            if low_text_pages:
                scan_warnings.append(f"{f.name}: テキストが少ないページ {low_text_pages}。画像PDFの場合はOCR追加が必要です。")
            df, conflicts=analyze_pages(f.name,pages)
            if not df.empty: all_df.append(df)
            all_conflicts.extend(conflicts)
    if not all_df:
        st.error("建具符号を抽出できませんでした。画像PDF、特殊な符号表記、または文字情報を持たない図面の可能性があります。")
        if scan_warnings:
            for x in scan_warnings: st.warning(x)
        st.stop()
    df=pd.concat(all_df,ignore_index=True)
    # Merge same symbol across files conservatively by retaining each source row.
    conflict_symbols={x["建具符号"] for x in all_conflicts}
    judgments=df.apply(lambda r: judge_row(r,conflict_symbols),axis=1)
    df["判定"]=[x[0] for x in judgments]
    df["要確認理由"]=[x[1] for x in judgments]
    df["人による修正値"]=""
    df["確認状態"]="未確認"
    df=df.sort_values(by="判定",key=lambda s:s.map(STATUS_ORDER)).reset_index(drop=True)
    st.session_state["result_df"]=df
    st.session_state["conflicts"]=all_conflicts
    st.session_state["project"]=project or "未設定案件"
    st.session_state["warnings"]=scan_warnings

if "result_df" in st.session_state:
    df=st.session_state["result_df"]
    counts=df["判定"].value_counts()
    c1,c2,c3,c4=st.columns(4)
    c1.metric("抽出建具数",len(df))
    c2.metric("見積準備OK",int(counts.get("見積準備OK",0)))
    c3.metric("要確認",int(counts.get("要確認",0)))
    c4.metric("見積前確認必須",int(counts.get("見積前確認必須",0)))
    for x in st.session_state.get("warnings",[]): st.warning(x)
    tab1,tab2,tab3,tab4=st.tabs(["要確認一覧","建具拾い出し","図面不整合","出力"])
    with tab1:
        issue_df=df[df["判定"]!="見積準備OK"]
        st.dataframe(issue_df[["建具符号","判定","要確認理由","W","H","数量","参照ファイル","参照ページ"]],use_container_width=True,hide_index=True)
    with tab2:
        display_cols=["建具符号","建具種別","材質","開閉形式","W","H","数量","建具構成","ガラス","防火・性能","色","付属品","工事関連","判定","参照ファイル","参照ページ"]
        st.dataframe(df[display_cols],use_container_width=True,hide_index=True)
        with st.expander("根拠抜粋を表示"):
            st.dataframe(df[["建具符号","参照ファイル","参照ページ","根拠抜粋"]],use_container_width=True,hide_index=True)
    with tab3:
        conflicts=st.session_state.get("conflicts",[])
        if conflicts: st.dataframe(pd.DataFrame(conflicts),use_container_width=True,hide_index=True)
        else: st.success("抽出範囲では図面内差異は検出されませんでした。")
    with tab4:
        xlsx=make_excel(df,st.session_state.get("conflicts",[]),st.session_state.get("project","未設定案件"))
        st.download_button("Excel報告書をダウンロード",xlsx,file_name=f"{st.session_state.get('project','案件')}_建具見積前チェック.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.caption("Excelには、建具拾い出し一覧・見積前要確認・図面不整合・案件情報を収録します。")

st.divider()
st.markdown("**重要:** このVer.1は、文字情報を持つPDFを対象にしたルールベースMVPです。画像スキャン図面にはOCR、図形からの数量カウントには画像解析、メーカー製品の決定には製品マスターが別途必要です。")
