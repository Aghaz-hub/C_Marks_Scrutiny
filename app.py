"""
Exam Cover Page -> Excel
------------------------
Faculty upload photos / scans / PDFs of examination-copy cover pages
(the page with the "Marks Distribution" table). The app reads each page with
Claude's vision, lets the faculty review and correct the extracted marks, and
downloads one Excel workbook:

  * Script Details       - exam details and a processing summary
  * Marks Check          - question-wise table per student with live formulas
  * Student-wise Marks   - ONE ROW PER STUDENT, sorted by roll number:
                           Roll No | Name | Q.No1(a) | Q.No1(b) | ... | Total
"""

import base64
import hmac
import io
import json
import os
import re
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #
LETTERS = "abcdefghij"
PART_RE = re.compile(r"^Q\.No(\d+)\(([a-j])\)$")
MODELS = ["claude-sonnet-5-5", "claude-opus-5-5", "claude-haiku-5-5"]
GEMINI_MODELS = ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-2.5-flash"]
PROVIDERS = {"Gemini (free tier)": "gemini", "Claude (paid)": "claude"}
SECRET_NAMES = {"gemini": "GEMINI_API_KEY", "claude": "ANTHROPIC_API_KEY"}
WORKERS_BY_PROVIDER = {"gemini": 2, "claude": 4}  # free tier has low per-minute limits
PDF_DPI = 150
MAX_IMAGE_SIDE = 2000

TOOL = {
    "name": "record_cover_page",
    "description": "Record everything legible on one university answer-script cover page.",
    "input_schema": {
        "type": "object",
        "properties": {
            "is_cover_page": {
                "type": "boolean",
                "description": "True only if this image shows an answer-script cover / marks-distribution page.",
            },
            "candidate_name": {"type": "string"},
            "roll_no": {"type": "string"},
            "examination": {"type": "string"},
            "subject": {"type": "string"},
            "subject_code": {"type": "string"},
            "exam_date": {"type": "string"},
            "session": {"type": "string"},
            "semester": {"type": "string"},
            "questions": {
                "type": "array",
                "description": "One entry per row of the Marks Distribution table (Q.No 1..N).",
                "items": {
                    "type": "object",
                    "properties": {
                        "q_no": {"type": "integer"},
                        "parts": {
                            "type": "array",
                            "description": "Only the part columns (a, b, c, ...) that contain a mark.",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "part": {"type": "string"},
                                    "marks": {"type": ["number", "null"]},
                                },
                                "required": ["part", "marks"],
                            },
                        },
                        "total": {"type": ["number", "null"]},
                    },
                    "required": ["q_no", "parts", "total"],
                },
            },
            "total_marks_obtained": {"type": ["number", "null"]},
            "marks_in_words": {"type": "string"},
            "uncertain": {
                "type": "array",
                "description": "Short notes on anything hard to read or ambiguous.",
                "items": {"type": "string"},
            },
        },
        "required": ["is_cover_page"],
    },
}

PROMPT = """This is a photo or scan of the cover page of a university examination answer script.

Layout: the LEFT half has a "Marks Distribution" table written by the examiner (usually in red ink): rows are Q.No 1..10, columns are parts a, b, c ... j, plus a Total column. Below it are "Total Marks Obtained" and "Marks (in Words)". The RIGHT half has the candidate's details (name, roll number, subject, date, semester, session) usually in blue ink.

Call record_cover_page with what is written. Rules:
- Transcribe exactly what is written. Do NOT calculate or correct any total yourself.
- A dash, blank cell or strike-through means no marks: leave that part out of "parts" and use null for the question total.
- Put each mark under the correct column letter (a, b, c ...). Be careful about which column a number sits in.
- Roll numbers look like 2K23CSUN01102 - transcribe them exactly, upper-case, no spaces.
- If a digit or word is hard to read, give your best reading and add a short note to "uncertain" (for example "Q3 total looks like 8 or 6").
- If the image is not a cover / marks page, set is_cover_page to false and leave everything else empty.
- Marks may be half marks (for example 2.5)."""

# --------------------------------------------------------------------------- #
# Helpers: numbers, words, sorting
# --------------------------------------------------------------------------- #
_UNITS = {
    "zero": 0, "nil": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11,
    "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16,
    "seventeen": 17, "eighteen": 18, "nineteen": 19,
}
_TENS = {"twenty": 20, "thirty": 30, "forty": 40, "fourty": 40, "fifty": 50,
         "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90}


def words_to_number(text):
    """'Thirty Three only' -> 33. Returns None if it cannot be parsed."""
    if not text or not str(text).strip():
        return None
    t = re.sub(r"[^a-z\s-]", " ", str(text).lower()).replace("-", " ")
    tokens = [w for w in t.split() if w not in ("only", "and", "marks", "mark")]
    if not tokens:
        return None
    total = 0
    for w in tokens:
        if w in _UNITS:
            total += _UNITS[w]
        elif w in _TENS:
            total += _TENS[w]
        elif w == "hundred":
            total = max(total, 1) * 100
        else:
            return None
    return total


def natural_key(value):
    s = "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()
    parts = [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", s)]
    return (s == "", parts)  # blank roll numbers go last


def clean(v):
    """NaN -> None, 3.0 -> 3 (for writing to Excel)."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def is_num(v):
    return clean(v) is not None and isinstance(clean(v), (int, float))


# --------------------------------------------------------------------------- #
# Reading pages and calling Claude
# --------------------------------------------------------------------------- #
def prep_image(img_bytes):
    """Fix rotation, downscale, return base64 JPEG."""
    img = Image.open(io.BytesIO(img_bytes))
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((MAX_IMAGE_SIDE, MAX_IMAGE_SIDE))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=88)
    return base64.b64encode(buf.getvalue()).decode()


def collect_pages(uploaded_files, max_pages):
    """Turn uploads (jpg/png/pdf) into a flat list of page images."""
    import pymupdf as fitz  # PyMuPDF

    pages, notes = [], []
    for f in uploaded_files:
        data, name = f.getvalue(), f.name
        ext = name.lower().rsplit(".", 1)[-1]
        try:
            if ext == "pdf":
                doc = fitz.open(stream=data, filetype="pdf")
                for i, page in enumerate(doc, start=1):
                    if len(pages) >= max_pages:
                        break
                    png = page.get_pixmap(dpi=PDF_DPI).tobytes("png")
                    pages.append({"label": f"{name} (page {i})", "source": f"{name} p{i}", "b64": prep_image(png)})
            else:
                if len(pages) < max_pages:
                    pages.append({"label": name, "source": name, "b64": prep_image(data)})
        except Exception as e:  # corrupt file etc.
            notes.append(f"Could not read {name}: {e}")
        if len(pages) >= max_pages:
            notes.append(f"Page limit ({max_pages}) reached - remaining pages were skipped.")
            break
    return pages, notes


def extract_page(client, model, page):
    """Claude backend (tool use gives guaranteed-structured output)."""
    resp = client.messages.create(
        model=model,
        max_tokens=3000,
        tools=[TOOL],
        tool_choice={"type": "tool", "name": TOOL["name"]},
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": page["b64"]}},
                {"type": "text", "text": PROMPT},
            ],
        }],
    )
    for block in resp.content:
        if block.type == "tool_use":
            return block.input
    raise RuntimeError("The model returned no structured result.")


JSON_SHAPE = """Reply with ONLY a JSON object (no markdown, no commentary) in exactly this shape:
{"is_cover_page": true, "candidate_name": "", "roll_no": "", "examination": "", "subject": "",
 "subject_code": "", "exam_date": "", "session": "", "semester": "",
 "questions": [{"q_no": 1, "parts": [{"part": "a", "marks": 3}], "total": 3}],
 "total_marks_obtained": 33, "marks_in_words": "", "uncertain": []}
Use null for a mark or total that is blank or a dash."""

GEMINI_PROMPT = PROMPT.replace("Call record_cover_page with what is written.",
                               "Extract what is written.") + "\n\n" + JSON_SHAPE


def parse_json(text):
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t)
    data = json.loads(t)
    if isinstance(data, list):
        data = data[0] if data else {}
    if not isinstance(data, dict):
        raise ValueError("Unexpected JSON shape")
    return data


def make_extractor(provider, key, model):
    """Return a function page -> dict for the chosen provider."""
    if provider == "claude":
        import anthropic
        client = anthropic.Anthropic(api_key=key, max_retries=3)
        return lambda page: extract_page(client, model, page)

    from google import genai
    from google.genai import types
    client = genai.Client(api_key=key)
    cfg = types.GenerateContentConfig(response_mime_type="application/json", temperature=0)

    def run(page):
        img = types.Part.from_bytes(data=base64.b64decode(page["b64"]), mime_type="image/jpeg")
        delay = 6
        for attempt in range(6):
            try:
                resp = client.models.generate_content(model=model, contents=[img, GEMINI_PROMPT], config=cfg)
                return parse_json(resp.text)
            except Exception as e:
                msg = str(e)
                transient = any(t in msg for t in ("429", "RESOURCE_EXHAUSTED", "500", "503", "UNAVAILABLE"))
                if not transient or attempt == 5:
                    raise
                time.sleep(delay)  # free tier rate limit: wait and retry
                delay = min(delay * 2, 60)

    return run


def run_extraction(extract_fn, pages, workers, progress_cb=None):
    results = [None] * len(pages)

    def work(i):
        try:
            return i, {"ok": True, "data": extract_fn(pages[i])}
        except Exception as e:
            return i, {"ok": False, "error": str(e)[:300]}

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = [ex.submit(work, i) for i in range(len(pages))]
        for n, fut in enumerate(as_completed(futures), start=1):
            i, res = fut.result()
            results[i] = res
            if progress_cb:
                progress_cb(n, len(pages))
    return results


# --------------------------------------------------------------------------- #
# Results -> DataFrame
# --------------------------------------------------------------------------- #
def part_columns(n_q, n_parts=len(LETTERS)):
    cols = []
    for q in range(1, n_q + 1):
        cols += [f"Q.No{q}({l})" for l in LETTERS[:n_parts]]
        cols.append(f"Q.No{q} Total")
    return cols


def results_to_df(pages, results, n_q):
    rows, log, meta_vals = [], [], {k: [] for k in
        ("examination", "subject", "subject_code", "exam_date", "session", "semester")}

    for page, res in zip(pages, results):
        if not res["ok"]:
            log.append({"Page": page["label"], "Status": "Error", "Detail": res["error"], "Roll No": ""})
            continue
        d = res["data"]
        has_content = d.get("questions") or d.get("roll_no") or d.get("candidate_name")
        if not d.get("is_cover_page", False) or not has_content:
            log.append({"Page": page["label"], "Status": "Skipped", "Detail": "Not a cover / marks page", "Roll No": ""})
            continue

        roll = re.sub(r"\s+", "", str(d.get("roll_no", ""))).upper()
        row = {"Roll No": roll, "Student Name": str(d.get("candidate_name", "")).strip()}
        qmap = {int(q.get("q_no", 0)): q for q in (d.get("questions") or [])}
        for q in range(1, n_q + 1):
            qd = qmap.get(q, {})
            pm = {str(p.get("part", "")).strip().lower(): p.get("marks") for p in qd.get("parts", [])}
            for l in LETTERS:
                row[f"Q.No{q}({l})"] = pm.get(l)
            row[f"Q.No{q} Total"] = qd.get("total")
        row["Total Marks"] = d.get("total_marks_obtained")
        row["Marks in Words"] = str(d.get("marks_in_words", "")).strip()
        row["Extraction Notes"] = "; ".join(d.get("uncertain") or [])
        row["Source"] = page["source"]
        rows.append(row)

        for k in meta_vals:
            if str(d.get(k, "")).strip():
                meta_vals[k].append(str(d[k]).strip())
        log.append({"Page": page["label"], "Status": "Read", "Detail": "", "Roll No": roll})

    df = pd.DataFrame(rows)
    if df.empty:
        return df, log, {}
    for c in part_columns(n_q) + ["Total Marks"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.sort_values("Roll No", key=lambda s: s.map(natural_key)).reset_index(drop=True)
    meta = {k: Counter(v).most_common(1)[0][0] if v else "" for k, v in meta_vals.items()}
    return df, log, meta


def hide_empty_part_columns(df, n_q):
    """Drop part columns (a..j) that are empty for every student."""
    keep = []
    for c in df.columns:
        m = PART_RE.match(c)
        if m and df[c].isna().all():
            continue
        keep.append(c)
    return df[keep]


# --------------------------------------------------------------------------- #
# Validation
# --------------------------------------------------------------------------- #
def fmt(v):
    v = clean(v)
    return "" if v is None else str(v)


def review_row(row, n_q):
    flags = []
    if not str(row.get("Roll No", "") or "").strip():
        flags.append("roll no missing")
    if not str(row.get("Student Name", "") or "").strip():
        flags.append("name missing")

    q_totals = []
    for q in range(1, n_q + 1):
        vals = [clean(row.get(f"Q.No{q}({l})")) for l in LETTERS]
        vals = [v for v in vals if v is not None]
        total = clean(row.get(f"Q.No{q} Total"))
        if vals and total is None:
            flags.append(f"Q{q} total missing")
            q_totals.append(sum(vals))
        elif vals and abs(sum(vals) - total) > 1e-9:
            flags.append(f"Q{q} parts add to {fmt(sum(vals))} but total is {fmt(total)}")
            q_totals.append(total)
        else:
            q_totals.append(total or 0)

    grand = clean(row.get("Total Marks"))
    if grand is None:
        flags.append("total marks missing")
    elif abs(sum(q_totals) - grand) > 1e-9:
        flags.append(f"question totals add to {fmt(sum(q_totals))} but sheet total is {fmt(grand)}")

    words_n = words_to_number(row.get("Marks in Words"))
    if words_n is not None and grand is not None and abs(words_n - grand) > 1e-9:
        flags.append(f"words say {words_n} but total is {fmt(grand)}")

    notes = str(row.get("Extraction Notes", "") or "").strip()
    if notes:
        flags.append(f"unclear: {notes}")
    return "OK" if not flags else "CHECK: " + "; ".join(flags)


def compute_review(df, n_q):
    if df.empty:
        return pd.Series([], dtype=str)
    review = df.apply(lambda r: review_row(r, n_q), axis=1)
    roll = df["Roll No"].astype(str).str.strip()
    dup = roll.duplicated(keep=False) & roll.ne("")
    for i in df.index[dup]:
        review[i] = ("CHECK: duplicate roll no" if review[i] == "OK"
                     else review[i].replace("CHECK: ", "CHECK: duplicate roll no; ", 1))
    return review


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
_F = "Arial"
_HDR_FILL = PatternFill("solid", fgColor="1F3864")
_HDR_FONT = Font(name=_F, bold=True, color="FFFFFF", size=10)
_BODY = Font(name=_F, size=10)
_INPUT = Font(name=_F, size=10, color="0000FF")
_BOLD = Font(name=_F, size=10, bold=True)
_TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
_BAD_FILL = PatternFill("solid", fgColor="F8CBAD")
_thin = Side(style="thin", color="999999")
_BOX = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
_NUM = '0.##;-0.##;"-"'


def _header(ws, row, headers):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row, j, h)
        c.font, c.fill, c.alignment, c.border = _HDR_FONT, _HDR_FILL, _CENTER, _BOX
    ws.row_dimensions[row].height = 32


def _letters_present(df):
    idx = [LETTERS.index(m.group(2)) for c in df.columns for m in [PART_RE.match(c)] if m]
    last = max(idx) if idx else 2
    return list(LETTERS[: max(last, 2) + 1])  # always at least a, b, c


def build_excel(df, meta, n_q):
    df = df.copy()
    df = df.sort_values("Roll No", key=lambda s: s.map(natural_key)).reset_index(drop=True)
    df["Review"] = compute_review(df, n_q)
    letters = _letters_present(df)
    wb = Workbook()

    # ---- Sheet 1: Script Details -------------------------------------------
    ws = wb.active
    ws.title = "Script Details"
    ws["A1"] = "Answer Script Cover Sheets - Summary"
    ws["A1"].font = Font(name=_F, size=13, bold=True)
    _header(ws, 3, ["Field", "Value"])
    n_review = int((df["Review"] != "OK").sum()) if len(df) else 0
    rows = [
        ("Examination", meta.get("examination", "")),
        ("Subject", meta.get("subject", "")),
        ("Subject Code", meta.get("subject_code", "")),
        ("Date of Exam", meta.get("exam_date", "")),
        ("Session", meta.get("session", "")),
        ("Semester", meta.get("semester", "")),
        ("Scripts in this file", len(df)),
        ("Scripts flagged for review", n_review),
        ("Generated on", datetime.now().strftime("%d-%m-%Y %H:%M")),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        ws.cell(i, 1, k).font = _BOLD
        ws.cell(i, 2, v).font = _INPUT if i < 10 else _BODY
        for c in (1, 2):
            ws.cell(i, c).border, ws.cell(i, c).alignment = _BOX, _LEFT
    ws["A14"] = ("Exam details are the most common value read across all pages. "
                 "Blue = read from the scans; black = computed.")
    ws["A14"].font = Font(name=_F, size=9, italic=True)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 36

    # ---- Sheet 2: Marks Check ----------------------------------------------
    mc = wb.create_sheet("Marks Check")
    mc["A1"] = "Marks Table Check (per student)"
    mc["A1"].font = Font(name=_F, size=13, bold=True)
    c_first = 4
    c_sum = c_first + len(letters)
    c_rec, c_diff, c_stat, c_box, c_rem = c_sum + 1, c_sum + 2, c_sum + 3, c_sum + 4, c_sum + 5
    L = get_column_letter
    _header(mc, 3, ["Roll No", "Student Name", "Q.No."] + [f"Part {l}" for l in letters] +
            ["Sum of Parts", "Total Recorded on Sheet", "Difference", "Status", "Sheet Total (box)", "Remark"])
    r = 4
    for _, s in df.iterrows():
        start = r
        for q in range(1, n_q + 1):
            mc.cell(r, 1, s["Roll No"])
            mc.cell(r, 2, s["Student Name"])
            mc.cell(r, 3, q)
            for j, l in enumerate(letters):
                v = clean(s.get(f"Q.No{q}({l})"))
                if v is not None:
                    mc.cell(r, c_first + j, v)
            mc.cell(r, c_sum, f"=SUM({L(c_first)}{r}:{L(c_sum - 1)}{r})")
            rec = clean(s.get(f"Q.No{q} Total"))
            if rec is not None:
                mc.cell(r, c_rec, rec)
            mc.cell(r, c_diff, f"={L(c_sum)}{r}-{L(c_rec)}{r}")
            mc.cell(r, c_stat, f'=IF({L(c_diff)}{r}=0,"OK","CHECK")')
            for col in range(1, c_rem + 1):
                cell = mc.cell(r, col)
                is_in = c_first <= col < c_sum or col == c_rec
                cell.font, cell.border = (_INPUT if is_in else _BODY), _BOX
                cell.alignment = _CENTER if col > 2 else _LEFT
                if c_first <= col <= c_diff:
                    cell.number_format = _NUM
            r += 1
        end = r - 1
        mc.cell(r, 1, s["Roll No"])
        mc.cell(r, 2, s["Student Name"])
        mc.cell(r, 3, "Total")
        mc.cell(r, c_sum, f"=SUM({L(c_sum)}{start}:{L(c_sum)}{end})")
        mc.cell(r, c_rec, f"=SUM({L(c_rec)}{start}:{L(c_rec)}{end})")
        mc.cell(r, c_diff, f"={L(c_sum)}{r}-{L(c_rec)}{r}")
        mc.cell(r, c_stat, f'=IF(AND({L(c_diff)}{r}=0,{L(c_box)}{r}={L(c_rec)}{r}),"OK","CHECK")')
        box = clean(s.get("Total Marks"))
        if box is not None:
            mc.cell(r, c_box, box)
        words = str(s.get("Marks in Words", "") or "").strip()
        wn = words_to_number(words)
        remark = f"In words: {words}" if words else "In words: (not read)"
        if wn is not None and box is not None and abs(wn - box) > 1e-9:
            remark += f"  |  WORDS DO NOT MATCH ({wn} vs {fmt(box)})"
        mc.cell(r, c_rem, remark)
        for col in range(1, c_rem + 1):
            cell = mc.cell(r, col)
            cell.font, cell.fill, cell.border = _BOLD, _TOTAL_FILL, _BOX
            cell.alignment = _CENTER if col > 2 else _LEFT
            if c_first <= col <= c_diff or col == c_box:
                cell.number_format = _NUM
        r += 1
    mc.conditional_formatting.add(f"{L(c_stat)}4:{L(c_stat)}{max(r, 5)}",
                                  CellIsRule(operator="equal", formula=['"CHECK"'], fill=_BAD_FILL))
    for col, w in {1: 17, 2: 24, 3: 8}.items():
        mc.column_dimensions[L(col)].width = w
    for col in range(c_first, c_sum):
        mc.column_dimensions[L(col)].width = 8
    for col, w in {c_sum: 12, c_rec: 16, c_diff: 11, c_stat: 10, c_box: 14, c_rem: 46}.items():
        mc.column_dimensions[L(col)].width = w
    mc.freeze_panes = "D4"

    # ---- Sheet 3: Student-wise Marks (one row per student) -----------------
    sw = wb.create_sheet("Student-wise Marks")
    sw["A1"] = "Student-wise Marks (sorted by roll number)"
    sw["A1"].font = Font(name=_F, size=13, bold=True)
    headers = ["Sr No", "Roll No", "Student Name"]
    for q in range(1, n_q + 1):
        headers += [f"Q.No{q}({l})" for l in letters] + [f"Q.No{q} Total"]
    headers += ["Total Marks", "Marks in Words", "Review", "Source"]
    _header(sw, 3, headers)
    for i, (_, s) in enumerate(df.iterrows(), start=1):
        rr = 3 + i
        vals = [i, s["Roll No"], s["Student Name"]]
        for q in range(1, n_q + 1):
            vals += [clean(s.get(f"Q.No{q}({l})")) for l in letters] + [clean(s.get(f"Q.No{q} Total"))]
        vals += [clean(s.get("Total Marks")), s.get("Marks in Words", ""), s["Review"], s.get("Source", "")]
        for j, v in enumerate(vals, start=1):
            cell = sw.cell(rr, j, v)
            cell.font, cell.border = _BODY, _BOX
            cell.alignment = _LEFT if j in (2, 3) or j > len(vals) - 3 else _CENTER
            if isinstance(v, (int, float)) and j > 3:
                cell.number_format = _NUM
    last_row = 3 + max(len(df), 1)
    last_col = len(headers)
    sw.auto_filter.ref = f"A3:{L(last_col)}{last_row}"
    rev_col = L(last_col - 1)
    sw.conditional_formatting.add(f"{rev_col}4:{rev_col}{last_row}",
                                  CellIsRule(operator="notEqual", formula=['"OK"'], fill=_BAD_FILL))
    for j, h in enumerate(headers, start=1):
        sw.column_dimensions[L(j)].width = (
            7 if h == "Sr No" else 17 if h == "Roll No" else 24 if h == "Student Name"
            else 12 if h.endswith("Total") or h == "Total Marks" else 20 if h == "Marks in Words"
            else 50 if h == "Review" else 24 if h == "Source" else 8)
    sw.freeze_panes = "D4"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Streamlit UI
# --------------------------------------------------------------------------- #
def _clean_key(k):
    return str(k).strip().strip("\"'").strip() if k else ""


def get_api_key_with_source(provider):
    """Sidebar box > Streamlit secret > environment variable. Returns (key, source)."""
    name = SECRET_NAMES[provider]
    manual = _clean_key(st.session_state.get(f"manual_key_{provider}"))
    if manual:
        return manual, "sidebar box"
    try:
        secret = _clean_key(st.secrets.get(name))
    except Exception:
        secret = ""
    if secret:
        return secret, "Streamlit secret"
    env = _clean_key(os.environ.get(name))
    if env:
        return env, "environment variable"
    return "", "none"


def key_diagnostics(key, source, provider):
    shown = f"{key[:5]}...{key[-4:]}" if len(key) > 12 else "(too short)"
    lines = [f"Source: {source}", f"Length: {len(key)} characters", f"Looks like: {shown}"]
    if provider == "claude" and not key.startswith("sk-ant-"):
        lines.append("Problem: an Anthropic key should start with sk-ant-")
    if provider == "gemini" and not key.startswith("AIza"):
        lines.append("Note: Gemini keys usually start with AIza - check you copied it from AI Studio")
    if any(ch.isspace() for ch in key):
        lines.append("Problem: the key contains spaces or line breaks")
    if len(key) < 30:
        lines.append("Problem: the key looks too short - it may be cut off")
    return lines


def test_key(provider):
    key, source = get_api_key_with_source(provider)
    if not key:
        st.sidebar.error(f"No key found. Paste one in the box above or set {SECRET_NAMES[provider]} in secrets.")
        return
    for line in key_diagnostics(key, source, provider):
        st.sidebar.caption(line)
    try:
        if provider == "claude":
            import anthropic
            anthropic.Anthropic(api_key=key, max_retries=0).models.list(limit=1)
            st.sidebar.success("Key works.")
        else:
            from google import genai
            names = [m.name.replace("models/", "") for m in genai.Client(api_key=key).models.list()]
            flash = [n for n in names if "flash" in n and "image" not in n and "tts" not in n][:8]
            st.sidebar.success("Key works.")
            if flash:
                st.sidebar.caption("Flash models available to this key: " + ", ".join(flash))
    except Exception as e:
        msg = str(e)
        if any(t in msg for t in ("401", "403", "API key not valid", "API_KEY_INVALID", "authentication")):
            st.sidebar.error("The key was rejected. Create a new key and paste it in the box above to try it.")
        else:
            st.sidebar.warning(f"Could not complete the test: {msg[:200]}")


def get_app_password():
    try:
        pw = st.secrets.get("APP_PASSWORD")
    except Exception:
        pw = None
    return str(pw) if pw else os.environ.get("APP_PASSWORD", "")


def check_password():
    """Show a login form until the correct password is entered. Fails closed."""
    expected = get_app_password()
    if not expected:
        st.error("No app password is configured. Add APP_PASSWORD to the app's secrets "
                 "(Streamlit Cloud: Settings -> Secrets), then reboot the app.")
        st.code('APP_PASSWORD = "choose-a-strong-password"', language="toml")
        return False
    if st.session_state.get("authenticated"):
        return True
    st.title("🔒 Exam Marks Extractor")
    with st.form("login"):
        entered = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in", type="primary")
    if submitted:
        if hmac.compare_digest(entered.encode(), expected.encode()):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            time.sleep(1)  # slows down guessing
            st.error("Incorrect password.")
    return False


def main():
    st.set_page_config(page_title="Exam Marks Extractor", page_icon="📝", layout="wide")
    if not check_password():
        st.stop()
    st.title("📝 Exam Cover Page → Excel")
    st.caption("Upload photos, scans or PDFs of answer-script cover pages. "
               "Review the extracted marks, then download the Excel workbook.")

    with st.sidebar:
        st.header("Settings")
        if st.button("Log out"):
            st.session_state.clear()
            st.rerun()
        provider_label = st.radio("Reading engine", list(PROVIDERS), index=0,
                                  help="Gemini has a free tier (rate-limited). Claude is paid but was built "
                                       "and tested for this layout.")
        provider = PROVIDERS[provider_label]
        st.text_input(f"{'Gemini' if provider == 'gemini' else 'Anthropic'} API key (optional)",
                      type="password", key=f"manual_key_{provider}",
                      help=f"Overrides the app's saved key. Leave empty to use {SECRET_NAMES[provider]} from Streamlit secrets.")
        if st.button("Test API key"):
            test_key(provider)
        if provider == "gemini":
            model = st.selectbox("Model", GEMINI_MODELS, index=0,
                                 help="Flash models are fast and free-tier friendly.")
            custom = st.text_input("Custom model name (optional)",
                                   help="Use this if a model above is not available to your key (see Test API key).")
            model = custom.strip() or model
        else:
            model = st.selectbox("Model", MODELS, index=0,
                                 help="Sonnet is the recommended balance of accuracy and cost. "
                                      "Haiku is cheaper but may misread handwriting more often.")
        n_q = st.number_input("Number of questions on the sheet", 1, 20, 10)
        hide_empty = st.checkbox("Hide part columns that are empty for all students", value=True)
        max_pages = st.number_input("Maximum pages to process", 1, 500, 100)
        st.divider()
        if provider == "gemini":
            st.caption("Privacy: page images are sent to Google's Gemini API. On the free tier Google may use "
                       "submitted data to improve its products. Check your university's examination rules "
                       "before using real scripts.")
        else:
            st.caption("Privacy: page images are sent to the Anthropic API for reading. "
                       "Check that your university's examination rules allow this before using real scripts.")

    files = st.file_uploader("Cover pages (JPG, PNG or PDF - multiple files allowed)",
                             type=["jpg", "jpeg", "png", "pdf"], accept_multiple_files=True)
    st.caption("Tip: lay the page flat in good light and keep the whole page in the frame. "
               "A PDF can hold many cover pages - each page is read separately.")

    if st.button("Extract marks", type="primary", disabled=not files):
        key, _src = get_api_key_with_source(provider)
        if not key:
            st.error(f"Please enter an API key in the sidebar (or set {SECRET_NAMES[provider]} in secrets).")
            st.stop()
        try:
            extract_fn = make_extractor(provider, key, model)
        except ImportError as e:
            st.error(f"A required package is missing: {e}. Check requirements.txt and reboot the app.")
            st.stop()

        with st.spinner("Reading uploaded files..."):
            pages, notes = collect_pages(files, int(max_pages))
        for n in notes:
            st.warning(n)
        if not pages:
            st.error("No readable pages found.")
            st.stop()

        bar = st.progress(0.0, text=f"Reading page 0 of {len(pages)}")
        results = run_extraction(
            extract_fn, pages, WORKERS_BY_PROVIDER[provider],
            lambda i, n: bar.progress(i / n, text=f"Reading page {i} of {n}"))
        bar.empty()

        df, log, meta = results_to_df(pages, results, int(n_q))
        st.session_state.update(df=df, log=log, meta=meta, n_q=int(n_q))

    if "df" not in st.session_state:
        st.info("Upload cover pages above and press **Extract marks**.")
        return

    df, log, meta, n_q = (st.session_state[k] for k in ("df", "log", "meta", "n_q"))
    log_df = pd.DataFrame(log)
    n_err = int((log_df["Status"] == "Error").sum()) if len(log_df) else 0
    n_skip = int((log_df["Status"] == "Skipped").sum()) if len(log_df) else 0

    if df.empty:
        st.error("No cover pages could be read.")
        st.dataframe(log_df, width="stretch")
        return

    review = compute_review(df, n_q)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pages processed", len(log_df))
    c2.metric("Students found", len(df))
    c3.metric("Need review", int((review != "OK").sum()))
    c4.metric("Skipped / errors", f"{n_skip} / {n_err}")

    if n_err or n_skip:
        with st.expander("Pages that were skipped or failed"):
            st.dataframe(log_df[log_df["Status"] != "Read"], width="stretch", hide_index=True)

    st.subheader("Review and correct")
    st.caption("Rows marked CHECK have totals that do not add up or handwriting the model was unsure about. "
               "Compare with the original page and edit any cell. You can also add or delete rows.")
    show = df.copy()
    show.insert(2, "Review", review)
    if hide_empty:
        show = hide_empty_part_columns(show, n_q)
    num_cfg = {c: st.column_config.NumberColumn(c, format="%g")
               for c in show.columns if PART_RE.match(c) or c.endswith(" Total") or c == "Total Marks"}
    edited = st.data_editor(
        show, num_rows="dynamic", width="stretch", hide_index=True,
        column_config=num_cfg, disabled=["Review", "Extraction Notes", "Source"], key="editor")

    edited = edited.drop(columns=["Review"], errors="ignore")
    edited = edited.dropna(how="all")
    xlsx = build_excel(edited, meta, n_q)
    name = re.sub(r"[^A-Za-z0-9]+", "_", meta.get("subject", "") or "Marks").strip("_")
    st.download_button("⬇️ Download Excel", xlsx,
                       file_name=f"{name}_marks_{datetime.now():%Y%m%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    st.caption("The workbook has three tabs: Script Details, Marks Check and Student-wise Marks "
               "(one row per student, sorted by roll number).")


def _has_secrets():
    try:
        return "ANTHROPIC_API_KEY" in st.secrets
    except Exception:
        return False


if __name__ == "__main__":
    main()
