import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

# Hyphen variants that appear in PDFs
HYPHENS = r"[-\u2010\u2011\u2012\u2013\u2014\u2212\uFE63\uFF0D]"
DATE_WORD_RE = re.compile(rf"^\d{{4}}{HYPHENS}\d{{2}}{HYPHENS}\d{{2}}\*?$")

WORKTYPE_RULES = [
    ("Pile diagram", r"\bpile\s+diagram\b"),
    ("Issue pile drawing", r"\bissue\s+pile\s+drawing\b"),
    ("Issue DED drawing", r"\bissue\s+ded\s+drawing\b"),
    ("Issue final DED drawing", r"\bissue\s+final\s+ded\s+drawing\b"),
    ("Other professional drawings", r"\bother\s+professional\s+drawings\b"),
    ("Professional drawings", r"\bprofessional\s+drawings\b"),
    ("General layout", r"\bgeneral\s+layout\b"),
    ("Bidding", r"\bbidding\b"),
    ("Manufacturing", r"\bmanufacturing\b"),
    ("Shipping", r"\bshipping\b"),
]

def normalize_hyphens(s: str) -> str:
    return re.sub(HYPHENS, "-", s)

def normalize_date_word(w: str) -> str:
    w = (w or "").strip()
    has_star = w.endswith("*")
    w = w.replace("*", "")
    w = normalize_hyphens(w)
    # parse to ISO (YYYY-MM-DD)
    d = dtparse(w).date().isoformat()
    return d, has_star

def infer_work_type(name: str) -> str:
    s = (name or "").lower()
    for label, pat in WORKTYPE_RULES:
        if re.search(pat, s):
            return label
    if (name or "").startswith("MS"):
        return "Milestone"
    return "Other"

def is_package_code(act_id: str) -> bool:
    # S00, A10, U32, E01, M06 etc.
    return bool(re.match(r"^[A-Z]\d{2,3}$", act_id or ""))

def group_words_into_lines(words, y_tol=2.0):
    """
    words: list of (x0, y0, x1, y1, text, block_no, line_no, word_no)
    returns list of lines, each line is list of words sorted by x0
    """
    # sort by y then x
    words_sorted = sorted(words, key=lambda w: (w[1], w[0]))
    lines = []
    current = []
    current_y = None

    for w in words_sorted:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        if current_y is None:
            current_y = y0
            current = [w]
        else:
            if abs(y0 - current_y) <= y_tol:
                current.append(w)
            else:
                # finalize line
                current = sorted(current, key=lambda z: z[0])
                lines.append(current)
                # new line
                current_y = y0
                current = [w]

    if current:
        current = sorted(current, key=lambda z: z[0])
        lines.append(current)

    return lines

def extract(pdf_path: str, out_csv: str):
    rows = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    extracted_lines = 0
    debug_samples = []

    for page_i in range(total_pages):
        page = doc.load_page(page_i)

        # WORDS mode preserves columnar text far better than plain text
        words = page.get_text("words")  # list of tuples
        if not words:
            continue

        # optional: detect major group using plain text (best effort)
        text = page.get_text("text") or ""
        for ln in (text.splitlines() if text else []):
            l = ln.strip().lower()
            if l.startswith("detailed en"):
                current_major_group = "Detailed Engineering Design"
            elif l.startswith("procurement"):
                current_major_group = "Procurement"
            elif l.startswith("employer revi"):
                current_major_group = "Employer Review and Approval"
            elif l.startswith("main mile"):
                current_major_group = "Main Milestones"

        lines = group_words_into_lines(words, y_tol=2.0)

        for line in lines:
            tokens = [t[4].strip() for t in line if t[4].strip()]
            if len(tokens) < 4:
                continue

            # skip obvious headers
            head = " ".join(tokens[:3]).lower()
            if head.startswith("activity id") or head.startswith("activityid"):
                continue
            if tokens[0].lower() in ("month", "page"):
                continue

            # find date-like words
            date_idxs = []
            for idx, tok in enumerate(tokens):
                if DATE_WORD_RE.match(tok):
                    date_idxs.append(idx)

            if len(date_idxs) < 2:
                continue

            s_idx, f_idx = date_idxs[-2], date_idxs[-1]
            start_tok, finish_tok = tokens[s_idx], tokens[f_idx]

            try:
                start_iso, start_star = normalize_date_word(start_tok)
                finish_iso, finish_star = normalize_date_word(finish_tok)
            except Exception:
                continue

            act_id = tokens[0]
            if len(act_id) < 2:
                continue

            # activity name tokens are between act_id and start date
            name_tokens = tokens[1:s_idx]
            name = " ".join(name_tokens).strip()
            if not name:
                continue

            extracted_lines += 1
            if len(debug_samples) < 10:
                debug_samples.append(" | ".join(tokens))

            # package context update if package summary line
            if is_package_code(act_id):
                current_package_code = act_id
                current_package_name = name

            # duration
            try:
                dur_days = (dtparse(finish_iso).date() - dtparse(start_iso).date()).days
            except Exception:
                dur_days = None

            rows.append({
                "major_group": current_major_group or "Unknown",
                "package_code": current_package_code,
                "package_name": current_package_name,
                "activity_id": act_id,
                "activity_name": name,
                "work_type": infer_work_type(name),
                "start": start_iso,
                "finish": finish_iso,
                "duration_days": dur_days,
                "is_milestone": start_iso == finish_iso,
                "source_page": page_i + 1,
                "pdf_pages": total_pages,
                "start_star": bool(start_star),
                "finish_star": bool(finish_star),
            })

    print(f"[DEBUG] PDF pages: {total_pages}")
    print(f"[DEBUG] Extracted rows: {extracted_lines}")
    print("[DEBUG] Sample reconstructed rows (tokenized):")
    for s in debug_samples:
        print("   ", s)

    headers = [
        "major_group","package_code","package_name","activity_id","activity_name","work_type",
        "start","finish","duration_days","is_milestone","source_page","pdf_pages","start_star","finish_star"
    ]
    out_path = Path(out_csv)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(rows)
    if df.empty:
        print("[WARN] No rows extracted. Writing empty CSV with headers.")
        pd.DataFrame(columns=headers).to_csv(out_csv, index=False)
        return

    df = df.drop_duplicates(subset=["activity_id","activity_name","start","finish"])
    df = df.sort_values(["package_code","start","finish","activity_id"], na_position="last").reset_index(drop=True)
    df.to_csv(out_csv, index=False)
    print(f"[OK] Saved {len(df)} rows → {out_csv}")

if __name__ == "__main__":
    extract("ProjectSchedule.pdf", "data/primavera.csv")
