import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

# Hyphen variants that appear in PDFs
HYPHENS = r"[-\u2010\u2011\u2012\u2013\u2014\u2212\uFE63\uFF0D]"
DATE_WORD_RE = re.compile(rf"^\d{{4}}{HYPHENS}\d{{2}}{HYPHENS}\d{{2}}\*?$")

# Looser Activity ID: DD1050 / MS1010 / PR1100 / ABC1234 / etc.
ACT_ID_RE = re.compile(r"^[A-Z]{1,5}\d{2,7}$", re.IGNORECASE)
PKG_RE = re.compile(r"^[A-Z]\d{2,3}$", re.IGNORECASE)  # S00, A10, U32...

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

def parse_date_word(w: str):
    w = (w or "").strip()
    star = w.endswith("*")
    w = w.replace("*", "")
    w = normalize_hyphens(w)
    d = dtparse(w).date()
    return d, star

def infer_work_type(name: str) -> str:
    s = (name or "").lower()
    for label, pat in WORKTYPE_RULES:
        if re.search(pat, s):
            return label
    if (name or "").upper().startswith("MS"):
        return "Milestone"
    return "Other"

def is_package_code(tok: str) -> bool:
    return bool(PKG_RE.match(tok or ""))

def looks_like_activity_id(tok: str) -> bool:
    return bool(ACT_ID_RE.match(tok or "")) and not DATE_WORD_RE.match(tok or "")

def group_words_by_y(words, y_tol=0.6):
    """
    Group word tuples into 'rows' by similar y0.
    words tuples: (x0, y0, x1, y1, text, block_no, line_no, word_no)
    """
    words_sorted = sorted(words, key=lambda w: (w[1], w[0]))  # y0, x0
    groups = []
    cur = []
    cur_y = None

    for w in words_sorted:
        x0, y0, x1, y1, txt = w[0], w[1], w[2], w[3], str(w[4]).strip()
        if not txt:
            continue
        if cur_y is None:
            cur_y = y0
            cur = [w]
        else:
            if abs(y0 - cur_y) <= y_tol:
                cur.append(w)
            else:
                groups.append(sorted(cur, key=lambda z: z[0]))
                cur_y = y0
                cur = [w]

    if cur:
        groups.append(sorted(cur, key=lambda z: z[0]))
    return groups

def extract(pdf_path: str, out_csv: str):
    rows = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    debug_rows_seen = 0
    debug_samples = []

    for page_i in range(total_pages):
        page = doc.load_page(page_i)

        # Best effort major group detection from plain text headings
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

        words = page.get_text("words")
        if not words:
            continue

        # Tight row grouping to avoid merging multiple lines
        line_groups = group_words_by_y(words, y_tol=0.6)

        for g in line_groups:
            tokens = [str(w[4]).strip() for w in g if str(w[4]).strip()]
            if len(tokens) < 4:
                continue

            # skip obvious headers
            head = " ".join(tokens[:3]).lower()
            if head.startswith("activity id") or head.startswith("activityid"):
                continue
            if tokens[0].lower() in ("month", "page"):
                continue

            # Find all date tokens
            date_idxs = [i for i, t in enumerate(tokens) if DATE_WORD_RE.match(t)]
            if len(date_idxs) < 2:
                continue

            # Debug: we found a row-like line with dates
            debug_rows_seen += 1
            if len(debug_samples) < 12:
                debug_samples.append(" | ".join(tokens))

            # Choose start/finish as last two date tokens
            s_idx, f_idx = date_idxs[-2], date_idxs[-1]
            start_tok, finish_tok = tokens[s_idx], tokens[f_idx]

            try:
                start_d, start_star = parse_date_word(start_tok)
                finish_d, finish_star = parse_date_word(finish_tok)
            except Exception:
                continue

            # Find an activity id token BEFORE the start date (usually first token)
            act_id = None
            act_pos = None
            for i in range(0, s_idx):
                if looks_like_activity_id(tokens[i]):
                    act_id = tokens[i]
                    act_pos = i
                    break

            # If we can't find an activity id, skip (avoids garbage rows)
            if not act_id:
                continue

            # Activity name is between act_id and start date
            name_tokens = tokens[act_pos + 1 : s_idx]
            name = " ".join(name_tokens).strip()
            if not name:
                continue

            # Update package context if this looks like package summary row
            if is_package_code(act_id):
                current_package_code = act_id.upper()
                current_package_name = name

            duration_days = (finish_d - start_d).days
            if duration_days < 0:
                continue

            rows.append({
                "major_group": current_major_group or "Unknown",
                "package_code": current_package_code,
                "package_name": current_package_name,
                "activity_id": act_id.upper(),
                "activity_name": name,
                "work_type": infer_work_type(name),
                "start": start_d.isoformat(),
                "finish": finish_d.isoformat(),
                "duration_days": duration_days,
                "is_milestone": start_d == finish_d,
                "source_page": page_i + 1,
                "pdf_pages": total_pages,
                "start_star": bool(start_star),
                "finish_star": bool(finish_star),
            })

    print(f"[DEBUG] PDF pages: {total_pages}")
    print(f"[DEBUG] Lines with >=2 dates detected (pre-filter): {debug_rows_seen}")
    print(f"[DEBUG] Extracted final rows: {len(rows)}")
    print("[DEBUG] Sample date-lines (tokenized):")
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
