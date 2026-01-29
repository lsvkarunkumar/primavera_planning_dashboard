import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

DATE_TOKEN = re.compile(r"^\d{4}-\d{2}-\d{2}\*?$")

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

def normalize_date(token: str):
    token = (token or "").strip()
    has_star = token.endswith("*")
    token_clean = token.replace("*", "")
    try:
        d = dtparse(token_clean).date()
        return d, has_star
    except Exception:
        return None, has_star

def infer_work_type(name: str) -> str:
    s = (name or "").lower()
    for label, pat in WORKTYPE_RULES:
        if re.search(pat, s):
            return label
    if (name or "").startswith("MS"):
        return "Milestone"
    return "Other"

def is_package_code(act_id: str) -> bool:
    # S00, A10, U32, E01, M06, etc.
    return bool(re.match(r"^[A-Z]\d{2,3}$", act_id or ""))

def find_last_two_dates(parts):
    """
    Find last two date-like tokens anywhere in the split tokens.
    Returns (start_token, finish_token, start_index, finish_index) or (None,...)
    """
    idxs = [i for i, p in enumerate(parts) if DATE_TOKEN.match(p)]
    if len(idxs) < 2:
        return None, None, None, None
    s_idx, f_idx = idxs[-2], idxs[-1]
    return parts[s_idx], parts[f_idx], s_idx, f_idx

def extract(pdf_path: str, out_csv: str):
    rows = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    debug_found_lines = 0
    debug_sample_activity_lines = []

    for page_i in range(total_pages):
        text = doc.load_page(page_i).get_text("text") or ""
        if not text.strip():
            continue

        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

        # Major group detection
        for ln in lines:
            l = ln.lower()
            if l.startswith("detailed en"):
                current_major_group = "Detailed Engineering Design"
            elif l.startswith("procurement"):
                current_major_group = "Procurement"
            elif l.startswith("employer revi"):
                current_major_group = "Employer Review and Approval"
            elif l.startswith("main mile"):
                current_major_group = "Main Milestones"

        for ln in lines:
            l = ln.lower()
            # skip headers/axes/footers
            if l.startswith("activity id") or l.startswith("month") or l.startswith("page "):
                continue
            if ln.startswith("-1 "):
                continue

            parts = ln.split()
            if len(parts) < 4:
                continue

            start_tok, finish_tok, s_idx, f_idx = find_last_two_dates(parts)
            if start_tok is None:
                continue

            # activity id is usually first token
            act_id = parts[0]

            # activity name is everything between act_id and the start date token
            name_tokens = parts[1:s_idx]
            name = " ".join(name_tokens).strip()

            start_d, start_star = normalize_date(start_tok)
            finish_d, finish_star = normalize_date(finish_tok)
            if start_d is None or finish_d is None:
                continue

            debug_found_lines += 1
            if len(debug_sample_activity_lines) < 12:
                debug_sample_activity_lines.append(ln)

            # update package context if this is a package summary line
            if is_package_code(act_id):
                current_package_code = act_id
                current_package_name = name

            rows.append({
                "major_group": current_major_group or "Unknown",
                "package_code": current_package_code,
                "package_name": current_package_name,
                "activity_id": act_id,
                "activity_name": name,
                "work_type": infer_work_type(name),
                # write dates as ISO strings (most reliable for Streamlit parsing)
                "start": start_d.isoformat(),
                "finish": finish_d.isoformat(),
                "duration_days": (finish_d - start_d).days,
                "is_milestone": start_d == finish_d,
                "source_page": page_i + 1,
                "pdf_pages": total_pages,
                "start_star": bool(start_star),
                "finish_star": bool(finish_star),
            })

    print(f"[DEBUG] PDF pages: {total_pages}")
    print(f"[DEBUG] Extracted activity-like lines: {debug_found_lines}")
    print("[DEBUG] Sample extracted lines:")
    for s in debug_sample_activity_lines:
        print("   ", s)

    headers = [
        "major_group","package_code","package_name","activity_id","activity_name","work_type",
        "start","finish","duration_days","is_milestone","source_page","pdf_pages","start_star","finish_star"
    ]

    df = pd.DataFrame(rows)
    Path(out_csv).parent.mkdir(parents=True, exist_ok=True)

    if df.empty:
        print("[WARN] No rows extracted. Writing empty CSV with headers.")
        pd.DataFrame(columns=headers).to_csv(out_csv, index=False)
        return

    df = df.drop_duplicates(subset=["activity_id","activity_name","start","finish"])
    # safe sort (package_code can be None)
    df = df.sort_values(["package_code","start","finish","activity_id"], na_position="last").reset_index(drop=True)

    df.to_csv(out_csv, index=False)
    print(f"[OK] Saved {len(df)} rows → {out_csv}")

if __name__ == "__main__":
    extract("ProjectSchedule.pdf", "data/primavera.csv")
