import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

# Accept 2026-02-15 or 2026-02-15*
DATE_TOKEN = re.compile(r"\d{4}-\d{2}-\d{2}\*?")
# Line ending with two date tokens
LINE_ENDS_WITH_TWO_DATES = re.compile(r"(\d{4}-\d{2}-\d{2}\*?)\s+(\d{4}-\d{2}-\d{2}\*?)\s*$")

WORKTYPE_RULES = [
    ("Pile diagram", r"\bpile\s+diagram\b"),
    ("Issue pile drawing", r"\bissue\s+pile\s+drawing\b"),
    ("Issue DED drawing", r"\bissue\s+ded\s+drawing\b"),
    ("Other professional drawings", r"\bother\s+professional\s+drawings\b"),
    ("Professional drawings", r"\bprofessional\s+drawings\b"),
]

def normalize_date(token: str):
    has_star = token.endswith("*")
    token_clean = token.replace("*", "")
    try:
        return dtparse(token_clean).date(), has_star
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
    # S00, A10, U32, E01, M06 etc.
    return bool(re.match(r"^[A-Z]\d{2,3}$", act_id or ""))

def extract(pdf_path: str, out_csv: str):
    rows = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    debug_lines_with_dates = 0
    debug_first_lines = []

    for i in range(total_pages):
        text = doc.load_page(i).get_text("text") or ""
        if not text.strip():
            continue

        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

        # keep some debug samples from early pages
        if len(debug_first_lines) < 20:
            debug_first_lines.extend(lines[:5])

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
            # skip obvious headers/footers
            if l.startswith("activity id") or l.startswith("month") or l.startswith("page "):
                continue
            if ln.startswith("-1 "):
                continue

            m = LINE_ENDS_WITH_TWO_DATES.search(ln)
            if not m:
                continue

            debug_lines_with_dates += 1

            start_tok = m.group(1)
            finish_tok = m.group(2)

            # split safely: remove the two date tokens from the end
            # Example line: "DD1050 Pile diagram 2025-02-03 2025-03-10"
            # We'll remove last occurrence of those tokens.
            parts = ln.split()
            if len(parts) < 4:
                continue

            # last two tokens should match extracted date tokens (tolerant)
            if not (DATE_TOKEN.fullmatch(parts[-2]) and DATE_TOKEN.fullmatch(parts[-1])):
                continue

            act_id = parts[0]
            name = " ".join(parts[1:-2]).strip()

            start_d, start_star = normalize_date(parts[-2])
            finish_d, finish_star = normalize_date(parts[-1])
            if start_d is None or finish_d is None:
                continue

            # update package context (if this line is a package summary line)
            if is_package_code(act_id):
                current_package_code = act_id
                current_package_name = name

            rows.append({
                "major_group": current_major_group or "Unknown",
                "package_code": current_package_code,      # can be None (still OK)
                "package_name": current_package_name,      # can be None
                "activity_id": act_id,
                "activity_name": name,
                "work_type": infer_work_type(name),
                "start": start_d,
                "finish": finish_d,
                "duration_days": (finish_d - start_d).days,
                "is_milestone": start_d == finish_d,
                "source_page": i + 1,
                "pdf_pages": total_pages,
                "start_star": start_star,
                "finish_star": finish_star,
            })

    df = pd.DataFrame(rows)

    # ---- DEBUG output in Actions logs ----
    print(f"[DEBUG] PDF pages: {total_pages}")
    print(f"[DEBUG] Lines that look like activities (end with two dates): {debug_lines_with_dates}")
    print("[DEBUG] Sample lines from PDF:")
    for s in debug_first_lines[:15]:
        print("   ", s)

    # If nothing extracted, still write a CSV with headers so the app doesn't break
    headers = [
        "major_group","package_code","package_name","activity_id","activity_name","work_type",
        "start","finish","duration_days","is_milestone","source_page","pdf_pages","start_star","finish_star"
    ]
    if df.empty:
        print("[WARN] No activities extracted. Writing empty CSV with headers.")
        Path(out_csv).parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(columns=headers).to_csv(out_csv, index=False)
        return

    # Clean + sort safely (package_code may be missing sometimes)
    df = df.drop_duplicates(subset=["activity_id","activity_name","start","finish"])
    sort_cols = ["package_code","start","finish","activity_id"]
    for c in sort_cols:
        if c not in df.columns:
            df[c] = None
    df = df.sort_values(sort_cols, na_position="last").reset_index(drop=True)

    Path(out_csv).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out_csv, index=False)
    print(f"Saved {len(df)} rows → {out_csv}")

if __name__ == "__main__":
    extract("ProjectSchedule.pdf", "data/primavera.csv")
