import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}\*?$")
LINE_TWO_DATES = re.compile(r"\d{4}-\d{2}-\d{2}\*?\s+\d{4}-\d{2}-\d{2}\*?$")

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

    for i in range(total_pages):
        text = doc.load_page(i).get_text("text") or ""
        if not text.strip():
            continue
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

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
            if l.startswith("activity id") or l.startswith("month") or l.startswith("page "):
                continue
            if ln.startswith("-1 "):
                continue
            if not LINE_TWO_DATES.search(ln):
                continue

            parts = ln.split()
            if len(parts) < 4:
                continue

            start_tok, finish_tok = parts[-2], parts[-1]
            if not (DATE_RE.match(start_tok) and DATE_RE.match(finish_tok)):
                continue

            act_id = parts[0]
            name = " ".join(parts[1:-2]).strip()

            start_d, start_star = normalize_date(start_tok)
            finish_d, finish_star = normalize_date(finish_tok)
            if start_d is None or finish_d is None:
                continue

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
                "start": start_d,
                "finish": finish_d,
                "duration_days": (finish_d - start_d).days,
                "is_milestone": start_d == finish_d,
                "source_page": i + 1,
                "pdf_pages": total_pages,
            })

    df = pd.DataFrame(rows).drop_duplicates(subset=["activity_id","activity_name","start","finish"])
    df = df.sort_values(["package_code","start","finish","activity_id"], na_position="last").reset_index(drop=True)

    Path(out_csv).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out_csv, index=False)
    print(f"Saved {len(df)} rows → {out_csv}")

if __name__ == "__main__":
    extract("ProjectSchedule.pdf", "data/primavera.csv")
