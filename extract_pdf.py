import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

# Accept normal hyphen '-' and common unicode hyphens in PDFs
HYPHENS = r"[-\u2010\u2011\u2012\u2013\u2014\u2212\uFE63\uFF0D]"
DATE_IN_TEXT = re.compile(rf"(\d{{4}}){HYPHENS}(\d{{2}}){HYPHENS}(\d{{2}})\*?")

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

def infer_work_type(name: str) -> str:
    s = (name or "").lower()
    for label, pat in WORKTYPE_RULES:
        if re.search(pat, s):
            return label
    if (name or "").startswith("MS"):
        return "Milestone"
    return "Other"

def is_package_code(act_id: str) -> bool:
    return bool(re.match(r"^[A-Z]\d{2,3}$", act_id or ""))  # S01, A00, U32, etc.

def normalize_date_str(date_str: str) -> str:
    # Convert unicode hyphens to normal '-' and strip trailing '*'
    clean = date_str.replace("*", "")
    clean = re.sub(HYPHENS, "-", clean)
    # Parse to ISO to guarantee CSV consistency
    return dtparse(clean).date().isoformat()

def extract(pdf_path: str, out_csv: str):
    rows = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    debug_activity_lines = []
    extracted_lines_count = 0

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
            if l.startswith("activity id") or l.startswith("month") or l.startswith("page "):
                continue
            if ln.startswith("-1 "):
                continue

            # Find all date occurrences in the line (supports unicode hyphens)
            dates = [m.group(0) for m in DATE_IN_TEXT.finditer(ln)]
            if len(dates) < 2:
                continue

            # Take last two as Start/Finish
            start_raw, finish_raw = dates[-2], dates[-1]
            try:
                start_iso = normalize_date_str(start_raw)
                finish_iso = normalize_date_str(finish_raw)
            except Exception:
                continue

            # Activity ID is typically first token
            parts = ln.split()
            if len(parts) < 2:
                continue
            act_id = parts[0]

            # Activity name = line with dates removed, minus activity id
            # Remove the last two dates from the line text
            ln_wo_finish = re.sub(re.escape(finish_raw), "", ln, count=1)
            ln_wo_both = re.sub(re.escape(start_raw), "", ln_wo_finish, count=1)
            # Now remove the activity id from the beginning
            name = ln_wo_both.strip()
            if name.startswith(act_id):
                name = name[len(act_id):].strip()

            if not name:
                continue

            extracted_lines_count += 1
            if len(debug_activity_lines) < 12:
                debug_activity_lines.append(ln)

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
                "start": start_iso,
                "finish": finish_iso,
                "source_page": page_i + 1,
                "pdf_pages": total_pages,
            })

    print(f"[DEBUG] PDF pages: {total_pages}")
    print(f"[DEBUG] Extracted candidate lines: {extracted_lines_count}")
    print("[DEBUG] Sample extracted lines:")
    for s in debug_activity_lines:
        print("   ", s)

    # Always write a CSV (even if empty) but now it should not be empty
    headers = [
        "major_group","package_code","package_name","activity_id","activity_name","work_type",
        "start","finish","source_page","pdf_pages"
    ]
    df = pd.DataFrame(rows)

    Path(out_csv).parent.mkdir(parents=True, exist_ok=True)

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
