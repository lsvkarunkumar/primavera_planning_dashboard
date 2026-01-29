import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

# Hyphen variants that appear in PDFs
HYPHENS = r"[-\u2010\u2011\u2012\u2013\u2014\u2212\uFE63\uFF0D]"
DATE_WORD_RE = re.compile(rf"^\d{{4}}{HYPHENS}\d{{2}}{HYPHENS}\d{{2}}\*?$")

# Activity IDs are usually letters+digits (DD1050 / MS1010 / PR1100 etc.)
ACT_ID_RE = re.compile(r"^[A-Z]{1,6}\d{2,7}$", re.IGNORECASE)
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
    tok = (tok or "").strip()
    return bool(ACT_ID_RE.match(tok)) and not bool(DATE_WORD_RE.match(tok))

def extract(pdf_path: str, out_csv: str):
    rows = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    debug_fragments = 0
    debug_merged_rows = 0
    debug_final_rows = 0
    debug_samples = []

    for page_i in range(total_pages):
        page = doc.load_page(page_i)

        # Best-effort major group detection from headings
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

        words = page.get_text("words")  # (x0,y0,x1,y1,word,block,line,wordno)
        if not words:
            continue

        # 1) Build "line fragments" by (block_no, line_no)
        frag_map = {}
        for w in words:
            x0, y0, x1, y1, txt, block_no, line_no, word_no = w
            txt = str(txt).strip()
            if not txt:
                continue
            key = (block_no, line_no)
            frag_map.setdefault(key, []).append(w)

        fragments = []
        for key, ws in frag_map.items():
            ws_sorted = sorted(ws, key=lambda z: z[0])  # sort by x0
            tokens = [str(z[4]).strip() for z in ws_sorted if str(z[4]).strip()]
            if not tokens:
                continue
            x0 = min(z[0] for z in ws_sorted)
            x1 = max(z[2] for z in ws_sorted)
            y0 = min(z[1] for z in ws_sorted)
            y1 = max(z[3] for z in ws_sorted)
            ymid = (y0 + y1) / 2.0
            fragments.append({
                "x0": x0, "x1": x1, "y0": y0, "y1": y1, "ymid": ymid,
                "tokens": tokens
            })

        debug_fragments += len(fragments)

        # 2) Cluster fragments into logical rows by Y (merge columns)
        # Sort by ymid then x0
        fragments.sort(key=lambda f: (f["ymid"], f["x0"]))

        row_clusters = []
        cur = []
        cur_y = None
        Y_TOL = 2.0  # key change: merge left+right columns of same row

        for frag in fragments:
            if cur_y is None:
                cur = [frag]
                cur_y = frag["ymid"]
            else:
                if abs(frag["ymid"] - cur_y) <= Y_TOL:
                    cur.append(frag)
                else:
                    row_clusters.append(cur)
                    cur = [frag]
                    cur_y = frag["ymid"]
        if cur:
            row_clusters.append(cur)

        # 3) Parse each merged row
        for row in row_clusters:
            # Sort fragments left-to-right and flatten tokens
            row = sorted(row, key=lambda f: f["x0"])
            tokens = []
            for f in row:
                tokens.extend(f["tokens"])

            if len(tokens) < 4:
                continue

            # Skip obvious headers
            head = " ".join(tokens[:3]).lower()
            if head.startswith("activity id") or head.startswith("activityid"):
                continue
            if tokens[0].lower() in ("month", "page"):
                continue

            # Find date tokens anywhere in merged tokens
            date_idxs = [i for i, t in enumerate(tokens) if DATE_WORD_RE.match(t)]
            if len(date_idxs) < 2:
                continue

            debug_merged_rows += 1
            if len(debug_samples) < 10:
                debug_samples.append(" | ".join(tokens[:25]))  # keep short sample

            s_idx, f_idx = date_idxs[-2], date_idxs[-1]
            try:
                start_d, start_star = parse_date_word(tokens[s_idx])
                finish_d, finish_star = parse_date_word(tokens[f_idx])
            except Exception:
                continue

            # Find activity id token before the start date
            act_id = None
            act_pos = None
            for i in range(0, s_idx):
                if looks_like_activity_id(tokens[i]):
                    act_id = tokens[i].upper()
                    act_pos = i
                    break
            if not act_id:
                continue

            # Name is between activity id and start date
            name = " ".join(tokens[act_pos + 1 : s_idx]).strip()
            if not name:
                continue

            # Update package context (package row)
            if is_package_code(act_id):
                current_package_code = act_id
                current_package_name = name

            duration_days = (finish_d - start_d).days
            if duration_days < 0:
                continue

            debug_final_rows += 1
            rows.append({
                "major_group": current_major_group or "Unknown",
                "package_code": current_package_code,
                "package_name": current_package_name,
                "activity_id": act_id,
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
    print(f"[DEBUG] Total line-fragments built: {debug_fragments}")
    print(f"[DEBUG] Merged rows with >=2 dates: {debug_merged_rows}")
    print(f"[DEBUG] Final extracted rows saved: {debug_final_rows}")
    print("[DEBUG] Sample merged token rows (first 25 tokens each):")
    for s in debug_samples:
        print("   ", s)

    headers = [
        "major_group","package_code","package_name","activity_id","activity_name","work_type",
        "start","finish","duration_days","is_milestone","source_page","pdf_pages","start_star","finish_star"
    ]
    Path(out_csv).parent.mkdir(parents=True, exist_ok=True)
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
