import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

HYPHENS = r"[-\u2010\u2011\u2012\u2013\u2014\u2212\uFE63\uFF0D]"
DATE_WORD_RE = re.compile(rf"^\d{{4}}{HYPHENS}\d{{2}}{HYPHENS}\d{{2}}\*?$")

# Allow many Primavera-style IDs: DD1050, MS1010, PR1100, A00, S01, etc.
ACT_ID_RE = re.compile(r"^[A-Z]{1,6}\d{2,7}$", re.IGNORECASE)
PKG_RE = re.compile(r"^[A-Z]\d{2,3}$", re.IGNORECASE)

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

def looks_like_activity_id(tok: str) -> bool:
    tok = (tok or "").strip()
    return bool(ACT_ID_RE.match(tok)) and not bool(DATE_WORD_RE.match(tok))

def is_package_code(tok: str) -> bool:
    return bool(PKG_RE.match(tok or ""))

def build_lines(words):
    """Group by (block_no, line_no) to get stable fragments."""
    line_map = {}
    for w in words:
        x0, y0, x1, y1, txt, block_no, line_no, word_no = w
        txt = str(txt).strip()
        if not txt:
            continue
        key = (block_no, line_no)
        line_map.setdefault(key, []).append(w)

    lines = []
    for key, ws in line_map.items():
        ws_sorted = sorted(ws, key=lambda z: z[0])  # x0
        tokens = [str(z[4]).strip() for z in ws_sorted if str(z[4]).strip()]
        if not tokens:
            continue
        x0 = min(z[0] for z in ws_sorted)
        x1 = max(z[2] for z in ws_sorted)
        y0 = min(z[1] for z in ws_sorted)
        y1 = max(z[3] for z in ws_sorted)
        ymid = (y0 + y1) / 2.0
        lines.append({"x0": x0, "x1": x1, "y0": y0, "y1": y1, "ymid": ymid, "tokens": tokens})
    return lines

def extract(pdf_path: str, out_csv: str):
    rows = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    debug_left = 0
    debug_right = 0
    debug_joined = 0
    debug_samples = []

    for page_i in range(total_pages):
        page = doc.load_page(page_i)

        # Major group detection (best effort)
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

        page_width = page.rect.width
        split_x = page_width * 0.50  # left vs right split

        lines = build_lines(words)

        # Separate left and right line fragments by x0
        left_lines = [ln for ln in lines if ln["x0"] < split_x]
        right_lines = [ln for ln in lines if ln["x0"] >= split_x]

        # LEFT: keep only lines that have an activity id-like token
        left_rows = []
        for ln in left_lines:
            toks = ln["tokens"]
            # skip headers
            head = " ".join(toks[:3]).lower() if len(toks) >= 3 else " ".join(toks).lower()
            if head.startswith("activity id") or head.startswith("activityid"):
                continue
            if toks and toks[0].lower() in ("month", "page"):
                continue

            # find first activity id token in the line
            act_id = None
            act_pos = None
            for i, t in enumerate(toks[:6]):  # usually early
                if looks_like_activity_id(t):
                    act_id = t.upper()
                    act_pos = i
                    break
            if not act_id:
                continue

            name = " ".join(toks[act_pos + 1:]).strip()
            if not name:
                name = ""  # allow empty, we still want to match with dates

            left_rows.append({
                "ymid": ln["ymid"],
                "activity_id": act_id,
                "activity_name": name,
                "raw": " | ".join(toks[:25]),
            })
        left_rows.sort(key=lambda r: r["ymid"])
        debug_left += len(left_rows)

        # RIGHT: keep only lines with >=2 date tokens
        right_rows = []
        for ln in right_lines:
            toks = ln["tokens"]
            date_idxs = [i for i, t in enumerate(toks) if DATE_WORD_RE.match(t)]
            if len(date_idxs) < 2:
                continue
            s_idx, f_idx = date_idxs[-2], date_idxs[-1]
            right_rows.append({
                "ymid": ln["ymid"],
                "start_tok": toks[s_idx],
                "finish_tok": toks[f_idx],
                "raw": " | ".join(toks[:25]),
            })
        right_rows.sort(key=lambda r: r["ymid"])
        debug_right += len(right_rows)

        if not left_rows or not right_rows:
            continue

        # Helper: find nearest left row by y
        # (left rows are sorted by ymid)
        def nearest_left(y):
            # binary-ish scan: linear is ok (counts per page small), keep simple
            best = None
            best_d = 1e9
            for r in left_rows:
                d = abs(r["ymid"] - y)
                if d < best_d:
                    best_d = d
                    best = r
            return best, best_d

        # Join right date rows to nearest left id/name row
        Y_JOIN_MAX = 6.0  # allow a bit of offset
        for rr in right_rows:
            lr, dist = nearest_left(rr["ymid"])
            if not lr or dist > Y_JOIN_MAX:
                continue

            try:
                start_d, start_star = parse_date_word(rr["start_tok"])
                finish_d, finish_star = parse_date_word(rr["finish_tok"])
            except Exception:
                continue

            duration_days = (finish_d - start_d).days
            if duration_days < 0:
                continue

            act_id = lr["activity_id"]
            name = lr["activity_name"]

            # package context update
            if is_package_code(act_id):
                current_package_code = act_id
                current_package_name = name

            debug_joined += 1
            if len(debug_samples) < 10:
                debug_samples.append(
                    f"JOIN ydist={dist:.2f} | L: {lr['raw']} || R: {rr['raw']}"
                )

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
    print(f"[DEBUG] Left ID/name rows found: {debug_left}")
    print(f"[DEBUG] Right date rows found: {debug_right}")
    print(f"[DEBUG] Joined rows produced: {debug_joined}")
    print("[DEBUG] Sample joins:")
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

    df = df.drop_duplicates(subset=["activity_id","start","finish"])
    df = df.sort_values(["package_code","start","finish","activity_id"], na_position="last").reset_index(drop=True)
    df.to_csv(out_csv, index=False)
    print(f"[OK] Saved {len(df)} rows → {out_csv}")

if __name__ == "__main__":
    extract("ProjectSchedule.pdf", "data/primavera.csv")
