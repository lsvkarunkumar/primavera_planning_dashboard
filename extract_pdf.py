import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

# Unicode hyphens inside tokens
HYPHEN_CLASS = r"[-\u2010\u2011\u2012\u2013\u2014\u2212\uFE63\uFF0D]"
FULL_DATE_RE = re.compile(rf"^\s*(\d{{4}}){HYPHEN_CLASS}(\d{{2}}){HYPHEN_CLASS}(\d{{2}})\*?\s*$")

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

def normalize_token(t: str) -> str:
    t = (t or "").strip()
    t = t.replace("\u200b", "").replace("\ufeff", "").replace("\xa0", " ")
    # normalize unicode hyphens to "-"
    t = re.sub(HYPHEN_CLASS, "-", t)
    return t

def looks_like_activity_id(tok: str) -> bool:
    tok = normalize_token(tok)
    return bool(ACT_ID_RE.match(tok))

def is_package_code(tok: str) -> bool:
    tok = normalize_token(tok)
    return bool(PKG_RE.match(tok))

def infer_work_type(name: str) -> str:
    s = (name or "").lower()
    for label, pat in WORKTYPE_RULES:
        if re.search(pat, s):
            return label
    if (name or "").upper().startswith("MS"):
        return "Milestone"
    return "Other"

def parse_iso(iso: str):
    return dtparse(iso).date()

def extract_full_dates_from_tokens(tokens):
    """
    Return list of found full-date tokens in order.
    Each item: {"iso": "YYYY-MM-DD", "star": bool}
    """
    out = []
    for tok in tokens:
        raw = (tok or "").strip()
        star = raw.endswith("*")
        norm = normalize_token(raw).replace("*", "")
        m = FULL_DATE_RE.match(norm)
        if m:
            out.append({"iso": f"{m.group(1)}-{m.group(2)}-{m.group(3)}", "star": star})
    return out

def build_fragments(words):
    """
    Build line fragments by (block_no,line_no) and keep their y-mid for clustering.
    """
    line_map = {}
    for w in words:
        x0, y0, x1, y1, txt, block_no, line_no, word_no = w
        txt = normalize_token(str(txt))
        if not txt:
            continue
        key = (block_no, line_no)
        line_map.setdefault(key, []).append((x0, y0, x1, y1, txt))

    frags = []
    for ws in line_map.values():
        ws = sorted(ws, key=lambda z: z[0])
        tokens = [z[4] for z in ws if z[4]]
        if not tokens:
            continue
        y0 = min(z[1] for z in ws)
        y1 = max(z[3] for z in ws)
        ymid = (y0 + y1) / 2.0
        x0 = min(z[0] for z in ws)
        frags.append({"ymid": ymid, "x0": x0, "tokens": tokens})
    return frags

def cluster_by_y(frags, y_tol=2.0):
    """
    Cluster fragments into rows by ymid; within each row keep tokens ordered by x0.
    """
    frags = sorted(frags, key=lambda f: (f["ymid"], f["x0"]))
    clusters = []
    cur = []
    cur_y = None

    for f in frags:
        if cur_y is None:
            cur = [f]
            cur_y = f["ymid"]
        else:
            if abs(f["ymid"] - cur_y) <= y_tol:
                cur.append(f)
            else:
                clusters.append(cur)
                cur = [f]
                cur_y = f["ymid"]
    if cur:
        clusters.append(cur)

    rows = []
    for cl in clusters:
        cl = sorted(cl, key=lambda f: f["x0"])
        tokens = []
        for f in cl:
            tokens.extend(f["tokens"])
        rows.append({"ymid": sum(f["ymid"] for f in cl) / len(cl), "tokens": tokens})
    return rows

def nearest_row(rows, y):
    best = None
    best_d = 1e18
    for r in rows:
        d = abs(r["ymid"] - y)
        if d < best_d:
            best_d = d
            best = r
    return best, best_d

def extract(pdf_path: str, out_csv: str):
    rows_out = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    debug_date_rows = 0
    debug_id_rows = 0
    debug_offset = []
    debug_joined = 0
    debug_samples = []

    for page_i in range(total_pages):
        page = doc.load_page(page_i)

        # Major group headings (best effort)
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

        frags = build_fragments(words)

        # Row clusters for whole page (this is where dates finally group together)
        y_rows = cluster_by_y(frags, y_tol=2.0)

        # Build date-rows: rows that contain >=2 FULL dates (like 2026-02-15)
        date_rows = []
        for r in y_rows:
            dates = extract_full_dates_from_tokens(r["tokens"])
            if len(dates) >= 2:
                date_rows.append({"ymid": r["ymid"], "dates": dates, "raw": " | ".join(r["tokens"][:35])})
        date_rows.sort(key=lambda r: r["ymid"])

        # Build id-rows: rows that contain an activity id
        id_rows = []
        for r in y_rows:
            toks = r["tokens"]
            # skip obvious headers
            head = " ".join(toks[:3]).lower() if len(toks) >= 3 else " ".join(toks).lower()
            if head.startswith("activity id") or head.startswith("activityid"):
                continue
            if toks and toks[0].lower() in ("month", "page"):
                continue

            act_id = None
            act_pos = None
            for i in range(min(8, len(toks))):
                if looks_like_activity_id(toks[i]):
                    act_id = normalize_token(toks[i]).upper()
                    act_pos = i
                    break
            if act_id:
                name = " ".join(normalize_token(x) for x in toks[act_pos+1:]).strip()
                id_rows.append({"ymid": r["ymid"], "activity_id": act_id, "activity_name": name, "raw": " | ".join(toks[:35])})
        id_rows.sort(key=lambda r: r["ymid"])

        debug_date_rows += len(date_rows)
        debug_id_rows += len(id_rows)

        if not date_rows or not id_rows:
            continue

        # Learn page-specific offset between date rows and id rows
        # (because the two columns may be vertically shifted)
        diffs = []
        for dr in date_rows:
            nr, dist = nearest_row(id_rows, dr["ymid"])
            if nr and dist < 20.0:
                diffs.append(nr["ymid"] - dr["ymid"])
        if diffs:
            diffs_sorted = sorted(diffs)
            offset = diffs_sorted[len(diffs_sorted)//2]  # median
        else:
            offset = 0.0

        debug_offset.append(offset)

        # Join using offset-corrected y
        Y_JOIN_MAX = 25.0
        for dr in date_rows:
            target_y = dr["ymid"] + offset
            ir, dist = nearest_row(id_rows, target_y)
            if not ir or dist > Y_JOIN_MAX:
                continue

            start = dr["dates"][-2]
            finish = dr["dates"][-1]

            try:
                start_d = parse_iso(start["iso"])
                finish_d = parse_iso(finish["iso"])
            except Exception:
                continue

            duration_days = (finish_d - start_d).days
            if duration_days < 0:
                continue

            act_id = ir["activity_id"]
            name = ir["activity_name"]

            if is_package_code(act_id):
                current_package_code = act_id
                current_package_name = name

            debug_joined += 1
            if len(debug_samples) < 8:
                debug_samples.append(
                    f"offset={offset:.2f} ydist={dist:.2f} | ID: {ir['raw']} || DATES: {dr['raw']}"
                )

            rows_out.append({
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
                "start_star": bool(start.get("star", False)),
                "finish_star": bool(finish.get("star", False)),
            })

    print(f"[DEBUG] PDF pages: {total_pages}")
    print(f"[DEBUG] Date rows found (clustered): {debug_date_rows}")
    print(f"[DEBUG] ID rows found (clustered): {debug_id_rows}")
    if debug_offset:
        off_med = sorted(debug_offset)[len(debug_offset)//2]
        print(f"[DEBUG] Learned median offset across pages: {off_med:.2f}")
    print(f"[DEBUG] Joined rows produced: {debug_joined}")
    print("[DEBUG] Sample joins:")
    for s in debug_samples:
        print("   ", s)

    headers = [
        "major_group","package_code","package_name","activity_id","activity_name","work_type",
        "start","finish","duration_days","is_milestone","source_page","pdf_pages","start_star","finish_star"
    ]
    Path(out_csv).parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(rows_out)
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
