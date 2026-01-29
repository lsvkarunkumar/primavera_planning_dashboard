import re
from pathlib import Path
from dateutil.parser import parse as dtparse
import pandas as pd
import fitz  # PyMuPDF

# Hyphen variants that appear inside PDF text
HYPHEN_CLASS = r"[-\u2010\u2011\u2012\u2013\u2014\u2212\uFE63\uFF0D]"
HYPHEN_CHARS = set(["-", "‐", "-", "‒", "–", "—", "−", "﹣", "－"])

# Full date inside one token (supports unicode hyphens and optional '*')
FULL_DATE_RE = re.compile(rf"^\s*(\d{{4}}){HYPHEN_CLASS}(\d{{2}}){HYPHEN_CLASS}(\d{{2}})\*?\s*$")

# Split parts
YEAR_RE = re.compile(r"^\d{4}$")
MD_RE = re.compile(r"^\d{2}$")

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
    """Remove hidden chars and normalize unicode hyphens inside tokens."""
    t = (t or "").strip()
    t = t.replace("\u200b", "").replace("\ufeff", "").replace("\xa0", " ")
    # Normalize any unicode hyphens to '-'
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

def build_lines(words):
    """Group by (block_no, line_no) for stable fragments."""
    line_map = {}
    for w in words:
        x0, y0, x1, y1, txt, block_no, line_no, word_no = w
        txt = normalize_token(str(txt))
        if not txt:
            continue
        key = (block_no, line_no)
        line_map.setdefault(key, []).append((x0, y0, x1, y1, txt))

    lines = []
    for ws in line_map.values():
        ws_sorted = sorted(ws, key=lambda z: z[0])  # x0
        tokens = [z[4] for z in ws_sorted if z[4]]
        if not tokens:
            continue
        x0 = min(z[0] for z in ws_sorted)
        x1 = max(z[2] for z in ws_sorted)
        y0 = min(z[1] for z in ws_sorted)
        y1 = max(z[3] for z in ws_sorted)
        ymid = (y0 + y1) / 2.0
        lines.append({"x0": x0, "x1": x1, "ymid": ymid, "tokens": tokens})
    return lines

def extract_dates_from_tokens(tokens):
    """
    Return list of found dates in reading order.
    Each item: {"idx": token_index, "iso": "YYYY-MM-DD", "star": bool}
    Detects:
      A) Full date token: 2030-08-15 or 2026-02-15*
      B) Split tokens: 2030 - 08 - 15 *
    """
    dates = []

    # A) full-token dates
    for i, tok in enumerate(tokens):
        raw = (tok or "").strip()
        star = raw.endswith("*")
        norm = normalize_token(raw).replace("*", "")
        m = FULL_DATE_RE.match(norm)
        if m:
            iso = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            dates.append({"idx": i, "iso": iso, "star": star})

    # B) split dates (only if needed, but safe to run always)
    i = 0
    n = len(tokens)
    while i < n:
        t0 = normalize_token(tokens[i])
        if YEAR_RE.match(t0) and i + 4 < n:
            t1 = normalize_token(tokens[i+1])
            t2 = normalize_token(tokens[i+2])
            t3 = normalize_token(tokens[i+3])
            t4 = normalize_token(tokens[i+4])

            # hyphen token might be its own token OR already normalized to '-'
            if (t1 in HYPHEN_CHARS or t1 == "-") and MD_RE.match(t2) and (t3 in HYPHEN_CHARS or t3 == "-") and MD_RE.match(t4):
                star = False
                j = i + 5
                if j < n:
                    nxt = normalize_token(tokens[j])
                    if nxt == "*":
                        star = True
                        j += 1
                iso = f"{t0}-{t2}-{t4}"
                dates.append({"idx": i, "iso": iso, "star": star})
                i = j
                continue
        i += 1

    # De-duplicate by (iso, idx) while keeping order
    seen = set()
    out = []
    for d in sorted(dates, key=lambda x: x["idx"]):
        key = (d["iso"], d["idx"])
        if key not in seen:
            seen.add(key)
            out.append(d)
    return out

def extract(pdf_path: str, out_csv: str):
    rows = []
    current_package_code = None
    current_package_name = None
    current_major_group = None

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count

    debug_id_lines = 0
    debug_date_lines = 0
    debug_joined = 0
    debug_date_samples = []
    debug_join_samples = []

    for page_i in range(total_pages):
        page = doc.load_page(page_i)

        # Major group (best effort)
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

        lines = build_lines(words)

        id_lines = []
        date_lines = []

        for ln in lines:
            toks = ln["tokens"]
            if len(toks) < 2:
                continue

            head = " ".join(toks[:3]).lower() if len(toks) >= 3 else " ".join(toks).lower()
            if head.startswith("activity id") or head.startswith("activityid"):
                continue
            if toks[0].lower() in ("month", "page"):
                continue

            # ID line
            act_id = None
            act_pos = None
            for i in range(min(6, len(toks))):
                if looks_like_activity_id(toks[i]):
                    act_id = normalize_token(toks[i]).upper()
                    act_pos = i
                    break
            if act_id:
                name = " ".join(normalize_token(x) for x in toks[act_pos+1:]).strip()
                id_lines.append({"ymid": ln["ymid"], "activity_id": act_id, "activity_name": name, "raw": " | ".join(toks[:30])})

            # Date line (using robust date extraction)
            found_dates = extract_dates_from_tokens(toks)
            if len(found_dates) >= 2:
                date_lines.append({"ymid": ln["ymid"], "dates": found_dates, "raw": " | ".join(toks[:40])})
                if len(debug_date_samples) < 8:
                    debug_date_samples.append(" | ".join(toks[:40]))

        id_lines.sort(key=lambda r: r["ymid"])
        date_lines.sort(key=lambda r: r["ymid"])

        debug_id_lines += len(id_lines)
        debug_date_lines += len(date_lines)

        if not id_lines or not date_lines:
            continue

        # Join by nearest Y
        Y_JOIN_MAX = 14.0

        def nearest_id(y):
            best = None
            best_d = 1e9
            for r in id_lines:
                d = abs(r["ymid"] - y)
                if d < best_d:
                    best_d = d
                    best = r
            return best, best_d

        for dl in date_lines:
            lr, dist = nearest_id(dl["ymid"])
            if not lr or dist > Y_JOIN_MAX:
                continue

            start = dl["dates"][-2]
            finish = dl["dates"][-1]

            try:
                start_d = parse_iso(start["iso"])
                finish_d = parse_iso(finish["iso"])
            except Exception:
                continue

            duration_days = (finish_d - start_d).days
            if duration_days < 0:
                continue

            act_id = lr["activity_id"]
            name = lr["activity_name"]

            if is_package_code(act_id):
                current_package_code = act_id
                current_package_name = name

            debug_joined += 1
            if len(debug_join_samples) < 6:
                debug_join_samples.append(f"ydist={dist:.2f} | ID: {lr['raw']} || DATES: {dl['raw']}")

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
                "start_star": bool(start.get("star", False)),
                "finish_star": bool(finish.get("star", False)),
            })

    print(f"[DEBUG] PDF pages: {total_pages}")
    print(f"[DEBUG] ID lines found total: {debug_id_lines}")
    print(f"[DEBUG] Date lines found total: {debug_date_lines}")
    print(f"[DEBUG] Joined rows produced: {debug_joined}")
    print("[DEBUG] Sample date lines:")
    for s in debug_date_samples:
        print("   ", s)
    print("[DEBUG] Sample joins:")
    for s in debug_join_samples:
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
