"""Regression tests for reconcile.py — run this before shipping any change.

    python test_reconcile.py

Two guarantees are locked in here so past mistakes cannot silently recur:

  1. LIMITS/DECIMALS behave per the agreed rules (target blank, decimals from the
     limits, no trailing .0, accuracy on Lower/Upper).
  2. DIMENSIONS are never dropped: every genuine tolerance (X±Y / X+a/-b / X~Y)
     present in a sample's OCR must appear in the extracted PDF rows.

The samples are driven from their cached raw OCR (.ocr_cache/<hash>_raw220.json),
so the test is fast and offline — no PDF or re-OCR needed.
"""
import importlib.util
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, ".ocr_cache")

# Load reconcile.py without running its Streamlit UI (the UI code at import time
# raises once there is no upload; every function is defined before that point).
_spec = importlib.util.spec_from_file_location("rec", os.path.join(HERE, "reconcile.py"))
rec = importlib.util.module_from_spec(_spec)
try:
    _spec.loader.exec_module(rec)
except Exception:
    pass

# Cached OCR by hash — the three known samples (two MIS speedometers, one SMIR mirror).
SAMPLES = {
    "34100M66RR0 (header-less MIS)": "cf179b64bf7b834041baa7791e2afa34",
    "sample3 subset (headered MIS)": "f93589919949a8e0ae2107fdc73b13f3",
    "sample2 (SMIR mirror)":         "a3b27ab225446f40e561b826ddd161ed",
    "M31100 (Denso motor, nested perf.)": "535b3274d89dd7ea454e3f4d74ac2163",
    "M72R03 (GSISP motor, min/max)":       "0dd77b8c88cbc589f43aaaff536f624e",
    "51490M77P10 (tube, OCR ± garble)":    "3a66228ca92174895213f52c722ad694",
    "51510M67L00 (tube, range-vs-±)":      "7a4af0a7515acd83b4ff92482592447e",
}

# Optional end-to-end coverage checks — run only when the source Excel is present
# (drives the full align + value-match pipeline and asserts per-MIC coverage floors).
COVERAGE = {
    "535b3274d89dd7ea454e3f4d74ac2163": (
        r"C:\Users\rstyt\Downloads\31100M54T00_D178_Simplified Excel.xlsx",
        {"Appearance": 6, "Dimension": 9, "Material": 1,
         "Performance": 11, "Revalidation & Layout Inspection": 2},  # floors (100%)
    ),
    "0dd77b8c88cbc589f43aaaff536f624e": (
        r"C:\Users\rstyt\Downloads\31100M72R03_L059_Simplified Excel.xlsx",
        {"Appearance": 5, "Dimension": 6, "Performance": 11},        # floors (85%)
    ),
    "3a66228ca92174895213f52c722ad694": (
        r"C:\Users\rstyt\Downloads\51490M77P10_B256_Simplified Excel.xlsx",
        {"Appearance": 2, "Dimension": 20, "Material": 3, "Performance": 14},  # 100%
    ),
    "7a4af0a7515acd83b4ff92482592447e": (
        r"C:\Users\rstyt\Downloads\51510M67L00_B256_Simplified Excel.xlsx",
        {"Appearance": 2, "Dimension": 20, "Material": 2, "Performance": 13},  # 88%
    ),
}

_fails = []


def check(cond, msg):
    print(("  ok   " if cond else "  FAIL ") + msg)
    if not cond:
        _fails.append(msg)


def _rows_from_cache(h):
    cache = json.load(open(os.path.join(CACHE, f"{h}_raw220.json"), encoding="utf-8"))
    rec._ocr_raw_pages = lambda pdf_bytes, dpi=220, c=cache: c
    return rec._rows_bands(b"x"), cache


def test_limits():
    print("[limits/decimals rules]")
    # (param, stored L,T,U,dp) -> (status, expected dict subset)
    st, exp = rec.limit_check("1.16 +0.06/-0.06", None, None, None, None)
    check(exp["Decimal Places"] == 2 and exp["Lower Limit"] == 1.1 and exp["Upper Limit"] == 1.22,
          "1.1/1.22 -> decimals 2")
    st, exp = rec.limit_check("1.15 +0.05/-0.05", None, None, None, None)
    check(exp["Decimal Places"] == 1, "1.1/1.2 -> decimals 1")
    st, exp = rec.limit_check("2.5 +0.5/-0.5", None, None, None, None)
    check(exp["Decimal Places"] is None and exp["Lower Limit"] == 2 and exp["Upper Limit"] == 3,
          "2/3 -> decimals blank, ints (no .0)")
    st, exp = rec.limit_check("112 ± 0.5", None, None, None, None)
    check(exp["Target Value"] is None, "target always blank when limits exist")
    st, _ = rec.limit_check("91.8 ± 0.5", 91.3, None, 92.3, 1)
    check(st == "ok", "correct L/U with empty target -> ok (not flagged)")
    st, _ = rec.limit_check("91.8 ± 0.5", 91.3, 91.8, 92.3, 1)
    check(st == "ok", "present-but-fine target -> still ok")
    st, _ = rec.limit_check("42 ± 0.5", 41.5, None, 42.6, 1)
    check(st == "mismatch", "wrong upper -> mismatch")
    # a computed limit fix must never produce a Decimal Places of 0
    for p in ("2 ± 1", "4~15 Nm", "243~320HV", "2.5 +0.5/-0.5"):
        _, e = rec.limit_check(p, None, None, None, None)
        check(e["Decimal Places"] != 0, f"{p!r}: decimals is blank not 0")
    # single-sided MIN/MAX: fills only one bound, the other stays BLANK (no 0s)
    _, e = rec.limit_check("0.3 MIN", None, None, None, None)
    check(e["Lower Limit"] == 0.3 and e["Upper Limit"] is None, "'0.3 MIN' -> lower only")
    _, e = rec.limit_check("8 MAX", None, None, None, None)
    check(e["Upper Limit"] == 8 and e["Lower Limit"] is None, "'8 MAX' -> upper only, no 0 lower")
    st, _ = rec.limit_check("0.3 MIN", None, None, None, None)
    check(st == "mismatch", "'0.3 MIN' with no stored limits -> flagged")
    st, _ = rec.limit_check("0.3 MIN", 0.3, None, None, 1)
    check(st == "ok", "'0.3 MIN' with correct lower, blank upper, dec=1 -> ok")
    # diameter-prefixed ranges: 'ø 6.6~7.1' / '0 6.6~7.1' -> lower/upper
    for p, lo, up in [("ø 6.6~7.1", 6.6, 7.1), ("0 6.6~7.1", 6.6, 7.1),
                      ("ø 3.0~3.7", 3.0, 3.7), ("Ø7.0~7.5", 7.0, 7.5)]:
        r = rec.parse_limits(p)
        check(r and r["lower"] == lo and r["upper"] == up, f"{p!r} -> {lo}/{up}")
    check(rec.parse_limits("Point 1 to 6") is None, "'Point 1 to 6' is NOT a range")
    # signed-deviation notation: 'X d1/d2' and single 'X ±a' (other side 0)
    for p, lo, up in [("7 0/-1", 6, 7), ("1.2 -1/-1.1", 0.1, 0.2), ("1.3 -1/-0.2", 0.3, 1.1),
                      ("5 -2", 3, 5), ("6 +2", 6, 8), ("6 + 2", 6, 8)]:
        _, e = rec.limit_check(p, None, None, None, None)   # via limit_check -> clean rounding
        check(e and e["Lower Limit"] == lo and e["Upper Limit"] == up,
              f"{p!r} -> {lo}/{up} (got {e['Lower Limit']}/{e['Upper Limit'] if e else None})")
    # must NOT match threads / min-max / OCR-garble as deviations
    for p in ("M10 x 1.25", "25 MICRONS MIN", "5.00.1"):
        r = rec.parse_spec(p)
        check(r is None, f"{p!r} is not a signed-deviation tolerance")


def test_residual_tokens():
    print("[residual :select:/:selected: tokens stripped]")
    for raw, want in [(":selected: 8 MAX", "8 MAX"), ("28 ± 1 :unselected:", "28 ± 1"),
                      ("A :select: B", "A B"), ("no tokens", "no tokens")]:
        got = rec.strip_residuals(raw)
        check(got == want, f"{raw!r} -> {got!r} (want {want!r})")


def test_spec_verb():
    print("[inspection verb split off PDF spec -> Info Field 1]")
    cases = [("28 ± 1 mm Check", "28 ± 1 mm", "Check"),
             ("M10x1.25 Check", "M10x1.25", "Check"),
             ("70-0.02/-0.10 Clieuk", "70-0.02/-0.10", "Check"),   # OCR variant
             ("52.5 ± 0.15 Click", "52.5 ± 0.15", "Check"),        # OCR variant
             ("As per standard Discharge Test", "As per standard", "Discharge Test"),
             ("0.3 MIN Measure", "0.3 MIN", "Measure"),
             ("28 ± 1 mm", "28 ± 1 mm", "")]                       # no verb
    for raw, spec, verb in cases:
        s, v = rec.spec_verb(raw)
        check(s == spec and v == verb, f"{raw!r} -> ({s!r}, {v!r}) want ({spec!r}, {verb!r})")


def test_dim_value_matching():
    print("[dimensions matched by VALUE, not generic item name]")
    pdf = [{"item": "DIMENSION", "spec": "263.8±1.1 MM"},
           {"item": "MTG. C.D.", "spec": "2-87±0.55MM"},
           {"item": "DIMENSION", "spec": "160±0.8"}]
    params = ["263.8±1.1 MM", "87±0.55 MM", "130±0.8 MM", "Should be smooth"]
    m = [0, 0, 0, 2]                      # a bad text alignment (all piled on row 0)
    out = rec.match_dims_by_value(params, pdf, m)
    check(out[0] == 0, "263.8 -> its own row")
    check(out[1] == 1, "87 -> '2-87' row (subset match, prefix-tolerant)")
    check(out[2] is None, "130 absent from PDF -> honest None, not a faked match")
    check(out[3] == 2, "non-dimension row left untouched")
    # distinctive item + value written differently (range vs ±) -> KEEP the text match
    pdf2 = [{"item": "SEATTHICKNESS", "spec": "1.4±0.2"},                 # = 1.2~1.6
            {"item": "PVC SLEEVE WALL THICKNES", "spec": "0.55~0.85"}]    # = 0.70±0.15
    params2 = ["1.2~1.6", "0.70±0.15"]
    items2 = ["SEAT THICKNESS", "PVC WALL THICKNESS"]
    out2 = rec.match_dims_by_value(params2, pdf2, [0, 1], xl_items=items2)
    check(out2[0] == 0, "distinctive SEAT THICKNESS kept despite range-vs-± value")
    check(out2[1] == 1, "distinctive PVC WALL kept despite range-vs-± value")
    # generic repeated item with wrong value is still nulled (no xl_items = strict)
    out3 = rec.match_dims_by_value(["130±0.8"], [{"item": "Dim", "spec": "263.8±1.1"}], [0])
    check(out3[0] is None, "generic 'Dim' wrong value still nulled")


def test_dimensions_never_dropped():
    print("[dimensions never dropped — content invariant]")
    for name, h in SAMPLES.items():
        if not os.path.exists(os.path.join(CACHE, f"{h}_raw220.json")):
            print(f"  skip  {name} (no cache)")
            continue
        rows, cache = _rows_from_cache(h)
        emitted = " ".join(rec._norm(r["spec"]) for r in rows)
        # every dimension tolerance in the spec column must appear in some row
        missing = []
        seen = set()
        # performance specs (voltage/frequency/torque/speed) can carry an incidental
        # ± or ~ but are NOT dimensions — exclude them from the dimension invariant.
        perf_unit = re.compile(r"\d[\d.]*\s*(?:v|hz|nm|rpm|hv|watt|amp|sec)\b", re.I)
        for pg in cache:
            w = pg["w"]
            for x, y, ht, t in pg["texts"]:
                if not (0.30 * w <= x < 0.50 * w):
                    continue
                if not rec._DIM_SPEC_RE.search(t) or perf_unit.search(t):
                    continue
                nt = rec._norm(t)
                if not nt or nt in seen:
                    continue
                seen.add(nt)
                if nt not in emitted:
                    missing.append(t.strip())
        check(not missing, f"{name}: {len(seen)} dims, {len(rows)} rows, missing={missing}")


def test_per_mic_coverage():
    print("[per-MIC coverage floors — end-to-end (skips if Excel absent)]")
    import pandas as pd
    for h, (xlsx, floors) in COVERAGE.items():
        if not (os.path.exists(os.path.join(CACHE, f"{h}_raw220.json")) and os.path.exists(xlsx)):
            print(f"  skip  {h} (cache or Excel missing)")
            continue
        rows, _ = _rows_from_cache(h)
        xl = pd.read_excel(xlsx, dtype=object)
        nmap = {rec._norm(c): c for c in xl.columns}   # tolerant column mapping (Target value ...)
        xl = xl.rename(columns={nmap[rec._norm(c)]: c for c in rec.SCHEMA if rec._norm(c) in nmap})[rec.SCHEMA]
        items = list(xl["Inspection Item"].astype(str))
        m = rec.align_embed(items, [r["item"] for r in rows])
        m = rec.match_dims_by_value(list(xl["Parameter"]), rows, m, xl_items=items)
        matched = ["" if j is None else rows[j]["item"] for j in m]
        from collections import Counter
        got = Counter()
        for i in range(len(xl)):
            if str(matched[i]).strip():
                got[str(xl.iloc[i]["MIC Name"])] += 1
        for mic, floor in floors.items():
            check(got[mic] >= floor, f"{mic}: {got[mic]} matched (>= {floor})")


def test_no_regression_row_counts():
    print("[row-count sanity — no collapse]")
    expect_min = {  # extraction must not fall below these known-good floors
        "cf179b64bf7b834041baa7791e2afa34": 20,   # was 24
        "f93589919949a8e0ae2107fdc73b13f3": 45,   # was 54
        "a3b27ab225446f40e561b826ddd161ed": 150,  # was 190 (208 after row_gap tune)
        "535b3274d89dd7ea454e3f4d74ac2163": 18,   # Denso motor
        "0dd77b8c88cbc589f43aaaff536f624e": 60,   # GSISP motor (87 rows)
        "3a66228ca92174895213f52c722ad694": 45,   # 51490 tube (63 rows)
        "7a4af0a7515acd83b4ff92482592447e": 30,   # 51510 tube (42 rows)
    }
    for name, h in SAMPLES.items():
        if not os.path.exists(os.path.join(CACHE, f"{h}_raw220.json")):
            continue
        rows, _ = _rows_from_cache(h)
        check(len(rows) >= expect_min[h], f"{name}: {len(rows)} rows (>= {expect_min[h]})")


if __name__ == "__main__":
    test_limits()
    test_residual_tokens()
    test_spec_verb()
    test_dim_value_matching()
    test_dimensions_never_dropped()
    test_per_mic_coverage()
    test_no_regression_row_counts()
    print()
    if _fails:
        print(f"FAILED ({len(_fails)}):")
        for m in _fails:
            print("  - " + m)
        sys.exit(1)
    print("ALL TESTS PASSED")
