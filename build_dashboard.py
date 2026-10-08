"""
build_dashboard.py  -  Inflation Project (Pre-Baseline 1 & 2) dashboard builder.

    python build_dashboard.py                 build index.html from the newest real exports
    python build_dashboard.py --dummy         force the synthetic demo data
    python build_dashboard.py --set-password  store a new access password (local, git-ignored)

DATA DROP (newest file wins, .csv / .dta / .xlsx, SurveyCTO "WIDE" export):
    Pre-Baseline 1  ->  ..\\Data 1\\   (or this folder)   file name contains "PreBaseline_1" / "Pre-Baseline 1" / "PB1"
    Pre-Baseline 2  ->  ..\\Data 2\\   (or this folder)   file name contains "PreBaseline_2" / "Pre-Baseline 2" / "PB2"
If neither survey has real data yet, dummy_data/ is used automatically (banner says so).

PRIVACY
    * The page embeds de-identified interview rows (no names, phones, GPS, shop names, shop IDs,
      free text) so filters can recompute charts in the browser.
    * That payload is AES-256-GCM encrypted with the access password (PBKDF2-SHA256), so the public
      index.html is unreadable without the password.  The password itself is never committed:
      it lives in password.local.txt (git-ignored) or the DASHBOARD_PASSWORD environment variable.
    * Every column that is not on the whitelist below is dropped before anything is written.
"""
import warnings
warnings.filterwarnings("ignore", message="Could not infer format")
import argparse, base64, getpass, json, os, re, sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
PW_FILE = HERE / "password.local.txt"
CFG_FILE = HERE / "dashboard_config.json"

# --------------------------------------------------------------------------------------
# Configuration (edit dashboard_config.json, not this file)
# --------------------------------------------------------------------------------------
DEFAULT_CFG = {
    "title": "Inflation Project",
    "subtitle": "Pre-Baseline Surveys · Firm Inflation Expectations",
    "survey_labels": {"1": "Pre-Baseline 1", "2": "Pre-Baseline 2"},
    # Target = number of completed interviews wanted.  null -> count of sample_role == 1 in the prefill frame.
    "targets": {"1": None, "2": None},
    "valid_inflation_range": [-20, 100],
    "pbkdf2_iterations": 200000,
}
cfg = dict(DEFAULT_CFG)
if CFG_FILE.exists():
    cfg.update(json.loads(CFG_FILE.read_text(encoding="utf-8")))

SURVEYS = (1, 2)
SFX = {1: "", 2: "_2"}                       # PB2 repeats admin variables with a _2 suffix
ED_COL = {1: "pb1f_1_educ", 2: "pb2c_1_educ"}
AGE_COL = {1: "pb1f_2_age", 2: "pb2c_2_age"}
TEN_Y = {1: "pb1f_3_years", 2: "pb2c_3_years"}
TEN_M = {1: "pb1f_3_months", 2: "pb2c_3_months"}
AREA = {1: "pb1g_1_area", 2: "pb2d_1_area"}
EMP = {1: "pb1g_2_employees", 2: "pb2d_2_employees"}
FY = {1: "pb1g_3_years", 2: "pb2d_3_years"}
FM = {1: "pb1g_3_months", 2: "pb2d_3_months"}
MARGINS = ["1_price", "2_wage", "3_debt", "4_invest", "5_hours", "6_emp_ft", "7_emp_pt", "8_emp_unpaid"]
SRC = ["a_costs", "b_domestic", "c_intl", "d_firms", "e_social"]
FREQ = ["any", "prices", "costs", "infl"]
MQ = ["start", "knew", "numbers", "advice", "check", "adjust", "prompt"]
AGE_BINS = [(18, 24, "18–24"), (25, 34, "25–34"), (35, 44, "35–44"), (45, 54, "45–54"), (55, 120, "55+")]

# Blocks of Allama Iqbal Town: spelling variants in the Google-Places addresses -> one name
BLOCK_ALIASES = {"sutlej": "Satluj", "satluj": "Satluj", "hunza": "Hunza", "pak": "Pak", "huma": "Huma",
                 "jahanzeb": "Jahanzeb", "asif": "Asif", "muslim": "Muslim", "gulshan": "Gulshan",
                 "chenab": "Chenab", "kashmir": "Kashmir", "khyber": "Khyber", "zeenat": "Zeenat",
                 "ravi": "Ravi", "neelum": "Neelum", "karim": "Karim"}


# --------------------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------------------
def num(s):
    return pd.to_numeric(s, errors="coerce")


def clean(v):
    """JSON-safe scalar."""
    if v is None:
        return None
    if isinstance(v, (float, np.floating)):
        return None if np.isnan(v) else (int(v) if float(v).is_integer() else round(float(v), 2))
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    return v


def block_of(address):
    a = str(address)
    if "moon market" in a.lower():
        return "Moon Market"
    m = re.search(r"([A-Za-z\-]+)\s+Block", a)
    if m:
        return BLOCK_ALIASES.get(m.group(1).lower(), "Other areas")
    return "Other areas"


def find_file(folder_list, tokens, exclude=("prefill", "dummy", "~$")):
    """Newest data file whose name matches any token."""
    best = None
    for folder in folder_list:
        if not folder.exists():
            continue
        for p in folder.iterdir():
            if p.suffix.lower() not in (".csv", ".dta", ".xlsx", ".xls") or not p.is_file():
                continue
            low = p.name.lower().replace(" ", "").replace("-", "").replace("_", "")
            if any(x in low for x in exclude) or not any(t in low for t in tokens):
                continue
            if p.name.lower().startswith("inflation_prebaseline_") and p.suffix.lower() in (".xlsx", ".xls") and "wide" not in low:
                continue                      # the XLSForm itself, not an export
            if best is None or p.stat().st_mtime > best.stat().st_mtime:
                best = p
    return best


def read_any(p):
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p, low_memory=False, encoding="utf-8-sig")
    if p.suffix.lower() == ".dta":
        return pd.read_stata(p, convert_categoricals=False)
    return pd.read_excel(p)


def load_choices():
    """{list_name: [(value, label_en), ...]} from the XLSForm choices sheet (both forms)."""
    out = {}
    for n in SURVEYS:
        for base in (HERE, PROJECT / "Instruments"):
            f = base / f"Inflation_PreBaseline_{n}.xlsx"
            if f.exists():
                ch = pd.read_excel(f, sheet_name="choices").fillna("")
                lab = "label: eng" if "label: eng" in ch.columns else "label"
                for ln, grp in ch.groupby("list_name", sort=False):
                    if ln and ln not in out:
                        out[ln] = [(str(v).strip(), str(l).strip()) for v, l in zip(grp["value"], grp[lab]) if str(v).strip()]
                break
    return out


def to_code(series, options):
    """Numeric code, tolerant of exports that contain the choice *label* instead of the value."""
    c = num(series)
    if c.isna().all() and series.notna().any():
        lab = {l.lower(): float(v) for v, l in options if re.fullmatch(r"-?\d+(\.\d+)?", v)}
        c = series.astype(str).str.strip().str.lower().map(lab)
    return c


def parse_dates(df):
    for col in ("date", "starttime", "SubmissionDate"):
        if col in df.columns:
            d = pd.to_datetime(df[col], errors="coerce")
            if d.notna().mean() > .5:
                return d.fillna(pd.to_datetime(df.get("SubmissionDate"), errors="coerce"))
    return pd.Series(pd.NaT, index=df.index)


def age_group(a):
    if pd.isna(a) or a < 15 or a > 90:
        return None
    for lo, hi, lab in AGE_BINS:
        if lo <= a <= hi:
            return lab
    return AGE_BINS[0][2] if a < 18 else None


def multi(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return []
    return [int(x) for x in re.findall(r"\d+", str(v)) if x]


# --------------------------------------------------------------------------------------
# Frame (prefill) -> blocks + targets
# --------------------------------------------------------------------------------------
def load_frame(n):
    for base in (HERE, PROJECT / "Instruments" / "Attachments"):
        f = base / f"prefill_prebaseline{n}.xlsx"
        if f.exists():
            fr = pd.read_excel(f)
            fr["blk"] = fr["firm_address"].map(block_of)
            return fr[["firm_id", "blk", "sample_role"]]
    print(f"  ! prefill_prebaseline{n}.xlsx not found - blocks/targets unavailable for PB{n}")
    return pd.DataFrame(columns=["firm_id", "blk", "sample_role"])


# --------------------------------------------------------------------------------------
# Per-survey record builder
# --------------------------------------------------------------------------------------
def build_records(n, df, frame, choices, enum_map):
    sx = SFX[n]
    lab = lambda ln: {v: l for v, l in choices.get(ln, [])}
    shop_l, edu_l, status_l = lab("shop_type"), lab("educ4"), lab("status_survey")
    df = df.copy()
    if "KEY" in df.columns:
        df = df.drop_duplicates("KEY", keep="last")
    df["_d"] = parse_dates(df)
    df = df[df["_d"].notna()]
    blk_of = dict(zip(frame["firm_id"].astype(int), frame["blk"]))
    role_of = dict(zip(frame["firm_id"].astype(int), frame["sample_role"]))   # 1 = main sample, 2 = replacement
    fid = num(df["firm_id"]) if "firm_id" in df.columns else pd.Series(np.nan, index=df.index)

    g = lambda c: df[c] if c in df.columns else pd.Series(np.nan, index=df.index)
    status = to_code(g("survey_status"), choices.get("status_survey", []))
    consent = to_code(g("consent" + sx), choices.get("consent_yn", []))
    status = status.where(status.notna(), np.where(consent == 1, 1, np.where(consent == 0, 3, np.nan)))
    shop = to_code(g("obs_shop_type" + sx), choices.get("shop_type", []))
    edu = to_code(g(ED_COL[n]), choices.get("educ4", []))
    age = num(g(AGE_COL[n]))
    enum = to_code(g("enum_name"), choices.get("enum_name", []))
    dur = num(g("duration")) / 60.0
    ten = num(g(TEN_Y[n])) + num(g(TEN_M[n])).fillna(0) / 12
    fy = num(g(FY[n])) + num(g(FM[n])).fillna(0) / 12

    recs = []
    for i in df.index:
        sc = status.get(i)
        sc = None if pd.isna(sc) else int(sc)
        rec = {
            "s": n, "d": clean(df.at[i, "_d"]),
            "blk": blk_of.get(int(fid[i]), "Not in frame") if pd.notna(fid[i]) else "Not in frame",
            "rl": (int(role_of[int(fid[i])]) if pd.notna(fid[i]) and int(fid[i]) in role_of and pd.notna(role_of[int(fid[i])]) else None),
            "st": shop_l.get(str(int(shop[i])), "Other") if pd.notna(shop[i]) else None,
            "en": enum_map.get(int(enum[i]), "Other") if pd.notna(enum[i]) else "Not recorded",
            "oc": sc, "ou": status_l.get(str(sc)) if sc else "Not recorded",
            "cm": 1 if sc == 1 else 0,
            "ed": edu_l.get(str(int(edu[i]))) if pd.notna(edu[i]) else None,
            "age": clean(age[i]) if pd.notna(age[i]) and 15 <= age[i] <= 90 else None,
            "ag": age_group(age[i]),
            "dur": clean(round(dur[i], 1)) if pd.notna(dur[i]) and 0 < dur[i] < 300 else None,
            "mt": clean(round(ten[i], 1)) if pd.notna(ten[i]) else None,
            "fy": clean(round(fy[i], 1)) if pd.notna(fy[i]) else None,
            "emp": clean(num(g(EMP[n]))[i]), "area": clean(num(g(AREA[n]))[i]),
        }
        if n == 1:
            rec.update(
                e1=clean(num(g("pb1b_1_point"))[i]), elo=clean(num(g("pb1b_2_low"))[i]), ehi=clean(num(g("pb1b_3_high"))[i]),
                pm=clean(num(g("pb1d_2_likely"))[i]), plo=clean(num(g("pb1d_1_low"))[i]), phi=clean(num(g("pb1d_3_high"))[i]),
                un1=clean(num(g("pb1a_1_understand"))[i]), un2=clean(num(g("pb1a_3_understand_after"))[i]),
                calc=(None if pd.isna(num(g("pb1a_4_price"))[i]) else int(abs(num(g("pb1a_4_price"))[i] - 103) < 1e-6)),
                fr=clean(num(g("pb1a_5_framing"))[i]),
                c=[clean(num(g(f"pb1c_{m}_3m"))[i]) for m in MARGINS],
                e=[clean(num(g(f"pb1e_{m}_3m"))[i]) for m in MARGINS])
        else:
            rec.update(
                e1=clean(num(g("pb2a_1_point"))[i]),
                src=[clean(num(g(f"pb2b_1_{k}"))[i]) for k in SRC],
                nf=clean(num(g("pb2b_2_n_firms"))[i]), dcat=clean(num(g("pb2b_2a_dist_cat"))[i]),
                fq=[clean(num(g(f"pb2b_3_{k}"))[i]) for k in FREQ],
                mem=clean(num(g("pb2b_4_memorable"))[i]),
                mq=[clean(num(g(f"pb2b_4b_{j + 1}_{k}"))[i]) for j, k in enumerate(MQ)],
                rel=multi(g("pb2b_4c_1_bus")[i]), soc=clean(num(g("pb2b_4c_4_social"))[i]),
                kn=clean(num(g("pb2b_4c_5_known"))[i]), tw=clean(num(g("pb2b_4c_7_times"))[i]),
                attr=multi(g("pb2b_5_attr")[i]) if str(g("pb2b_5_firm")[i]).strip() not in ("999", "nan", "") else [])
        recs.append(rec)
    return recs


# --------------------------------------------------------------------------------------
# Encryption
# --------------------------------------------------------------------------------------
def encrypt(payload: dict, password: str) -> dict:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    salt, iv = os.urandom(16), os.urandom(12)
    it = int(cfg["pbkdf2_iterations"])
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=it).derive(password.encode("utf-8"))
    ct = AESGCM(key).encrypt(iv, json.dumps(payload, separators=(",", ":")).encode("utf-8"), None)
    b = lambda x: base64.b64encode(x).decode()
    return {"salt": b(salt), "iv": b(iv), "ct": b(ct), "it": it}


def get_password():
    pw = os.environ.get("DASHBOARD_PASSWORD") or (PW_FILE.read_text(encoding="utf-8").strip() if PW_FILE.exists() else "")
    if not pw:
        sys.exit("No password set. Run:  python build_dashboard.py --set-password")
    return pw


# --------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dummy", action="store_true", help="force synthetic demo data")
    ap.add_argument("--set-password", action="store_true")
    args = ap.parse_args()

    if args.set_password:
        pw = getpass.getpass("New dashboard password: ")
        if len(pw) < 6:
            sys.exit("Use at least 6 characters.")
        PW_FILE.write_text(pw, encoding="utf-8")
        print("Saved to password.local.txt (git-ignored). Rebuild and push to apply.")
        return

    choices = load_choices()
    enum_codes = [v for v, _ in choices.get("enum_name", []) if v != "777"]
    enum_map = {int(v): f"Enumerator {i + 1:02d}" for i, v in enumerate(sorted(enum_codes, key=int))}
    enum_map[777] = "Other enumerator"

    frames = {n: load_frame(n) for n in SURVEYS}
    sources, raw = {}, {}
    if not args.dummy:
        for n in SURVEYS:
            tokens = [f"prebaseline{n}", f"pb{n}"]
            f = find_file([PROJECT / f"Data {n}", HERE], tokens)
            if f is not None:
                sources[n], raw[n] = f.name, read_any(f)
    is_dummy = args.dummy or not raw
    if is_dummy:
        sources, raw = {}, {}
        dd = HERE / "dummy_data"
        if not (dd / "Inflation_PreBaseline_1_dummy_WIDE.csv").exists():
            import subprocess
            subprocess.run([sys.executable, str(HERE / "generate_dummy_data.py")], check=True)
        for n in SURVEYS:
            f = dd / f"Inflation_PreBaseline_{n}_dummy_WIDE.csv"
            sources[n], raw[n] = f.name, read_any(f)
    for n in SURVEYS:
        print(f"PB{n}: {sources.get(n, '— no data yet —')}  ({len(raw[n]) if n in raw else 0} rows)")

    records, survey_meta = [], {}
    for n in SURVEYS:
        fr = frames[n]
        if n in raw:
            records += build_records(n, raw[n], fr, choices, enum_map)
        tgt = cfg["targets"].get(str(n))
        role1 = fr[fr["sample_role"] == 1]
        survey_meta[str(n)] = {
            "label": cfg["survey_labels"][str(n)],
            "target": int(tgt) if tgt else int(len(role1)),
            "frame_n": int(len(fr)),
            "repl_n": int((fr["sample_role"] == 2).sum()),
            "repl_blocks": {k: int(v) for k, v in fr[fr["sample_role"] == 2]["blk"].value_counts().items()},
            "blocks": {k: int(v) for k, v in role1["blk"].value_counts().items()},
            "frame_blocks": {k: int(v) for k, v in fr["blk"].value_counts().items()},
            "source": sources.get(n),
        }

    valid = [r for r in records if r["d"]]
    last = max((r["d"] for r in valid), default=None)
    now = datetime.now()
    payload = {
        "meta": {"is_dummy": bool(is_dummy), "last_date": last,
                 "generated": now.strftime("%d %b %Y, %I:%M %p").lstrip("0"),
                 "surveys": survey_meta, "valid_range": cfg["valid_inflation_range"]},
        "records": records,
        "lists": {k: choices.get(k, []) for k in
                  ("educ4", "understand3", "framing3", "dist_cat", "freq5", "ymnd4", "bus_rel", "soc_rel", "known_dur", "key_attr", "status_survey")},
    }
    # ---- PII guard: only whitelisted keys may be present ----
    allowed = {"s", "d", "blk", "rl", "st", "en", "oc", "ou", "cm", "ed", "age", "ag", "dur", "mt", "fy", "emp", "area", "e1", "elo", "ehi",
               "pm", "plo", "phi", "un1", "un2", "calc", "fr", "c", "e", "src", "nf", "dcat", "fq", "mem", "mq", "rel", "soc", "kn", "tw", "attr"}
    bad = {k for r in records for k in r} - allowed
    assert not bad, f"Non-whitelisted fields in payload: {bad}"

    tpl = (HERE / "dashboard_template.html").read_text(encoding="utf-8")
    public = {"title": cfg["title"], "is_dummy": bool(is_dummy), "generated": payload["meta"]["generated"]}
    html = (tpl.replace("__PUBLIC__", json.dumps(public))
               .replace("__PAYLOAD__", json.dumps(encrypt(payload, get_password()))))
    (HERE / "index.html").write_text(html, encoding="utf-8")

    done = {n: sum(1 for r in records if r["s"] == n and r["cm"]) for n in SURVEYS}
    print(f"\nindex.html written ({len(html) / 1024:.0f} KB) · completed PB1={done[1]} PB2={done[2]} · "
          f"{'DUMMY data' if is_dummy else 'LIVE data'} · encrypted payload")


if __name__ == "__main__":
    main()
