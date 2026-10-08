"""
generate_dummy_data.py  -  synthetic SurveyCTO-style exports for the Inflation Project dashboard.

Writes two WIDE csv files into ./dummy_data/ that use the *real* variable names from the
XLSForms (Inflation_PreBaseline_1.xlsx / _2.xlsx), so the dashboard builder exercises exactly
the same code path it will use for real exports.  Firm IDs, addresses and shop types are
sampled from the real prefill frames; every answer is invented.  No names, phones or GPS.

Usage:   python generate_dummy_data.py            (seeded - output is reproducible)
"""
from pathlib import Path
import numpy as np
import pandas as pd

SEED = 20261008
HERE = Path(__file__).parent
OUT = HERE / "dummy_data"
OUT.mkdir(exist_ok=True)
rng = np.random.default_rng(SEED)

ENUMS = [120, 121, 122, 123, 124, 125, 126, 127]
# frame_type (Google Places) -> instrument shop_type code
TYPE_MAP = {
    "grocery_or_supermarket": 1, "supermarket": 1, "convenience_store": 1, "food": 1, "bakery": 1,
    "clothing_store": 2, "shoe_store": 2, "jewelry_store": 2,
    "electronics_store": 3, "pharmacy": 4, "drugstore": 4,
    "restaurant": 5, "cafe": 5, "meal_takeaway": 5,
    "beauty_salon": 6, "hair_care": 6, "car_repair": 6, "laundry": 6,
}


def frame_path(n):
    for p in (HERE / f"prefill_prebaseline{n}.xlsx",
              HERE.parent / "Instruments" / "Attachments" / f"prefill_prebaseline{n}.xlsx"):
        if p.exists():
            return p
    raise FileNotFoundError(f"prefill_prebaseline{n}.xlsx not found")


def shop_code(frame_type):
    for t in str(frame_type).replace("[", "").replace("]", "").replace("'", "").split(","):
        if t.strip() in TYPE_MAP:
            return TYPE_MAP[t.strip()]
    return 77


def pick(p):
    keys = list(p)
    w = np.array([p[k] for k in keys], float)
    return keys[rng.choice(len(keys), p=w / w.sum())]


def heaped_inflation(edu, base=9.0):
    """Skewed, heaped (5/10/15/20...) expectation like real firm surveys."""
    x = rng.lognormal(np.log(base), 0.45) + {1: 1.5, 2: 1.0, 3: 0, 4: -1.2}[edu]
    if rng.random() < 0.55:
        x = round(x / 5) * 5 or 5
    return float(np.clip(round(x, 1), 1, 60))


def build(n, n_attempts, n_done, days, seed_off):
    fr = pd.read_excel(frame_path(n))
    fr = fr.sort_values("sample_role", kind="stable")  # role 1 (target sample) is visited first
    firms = fr.head(n_attempts).sample(frac=1, random_state=SEED + seed_off).reset_index(drop=True)
    sfx = "" if n == 1 else "_2"
    outcomes = ["done"] * n_done
    rest = n_attempts - n_done
    others = ["partial"] * 2 + ["ref_time"] * 4 + ["ref_trust"] * 2 + ["ref_other"] + ["locked"] * 3 + \
             ["appt"] * 2 + ["mgr_na"] * 4 + ["criteria"]
    outcomes += [others[i % len(others)] for i in range(rest)]
    rng.shuffle(outcomes)
    status_code = dict(done=1, partial=2, ref_time=3, ref_trust=5, ref_other=6, locked=7, appt=8,
                       criteria=9, mgr_na=10)
    rows = []
    for i, (_, f) in enumerate(firms.iterrows()):
        oc = outcomes[i]
        day = days[min(int(i / len(firms) * len(days)), len(days) - 1)]
        start = pd.Timestamp(day) + pd.Timedelta(hours=int(rng.integers(9, 18)), minutes=int(rng.integers(0, 60)))
        enum = int(rng.choice(ENUMS, p=np.array([.19, .15, .14, .13, .12, .1, .09, .08])))
        r = {
            "KEY": f"uuid:dummy-{n}-{i:04d}", "SubmissionDate": start.strftime("%b %d, %Y %I:%M:%S %p"),
            "starttime": start.strftime("%b %d, %Y %I:%M:%S %p"), "date": start.strftime("%Y-%m-%d"),
            "enum_name": enum, "firm_id": int(f.firm_id),
            f"obs_shop_type{sfx}": np.nan if oc == "locked" else shop_code(f.frame_type),
            f"obs_employees{sfx}": int(rng.poisson(2.2)), f"obs_condition{sfx}": int(np.clip(rng.normal(6, 2), 0, 10)),
            "survey_status": status_code[oc],
        }
        if oc in ("done", "partial", "ref_time", "ref_trust", "ref_other"):
            r[f"is_manager{sfx}"] = int(rng.choice([1, 2], p=[.78, .22]))
            r[f"consent{sfx}"] = 1 if oc in ("done", "partial") else 0
        elif oc == "mgr_na":
            r[f"is_manager{sfx}"] = 3
        if oc in ("done", "partial"):
            dur_min = float(np.clip(rng.normal(21 if n == 1 else 26, 4.5), 9, 55))
            r["duration"] = int(dur_min * 60)
            edu = pick({1: .08, 2: .2, 3: .42, 4: .30})
            age = int(np.clip(rng.normal(38, 10), 20, 68))
            tenure = float(np.clip(rng.gamma(2.2, 2.2), 0, 30))
            emp = int(rng.poisson(2.6))
            fy = float(np.clip(rng.gamma(2.5, 3.5), 0.5, 40))
            area = float(np.clip(rng.lognormal(3.4, .6), 6, 400))
            pre = {f"pb{n}{'f' if n == 1 else 'c'}_1_educ": edu,
                   f"pb{n}{'f' if n == 1 else 'c'}_2_age": age,
                   f"pb{n}{'f' if n == 1 else 'c'}_3_years": int(tenure),
                   f"pb{n}{'f' if n == 1 else 'c'}_3_months": int((tenure % 1) * 12),
                   f"pb{n}{'g' if n == 1 else 'd'}_1_area": round(area, 1),
                   f"pb{n}{'g' if n == 1 else 'd'}_2_employees": emp,
                   f"pb{n}{'g' if n == 1 else 'd'}_3_years": int(fy),
                   f"pb{n}{'g' if n == 1 else 'd'}_3_months": int((fy % 1) * 12)}
            r.update(pre)
            if oc == "partial":  # partial interviews stop before the module questions
                r.update({f"pb{n}{'f' if n == 1 else 'c'}_1_educ": np.nan, f"pb{n}{'f' if n == 1 else 'c'}_2_age": np.nan})
                rows.append(r)
                continue
            prior = heaped_inflation(edu)
            if n == 1:
                u1 = int(rng.choice([0, 1, 2], p={1: [.4, .4, .2], 2: [.2, .45, .35], 3: [.1, .4, .5], 4: [.03, .27, .7]}[edu]))
                r["pb1a_1_understand"] = u1
                r["pb1a_3_understand_after"] = min(2, u1 + int(rng.random() < .55))
                r["pb1a_4_price"] = 103 if rng.random() < {1: .45, 2: .6, 3: .75, 4: .9}[edu] else float(rng.choice([3, 100, 130, 106, 110]))
                r["pb1a_5_framing"] = int(rng.choice([1, 2, 3], p=[.3, .5, .2]))
                lo = max(0.5, prior - abs(rng.normal(3, 1.5))); hi = prior + abs(rng.normal(5, 2.5))
                r.update(pb1b_1_point=prior, pb1b_2_low=round(lo, 1), pb1b_3_high=round(hi, 1))
                shown = 8.4
                post = prior + (0.45 + 0.1 * rng.random()) * (shown - prior) + rng.normal(0, .8)
                post = round(float(np.clip(post, 1, 50)), 1)
                r.update(pb1d_2_likely=post, pb1d_1_low=round(max(.5, post - abs(rng.normal(2.5, 1))), 1),
                         pb1d_3_high=round(post + abs(rng.normal(3.5, 1.5)), 1))
                names = ["price", "wage", "debt", "invest", "hours", "emp_ft", "emp_pt", "emp_unpaid"]
                keys = ["1_price", "2_wage", "3_debt", "4_invest", "5_hours", "6_emp_ft", "7_emp_pt", "8_emp_unpaid"]
                scale = np.array([.6, .35, .1, -.15, -.05, -.02, -.02, -.01])
                for k, sc in zip(keys, scale):
                    nm = k.split("_", 1)[1]
                    pre_v = round(float(rng.normal(prior * sc, 3)) / 1) if nm in ("price", "wage") else float(round(rng.normal(prior * sc, 4)))
                    pst_v = round(float(rng.normal(post * sc, 3)))
                    r[f"pb1c_{k}_3m"] = pre_v
                    r[f"pb1e_{k}_3m"] = pst_v
            else:
                r["pb2a_1_point"] = prior
                means = [8.0, 4.8, 3.4, 6.2, 5.0]
                for key, m in zip(["a_costs", "b_domestic", "c_intl", "d_firms", "e_social"], means):
                    r[f"pb2b_1_{key}"] = int(np.clip(round(rng.normal(m, 2.4)), 0, 10))
                nf = int(min(15, rng.poisson(2.6)))
                r["pb2b_2_n_firms"] = nf
                if nf > 0:
                    r["pb2b_2a_dist_cat"] = int(rng.choice([1, 2, 3, 4, 5, 6], p=[.28, .2, .2, .12, .15, .05]))
                for key, p in zip(["any", "prices", "costs", "infl"],
                                  [[.2, .3, .25, .1, .15], [.12, .25, .3, .13, .2], [.1, .25, .3, .15, .2], [.04, .12, .3, .2, .34]]):
                    r[f"pb2b_3_{key}"] = int(rng.choice([1, 2, 3, 4, 5], p=p))
                mem = int(rng.random() < .42 and nf > 0)
                r["pb2b_4_memorable"] = mem
                if mem:
                    for j, p in enumerate([[.55, .25, .15, .05]] * 7):
                        r[f"pb2b_4b_{j + 1}_" + ["start", "knew", "numbers", "advice", "check", "adjust", "prompt"][j]] = int(rng.choice([1, 2, 3, 999], p=p))
                    rels = rng.choice([1, 2, 3, 4, 5], size=int(rng.integers(1, 3)), replace=False, p=[.28, .22, .25, .1, .15])
                    r["pb2b_4c_1_bus"] = " ".join(str(x) for x in sorted(rels))
                    r["pb2b_4c_4_social"] = int(rng.choice([1, 2, 3], p=[.35, .5, .15]))
                    r["pb2b_4c_5_known"] = int(rng.choice([1, 2, 3], p=[.1, .15, .75]))
                    r["pb2b_4c_7_times"] = int(rng.integers(1, 12))
                    r["pb2b_4c_8_assoc"] = int(rng.random() < .3)
                if rng.random() < .6:
                    r["pb2b_5_firm"] = "dummy"
                    r["pb2b_5_attr"] = " ".join(str(x) for x in sorted(rng.choice([1, 2, 3, 4, 5, 6], size=int(rng.integers(1, 3)), replace=False)))
                else:
                    r["pb2b_5_firm"] = "999"
        rows.append(r)
    return pd.DataFrame(rows)


if __name__ == "__main__":
    d1 = build(1, 62, 47, pd.date_range("2026-09-29", "2026-10-05"), 1)
    d2 = build(2, 78, 44, pd.date_range("2026-10-01", "2026-10-08"), 2)
    d1.to_csv(OUT / "Inflation_PreBaseline_1_dummy_WIDE.csv", index=False)
    d2.to_csv(OUT / "Inflation_PreBaseline_2_dummy_WIDE.csv", index=False)
    print("PB1", len(d1), "rows;  PB2", len(d2), "rows  ->", OUT)
