#!/usr/bin/env python3
"""
o3_input_report_lungpilot.py - Summarise an O3 JSON record set into a Word (.docx) report.

Everything runs locally. Nothing is ever downloaded: pgeocode is only used
when its postal-code data file is already on disk (see --analyze-sources and
--install-source).

Examples
  python o3_input_report_lungpilot.py data.json -o report.docx
  python o3_input_report_lungpilot.py data.json --analyze-sources              # console only
  python o3_input_report_lungpilot.py data.json -o report.docx --analyze-sources
  python o3_input_report_lungpilot.py data.json --savewmods      # -> data_202611081403.json
  python o3_input_report_lungpilot.py --install-source C:\\downloads\\US.zip   # one-time setup
  python o3_input_report_lungpilot.py data.json -o r.docx --country US --units km

Requires: python-docx, matplotlib, scipy, numpy, pandas, pgeocode.
The report contains ZIP codes only - no street addresses or dates of birth.
"""
import argparse
import collections
import io
import json
import math
import os
import re
import sys
import zipfile

GEONAMES_URLS = [
    "https://download.geonames.org/export/zip/{country}.zip",
    "https://symerio.github.io/postal-codes-data/data/geonames/{country}.txt",
]
US_STATES = set(
    "AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC "
    "ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY DC PR".split())
CA_RE = re.compile(r"^[A-Z]\d[A-Z]\s?\d[A-Z]\d$", re.I)
UK_RE = re.compile(r"^[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}$", re.I)
US_RE = re.compile(r"^\d{5}(-\d{4})?$")
KM_PER_MI = 1.609344
LOST_MONTHS = 18
LOST_DAYS = round(LOST_MONTHS * 365.25 / 12)  # 548


# ----------------------------------------------------------------- helpers
def data_dir():
    return os.environ.get("PGEOCODE_DATA_DIR",
                          os.path.join(os.path.expanduser("~"), ".cache", "pgeocode"))


def load(path):
    with open(path, encoding="utf-8-sig") as f:
        d = json.load(f)
    if not isinstance(d, dict) or "Patient_List" not in d:
        sys.exit("error: %s does not look like an O3 record set (no Patient_List)" % path)
    if not d["Patient_List"]:
        sys.exit("error: %s has an empty Patient_List; nothing to analyse" % path)
    return d


def first(seq, key, default=""):
    for x in seq or []:
        v = x.get(key)
        if v not in (None, ""):
            return v
    return default


def zip5(z):
    """US ZIP -> first 5 digits; any other postal code is returned stripped and upper-cased (so --country other than US can work)."""
    z = (z or "").strip()
    m = re.match(r"(\d{5})", z)
    return m.group(1) if m else z.upper()


# ------------------------------------------------------------- source scan
def guess_country(code, address, default):
    """Return (country, certainty) for a postal code."""
    c = (code or "").strip()
    if CA_RE.match(c):
        return "CA", "format"
    if UK_RE.match(c):
        return "GB", "format"
    if US_RE.match(c):
        m = re.search(r",\s*([A-Z]{2})\b", address or "")
        if m and m.group(1) in US_STATES:
            return "US", "address state %s" % m.group(1)
        return default, "assumed (5-digit code also fits DE/FR/ES/IT/...; use --country)"
    return default, "unrecognised format"


def collect_postal(d, default_country):
    """[(patient id, field name, postal code, matching address text)] from every *PostalCode* field in the patient data."""
    out = []
    for idx, p in enumerate(d["Patient_List"], 1):
        pid = patient_label(p, idx)

        def walk(o, path=""):
            if isinstance(o, dict):
                addr_for = {}
                for k, v in o.items():
                    if k.endswith("Address") or k.endswith("AddressOfDomicile"):
                        addr_for[k] = v
                for k, v in o.items():
                    if isinstance(v, (dict, list)):
                        walk(v, path + "." + k)
                    elif "PostalCode" in k and v:
                        base = k.split("PostalCode")[0]
                        addr = next((a for ak, a in addr_for.items() if ak.startswith(base.rstrip("_"))), "")
                        out.append((pid, k, str(v), str(addr)))
            elif isinstance(o, list):
                for x in o:
                    walk(x, path)
        walk(p)
    return out


def analyze_sources(d, default_country):
    rows = collect_postal(d, default_country)
    by_cty = collections.defaultdict(lambda: {"codes": set(), "why": collections.Counter()})
    for _, _, code, addr in rows:
        cty, why = guess_country(code, addr, default_country)
        by_cty[cty]["codes"].add(zip5(code) if cty == "US" else code.strip().upper())
        by_cty[cty]["why"][why] += 1
    res = []
    for cty, info in sorted(by_cty.items()):
        path = os.path.join(data_dir(), cty + ".txt")
        present = os.path.exists(path)
        missing = None
        if present:
            try:
                import pandas as pd
                df = pd.read_csv(path, dtype={"postal_code": str}, keep_default_na=False,
                                 usecols=["postal_code"])
                known = set(df["postal_code"])
                missing = sorted(c for c in info["codes"] if c not in known)
            except Exception as e:  # noqa
                missing = ["<could not read: %s>" % e]
        res.append(dict(country=cty, codes=sorted(info["codes"]), why=dict(info["why"]),
                        file=cty + ".zip", urls=[u.format(country=cty) for u in GEONAMES_URLS],
                        cache_path=path, present=present, missing=missing))
    return res


def print_sources(src):
    print("\n=== Offline pgeocode source files needed ===")
    print("pgeocode data folder: %s   (override with PGEOCODE_DATA_DIR or --data-dir)" % data_dir())
    for s in src:
        print("\nCountry %s  - %d distinct postal code(s): %s" % (
            s["country"], len(s["codes"]), ", ".join(s["codes"][:12]) + (" ..." if len(s["codes"]) > 12 else "")))
        print("  basis      : %s" % "; ".join("%s (x%d)" % kv for kv in s["why"].items()))
        print("  download   : %s   (on a machine that may go online; the OS file only, no patient data)" % s["urls"][0])
        print("  alt source : %s" % s["urls"][1])
        print("  install    : python o3_input_report_lungpilot.py --install-source %s" % s["file"])
        print("  installs to: %s  -> %s" % (s["cache_path"], "PRESENT" if s["present"] else "MISSING"))
        if s["present"] and s["missing"]:
            print("  WARNING    : codes not found in installed file: %s" % ", ".join(s["missing"]))
        elif s["present"]:
            print("  all codes found in the installed file")


# ----------------------------------------------------------- install source
def install_source(path):
    """Convert a raw GeoNames <CC>.zip / <CC>.txt into the CSV layout pgeocode reads from its cache."""
    import pandas as pd
    os.makedirs(data_dir(), exist_ok=True)
    cc = os.path.splitext(os.path.basename(path))[0].upper()
    if cc.endswith("_FULL"):
        cc = cc[:-5]
    import pgeocode
    if path.lower().endswith(".zip"):
        with zipfile.ZipFile(path) as z:
            name = next((n for n in z.namelist() if n.upper() == cc + ".TXT"), None)
            if not name:
                sys.exit("error: %s.txt not found inside %s" % (cc, path))
            raw = io.BytesIO(z.read(name))
    else:
        raw = path
    try:  # raw GeoNames files are tab-separated with no header
        df = pd.read_csv(raw, sep="\t", header=None, names=pgeocode.DATA_FIELDS, dtype={"postal_code": str},
                         na_values=pgeocode.NA_VALUES, keep_default_na=False)
    except Exception as e:
        sys.exit("error: cannot parse %s: %s" % (path, e))
    out = os.path.join(data_dir(), cc + ".txt")
    df.to_csv(out, index=None)
    print("installed %d rows -> %s" % (len(df), out))


# --------------------------------------------------------------- analysis
def parse_date(v):
    """Date part of an O3 timestamp; None for blank or the 1001-01-01 'no data' sentinel."""
    import datetime
    try:
        dt = datetime.datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    return None if dt.year <= 1001 else dt


def days_dx_to_first_tx(p):
    """Days from earliest diagnosis (DiagStaging date) to the earlier of first RTCourse / SysTherCourse date."""
    dx = [x for x in (parse_date(r.get("DiagStaging_DateOfRecord")) for r in p.get("DiagStaging_List", [])) if x]
    tx = [x for x in (parse_date(r.get("RTCourse_DateOfRecord")) for r in p.get("RTCourse_List", [])) if x]
    tx += [x for x in (parse_date(r.get("SysTherCourse_DateOfRecord")) for r in p.get("SysTherCourse_List", [])) if x]
    return (min(tx) - min(dx)).days if dx and tx else None


def followup_days(p):
    """Days from RT course start (earliest RTCourse date) to the last PatientInformation record with a vital status."""
    st = [x for x in (parse_date(r.get("RTCourse_DateOfRecord")) for r in p.get("RTCourse_List", [])) if x]
    vs = [x for x in (parse_date(i.get("PatientInformation_DateOfRecord")) for i in p.get("PatientInformation_List", [])
                      if i.get("PatientInformation_VitalStatus")) if x]
    return (max(vs) - min(st)).days if st and vs else None


def last_vital(p):
    """(date, status) of the latest PatientInformation record that has a vital status and a usable date."""
    v = [(parse_date(i.get("PatientInformation_DateOfRecord")), i["PatientInformation_VitalStatus"])
         for i in p.get("PatientInformation_List", []) if i.get("PatientInformation_VitalStatus")]
    v = [x for x in v if x[0]]
    return max(v, key=lambda x: x[0]) if v else (None, "")


def vital_status(p):
    """'Alive' / 'Dead' / 'Lost to Followup' / 'No vital status' for one patient (see rule below)."""
    rt = [x for x in (parse_date(r.get("RTCourse_DateOfRecord")) for r in p.get("RTCourse_List", [])) if x]
    vs = [(parse_date(i.get("PatientInformation_DateOfRecord")), i["PatientInformation_VitalStatus"])
          for i in p.get("PatientInformation_List", []) if i.get("PatientInformation_VitalStatus")]
    vs = [x for x in vs if x[0]]
    if not rt or not vs:
        return "No vital status"
    latest = max(vs, key=lambda x: x[0])[1]
    if latest == "Dead":  # a recorded death is never "lost"
        return "Dead"
    if not any(abs((dt - min(rt)).days) <= LOST_DAYS for dt, _ in vs):
        return "Lost to Followup"
    return latest


def vital_counts(d):
    """Alive / Dead / Lost to Followup, per patient.
    Lost to Followup = no vital status date within LOST_MONTHS months (LOST_DAYS days, either side) of the patient's RTCourse date.
    Otherwise the status of the patient's latest dated vital status record (Alive / Dead)."""
    return collections.Counter(vital_status(p) for p in d["Patient_List"])


def stage_group(p):
    """Overall stage group I/II/III/IV from the earliest DiagStaging record with a non-blank DiagStaging_Overall (else None)."""
    v = staging_value(p, "DiagStaging_Overall")
    m = re.match(r"\s*(?:stage\s*)?(IV|III|II|I)", str(v).strip(), re.I)
    return m.group(1).upper() if m else None


def assoc_stage_group(d, values, label):
    """Numeric measure across overall stage groups I-IV: Kruskal-Wallis plus Spearman ordinal trend."""
    from scipy.stats import kruskal, spearmanr
    order = ["I", "II", "III", "IV"]
    g = collections.defaultdict(list)
    for p, v in zip(d["Patient_List"], values):
        sg = stage_group(p)
        if sg and v is not None:
            g[sg].append(v)
    desc = "; ".join("Stage %s: %s" % (k, fmt_summary(g[k], 0)) for k in order if g.get(k))
    blank = sum(1 for p in d["Patient_List"] if not stage_group(p))
    groups = [g[k] for k in order if g.get(k)]
    notes = "Stage groups: Ia/Ib->I, IIa/IIb->II, IIIa/b/c->III, IVa/b->IV; stage from earliest diagnosis record with a non-blank overall stage."
    if blank:
        notes += " Patients with blank overall stage excluded (n=%d)." % blank
    if len(groups) < 2:
        return (label, desc, "not computed", "n/a", notes + " Fewer than 2 stage groups.")
    kw = kruskal(*groups)
    xs = [order.index(k) + 1 for k in order for _ in g.get(k, [])]
    ys = [v for k in order for v in g.get(k, [])]
    tr = spearmanr(xs, ys)
    test = "Kruskal-Wallis across stage groups (H=%.2f, %d groups); ordinal trend Spearman rho=%.2f (n=%d)" % (
        kw.statistic, len(groups), tr.statistic, len(ys))
    small = ", ".join("%s n=%d" % (k, len(g[k])) for k in order if g.get(k) and len(g[k]) <= 2)
    if small:
        notes += " Groups with only 1-2 patients (%s) limit power." % small
    notes += " The trend test treats stage as ordered 1-4. Exploratory and unadjusted."
    return (label, desc, test, "Kruskal-Wallis %.3f; trend %.3f" % (kw.pvalue, tr.pvalue), notes)


def assoc_categorical(d, values, field, label):
    """Numeric measure across the categories of a patient-level field: Mann-Whitney (2 groups) or Kruskal-Wallis (3+)."""
    from scipy.stats import kruskal, mannwhitneyu
    g = collections.defaultdict(list)
    for p, v in zip(d["Patient_List"], values):
        if v is not None:
            g[p.get(field) or "Blank"].append(v)
    desc = "; ".join("%s: %s" % (k, fmt_summary(vs_, 0)) for k, vs_ in sorted(g.items()))
    if len(g) < 2:
        return (label, desc, "not computable", "n/a",
                "Only one category present (%s), so there is no comparison to make. %s" % (
                    next(iter(g), "none"), "Will compute when the field has 2 or more categories."))
    groups = list(g.values())
    if len(groups) == 2:
        r = mannwhitneyu(*groups, alternative="two-sided")
        test, pv = "Mann-Whitney U (U=%g)" % r.statistic, "%.3f" % r.pvalue
    else:
        r = kruskal(*groups)
        test, pv = "Kruskal-Wallis (H=%.2f, %d groups)" % (r.statistic, len(groups)), "%.3f" % r.pvalue
    small = ", ".join("%s n=%d" % (k, len(v)) for k, v in sorted(g.items()) if len(v) < 3)
    notes = "Rank test. Exploratory and unadjusted." + (" Groups with fewer than 3 patients (%s) limit power." % small if small else "")
    return (label, desc, test, pv, notes)


def assoc_spearman(x, y, label, xname, yname):
    """Row for the Associations table: Spearman rank correlation between two numeric patient-level measures."""
    from scipy.stats import spearmanr
    pairs = [(a, b) for a, b in zip(x, y) if a is not None and b is not None]
    n = len(pairs)
    xs, ys = [a for a, _ in pairs], [b for _, b in pairs]
    desc = "%s: %s; %s: %s" % (xname, fmt_summary(xs, 1), yname, fmt_summary(ys, 1))
    if n < 3:
        return (label, desc, "Spearman rank correlation: not computed", "n/a", "Fewer than 3 patients with both measures.")
    if len(set(xs)) < 2 or len(set(ys)) < 2:
        const = xname if len(set(xs)) < 2 else yname
        extra = (" Distance is the same for every patient because they share one domicile ZIP, likely a placeholder in this export; "
                 "this row will compute once real per-patient ZIPs are present." if "istance" in const else "")
        return (label, desc, "Spearman rank correlation: not computable", "n/a",
                "%s is identical for all %d patients, so there is nothing to correlate.%s" % (const, n, extra))
    r = spearmanr(xs, ys)
    notes = ["Rank correlation, less influenced by outliers than Pearson.", "Exploratory and unadjusted; n=%d." % n]
    if len(x) - n:
        notes.append("%d patient(s) excluded for a missing value." % (len(x) - n))
    if abs(r.statistic) >= 0.3 and r.pvalue >= 0.05:
        notes.append("Moderate effect size but not statistically significant at 0.05; may reflect low power.")
    return (label, desc, "Spearman rank correlation, rho=%.2f (n=%d)" % (r.statistic, n), "%.3f" % r.pvalue, " ".join(notes))


def patient_distances(d, rows):
    """Per-patient distance: the filled RTCourse distance field if present, else the offline-computed value (row['dist'])."""
    out = []
    for p, r in zip(d["Patient_List"], rows):
        v = next((rc["RTCourse_DistanceToPatientPrimaryResidenceDirect"] for rc in p.get("RTCourse_List", [])
                  if not is_blank_distance(rc.get("RTCourse_DistanceToPatientPrimaryResidenceDirect"))
                  and isinstance(rc["RTCourse_DistanceToPatientPrimaryResidenceDirect"], (int, float))), None)
        out.append(v if v is not None else r.get("dist"))
    return out


def assoc_numeric_by_vital(d, values, label):
    """Row for the Associations table: numeric measure vs vital status (Mann-Whitney U, Alive vs Dead)."""
    from scipy.stats import mannwhitneyu
    g = collections.defaultdict(list)
    for p, v in zip(d["Patient_List"], values):
        if v is not None:
            g[vital_status(p)].append(v)
    desc = "; ".join("%s: %s" % (k, fmt_summary(g[k], 0)) for k in ("Alive", "Dead", "Lost to Followup") if g.get(k))
    if len(g.get("Alive", [])) >= 2 and len(g.get("Dead", [])) >= 2:
        r = mannwhitneyu(g["Alive"], g["Dead"], alternative="two-sided")
        test, pv = "Mann-Whitney U, Alive vs Dead (U=%g)" % r.statistic, "%.3f" % r.pvalue
    else:
        test, pv = "not computed", "n/a"
    notes = ["Rank test chosen because the measure is skewed with outliers (max %g)." % max([v for vs_ in g.values() for v in vs_] or [0]),
             "Exploratory and unadjusted; n=%d in the test, so power is low." % (len(g.get("Alive", [])) + len(g.get("Dead", [])))]
    if pv == "n/a":
        notes.insert(0, "Fewer than 2 patients in Alive or Dead.")
    if g.get("Lost to Followup"):
        notes.append("Lost to Followup (n=%d) is shown but excluded from the test: too few to compare." % len(g["Lost to Followup"]))
    return (label, desc, test, pv, " ".join(notes))


def fmt_vital(c, n):
    return ", ".join("%s %d (%.0f%%)" % (k, c.get(k, 0), 100.0 * c.get(k, 0) / n)
                     for k in ("Alive", "Dead", "Lost to Followup", "No vital status") if k != "No vital status" or c.get(k))


def staging_counts(d, field):
    """Per patient: value of `field` from the earliest DiagStaging record where it is non-blank (else 'Blank')."""
    far = parse_date("9999-12-31")
    c = collections.Counter()
    for p in d["Patient_List"]:
        recs = sorted(p.get("DiagStaging_List", []),
                      key=lambda r: parse_date(r.get("DiagStaging_DateOfRecord")) or far)
        c[next((r[field] for r in recs if r.get(field)), "Blank")] += 1
    return c


def fmt_counts(c, n):
    return ", ".join("%s %d (%.0f%%)" % (k, v, 100.0 * v / n) for k, v in sorted(c.items(), key=lambda kv: (kv[0] == "Blank", kv[0])))


def gtv_volume_cc(p):
    """Total GTV volume (cc): sum of Volume[cc] over structures named GTV* (GTVp, GTVn, GTV01, ...).
    If a patient has no GTV* structure, the IGTV* structures are summed instead. A structure repeated
    across plans is counted once. None if neither exists."""
    gtv, igtv = {}, {}
    for rc in p.get("RTCourse_List", []):
        for pl in rc.get("RTTreatedPlan_List", []):
            for m in pl.get("DVHMetric_List", []):
                nm = (m.get("DVHMetric_StructureName") or "").strip().upper()
                if m.get("DVHMetric_Name") != "Volume[cc]" or not isinstance(m.get("DVHMetric_Value"), (int, float)):
                    continue
                if nm.startswith("GTV"):
                    gtv.setdefault(nm, m["DVHMetric_Value"])
                elif nm.startswith("IGTV"):
                    igtv.setdefault(nm, m["DVHMetric_Value"])
    use = gtv or igtv
    return sum(use.values()) if use else None


CHEMO_AGENTS = {  # canonical name -> regex of aliases / common misspellings (matched on word starts)
    "Carboplatin": r"carbo\w*platin|paraplatin",
    "Cisplatin": r"cis-?platin",
    "Oxaliplatin": r"oxali\w*platin",
    "Paclitaxel": r"pacl?i?t?a[cx]\w*|taxol|abraxane",
    "Docetaxel": r"docetaxel|docetaxal|taxotere",
    "Gemcitabine": r"gemcitabine|gemzar",
    "Pemetrexed": r"pemetrexed|alimta",
    "Etoposide": r"etoposide|vp-?16",
    "Vinorelbine": r"vinorelbine|navelbine",
    "Irinotecan": r"irinotecan",
    "Topotecan": r"topotecan",
    "Doxorubicin": r"doxorubicin|adriamycin",
    "Cyclophosphamide": r"cyclophosphamide",
    "Fluorouracil": r"fluorouracil|5-?fu",
    "Capecitabine": r"capecitabine|xeloda",
    "Temozolomide": r"temozolomide",
}


IMMUNO_AGENTS = {  # canonical name -> regex of aliases
    "Pembrolizumab": r"pembrolizumab|keytruda",
    "Nivolumab": r"nivolumab|opdivo",
    "Durvalumab": r"durvalumab|imfinzi",
    "Atezolizumab": r"atezolizumab|tecentriq",
    "Ipilimumab": r"ipilimumab|yervoy",
    "Cemiplimab": r"cemiplimab|libtayo",
    "Tremelimumab": r"tremelimumab",
}


def chemo_agents(p, agents=None):
    """Set of chemotherapy agents for a patient, parsed from SysTherCourse/Cycle names and structured chemo drug entries."""
    texts = []
    for sc in p.get("SysTherCourse_List", []):
        texts.append(sc.get("SysTherCourse_Name") or "")
        for cy in sc.get("SysTherCycle_List", []):
            texts.append(cy.get("SysTherCycle_Name") or "")
            texts += [x.get("SysTherCycleDrugsChemo_DrugName") or "" for x in cy.get("SysTherCycleDrugsChemo_List", [])]
            texts += [x.get(k) or "" for x in cy.get("SysTherCycleDrugsImmuno_List", []) for k in x if k.endswith("DrugName")]
    txt = " ".join(texts)
    return {k for k, rx in (agents or CHEMO_AGENTS).items() if re.search(r"(?<![a-z])(?:%s)" % rx, txt, re.I)}


def chemo_text(d, n, agents=None, none_label="No chemotherapy"):
    ag = [chemo_agents(p, agents) for p in d["Patient_List"]]
    c = collections.Counter(a for s_ in ag for a in s_)
    parts = ["%s %d (%.0f%%)" % (k, v, 100.0 * v / n) for k, v in sorted(c.items(), key=lambda kv: (-kv[1], kv[0]))]
    none = sum(1 for s_ in ag if not s_)
    parts.append("%s %d (%.0f%%)" % (none_label, none, 100.0 * none / n))
    return ", ".join(parts)


def fmt_summary(vals, nd=1):
    """median [Q1, Q3] (min, max)"""
    import numpy as np
    v = [float(x) for x in vals if x is not None]
    if not v:
        return "n/a"
    q1, med, q3 = np.percentile(v, [25, 50, 75])
    f = lambda x: ("%.*f" % (nd, x)).rstrip("0").rstrip(".") if nd else "%d" % round(x)
    return "%s [%s, %s] (%s, %s)  n=%d" % (f(med), f(q1), f(q3), f(min(v)), f(max(v)), len(v))


def site_patients_per_year(d):
    """{year: patients treated} from HCO_SiteCount 'Number of Patients' records. Duplicate (site, year) records are collapsed
    (conflicting duplicates keep the first value); sites are summed within a year. Returns (dict, n_records, n_conflicts)."""
    seen, nrec, conflicts = {}, 0, 0
    for site in d.get("HCO_Site_List", []):
        for c in site.get("HCO_SiteCount_List", []):
            if c.get("HCO_SiteCount_Type") != "Number of Patients" or c.get("HCO_SiteCount_TimePointType") != "Calendar Year":
                continue
            v = c.get("HCO_SiteCount_Value")
            if not isinstance(v, (int, float)):
                continue
            nrec += 1
            key = (site.get("HCO_Site_GUID") or site.get("HCO_Site_Name"), c.get("HCO_SiteCount_TimePointValue"))
            if key in seen:
                conflicts += seen[key] != v
            else:
                seen[key] = v
    per = collections.Counter()
    for (_, yr), v in seen.items():
        per[yr] += v
    return dict(per), nrec, conflicts


def treatment_units(d):
    """HCO_SiteImgTrtUnitInfo_List items across sites -> (n_items, unique_units{key: item}).
    A unit is identified by (site, name, serial number); repeated list items for the same unit are collapsed."""
    items, uniq = 0, {}
    for site in d.get("HCO_Site_List", []):
        for u in site.get("HCO_SiteImgTrtUnitInfo_List", []) or []:
            items += 1
            key = (site.get("HCO_Site_GUID"), u.get("HCO_SiteImgTrtUnitInfo_Name"), u.get("HCO_SiteImgTrtUnitInfo_SerialNumber"))
            uniq.setdefault(key, u)
    return items, uniq


def staging_value(p, field):
    """Earliest dated non-blank value of a DiagStaging field for one patient ('' if none)."""
    far = parse_date("9999-12-31")
    recs = sorted(p.get("DiagStaging_List", []), key=lambda r: parse_date(r.get("DiagStaging_DateOfRecord")) or far)
    return next((r[field] for r in recs if r.get(field)), "")


def curation_flags(d, rows, dists):
    """Per patient: (id, [missing / inconsistent items], [outlier items]).
    Outliers: Tukey fences (1.5 x IQR beyond Q1/Q3) across patients, on a log scale for skewed measures
    (days, GTV); only when at least 8 patients have a value and the IQR is non-zero."""
    import numpy as np
    pats = d["Patient_List"]
    measures = [  # key, label, unit, log scale, values
        ("dx_to_tx", "Days diagnosis to treatment", "d", True, [r["dx_to_tx"] for r in rows]),
        ("followup", "Follow-up", "d", False, [r["followup"] for r in rows]),
        ("age", "Age at treatment", "y", False, [r["age"] for r in rows]),
        ("fx", "Fractions", "", False, [r["fx"] for r in rows]),
        ("dose", "Max target dose", "Gy", False, [r["dose"] for r in rows]),
        ("gtv", "GTV volume", "cc", True, [r["gtv"] for r in rows]),
        ("dist", "Distance to facility", "", False, dists),
    ]
    out = [(r["id"], [], []) for r in rows]
    for _, label, unit, lg, vals in measures:
        idx = [i for i, v in enumerate(vals) if v is not None and (v > 0 or not lg)]
        if len(idx) < 8:
            continue
        arr = np.array([vals[i] for i in idx], float)
        w = np.log10(arr) if lg else arr
        q1, q3 = np.percentile(w, [25, 75])
        iqr = q3 - q1
        if iqr == 0:
            continue
        lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        for i, x, raw in zip(idx, w, arr):
            if x < lo or x > hi:
                out[i][2].append("%s %s%s (%s)" % (label, ("%g" % raw), (" " + unit) if unit else "", "high" if x > hi else "low"))
    for i, (p, r) in enumerate(zip(pats, rows)):
        miss = out[i][1]
        rc = (p.get("RTCourse_List") or [{}])[0]
        if r["dx_to_tx"] is None:
            miss.append("no usable diagnosis or treatment date")
        elif r["dx_to_tx"] < 0:
            miss.append("treatment starts %d days before earliest diagnosis record" % -r["dx_to_tx"])
        if r["n_courses"] > 1:
            miss.append("%d RT courses; fractions, dose, technique, age and distance use the first only" % r["n_courses"])
        if r["dose"] is None:
            miss.append("no target dose entry")
        if r["gtv"] is None:
            miss.append("no GTV/IGTV volume")
        if not staging_value(p, "DiagStaging_Overall"):
            miss.append("overall stage blank")
        for f, nm in (("DiagStaging_StagingT", "T"), ("DiagStaging_StagingN", "N"), ("DiagStaging_StagingM", "M")):
            if not staging_value(p, f):
                miss.append("%s stage blank" % nm)
        if r["tech"] in ("n/a", "None", None):
            miss.append("RT technique missing")
        if r["age_dx"] is None:
            miss.append("no age at diagnosis")
        if last_vital(p)[0] is None:
            miss.append("no dated vital status")
        elif vital_status(p) == "Lost to Followup":
            miss.append("lost to follow-up (no vital status within %d months of RT)" % LOST_MONTHS)
        ns = rc.get("RTCourse_NSessionsInCourse")
        if isinstance(ns, int) and ns != len(rc.get("RTCourseSession_List", [])):
            miss.append("session count mismatch (%d listed vs %d recorded)" % (len(rc.get("RTCourseSession_List", [])), ns))
        if not r["dz"]:
            miss.append("domicile postal code missing (no distance calculated)")
        if not r["fz"]:
            miss.append("facility postal code missing (no distance calculated)")
        if dists[i] is None and r["dz"] and r["fz"]:
            miss.append("no distance value")
    return out


def summary_note(label, d, rows, n):
    """Explanation / caveat text for a row of the summary table (includes data-driven counts)."""
    def nmiss(key):
        return sum(1 for r in rows if r[key] is None)

    def one_cat(field):
        vals = {p.get(field) or "Blank" for p in d["Patient_List"]}
        return " Only one category present (%s): no information for analysis." % next(iter(vals)) if len(vals) == 1 else ""
    l = label.lower()
    if l.startswith("time from earliest diagnosis"):
        mx = max([r["dx_to_tx"] for r in rows if r["dx_to_tx"] is not None] or [0])
        return ("Earliest DiagStaging date to earliest RTCourse or SysTherCourse date. 1001-01-01 placeholder dates ignored; "
                "%d of %d patients missing a usable interval. Maximum (%d d) is a strong outlier, so medians/quantiles are "
                "used. Dates appear shifted per patient; intervals are meaningful, calendar dates may not be." % (nmiss("dx_to_tx"), n, mx))
    if l.startswith("follow-up"):
        return ("Earliest RTCourse date to the latest dated vital status record. This is time to last contact, not time to "
                "death: the data has no death date. %d of %d patients missing." % (nmiss("followup"), n))
    if l.startswith("vital status"):
        return ("Dead/Alive = status of the latest dated vital status record. Lost to Followup = no vital status date within %d months "
                "(%d days, either side) of the patient's RTCourse date. Most patients have a baseline Alive record near RT start, "
                "which keeps this category small. Percentages are of all %d patients." % (LOST_MONTHS, LOST_DAYS, n))
    if l.startswith("sex"):
        return "Patient_SexAtBirth." + one_cat("Patient_SexAtBirth")
    if l.startswith("race"):
        return "Patient_Race." + one_cat("Patient_Race")
    if l.startswith("ethnicity"):
        return "Patient_Ethnicity." + one_cat("Patient_Ethnicity")
    if l.split()[0] in ("t", "n", "m", "overall") and "stage" in l:
        return ("Per patient, from the earliest dated DiagStaging record with a non-blank value for this field; "
                "patients with none are shown as Blank. Stage values sorted alphabetically, not clinically.")
    if l.startswith("gtv"):
        return ("Sum of Volume[cc] for structures named GTV* (GTVp, GTVn, GTV01...); IGTV* used only if the patient has no GTV*. "
                "CTV, IGRV and other structures are not counted. %d of %d patients have neither, so n is smaller." % (nmiss("gtv"), n))
    if l.startswith("chemotherapy"):
        return ("Parsed from systemic therapy course/cycle names and structured chemo drug entries using a local drug-name list "
                "(includes brand names and misspellings). Agents not in the list are not counted. Each patient counts once per agent. "
                "Immunotherapy is excluded here. Percentages are of all %d patients." % n)
    if l.startswith("immunotherapy"):
        return ("Parsed like chemotherapy, using a local list of checkpoint inhibitors (pembrolizumab, nivolumab, durvalumab...). "
                "Each patient counts once per agent. Percentages are of all %d patients." % n)
    if l.startswith("age at diagnosis"):
        return "Age at the earliest diagnosis record (DiagStaging_AgeAtRecord); %d of %d patients missing." % (nmiss("age_dx"), n)
    if l.startswith("fractions"):
        return "RTCourse_NSessionsInCourse (first RT course). 5-fraction courses are likely SBRT."
    if l.startswith("maximum target dose"):
        return ("Highest RTCourseTargetDose per patient (planned dose to target). %d of %d patients have no target-dose entry "
                "and are excluded." % (nmiss("dose"), n))
    if l.startswith("patients by year"):
        return "Year of the earliest RTCourse date. Dates appear shifted/obfuscated in this export, so years may not be real treatment years."
    if l.startswith("patients treated per year"):
        per, nrec, conf = site_patients_per_year(d)
        return ("HCO_SiteCount_Value where HCO_SiteCount_Type is 'Number of Patients', one value per calendar year (HCO_SiteCount_TimePointValue). "
                "Whole-site volume, not this %d-patient sample. %d records collapsed to %d unique year(s) (%s)%s; "
                "quantiles over so few years are only descriptive." % (
                    n, nrec, len(per), ", ".join(str(y) for y in sorted(per)),
                    "; %d duplicate(s) had conflicting values (first kept)" % conf if conf else "; duplicate records had identical values"))
    if l.startswith("treatment machines"):
        items, uniq = treatment_units(d)
        models = collections.Counter((u.get("HCO_SiteImgTrtUnitInfo_Model") or "unknown") for u in uniq.values())
        makes = sorted({(u.get("HCO_SiteImgTrtUnitInfo_Make") or "unknown") for u in uniq.values()})
        depts = collections.Counter((u.get("HCO_SiteImgTrtUnitInfo_DepartmentName") or "unknown") for u in uniq.values())
        return ("Count of HCO_SiteImgTrtUnitInfo_List items. %d list items describe %d distinct machines (same name and serial number), "
                "so repeated entries are collapsed. Make: %s. Models: %s. Departments: %s. Machine type and capabilities fields are blank, "
                "so linacs, Gamma Knife and other units are not distinguished here." % (
                    items, len(uniq), ", ".join(makes), ", ".join("%s %d" % kv for kv in models.most_common()),
                    ", ".join("%s %d" % kv for kv in depts.most_common())))
    if l.startswith("distance"):
        vals = {round(r["dist"], 1) for r in rows if r.get("dist") is not None}
        same = (" All patients have the same distance (one shared domicile ZIP), likely a placeholder in this export."
                if len(vals) == 1 else "")
        has_field = any(not is_blank_distance(rc.get("RTCourse_DistanceToPatientPrimaryResidenceDirect"))
                        and isinstance(rc.get("RTCourse_DistanceToPatientPrimaryResidenceDirect"), (int, float))
                        for p in d["Patient_List"] for rc in p.get("RTCourse_List", []))
        src_txt = ("Taken from the RTCourse distance field in the JSON." if has_field
                   else "Computed offline from postal-code centroids.")
        return ("%s Intended as straight-line (great-circle) distance between patient domicile and RT facility; not street "
                "address, not driving distance (check DirectMethod if taken from the field).%s" % (src_txt, same))
    return ""


def patient_label(p, idx):
    """Patient_AnonPatID (never the MRN); if blank, say which item of Patient_List it is (1-based)."""
    pid = str(p.get("Patient_AnonPatID") or "").strip()
    return pid if pid else "(blank AnonPatID - Patient_List item %d)" % idx


def patient_table(d):
    rows = []
    for idx, p in enumerate(d["Patient_List"], 1):
        rc = (p.get("RTCourse_List") or [{}])[0]
        pinfo = p.get("PatientInformation_List") or []
        tech = first([t.get("RTTreatedPlanDetailsXRT") or {} for t in rc.get("RTTreatedPlan_List", [])],
                     "RTTreatedPlanDetailsXRT_Technique", "n/a")
        doses = {t["RTCourseTarget_VolumeDose"] for t in rc.get("RTCourseTargetDose_List", [])
                 if isinstance(t.get("RTCourseTarget_VolumeDose"), (int, float))}
        ages = [e["RTCourseEvent_AgeAtRecord"] for e in rc.get("RTCourseEvent_List", [])
                if e.get("RTCourseEvent_AgeAtRecord") is not None]
        rows.append(dict(
            id=patient_label(p, idx),
            icd=staging_value(p, "DiagStaging_ICDCode"),
            stage=staging_value(p, "DiagStaging_Overall") or "Blank",
            tnm="".join(staging_value(p, k) for k in ("DiagStaging_StagingT", "DiagStaging_StagingN", "DiagStaging_StagingM")),
            fx=rc.get("RTCourse_NSessionsInCourse"),
            dose=max(doses) if doses else None,
            tech=tech,
            age=min(ages) if ages else None,
            chemo=len(p.get("SysTherCourse_List", [])) > 0,
            vital=vital_status(p),
            n_courses=len(p.get("RTCourse_List", [])),
            images=len(p.get("Image_List", [])),
            dx_to_tx=days_dx_to_first_tx(p),
            followup=followup_days(p),
            gtv=gtv_volume_cc(p),
            age_dx=min([x["DiagStaging_AgeAtRecord"] for x in p.get("DiagStaging_List", []) if x.get("DiagStaging_AgeAtRecord") is not None] or [None], key=lambda z: (z is None, z)),
            dz=zip5(first(pinfo, "PatientInformation_PostalCodeDomicile")),
            fz=zip5(rc.get("RTCourse_FacilityPostalCode", "")),
        ))
    return rows


def compute_distances(rows, country, units):
    """Offline distances. Returns (list_or_None, note). Never triggers a download."""
    path = os.path.join(data_dir(), country + ".txt")
    if not os.path.exists(path):
        return None, "pgeocode data file %s is not installed, so distances were not computed " \
                     "(run with --analyze-sources, then --install-source)." % path
    import pgeocode
    import numpy as np
    gd = pgeocode.GeoDistance(country)
    vals = np.full(len(rows), np.nan)
    ok = [i for i, r in enumerate(rows) if r["dz"] and r["fz"]]  # missing postal code -> no distance is calculated
    if ok:
        km = np.asarray(gd.query_postal_code([rows[i]["dz"] for i in ok], [rows[i]["fz"] for i in ok]), dtype=float)
        vals[ok] = km if units == "km" else km / KM_PER_MI
    for r, v in zip(rows, vals):
        r["dist"] = None if np.isnan(v) else float(v)
    n_missing = len(rows) - len(ok)
    n_notfound = int(np.isnan(vals).sum()) - n_missing
    note = "Straight-line (great-circle) distance between postal-code centroids, computed offline with pgeocode."
    if n_missing:
        note += " %d patient(s) excluded: domicile or facility postal code missing, so no distance was calculated." % n_missing
    if n_notfound:
        note += " %d patient(s) excluded: postal code not found in the installed file." % n_notfound
    return [float(v) for v in vals if not np.isnan(v)], note


def is_blank_distance(v):
    """Blank = missing, null, empty string, or the O3 'no data' sentinel -9999."""
    if v is None or (isinstance(v, str) and v.strip() in ("", "-9999")):
        return True
    return isinstance(v, (int, float)) and v == -9999


def fill_distances(d, country, units, out):
    """Write a COPY of the record set with the RTCourse direct-distance fields filled in (offline)."""
    path = os.path.join(data_dir(), country + ".txt")
    if not os.path.exists(path):
        sys.exit("error: %s not installed; run --analyze-sources / --install-source first" % path)
    import pgeocode
    import numpy as np
    gd = pgeocode.GeoDistance(country)
    method = "Postal Code"
    n = skipped = kept = 0
    for p in d["Patient_List"]:
        dz = zip5(first(p.get("PatientInformation_List"), "PatientInformation_PostalCodeDomicile"))
        for rc in p.get("RTCourse_List", []):
            if not is_blank_distance(rc.get("RTCourse_DistanceToPatientPrimaryResidenceDirect")):
                kept += 1  # already has a value: leave value, method and units untouched
                continue
            fz = zip5(rc.get("RTCourse_FacilityPostalCode", ""))
            km = float(np.asarray(gd.query_postal_code([dz], [fz]), dtype=float)[0]) if dz and fz else float("nan")
            if km != km:
                skipped += 1
                continue
            rc["RTCourse_DistanceToPatientPrimaryResidenceDirect"] = round(km if units == "km" else km / KM_PER_MI, 1)
            rc["RTCourse_DistanceToPatientPrimaryResidenceDirectMethod"] = method
            rc["RTCourse_DistanceToPatientDomicileDirectUnits"] = units
            n += 1
    with open(out, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=2, ensure_ascii=False)
    print("filled %d RT course(s), left %d with existing values, %d not computable (ZIP missing or not in installed file); wrote %s" % (n, kept, skipped, out))
    print("WARNING: the output is a full copy of the input and still contains any DOB / street address fields. Handle it as PHI.")


def make_hist_grid(panels, ncols=2):
    """panels: list of (title, values, xlabel, log_x). One PNG with a histogram per panel. Returns (BytesIO, n_panels)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np
    nrows = math.ceil(len(panels) / ncols)
    fig, axes = plt.subplots(nrows, ncols, figsize=(5.0 * ncols, 3.0 * nrows), squeeze=False)
    for ax in axes.flat[len(panels):]:
        ax.axis("off")
    for ax, (title, vals, xlabel, log_x) in zip(axes.flat, panels):
        v = np.asarray(vals, dtype=float)
        if log_x and v.min() > 0:
            bins = np.logspace(np.log10(v.min()), np.log10(v.max()), 9) if v.max() > v.min() else 5
            ax.set_xscale("log")
        else:
            bins = min(10, max(3, len(set(np.round(v, 1)))))
        ax.hist(v, bins=bins, color="#3b6ea5", edgecolor="white")
        ax.axvline(np.median(v), color="#c0392b", lw=1.2, ls="--")
        ax.set_title("%s (n=%d)" % (title, len(v)), fontsize=9)
        ax.set_xlabel(xlabel + (" [log scale]" if log_x else ""), fontsize=8)
        ax.set_ylabel("Patients", fontsize=8)
        ax.tick_params(labelsize=8)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf, len(panels)


def logistic(x, L, k, x0):
    import numpy as np
    return L / (1.0 + np.exp(-k * (x - x0)))


def fit_logistic(x, y):
    """Fit y = L / (1 + exp(-k (x - x0))). Returns dict with params, std errors, R2, or {'error': msg}."""
    import numpy as np
    from scipy.optimize import curve_fit
    x, y = np.asarray(x, float), np.asarray(y, float)
    if len(x) < 4:
        return {"error": "fewer than 4 points (3 parameters need more data)"}
    if np.ptp(x) == 0:
        return {"error": "x values identical"}
    best = None
    kmax = 180.0 / np.ptp(x)  # steepest allowed: ~ a full 0-to-L transition within 1/45 of the x range
    starts = [sg * c / np.ptp(x) for c in (2, 5, 10, 20, 50) for sg in (1, -1)]
    for k0 in starts:
        try:
            popt, pcov = curve_fit(logistic, x, y, p0=[max(y.max(), 1.0) * 1.1, k0, float(np.median(x))],
                                   bounds=([0, -kmax, x.min() - np.ptp(x)], [max(y.max(), 1.0) * 10, kmax, x.max() + np.ptp(x)]), maxfev=20000)
        except Exception:
            continue
        sse = float(np.sum((y - logistic(x, *popt)) ** 2))
        if best is None or sse < best[0]:
            best = (sse, popt, pcov)
    if best is None:
        return {"error": "fit did not converge"}
    sse, popt, pcov = best
    sst = float(np.sum((y - y.mean()) ** 2))
    se = np.sqrt(np.abs(np.diag(pcov))) if np.all(np.isfinite(pcov)) else [float("nan")] * 3
    return {"L": popt[0], "k": popt[1], "x0": popt[2], "se": list(se), "r2": 1 - sse / sst if sst > 0 else float("nan"), "n": len(x), "kmax": kmax}


def scatter_logistic_png(x, y, xlabel, ylabel, fit, log_y=False):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np
    fig, ax = plt.subplots(figsize=(7.5, 4.5))
    ax.scatter(x, y, color="#3b6ea5", s=35, zorder=3, label="Patients (n=%d)" % len(x))
    if "error" not in fit:
        xs = np.linspace(min(x), max(x), 300)
        ax.plot(xs, logistic(xs, fit["L"], fit["k"], fit["x0"]), color="#c0392b", lw=2,
                label="Logistic fit  y = L / (1 + exp(-k (x - x0)))")
    if log_y:
        ax.set_yscale("log")
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel + (" [log scale]" if log_y else ""))
    ax.legend(fontsize=8)
    ax.grid(alpha=0.25)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf


# -------------------------------------------------------------------- docx
def add_logistic_plot(doc, title, pairs, xlabel, xunit):
    """Scatter of days from diagnosis to treatment (y) vs x with a logistic fit, coefficient table and notes."""
    from docx.shared import Inches
    doc.add_heading(title, 2)
    pts = [(a, b) for a, b in pairs if a is not None and b is not None]
    result = {"plot": title, "n": len(pts)}
    if len(pts) < 2:
        doc.add_paragraph("Fewer than 2 patients with both values; plot not drawn.")
        result["fit"] = None
        result["notes"] = "Fewer than 2 patients with both values; plot not drawn."
        return result
    xa, ya = [a for a, _ in pts], [b for _, b in pts]
    fit = fit_logistic(xa, ya)
    doc.add_picture(scatter_logistic_png(xa, ya, xlabel, "Days from diagnosis to treatment", fit), width=Inches(6.5))
    result["model"] = "days = L / (1 + exp(-k (x - x0)))"
    if "error" in fit:
        result["fit"] = None
        result["notes"] = "Logistic fit not available: %s." % fit["error"]
        doc.add_paragraph("Logistic fit not available: %s. %s" % (fit["error"], (
            "All patients have the same x value (e.g. one shared domicile ZIP in this export), so no curve can be fitted; "
            "the fit will run automatically on data sets with varying values." if fit["error"] == "x values identical" else "")))
        return result
    ft = doc.add_table(rows=1, cols=3)
    ft.style = "Light Grid Accent 1"
    for i, h in enumerate(["Coefficient", "Estimate", "Std. error"]):
        ft.rows[0].cells[i].text = h
    for nm, val, se in (("L (upper asymptote, days)", fit["L"], fit["se"][0]),
                        ("k (steepness, per %s)" % xunit.rstrip("s"), fit["k"], fit["se"][1]),
                        ("x0 (midpoint, %s)" % xunit, fit["x0"], fit["se"][2])):
        c = ft.add_row().cells
        c[0].text, c[1].text, c[2].text = nm, "%.4g" % val, "%.3g" % se
    doc.add_paragraph("Model: days = L / (1 + exp(-k (x - x0))), least-squares fit (n=%d), R-squared = %.3f." % (fit["n"], fit["r2"]))
    med = sorted(ya)[len(ya) // 2]
    notes = ["Exploratory: with %d points and 3 parameters the fit is poorly constrained; standard errors are approximate." % fit["n"]]
    if fit["r2"] < 0.1:
        notes.append("The curve explains little of the variation (R-squared %.2f)." % fit["r2"])
    if max(ya) > 5 * med:
        notes.append("One or more very long intervals (max %d days vs median %g) pull the least-squares fit; the logistic form was not chosen from the data." % (max(ya), med))
    if abs(fit["k"]) >= 0.99 * fit["kmax"] or fit["L"] >= 9.9 * max(max(ya), 1.0):
        notes.append("A coefficient reached its fitting bound, so the logistic shape is not identifiable from these data.")
    if (not math.isnan(fit["se"][1])) and fit["se"][1] > abs(fit["k"]):
        notes.append("The standard error of k exceeds k itself, so the steepness is essentially undetermined.")
    doc.add_paragraph(" ".join(notes))
    result["fit"] = {"L": {"estimate": fit["L"], "std_error": fit["se"][0], "unit": "days"},
                     "k": {"estimate": fit["k"], "std_error": fit["se"][1], "unit": "per " + xunit.rstrip("s")},
                     "x0": {"estimate": fit["x0"], "std_error": fit["se"][2], "unit": xunit},
                     "r_squared": fit["r2"]}
    result["notes"] = " ".join(notes)
    return result


def json_clean(o):
    """Make results JSON-safe: NaN/inf -> None, numpy numbers -> Python numbers."""
    if isinstance(o, dict):
        return {str(k): json_clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple, set)):
        return [json_clean(v) for v in o]
    if hasattr(o, "item") and not isinstance(o, (str, bytes)):
        try:
            o = o.item()
        except Exception:
            pass
    if isinstance(o, float) and (o != o or o in (float("inf"), float("-inf"))):
        return None
    return o


def build_docx(d, rows, dist, dist_note, units, src, out):
    """Build the report and return the analysis results in report order (no patient/curation tables, no figures).
    The .docx is saved only if `out` is given."""
    res = {}
    from docx import Document
    from docx.shared import Inches, Pt
    doc = Document()
    doc.styles["Normal"].font.size = Pt(10)
    sec = doc.sections[0]  # landscape, narrower margins for the wide tables
    from docx.enum.section import WD_ORIENT
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(0.7))
    doc.add_heading("O3 Record Set Summary", 0)
    doc.add_paragraph("Dataset: %s" % d.get("O3DataSetName", "(unnamed)"))
    n = len(rows)
    res["dataset"] = {"name": d.get("O3DataSetName", ""), "patients": n,
                      "rt_courses": sum(len(p.get("RTCourse_List", [])) for p in d["Patient_List"]),
                      "images": sum(r["images"] for r in rows),
                      "systemic_therapy_courses": sum(len(p.get("SysTherCourse_List", [])) for p in d["Patient_List"])}
    doc.add_paragraph("Patients: %d   RT courses: %d   Images: %d   Systemic therapy courses: %d" % (
        n, sum(len(p.get("RTCourse_List", [])) for p in d["Patient_List"]),
        sum(r["images"] for r in rows), sum(len(p.get("SysTherCourse_List", [])) for p in d["Patient_List"])))

    def counts(title, key):
        doc.add_heading(title, 2)
        c = collections.Counter(str(r[key]) for r in rows)
        res.setdefault("distributions", {})[title] = dict(sorted(c.items()))
        t = doc.add_table(rows=1, cols=2)
        t.style = "Light Grid Accent 1"
        t.rows[0].cells[0].text, t.rows[0].cells[1].text = "Value", "Patients"
        for k, v in sorted(c.items()):
            cells = t.add_row().cells
            cells[0].text, cells[1].text = k, str(v)

    doc.add_heading("Patients", 1)
    hdr = ["ID", "ICD", "Stage", "TNM", "Fx", "Dose (Gy)", "Technique", "Age at Tx", "Chemo", "Vital"]
    t = doc.add_table(rows=1, cols=len(hdr))
    t.style = "Light Grid Accent 1"
    for i, h in enumerate(hdr):
        t.rows[0].cells[i].text = h
    for r in rows:
        vals = [r["id"], r["icd"], r["stage"], r["tnm"], r["fx"], r["dose"], r["tech"],
                None if r["age"] is None else round(r["age"], 1), "Y" if r["chemo"] else "N", r["vital"]]
        for c, v in zip(t.add_row().cells, vals):
            c.text = "" if v is None else str(v)

    doc.add_page_break()
    doc.add_heading("Curation Check", 2)
    flags = curation_flags(d, rows, patient_distances(d, rows))
    flagged = [(pid, m_, o_) for pid, m_, o_ in flags if m_ or o_]
    doc.add_paragraph("Missing or inconsistent items and statistical outliers, per patient. Outliers are values beyond 1.5 x IQR "
                      "from the quartiles across all patients (log scale for days and GTV); they are prompts to verify the source "
                      "data, not errors. %d of %d patients have at least one flag; patients with no flags are omitted." % (len(flagged), len(flags)))
    ct = doc.add_table(rows=1, cols=3)
    ct.style = "Light Grid Accent 1"
    for i, h in enumerate(["Patient", "Missing / inconsistent data", "Outliers"]):
        ct.rows[0].cells[i].text = h
    for pid, m_, o_ in flagged:
        c = ct.add_row().cells
        c[0].text, c[1].text, c[2].text = pid, "; ".join(m_) or "-", "; ".join(o_) or "-"
    doc.add_paragraph()

    doc.add_page_break()
    doc.add_heading("Summary of outcome measures and input features", 1)
    doc.add_paragraph("Values are median [25th percentile, 75th percentile] (min, max).")
    t = doc.add_table(rows=1, cols=3)
    t.style = "Light Grid Accent 1"
    for i, h in enumerate(["Measure", "Median [Q1, Q3] (min, max) or counts", "Notes"]):
        t.rows[0].cells[i].text = h
    vc = vital_counts(d)
    yr_row = ("Patients by year of RT course", ", ".join("%s %d" % (y, c) for y, c in sorted(collections.Counter(
            str(min(x.year for x in (parse_date(r.get("RTCourse_DateOfRecord")) for r in p.get("RTCourse_List", [])) if x)
                ) if any(parse_date(r.get("RTCourse_DateOfRecord")) for r in p.get("RTCourse_List", [])) else "Unknown"
            for p in d["Patient_List"]).items())))
    summary_rows = [
        ("Time from earliest diagnosis to first RT or systemic therapy course (days)", fmt_summary([r["dx_to_tx"] for r in rows], 0)),
        ("Follow-up: RT course start to last vital status date (days)", fmt_summary([r["followup"] for r in rows], 0)),
        ("Vital status: Alive / Dead / Lost to Followup (no vital status date within %d months of the patient's RT course date)" % LOST_MONTHS, fmt_vital(vc, n)),
        ("Sex at birth", ", ".join("%s %d (%.0f%%)" % (k, v, 100.0 * v / n) for k, v in sorted(collections.Counter(
            (p.get("Patient_SexAtBirth") or "Blank") for p in d["Patient_List"]).items()))),
        ("Race", ", ".join("%s %d (%.0f%%)" % (k, v, 100.0 * v / n) for k, v in sorted(collections.Counter(
            (p.get("Patient_Race") or "Blank") for p in d["Patient_List"]).items()))),
        ("Ethnicity", ", ".join("%s %d (%.0f%%)" % (k, v, 100.0 * v / n) for k, v in sorted(collections.Counter(
            (p.get("Patient_Ethnicity") or "Blank") for p in d["Patient_List"]).items()))),
        ("T stage", fmt_counts(staging_counts(d, "DiagStaging_StagingT"), n)),
        ("N stage", fmt_counts(staging_counts(d, "DiagStaging_StagingN"), n)),
        ("M stage", fmt_counts(staging_counts(d, "DiagStaging_StagingM"), n)),
        ("Overall stage", fmt_counts(staging_counts(d, "DiagStaging_Overall"), n)),
        ("GTV volume, summed over GTV structures; IGTV used if no GTV (cc)", fmt_summary([r["gtv"] for r in rows], 1)),
        ("Chemotherapy agents (patients receiving each, parsed from systemic therapy course names)", chemo_text(d, n)),
        ("Immunotherapy agents (patients receiving each, parsed from systemic therapy course names)",
         chemo_text(d, n, IMMUNO_AGENTS, "No immunotherapy")),
        ("Age at diagnosis (years)", fmt_summary([r["age_dx"] for r in rows], 1)),
        ("Fractions in RT course", fmt_summary([r["fx"] for r in rows], 0)),
        ("Maximum target dose (Gy)", fmt_summary([r["dose"] for r in rows], 1)),
    ]
    summary_rows.append(yr_row)
    dvals, dunits = [], set()
    for p in d["Patient_List"]:
        for rc in p.get("RTCourse_List", []):
            v = rc.get("RTCourse_DistanceToPatientPrimaryResidenceDirect")
            if not is_blank_distance(v) and isinstance(v, (int, float)):
                dvals.append(v)
                dunits.add(rc.get("RTCourse_DistanceToPatientDomicileDirectUnits") or "units not stated")
    dist_units = "/".join(sorted(dunits)) if dvals else units
    if dvals:
        summary_rows.append(("Distance to RT course facility (%s; from RTCourse distance field)" % "/".join(sorted(dunits)),
                             fmt_summary(dvals, 1)))
    elif dist:
        summary_rows.append(("Distance to RT course facility (%s; computed from postal codes)" % units, fmt_summary(dist, 1)))
    per, _, _ = site_patients_per_year(d)
    if per:
        summary_rows.append(("Patients treated per year at the site (HCO_SiteCount, Number of Patients)", fmt_summary(list(per.values()), 0)))
    items, uniq = treatment_units(d)
    if items:
        summary_rows.append(("Treatment machines (HCO_SiteImgTrtUnitInfo_List)",
                             "%d distinct machines (%d list items)" % (len(uniq), items)))
    res["summary"] = []
    for k, v in summary_rows:
        c = t.add_row().cells
        note = summary_note(k, d, rows, n)
        c[0].text, c[1].text, c[2].text = k, v, note
        res["summary"].append({"measure": k, "value": v, "notes": note})

    doc.add_heading("Associations", 1)
    doc.add_paragraph("Vital status: Alive / Dead / Lost to Followup as defined above. Medians are [Q1, Q3] (min, max) in days. "
                      "Small sample: p-values are exploratory and unadjusted.")
    at = doc.add_table(rows=1, cols=5)
    at.style = "Light Grid Accent 1"
    for i, h in enumerate(["Association", "Values by group", "Test", "p-value", "Notes"]):
        at.rows[0].cells[i].text = h
    assoc_rows = [
        assoc_numeric_by_vital(d, [r["dx_to_tx"] for r in rows], "Vital status vs days from diagnosis to treatment"),
        assoc_spearman([r["followup"] for r in rows], patient_distances(d, rows),
                       "Time to follow-up vs distance to RT facility", "Follow-up (days)", "Distance (%s)" % dist_units),
        assoc_spearman([r["dx_to_tx"] for r in rows], patient_distances(d, rows),
                       "Days from diagnosis to treatment vs distance to RT facility",
                       "Diagnosis to treatment (days)", "Distance (%s)" % dist_units),
        assoc_spearman([r["dx_to_tx"] for r in rows], [r["gtv"] for r in rows],
                       "Days from diagnosis to treatment vs GTV volume",
                       "Diagnosis to treatment (days)", "GTV volume (cc)"),
        assoc_spearman([r["dx_to_tx"] for r in rows], [r["age"] for r in rows],
                       "Days from diagnosis to treatment vs patient age at treatment",
                       "Diagnosis to treatment (days)", "Age at treatment (years)"),
        assoc_categorical(d, [r["dx_to_tx"] for r in rows], "Patient_Race",
                          "Race vs days from diagnosis to treatment"),
        assoc_categorical(d, [r["dx_to_tx"] for r in rows], "Patient_SexAtBirth",
                          "Sex vs days from diagnosis to treatment"),
        assoc_stage_group(d, [r["dx_to_tx"] for r in rows],
                          "Overall stage group (I-IV) vs days from diagnosis to treatment"),
    ]
    res["associations"] = [dict(zip(("association", "values_by_group", "test", "p_value", "notes"), ar)) for ar in assoc_rows]
    for ar in assoc_rows:
        for c, v in zip(at.add_row().cells, ar):
            c.text = v

    doc.add_heading("Distributions", 1)
    counts("Overall stage", "stage")
    counts("Fractions (sessions in course)", "fx")
    counts("Technique", "tech")
    counts("Vital status (as defined in the summary table)", "vital")

    if src:
        res["offline_pgeocode_sources"] = src
        doc.add_heading("Offline pgeocode source files", 1)
        doc.add_paragraph("pgeocode data folder: %s" % data_dir())
        t = doc.add_table(rows=1, cols=5)
        t.style = "Light Grid Accent 1"
        for i, h in enumerate(["Country", "Codes", "Download", "Install to", "Status"]):
            t.rows[0].cells[i].text = h
        for s in src:
            status = "present" if s["present"] else "MISSING"
            if s["missing"]:
                status += "; not found: " + ", ".join(s["missing"])
            vals = [s["country"], ", ".join(s["codes"]), s["urls"][0], s["cache_path"], status]
            for c, v in zip(t.add_row().cells, vals):
                c.text = v
        doc.add_paragraph("Install after downloading: python o3_input_report_lungpilot.py --install-source <CC>.zip")

    doc.add_heading("Histograms", 1)
    cand = [
        ("Days from diagnosis to treatment", [r["dx_to_tx"] for r in rows], "days", True),
        ("Follow-up (RT start to last vital status)", [r["followup"] for r in rows], "days", False),
        ("Age at diagnosis", [r["age_dx"] for r in rows], "years", False),
        ("Age at treatment", [r["age"] for r in rows], "years", False),
        ("Fractions in RT course", [r["fx"] for r in rows], "fractions", False),
        ("Maximum target dose", [r["dose"] for r in rows], "Gy", False),
        ("GTV volume (IGTV if no GTV)", [r["gtv"] for r in rows], "cc", True),
    ]
    dd = patient_distances(d, rows)
    cand.append(("Distance to RT facility", dd, dist_units, False))
    panels, skipped = [], []
    for title, vals, xl, lg in cand:
        v = [x for x in vals if x is not None]
        if title.startswith("Distance"):  # always plotted when any value exists, even if all identical
            if v:
                panels.append((title + (" - all identical" if len(set(v)) < 2 else ""), v, xl, lg))
            else:
                skipped.append("%s (no values)" % title)
        elif len(v) < 2:
            skipped.append("%s (fewer than 2 values)" % title)
        elif len(set(v)) < 2:
            skipped.append("%s (all %d values identical: %g)" % (title, len(v), v[0]))
        else:
            panels.append((title, v, xl, lg))
    doc.add_paragraph("Dashed red line = median. Log-scaled axes are used for skewed measures. Histograms with 20 or fewer patients "
                      "are coarse; bin choice affects the appearance.")
    import numpy as np
    res["histograms"] = []
    for i, panel in enumerate(panels):
        title, v, xl, lg = panel
        if i > 0:
            doc.add_page_break()  # one histogram per page
        doc.add_heading(title, 2)
        buf, _ = make_hist_grid([panel], ncols=1)
        doc.add_picture(buf, width=Inches(7.5))
        q1, med, q3 = np.percentile(v, [25, 50, 75])
        stat = {"histogram": title, "n": len(v), "unit": xl, "min": min(v), "q1": float(q1), "median": float(med),
                "q3": float(q3), "max": max(v)}
        line = "n=%d  median %.1f [Q1 %.1f, Q3 %.1f]  min %.1f  max %.1f %s." % (len(v), med, q1, q3, min(v), max(v), xl)
        if title.startswith("Distance"):
            dnote = ("Taken from the RTCourse distance field in the JSON." if dvals else dist_note)
            stat["note"] = dnote
            line += " " + dnote
            if len(set(round(x, 1) for x in v)) == 1:
                extra = ("All patients have the same distance, i.e. the same domicile and facility postal codes. "
                         "The domicile field may be a placeholder in this export.")
                stat["note"] += " " + extra
                line += " " + extra
        doc.add_paragraph(line)
        res["histograms"].append(stat)
    if skipped:
        doc.add_paragraph("Not plotted: " + "; ".join(skipped) + ".")
        res["histograms_not_plotted"] = skipped

    if panels:
        doc.add_page_break()  # keep the last histogram page to itself
    doc.add_heading("Analysis Plots", 1)

    def _plot(*args):
        res.setdefault("analysis_plots", []).append(add_logistic_plot(doc, *args))

    _plot("Time from diagnosis to treatment vs patient age at treatment",
                      [(r["age"], r["dx_to_tx"]) for r in rows], "Patient age at treatment (years)", "years")
    _plot("Time from diagnosis to treatment vs GTV volume",
                      [(r["gtv"], r["dx_to_tx"]) for r in rows], "GTV volume (cc; IGTV if no GTV)", "cc")
    _plot("Time from diagnosis to treatment vs distance to RT facility",
                      list(zip(patient_distances(d, rows), [r["dx_to_tx"] for r in rows])),
                      "Distance, patient domicile to RT facility (%s)" % dist_units, dist_units)

    doc.add_paragraph("\nGenerated locally. No street addresses or birth dates. Postal codes appear only in the "
                      "source-file table (--analyze-sources).")
    if out:
        doc.save(out)
    return res


# --------------------------------------------------------------------- main
HELP_EPILOG = r'''
WHAT IT DOES
  Reads an O3 record set (.json) and writes a Word report (AnalysisSummary_<YYYYMMDDHHMM>.docx) with a
  patient table, curation check, summary table, associations, distributions, histograms and fitted plots.
  Everything runs on this computer: no network calls, no uploads. pgeocode is used only when its postal-code
  data file is already installed on disk, and nothing is ever downloaded.

ORDER OF STEPS
  1. (One time) Find out which postal-code file(s) the data needs:
         python o3_input_report_lungpilot.py data.json --analyze-sources
     This lists the country file (for example US.zip), where to get it, where it must be installed and
     whether it is already present. It writes no report unless you also give -o.
  2. (One time, on a computer that may go online) download the file named in step 1 (GeoNames US.zip),
     copy it to this computer, then install it:
         python o3_input_report_lungpilot.py --install-source C:\downloads\US.zip
     Use --data-dir to install to, and read from, a different folder.
  3. (Optional) Save a copy of the JSON with the RTCourse distance fields filled in
     (value, method "Postal Code", units). Only blank or -9999 values are changed; the original is untouched:
         python o3_input_report_lungpilot.py data.json --savewmods
  4. Write the report (and optionally a JSON of the results):
         python o3_input_report_lungpilot.py data.json
         python o3_input_report_lungpilot.py data.json --results-json

OPTIONS IN DETAIL
  json                    The O3 record set to analyse. Required except with --install-source.
  -o, --output PATH       Where the Word report goes. A .docx file name is used exactly as given; a folder
                          gets AnalysisSummary_<YYYYMMDDHHMM>.docx inside it. Default: that name next to the
                          input JSON.
  --analyze-sources       Scan every postal-code field, work out the country, and list the pgeocode source
                          files needed for offline distances, with download URL, install command, install
                          path, present/missing, and any codes not found in an installed file. Adds the same
                          table to the report if a report is written.
  --install-source FILE   Install a downloaded GeoNames <CC>.zip or <CC>.txt into the pgeocode cache
                          (converted to the layout pgeocode reads), then exit.
  --country CC            Country for postal codes that could be several countries (a bare 5-digit code).
                          Default US. Codes with a US state in the address, or in Canadian/UK format, are
                          detected automatically.
  --units mi|km           Distance units for the report and for filled-in distances. Default mi.
  --savewmods             Save a copy of the input JSON, with the distance fields filled in, as
                          <input>_<YYYYMMDDHHMM>.json next to the input, then exit. Fields filled:
                          RTCourse_DistanceToPatientPrimaryResidenceDirect (value),
                          ...DirectMethod ("Postal Code") and RTCourse_DistanceToPatientDomicileDirectUnits.
                          Courses that already have a value are left alone. Distances are straight-line
                          between postal-code centroids, not street-address or driving distance.
  --fill-distances OUT    Same as --savewmods but you choose the output file. Refuses to overwrite the input.
  --results-json [FILE]   Also write the analysis results to JSON in the same order as the report
                          (dataset, summary, associations, distributions, histograms, analysis-plot fits).
                          Excludes the patients and curation tables and all figures. FILE or a folder;
                          default AnalysisResults_<YYYYMMDDHHMM>.json next to the input.
  --no-report             Do not write the Word report (use with --results-json for JSON only).
  --data-dir DIR          pgeocode data folder (sets PGEOCODE_DATA_DIR). Default ~/.cache/pgeocode.
  -h, --help              Show this guide.

EXAMPLES
  Check what postal-code data is needed, without writing a report:
      python o3_input_report_lungpilot.py data.json --analyze-sources
  Install the downloaded US postal-code file:
      python o3_input_report_lungpilot.py --install-source C:\downloads\US.zip
  Write the report next to the data file:
      python o3_input_report_lungpilot.py data.json
  Write the report to a folder, in kilometres, with a results JSON beside the input:
      python o3_input_report_lungpilot.py data.json -o C:\reports --units km --results-json
  Report plus the source-file table:
      python o3_input_report_lungpilot.py data.json --analyze-sources -o C:\reports
  Results JSON only:
      python o3_input_report_lungpilot.py data.json --results-json results.json --no-report
  Save a copy of the data with distances filled in:
      python o3_input_report_lungpilot.py data.json --savewmods

NOTES
  - Run with plain "python" (not "python -I"): python-docx may live in your user site-packages.
  - Without the postal-code file installed, the report is still written but distances are left out and
    the distance rows say so.
  - Statistics are exploratory and unadjusted; the report's Notes columns explain each test and caveat.
  - Requires: python-docx, matplotlib, scipy, numpy, pandas, pgeocode.
'''


def main():
    ap = argparse.ArgumentParser(
        prog="o3_input_report_lungpilot.py", formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Create a Word summary report (and optional results JSON) from an O3 JSON record set, entirely offline.",
        epilog=HELP_EPILOG)
    ap.add_argument("json", nargs="?", help="O3 record set .json")
    ap.add_argument("-o", "--output", help="output .docx file, or a folder (default: AnalysisSummary_YYYYMMDDHHMM.docx next to the input JSON)")
    ap.add_argument("--analyze-sources", action="store_true",
                    help="list the pgeocode source file(s) needed to compute distances offline, "
                         "and whether they are installed. With no -o, no report is written.")
    ap.add_argument("--install-source", metavar="FILE",
                    help="install a downloaded GeoNames <CC>.zip/.txt into the pgeocode cache and exit")
    ap.add_argument("--country", default="US", help="default country for ambiguous postal codes (default US)")
    ap.add_argument("--units", choices=["mi", "km"], default="mi", help="distance units (default mi)")
    ap.add_argument("--fill-distances", metavar="OUT.json",
                    help="write a copy of the JSON with RTCourse direct-distance value/method/units filled in (only where the value is blank/-9999), and exit")
    ap.add_argument("--savewmods", action="store_true",
                    help="save a copy of the input JSON with the distance fields filled in (blank values only), "
                         "named <input>_YYYYMMDDHHMM.json in the input file's folder, and exit")
    ap.add_argument("--results-json", nargs="?", const=True, metavar="FILE",
                    help="also write the analysis results (dataset counts, summary, associations, distributions, histograms, "
                         "analysis-plot fits) to JSON in report order, without the patients/curation tables or figures. "
                         "FILE or folder; default AnalysisResults_YYYYMMDDHHMM.json next to the input")
    ap.add_argument("--no-report", action="store_true", help="do not write the Word report (use with --results-json)")
    ap.add_argument("--data-dir", help="pgeocode data folder (sets PGEOCODE_DATA_DIR)")
    a = ap.parse_args()
    a.country = a.country.upper()
    if a.data_dir:
        os.environ["PGEOCODE_DATA_DIR"] = os.path.abspath(a.data_dir)

    if a.install_source:
        return install_source(a.install_source)
    if not a.json:
        ap.error("json file required")

    d = load(a.json)
    if a.savewmods and not a.fill_distances:
        import datetime
        base, ext = os.path.splitext(a.json)
        a.fill_distances = "%s_%s%s" % (base, datetime.datetime.now().strftime("%Y%m%d%H%M"), ext or ".json")
    if a.fill_distances:
        if os.path.abspath(a.fill_distances) == os.path.abspath(a.json):
            sys.exit("error: refusing to overwrite the input file")
        return fill_distances(d, a.country, a.units, a.fill_distances)
    src = None
    if a.analyze_sources:
        src = analyze_sources(d, a.country)
        print_sources(src)
        if not a.output:
            return
    import datetime
    stamp = datetime.datetime.now().strftime("%Y%m%d%H%M")
    default_name = "AnalysisSummary_%s.docx" % stamp
    if not a.output:  # default: next to the input file
        out = os.path.join(os.path.dirname(os.path.abspath(a.json)), default_name)
    elif os.path.isdir(a.output):  # -o folder: default name inside it
        out = os.path.join(a.output, default_name)
    else:  # -o file: used as given, with .docx added if missing
        out = a.output if a.output.lower().endswith(".docx") else a.output + ".docx"
    rows = patient_table(d)
    dist, note = compute_distances(rows, a.country, a.units)
    res = build_docx(d, rows, dist, note, a.units, src, None if a.no_report else out)
    if not a.no_report:
        print("wrote %s" % out)
    if a.results_json:
        base = os.path.dirname(os.path.abspath(a.json))
        jname = "AnalysisResults_%s.json" % stamp
        if a.results_json is True:
            jout = os.path.join(base, jname)
        elif os.path.isdir(a.results_json):
            jout = os.path.join(a.results_json, jname)
        else:
            jout = a.results_json
        with open(jout, "w", encoding="utf-8") as f:
            json.dump(json_clean(res), f, indent=2, ensure_ascii=False)
        print("wrote %s" % jout)
    if dist is None:
        print("note: " + note)


if __name__ == "__main__":
    main()
