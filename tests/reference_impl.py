"""Independent pandas re-implementation of the v6 BLOT best-execution rule.

One BLOT row = one dealer leg of a quote request.
  request  = rows whose ID is equal except the last ID_SUFFIX_LEN characters (no ID: date|ISIN|side|qty)
  unit     = the request (GROUP_LEVEL=INQUIRY) or all requests with the same date|ISIN|side (ORDER)
  asked    = distinct dealers (DEALER_MAP applied), whatever their status
  quoted   = distinct dealers whose row has a price > 0
  filled   = a row whose status contains a FILL_STATUS token
  excluded = rows of EXCLUDE_USERS (Username), left out first
  cohort   = a fill + the earlier unfilled requests of the same day / ISIN / side: same quantity (later fill),
             or, for an OTC fill (PNT), trade date = issue date and price = re-offer price (any quantity)
  check    = KEY? | OTC-REVIEW (OTC fill not joined) | BREACH (fill, cohort asked < MIN_Q)
             | OK-LINKED | LINKED (request joined to a fill) | LOW (not filled, asked < MIN_Q) | OK

    python tests/reference_impl.py                          # reference vs expected
    python tests/reference_impl.py --vba-dir tests/vba_out  # + reference vs VBA, unit by unit
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import re
import sys
import unicodedata
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
DEFAULTS = {
    "MIN_Q": "3", "GROUP_LEVEL": "INQUIRY", "ID_SUFFIX_LEN": "2", "DATE_FORMAT_TEXT": "DD/MM/YYYY",
    "FILL_STATUS": "done|filled|accepted|traded|executed|complete",
    "HDR_DEALER": "brkrname|broker name|broker|dealer|counterparty", "HDR_ID": "id|ticket id|trade id",
    "HDR_SECURITY": "security|description", "HDR_ISIN": "isin", "HDR_SIDE": "side|b/s", "HDR_STATUS": "status",
    "HDR_QTY": "qty (m)|qty|quantity", "HDR_PRICE": "price", "HDR_DATE": "trade dt|trade date",
    "HDR_USER": "username|user name|user", "HDR_PNT": "pnt", "HDR_ISSUE_DATE": "issue date|issue dt|iss date|iss dt",
    "HDR_REOFFER": "reoffer price|re-offer price|reoffer px|re-offer px|reoffer|re-offer",
    "HDR_TIME": "time|trade time|exec time|execution time|inquiry time|create time",
    "EXCLUDE_USERS": "ROGERIO SANTOS", "LINK_LATER_FILLS": "TRUE", "PX_TOL": "0.0001",
}
ROLES = {"dl": "HDR_DEALER", "id": "HDR_ID", "sec": "HDR_SECURITY", "isin": "HDR_ISIN", "side": "HDR_SIDE",
         "st": "HDR_STATUS", "qty": "HDR_QTY", "px": "HDR_PRICE", "dt": "HDR_DATE", "user": "HDR_USER",
         "pnt": "HDR_PNT", "iss": "HDR_ISSUE_DATE", "reo": "HDR_REOFFER", "time": "HDR_TIME"}


def norm_name(s: str) -> str:
    for ch in ("\t", "\xa0", "\r", "\n"):
        s = s.replace(ch, " ")
    return re.sub(" {2,}", " ", s).strip(" ").upper()


def header_index(hdr: list[str], aliases: str) -> int:
    names = [norm_name(h).lower() for h in hdr]
    alist = [norm_name(a).lower() for a in aliases.split("|")]
    for a in alist:
        if a and a in names:
            return names.index(a) + 1
    for a in alist:
        if len(a) >= 5:
            for j, n in enumerate(names):
                if a in n:
                    return j + 1
    return 0


def read_csv(text: str) -> list[list[str]]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")[:20]
    delim, best = ",", 0
    for cand in ("\t", ";", ","):
        mx = max((ln.count(cand) for ln in lines), default=0)
        if mx > best:
            delim, best = cand, mx
    rows = [r if r else [""] for r in csv.reader(io.StringIO(text, newline=""), delimiter=delim)]
    width = max(len(r) for r in rows)
    return [r + [""] * (width - len(r)) for r in rows]


def find_block(rows, cfg):
    for r, row in enumerate(rows[:40]):
        last = max((j + 1 for j, v in enumerate(row) if v.strip(" ")), default=0)
        h = row[:last]
        if last >= 3 and header_index(h, cfg["HDR_DEALER"]) and header_index(h, cfg["HDR_ID"]) and header_index(h, cfg["HDR_STATUS"]):
            return [x[:last] for x in rows[r:]]
    return None


def parse_date(s: str, fmt: str) -> date | None:
    s = s.strip(" ")
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    try:
        if m:
            return date(int(m[1]), int(m[2]), int(m[3]))
        order = [t[0] for t in re.findall(r"D+|M+|Y+", fmt.upper())]
        parts = re.findall(r"[0-9A-Za-z]+", s.split(" ")[0])
        if len(order) != 3 or len(parts) != 3 or not all(p.isdigit() for p in parts):
            return None
        v = dict(zip(order, map(int, parts)))
        y = v["Y"] + (2000 if v["Y"] < 100 else 0)
        return date(y, v["M"], v["D"])
    except (ValueError, KeyError):
        return None


def price(s: str) -> float:
    s = s.replace(" ", "").replace("\xa0", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if s.count(",") == 1 else s.replace(",", "")
    return float(s) if re.fullmatch(r"-?\d+(\.\d+)?", s) else 0.0


FILL_KILL = ("cancel", "reject")
TRUTHY = {"TRUE", "Y", "YES", "1", "X", "SIM", "S", "VERDADEIRO", "-1"}
EPOCH = date(1899, 12, 30)


def serial(d: date) -> float:
    return float((d - EPOCH).days)


def norm_user(s: str) -> str:
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return norm_name(s)


def id_time(id_: str, d: date) -> float:
    """Unix seconds in the first 8 hex digits after the last ':' of a Bloomberg request ID (kept if within a day)."""
    hx = id_.strip().rsplit(":", 1)[-1][:8].upper()
    if len(hx) < 8 or any(c not in "0123456789ABCDEF" for c in hx):
        return 0.0
    t = 25569.0 + int(hx, 16) / 86400.0
    return t if abs(int(t) - serial(d)) <= 1 else 0.0


def time_of(s: str, d: date) -> float:
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", s.strip())
    if m:
        h, mi, se = int(m[1]), int(m[2]), int(m[3] or 0)
        if h <= 23 and mi <= 59 and se <= 59:
            return serial(d) + (h * 3600 + mi * 60 + se) / 86400.0
    return 0.0


def analyse(block, cfg, dmap, default_td):
    hdr = block[0]
    col = {k: header_index(hdr, cfg[v]) for k, v in ROLES.items()}
    min_q, n_suffix = int(cfg["MIN_Q"]), int(cfg["ID_SUFFIX_LEN"])
    fills = [t.strip() for t in cfg["FILL_STATUS"].lower().split("|") if t.strip()]
    excl = {norm_user(x) for x in cfg["EXCLUDE_USERS"].replace(";", "|").split("|") if norm_user(x)}
    tol = float(cfg["PX_TOL"].replace(",", "."))
    link_later = cfg["LINK_LATER_FILLS"].upper() in ("TRUE", "1", "YES", "VERDADEIRO", "SIM")
    df = pd.DataFrame(block[1:], columns=range(1, len(hdr) + 1)).astype(str).apply(lambda c: c.str.strip(" "))
    get = lambda r, k: r[col[k]] if col[k] else ""
    missing = (["PNT"] if not col["pnt"] else []) + (["Username"] if not col["user"] and excl else [])

    units, order, reqs, bad = {}, [], set(), set()
    seen_isin, seen_side, seen_date = {}, {}, {}
    n_excl = n_legs = 0
    for _, r in df.iterrows():
        if col["user"] and excl and norm_user(get(r, "user")) in excl:
            n_excl += 1
            continue
        dealer = get(r, "dl")
        if not dealer:
            continue
        d = parse_date(get(r, "dt"), cfg["DATE_FORMAT_TEXT"]) or default_td
        isin, sec, side, qty, rid = get(r, "isin").upper(), get(r, "sec"), get(r, "side").upper(), get(r, "qty"), get(r, "id")
        bond = isin or sec.upper()
        if rid:
            req = "INQ:" + (rid[:-n_suffix] if len(rid) > n_suffix + 3 and n_suffix else rid)
        else:
            req = f"CMP:{d.isoformat()}|{bond}|{side}|{qty}"
        reqs.add(req)
        for store, val in ((seen_isin, isin), (seen_side, side), (seen_date, d.isoformat())):
            if val:
                if req not in store:
                    store[req] = val
                elif store[req] != val:
                    bad.add(req)
        key = f"ORD:{d.isoformat()}|{bond}|{side}" if cfg["GROUP_LEVEL"].upper() == "ORDER" else req
        if key not in units:
            units[key] = {"key": key, "date": d, "bond": bond, "side": side, "qn": price(qty), "reqs": set(),
                          "asked": set(), "quoted": set(), "filled": False, "winners": [], "px": None,
                          "t": 0.0, "otc": False, "iss": None, "reo": 0.0}
            order.append(key)
        u = units[key]
        u["reqs"].add(req)
        t = time_of(get(r, "time"), d) if col["time"] else 0.0
        if not t and col["id"]:
            t = id_time(rid, d)
        if t and (not u["t"] or t < u["t"]):
            u["t"] = t
        if col["iss"] and u["iss"] is None:
            u["iss"] = parse_date(get(r, "iss"), cfg["DATE_FORMAT_TEXT"])
        if col["reo"] and not u["reo"]:
            u["reo"] = price(get(r, "reo"))
        name = norm_name(dealer)
        name = dmap.get(name, name)
        st = get(r, "st")
        px = price(get(r, "px")) if col["px"] else 0.0
        pnt = bool(col["pnt"]) and get(r, "pnt").strip().upper() in TRUTHY
        filled = (bool(st) and any(t_ in st.lower() for t_ in fills)) or (pnt and not any(k in st.lower() for k in FILL_KILL))
        u["asked"].add(name)
        if px > 0:
            u["quoted"].add(name)
        if filled:
            u["filled"] = True
            u["otc"] = u["otc"] or pnt
            if name not in u["winners"]:
                u["winners"].append(name)
            if u["px"] is None and px > 0:
                u["px"] = px
        n_legs += 1

    # fills processed in time order (stable insertion sort; an unknown time is never "after")
    idx = list(order)
    for i in range(1, len(idx)):
        t, j = idx[i], i - 1
        while j >= 0 and units[idx[j]]["t"] and units[t]["t"] and units[idx[j]]["t"] > units[t]["t"]:
            idx[j + 1] = idx[j]
            j -= 1
        idx[j + 1] = t
    head, cohort_asked, check = {}, {}, {}
    for k in idx:
        f = units[k]
        if not f["filled"] or f["reqs"] & bad:
            continue
        cands = [c for c in order if c != k and not units[c]["filled"] and not (units[c]["reqs"] & bad) and c not in head
                 and units[c]["date"] == f["date"] and units[c]["bond"] == f["bond"] and units[c]["side"] == f["side"]
                 and (not units[c]["t"] or not f["t"] or units[c]["t"] <= f["t"])]
        sel = []
        if f["otc"] and len(f["asked"]) < min_q:
            ok = bool(col["iss"] and col["reo"]) and f["iss"] is not None and f["iss"] == f["date"] \
                and f["reo"] > 0 and f["px"] is not None and abs(f["px"] - f["reo"]) <= tol
            if ok and cands:
                sel = cands
            else:
                check[k] = "OTC-REVIEW"
        elif link_later:
            sel = [c for c in cands if f["qn"] != 0 and abs(units[c]["qn"] - f["qn"]) < 1e-7]
        if sel:
            union = set(f["asked"])
            for c in sel:
                head[c] = k
                union |= units[c]["asked"]
            cohort_asked[k] = len(union)
    out = []
    for k in order:
        u = units[k]
        if u["reqs"] & bad:
            c = "KEY?"
        elif check.get(k) == "OTC-REVIEW":
            c = "OTC-REVIEW"
        elif k in head:
            c = "LINKED"
        elif u["filled"]:
            n = cohort_asked.get(k, len(u["asked"]))
            c = "BREACH" if n < min_q else ("OK-LINKED" if k in cohort_asked else "OK")
        else:
            c = "LOW" if len(u["asked"]) < min_q else "OK"
        out.append({"key": k, "date": u["date"].isoformat(), "requests": len(u["reqs"]), "asked": len(u["asked"]),
                    "quoted": len(u["quoted"]), "filled": "YES" if u["filled"] else "NO", "winner": ";".join(u["winners"]),
                    "check": c, "cohort": k if k in cohort_asked else head.get(k, ""),
                    "cohort_asked": cohort_asked.get(k, ""), "otc": "YES" if u["otc"] else ""})
    return {"ok": True, "units": out, "n_units": len(out), "n_requests": len(reqs), "n_legs": n_legs,
            "n_filled": sum(x["filled"] == "YES" for x in out), "n_breach": sum(x["check"] == "BREACH" for x in out),
            "n_low": sum(x["check"] == "LOW" for x in out), "n_key": sum(x["check"] == "KEY?" for x in out),
            "n_otc_review": sum(x["check"] == "OTC-REVIEW" for x in out), "n_linked": len(cohort_asked),
            "n_excluded": n_excl, "missing_cols": ", ".join(missing),
            "n_quoted_low": sum(x["quoted"] < min_q and x["check"] != "LINKED" for x in out)}


def compare_expected(res, exp):
    errs = [f"{k}: expected {exp[k]}, got {res[k]}" for k in
            ("n_units", "n_requests", "n_legs", "n_filled", "n_breach", "n_low", "n_key", "n_quoted_low",
             "n_otc_review", "n_linked", "n_excluded", "missing_cols")
            if k in exp and res[k] != exp[k]]
    got = {u["key"]: u for u in res["units"]}
    if "units" in exp:
        if len(got) != len(exp["units"]):
            errs.append(f"unit count: expected {len(exp['units'])}, got {len(got)}")
        for e in exp["units"]:
            u = got.get(e["key"])
            if u is None:
                errs.append(f"missing unit {e['key']}")
                continue
            errs += [f"{e['key']} {f}: expected {e[f]!r}, got {u[f]!r}" for f in e if f != "key" and u[f] != e[f]]
    return errs


def compare_vba(res, path: Path):
    if not path.exists():
        return [f"VBA dump not found: {path}"]
    with path.open(newline="", encoding="cp1252") as f:
        vba = {r["key"]: r for r in csv.DictReader(f)}
    ref = {u["key"]: u for u in res["units"]}
    errs = []
    if set(ref) != set(vba):
        errs.append(f"keys differ: only reference {sorted(set(ref) - set(vba))}, only VBA {sorted(set(vba) - set(ref))}")
    for k in sorted(set(ref) & set(vba)):
        errs += [f"{k} {f}: reference {ref[k][f]!r}, VBA {vba[k][f]!r}" for f in ref[k] if f != "key" and str(ref[k][f]) != vba[k][f]]
    return errs


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--vba-dir", type=Path)
    args = ap.parse_args()
    failures = checked = 0
    for fx in sorted((ROOT / "fixtures").glob("*.csv")):
        spec = json.loads((ROOT / "expected" / (fx.stem + ".json")).read_text(encoding="utf-8"))
        rows = read_csv(fx.read_text(encoding="utf-8"))
        for case in spec["cases"]:
            cfg = dict(DEFAULTS)
            cfg.update({k.upper(): str(v) for k, v in case.get("config", {}).items()})
            dmap = {norm_name(a): norm_name(e) for a, e in case.get("dealer_map", spec.get("dealer_map", {})).items()}
            block = find_block(rows, cfg)
            if block is None:
                errs, res = ["header row not found"], None
            else:
                res = analyse(block, cfg, dmap, date.fromisoformat(spec["default_trade_date"]))
                errs = compare_expected(res, case["expect"])
                if args.vba_dir:
                    errs += compare_vba(res, args.vba_dir / f"{fx.stem}__{case['name']}.csv")
            checked += 1
            label = f"{fx.name} [{case['name']}]"
            if errs:
                failures += 1
                print(f"FAIL {label}")
                for e in errs:
                    print(f"     {e}")
            else:
                print(f"PASS {label}: {res['n_units']} unit(s), {res['n_filled']} filled, {res['n_breach']} breach, "
                      f"{res['n_low']} low, {res['n_key']} key?" + (" | matches VBA" if args.vba_dir else ""))
    print(f"\n{checked - failures}/{checked} case(s) passed" + (" (reference vs expected and VBA)" if args.vba_dir else ""))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
