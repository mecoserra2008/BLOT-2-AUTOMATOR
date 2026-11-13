"""Independent pandas re-implementation of the v5 BLOT best-execution rule.

One BLOT row = one dealer leg of a quote request.
  request  = rows whose ID is equal except the last ID_SUFFIX_LEN characters (no ID: date|ISIN|side|qty)
  unit     = the request (GROUP_LEVEL=INQUIRY) or all requests with the same date|ISIN|side (ORDER)
  asked    = distinct dealers (DEALER_MAP applied), whatever their status
  quoted   = distinct dealers whose row has a price > 0
  filled   = a row whose status contains a FILL_STATUS token
  check    = KEY? (request with two ISINs / sides / dates) | BREACH (filled, asked < MIN_Q)
             | LOW (not filled, asked < MIN_Q) | OK

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
}
ROLES = {"dl": "HDR_DEALER", "id": "HDR_ID", "sec": "HDR_SECURITY", "isin": "HDR_ISIN", "side": "HDR_SIDE",
         "st": "HDR_STATUS", "qty": "HDR_QTY", "px": "HDR_PRICE", "dt": "HDR_DATE"}


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
        if last >= 3 and header_index(h, cfg["HDR_DEALER"]) and (header_index(h, cfg["HDR_ID"]) or header_index(h, cfg["HDR_ISIN"])):
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


def analyse(block, cfg, dmap, default_td):
    hdr = block[0]
    col = {k: header_index(hdr, cfg[v]) for k, v in ROLES.items()}
    df = pd.DataFrame(block[1:], columns=range(1, len(hdr) + 1)).astype(str).apply(lambda c: c.str.strip(" "))
    g = lambda k: df[col[k]] if col[k] else pd.Series([""] * len(df), index=df.index)
    df = pd.DataFrame({k: g(k) for k in col})
    df = df[df["dl"] != ""].copy()
    n_suffix = int(cfg["ID_SUFFIX_LEN"])
    min_q = int(cfg["MIN_Q"])
    fills = [t.strip() for t in cfg["FILL_STATUS"].lower().split("|") if t.strip()]

    df["d"] = df["dt"].map(lambda s: parse_date(s, cfg["DATE_FORMAT_TEXT"]) or default_td)
    df["isinU"] = df["isin"].str.upper()
    df["bond"] = df["isinU"].where(df["isinU"] != "", df["sec"].str.upper())
    df["sideU"] = df["side"].str.upper()

    def req_key(r):
        if r["id"]:
            return "INQ:" + (r["id"][:-n_suffix] if len(r["id"]) > n_suffix + 3 and n_suffix else r["id"])
        return f"CMP:{r['d'].isoformat()}|{r['bond']}|{r['sideU']}|{r['qty']}"
    df["req"] = df.apply(req_key, axis=1)
    bad_req = set()
    for k, gg in df.groupby("req"):
        for c in ("isinU", "sideU"):
            if gg[c][gg[c] != ""].nunique() > 1:
                bad_req.add(k)
        if gg["d"].nunique() > 1:
            bad_req.add(k)
    if cfg["GROUP_LEVEL"].upper() == "ORDER":
        df["unit"] = "ORD:" + df["d"].map(date.isoformat) + "|" + df["bond"] + "|" + df["sideU"]
    else:
        df["unit"] = df["req"]
    df["dealer"] = df["dl"].map(norm_name).map(lambda n: dmap.get(n, n))
    df["quoted"] = df["px"].map(price) > 0
    df["filled"] = df["st"].str.lower().map(lambda s: bool(s) and any(t in s for t in fills))

    units = []
    for k, gg in df.groupby("unit", sort=False):
        asked = gg["dealer"].nunique()
        filled = bool(gg["filled"].any())
        winners = list(dict.fromkeys(gg.loc[gg["filled"], "dealer"]))
        if set(gg["req"]) & bad_req:
            check = "KEY?"
        elif asked < min_q:
            check = "BREACH" if filled else "LOW"
        else:
            check = "OK"
        units.append({"key": k, "date": gg["d"].iloc[0].isoformat(), "requests": gg["req"].nunique(), "asked": asked,
                      "quoted": gg.loc[gg["quoted"], "dealer"].nunique(), "filled": "YES" if filled else "NO",
                      "winner": ";".join(winners), "check": check})
    return {"units": units, "n_units": len(units), "n_requests": df["req"].nunique(), "n_legs": len(df),
            "n_filled": sum(u["filled"] == "YES" for u in units), "n_breach": sum(u["check"] == "BREACH" for u in units),
            "n_low": sum(u["check"] == "LOW" for u in units), "n_key": sum(u["check"] == "KEY?" for u in units),
            "n_quoted_low": sum(u["quoted"] < min_q for u in units)}


def compare_expected(res, exp):
    errs = [f"{k}: expected {exp[k]}, got {res[k]}" for k in
            ("n_units", "n_requests", "n_legs", "n_filled", "n_breach", "n_low", "n_key", "n_quoted_low")
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
