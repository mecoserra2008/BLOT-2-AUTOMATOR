"""Independent pandas re-implementation of the BLOT best-execution counting logic (CLAUDE.md section 4).

Runs every fixture in tests/fixtures against tests/expected/*.json and, with --vba-dir,
also compares deal by deal with the CSV dumps written by the VBA (DumpSelfTestResults).

    python tests/reference_impl.py                     # reference vs expected
    python tests/reference_impl.py --vba-dir tests/vba_out   # + reference vs VBA

Exit code 0 = everything matches.
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent

DEFAULTS = {
    "DEALER_MODE": "AUTO",
    "GROUP_LEVEL": "RFQ",
    "MIN_Q": "3",
    "LIST_DELIMS": ";,/|",
    "EXCL_STATUS": "cancel|cxl|reject|expired",
    "DATE_FORMAT_TEXT": "",
    "CSV_DELIM": "AUTO",
    "HDR_DEALERS": "dealers in comp|dealers in competition|dealers queried|num dealers|# dealers|no. of dealers|nb dealers|dealers|dealer|counterparties|counterparty",
    "HDR_RFQ": "inquiry id|inquiry number|rfq id|rfq number|request id|quote request id",
    "HDR_ORDER": "order id|order number|parent order id|order",
    "HDR_ID": "trade id|ticket number|ticket|tkt|order number|order id|id",
    "HDR_SEC": "isin|security|description|ticker|instrument",
    "HDR_DATE": "trade date|tradedate|date",
    "HDR_SIDE": "side|b/s|buy/sell",
    "HDR_QTY": "quantity|notional|amount|qty|size",
    "HDR_STATUS": "status|order status|state",
    "HDR_TIME": "inquiry time|rfq time|request time|create time|time",
    "HDR_TRADER": "trader|trader name|trader uuid|user",
}
ROLES = {"dl": "HDR_DEALERS", "rfq": "HDR_RFQ", "ord": "HDR_ORDER", "id": "HDR_ID", "sec": "HDR_SEC",
         "dt": "HDR_DATE", "side": "HDR_SIDE", "qty": "HDR_QTY", "st": "HDR_STATUS", "time": "HDR_TIME",
         "trd": "HDR_TRADER"}
MONTHS_EN = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
MONTHS_PT = "JANFEVMARABRMAIJUNJULAGOSETOUTNOVDEZ"


# --------------------------------------------------------------------------- text helpers
def norm_name(s: str) -> str:
    """Upper case; tab / NBSP / CR / LF -> space; repeated spaces collapsed; trimmed."""
    for ch in ("\t", "\xa0", "\r", "\n"):
        s = s.replace(ch, " ")
    return re.sub(" {2,}", " ", s).strip(" ").upper()


def norm_header(s: str) -> str:
    return norm_name(s).lower()


def is_plain_number(s: str) -> bool:
    return re.fullmatch(r"-?\d+(\.\d+)?", s.strip(" ")) is not None


# --------------------------------------------------------------------------- csv
def read_text(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8")
    if raw.startswith(b"\xff\xfe"):
        return raw[2:].decode("utf-16-le")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp1252")


def detect_delim(text: str, cfg: str) -> str:
    cfg = cfg.strip().upper()
    if cfg in (";", ","):
        return cfg
    if cfg == "TAB":
        return "\t"
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")[:20]
    best, best_n = ",", 0
    for cand in (";", ",", "\t"):
        mx = max((ln.count(cand) for ln in lines), default=0)
        if mx > best_n:
            best, best_n = cand, mx
    return best


def parse_csv(text: str, cfg_delim: str) -> list[list[str]]:
    delim = detect_delim(text, cfg_delim)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    rows = [r if r else [""] for r in csv.reader(io.StringIO(text, newline=""), delimiter=delim)]
    width = max((len(r) for r in rows), default=0)
    return [r + [""] * (width - len(r)) for r in rows]


# --------------------------------------------------------------------------- header
def header_index(hdr: list[str], aliases: str) -> int:
    """1-based column of the first alias (exact match first, then contains for aliases >= 5 chars); 0 if none."""
    names = [norm_header(h) for h in hdr]
    alist = [norm_header(a) for a in aliases.split("|")]
    for a in alist:
        if a and a in names:
            return names.index(a) + 1
    for a in alist:
        if len(a) >= 5:
            for j, n in enumerate(names):
                if a in n:
                    return j + 1
    return 0


def find_block(rows: list[list[str]], cfg: dict) -> list[list[str]] | None:
    for r, row in enumerate(rows[:40]):
        last = max((j + 1 for j, v in enumerate(row) if v.strip(" ")), default=0)
        if last >= 3 and header_index(row[:last], cfg["HDR_DEALERS"]) > 0:
            return [x[:last] for x in rows[r:]]
    return None


# --------------------------------------------------------------------------- dates
def _make_date(y: int, m: int, d: int) -> date | None:
    try:
        return date(y, m, d) if 1900 <= y <= 9999 else None
    except ValueError:
        return None


def _format_runs(fmt: str):
    """Letter runs of a D/M/Y format and whether they are separated; None if unsupported."""
    runs, seps, gaps, cur, prev_sep = [], 0, 0, "", False
    for i, ch in enumerate(fmt):
        if ch.isalpha():
            if ch not in "DMY":
                return None
            if cur and cur[0] != ch:
                runs.append(cur)
                cur = ""
                if not prev_sep:
                    gaps += 1
            cur += ch
            prev_sep = False
        else:
            if cur:
                runs.append(cur)
                cur = ""
            if i > 0:
                seps += 1
            prev_sep = True
    if cur:
        runs.append(cur)
    if len(runs) != 3 or (gaps and seps):
        return None
    separated = seps > 0
    for r in runs:
        if r not in ("D", "DD", "M", "MM", "MMM", "MMMM", "YY", "YYYY"):
            return None
        if not separated and r in ("D", "M", "MMM", "MMMM"):
            return None
    return runs, separated


def _month_from_name(t: str) -> int:
    t = t[:3].upper()
    if len(t) < 3:
        return 0
    for table in (MONTHS_EN, MONTHS_PT):
        p = table.find(t)
        if p >= 0 and p % 3 == 0:
            return p // 3 + 1
    return 0


def parse_date_text(s: str, fmt: str) -> date | None:
    s = s.strip(" ")
    if not s:
        return None
    if re.match(r"\d{4}-\d{2}-\d{2}", s[:10]) and len(s) >= 10:
        return _make_date(int(s[:4]), int(s[5:7]), int(s[8:10]))
    fmt = fmt.strip().upper()
    if not fmt:
        return None
    s = s.split(" ")[0]
    fr = _format_runs(fmt)
    if fr is None:
        return None
    runs, separated = fr
    if separated:
        parts = re.findall(r"[0-9A-Za-z]+", s)
        if len(parts) != len(runs):
            return None
    else:
        if len(s) != len(fmt):
            return None
        parts, k = [], 0
        for r in runs:
            parts.append(s[k:k + len(r)])
            k += len(r)
    y = m = d = 0
    for r, t in zip(runs, parts):
        if r[0] == "D":
            if not t.isdigit() or len(t) > 2:
                return None
            d = int(t)
        elif r[0] == "M":
            if len(r) >= 3:
                m = _month_from_name(t)
                if not m:
                    return None
            else:
                if not t.isdigit() or len(t) > 2:
                    return None
                m = int(t)
        else:
            if not t.isdigit() or len(t) != len(r):
                return None
            y = int(t) + (2000 if len(r) == 2 else 0)
    return _make_date(y, m, d)


# --------------------------------------------------------------------------- core
def analyse(block: list[list[str]], cfg: dict, dealer_map: dict, default_td: date) -> dict:
    hdr = block[0]
    col = {role: header_index(hdr, cfg[key]) for role, key in ROLES.items()}
    if col["dl"] == 0:
        return {"ok": False}
    order_level = cfg["GROUP_LEVEL"].upper() == "ORDER" and col["ord"] > 0
    if not order_level:
        col["ord"] = 0

    df = pd.DataFrame(block[1:], columns=range(1, len(hdr) + 1)).astype(str)
    df = df.apply(lambda s: s.str.strip(" "))
    get = lambda role: df[col[role]] if col[role] else pd.Series([""] * len(df), index=df.index)
    dl, rfq, ordc, tid, sec = get("dl"), get("rfq"), get("ord"), get("id"), get("sec")
    side, qty, st, tm, trd, dtc = get("side"), get("qty"), get("st"), get("time"), get("trd"), get("dt")

    child = (ordc == "") & (rfq == "") & (tid == "") & (sec == "")
    # ticket is a usable key only if every non-child dealer row without RFQ/order id carries one
    ticket_ok = col["id"] > 0 and not ((dl != "") & ~child & (ordc == "") & (rfq == "") & (tid == "")).any()

    def comp_date(v: str) -> str:
        d = parse_date_text(v, cfg["DATE_FORMAT_TEXT"])
        return d.isoformat() if d else v

    own_key = pd.Series("", index=df.index)
    own_key = own_key.mask(ordc != "", "ORD:" + ordc)
    own_key = own_key.mask((own_key == "") & (rfq != ""), "RFQ:" + rfq)
    if ticket_ok:
        own_key = own_key.mask((own_key == "") & (tid != ""), "TKT:" + tid)
    cmp = ("CMP:" + dtc.map(comp_date) + "|" + sec.str.upper() + "|" + side.str.upper() + "|" + qty
           + "|" + tm + "|" + trd.str.upper())
    own_key = own_key.mask(own_key == "", cmp)
    own_key = own_key.mask(child, None)
    key = own_key.ffill()                                   # child rows inherit the previous key
    key = key.mask(child & (dl == ""), None)                # child rows without dealer are ignored
    df["key"] = key
    rows = df[df["key"].notna()].copy()

    mode = cfg["DEALER_MODE"].upper()
    if mode not in ("ROWS", "LIST", "COUNT"):
        dealer_rows = rows[rows[col["dl"]] != ""]
        n_num = dealer_rows[col["dl"]].map(is_plain_number).sum()
        n_txt = len(dealer_rows) - n_num
        multi = (dealer_rows.groupby("key").size() > 1).any()
        mode = "COUNT" if n_num > 0 and n_num >= n_txt else ("ROWS" if multi or n_txt == 0 else "LIST")

    excl = [t.strip().lower() for t in cfg["EXCL_STATUS"].split("|") if t.strip()]
    is_live_row = lambda s: not s or not any(t in s.lower() for t in excl)
    min_q = int(float(cfg["MIN_Q"]))
    delims = cfg["LIST_DELIMS"]

    def dealers_of(cells: pd.Series) -> set[str]:
        out = set()
        for c in cells:
            if not c:
                continue
            toks = [c]
            if mode == "LIST":
                toks = re.split("[" + re.escape(delims) + "\r\n]", c) if delims else re.split("[\r\n]", c)
            for t in toks:
                u = norm_name(t)
                if u:
                    out.add(dealer_map.get(u, u))
        return out

    deals, n_dropped = [], 0
    for k, g in rows.groupby("key", sort=False):
        if not g[col["st"]].map(is_live_row).any() if col["st"] else False:
            n_dropped += 1
            continue
        secs = g[col["sec"]][g[col["sec"]] != ""] if col["sec"] else pd.Series(dtype=str)
        sides = g[col["side"]][g[col["side"]] != ""] if col["side"] else pd.Series(dtype=str)
        bad = secs.str.upper().nunique() > 1 or sides.str.upper().nunique() > 1
        if mode == "COUNT":
            nums = [int(float(c)) for c in g[col["dl"]] if is_plain_number(c)]
            n, names = max([0] + nums), ""
        else:
            ds = dealers_of(g[col["dl"]])
            n, names = len(ds), ";".join(sorted(ds))
        dates = g[col["dt"]][g[col["dt"]] != ""] if col["dt"] else pd.Series(dtype=str)
        td = parse_date_text(dates.iloc[0], cfg["DATE_FORMAT_TEXT"]) if len(dates) else None
        deals.append({
            "key": k, "date": (td or default_td).isoformat(), "fallback": td is None, "n": n,
            "status": "KEY?" if bad else ("OK" if n >= min_q else f"< {min_q}"),
            "dealers": names, "security": secs.iloc[0] if len(secs) else "",
        })
    n_below = sum(d["n"] < min_q for d in deals)
    return {"ok": True, "mode": mode, "deals": deals, "n_trades": len(deals), "n_below": n_below,
            "n_key_issue": sum(d["status"] == "KEY?" for d in deals), "n_dropped": n_dropped,
            "n_date_fallback": sum(d["fallback"] for d in deals),
            "pct_below": n_below / len(deals) if deals else 0.0}


# --------------------------------------------------------------------------- comparison
def compare_expected(res: dict, exp: dict) -> list[str]:
    errs = []
    for k in ("mode", "n_trades", "n_below", "n_key_issue", "n_dropped", "n_date_fallback"):
        if k in exp and res[k] != exp[k]:
            errs.append(f"{k}: expected {exp[k]!r}, got {res[k]!r}")
    if "pct_below" in exp and abs(res["pct_below"] - exp["pct_below"]) > 0.0005:
        errs.append(f"pct_below: expected {exp['pct_below']}, got {res['pct_below']:.4f}")
    if "deals" in exp:
        got = {d["key"]: d for d in res["deals"]}
        if len(got) != len(exp["deals"]):
            errs.append(f"deal count: expected {len(exp['deals'])}, got {len(got)} ({sorted(got)})")
        for e in exp["deals"]:
            d = got.get(e["key"])
            if d is None:
                errs.append(f"missing deal {e['key']}")
                continue
            for f in ("n", "status", "date", "dealers", "security"):
                if f in e and d[f] != e[f]:
                    errs.append(f"{e['key']} {f}: expected {e[f]!r}, got {d[f]!r}")
    return errs


def compare_vba(res: dict, path: Path) -> list[str]:
    if not path.exists():
        return [f"VBA dump not found: {path}"]
    with path.open(newline="", encoding="cp1252") as f:
        vba = {r["key"]: r for r in csv.DictReader(f)}
    errs = []
    ref = {d["key"]: d for d in res["deals"]}
    if set(ref) != set(vba):
        errs.append(f"deal keys differ: only reference {sorted(set(ref) - set(vba))}, only VBA {sorted(set(vba) - set(ref))}")
    for k in sorted(set(ref) & set(vba)):
        r, v = ref[k], vba[k]
        for f in ("date", "status", "dealers", "security"):
            if str(r[f]) != v[f]:
                errs.append(f"{k} {f}: reference {r[f]!r}, VBA {v[f]!r}")
        if str(r["n"]) != v["n"]:
            errs.append(f"{k} n: reference {r['n']}, VBA {v['n']}")
        if res["mode"] != v["mode"]:
            errs.append(f"{k} mode: reference {res['mode']}, VBA {v['mode']}")
    return errs


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--fixtures", default=ROOT / "fixtures", type=Path)
    ap.add_argument("--expected", default=ROOT / "expected", type=Path)
    ap.add_argument("--vba-dir", type=Path, help="folder with the VBA dumps (<fixture>__<case>.csv)")
    args = ap.parse_args()

    failures, checked = 0, 0
    for fx in sorted(args.fixtures.glob("*.csv")):
        spec_path = args.expected / (fx.stem + ".json")
        if not spec_path.exists():
            print(f"FAIL {fx.name}: no expected file")
            failures += 1
            continue
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        default_td = date.fromisoformat(spec.get("default_trade_date", (date.today() - timedelta(days=1)).isoformat()))
        text = read_text(fx)
        for case in spec["cases"]:
            cfg = dict(DEFAULTS)
            cfg.update({k.upper(): ("" if v is None else str(v)) for k, v in case.get("config", {}).items()})
            dmap = {norm_name(a): norm_name(e) for a, e in case.get("dealer_map", spec.get("dealer_map", {})).items()}
            block = find_block(parse_csv(text, cfg["CSV_DELIM"]), cfg)
            res = analyse(block, cfg, dmap, default_td) if block else {"ok": False}
            errs = [] if res["ok"] else ["header row / dealer column not found"]
            if res["ok"]:
                errs += compare_expected(res, case["expect"])
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
                print(f"PASS {label}: {res['n_trades']} deal(s), {res['n_below']} below min, mode {res['mode']}"
                      + (" | matches VBA" if args.vba_dir else ""))
    print(f"\n{checked - failures}/{checked} case(s) passed" + (" (reference vs expected and VBA)" if args.vba_dir else ""))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
