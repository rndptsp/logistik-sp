"""Build the encrypted data bundle for the Outbound Logistic site from the master Excel.

    python tools/build_site.py            # uses tools/config.json
    python tools/build_site.py --setup    # (re)create tools/config.json

Reads MASTER_DATA_OUTBOUND_LOGISTIC.xlsx (one sheet per dataset, append-only), aggregates it,
gzips + encrypts it (AES-GCM, key = PBKDF2-SHA256(password)) and writes data/site.enc + data/meta.json.
Never touches git: review, commit and push yourself.

Business rules (same as the old logistik-sp dashboard):
  - Realisasi date = TGL_SPJ; only rows with tonnage are counted.
  - Target = SNOP only (no RKAP/CONTRACT fallback). Target MTD = sum of daily D1..D<last data day>.
  - FOT has no target; achievement % is FRC-only. The site's filters decide incoterm/source;
    PP BELAWAN SBA is excluded by default in the site, not here (all data is shipped).
"""
import base64, datetime, gzip, json, os, re, secrets, sys, time
from collections import defaultdict

import openpyxl
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CFG_PATH = os.path.join(ROOT, "tools", "config.json")
ITERATIONS = 200_000

MONTH_PREFIX = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MEI": 5, "JUN": 6, "JUL": 7,
                "AGU": 8, "AUG": 8, "SEP": 9, "OKT": 10, "OCT": 10, "NOV": 11, "DES": 12, "DEC": 12}


def setup():
    cfg = load_cfg(required=False) or {}
    x = input(f"Path to MASTER_DATA_OUTBOUND_LOGISTIC.xlsx [{cfg.get('master_excel', '')}]: ").strip()
    if x:
        cfg["master_excel"] = x.strip('"')
    if not cfg.get("password") or input("Generate a NEW site password? [y/N]: ").lower() == "y":
        cfg["password"] = secrets.token_urlsafe(12)
        cfg.pop("salt", None)
        print("New password written to tools/config.json - share it with your team privately.")
    with open(CFG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    return cfg


def load_cfg(required=True):
    if not os.path.exists(CFG_PATH):
        if required:
            return setup()
        return None
    with open(CFG_PATH, encoding="utf-8") as f:
        return json.load(f)


def num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return 0.0


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v or "").strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def up(v):
    return str(v or "").strip().upper()


def sheet_rows(wb, name):
    """Yield dicts for a sheet, skipping a leading banner row if present."""
    it = wb[name].iter_rows(values_only=True)
    header = next(it)
    for row in it:
        if row is None or all(v is None for v in row):
            continue
        yield dict(zip(header, row))


class Dict:
    """String interning: values are stored once in a list, rows reference them by index."""
    def __init__(self):
        self.items, self.idx = [], {}

    def __call__(self, v):
        v = v or ""
        if v not in self.idx:
            self.idx[v] = len(self.items)
            self.items.append(v)
        return self.idx[v]


def build(cfg):
    t0 = time.time()
    path = cfg["master_excel"]
    print(f"Reading {path} ...")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    prov, dist, eksp, src, distr, truck = Dict(), Dict(), Dict(), Dict(), Dict(), Dict()
    dist_prov = {}

    # ---------------- Realisasi (FRC + FOT) ----------------
    facts = defaultdict(lambda: [0.0, 0, 0.0, 0])        # (day, inc, src, prov, dist, eksp) -> ton, trips, dwellSum, dwellN
    fot_dist = defaultdict(lambda: [0.0, 0])               # (month, prov, distributor) -> ton, trips
    trucks = defaultdict(lambda: [0, 0.0, 0.0, 0, 99, 0, set(), defaultdict(float)])
    last_day = {}
    n_real = 0
    for sheet, inc in (("Realisasi FRC", "FRC"), ("Realisasi H", "FOT")):
        for r in sheet_rows(wb, sheet):
            d = as_date(r.get("TGL_SPJ"))
            ton = num(r.get("TONASE"))
            if not d or ton <= 0:
                continue
            n_real += 1
            month = d.strftime("%Y-%m")
            last_day[month] = max(last_day.get(month, 0), d.day)
            p, di = prov(up(r.get("NAMA_PROPINSI"))), dist(up(r.get("NAMA_AREA_DISTRIK")))
            dist_prov.setdefault(di, p)
            e = eksp(up(r.get("EXPEDITUR_KODE")) if inc == "FRC" else "")
            s = src(up(r.get("SOURCE_PLANT")))
            dwell = r.get("DWELL_TIME_JAM")
            ok_dwell = dwell not in (None, "") and not r.get("DWELL_WAITING_FLAG") and 0 <= num(dwell) <= 48
            f = facts[(d.isoformat(), inc, s, p, di, e)]
            f[0] += ton; f[1] += 1
            if ok_dwell:
                f[2] += num(dwell); f[3] += 1
            if inc == "FOT":
                g = fot_dist[(month, p, distr(up(r.get("DISTRIBUTOR"))))]
                g[0] += ton; g[1] += 1
            nopol = up(r.get("NOPOL"))
            if nopol:
                tk = trucks[(month, truck(nopol), inc, e, s)]
                tk[0] += 1; tk[1] += ton
                if ok_dwell:
                    tk[2] += num(dwell); tk[3] += 1
                tk[4] = min(tk[4], d.day); tk[5] = max(tk[5], d.day); tk[6].add(d.day); tk[7][p] += ton
    print(f"  realisasi rows: {n_real:,}")

    # ---------------- SNOP targets ----------------
    targets = []
    for r in sheet_rows(wb, "Target Ekspeditur"):
        if up(r.get("SOURCE_TYPE")) != "SNOP":
            continue
        month = f"{int(r['TAHUN'])}-{int(r['BULAN_NO']):02d}"
        daily = [num(r.get(f"D{i}")) for i in range(1, 32)]
        has_daily = any(r.get(f"D{i}") is not None for i in (1, 15))
        p, di = prov(up(r.get("PROVINSI"))), dist(up(r.get("DISTRIK")))
        dist_prov.setdefault(di, p)
        targets.append([month, p, di, eksp(up(r.get("EXPEDITUR_KODE"))), src(up(r.get("SOURCE_PLANT"))),
                        round(num(r.get("TARGET_TON_BULAN")), 3),
                        [round(x, 3) for x in daily] if has_daily else None])
    print(f"  SNOP target rows: {len(targets):,}")

    # ---------------- Sales Order (by PERIODE month) ----------------
    so = defaultdict(float)
    for r in sheet_rows(wb, "Sales Order"):
        per = up(r.get("PERIODE"))[:3]
        mno = MONTH_PREFIX.get(per)
        d = as_date(r.get("TGL_SO"))
        if not mno:
            continue
        year = 2026 if not d else (d.year + (1 if d.month == 12 and mno == 1 else 0))
        p, di = prov(up(r.get("PROPINSI"))), dist(up(r.get("NAMA_DISTRIK")))
        dist_prov.setdefault(di, p)
        so[(f"{year}-{mno:02d}", p, di, eksp(up(r.get("EXPEDITUR_KODE"))))] += num(r.get("TONASE_KIRIM"))
    print(f"  SO groups: {len(so):,}")

    # ---------------- Latest Prognosa snapshot + SO H+1..3 ----------------
    snap_rows = list(sheet_rows(wb, "Prognosa Snapshot"))
    def snap_key(r):
        d = as_date(r.get("SNAPSHOT_DATE"))
        return ((d.isoformat() if d else ""), str(r.get("SNAPSHOT_TIME") or ""))
    latest = max(snap_key(r) for r in snap_rows) if snap_rows else ("", "")
    fields = ["antriUnit", "antriTon", "timbangUnit", "timbangTon", "bookUnit", "bookTon", "sudahRilis",
              "potensi", "prognose", "snopTarget", "so", "h1TargetSnop", "h1So", "h2TargetSnop", "h2So",
              "h3TargetSnop", "h3So"]
    prog = []
    for r in snap_rows:
        if snap_key(r) != latest:
            continue
        prog.append([up(r.get("plantName")), up(r.get("incoterm")), up(r.get("propinsi"))] +
                    [round(num(r.get(k)), 2) for k in fields])
    soh = defaultdict(float)
    for h in (1, 2, 3):
        for r in sheet_rows(wb, f"SO H{h}"):
            if snap_key(r) != latest:
                continue
            p, di = prov(up(r.get("Propinsi"))), dist(up(r.get("Nama Distrik")))
            dist_prov.setdefault(di, p)
            soh[(h, p, di, up(r.get("Incoterm")))] += num(r.get("Tonase"))
    print(f"  prognosa snapshot {latest[0]} {latest[1]}: {len(prog)} rows")

    # ---------------- Assemble ----------------
    days = sorted({k[0] for k in facts})
    day_idx = {d: i for i, d in enumerate(days)}
    bundle = {
        "v": 1,
        "built": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "asOf": days[-1] if days else None,
        "lastDay": last_day,
        "dims": {"prov": prov.items, "dist": dist.items, "eksp": eksp.items, "src": src.items,
                 "distr": distr.items, "truck": truck.items,
                 "distProv": [dist_prov.get(i, -1) for i in range(len(dist.items))]},
        "days": days,
        "facts": [[day_idx[k[0]], 0 if k[1] == "FRC" else 1, k[2], k[3], k[4], k[5],
                   round(v[0], 2), v[1], round(v[2], 2), v[3]] for k, v in facts.items()],
        "fotDist": [[k[0], k[1], k[2], round(v[0], 2), v[1]] for k, v in fot_dist.items()],
        "trucks": [[k[0], k[1], 0 if k[2] == "FRC" else 1, k[3], k[4], v[0], round(v[1], 2), round(v[2], 2), v[3],
                    v[4], v[5], len(v[6]), max(v[7], key=v[7].get)] for k, v in trucks.items()],
        "targets": targets,
        "so": [[k[0], k[1], k[2], k[3], round(v, 2)] for k, v in so.items() if v],
        "prog": {"date": latest[0], "time": latest[1], "fields": fields, "rows": prog},
        "soh": [[k[0], k[1], k[2], k[3], round(v, 2)] for k, v in soh.items()],
    }
    plain = json.dumps(bundle, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    packed = gzip.compress(plain, 9)

    # salt is fixed per password (kept in config.json) so open sessions survive data rebuilds; IV is new every build
    if not cfg.get("salt"):
        cfg["salt"] = base64.b64encode(os.urandom(16)).decode()
        with open(CFG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    salt = base64.b64decode(cfg["salt"])
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITERATIONS).derive(
        cfg["password"].encode("utf-8"))
    iv = os.urandom(12)
    ct = AESGCM(key).encrypt(iv, packed, None)
    os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
    with open(os.path.join(ROOT, "data", "site.enc"), "wb") as f:
        f.write(ct)
    meta = {"build": "build." + datetime.datetime.now().strftime("%Y.%m.%d.%H"),
            "asOf": bundle["asOf"], "salt": base64.b64encode(salt).decode(),
            "iv": base64.b64encode(iv).decode(), "iterations": ITERATIONS}
    with open(os.path.join(ROOT, "data", "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=1)
    # cache-bust the page's own css/js so browsers pick up a new build immediately
    idx = os.path.join(ROOT, "index.html")
    with open(idx, encoding="utf-8") as f:
        html = f.read()
    html = re.sub(r'(style\.css|app\.js)\?v=[^"]*', lambda mm: mm.group(1) + "?v=" + meta["build"], html)
    with open(idx, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Done in {time.time()-t0:.0f}s: {len(plain)/1e6:.1f} MB json -> {len(ct)/1e6:.1f} MB encrypted, "
          f"data through {bundle['asOf']}, {meta['build']}")


if __name__ == "__main__":
    if "--setup" in sys.argv:
        setup()
    else:
        build(load_cfg())
