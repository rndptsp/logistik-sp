"""Build the encrypted data bundle for the Outbound Logistic site from the master Excel.

    python tools/build_site.py            # uses tools/config.json
    python tools/build_site.py --setup    # (re)create tools/config.json
    python tools/build_site.py --password # ganti password situs, lalu build ulang
    python tools/build_site.py --lock     # sekali: pasang kunci publik (build tanpa password)
    python tools/build_site.py --master X.xlsx   # build dari file lain (mis. di server), tanpa config.json

Reads MASTER_DATA_OUTBOUND_LOGISTIC.xlsx (one sheet per dataset, append-only), aggregates it,
gzips + encrypts it with AES-GCM and writes data/site.enc + data/meta.json. Never touches git.

Encryption, two modes:
  - Public-key lock (tools/lock.json present, the normal mode after --lock): each build makes a random data key,
    encrypts the data with it and locks the data key with an RSA public key. The matching private key is stored
    locked with the viewer password (AES-GCM, key = PBKDF2-SHA256(password)). The build needs no password, so it
    can run on a server; lock.json holds nothing secret and is committed.
  - Legacy (no lock.json): data encrypted directly with the password-derived key from config.json.

Business rules (same as the previous logistik-sp dashboard, now logistik-sp-lama):
  - Realisasi date = TGL_SPJ; only rows with tonnage are counted.
  - Target = SNOP only (no RKAP/CONTRACT fallback). Target MTD = sum of daily D1..D<last data day>.
  - Sheet "SNOP" (the SNOP export pasted as-is) is the target source for every month it covers: zak, every
    plant and incoterm (FRC/FOT/FOB/FAS), realisasi or not. FRC is split to ekspeditur x day like the "Target
    Ekspeditur" rows of the same month (allocate_snop); FOT/FOB/FAS are distributor targets, split by the
    distributors' realisasi of that incoterm. Realisasi INCOTERM may be FRC/FOT/FOB/FAS. Months not in the
    SNOP sheet keep the "Target Ekspeditur" rows (FRC). The site's filters decide incoterm/source
    (default: FRC, CP INDARUNG); all data is shipped.
"""
import base64, datetime, gzip, json, os, re, secrets, sys, time
from collections import defaultdict

import openpyxl
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CFG_PATH = os.path.join(ROOT, "tools", "config.json")
LOCK_PATH = os.path.join(ROOT, "tools", "lock.json")
ITERATIONS = 200_000
OAEP = padding.OAEP(mgf=padding.MGF1(algorithm=hashes.SHA256()), algorithm=hashes.SHA256(), label=None)


def pw_key(password, salt_b64, iterations=ITERATIONS):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=base64.b64decode(salt_b64),
                      iterations=iterations).derive(password.encode("utf-8"))


def load_lock():
    if not os.path.exists(LOCK_PATH):
        return None
    with open(LOCK_PATH, encoding="utf-8") as f:
        return json.load(f)


def write_lock(private_key, password, salt_b64):
    """Store the public key and the private key locked with the viewer password (nothing secret in the file)."""
    der = private_key.private_bytes(serialization.Encoding.DER, serialization.PrivateFormat.PKCS8,
                                    serialization.NoEncryption())
    kiv = os.urandom(12)
    lock = {"v": 2, "salt": salt_b64, "iterations": ITERATIONS,
            "pub": private_key.public_key().public_bytes(serialization.Encoding.PEM,
                                                         serialization.PublicFormat.SubjectPublicKeyInfo).decode(),
            "kiv": base64.b64encode(kiv).decode(),
            "kpriv": base64.b64encode(AESGCM(pw_key(password, salt_b64)).encrypt(kiv, der, None)).decode()}
    with open(LOCK_PATH, "w", encoding="utf-8") as f:
        json.dump(lock, f, indent=1)


def unlock_private(lock, password):
    der = AESGCM(pw_key(password, lock["salt"], lock["iterations"])).decrypt(
        base64.b64decode(lock["kiv"]), base64.b64decode(lock["kpriv"]), None)   # raises InvalidTag on a wrong password
    return serialization.load_der_private_key(der, password=None)


def make_lock():
    """One-time: create the key pair, lock the private key with the current password (same salt, so open
    sessions stay valid), write tools/lock.json, then rebuild."""
    cfg = load_cfg()
    if load_lock() and input("tools/lock.json sudah ada. Buat kunci BARU? [y/N]: ").lower() != "y":
        return
    if not cfg.get("salt"):
        cfg["salt"] = base64.b64encode(os.urandom(16)).decode()
        with open(CFG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    write_lock(rsa.generate_private_key(public_exponent=65537, key_size=3072), cfg["password"], cfg["salt"])
    print("tools/lock.json dibuat (kunci publik + kunci privat terkunci password). Membangun ulang data ...")
    build(cfg)
    print("SELESAI. Commit tools/lock.json + data/ lalu push. Password tim TIDAK berubah.")

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


def change_password():
    """Ask for a new site password (typed twice, not shown on screen), save it, then rebuild."""
    import getpass
    cfg = load_cfg()
    while True:
        pw = getpass.getpass("Password baru (min. 8 karakter, kosongkan = dibuat acak): ")
        if not pw:
            pw = secrets.token_urlsafe(12)
            print("Password acak dibuat - lihat di tools/config.json.")
            break
        if len(pw) < 8:
            print("Terlalu pendek, minimal 8 karakter.")
            continue
        if getpass.getpass("Ulangi password baru: ") != pw:
            print("Tidak sama, coba lagi.")
            continue
        break
    old_pw, lock = cfg.get("password"), load_lock()
    cfg["password"] = pw
    cfg.pop("salt", None)   # new password -> new salt
    if lock:   # same key pair, private key re-locked with the new password
        cfg["salt"] = base64.b64encode(os.urandom(16)).decode()
        write_lock(unlock_private(lock, old_pw), pw, cfg["salt"])
    with open(CFG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    print("Password disimpan. Membuat ulang data dengan password baru ...")
    build(cfg)
    print("SELESAI. Push lewat GitHub Desktop, lalu bagikan password baru ke tim secara pribadi.")


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


def dwell_hours(r):
    """Hours between TGL/JAM_MASUK and TGL/JAM_KELUAR, or None when either is missing."""
    def at(tgl, jam):
        d = as_date(tgl)
        if not d or not jam:
            return None
        try:
            h, m, *sec = [int(x) for x in str(jam).strip().split(":")]
            return datetime.datetime(d.year, d.month, d.day, h, m, sec[0] if sec else 0)
        except ValueError:
            return None
    a, b = at(r.get("TGL_MASUK"), r.get("JAM_MASUK")), at(r.get("TGL_KELUAR"), r.get("JAM_KELUAR"))
    return round((b - a).total_seconds() / 3600, 3) if a and b else None


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


INCOTERMS = ["FRC", "FOT", "FOB", "FAS"]   # index = incoterm code in the bundle (facts/ship/trucks/targets)
PLANT_ALIAS = {"CP BATURAJA SMBR": "SMBR"}   # SNOP plant name -> SOURCE_PLANT name in the realisasi sheets


def plant_name(v):
    s = re.sub(r"\s+SP$", "", up(v))
    return PLANT_ALIAS.get(s, s)


def snop_rows(wb, check):
    """Rows of sheet SNOP that become targets: (month, incoterm, provinsi, distrik, plant, ton). Zak only, every
    plant, incoterm and handling (container/polysling included), with or without realisasi yet."""
    if "SNOP" not in wb.sheetnames:
        return []
    out = []
    for r in sheet_rows(wb, "SNOP"):
        g = {up(k): v for k, v in r.items() if k}
        ton, term = num(g.get("VOLUME")), up(g.get("TERM"))
        mno = MONTH_PREFIX.get(up(g.get("PERIODE"))[:3])
        if ton <= 0 or not mno or not g.get("TAHUN"):
            continue
        if up(g.get("KEMASAN")) != "ZAK":
            check["SNOP curah/jumbo (bukan zak, tidak dipakai)"] += 1
            continue
        if term not in INCOTERMS:
            check[f"SNOP incoterm {term or '(kosong)'} (tidak dikenal, tidak dipakai)"] += 1
            continue
        out.append((f"{int(num(g['TAHUN']))}-{mno:02d}", term, up(g.get("PROVINSI")), up(g.get("SALES DISTRIK")),
                    plant_name(g.get("DELIV. PLANT")), ton))
    return out


def allocate_snop(rows, base, ship, prov, dist, eksp, src, distr, dist_prov, check):
    """SNOP (month x distrik x plant) -> target rows [month, prov, dist, eksp, src, ton, daily31, inc, distributor].
    FRC (inc 0), first match wins:
      ekspeditur split + daily shape of the Target Ekspeditur rows for the same month, distrik and plant;
      else ekspeditur shares from FRC realisasi (same month/distrik/plant, then any month, then the distrik's
      Target Ekspeditur rows, then any plant), with the daily shape of the distrik (or province) target that month;
      else one row without ekspeditur.
    FOT/FOB/FAS (inc 1-3): distributor shares from realisasi of the same incoterm (month/distrik/plant, month/distrik, distrik/plant, distrik),
      no ekspeditur; else one row without distributor.
    base = Target Ekspeditur rows [month, prov, dist, eksp, src, ton, daily|None]; months covered by SNOP replace them."""
    dkey = lambda s: re.sub("[^A-Z]", "", s)
    known = {dkey(n): n for n in dist.items if n}   # SNOP spelling -> realisasi spelling ("KAB. TAPANULI SELATAN")
    vol = defaultdict(float)
    for m, term, p, d, s, ton in rows:
        pi, di = prov(p), dist(known.get(dkey(d), d))
        dist_prov.setdefault(di, pi)
        vol[(m, INCOTERMS.index(term), pi, di, src(s))] += ton
    months = {k[0] for k in vol}
    ndays = lambda m: (datetime.date(int(m[:4]) + int(m[5:]) // 12, int(m[5:]) % 12 + 1, 1) - datetime.timedelta(days=1)).day

    # patterns from Target Ekspeditur
    mrows = defaultdict(list)                                    # (m, dist, src) -> rows
    shape = defaultdict(lambda: [0.0] * 31)                      # (m, dist, src) / (m, dist) / (m, prov) -> daily sum
    m_eksp = defaultdict(lambda: defaultdict(float))             # (m, dist) -> eksp -> ton
    for t in base:
        m, p, di, e, s, ton, daily = t
        n = ndays(m)
        dl = daily if daily else [ton / n if i < n else 0.0 for i in range(31)]
        mrows[(m, di, s)].append((e, ton, dl))
        m_eksp[(m, di)][e] += ton
        for k in ((m, di, s), (m, di), (m, p)):
            sh = shape[k]
            for i in range(31):
                sh[i] += dl[i]
    # patterns from realisasi
    no_e, no_d = eksp(""), distr("")
    r_eksp = defaultdict(lambda: defaultdict(float))
    r_distr = defaultdict(lambda: defaultdict(float))
    for (day, inc, s, p, di, dr, e, _), v in ship.items():
        m = day[:7]
        if inc == "FRC" and e != no_e:
            for k in ((m, di, s), (di, s), (m, di)):
                r_eksp[k][e] += v[0]
        elif inc != "FRC" and dr != no_d:
            for k in ((m, di, s), (m, di), (di, s), (di,)):
                r_distr[(INCOTERMS.index(inc),) + k][dr] += v[0]

    def pick(*cands):
        for c in cands:
            if c and sum(c.values()) > 0:
                return c
        return None

    def daily_of(m, p, di, s, ton):
        n = ndays(m)
        for k in ((m, di, s), (m, di), (m, p)):
            sh = shape.get(k)
            if sh and sum(sh[:n]) > 0:
                tot = sum(sh[:n])
                return [round(ton * sh[i] / tot, 3) if i < n else 0.0 for i in range(31)]
        return [round(ton / n, 3) if i < n else 0.0 for i in range(31)]

    out = [list(t) + [0, -1] for t in base if t[0] not in months]
    for (m, inc, p, di, s), V in sorted(vol.items()):
        if inc == 0:
            own = mrows.get((m, di, s))
            tot = sum(x[1] for x in own) if own else 0
            if tot > 0:
                for e, ton, dl in own:
                    out.append([m, p, di, e, s, round(V * ton / tot, 3), [round(V * x / tot, 3) for x in dl], 0, -1])
                continue
            shares = pick(r_eksp.get((m, di, s)), r_eksp.get((di, s)), m_eksp.get((m, di)), r_eksp.get((m, di)))
            if not shares:
                check["target FRC tanpa pola ekspeditur (baris tanpa ekspeditur)"] += 1
                shares = {no_e: 1.0}
            tot = sum(shares.values())
            for e, w in shares.items():
                out.append([m, p, di, e, s, round(V * w / tot, 3), daily_of(m, p, di, s, V * w / tot), 0, -1])
        else:
            shares = pick(*(r_distr.get((inc,) + k) for k in ((m, di, s), (m, di), (di, s), (di,))))
            if not shares:
                check[f"target {INCOTERMS[inc]} tanpa realisasi distributor (baris tanpa distributor)"] += 1
                shares = {no_d: 1.0}
            tot = sum(shares.values())
            for dr, w in shares.items():
                out.append([m, p, di, no_e, s, round(V * w / tot, 3), daily_of(m, p, di, s, V * w / tot), inc, dr])
    return out


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

    prov, dist, eksp, src, distr, truck, toko = Dict(), Dict(), Dict(), Dict(), Dict(), Dict(), Dict()
    toko_meta = {}   # toko index -> [kode, nama, alamat]
    dist_prov = {}

    # ---------------- Realisasi (FRC + FOT) ----------------
    facts = defaultdict(lambda: [0.0, 0, 0.0, 0])        # (day, inc, src, prov, dist, eksp) -> ton, trips, dwellSum, dwellN
    ship = defaultdict(lambda: [0.0, 0])                   # (day, inc, src, prov, dist, distributor, eksp, toko) -> ton, trips
    trucks = defaultdict(lambda: [0, 0.0, 0.0, 0])   # (day, truck, inc, eksp, src, prov) -> trips, ton, dwellSum, dwellN
    last_day = {}
    n_real = 0
    check = defaultdict(int)          # data-quality counters, printed at the end and shown on the site
    toko_q = defaultdict(lambda: [0, 0])   # (month, inc) -> [rows without ship-to, rows]
    seen_spj = set()
    today = datetime.date.today()
    for sheet, sheet_inc in (("Realisasi FRC", "FRC"), ("Realisasi H", "FOT")):
        rows = list(sheet_rows(wb, sheet))
        # learn ekspeditur name -> code from rows that have both, to fill rows pasted without the code
        name2code = {}
        for r in rows:
            if r.get("EXPEDITUR_KODE") and r.get("EXPEDITUR_NAMA"):
                name2code.setdefault(up(r["EXPEDITUR_NAMA"]), up(r["EXPEDITUR_KODE"]))
        for r in rows:
            d = as_date(r.get("TGL_SPJ"))
            ton = num(r.get("TONASE"))
            if not d or ton <= 0:
                check["baris tanpa tanggal/tonase (dilewati)"] += 1
                continue
            if d >= today:   # the site shows plan vs realisasi up to H-1; today's rows are partial
                check[f"baris tgl {d.strftime('%d-%m')} (hari ini/ke depan, dilewati: situs s.d. H-1)"] += 1
                continue
            spj = str(r.get("NO_SPJ") or "").strip().split(".")[0]
            if spj:
                if spj in seen_spj:
                    check["SPJ duplikat (dilewati)"] += 1
                    continue
                seen_spj.add(spj)
            # the row's own INCOTERM wins over the sheet it was pasted into (FRC rows pasted into the FOT sheet)
            inc = up(r.get("INCOTERM"))
            if inc not in INCOTERMS:
                inc = sheet_inc
            elif (inc == "FRC") != (sheet_inc == "FRC"):   # FOB/FAS rows in Realisasi H are expected
                check[f"{inc} di sheet {sheet} (dihitung sebagai {inc})"] += 1
            if inc == "FRC" and not r.get("EXPEDITUR_KODE"):
                code = name2code.get(up(r.get("EXPEDITUR_NAMA")))
                if code:
                    r["EXPEDITUR_KODE"] = code; check["kode ekspeditur diisi dari nama"] += 1
                else:
                    check["FRC tanpa kode ekspeditur"] += 1
            if r.get("DWELL_TIME_JAM") in (None, ""):
                dw = dwell_hours(r)
                if dw is not None:
                    r["DWELL_TIME_JAM"] = dw; check["dwell dihitung dari jam masuk/keluar"] += 1
            if not str(r.get("KODE_TOKO") or "").strip():
                check[inc + " tanpa toko tujuan"] += 1
            n_real += 1
            month = d.strftime("%Y-%m")
            tq = toko_q[(month, inc)]
            tq[1] += 1
            if not str(r.get("KODE_TOKO") or "").strip():
                tq[0] += 1
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
            kt = str(r.get("KODE_TOKO") or "").strip().split(".")[0]
            ti = toko(kt)
            if kt and ti not in toko_meta:
                toko_meta[ti] = [kt, str(r.get("NAMA_TOKO") or "").strip(), str(r.get("ALAMAT_TOKO") or "").strip()]
            g = ship[(d.isoformat(), inc, s, p, di, distr(up(r.get("DISTRIBUTOR"))), e, ti)]
            g[0] += ton; g[1] += 1
            nopol = up(r.get("NOPOL"))
            if nopol:
                tk = trucks[(d.isoformat(), truck(nopol), inc, e, s, p)]
                tk[0] += 1; tk[1] += ton
                if ok_dwell:
                    tk[2] += num(dwell); tk[3] += 1
    print(f"  realisasi rows: {n_real:,}")
    for k, v in sorted(check.items()):
        print(f"    cek: {k}: {v:,}")

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
    print(f"  Target Ekspeditur rows: {len(targets):,}")
    snop = snop_rows(wb, check)
    targets = allocate_snop(snop, targets, ship, prov, dist, eksp, src, distr, dist_prov, check)
    print(f"  sheet SNOP: {len(snop):,} rows ({', '.join(sorted({r[0] for r in snop})) or 'tidak ada'}); target rows: {len(targets):,}")

    # ---------------- Sales Order ----------------
    # With TGL_KIRIM: SO per delivery day = SOCC as in the daily report (sum of tonase by Tgl Kirim).
    # Without it: the whole PERIODE month (older rows).
    so, sod = defaultdict(float), defaultdict(float)
    for r in sheet_rows(wb, "Sales Order"):
        dk = as_date(r.get("TGL_KIRIM"))
        if dk:
            p, di = prov(up(r.get("PROPINSI"))), dist(up(r.get("NAMA_DISTRIK")))
            dist_prov.setdefault(di, p)
            sod[(dk.isoformat(), p, di, eksp(up(r.get("EXPEDITUR_KODE"))))] += num(r.get("TONASE_KIRIM"))
            continue
        per = up(r.get("PERIODE"))[:3]
        mno = MONTH_PREFIX.get(per)
        d = as_date(r.get("TGL_SO"))
        if not mno:
            continue
        year = 2026 if not d else (d.year + (1 if d.month == 12 and mno == 1 else 0))
        p, di = prov(up(r.get("PROPINSI"))), dist(up(r.get("NAMA_DISTRIK")))
        dist_prov.setdefault(di, p)
        so[(f"{year}-{mno:02d}", p, di, eksp(up(r.get("EXPEDITUR_KODE"))))] += num(r.get("TONASE_KIRIM"))
    print(f"  SO groups: {len(so):,} per bulan, {len(sod):,} per tgl kirim (SOCC)")

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
        "dims": {"inc": INCOTERMS, "prov": prov.items, "dist": dist.items, "eksp": eksp.items, "src": src.items,
                 "distr": distr.items, "truck": truck.items,
                 "toko": [toko_meta.get(i, ["", "", ""]) for i in range(len(toko.items))],
                 "distProv": [dist_prov.get(i, -1) for i in range(len(dist.items))]},
        "days": days,
        "facts": [[day_idx[k[0]], INCOTERMS.index(k[1]), k[2], k[3], k[4], k[5],
                   round(v[0], 2), v[1], round(v[2], 2), v[3]] for k, v in facts.items()],
        # one row per day x incoterm x source x province x district x distributor x ekspeditur x ship-to
        "ship": [[day_idx[k[0]], INCOTERMS.index(k[1]), k[2], k[3], k[4], k[5], k[6], k[7], round(v[0], 2), v[1]]
                 for k, v in ship.items()],
        # per day, so any date range can be cut exactly: [day, truck, inc, eksp, src, prov, trips, ton, dwellSum, dwellN]
        "trucks": [[day_idx[k[0]], k[1], INCOTERMS.index(k[2]), k[3], k[4], k[5], v[0], round(v[1], 2), round(v[2], 2), v[3]]
                   for k, v in trucks.items()],
        "targets": targets,
        "so": [[k[0], k[1], k[2], k[3], round(v, 2)] for k, v in so.items() if v],
        "sod": [[k[0], k[1], k[2], k[3], round(v, 2)] for k, v in sod.items() if v],
        "prog": {"date": latest[0], "time": latest[1], "fields": fields, "rows": prog},
        "soh": [[k[0], k[1], k[2], k[3], round(v, 2)] for k, v in soh.items()],
        "quality": {"checks": sorted(check.items()), "toko": [[k[0], k[1], v[0], v[1]] for k, v in sorted(toko_q.items())]},
    }
    plain = json.dumps(bundle, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    packed = gzip.compress(plain, 9)

    iv = os.urandom(12)
    meta = {"build": "build." + datetime.datetime.now().strftime("%Y.%m.%d.%H"), "asOf": bundle["asOf"],
            "iv": base64.b64encode(iv).decode()}
    lock = load_lock()
    if lock:
        # public-key lock: fresh data key per build, locked with the public key; no password needed here
        data_key = AESGCM.generate_key(bit_length=256)
        ct = AESGCM(data_key).encrypt(iv, packed, None)
        pub = serialization.load_pem_public_key(lock["pub"].encode())
        meta.update({"v": 2, "salt": lock["salt"], "iterations": lock["iterations"], "kiv": lock["kiv"],
                     "kpriv": lock["kpriv"], "dk": base64.b64encode(pub.encrypt(data_key, OAEP)).decode()})
    else:
        # legacy: salt is fixed per password (kept in config.json) so open sessions survive data rebuilds
        if not cfg.get("password"):
            sys.exit("Tidak ada tools/lock.json dan tidak ada password di config.json: tidak bisa mengenkripsi.")
        if not cfg.get("salt"):
            cfg["salt"] = base64.b64encode(os.urandom(16)).decode()
            with open(CFG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        ct = AESGCM(pw_key(cfg["password"], cfg["salt"])).encrypt(iv, packed, None)
        meta.update({"salt": cfg["salt"], "iterations": ITERATIONS})
    os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
    with open(os.path.join(ROOT, "data", "site.enc"), "wb") as f:
        f.write(ct)
    with open(os.path.join(ROOT, "data", "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=1)
    # cache-bust the page's own css/js so browsers pick up a new build immediately
    idx = os.path.join(ROOT, "index.html")
    with open(idx, encoding="utf-8") as f:
        html = f.read()
    html = re.sub(r'(style\.css|app\.js|sumatra\.js)\?v=[^"]*', lambda mm: mm.group(1) + "?v=" + meta["build"], html)
    with open(idx, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Done in {time.time()-t0:.0f}s: {len(plain)/1e6:.1f} MB json -> {len(ct)/1e6:.1f} MB encrypted, "
          f"data through {bundle['asOf']}, {meta['build']}")


if __name__ == "__main__":
    if "--setup" in sys.argv:
        setup()
    elif "--password" in sys.argv:
        change_password()
    elif "--lock" in sys.argv:
        make_lock()
    elif "--master" in sys.argv:   # e.g. on a build server: no config.json, encryption via tools/lock.json
        cfg = load_cfg(required=False) or {}
        cfg["master_excel"] = sys.argv[sys.argv.index("--master") + 1]
        build(cfg)
    else:
        build(load_cfg())
