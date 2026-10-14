"""Build the encrypted data bundle from MASTER BARU (data/rilis.csv.gz, data/so.csv.gz, 01_TARGET.xlsx).

Versi untuk master baru (3 tabel). Builder lama build_site.py (master lama) tetap sebagai cadangan.

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


INCOTERMS = ["FRC", "FOT", "FOB", "FAS", "FRC CONTAINER", "FRC KAPAL", "FOT CONTAINER"]   # index = incoterm code in the bundle (facts/ship/trucks/targets)
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


def read_store(path):
    import csv
    with gzip.open(path, "rt", encoding="utf-8", newline="") as f:
        yield from csv.DictReader(f)


def build(cfg):
    t0 = time.time()
    base = cfg.get("master_baru") or os.environ.get("MASTER_BARU") or os.path.join(os.path.dirname(ROOT), "MASTER BARU")
    print(f"Membaca master baru: {base}")
    prov, dist, eksp, src, distr, truck, toko = Dict(), Dict(), Dict(), Dict(), Dict(), Dict(), Dict()
    kem_d = Dict()
    kem_d("ZAK"); kem_d("CURAH"); kem_d("JUMBO BAG")
    kemasan = lambda v: kem_d(up(v) or "ZAK")
    toko_meta = {}
    dist_prov = {}

    # ---------------- Realisasi (data/rilis.csv.gz) ----------------
    facts = defaultdict(lambda: [0.0, 0, 0.0, 0])
    ship = defaultdict(lambda: [0.0, 0])
    trucks = defaultdict(lambda: [0, 0.0, 0.0, 0])
    last_day = {}
    n_real = 0
    check = defaultdict(int)
    toko_q = defaultdict(lambda: [0, 0])
    seen_spj = set()
    today = datetime.date.today()
    for r in read_store(os.path.join(base, "data", "rilis.csv.gz")):
        d = as_date(r.get("TGL_SPJ"))
        ton = num(r.get("TONASE"))
        if not d or ton <= 0:
            check["baris tanpa tanggal/tonase (dilewati)"] += 1
            continue
        if d >= today:
            check[f"baris tgl {d.strftime('%d-%m')} (hari ini/ke depan, dilewati: situs s.d. H-1)"] += 1
            continue
        spj = str(r.get("NO_SPJ") or "").strip().split(".")[0]
        if spj:
            if spj in seen_spj:
                check["SPJ duplikat (dilewati)"] += 1
                continue
            seen_spj.add(spj)
        inc = up(r.get("INCOTERM"))
        if inc not in INCOTERMS:
            check[f"incoterm {inc or '(kosong)'} tidak dikenal (dilewati)"] += 1
            continue
        frc = inc.startswith("FRC")
        if frc and not r.get("EXPEDITUR_KODE"):
            check["FRC tanpa kode ekspeditur"] += 1
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
        e = eksp(up(r.get("EXPEDITUR_KODE")) if frc else "")
        s = src(up(r.get("SOURCE_PLANT")))
        dwell = r.get("DWELL_TIME_JAM")
        ok_dwell = dwell not in (None, "") and not r.get("DWELL_WAITING_FLAG") and 0 <= num(dwell) <= 48
        kem = kemasan(r.get("KEMASAN"))
        f = facts[(d.isoformat(), inc, s, p, di, e, kem)]
        f[0] += ton; f[1] += 1
        if ok_dwell:
            f[2] += num(dwell); f[3] += 1
        kt = str(r.get("KODE_TOKO") or "").strip().split(".")[0]
        ti = toko(kt)
        if kt and ti not in toko_meta:
            toko_meta[ti] = [kt, str(r.get("NAMA_TOKO") or "").strip(), str(r.get("ALAMAT_TOKO") or "").strip()]
        g = ship[(d.isoformat(), inc, s, p, di, distr(up(r.get("DISTRIBUTOR"))), e, ti, kem)]
        g[0] += ton; g[1] += 1
        nopol = up(r.get("NOPOL"))
        if nopol:
            tk = trucks[(d.isoformat(), truck(nopol), inc, e, s, p, kem)]
            tk[0] += 1; tk[1] += ton
            if ok_dwell:
                tk[2] += num(dwell); tk[3] += 1
    print(f"  realisasi rows: {n_real:,}")
    for k, v in sorted(check.items()):
        print(f"    cek: {k}: {v:,}")

    # ---------------- Target (01_TARGET.xlsx: sheet Target Ekspeditur) ----------------
    targets = []
    wb = openpyxl.load_workbook(os.path.join(base, "01_TARGET.xlsx"), read_only=True, data_only=True)
    for r in sheet_rows(wb, "Target Ekspeditur"):
        inc = up(r.get("INCOTERM"))
        if inc not in INCOTERMS:
            check[f"target incoterm {inc or '(kosong)'} tidak dikenal (dilewati)"] += 1
            continue
        month = f"{int(r['TAHUN'])}-{int(r['BULAN_NO']):02d}"
        daily = [round(num(r.get(f"D{i}")), 3) for i in range(1, 32)]
        p, di = prov(up(r.get("PROVINSI"))), dist(up(r.get("DISTRIK")))
        dist_prov.setdefault(di, p)
        code = up(r.get("EXPEDITUR_KODE"))
        s = src(up(r.get("SOURCE_PLANT")))
        ton = round(num(r.get("TARGET_TON_BULAN")), 3)
        if inc.startswith("FRC"):
            targets.append([month, p, di, eksp(code), s, ton, daily, INCOTERMS.index(inc), -1])
        else:   # distributor target: EXPEDITUR_KODE holds the distributor
            targets.append([month, p, di, eksp(""), s, ton, daily, INCOTERMS.index(inc), distr(code)])
    print(f"  target rows: {len(targets):,} ({', '.join(sorted({t[0] for t in targets}))})")

    # ---------------- Sales Order (03_SO.xlsx / data/so.csv.gz): FRC only, SOCC by delivery day ----------------
    so, sod = defaultdict(float), defaultdict(float)
    for r in read_store(os.path.join(base, "data", "so.csv.gz")):
        if not up(r.get("INCOTERM")).startswith("FRC"):
            continue
        dk = as_date(r.get("TGL_KIRIM"))
        if not dk:
            continue
        p, di = prov(up(r.get("PROPINSI"))), dist(up(r.get("NAMA_DISTRIK")))
        dist_prov.setdefault(di, p)
        sod[(dk.isoformat(), p, di, eksp(up(r.get("EXPEDITUR_KODE"))), kemasan(r.get("KEMASAN")))] += num(r.get("TON_SO"))
    print(f"  SO groups: {len(sod):,} per tgl kirim (SOCC)")

    # Provinsi non-kinerja (config MASTER BARU/settings.json): target 0 per kombinasi yang punya realisasi -> tampil 100%
    try:
        with open(os.path.join(base, "config", "settings.json"), encoding="utf-8") as f:
            nonkin = [prov(p) for p in json.load(f).get("provinsi_non_kinerja", [])]
    except OSError:
        nonkin = []
    zero = set()
    for (day, inc_, s_, p_, di_, e_, k_) in facts:
        if p_ in nonkin:
            zero.add((day[:7], p_, di_, e_, s_, INCOTERMS.index(inc_))) if k_ == kem_d("ZAK") else None
    for m_, p_, di_, e_, s_, ii in sorted(zero):
        targets.append([m_, p_, di_, e_, s_, 0.0, [0.0] * 31, ii, -1 if ii in (0, 4, 5) else distr("")])
    print(f"  provinsi non-kinerja: {[prov.items[i] for i in nonkin]}; target 0: {len(zero)} baris")

    try:
        with open(os.path.join(base, "config", "ekspeditur_map.json"), encoding="utf-8") as f:
            kamus = json.load(f).get("kamus", {})
    except OSError:
        kamus = {}

    # Prognosa / SO H+1..3: sumber snapshot monitoring tidak dipakai lagi (master baru).
    latest = ("", "")
    fields, prog, soh = [], [], {}


    # ---------------- Assemble ----------------
    days = sorted({k[0] for k in facts})
    day_idx = {d: i for i, d in enumerate(days)}
    bundle = {
        "v": 1,
        "master": "baru",
        "built": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "asOf": days[-1] if days else None,
        "lastDay": last_day,
        "dims": {"inc": INCOTERMS, "prov": prov.items, "dist": dist.items, "eksp": eksp.items, "src": src.items,
                 "distr": distr.items, "truck": truck.items, "nonKin": nonkin, "kem": kem_d.items, "ekspNama": [kamus.get(n, "") for n in eksp.items],
                 "toko": [toko_meta.get(i, ["", "", ""]) for i in range(len(toko.items))],
                 "distProv": [dist_prov.get(i, -1) for i in range(len(dist.items))]},
        "days": days,
        "facts": [[day_idx[k[0]], INCOTERMS.index(k[1]), k[2], k[3], k[4], k[5],
                   round(v[0], 2), v[1], round(v[2], 2), v[3], k[6]] for k, v in facts.items()],
        # one row per day x incoterm x source x province x district x distributor x ekspeditur x ship-to
        "ship": [[day_idx[k[0]], INCOTERMS.index(k[1]), k[2], k[3], k[4], k[5], k[6], k[7], round(v[0], 2), v[1], k[8]]
                 for k, v in ship.items()],
        # per day, so any date range can be cut exactly: [day, truck, inc, eksp, src, prov, trips, ton, dwellSum, dwellN]
        "trucks": [[day_idx[k[0]], k[1], INCOTERMS.index(k[2]), k[3], k[4], k[5], v[0], round(v[1], 2), round(v[2], 2), v[3], k[6]]
                   for k, v in trucks.items()],
        "targets": targets,
        "so": [[k[0], k[1], k[2], k[3], round(v, 2)] for k, v in so.items() if v],
        "sod": [[k[0], k[1], k[2], k[3], round(v, 2), k[4]] for k, v in sod.items() if v],
        "prog": {"date": "", "time": "", "fields": [], "rows": []},
        "soh": [],
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
