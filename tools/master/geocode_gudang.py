#!/usr/bin/env python3
"""Cari koordinat toko/gudang tujuan dari nama + alamat (dipanggil run_web.py setiap build; gagal tidak menghentikan build).

  MASTER_HOME=<folder store> [GOOGLE_MAPS_KEY=...] python3 tools/master/geocode_gudang.py

Sumber toko: data/rilis.csv.gz (KODE_TOKO, NAMA_TOKO, ALAMAT_TOKO, distrik, provinsi); ALAMAT_LENGKAP yang diupload
(slot gudang) dipakai bila ada, karena ALAMAT TOKO di Rilis terpotong 30 karakter. Hasil: data/gudang.csv.gz
(kolom SUMBER: manual | google-tempat | google-alamat | osm | gagal:<penyedia>). Koordinat manual (upload admin) tidak
pernah ditimpa; toko yang sudah dicari tidak dicari ulang kecuali penyedianya lebih baik (Google setelah OSM).
Penyedia: Google Maps (Places Text Search, lalu Geocoding) bila GOOGLE_MAPS_KEY ada; selain itu OpenStreetMap Nominatim
(gratis, 1 permintaan/detik, dibatasi waktu per build). Setiap hasil harus jatuh DI DALAM batas kabupaten/kota toko itu
(assets/districts.js), kalau tidak ditolak. Hasil yang hanya menunjuk pusat kota/kabupaten juga ditolak.
Hanya jumlah yang dicetak.
"""
import collections, csv, gzip, io, json, os, re, sys, time, urllib.parse, urllib.request
from pathlib import Path

BASE = Path(os.environ.get('MASTER_HOME') or '.')
ROOT = Path(__file__).resolve().parents[2]
OUT = BASE / 'data' / 'gudang.csv.gz'
COLS = ['KODE_TOKO', 'LAT', 'LON', 'KECAMATAN', 'SUMBER', 'ALAMAT_LENGKAP']
KEY = os.environ.get('GOOGLE_MAPS_KEY', '').strip()
PROV = 'google' if KEY else 'osm'
LIMIT = int(os.environ.get('GEO_MAX') or 3000)
BUDGET = float(os.environ.get('GEO_SECONDS') or 8 * 60)   # time budget per build (OSM: ~1-2 s per ship-to); the rest waits for the next build
UA = {'User-Agent': 'logistik-sp-dashboard/1.0 (internal company dashboard; geocoding ship-to addresses)'}
ALIAS = {'RIAU KEPULAUAN': 'KEPULAUAN RIAU'}
ABBR = {'TANJ': 'TANJUNG', 'TJ': 'TANJUNG', 'BRT': 'BARAT', 'TIM': 'TIMUR', 'TMR': 'TIMUR', 'SEL': 'SELATAN', 'SLT': 'SELATAN',
        'UTR': 'UTARA', 'TGH': 'TENGAH', 'TNG': 'TENGAH', 'KEP': 'KEPULAUAN'}


def up(v):
    return re.sub(r'\s+', ' ', str(v or '')).strip().upper()


# ---------- district borders (same file and projection as the site: x = (lon-95)*200, y = (6.2-lat)*200) ----------
def load_districts():
    s = (ROOT / 'assets' / 'districts.js').read_text(encoding='utf-8')
    return json.loads(s[s.index('=') + 1:].strip().rstrip(';'))


def dkey(name):
    s = up(name)
    s = re.sub(r'^KAB(UPATEN)?\b\.?\s*', '', s)
    s = re.sub(r'[A-Z]+', lambda m: ABBR.get(m.group(0), m.group(0)), s)
    kota = bool(re.match(r'^KOTA\b', s))
    s = re.sub(r'^KOTA\b\.?\s*', '', s)
    return kota, re.sub('[^A-Z]', '', s)


def find_shape(D, prov, dist):
    shapes = (D.get(ALIAS.get(prov, prov)) or {}).get('d', [])
    kota, q = dkey(dist)
    best = None
    for g in shapes:
        gk = g['n'].startswith('Kota ')
        t = re.sub('[^A-Z]', '', up(g['n'][5:] if gk else g['n']))
        if gk != kota:
            continue
        if t == q:
            return g
        if best is None and len(q) >= 5 and t.startswith(q):
            best = g
    return best


def rings(g):
    return [[tuple(map(float, p.split(','))) for p in r.rstrip('Z').split('L') if p] for r in g['d'].split('M') if r]


def inside(lat, lon, rs):
    x, y, hit = (lon - 95) * 200, (6.2 - lat) * 200, False
    for r in rs:
        for (x0, y0), (x1, y1) in zip(r, r[1:] + r[:1]):
            if (y0 > y) != (y1 > y) and x < (x1 - x0) * (y - y0) / (y1 - y0) + x0:
                hit = not hit
    return hit


def bbox(rs):
    xs = [x for r in rs for x, _ in r]; ys = [y for r in rs for _, y in r]
    return 6.2 - max(ys) / 200, 95 + min(xs) / 200, 6.2 - min(ys) / 200, 95 + max(xs) / 200   # S, W, N, E


# ---------- providers ----------
def http(url, data=None, headers=None):
    req = urllib.request.Request(url, data=data, headers=dict(UA, **(headers or {})))
    return json.loads(urllib.request.urlopen(req, timeout=30).read())


def google(t, rs, bb):
    s, w, n, e = bb
    q = ' '.join(x for x in (t['nama'], t['alamat'], t['dist'].title(), t['prov'].title()) if x)
    try:
        r = http('https://places.googleapis.com/v1/places:searchText',
                 json.dumps({'textQuery': q, 'regionCode': 'ID', 'languageCode': 'id',
                             'locationRestriction': {'rectangle': {'low': {'latitude': s, 'longitude': w}, 'high': {'latitude': n, 'longitude': e}}}}).encode(),
                 {'Content-Type': 'application/json', 'X-Goog-Api-Key': KEY, 'X-Goog-FieldMask': 'places.location,places.types'})
        for p in r.get('places', []):
            lat, lon = p['location']['latitude'], p['location']['longitude']
            if 'political' not in p.get('types', []) and inside(lat, lon, rs):
                return lat, lon, 'google-tempat'
    except Exception:   # noqa: BLE001 - fall through to Geocoding
        pass
    if t['alamat']:
        r = http('https://maps.googleapis.com/maps/api/geocode/json?' + urllib.parse.urlencode({
            'address': ', '.join(x for x in (t['alamat'], t['dist'].title(), t['prov'].title(), 'Indonesia') if x),
            'components': 'country:ID', 'bounds': f'{s},{w}|{n},{e}', 'key': KEY}))
        for g in r.get('results', []):
            loc, kind = g['geometry']['location'], g['geometry'].get('location_type')
            if kind != 'APPROXIMATE' and inside(loc['lat'], loc['lng'], rs):
                return loc['lat'], loc['lng'], 'google-alamat'
    return None


def osm(t, rs, bb):
    s, w, n, e = bb
    for q in (f"{t['alamat']}, {t['dist'].title()}" if t['alamat'] else None, t['nama']):
        if not q:
            continue
        time.sleep(1.1)   # Nominatim usage policy: max 1 request/second
        r = http('https://nominatim.openstreetmap.org/search?' + urllib.parse.urlencode({
            'q': q, 'format': 'jsonv2', 'countrycodes': 'id', 'limit': 3, 'bounded': 1, 'viewbox': f'{w},{n},{e},{s}'}))
        for p in r:
            if p.get('category') == 'boundary' or p.get('addresstype') in ('city', 'county', 'state', 'municipality', 'regency'):
                continue
            lat, lon = float(p['lat']), float(p['lon'])
            if inside(lat, lon, rs):
                return lat, lon, 'osm'
    return None


def main():
    rilis = BASE / 'data' / 'rilis.csv.gz'
    if not rilis.exists():
        print('geocode: rilis tidak ada, dilewati'); return
    seen = collections.defaultdict(collections.Counter); meta = {}
    with gzip.open(rilis, 'rt', encoding='utf-8') as f:
        for r in csv.DictReader(f):
            k = str(r.get('KODE_TOKO') or '').strip().split('.')[0]
            if not k:
                continue
            seen[k][(up(r.get('NAMA_PROPINSI')), up(r.get('NAMA_AREA_DISTRIK')))] += 1
            meta.setdefault(k, (str(r.get('NAMA_TOKO') or '').strip(), str(r.get('ALAMAT_TOKO') or '').strip()))
    cur = {}
    if OUT.exists():
        with gzip.open(OUT, 'rt', encoding='utf-8') as f:
            cur = {r['KODE_TOKO']: {c: r.get(c, '') for c in COLS} for r in csv.DictReader(f)}
    D = load_districts()
    todo = []
    for k, c in seen.items():
        o = cur.get(k, {})
        src = o.get('SUMBER', '') or ('manual' if o.get('LAT') else '')
        if src == 'manual' or src.startswith('google') or src == 'gagal:' + PROV or (src == 'osm' and PROV == 'osm'):
            continue
        prov, dist = c.most_common(1)[0][0]
        todo.append({'kode': k, 'nama': meta[k][0], 'alamat': o.get('ALAMAT_LENGKAP') or meta[k][1], 'prov': prov, 'dist': dist, 'n': sum(c.values())})
    todo.sort(key=lambda t: -t['n'])   # busiest ship-tos first
    found = failed = noshape = 0
    t0 = time.time()
    for t in todo[:LIMIT]:
        if time.time() - t0 > BUDGET:
            print('geocode: batas waktu build tercapai, sisanya dilanjutkan build berikutnya'); break
        g = find_shape(D, t['prov'], t['dist'])
        if not g:
            noshape += 1; continue
        rs = rings(g)
        try:
            hit = (google if KEY else osm)(t, rs, bbox(rs))
        except Exception as e:   # noqa: BLE001 - network/quota trouble: stop, keep what we have
            print('geocode: berhenti,', type(e).__name__); break
        o = cur.setdefault(t['kode'], {c: '' for c in COLS})
        o['KODE_TOKO'] = t['kode']
        if hit:
            o['LAT'], o['LON'], o['SUMBER'] = round(hit[0], 6), round(hit[1], 6), hit[2]; found += 1
        else:
            o['SUMBER'] = 'gagal:' + PROV; failed += 1
    OUT.parent.mkdir(parents=True, exist_ok=True)
    buf = io.StringIO(); w = csv.DictWriter(buf, COLS); w.writeheader(); w.writerows(cur[k] for k in sorted(cur))
    with gzip.open(OUT, 'wt', encoding='utf-8') as f:
        f.write(buf.getvalue())
    print(f'geocode ({PROV}): {len(todo)} toko perlu dicari, {min(len(todo), LIMIT)} dicoba, {found} ketemu, {failed} tidak ketemu, '
          f'{noshape} distrik tanpa batas peta; total berkoordinat {sum(1 for v in cur.values() if v["LAT"])}')


if __name__ == '__main__':
    main()
