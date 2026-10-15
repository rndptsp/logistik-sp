#!/usr/bin/env python3
"""Lokasi gudang/toko tujuan -> MASTER_HOME/data/gudang.csv.gz (dipanggil run_web.py bila slot "gudang" diupload).

  MASTER_HOME=<folder store> python3 tools/master/update_gudang.py --input gudang.xlsx

Sheet pertama, baris judul berisi minimal KODE_TOKO dan salah satu dari: LAT + LON, KECAMATAN, ALAMAT_LENGKAP.
ALAMAT_LENGKAP (alamat ship-to utuh dari master customer, mis. SAP/CSMS) dipakai geocode_gudang.py untuk mencari
koordinat, menggantikan ALAMAT TOKO di Rilis yang terpotong 30 karakter. Kolom lain diabaikan.
Upsert per KODE_TOKO: baris baru menimpa yang lama, toko yang tidak ada di file tetap tersimpan. Sel LAT/LON kosong
tidak menghapus koordinat yang sudah ada. Hanya jumlah baris yang dicetak (log build terbaca publik).
"""
import argparse, csv, gzip, io, os, re, sys
from pathlib import Path
import openpyxl

BASE = Path(os.environ.get('MASTER_HOME') or '.')
OUT = BASE / 'data' / 'gudang.csv.gz'
COLS = ['KODE_TOKO', 'LAT', 'LON', 'KECAMATAN', 'SUMBER', 'ALAMAT_LENGKAP']   # SUMBER: manual here; geocode_gudang.py fills the rest
IN_COLS = ['KODE_TOKO', 'LAT', 'LON', 'KECAMATAN', 'ALAMAT_LENGKAP']


def coord(v, lo, hi):
    if v in (None, ''):
        return None
    try:
        x = float(str(v).replace(',', '.').strip())
    except ValueError:
        return False
    return round(x, 6) if lo <= x <= hi else False


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--input', required=True); a = ap.parse_args()
    ws = openpyxl.load_workbook(a.input, read_only=True, data_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    head = None
    for r in rows:   # header may sit below a title row
        names = [re.sub(r'[^A-Z_]', '', str(c or '').upper().replace(' ', '_')) for c in r]
        if 'KODE_TOKO' in names:
            head = names; break
    if not head:
        print('GAGAL: kolom KODE_TOKO tidak ditemukan di sheet pertama.'); sys.exit(1)
    ix = {c: head.index(c) for c in IN_COLS if c in head}
    if not ('LAT' in ix and 'LON' in ix) and 'KECAMATAN' not in ix and 'ALAMAT_LENGKAP' not in ix:
        print('GAGAL: butuh kolom LAT dan LON, KECAMATAN, atau ALAMAT_LENGKAP.'); sys.exit(1)
    cur = {}
    if OUT.exists():
        with gzip.open(OUT, 'rt', encoding='utf-8') as f:
            cur = {r['KODE_TOKO']: {c: r.get(c, '') for c in COLS} for r in csv.DictReader(f)}
    n = bad = 0
    for r in rows:
        kode = str(r[ix['KODE_TOKO']] or '').strip().split('.')[0]
        if not kode or kode.upper() == 'CONTOH':
            continue
        old = cur.get(kode, {c: '' for c in COLS})
        lat = coord(r[ix['LAT']], -11, 8) if 'LAT' in ix else None
        lon = coord(r[ix['LON']], 94, 120) if 'LON' in ix else None
        if lat is False or lon is False or (lat is None) != (lon is None):
            bad += 1; lat = lon = None
        kec = str(r[ix['KECAMATAN']] or '').strip().upper() if 'KECAMATAN' in ix else ''
        full = re.sub(r'\s+', ' ', str(r[ix['ALAMAT_LENGKAP']] or '')).strip() if 'ALAMAT_LENGKAP' in ix else ''
        new = {'KODE_TOKO': kode, 'LAT': lat if lat is not None else old['LAT'], 'LON': lon if lon is not None else old['LON'],
               'KECAMATAN': kec or old['KECAMATAN'], 'SUMBER': 'manual' if lat is not None else old['SUMBER'],
               'ALAMAT_LENGKAP': full or old['ALAMAT_LENGKAP']}
        if full and full != old['ALAMAT_LENGKAP'] and str(new['SUMBER']).startswith('gagal'):
            new['SUMBER'] = ''   # a new full address: search this ship-to again
        if new['LAT'] or new['KECAMATAN'] or new['ALAMAT_LENGKAP']:
            cur[kode] = new; n += 1
    OUT.parent.mkdir(parents=True, exist_ok=True)
    buf = io.StringIO(); w = csv.DictWriter(buf, COLS); w.writeheader(); w.writerows(cur[k] for k in sorted(cur))
    with gzip.open(OUT, 'wt', encoding='utf-8') as f:
        f.write(buf.getvalue())
    print(f'gudang: {n} baris diproses, {bad} koordinat tidak valid (dilewati), total tersimpan {len(cur)} '
          f'({sum(1 for v in cur.values() if v["LAT"])} dengan koordinat)')


if __name__ == '__main__':
    main()
