#!/usr/bin/env python3
"""Master kapal -> MASTER_HOME/data/kapal.csv.gz (dipanggil run_web.py bila slot "kapal" diupload).

  MASTER_HOME=<folder store> python3 tools/master/update_kapal.py --input kapal.xlsx

Sheet pertama, baris judul berisi minimal NAMA_KAPAL (persis seperti di Rilis, tanpa awalan "KAPAL").
Kolom lain opsional (lihat COLS). Upsert per NAMA_KAPAL: sel kosong tidak menghapus isi lama, kapal yang tidak ada
di file tetap tersimpan. FOTO_URL hanya diterima dari Wikimedia Commons (berlisensi bebas). Hanya jumlah baris
yang dicetak (log build terbaca publik).
"""
import argparse, csv, gzip, io, os, re, sys
from pathlib import Path
import openpyxl

BASE = Path(os.environ.get('MASTER_HOME') or '.')
OUT = BASE / 'data' / 'kapal.csv.gz'
COLS = ['NAMA_KAPAL', 'NAMA_RESMI', 'JENIS', 'IMO', 'MMSI', 'GT', 'DWT', 'PANJANG_M', 'TAHUN_BUAT', 'OPERATOR',
        'PELABUHAN_ASAL', 'FOTO_URL', 'FOTO_KREDIT', 'SUMBER', 'KEYAKINAN', 'CATATAN']
FOTO_OK = re.compile(r'^https://(upload\.wikimedia\.org|commons\.wikimedia\.org)/')


def key(v):
    return re.sub(r'\s+', ' ', re.sub(r'^\s*KAPAL\s+', '', str(v or '').upper())).strip()


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--input', required=True); a = ap.parse_args()
    ws = openpyxl.load_workbook(a.input, read_only=True, data_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    head = None
    for r in rows:   # header may sit below a title row
        names = [re.sub(r'[^A-Z_]', '', str(c or '').upper().replace(' ', '_').replace('/', '_')) for c in r]
        if 'NAMA_KAPAL' in names:
            head = names; break
    if not head:
        print('GAGAL: kolom NAMA_KAPAL tidak ditemukan di sheet pertama.'); sys.exit(1)
    ix = {c: head.index(c) for c in COLS if c in head}
    cur = {}
    if OUT.exists():
        with gzip.open(OUT, 'rt', encoding='utf-8') as f:
            cur = {r['NAMA_KAPAL']: {c: r.get(c, '') for c in COLS} for r in csv.DictReader(f)}
    n = bad = 0
    for r in rows:
        nm = key(r[ix['NAMA_KAPAL']])
        if not nm or nm == 'CONTOH':
            continue
        old = cur.get(nm, {c: '' for c in COLS})
        new = dict(old, NAMA_KAPAL=nm)
        for c in COLS[1:]:
            if c not in ix:
                continue
            v = re.sub(r'\s+', ' ', str(r[ix[c]] if r[ix[c]] is not None else '')).strip()
            if re.fullmatch(r'\d+\.0', v):
                v = v[:-2]
            if c == 'FOTO_URL' and v and not FOTO_OK.match(v):
                bad += 1; continue
            if v:
                new[c] = v
        cur[nm] = new; n += 1
    OUT.parent.mkdir(parents=True, exist_ok=True)
    buf = io.StringIO(); w = csv.DictWriter(buf, COLS); w.writeheader(); w.writerows(cur[k] for k in sorted(cur))
    with gzip.open(OUT, 'wt', encoding='utf-8') as f:
        f.write(buf.getvalue())
    print(f'kapal: {n} baris diproses, {bad} URL foto bukan Wikimedia (dilewati), total tersimpan {len(cur)}')


if __name__ == '__main__':
    main()
