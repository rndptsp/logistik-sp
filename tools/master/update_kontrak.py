#!/usr/bin/env python3
"""Kontrak angkutan (rute pemenang, kuota, estimasi volume, TOA) -> MASTER_HOME/data/kontrak.csv.gz.

  MASTER_HOME=<folder store> python3 tools/master/update_kontrak.py --input kontrak.xlsx

Sheet pertama, satu baris per ekspeditur x rute pemenang, kolom = COLS. Seluruh tabel diganti (bukan upsert):
kontrak adalah satu dokumen utuh; amandemen ditulis di kolom AMEND_* atau BERLAKU_SAMPAI pada baris yang sama.
Estimasi volume ekspeditur = volume rute (PPK Lampiran 1) x kuota (Monitoring), kecuali diubah amandemen.
Isi kontrak rahasia: hanya jumlah baris yang dicetak (log build terbaca publik).
"""
import argparse, csv, gzip, io, os, re, sys
from pathlib import Path
import openpyxl

BASE = Path(os.environ.get('MASTER_HOME') or '.')
OUT = BASE / 'data' / 'kontrak.csv.gz'
COLS = ['EKSP_KODE', 'PLANT', 'PROVINSI', 'DISTRIK', 'SUB', 'KUOTA', 'VOL_RUTE_THN', 'VOL_EKSP_THN', 'LEAD_HARI', 'CYCLE_HARI',
        'TRUK_DEDICATED_RUTE', 'TOA', 'BERLAKU_MULAI', 'BERLAKU_SAMPAI', 'AMEND_VOL_THN', 'AMEND_MULAI', 'CATATAN']


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--input', required=True); a = ap.parse_args()
    ws = openpyxl.load_workbook(a.input, read_only=True, data_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    head = None
    for r in rows:
        names = [re.sub(r'[^A-Z_]', '', str(c or '').upper().replace(' ', '_')) for c in r]
        if 'EKSP_KODE' in names and 'DISTRIK' in names:
            head = names; break
    if not head:
        print('GAGAL: kolom EKSP_KODE dan DISTRIK tidak ditemukan di sheet pertama.'); sys.exit(1)
    ix = {c: head.index(c) for c in COLS if c in head}
    out = []
    for r in rows:
        if not r or not r[ix['EKSP_KODE']]:
            continue
        row = {}
        for c in COLS:
            v = r[ix[c]] if c in ix else None
            v = '' if v is None else (v.strftime('%Y-%m') if hasattr(v, 'strftime') else str(v).strip())
            row[c] = v
        if row['EKSP_KODE'].upper() == 'CONTOH':
            continue
        out.append(row)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    buf = io.StringIO(); w = csv.DictWriter(buf, COLS); w.writeheader(); w.writerows(out)
    with gzip.open(OUT, 'wt', encoding='utf-8') as f:
        f.write(buf.getvalue())
    print(f'kontrak: {len(out)} baris ekspeditur x rute, {len({r["EKSP_KODE"] for r in out})} ekspeditur')


if __name__ == '__main__':
    main()
