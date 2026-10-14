#!/usr/bin/env python3
"""Update master dari file yang diupload lewat /admin (dipanggil oleh .github/workflows/build.yml).

  MASTER_HOME=<folder store>  python3 tools/master/run_web.py --inbox <folder> [--month N] [--force]

<folder store> berisi data/*.csv.gz dan config/*.json (diunduh dari R2, bukan dari repo).
<folder inbox> berisi file upload (nama tetap per slot): rilis.xlsx, so.xlsx, tactical.xlsx, snop.xlsx, monitoring.xlsx.
Urutan: Target (jika tactical+snop+monitoring ada) -> tulis 01_TARGET.xlsx -> Rilis -> SO -> rapikan.
Target harus lebih dulu: proporsi ekspeditur SO dibaca dari 01_TARGET.xlsx. Berhenti di langkah pertama yang gagal.
Slot yang tidak diupload dilewati; tanpa satu pun slot, hanya menulis ulang 01_TARGET.xlsx (untuk build ulang / rollback).
"""
import argparse, datetime as dt, os, subprocess, sys
from pathlib import Path
HERE = Path(__file__).resolve().parent

def run(args, env):
    print('\n$', ' '.join(str(x) for x in args), flush=True)
    return subprocess.call([sys.executable, '-I'] + [str(x) for x in args], cwd=HERE, env=env) == 0

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--inbox', required=True); ap.add_argument('--month', type=int); ap.add_argument('--force', action='store_true')
    a = ap.parse_args()
    opt = Path(a.inbox) / '_options.json'
    if opt.exists():
        import json; o = json.load(open(opt))
        a.month = a.month or (int(o['month']) if o.get('month') else None); a.force = a.force or bool(o.get('force'))
    home = os.environ.get('MASTER_HOME')
    if not home or not (Path(home) / 'data' / 'rilis.csv.gz').exists():
        print('GAGAL: store master (MASTER_HOME/data/rilis.csv.gz) tidak ada. Muat master awal lewat halaman admin dulu.'); sys.exit(1)
    for d in ('logs', 'data/backup'): (Path(home) / d).mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    inbox = Path(a.inbox); f = {k: inbox / (k + '.xlsx') for k in ('rilis', 'so', 'tactical', 'snop', 'monitoring')}
    have = {k: p.exists() for k, p in f.items()}
    part = [k for k in ('tactical', 'snop', 'monitoring') if have[k]]
    if part and len(part) < 3:
        print('GAGAL: Target butuh tiga file (Tactical, SNOP, Monitoring); yang ada baru:', ', '.join(part)); sys.exit(1)
    mn = a.month or dt.datetime.now().month
    fr = ['--force'] if a.force else []
    steps = []
    if part:
        steps.append(('target', ['update_target.py', 'target', '--month', mn, '--tactical', f['tactical'], '--snop', f['snop'], '--monitoring', f['monitoring']]))
    steps.append(('target_xlsx', ['update_target.py', 'views']))
    if have['rilis']: steps.append(('rilis', ['update_master.py', 'rilis', '--input', f['rilis']] + fr))
    if have['so']: steps.append(('so', ['update_master.py', 'so', '--input', f['so']] + fr))
    steps.append(('rapikan', ['update_master.py', 'rapikan']))
    print('Rencana:', ' -> '.join(n for n, _ in steps), '| bulan target:', mn)
    for n, args in steps:
        if not run(args, env):
            print('\nBERHENTI di langkah', n, '- data store tidak dipublikasikan.'); sys.exit(1)
    print('\nSELESAI (master)')
if __name__ == '__main__': main()
