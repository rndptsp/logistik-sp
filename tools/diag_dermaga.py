#!/usr/bin/env python3
"""One-off: which store columns carry dermaga codes (DBS/DBU/TKS...)? Prints column names + code counts only
(log is public): no row data, no names, no tonnage."""
import csv, gzip, os, re, sys, collections
home = sys.argv[1]
PAT = re.compile(r'(?<![A-Z])(DBS|DBU|TKS[\s\-_]*[TB]|TKS|DERMAGA[\s\-_]*[A-Z0-9]{1,6})(?![A-Z])')
for f in ('data/rilis.csv.gz', 'data/so.csv.gz'):
    p = os.path.join(home, f)
    if not os.path.exists(p): print(f, 'tidak ada'); continue
    hits, n = collections.defaultdict(collections.Counter), 0
    with gzip.open(p, 'rt', encoding='utf-8', errors='replace') as fh:
        rd = csv.DictReader(fh); cols = rd.fieldnames
        for r in rd:
            n += 1
            for c, v in r.items():
                if v:
                    for m in PAT.findall(str(v).upper()): hits[c][re.sub(r'[\s_]+', '-', m)] += 1
    print(f, 'baris', n, 'kolom', len(cols))
    for c, v in hits.items(): print('  kolom', c, '->', dict(v.most_common(12)))
    if not hits: print('  tidak ada kode dermaga di kolom mana pun')
