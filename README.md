# Outbound Logistics — PT Semen Padang

Dashboard pengiriman semen (FRC & FOT) per provinsi, distrik dan ekspeditur. Situs statis untuk GitHub Pages;
seluruh data terenkripsi (AES-GCM) di `data/site.enc` dan hanya terbuka dengan kata sandi.

## Isi repo

| Path | Isi |
|---|---|
| `index.html`, `style.css`, `app.js` | Situs (tampilan concept v3, SIG brand) |
| `assets/` | Logo & foto |
| `data/site.enc`, `data/meta.json` | Data terenkripsi + parameter dekripsi (dibuat oleh build) |
| `tools/build_site.py` | Membaca master Excel → membuat `data/` |
| `tools/config.json` | **Tidak di-commit.** Lokasi master Excel + kata sandi |

## Update data (rutin)

1. Tambahkan baris baru di **MASTER_DATA_OUTBOUND_LOGISTIC.xlsx** (satu sheet per jenis data, jangan menimpa baris lama).
2. Jalankan build:
   - Windows: double-click `tools\run_build.bat`
   - Mac: `python3 tools/build_site.py`
3. Cek hasil di browser (opsional): `python3 -m http.server` lalu buka http://localhost:8000
4. Commit & push `data/` dan `index.html` (`git pull` dulu sebelum mulai).

Build pertama kali akan menanyakan lokasi master Excel bila `tools/config.json` belum ada.
Ganti kata sandi: `python tools/build_site.py --setup` → jawab `y` → build → push, lalu bagikan kata sandi baru secara pribadi.

Kebutuhan: Python 3.9+, `pip install openpyxl cryptography`.

## Aturan perhitungan

- Tanggal realisasi = `TGL_SPJ`.
- Target = **SNOP saja**. Target MTD = jumlah target harian (kolom D1..D31) dari tgl 1 s.d. hari data terakhir.
- Capaian % = realisasi **FRC** ÷ target SNOP MTD. FOT tidak memiliki target.
- Default filter: FRC, tanpa PP Belawan SBA (bisa diubah di menu Filter).
- Real/SO = realisasi FRC ÷ SO (sheet Sales Order, per kolom PERIODE). SO per ekspeditur = atribusi.
- Forecast harian = snapshot Prognosa terakhir di master Excel; mingguan/bulanan = proyeksi laju rata-rata.

Angka September 2026 sudah dicocokkan dengan *Report Pengiriman OL MTD 30 September 2026* (per provinsi, ekspeditur dan minggu).
