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
| `assets/districts.js`, `tools/make_districts.py` | Batas kabupaten/kota (geoBoundaries, CC BY 4.0) untuk peta distrik di halaman provinsi/distrik. Dibuat sekali; jalankan ulang skrip hanya bila batas wilayah berubah. |
| `tools/config.json` | **Tidak di-commit.** Lokasi master Excel + kata sandi |
| `tools/lock.json` | Kunci publik + kunci privat yang terkunci kata sandi (dibuat sekali dengan `--lock`). **Tidak rahasia**, di-commit. Dengan file ini build tidak butuh kata sandi. |

## Update data lewat web (admin)

Lihat **ADMIN.md**: admin login Google di `/admin`, upload master Excel, dashboard ter-update otomatis. Cara lewat Mac di bawah tetap berlaku.

## Update data (rutin)

**Satu-satunya sumber: `02. DASHBOARD/MASTER_DATA_OUTBOUND_LOGISTIC.xlsx`.** Tab *PANDUAN* di file itu berisi daftar sheet, sumber dan kolom wajib.

1. Tutup master Excel di komputer lain, lalu tambahkan baris baru (jangan menimpa baris lama).
2. Simpan & tutup master Excel.
3. Jalankan build:
   - Windows: double-click `tools\run_build.bat`
   - Mac: `python3 tools/build_site.py`
4. Baca baris `cek:` di hasil build (SPJ duplikat otomatis diabaikan; dwell dihitung otomatis bila kosong; kode ekspeditur diisi dari nama bila kosong).
5. GitHub Desktop › **Push origin**. Situs ter-update ±1–2 menit.

### Kunci publik (sekali saja, di Mac)

`python3 tools/build_site.py --lock` membuat `tools/lock.json` dan membangun ulang data. Kata sandi tim **tidak berubah** dan sesi yang terbuka tetap jalan. Commit `tools/lock.json` + `data/`, lalu push.
Sesudahnya setiap build mengunci data dengan kunci publik, jadi build juga bisa jalan di server tanpa kata sandi: `python3 tools/build_site.py --master FILE.xlsx`.

Halaman **Kualitas data** (`#/kualitas`, link di footer): tanggal tanpa realisasi FRC/FOT, baris tanpa toko, hasil `cek:` build, distrik yang tidak cocok dengan peta.

Ganti kata sandi: `tools\ganti_password.bat` (Windows) atau `python3 tools/build_site.py --password`, lalu push dan bagikan password baru secara pribadi.

Kebutuhan: Python 3.9+, `pip install openpyxl cryptography`.

Dashboard lama diarsipkan sebagai repo `logistik-sp-lama`; skrip `AUTO_UPDATE/update_dashboard.py` miliknya tidak dipakai lagi.

## Aturan perhitungan

- Tanggal realisasi = `TGL_SPJ`.
- Situs menampilkan plan vs realisasi **s.d. H-1**: baris realisasi bertanggal hari build atau sesudahnya otomatis dilewati (lihat baris `cek:`), jadi data hari H boleh ikut tertempel di master.
- Target = **SNOP saja**. Target MTD = jumlah target harian (kolom D1..D31) dari tgl 1 s.d. hari data terakhir.
- Capaian % = realisasi **FRC** ÷ target SNOP MTD. FOT tidak memiliki target.
- Default filter: FRC, tanpa PP Belawan SBA (bisa diubah di menu Filter).
- Real/SO = realisasi FRC ÷ **SO sebulan penuh**: per bulan kirim (kolom TGL_KIRIM) bila ada, selain itu per kolom PERIODE. SO per ekspeditur = atribusi. Real/SO diwarnai netral (bukan target): di awal bulan nilainya wajar rendah.
- Forecast harian = snapshot Prognosa terakhir di master Excel; mingguan/bulanan = proyeksi laju rata-rata.

Angka September 2026 sudah dicocokkan dengan *Report Pengiriman OL MTD 30 September 2026* (per provinsi, ekspeditur dan minggu).
