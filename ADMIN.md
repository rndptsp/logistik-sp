# Halaman admin: update data dari web

Admin membuka `https://logistik-sp.pages.dev/admin`, masuk dengan Google, mengunggah file ekspor apa adanya
(Rilis + SO tiap hari; Tactical + SNOP + Monitoring di awal bulan), lalu menekan **Proses update**.
Dashboard ter-update otomatis dalam ±5–15 menit, tanpa Mac atau GitHub Desktop.

```
Admin (HP/laptop) ─ Google login (Cloudflare Access + cek di kode) ─► /admin
   pilih file ─► dipotong 16 MB di browser ─► R2 privat: inbox/<slot>.xlsx
   "Proses update" ─► GitHub Action (repository_dispatch)
        unduh store master (R2) + inbox ─► tools/master/run_web.py
          Target (jika 3 file ada) → 01_TARGET.xlsx → Rilis → SO → rapikan kemasan
        ─► tools/build_site_baru.py (kunci publik tools/lock.json, tanpa kata sandi)
        ─► commit data/site.enc + data/meta.json + index.html
        ─► store baru ke R2: versi lama disimpan (7 terakhir), inbox yang terproses dihapus
        ─► status + log dikirim balik ke /admin
```

## Isi R2 (bucket privat `logistik-sp-master`)
| Awalan | Isi |
|---|---|
| `store/current/` | master aktif: `data/rilis.csv.gz`, `so.csv.gz`, `target_*.csv.gz`, `config/*.json`, `_meta.json` |
| `store/versions/<waktu>/` | 7 salinan terakhir sebelum tiap update / rollback |
| `store/staging/` | hasil run yang belum dipromosikan (dipakai sesaat) |
| `inbox/<slot>.xlsx` | file mentah yang diunggah (rilis, so, tactical, snop, monitoring); dihapus setelah berhasil diproses |
| `state/` | status build, opsi, riwayat |

Master (rilis/SO/target) **tidak pernah** ada di repo. Di mesin build hanya ada di folder sementara runner dan
dihapus di akhir. Output tiap langkah ditulis ke berkas log yang dikirim ke halaman admin, bukan ke log GitHub (publik).

## Pemakaian
1. **Sekali saja (bagian 0):** pilih ketujuh file master dari folder `MASTER BARU` (4 × `data/*.csv.gz`, 3 × `config/*.json`).
2. **Harian (bagian 1):** pilih Rilis dan SO → tunggu "Terunggah" → **Proses update**. Upload ulang hari yang sama aman
   (kunci NO SPJ / SO+produk+line menggantikan baris lama, tidak ada duplikat).
3. **Awal bulan:** tambah Tactical + SNOP (versi sheet "Rincian") + Monitoring, pilih bulan target.
4. **Rollback (bagian 4):** tombol *Kembalikan* pada versi yang diinginkan memulihkan master lalu membangun ulang dashboard.
   Rollback juga tersimpan sebagai versi, jadi bisa dibatalkan.
- Guardrail (SPJ dobel, tanggal di luar rentang, jumlah baris turun >10%, tonase tidak valid) menghentikan build
  dan menampilkan alasannya di log admin; master tersimpan dan dashboard tidak berubah. Centang *Abaikan guardrail* hanya bila sengaja.
- Tidak ada file mentah yang diganti nama: slot menentukan jenis file. Format kolom: lihat `TEMPLATE_UPDATE_HARIAN.xlsx`
  (di folder MASTER BARU, bukan di repo) atau ringkasan di bawah.

### Kolom wajib
- **Rilis** (satu sheet per plant, nama sheet diawali kode plant): `NO SPJ`, `TGL SPJ`, `NAMA PROPINSI`, `INCOTERM`, `TONASE`; kolom lain sesuai ekspor.
- **SO** (sheet `List SO`): `Sales Order`, `Product Code`, `Line`, `Tanggal Kirim`, `Incoterm`, `Qty SO`, `UOM`, `Qty Release`, `Qty Rest SO`, `Propinsi`, `Rejection`, `Blocking`, `Keterangan` (hanya `AP1` diambil).
- **Target:** tiga file ekspor sekaligus; skrip `tools/master/update_target.py` yang membacanya.

## Setup sekali (±60 menit)

### 1. Cloudflare Pages
1. Buat akun di dash.cloudflare.com.
2. **Workers & Pages → Create → Pages → Connect to Git** → pilih `rndptsp/logistik-sp`, branch `main`.
   Framework: *None*, build command: kosong, output: `.` (dibaca dari `wrangler.toml`).
3. Nama proyek `logistik-sp` → alamat `https://logistik-sp.pages.dev`. Kalau nama terpakai, catat alamat yang diberikan dan pakai alamat itu di langkah berikut.

### 2. R2 (penyimpanan master)
**R2 → Create bucket** → nama `logistik-sp-master` (lokasi: Asia-Pacific). Cloudflare bisa meminta metode pembayaran untuk mengaktifkan R2, walau pemakaian sekecil ini masuk kuota gratis.

### 3. Google login (Cloudflare Zero Trust / Access)
1. **Zero Trust** → pilih nama tim (mis. `semenpadang-ol`) → alamat tim `https://<tim>.cloudflareaccess.com`.
2. Google Cloud Console → **APIs & Services → Credentials → Create OAuth client ID** (Web application).
   Authorized redirect URI: `https://<tim>.cloudflareaccess.com/cdn-cgi/access/callback`. Salin Client ID + Secret.
3. Zero Trust → **Settings → Authentication → Login methods → Add → Google** → tempel Client ID + Secret.
   (Opsional: tambah juga *One-time PIN* sebagai cadangan.)
4. Zero Trust → **Access → Applications → Add → Self-hosted**:
   - Domain: `logistik-sp.pages.dev`, path `admin`; tambah domain kedua yang sama dengan path `api/admin`.
   - Tambah juga `*.logistik-sp.pages.dev` dengan dua path yang sama (alamat preview).
   - Policy: **Allow**, *Emails*: email Google para admin.
   - Simpan, lalu salin **Application Audience (AUD) Tag**.

### 4. Variabel di Cloudflare Pages
Proyek `logistik-sp` → **Settings → Variables and Secrets** (Production):

| Nama | Isi | Jenis |
|---|---|---|
| `ACCESS_TEAM_DOMAIN` | `https://<tim>.cloudflareaccess.com` | Text |
| `ACCESS_AUD` | AUD tag dari langkah 3.4 | Text |
| `ADMIN_EMAILS` | email admin, dipisah koma (sama dengan policy) | Text |
| `GITHUB_TOKEN` | token GitHub dari langkah 5.1 | **Secret** |
| `BUILD_TOKEN` | string acak ≥32 karakter (buat sekali, simpan) | **Secret** |

Lalu **Deployments → Retry deployment** supaya variabel terpakai.

### 5. GitHub
1. github.com → **Settings → Developer settings → Fine-grained tokens → Generate**:
   repository `rndptsp/logistik-sp` saja, permission **Contents: Read and write**. Masa berlaku 1 tahun (catat tanggalnya).
2. Repo `logistik-sp` → **Settings → Secrets and variables → Actions → New repository secret**:
   - `ADMIN_URL` = `https://logistik-sp.pages.dev`
   - `BUILD_TOKEN` = string yang sama dengan langkah 4

### 6. Uji
Buka `https://logistik-sp.pages.dev/admin` → login Google → upload master → status berubah
*Menunggu build → Sedang dibangun → Berhasil* dengan "Data s.d." = kemarin.

## Menambah / mencabut admin
Ubah email di policy Access (langkah 3.4) **dan** di `ADMIN_EMAILS` (langkah 4). Keduanya harus cocok:
kode memeriksa daftar sendiri, jadi salah konfigurasi Access saja tidak membuka akses.

## Kalau build gagal
Status **Gagal** menampilkan log (mis. sheet/kolom hilang, guardrail). Perbaiki file lalu unggah ulang slot itu dan tekan Proses update lagi.
Dashboard dan master tersimpan tetap versi terakhir yang berhasil. Jika dashboard sudah terbit tetapi simpan master gagal
(pesan di akhir log), unggah file hari itu lagi: upsert membuatnya konsisten.

## Cadangan: build lewat Mac
`tools/run_build.bat` (master di folder `MASTER BARU`, `run_harian.py`) tetap bisa dipakai. Hati-hati: master di Mac dan di R2 adalah dua salinan;
pilih satu jalur per periode, atau muat ulang master awal ke R2 (hapus `store/` di R2 lebih dulu).
