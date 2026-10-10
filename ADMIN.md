# Halaman admin: upload master Excel dari web

Admin membuka `https://outbound-sp.pages.dev/admin`, masuk dengan Google, lalu meng-upload
`MASTER_DATA_OUTBOUND_LOGISTIC.xlsx`. Dashboard ter-update otomatis dalam ±3–5 menit.

```
Admin (HP/laptop) ─ Google login (Cloudflare Access + cek di kode) ─► /admin
   upload .xlsx ─► R2 (privat) ─► memicu GitHub Action
                                   └─ unduh master → build_site.py (tanpa kata sandi, pakai tools/lock.json)
                                      → push data/ → Cloudflare Pages + GitHub Pages ter-update
                                      → status + log build dikirim balik ke /admin
```

- Kata sandi penonton **tidak disimpan di server mana pun**. Ganti kata sandi dikerjakan di browser admin (kunci privat dikunci ulang), server hanya menerima kunci yang sudah terkunci.
- Master Excel hanya ada di R2 (bucket privat, versi terakhir saja) dan sesaat di mesin build. Tidak pernah masuk repo atau log GitHub (repo ini publik); log build hanya tampil di halaman admin.
- Build lewat Mac tetap bisa seperti biasa.

## Setup sekali (±60 menit)

### 1. Cloudflare Pages
1. Buat akun di dash.cloudflare.com.
2. **Workers & Pages → Create → Pages → Connect to Git** → pilih `rndptsp/outbound-sp`, branch `main`.
   Framework: *None*, build command: kosong, output: `.` (dibaca dari `wrangler.toml`).
3. Nama proyek `outbound-sp` → alamat `https://outbound-sp.pages.dev`. Kalau nama terpakai, catat alamat yang diberikan dan pakai alamat itu di langkah berikut.

### 2. R2 (penyimpanan master)
**R2 → Create bucket** → nama `outbound-sp-master` (lokasi: Asia-Pacific). Cloudflare bisa meminta metode pembayaran untuk mengaktifkan R2, walau pemakaian sekecil ini masuk kuota gratis.

### 3. Google login (Cloudflare Zero Trust / Access)
1. **Zero Trust** → pilih nama tim (mis. `semenpadang-ol`) → alamat tim `https://<tim>.cloudflareaccess.com`.
2. Google Cloud Console → **APIs & Services → Credentials → Create OAuth client ID** (Web application).
   Authorized redirect URI: `https://<tim>.cloudflareaccess.com/cdn-cgi/access/callback`. Salin Client ID + Secret.
3. Zero Trust → **Settings → Authentication → Login methods → Add → Google** → tempel Client ID + Secret.
   (Opsional: tambah juga *One-time PIN* sebagai cadangan.)
4. Zero Trust → **Access → Applications → Add → Self-hosted**:
   - Domain: `outbound-sp.pages.dev`, path `admin`; tambah domain kedua yang sama dengan path `api/admin`.
   - Tambah juga `*.outbound-sp.pages.dev` dengan dua path yang sama (alamat preview).
   - Policy: **Allow**, *Emails*: email Google para admin.
   - Simpan, lalu salin **Application Audience (AUD) Tag**.

### 4. Variabel di Cloudflare Pages
Proyek `outbound-sp` → **Settings → Variables and Secrets** (Production):

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
   repository `rndptsp/outbound-sp` saja, permission **Contents: Read and write**. Masa berlaku 1 tahun (catat tanggalnya).
2. Repo `outbound-sp` → **Settings → Secrets and variables → Actions → New repository secret**:
   - `ADMIN_URL` = `https://outbound-sp.pages.dev`
   - `BUILD_TOKEN` = string yang sama dengan langkah 4

### 6. Uji
Buka `https://outbound-sp.pages.dev/admin` → login Google → upload master → status berubah
*Menunggu build → Sedang dibangun → Berhasil* dengan "Data s.d." = kemarin.

## Menambah / mencabut admin
Ubah email di policy Access (langkah 3.4) **dan** di `ADMIN_EMAILS` (langkah 4). Keduanya harus cocok:
kode memeriksa daftar sendiri, jadi salah konfigurasi Access saja tidak membuka akses.

## Kalau build gagal
Status **Gagal** di halaman admin menampilkan log build (mis. sheet atau kolom yang hilang). Perbaiki master, upload ulang.
Data di dashboard tetap versi terakhir yang berhasil.
