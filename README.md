# Profitability Command · Segment 3

Dashboard profitabilitas Segmen 3 dengan login Firebase dan data **realtime dari Google Sheets**.

```
Google Sheets ──(Apps Script, akun "Sync")──▶ Firestore ──(listener realtime)──▶ Dashboard di GitHub Pages
  edit angka          ±1–2 menit                datasets/segment3-ytd2026          login email + password
```

- Halaman: `https://nandar0503.github.io/dashboard_performance/dashboard_profitability_govseg3_2026.html` (URL sama seperti sebelumnya).
- File HTML **tidak berisi data**. Angka hanya diambil dari Firestore setelah login dengan akun berperan aktif.
- Setiap perubahan di Google Sheets otomatis dikirim ke Firestore, dan semua dashboard yang sedang terbuka ikut ter-update tanpa reload. Status terlihat di chip **● Live** di bagian atas.

## Isi repo

| File | Fungsi |
|---|---|
| `dashboard_profitability_govseg3_2026.html` | Portal login + dashboard (menggantikan file lama) |
| `firebase-config.js` | Konfigurasi Web App Firebase. **Wajib diisi** (apiKey, messagingSenderId, appId) |
| `index.html` | Mengalihkan ke portal |
| `firestore.rules` | Aturan keamanan Firestore (peran, akun Sync, validasi dataset, audit) |
| `apps-script/Code.gs`, `apps-script/appsscript.json` | Skrip sinkronisasi untuk Google Sheets |
| `firebase.json`, `.firebaserc` | Untuk deploy rules lewat Firebase CLI (opsional) |

---

## A. Firebase (sekali saja)

1. **Authentication → Sign-in method → Email/Password → Enable.**
2. **Authentication → Settings → Authorized domains → Add domain:** `nandar0503.github.io`.
3. **Firestore Database → Create database** (production mode, lokasi `asia-southeast2`).
4. **Firestore → Rules:** salin seluruh isi `firestore.rules`, lalu **Publish**.
   (Atau lewat CLI: `firebase deploy --only firestore:rules`.)
5. **Project settings → General → Your apps → Add app → Web (</>)**. Salin `apiKey`, `messagingSenderId`, `appId` ke `firebase-config.js`.

### Akun Director pertama
1. **Authentication → Users → Add user** (email + password Anda). Salin **User UID**.
2. **Firestore → Start collection `users`** → Document ID = UID tersebut:

| Field | Tipe | Nilai |
|---|---|---|
| `name` | string | Nama Anda |
| `email` | string | Email Anda |
| `role` | string | `Director` |
| `active` | boolean | `true` |

### Akun Sync (untuk Google Sheets)
Login ke portal sebagai Director → **14 · Kelola Pengguna → Tambah pengguna**:
nama `Google Sheets Sync`, email mis. `sync.profitability@domain-anda`, password kuat, peran **Sync**.
Akun ini hanya bisa menulis dataset; tidak bisa membuka dashboard.

## B. GitHub Pages

Ganti isi repo `dashboard_performance` dengan file di folder ini, lalu commit:

```bash
git clone https://github.com/nandar0503/dashboard_performance
cd dashboard_performance
# salin semua file dari folder ini ke sini (timpa file lama)
git add -A
git commit -m "Portal login Firebase + data realtime dari Google Sheets"
git push
```

> **Penting:** file lama berisi seluruh angka keuangan dan masih tersimpan di riwayat commit publik.
> Untuk menghapusnya, buat ulang repo atau ganti riwayat dengan branch baru:
> ```bash
> git checkout --orphan bersih && git add -A && git commit -m "Portal tanpa data"
> git branch -D main && git branch -m main && git push -f origin main
> ```

## C. Google Sheets

1. Buka spreadsheet → **Extensions → Apps Script**.
2. Ganti isi `Code.gs` dengan `apps-script/Code.gs`.
   Untuk `appsscript.json`: **Project Settings → centang "Show appsscript.json"**, lalu ganti isinya.
3. Simpan, kembali ke spreadsheet, **muat ulang**. Menu **Profitability Sync** muncul.
4. **Profitability Sync → 1. Atur koneksi Firebase**: isi Web API Key, email & password akun Sync. Setujui izin akses saat diminta.
5. **Profitability Sync → Sinkronkan sekarang**: data pertama terkirim, dashboard terisi.
6. **Profitability Sync → 2. Aktifkan sinkron otomatis**: selanjutnya setiap perubahan terkirim dalam ±1–2 menit.

> **Penting:** ubah akses berbagi spreadsheet menjadi **Restricted** (hanya orang tertentu).
> Jika link spreadsheet bisa dibuka siapa saja, login di dashboard tidak melindungi datanya.

### Struktur sheet yang dibaca
Skrip membaca format yang sudah ada, jadi tidak perlu template baru:

| Sheet | Yang dibaca |
|---|---|
| `Summary` | Tabel "Layanan": revenue & direct cost SAP/OPS per pelanggan (angka resmi) |
| `Breakdown Pareto` | 5 blok "Beban …" per pelanggan + tabel total per kategori |
| `Rekap SAP S3` | Baris pertama = bulan (format tanggal) + "Grand Total"; dipakai untuk tren bulanan, depresiasi, catatan |

- **Menambah bulan:** sisipkan kolom tanggal baru (mis. Okt-26) sebelum "Grand Total" di `Rekap SAP S3`. Dashboard otomatis menjadi "YTD Jan–Okt".
- **Label tertukar** di `Summary` terdeteksi otomatis dengan mencocokkan angka ke ledger & Pareto, lalu dikoreksi di dashboard. Saat ini terdeteksi 4 pasang: Kementerian HAM ↔ Dirjen Pajak, Kejagung ↔ Ditjen Minerba, Universitas Terbuka ↔ BP Tapera, Al Azhar ↔ Universitas Indonesia.
- Catatan di kolom setelah "Grand Total" muncul di modul 11 dan Executive Insights.
- Riwayat sinkron tercatat di sheet **Log Sinkronisasi** (dibuat otomatis) dan di **Audit Trail** dashboard.

## Peran & akses

| Peran | Dashboard | Ubah dataset | Audit Trail | Kelola pengguna |
|---|---|---|---|---|
| Director | Semua modul | ✓ | ✓ | ✓ |
| Finance | Semua modul | ✓ | ✓ | – |
| Business Unit Head / Analyst | 01–10, 12 | – | – | – |
| Project Manager | 01, 04, 05, 06, 08, 09, 12 | – | – | – |
| Sync | – (tidak bisa login ke dashboard) | ✓ hanya `segment3-ytd2026` | – | – |

## Pemecahan masalah

| Gejala | Penyebab & solusi |
|---|---|
| "Konfigurasi Firebase tidak ditemukan" | `firebase-config.js` belum diisi atau belum di-push |
| "Akun belum diaktifkan" | Dokumen `users/{UID}` belum ada / `active` bukan `true` |
| Log: "Ditolak firestore.rules" | Akun Sync belum punya profil `role: "Sync"` aktif, atau rules belum di-publish |
| Log: "Login Firebase gagal" | Web API Key / email / password akun Sync salah (menu 1) |
| Chip "○ Koneksi realtime terputus" | Jaringan terputus; Firestore menyambung ulang otomatis |
| Data tidak berubah setelah edit | Cek menu **Cek status**; sinkron otomatis menunggu 20 detik setelah edit terakhir |
