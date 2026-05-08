# SIMARSIP — Panduan Setup Lengkap

## Struktur File
```
surat-system/
├── index.html      → Halaman Login
├── admin.html      → Dashboard Admin
├── user.html       → Dashboard User
└── Code.gs         → Backend Google Apps Script
```

---

## LANGKAH 1: Setup Google Sheets

1. Buat Google Spreadsheet baru
2. Catat **Spreadsheet ID** dari URL:
   `https://docs.google.com/spreadsheets/d/**[SPREADSHEET_ID]**/edit`

---

## LANGKAH 2: Deploy Google Apps Script

1. Di Spreadsheet → **Extensions → Apps Script**
2. Hapus semua kode default, paste isi `Code.gs`
3. Sesuaikan baris di bagian KONFIGURASI:
   ```javascript
   const KODE_UNIT = 'PAN.4';    // ← Ganti sesuai instansi
   const NOMOR_PREFIX = 'W28';   // ← Ganti sesuai kode surat
   ```
4. **Jalankan fungsi setupSpreadsheet()** satu kali:
   - Pilih fungsi `setupSpreadsheet` di dropdown
   - Klik tombol Run (▶)
   - Izinkan akses yang diminta
5. Deploy sebagai Web App:
   - Klik **Deploy → New Deployment**
   - Type: **Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Klik **Deploy**
6. **Copy URL Web App** yang diberikan

---

## LANGKAH 3: Setup Frontend

1. Buka `admin.html` di browser
2. Login dengan akun default:
   - Email: `admin@instansi.go.id`
   - Password: `admin123`
3. Pergi ke **Pengaturan Sistem**
4. Paste URL Web App ke kolom "URL Google Apps Script"
5. Klik **Simpan Pengaturan**

---

## LANGKAH 4: Deploy ke Cloudflare Pages / GitHub Pages

### GitHub Pages:
1. Buat repository baru di GitHub
2. Upload semua file `.html`
3. Settings → Pages → Source: main branch
4. Akses di `https://[username].github.io/[repo]`

### Cloudflare Pages (Rekomendasi):
1. Hubungkan GitHub repo ke Cloudflare Pages
2. Build command: (kosongkan)
3. Output directory: `/`
4. Deploy!

---

## Login Default

| Role  | Email                      | Password  |
|-------|---------------------------|-----------|
| Admin | admin@instansi.go.id      | admin123  |
| User  | user@instansi.go.id       | user123   |

> ⚠️ **Segera ganti password setelah pertama login!**

---

## Format Nomor Surat

```
W28.001/PAN.4/2026
│    │   │     └─ Tahun
│    │   └──────── Kode Unit
│    └──────────── Nomor Urut (auto, 3 digit)
└───────────────── Prefix instansi
```

---

## Penambahan Tahun Baru

Tidak perlu mengubah kode apapun! Cukup:
1. Login sebagai Admin
2. Pengaturan Sistem → Ubah **Tahun Aktif** ke tahun baru
3. Sistem otomatis membuat tab baru di Spreadsheet

---

## Troubleshooting

| Masalah | Solusi |
|---------|--------|
| "Gagal terhubung" | Periksa URL GAS di Pengaturan Sistem |
| Nomor ganda | Pastikan hanya ada 1 deployment aktif |
| Tidak bisa login | Jalankan `setupSpreadsheet()` ulang |
| Tab tidak muncul | Deploy ulang GAS dengan "New Deployment" |


localStorage.setItem('GAS_URL', 'https://script.google.com/macros/s/AKfycbz7W34A80P0LNl0D2qZuMYd9tXYbiuT5u3uq9baDL0-4qK6QOL5gko-Kyyhuij_svZ3/exec')
