# Auto Sertifikat — Webinar DWP Kabupaten Bengkalis 2026

Alur: **Google Form → Google Sheets → Apps Script → PDF sertifikat (Drive) → email ke peserta**.

## Isi folder
- `template/sertifikat-background.png` — template sertifikat dari PDF Canva (nama masih kosong).
- `apps-script/Code.gs` — script otomatisasi.

## Langkah setup (±10 menit)

### 1. Buat template di Google Slides
1. Buka [slides.new](https://slides.new), lalu **File → Page setup → Custom** → `11,69 x 8,27 inci` (A4 landscape).
2. Hapus semua isi slide kosong, lalu **Slide → Change background → Choose image** → upload `template/sertifikat-background.png`.
3. **Insert → Text box**, ketik persis `{{NAMA}}`, lalu atur:
   - Posisi (Format options → Size & position): kiri `202 pt`, atas `250 pt`, lebar `438 pt`, tinggi `40 pt`
   - Rata tengah, font sesuai selera (mis. Playfair Display / Georgia Bold), ukuran `28`, warna biru tua `#1F2A44`
4. Salin **ID Slides** dari URL: `docs.google.com/presentation/d/`**`<ID>`**`/edit`.

### 2. Pasang script
1. Buka [script.new](https://script.new), hapus isi bawaan, paste isi `apps-script/Code.gs`.
2. Ganti `TEMPLATE_ID` di bagian `CONFIG` dengan ID Slides tadi.
3. Pilih fungsi `setup` → **Run** → izinkan akses Google (Form, Sheets, Drive, Slides, Gmail).
4. Buka **Execution log**: akan muncul link **Form** (untuk dibagikan ke peserta) dan link **Spreadsheet**.

### 3. Aktifkan halaman unduh (link sertifikat setelah submit)
1. Di Apps Script: **Deploy → New deployment → ikon roda gigi → Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
2. Klik **Deploy**, izinkan akses, lalu **salin Web app URL yang berakhiran `/exec`** (bukan `/dev`, itu hanya bisa dibuka pemilik).
3. Tempel di `CONFIG.WEBAPP_URL`, lalu jalankan fungsi `updateFormConfirmation()` sekali.
4. Pesan konfirmasi Form sekarang berisi link **UNDUH SERTIFIKAT**. Peserta memasukkan email yang sama dengan di form, lalu link PDF muncul (halaman menunggu otomatis sampai sertifikat selesai dibuat).

> Kalau kode di-update lagi nanti, buka **Deploy → Manage deployments → Edit → New version → Deploy** supaya halaman unduh ikut berubah. Jangan jalankan `setup()` ulang.

Selesai. Setiap peserta submit form, sertifikat langsung dibuat, PDF disimpan di folder Drive
`Sertifikat DWP Bengkalis 2026`, dan dikirim ke email peserta. Kolom **Status Sertifikat** dan
**Link Sertifikat (PDF)** terisi otomatis di Sheet.

## Catatan penting
- **Link PDF bersifat "siapa saja yang punya link"** (tanpa login), dan halaman unduh mencari berdasarkan email. Artinya orang yang tahu email peserta lain bisa melihat nama dan sertifikatnya. Risikonya kecil untuk acara ini, tapi baik diketahui.
- **Typo nama = sertifikat salah.** Form sudah memberi peringatan; kalau ada koreksi, hapus status di baris itu lalu jalankan `generateMissingCertificates()`.
- Baris yang gagal (status `GAGAL: ...`) bisa diulang dengan menjalankan `generateMissingCertificates()`.
- Kuota email Gmail biasa ±100/hari (Google Workspace ±1.500). Kalau pesertanya lebih banyak, set `SEND_EMAIL: false` dan bagikan link PDF dari Sheet.
- Form ini membuat form baru. Kalau mau tetap memakai form lama, hubungkan ke Spreadsheet dan pasang trigger `onFormSubmit` secara manual, pastikan judul kolom sama dengan di `COL`.
