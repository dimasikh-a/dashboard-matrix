# Dashboard Matriks Pelaksanaan Kegiatan

Aplikasi Next.js untuk memasukkan, mengubah, menghapus, mencari, memfilter, mengimpor, dan mengekspor data matriks kegiatan.

## Menjalankan aplikasi

```bash
npm install
npm run dev
```

Buka `http://localhost:3000` di browser.

## Menggunakan Excel sumber

1. Tekan **Impor Excel**.
2. Pilih workbook matriks kegiatan.
3. Aplikasi mengambil tabel dari sheet pertama dan menemukan baris header pertama yang berisi sedikitnya tiga kolom.
4. Semua kolom dari Excel akan dipertahankan pada tabel dan formulir edit.

Untuk pengembangan lokal, salin `.env.example` menjadi `.env`, isi `DATABASE_URL` dari Supabase, lalu jalankan `npm run db:setup`. Pada Vercel, simpan nilai yang sama sebagai environment variable `DATABASE_URL`. Data tersimpan pada PostgreSQL Supabase dan dapat dilihat melalui **Table Editor** di dashboard Supabase.
