# 🚦 Dashboard Analisis Data Lalu Lintas

Dashboard Streamlit untuk menganalisis data hasil rekap CCTV (jumlah orang/motor/mobil/truk/bus
per lokasi per hari). Cukup unggah file Excel/CSV, data langsung dibersihkan dan dianalisis, lalu
ditampilkan sebagai dashboard interaktif — lengkap dengan opsi unduh **laporan Word**.

## Fitur

- **Upload otomatis**: mendukung file `.xlsx`/`.csv`, termasuk format rekap CCTV "blok per lokasi"
  (tiap lokasi punya baris header + baris harian + baris Total sendiri) maupun format tabel biasa.
- **Pembersihan & validasi data otomatis**, termasuk deteksi lokasi yang tidak menghasilkan data
  sama sekali (indikasi kamera/alat penghitung tidak aktif) atau data yang sebagian kosong.
- **Filter interaktif**: pilih lokasi, rentang tanggal, dan agregasi periode
  (**Harian / Mingguan / Bulanan / Keseluruhan**).
- **Dashboard lengkap**: KPI utama, tren volume, perbandingan antar lokasi, komposisi jenis
  kendaraan/orang, status kualitas data, dan tabel data detail (bisa diunduh sebagai CSV).
- **Unduh hasil analisis** dalam 2 format tambahan:
  - 📄 **Laporan Word (.docx)** — ringkasan eksekutif otomatis (narasi insight), KPI, grafik, tabel
    ringkasan per lokasi, dan catatan kualitas data.

## Cara Menjalankan

1. Pastikan Python 3.9+ terpasang.
2. Install dependensi:
   ```bash
   pip install -r requirements.txt
   ```
3. Jalankan aplikasi:
   ```bash
   streamlit run app.py
   ```
4. Buka browser ke alamat yang ditampilkan (biasanya `http://localhost:8501`).
5. Unggah file data CCTV Anda di sidebar, atau centang **"Gunakan data contoh"** untuk mencoba
   dashboard dengan data contoh (`sample_data.xlsx`, sudah disertakan).

## Format Data yang Didukung

**Format 1 — Blok per lokasi (format rekap CCTV umum):**
Setiap lokasi memiliki baris header sendiri (kolom kedua berisi kata "Waktu"), diikuti baris data
harian, baris "Total", lalu baris kosong sebagai pemisah sebelum lokasi berikutnya. Urutan kolom
kategori (Orang/Motor/Mobil/Truk/Bus) boleh berbeda antar lokasi — sistem membaca susunan kolom
dari baris header masing-masing blok secara otomatis.

**Format 2 — Tabel tidy standar (fallback):**
Satu baris header di baris pertama, dengan kolom yang mencakup nama lokasi (`Lokasi`), tanggal
(`Tanggal`/`Waktu`), dan satu atau lebih kolom kategori (`Motor`, `Mobil`, `Truk`, `Bus`, `Orang`),
serta opsional kolom `Total`.

Jika file Anda tidak sesuai kedua format ini, dashboard akan menampilkan pesan kesalahan yang
menjelaskan bagian mana yang tidak terbaca.

## Struktur Proyek

```
cctv_dashboard/
├── app.py            # Aplikasi Streamlit utama (UI & orkestrasi)
├── parser.py          # Pembacaan & pembersihan data, deteksi kualitas data, agregasi periode
├── report.py          # Pembuatan laporan Word (.docx) dengan grafik & narasi otomatis
├── requirements.txt   # Daftar dependensi Python
├── sample_data.xlsx   # Data contoh untuk mencoba dashboard
└── README.md
```

## Catatan Validitas Analisis

- Kolom `Total` mengikuti nilai yang tercantum pada sumber data; jika kosong, dihitung otomatis
  sebagai penjumlahan seluruh kategori pada baris tersebut.
- Lokasi yang seluruh nilainya nol sepanjang periode ditandai **"Tidak ada data / kamera tidak
  aktif"** — bukan berarti benar-benar nol lalu lintas, melainkan indikasi data tidak terekam.
  Selalu periksa tab **Kualitas Data** sebelum mengambil kesimpulan dari angka agregat.
- Semua angka pada laporan & infografis dihitung dari data yang sudah difilter sesuai pilihan
  Anda di sidebar (lokasi, rentang tanggal, dan agregasi periode).
