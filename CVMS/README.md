# Aplikasi Desktop BNI — Manajemen Over-Limit & Prototipe Prediksi Pagu Kas

Aplikasi desktop berbasis Python (Tkinter + ttkbootstrap) untuk mengelola
data over-limit kas multi-cabang BNI, dilengkapi prototipe penelitian skripsi
*"Analisis Algoritma LSTM untuk Prediksi Pagu Kas Harian Outlet (Studi Kasus:
BNI Cabang Tebet)"*.

> **Catatan**: Modul prediksi pagu kas pada tab "Prediksi Pagu Kas" adalah
> **prototipe/proof-of-concept** untuk keperluan skripsi, **bukan** sistem
> operasional yang dipakai BNI.

## Fitur

- **Preview Data** — upload Excel data over-limit, lihat tabel ringkasan, simpan ke SQLite.
- **Prediksi Pagu Kas** — upload dataset historis pagu kas harian (kolom `pagu_buka_tebet`),
  jalankan baseline Moving Average 3 hari (MA3) dan model LSTM univariat, lalu
  bandingkan keduanya:
  - Grafik aktual vs prediksi untuk MA3 dan LSTM.
  - Tabel metrik pembanding (MAE, RMSE, MAPE, R²) — kedua model ditampilkan apa
    adanya, tanpa rekayasa agar salah satu "menang".
  - Panel **Validasi 2 Arah**: parameter `MinMaxScaler` (min/max dari data
    training), nilai ternormalisasi satu sampel uji, dan output prediksi
    model untuk sampel itu — untuk dicocokkan manual dengan perhitungan di
    workbook Excel.
  - Jika TensorFlow tidak tersedia di environment, aplikasi otomatis fallback
    ke MA3 saja dan memberi label jelas bahwa LSTM tidak aktif.

## Struktur Proyek

```
CVMS/
├── AplikasiSuratBNI.py     # Entry point GUI
├── models/
│   └── predictor.py        # Logika MA3 + LSTM, split kronologis, metrik, validasi 2 arah
├── requirements.txt
└── logo.png
```

## Instalasi & Menjalankan

```bash
python3 -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
python AplikasiSuratBNI.py
```

## Catatan Data

Dataset historis pagu kas (`.xlsx`) **tidak disertakan** di repo ini karena
bersifat internal/sensitif — `.gitignore` di root proyek memblokir semua
file `.xlsx/.xls/.csv/.db`. Siapkan file dataset sendiri dengan kolom target
`pagu_buka_tebet` pada sheet `5_Data_Mentah`, lalu upload melalui tombol
"Upload Dataset Pagu (.xlsx)" di tab Prediksi Pagu Kas.
