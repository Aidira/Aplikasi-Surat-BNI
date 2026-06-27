# Cara Membungkus Aplikasi Menjadi Aplikasi Mac/Windows

Aplikasi ini (Python + Tkinter/ttkbootstrap) dibungkus menjadi aplikasi
desktop mandiri menggunakan **PyInstaller**. Hasilnya:

- **Windows** → `AplikasiSuratBNI.exe` (di dalam folder)
- **macOS** → `AplikasiSuratBNI.app`

Pengguna akhir **tidak perlu** memasang Python — semua sudah dibundel,
termasuk TensorFlow agar **LSTM aktif**.

> ⚠️ **Penting:** PyInstaller **tidak bisa cross-compile**. Untuk
> menghasilkan `.exe` harus dibangun **di Windows**, dan untuk `.app`
> harus dibangun **di macOS**. Tidak bisa membuat `.exe` dari Mac, atau
> sebaliknya.

---

## Prasyarat
- Python 3.11 atau 3.12 terpasang di komputer build.
- Berada di folder `CVMS/` (tempat `AplikasiSuratBNI.py` dan `logo.png`).

## Cara Cepat

### Windows
Klik dua kali `build_windows.bat`, atau dari Command Prompt:
```
build_windows.bat
```
Hasil: `dist\AplikasiSuratBNI\AplikasiSuratBNI.exe`

### macOS
```
chmod +x build_mac.sh
./build_mac.sh
```
Hasil: `dist/AplikasiSuratBNI.app`

## Cara Manual (kedua OS)
```bash
python -m venv build_env
# Windows:  build_env\Scripts\activate
# macOS:    source build_env/bin/activate
pip install -r requirements.txt
pyinstaller AplikasiSuratBNI.spec
```

---

## Catatan Penting

**Ukuran besar.** Karena TensorFlow disertakan, hasil build bisa
**±800 MB – 1.5 GB**. Ini wajar untuk aplikasi yang memuat LSTM.

**Database.** Saat dijalankan sebagai aplikasi terbungkus, file
`database_bni.db` otomatis disimpan di folder data pengguna (bukan di
dalam bundle yang read-only):
- Windows: `%APPDATA%\AplikasiSuratBNI\`
- macOS: `~/Library/Application Support/AplikasiSuratBNI/`

**Distribusi.**
- Windows: kompres seluruh folder `dist\AplikasiSuratBNI` menjadi `.zip`.
- macOS: kompres `dist/AplikasiSuratBNI.app`. Di Mac lain, pertama kali
  buka dengan **klik kanan → Open** (aplikasi belum ditandatangani Apple).

**Ikon (opsional).**
- Windows: sediakan `logo.ico` di folder `CVMS/` (otomatis terpakai).
- macOS: sediakan `logo.icns` lalu set `icon='logo.icns'` pada `BUNDLE`
  di `AplikasiSuratBNI.spec`.

**Jika startup lambat / antivirus mengeluh** (umum pada bundel
PyInstaller berukuran besar): itu normal pada peluncuran pertama. Mode
folder (onedir) sudah dipilih agar lebih cepat & stabil dibanding
onefile, terutama karena TensorFlow.
