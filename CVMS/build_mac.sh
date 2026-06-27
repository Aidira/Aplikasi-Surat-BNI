#!/usr/bin/env bash
# ============================================================
#  Build aplikasi macOS (.app) dari AplikasiSuratBNI
#  Jalankan DI macOS (PyInstaller tidak bisa cross-compile).
# ============================================================
set -e

# Pilih interpreter Python terbaik yang tersedia (disarankan 3.11/3.12;
# Python 3.9 bawaan macOS terlalu lama untuk TensorFlow versi baru).
if command -v python3.12 >/dev/null 2>&1; then
    PY=python3.12
elif command -v python3.11 >/dev/null 2>&1; then
    PY=python3.11
else
    PY=python3
fi
echo "Menggunakan Python: $($PY --version)"

echo "[1/4] Membuat virtual environment..."
"$PY" -m venv build_env
source build_env/bin/activate

echo "[2/4] Memasang dependensi (termasuk TensorFlow, agak lama)..."
pip install --upgrade pip
pip install -r requirements.txt

echo "[3/4] Membersihkan hasil build lama..."
rm -rf build dist

echo "[4/4] Membangun aplikasi dengan PyInstaller..."
pyinstaller AplikasiSuratBNI.spec

echo ""
echo "SELESAI. Hasil ada di: dist/AplikasiSuratBNI.app"
echo "Untuk distribusi, kompres dist/AplikasiSuratBNI.app menjadi .zip."
echo "Catatan: di Mac lain, klik kanan > Open saat pertama kali (aplikasi belum ditandatangani Apple)."
