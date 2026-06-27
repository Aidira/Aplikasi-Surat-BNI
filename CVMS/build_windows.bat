@echo off
REM ============================================================
REM  Build aplikasi Windows (.exe) dari AplikasiSuratBNI
REM  Jalankan DI WINDOWS (PyInstaller tidak bisa cross-compile).
REM ============================================================

echo [1/4] Membuat virtual environment...
python -m venv build_env
call build_env\Scripts\activate.bat

echo [2/4] Memasang dependensi (termasuk TensorFlow, agak lama)...
pip install --upgrade pip
pip install -r requirements.txt

echo [3/4] Membersihkan hasil build lama...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [4/4] Membangun aplikasi dengan PyInstaller...
pyinstaller AplikasiSuratBNI.spec

echo.
echo SELESAI. Hasil ada di: dist\AplikasiSuratBNI\AplikasiSuratBNI.exe
echo Untuk distribusi, kompres seluruh folder dist\AplikasiSuratBNI menjadi .zip
pause
