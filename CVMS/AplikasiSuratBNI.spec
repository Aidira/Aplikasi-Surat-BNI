# -*- mode: python ; coding: utf-8 -*-
"""
Konfigurasi PyInstaller untuk membungkus AplikasiSuratBNI menjadi
aplikasi desktop Windows (.exe) atau macOS (.app).

Menyertakan: logo.png, paket model, TensorFlow (agar LSTM aktif),
scikit-learn, dan ttkbootstrap.

Mode onedir (COLLECT) dipakai—bukan onefile—karena TensorFlow jauh
lebih andal dan startup lebih cepat dalam mode folder.

Cara pakai:
    pyinstaller AplikasiSuratBNI.spec
Hasil ada di folder dist/.
"""

import os
import sys
from PyInstaller.utils.hooks import collect_all

# Ikon opsional: dipakai hanya jika file-nya tersedia, agar build tidak gagal.
_ICON = 'logo.ico' if (sys.platform.startswith('win') and os.path.exists('logo.ico')) else None

datas = [('logo.png', '.')]
binaries = []
hiddenimports = ['models', 'models.predictor', 'surat_generator']

# TensorFlow, scikit-learn, dan ttkbootstrap butuh pengumpulan menyeluruh
# (data file, binary, dan submodule) agar lengkap di dalam bundle.
for paket in ('tensorflow', 'sklearn', 'ttkbootstrap'):
    d, b, h = collect_all(paket)
    datas += d
    binaries += b
    hiddenimports += h

a = Analysis(
    ['AplikasiSuratBNI.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AplikasiSuratBNI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # aplikasi GUI: tanpa jendela terminal
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=_ICON,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AplikasiSuratBNI',
)

# macOS: bungkus folder hasil menjadi paket .app
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='AplikasiSuratBNI.app',
        icon=None,           # ganti ke 'logo.icns' bila punya ikon .icns
        bundle_identifier='id.bni.aplikasisuratbni',
        info_plist={
            'NSHighResolutionCapable': 'True',
            'CFBundleDisplayName': 'Aplikasi Surat BNI',
        },
    )
