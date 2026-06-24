"""
Modul prediksi pagu kas harian.

Berisi implementasi baseline Moving Average 3 hari (MA3) dan model LSTM
univariat sederhana, beserta utilitas split kronologis, pembentukan
sequence, perhitungan metrik evaluasi, dan validasi 2 arah (untuk
dicocokkan manual dengan perhitungan di workbook Excel).

Catatan penting (sesuai temuan skripsi):
- MinMaxScaler HARUS hanya di-fit pada data training untuk mencegah data leakage.
- Baseline MA3 terbukti mengungguli LSTM pada studi kasus ini; modul ini
  TIDAK direkayasa untuk membuat LSTM menang.
"""

import random

import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.preprocessing import MinMaxScaler

try:
    import tensorflow as tf
    from tensorflow.keras.callbacks import EarlyStopping
    from tensorflow.keras.layers import LSTM, Dense
    from tensorflow.keras.models import Sequential

    TF_AVAILABLE = True
except Exception:  # TensorFlow tidak terpasang / gagal load di environment ini
    TF_AVAILABLE = False

SEED = 42
TARGET_COL = "pagu_buka_tebet"
SHEET_NAME = "5_Data_Mentah"
# Kolom yang dibuang karena berpotensi menyebabkan data leakage (mengandung
# informasi yang sebenarnya tidak tersedia saat prediksi dilakukan).
LEAKAGE_COLS = [
    "pagu_tutup_tebet",
    "pagu_cabang_28_outlet",
    "average_outlet",
    "average_cabang",
    "total",
    "average_total",
    "status",
    "kode_hari",
]
WINDOW_DEFAULT = 10


def set_seed(seed=SEED):
    """Menetapkan seed numpy/random/tensorflow agar hasil dapat direproduksi."""
    random.seed(seed)
    np.random.seed(seed)
    if TF_AVAILABLE:
        tf.random.set_seed(seed)


def muat_dataset(path, sheet_name=SHEET_NAME):
    """
    Memuat dataset pagu kas harian dari Excel, membuang kolom leakage, dan
    memfilter baris tidak valid: 14 hari burn-in rolling statistics (sebelum
    2024-01-15) serta baris dengan target = 0.
    """
    df = pd.read_excel(path, sheet_name=sheet_name, parse_dates=["tanggal"])
    df = df.drop(columns=[c for c in LEAKAGE_COLS if c in df.columns], errors="ignore")
    if TARGET_COL not in df.columns:
        raise ValueError(f"Kolom target '{TARGET_COL}' tidak ditemukan pada sheet '{sheet_name}'.")
    df = df.sort_values("tanggal").reset_index(drop=True)
    df = df[df[TARGET_COL] > 0].reset_index(drop=True)
    if "tanggal" in df.columns:
        df = df[df["tanggal"] >= "2024-01-15"].reset_index(drop=True)
    return df


def split_kronologis(series, train_ratio=0.7, val_ratio=0.15):
    """Membagi series secara kronologis (tanpa shuffle) menjadi train/val/test."""
    n = len(series)
    n_train = int(n * train_ratio)
    n_val = int(n * val_ratio)
    train = series[:n_train]
    val = series[n_train:n_train + n_val]
    test = series[n_train + n_val:]
    return train, val, test


def buat_sequence(arr, window):
    """Membentuk pasangan sequence/window (X) dan nilai target (y) dari array 1D."""
    X, y = [], []
    for i in range(len(arr) - window):
        X.append(arr[i:i + window])
        y.append(arr[i + window])
    return np.array(X), np.array(y)


def hitung_metrik(y_true, y_pred):
    """Menghitung MAE, RMSE, MAPE, dan R-squared pada skala asli (Rupiah)."""
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    mae = mean_absolute_error(y_true, y_pred)
    rmse = np.sqrt(mean_squared_error(y_true, y_pred))
    mape = float(np.mean(np.abs((y_true - y_pred) / y_true)) * 100)
    r2 = r2_score(y_true, y_pred)
    return {"MAE": mae, "RMSE": rmse, "MAPE (%)": mape, "R2": r2}


def prediksi_ma3(train_tail, test_vals):
    """
    Baseline Moving Average 3 hari pada skala asli.
    `train_tail` adalah 2 nilai terakhir sebelum test, dipakai sebagai
    konteks kronologis sehingga prediksi test[0] tetap valid.
    """
    full = np.concatenate([train_tail[-2:], test_vals])
    s = pd.Series(full)
    pred = s.rolling(window=3).mean().shift(1)
    return pred.values[len(train_tail[-2:]):]


def latih_lstm(train_vals, val_vals, window=WINDOW_DEFAULT):
    """
    Melatih model LSTM univariat sederhana (1 layer LSTM + Dense(1)).
    Scaler MinMaxScaler di-fit HANYA pada data train untuk mencegah leakage.
    """
    if not TF_AVAILABLE:
        raise RuntimeError("TensorFlow tidak tersedia di environment ini.")
    set_seed()
    scaler = MinMaxScaler()
    train_scaled = scaler.fit_transform(np.asarray(train_vals).reshape(-1, 1)).flatten()
    val_scaled = scaler.transform(np.asarray(val_vals).reshape(-1, 1)).flatten()

    X_train, y_train = buat_sequence(train_scaled, window)
    val_with_context = np.concatenate([train_scaled[-window:], val_scaled])
    X_val, y_val = buat_sequence(val_with_context, window)

    X_train = X_train.reshape((-1, window, 1))
    X_val = X_val.reshape((-1, window, 1))

    model = Sequential([
        LSTM(50, activation="relu", input_shape=(window, 1)),
        Dense(1),
    ])
    model.compile(optimizer="adam", loss="mse")
    es = EarlyStopping(monitor="val_loss", patience=20, restore_best_weights=True)
    history = model.fit(
        X_train, y_train,
        validation_data=(X_val, y_val),
        epochs=200, batch_size=8,
        callbacks=[es], verbose=0,
    )
    return model, scaler, history


def prediksi_lstm(model, scaler, context_vals, test_vals, window=WINDOW_DEFAULT):
    """Memprediksi test set memakai konteks kronologis dari split sebelumnya."""
    full = np.concatenate([context_vals[-window:], test_vals])
    full_scaled = scaler.transform(full.reshape(-1, 1)).flatten()
    X_test, _ = buat_sequence(full_scaled, window)
    X_test = X_test.reshape((-1, window, 1))
    pred_scaled = model.predict(X_test, verbose=0).flatten()
    pred = scaler.inverse_transform(pred_scaled.reshape(-1, 1)).flatten()
    return pred


def validasi_dua_arah(scaler, sample_raw_value, model=None, context_vals=None, window=WINDOW_DEFAULT):
    """
    Menyiapkan data untuk panel "Validasi 2 Arah": parameter scaler (min/max
    train), nilai ternormalisasi satu sampel, dan output prediksi model untuk
    sampel tersebut -- agar bisa dicocokkan manual dengan workbook Excel.
    """
    data_min = float(scaler.data_min_[0])
    data_max = float(scaler.data_max_[0])
    normalized = float(scaler.transform([[sample_raw_value]])[0, 0])

    yhat = None
    if model is not None and context_vals is not None and len(context_vals) >= window:
        ctx_scaled = scaler.transform(np.asarray(context_vals[-window:]).reshape(-1, 1)).flatten()
        x = ctx_scaled.reshape((1, window, 1))
        yhat_scaled = model.predict(x, verbose=0).flatten()[0]
        yhat = float(scaler.inverse_transform([[yhat_scaled]])[0, 0])

    return {
        "scaler_min_train": data_min,
        "scaler_max_train": data_max,
        "sample_raw": float(sample_raw_value),
        "sample_normalized": normalized,
        "model_yhat": yhat,
    }
