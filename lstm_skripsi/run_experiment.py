"""
Eksperimen LSTM & Baseline MA3 – Prediksi Pagu Kas Harian Outlet BNI Tebet
Versi: script mandiri (isi sama dengan notebook)
"""

# ─────────────────────────────────────────────────────────────────────────────
# PARAMETER (ubah di sini jika path berbeda)
# ─────────────────────────────────────────────────────────────────────────────
DATASET_PATH = "d8fc77b5-DATASETPAGU_HASIL_PREPROCESSING.xlsx"
SHEET_NAME   = "5_Data_Mentah"
OUTPUT_DIR   = "output"
TARGET_COL   = "pagu_buka_tebet"
WINDOW_SIZES = [5, 7, 10]
SEED         = 42
LSTM_UNITS   = 50
EPOCHS       = 200
BATCH_SIZE   = 32
PATIENCE     = 20

# ─────────────────────────────────────────────────────────────────────────────
# 0. Seed global
# ─────────────────────────────────────────────────────────────────────────────
import os, random
os.environ["PYTHONHASHSEED"] = str(SEED)
random.seed(SEED)

import numpy as np
np.random.seed(SEED)

import tensorflow as tf
tf.random.set_seed(SEED)
# Matikan GPU nondeterminism jika ada GPU
os.environ["TF_DETERMINISTIC_OPS"] = "1"

# ─────────────────────────────────────────────────────────────────────────────
# 1. Import
# ─────────────────────────────────────────────────────────────────────────────
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from sklearn.preprocessing import MinMaxScaler
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import LSTM, Dense
from tensorflow.keras.callbacks import EarlyStopping

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# 2. Load data
# ─────────────────────────────────────────────────────────────────────────────
print("="*60)
print("2. LOAD DATA")
print("="*60)
df_raw = pd.read_excel(DATASET_PATH, sheet_name=SHEET_NAME, parse_dates=["tanggal"])
print(f"Shape mentah  : {df_raw.shape}")
print("Kolom         :", list(df_raw.columns))
print("Missing values:\n", df_raw.isnull().sum())

# ─────────────────────────────────────────────────────────────────────────────
# 3. Seleksi kolom & buang leakage
# ─────────────────────────────────────────────────────────────────────────────
#   Leakage:  pagu_tutup_tebet  → informasi akhir hari, belum tersedia saat prediksi
#             pagu_cabang_28_outlet, average_outlet, average_cabang, total, average_total
#             → agregat multi-outlet / turunan dari variabel lain (data leakage)
#   Dipertahankan: tanggal (index), kode_hari (opsional), pagu_buka_tebet (target)
#   Karena LSTM univariat, kita hanya pakai pagu_buka_tebet sebagai input sequence.
LEAKAGE_COLS = [
    "pagu_tutup_tebet",
    "pagu_cabang_28_outlet",
    "average_outlet",
    "average_cabang",
    "total",
    "average_total",
    "status",        # label post-hoc, bukan fitur prediktif murni
    "kode_hari",     # dibuang karena LSTM univariat; tidak dimasukkan sebagai fitur
]
df = df_raw.drop(columns=LEAKAGE_COLS, errors="ignore")
print("\nKolom setelah drop leakage:", list(df.columns))

# ─────────────────────────────────────────────────────────────────────────────
# 4. Filter baris – selaraskan dengan preprocessing pipeline resmi
#    • Baris pertama 14 hari (2024-01-01 s.d. 2024-01-14) dibuang:
#      dipakai sebagai burn-in untuk fitur rolling di sheet 7_Data_Final_Modeling
#    • Baris terakhir (2025-01-02) dibuang: pagu_buka_tebet = 0 (data tidak valid)
#    → tersisa 353 baris, konsisten dengan referensi riset
# ─────────────────────────────────────────────────────────────────────────────
df = df.sort_values("tanggal").reset_index(drop=True)
df = df[df[TARGET_COL] > 0].reset_index(drop=True)          # buang baris 0-value
df = df[df["tanggal"] >= "2024-01-15"].reset_index(drop=True)  # buang burn-in 14 hari
print(f"\nShape setelah filter : {df.shape}")
print(f"Rentang tanggal      : {df['tanggal'].min().date()} – {df['tanggal'].max().date()}")

# Ambil target sebagai 1-D array
data_values = df[TARGET_COL].values.astype(float)
dates_all   = df["tanggal"].values
n_total     = len(data_values)
print(f"Total baris          : {n_total}")

# ─────────────────────────────────────────────────────────────────────────────
# 5. Split kronologis 70 / 15 / 15
# ─────────────────────────────────────────────────────────────────────────────
n_train = int(n_total * 0.70)
n_val   = int(n_total * 0.15)
n_test  = n_total - n_train - n_val

train_data = data_values[:n_train]
val_data   = data_values[n_train : n_train + n_val]
test_data  = data_values[n_train + n_val :]

dates_test  = dates_all[n_train + n_val :]

print("\n5. SPLIT KRONOLOGIS")
print(f"  Train      : {n_train} baris  ({df['tanggal'].iloc[0].date()} – {df['tanggal'].iloc[n_train-1].date()})")
print(f"  Validation : {n_val}  baris  ({df['tanggal'].iloc[n_train].date()} – {df['tanggal'].iloc[n_train+n_val-1].date()})")
print(f"  Test       : {n_test}  baris  ({df['tanggal'].iloc[n_train+n_val].date()} – {df['tanggal'].iloc[-1].date()})")

# ─────────────────────────────────────────────────────────────────────────────
# 6. Scaling – fit HANYA pada train
# ─────────────────────────────────────────────────────────────────────────────
scaler = MinMaxScaler(feature_range=(0, 1))
train_scaled = scaler.fit_transform(train_data.reshape(-1, 1)).flatten()
val_scaled   = scaler.transform(val_data.reshape(-1, 1)).flatten()
test_scaled  = scaler.transform(test_data.reshape(-1, 1)).flatten()

print("\n6. SCALING  →  scaler fit pada train saja")
print(f"  Train min/max setelah scale: {train_scaled.min():.4f} / {train_scaled.max():.4f}")
print(f"  Val   min/max setelah scale: {val_scaled.min():.4f} / {val_scaled.max():.4f}")
print(f"  Test  min/max setelah scale: {test_scaled.min():.4f} / {val_scaled.max():.4f}")

# ─────────────────────────────────────────────────────────────────────────────
# 7. Fungsi bantu
# ─────────────────────────────────────────────────────────────────────────────
def create_sequences(series, window):
    """Bentuk pasangan (X, y) dari time series 1-D."""
    X, y = [], []
    for i in range(len(series) - window):
        X.append(series[i : i + window])
        y.append(series[i + window])
    return np.array(X).reshape(-1, window, 1), np.array(y)


def mape(y_true, y_pred):
    """MAPE, skip titik-titik di mana y_true = 0."""
    mask = y_true != 0
    return np.mean(np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])) * 100


def evaluate_metrics(y_true, y_pred, label):
    mae  = mean_absolute_error(y_true, y_pred)
    rmse = np.sqrt(mean_squared_error(y_true, y_pred))
    mape_val = mape(y_true, y_pred)
    r2   = r2_score(y_true, y_pred)
    print(f"  [{label}]  MAE={mae:,.0f}  RMSE={rmse:,.0f}  MAPE={mape_val:.4f}%  R²={r2:.4f}")
    return {"Model": label, "MAE": mae, "RMSE": rmse, "MAPE (%)": mape_val, "R²": r2}


# ─────────────────────────────────────────────────────────────────────────────
# 8. Baseline Moving Average 3 hari (MA3)
# ─────────────────────────────────────────────────────────────────────────────
print("\n8. BASELINE MA3")
all_data = np.concatenate([train_data, val_data, test_data])
ma3_pred = np.array([
    np.mean(all_data[n_train + n_val + i - 3 : n_train + n_val + i])
    for i in range(n_test)
])

metrics_list = []
metrics_list.append(evaluate_metrics(test_data, ma3_pred, "MA3"))

# ─────────────────────────────────────────────────────────────────────────────
# 9. LSTM untuk setiap window size
# ─────────────────────────────────────────────────────────────────────────────
print("\n9. LSTM TRAINING (window 5, 7, 10)")
lstm_results   = {}   # {window: (history, y_pred_inversed)}
best_window    = None
best_val_loss  = float("inf")

for W in WINDOW_SIZES:
    print(f"\n  --- Window = {W} ---")

    # Reset seed sebelum setiap model agar reprodusibel
    random.seed(SEED);  np.random.seed(SEED);  tf.random.set_seed(SEED)

    # Buat sequences
    # Training: input dari train_scaled saja
    X_train, y_train = create_sequences(train_scaled, W)

    # Validation: pinjam W nilai terakhir train sebagai konteks
    val_ctx  = np.concatenate([train_scaled[-W:], val_scaled])
    X_val, y_val = create_sequences(val_ctx, W)

    # Test: pinjam W nilai terakhir val sebagai konteks
    test_ctx = np.concatenate([val_scaled[-W:], test_scaled])
    X_test, y_test = create_sequences(test_ctx, W)

    print(f"  Shape X_train={X_train.shape}, X_val={X_val.shape}, X_test={X_test.shape}")

    # Bangun model
    model = Sequential([
        LSTM(LSTM_UNITS, activation="relu", input_shape=(W, 1)),
        Dense(1),
    ], name=f"LSTM_seq{W}")
    model.compile(optimizer="adam", loss="mse")

    es = EarlyStopping(
        monitor="val_loss",
        patience=PATIENCE,
        restore_best_weights=True,
        verbose=0,
    )

    history = model.fit(
        X_train, y_train,
        epochs=EPOCHS,
        batch_size=BATCH_SIZE,
        validation_data=(X_val, y_val),
        callbacks=[es],
        verbose=0,
    )

    stopped_epoch = len(history.history["loss"])
    print(f"  Berhenti epoch   : {stopped_epoch}")
    print(f"  Val loss terbaik : {min(history.history['val_loss']):.6f}")

    # Prediksi & inverse transform
    y_pred_scaled = model.predict(X_test, verbose=0).flatten()
    y_pred        = scaler.inverse_transform(y_pred_scaled.reshape(-1, 1)).flatten()
    y_true        = test_data   # skala asli

    result = evaluate_metrics(y_true, y_pred, f"LSTM_seq{W}")
    metrics_list.append(result)

    lstm_results[W] = {
        "history": history,
        "y_pred" : y_pred,
        "y_true" : y_true,
    }

    # Cari model terbaik (val_loss terendah saat berhenti)
    min_val_loss = min(history.history["val_loss"])
    if min_val_loss < best_val_loss:
        best_val_loss = min_val_loss
        best_window   = W

print(f"\n  Model terbaik: LSTM_seq{best_window}  (val_loss={best_val_loss:.6f})")

# ─────────────────────────────────────────────────────────────────────────────
# 10. Simpan tabel metrik
# ─────────────────────────────────────────────────────────────────────────────
df_metrics = pd.DataFrame(metrics_list)
df_metrics.to_csv(os.path.join(OUTPUT_DIR, "metrik_evaluasi.csv"), index=False)
print("\n10. TABEL METRIK AKHIR")
print(df_metrics.to_string(index=False))

# ─────────────────────────────────────────────────────────────────────────────
# 11. Grafik 1 – Loss curve model terbaik
# ─────────────────────────────────────────────────────────────────────────────
print("\n11. Membuat grafik loss curve ...")
best_history = lstm_results[best_window]["history"]

fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(best_history.history["loss"],     label="Loss Training",    color="#1f77b4", linewidth=1.8)
ax.plot(best_history.history["val_loss"], label="Loss Validasi",   color="#ff7f0e", linewidth=1.8, linestyle="--")
ax.set_xlabel("Epoch", fontsize=13)
ax.set_ylabel("Loss (MSE)", fontsize=13)
ax.legend(fontsize=12)
ax.tick_params(labelsize=11)
ax.grid(True, linestyle="--", alpha=0.5)
fig.tight_layout()
loss_path = os.path.join(OUTPUT_DIR, "loss_curve.png")
fig.savefig(loss_path, dpi=300, bbox_inches="tight")
plt.close(fig)
print(f"  Tersimpan → {loss_path}")

# ─────────────────────────────────────────────────────────────────────────────
# 12. Grafik 2 – Aktual vs MA3
# ─────────────────────────────────────────────────────────────────────────────
print("12. Membuat grafik aktual vs MA3 ...")
dates_test_dt = pd.to_datetime(dates_test)
y_true_m = test_data   # skala asli, sama untuk semua model

fig, ax = plt.subplots(figsize=(12, 5))
ax.plot(dates_test_dt, y_true_m / 1e9,  label="Aktual",       color="#1f77b4", linewidth=1.8)
ax.plot(dates_test_dt, ma3_pred  / 1e9, label="Prediksi MA3", color="#d62728", linewidth=1.8, linestyle="--")
ax.set_xlabel("Tanggal", fontsize=13)
ax.set_ylabel("Pagu Buka Tebet (Miliar Rp)", fontsize=13)
ax.legend(fontsize=12)
ax.tick_params(axis="x", rotation=30, labelsize=10)
ax.tick_params(axis="y", labelsize=11)
ax.grid(True, linestyle="--", alpha=0.5)
fig.tight_layout()
ma3_path = os.path.join(OUTPUT_DIR, "aktual_vs_ma3.png")
fig.savefig(ma3_path, dpi=300, bbox_inches="tight")
plt.close(fig)
print(f"  Tersimpan → {ma3_path}")

# ─────────────────────────────────────────────────────────────────────────────
# 13. Grafik 3 – Aktual vs LSTM seq10
# ─────────────────────────────────────────────────────────────────────────────
print("13. Membuat grafik aktual vs LSTM seq10 ...")
lstm10_pred = lstm_results[10]["y_pred"]
lstm10_true = lstm_results[10]["y_true"]

fig, ax = plt.subplots(figsize=(12, 5))
ax.plot(dates_test_dt, lstm10_true / 1e9,  label="Aktual",          color="#1f77b4", linewidth=1.8)
ax.plot(dates_test_dt, lstm10_pred  / 1e9, label="Prediksi LSTM seq10", color="#2ca02c", linewidth=1.8, linestyle="--")
ax.set_xlabel("Tanggal", fontsize=13)
ax.set_ylabel("Pagu Buka Tebet (Miliar Rp)", fontsize=13)
ax.legend(fontsize=12)
ax.tick_params(axis="x", rotation=30, labelsize=10)
ax.tick_params(axis="y", labelsize=11)
ax.grid(True, linestyle="--", alpha=0.5)
fig.tight_layout()
lstm_path = os.path.join(OUTPUT_DIR, "aktual_vs_lstm.png")
fig.savefig(lstm_path, dpi=300, bbox_inches="tight")
plt.close(fig)
print(f"  Tersimpan → {lstm_path}")

# ─────────────────────────────────────────────────────────────────────────────
# 14. Laporan akhir & perbandingan dengan referensi
# ─────────────────────────────────────────────────────────────────────────────
REF = {
    "MA3"        : {"MAPE (%)": 5.89,  "R²": 0.467},
    "LSTM_seq5"  : {"MAPE (%)": None,  "R²": None},
    "LSTM_seq7"  : {"MAPE (%)": None,  "R²": None},
    "LSTM_seq10" : {"MAPE (%)": 10.58, "R²": -0.396},
}

print("\n" + "="*70)
print("14. LAPORAN AKHIR — PERBANDINGAN DENGAN REFERENSI SKRIPSI")
print("="*70)
print(f"{'Model':<16} {'MAPE% Baru':>12} {'MAPE% Ref':>12} {'R² Baru':>10} {'R² Ref':>10}")
print("-"*70)
for row in metrics_list:
    m   = row["Model"]
    ref = REF.get(m, {})
    mp  = f"{row['MAPE (%)']:.4f}"
    mr  = f"{ref['MAPE (%)']:.2f}" if ref.get("MAPE (%)") else "—"
    r2  = f"{row['R²']:.4f}"
    r2r = f"{ref['R²']:.3f}"      if ref.get("R²") is not None else "—"
    print(f"{m:<16} {mp:>12} {mr:>12} {r2:>10} {r2r:>10}")

print("\nCATATAN INTERPRETASI:")
ma3_row  = next(r for r in metrics_list if r["Model"] == "MA3")
lstm10_row = next(r for r in metrics_list if r["Model"] == "LSTM_seq10")
if ma3_row["MAPE (%)"] < lstm10_row["MAPE (%)"]:
    print("  ✓ Pola REFERENSI terkonfirmasi: MA3 lebih baik dari LSTM_seq10.")
    print("    Ini konsisten dengan dataset kecil (247 train) + volatilitas keuangan.")
else:
    print("  ✗ PERHATIAN: LSTM_seq10 lebih baik dari MA3 pada run ini.")
    print("    Periksa kemungkinan data leakage atau perubahan seed/versi library.")

print(f"\n  Versi TensorFlow : {tf.__version__}")
print(f"  Seed global      : {SEED}")
print(f"  Model terbaik    : LSTM_seq{best_window}")
print("="*70)
print("Eksperimen selesai. Semua output tersimpan di folder:", OUTPUT_DIR)
