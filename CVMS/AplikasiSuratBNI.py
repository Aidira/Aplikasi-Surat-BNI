import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import numpy as np
import sqlite3
import os
import sys
from datetime import datetime
from PIL import Image, ImageTk

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from models import predictor as pred
import surat_generator as suratgen

# --- KONFIGURASI DATABASE ---
DB_NAME = "database_bni.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS riwayat_over (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal_input DATE,
            no_surat TEXT,
            cabang TEXT,
            mata_uang TEXT,
            saldo REAL,
            pagu REAL,
            over_limit REAL
        )
    ''')
    conn.commit()
    conn.close()

# --- LOGIKA APLIKASI ---
class AppBNI(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly", title="BNI Asuransi Dashboard - Desktop")
        self.geometry("1100x700")
        init_db()
        
        self.df_current = pd.DataFrame()
        self.create_widgets()

    def create_widgets(self):
        # Sidebar
        sidebar = ttk.Frame(self, bootstyle="light", width=250, padding=10)
        sidebar.pack(side=LEFT, fill=Y)

        ttk.Label(sidebar, text="PENGATURAN", font=("Helvetica", 12, "bold")).pack(pady=10)
        
        ttk.Label(sidebar, text="Nomor Surat:").pack(anchor=W)
        self.ent_no_surat = ttk.Entry(sidebar)
        self.ent_no_surat.pack(fill=X, pady=5)

        ttk.Label(sidebar, text="Nama Manager:").pack(anchor=W)
        self.ent_manager = ttk.Entry(sidebar)
        self.ent_manager.insert(0, "Hasbiallah")
        self.ent_manager.pack(fill=X, pady=5)

        ttk.Label(sidebar, text="Jenis Surat:").pack(anchor=W)
        self.cmb_jenis_surat = ttk.Combobox(
            sidebar, values=list(suratgen.TEMPLATE_SURAT.keys()), state="readonly",
        )
        self.cmb_jenis_surat.current(0)
        self.cmb_jenis_surat.pack(fill=X, pady=5)

        ttk.Label(sidebar, text="Ambang Batas Over (% dari Pagu):").pack(anchor=W, pady=(10, 0))
        self.ent_ambang_batas = ttk.Entry(sidebar)
        self.ent_ambang_batas.insert(0, "20")
        self.ent_ambang_batas.pack(fill=X, pady=5)
        self.ent_ambang_batas.bind("<Return>", lambda e: self._refresh_tree_preview())
        self.ent_ambang_batas.bind("<FocusOut>", lambda e: self._refresh_tree_preview())

        self.btn_upload = ttk.Button(sidebar, text="Upload Excel", bootstyle="info", command=self.load_excel)
        self.btn_upload.pack(fill=X, pady=20)

        self.btn_save = ttk.Button(sidebar, text="Simpan ke DB", bootstyle="success", command=self.save_data)
        self.btn_save.pack(fill=X, pady=5)

        self.btn_cetak = ttk.Button(sidebar, text="Cetak Surat (PDF)", bootstyle="warning", command=self.cetak_surat)
        self.btn_cetak.pack(fill=X, pady=5)

        # Main Area (Tabs)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)

        # Tab 1: Data Editor / Preview
        self.tab_preview = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_preview, text="Preview Data")
        
        ttk.Label(
            self.tab_preview,
            text="Baris berwarna merah = over-limit melebihi ambang batas (lihat isian \"Ambang Batas Over\" di sidebar).",
            font=("Helvetica", 8, "italic"), bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 5))

        self.tree = ttk.Treeview(self.tab_preview, columns=("cabang", "mata_uang", "saldo", "pagu", "over", "over_pct"), show="headings")
        kolom_label = {"cabang": "CABANG", "mata_uang": "MATA_UANG", "saldo": "SALDO", "pagu": "PAGU", "over": "OVER", "over_pct": "OVER (%)"}
        for col in self.tree["columns"]:
            self.tree.heading(col, text=kolom_label[col])
            self.tree.column(col, width=140, anchor=CENTER)
        self.tree.tag_configure("kritis", background="#f8d7da")
        self.tree.pack(fill=BOTH, expand=YES)

        # Tab 2: Prediksi Pagu Kas — dashboard prediksi untuk hari berikutnya
        self.tab_prediksi = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_prediksi, text="Prediksi Pagu Kas")
        self.build_tab_prediksi()

        # Tab 3: Preview Perhitungan — rincian langkah hitung MA3 & LSTM
        self.tab_perhitungan = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_perhitungan, text="Preview Perhitungan")
        self.build_tab_perhitungan()

        # Tab 4: Riwayat & Tren — riwayat over-limit tersimpan dan grafik tren per cabang
        self.tab_riwayat = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_riwayat, text="Riwayat & Tren")
        self.build_tab_riwayat()

    def build_tab_prediksi(self):
        """Membangun dashboard Prediksi Pagu Kas: upload dataset, lalu
        tampilkan prediksi pagu buka untuk hari berikutnya (MA3 sebagai
        model utama -- sesuai temuan skripsi bahwa MA3 lebih akurat -- dan
        LSTM sebagai pembanding), beserta tabel metrik evaluasi historis."""
        self.df_pagu = None
        self.lstm_model = None
        self.lstm_scaler = None
        self.hasil_perhitungan = None

        top = ttk.Frame(self.tab_prediksi)
        top.pack(fill=X, pady=(0, 10))

        ttk.Label(
            top,
            text="Prototipe Penelitian Skripsi — bukan sistem operasional BNI",
            font=("Helvetica", 9, "italic"),
            bootstyle="danger",
        ).pack(anchor=W)

        ttk.Button(
            top, text="Upload Dataset Pagu (.xlsx)", bootstyle="info",
            command=self.upload_dataset_pagu,
        ).pack(side=LEFT, pady=5)

        status_text = "LSTM aktif (TensorFlow tersedia)" if pred.TF_AVAILABLE else \
            "LSTM TIDAK AKTIF — TensorFlow tidak tersedia, hanya baseline MA3 yang berjalan"
        self.lbl_status_lstm = ttk.Label(
            top, text=status_text,
            bootstyle="success" if pred.TF_AVAILABLE else "warning",
            font=("Helvetica", 9, "bold"),
        )
        self.lbl_status_lstm.pack(side=LEFT, padx=15)

        body = ttk.Frame(self.tab_prediksi)
        body.pack(fill=BOTH, expand=YES)

        # Panel prediksi besok
        self.lbl_tanggal_prediksi = ttk.Label(
            body, text="Belum ada data — upload dataset untuk melihat prediksi.",
            font=("Helvetica", 11, "bold"),
        )
        self.lbl_tanggal_prediksi.pack(anchor=W, pady=(0, 10))

        kartu_frame = ttk.Frame(body)
        kartu_frame.pack(fill=X, pady=(0, 15))

        kartu_ma3 = ttk.Labelframe(kartu_frame, text="Model Utama: MA3", bootstyle="success", padding=15)
        kartu_ma3.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 10))
        self.lbl_pred_ma3 = ttk.Label(kartu_ma3, text="Rp -", font=("Helvetica", 22, "bold"), bootstyle="success")
        self.lbl_pred_ma3.pack()
        ttk.Label(kartu_ma3, text="Rata-rata 3 hari aktual terakhir", font=("Helvetica", 8)).pack()

        kartu_lstm = ttk.Labelframe(kartu_frame, text="Pembanding: LSTM", bootstyle="secondary", padding=15)
        kartu_lstm.pack(side=LEFT, fill=BOTH, expand=YES)
        self.lbl_pred_lstm = ttk.Label(kartu_lstm, text="Rp -", font=("Helvetica", 18, "bold"), bootstyle="secondary")
        self.lbl_pred_lstm.pack()
        ttk.Label(kartu_lstm, text="Sequence 10 hari aktual terakhir", font=("Helvetica", 8)).pack()

        ttk.Label(
            body,
            text="Catatan: MA3 ditampilkan sebagai model utama karena terbukti lebih akurat "
                 "pada evaluasi historis di bawah ini (lihat tab Preview Perhitungan untuk rincian langkahnya).",
            font=("Helvetica", 8, "italic"), wraplength=850,
        ).pack(anchor=W, pady=(0, 15))

        # Tabel metrik pembanding (evaluasi historis pada test set)
        ttk.Label(body, text="Tabel Metrik Pembanding (evaluasi historis, test set)",
                  font=("Helvetica", 11, "bold")).pack(anchor=W, pady=(0, 2))
        self.tree_metrik = ttk.Treeview(
            body, columns=("model", "mae", "rmse", "mape", "r2"), show="headings", height=3
        )
        for col, label in zip(
            ("model", "mae", "rmse", "mape", "r2"),
            ("Model", "MAE (Rp)", "RMSE (Rp)", "MAPE (%)", "R²"),
        ):
            self.tree_metrik.heading(col, text=label)
            self.tree_metrik.column(col, width=140, anchor=CENTER)
        self.tree_metrik.pack(fill=X)

    def build_tab_perhitungan(self):
        """Membangun tab Preview Perhitungan: rincian langkah hitung MA3 dan
        LSTM untuk prediksi besok, beserta grafik evaluasi aktual vs prediksi
        pada test set -- agar mudah dicocokkan/dijelaskan manual saat sidang."""
        body = ttk.Frame(self.tab_perhitungan)
        body.pack(fill=BOTH, expand=YES)

        ttk.Label(body, text="Langkah Perhitungan MA3", font=("Helvetica", 11, "bold")).pack(anchor=W)
        self.txt_perhitungan_ma3 = tk.Text(body, height=7, font=("Consolas", 9))
        self.txt_perhitungan_ma3.pack(fill=X, pady=(0, 10))
        self.txt_perhitungan_ma3.insert(END, "Belum ada data. Upload dataset pagu di tab Prediksi Pagu Kas.")
        self.txt_perhitungan_ma3.config(state="disabled")

        ttk.Label(body, text="Langkah Perhitungan LSTM", font=("Helvetica", 11, "bold")).pack(anchor=W)
        self.txt_perhitungan_lstm = tk.Text(body, height=9, font=("Consolas", 9))
        self.txt_perhitungan_lstm.pack(fill=X, pady=(0, 10))
        self.txt_perhitungan_lstm.insert(END, "Belum ada data. Upload dataset pagu di tab Prediksi Pagu Kas.")
        self.txt_perhitungan_lstm.config(state="disabled")

        ttk.Label(body, text="Evaluasi Historis: Aktual vs Prediksi (Test Set)",
                  font=("Helvetica", 11, "bold")).pack(anchor=W)
        chart_frame = ttk.Frame(body)
        chart_frame.pack(fill=BOTH, expand=YES)

        self.fig_pred = Figure(figsize=(9, 3.5), dpi=100)
        self.ax_ma3 = self.fig_pred.add_subplot(121)
        self.ax_lstm = self.fig_pred.add_subplot(122)
        self.ax_ma3.set_title("Aktual vs MA3")
        self.ax_lstm.set_title("Aktual vs LSTM")
        self.fig_pred.tight_layout()

        self.canvas_pred = FigureCanvasTkAgg(self.fig_pred, master=chart_frame)
        self.canvas_pred.get_tk_widget().pack(fill=BOTH, expand=YES)

    def build_tab_riwayat(self):
        """Membangun tab Riwayat & Tren: menampilkan riwayat over-limit yang
        tersimpan di database (filter per cabang), tombol export ke Excel,
        dan grafik tren over-limit per cabang dari waktu ke waktu."""
        top = ttk.Frame(self.tab_riwayat)
        top.pack(fill=X, pady=(0, 10))

        ttk.Label(top, text="Filter Cabang:").pack(side=LEFT, padx=(0, 5))
        self.cmb_filter_cabang = ttk.Combobox(top, state="readonly", width=25)
        self.cmb_filter_cabang.pack(side=LEFT, padx=(0, 10))

        ttk.Button(top, text="Muat Riwayat", bootstyle="info",
                   command=self.muat_riwayat).pack(side=LEFT, padx=5)
        ttk.Button(top, text="Export ke Excel", bootstyle="secondary",
                   command=self.export_riwayat_excel).pack(side=LEFT, padx=5)
        ttk.Button(top, text="Tampilkan Tren Cabang", bootstyle="primary",
                   command=self.tampilkan_tren_cabang).pack(side=LEFT, padx=5)

        body = ttk.Panedwindow(self.tab_riwayat, orient=VERTICAL)
        body.pack(fill=BOTH, expand=YES)

        tabel_frame = ttk.Frame(body)
        body.add(tabel_frame, weight=1)

        self.tree_riwayat = ttk.Treeview(
            tabel_frame,
            columns=("tanggal", "no_surat", "cabang", "mata_uang", "saldo", "pagu", "over"),
            show="headings", height=8,
        )
        label_riwayat = {
            "tanggal": "Tanggal Input", "no_surat": "No. Surat", "cabang": "Cabang",
            "mata_uang": "Mata Uang", "saldo": "Saldo", "pagu": "Pagu", "over": "Over",
        }
        for col in self.tree_riwayat["columns"]:
            self.tree_riwayat.heading(col, text=label_riwayat[col])
            self.tree_riwayat.column(col, width=120, anchor=CENTER)
        self.tree_riwayat.pack(fill=BOTH, expand=YES)

        chart_frame = ttk.Frame(body)
        body.add(chart_frame, weight=1)

        self.fig_tren = Figure(figsize=(9, 3), dpi=100)
        self.ax_tren = self.fig_tren.add_subplot(111)
        self.ax_tren.set_title("Tren Over-Limit per Cabang")
        self.fig_tren.tight_layout()
        self.canvas_tren = FigureCanvasTkAgg(self.fig_tren, master=chart_frame)
        self.canvas_tren.get_tk_widget().pack(fill=BOTH, expand=YES)

        self.muat_riwayat()

    def _query_riwayat(self, cabang=None):
        """Mengambil data riwayat_over dari database, opsional difilter per cabang."""
        conn = sqlite3.connect(DB_NAME)
        if cabang:
            df = pd.read_sql_query(
                "SELECT * FROM riwayat_over WHERE cabang = ? ORDER BY tanggal_input", conn, params=(cabang,),
            )
        else:
            df = pd.read_sql_query("SELECT * FROM riwayat_over ORDER BY tanggal_input", conn)
        conn.close()
        return df

    def muat_riwayat(self):
        """Memuat ulang tabel riwayat dari database sesuai filter cabang yang dipilih,
        dan memperbarui daftar pilihan cabang pada combobox filter."""
        df_semua = self._query_riwayat()
        daftar_cabang = sorted(df_semua["cabang"].unique().tolist()) if not df_semua.empty else []
        self.cmb_filter_cabang["values"] = ["(Semua Cabang)"] + daftar_cabang
        if not self.cmb_filter_cabang.get():
            self.cmb_filter_cabang.current(0)

        pilihan = self.cmb_filter_cabang.get()
        cabang = None if pilihan in ("", "(Semua Cabang)") else pilihan
        df = self._query_riwayat(cabang)

        for i in self.tree_riwayat.get_children():
            self.tree_riwayat.delete(i)
        for _, row in df.iterrows():
            self.tree_riwayat.insert("", END, values=(
                row["tanggal_input"], row["no_surat"], row["cabang"], row["mata_uang"],
                f"{row['saldo']:,.0f}", f"{row['pagu']:,.0f}", f"{row['over_limit']:,.0f}",
            ))
        self._df_riwayat_terkini = df

    def export_riwayat_excel(self):
        """Mengekspor riwayat over-limit yang sedang tampil ke file Excel."""
        df = getattr(self, "_df_riwayat_terkini", pd.DataFrame())
        if df.empty:
            messagebox.showwarning("Peringatan", "Tidak ada data riwayat untuk diekspor.")
            return
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
            initialfile="Riwayat_Over_Limit.xlsx",
        )
        if not output_path:
            return
        try:
            df.to_excel(output_path, index=False)
            messagebox.showinfo("Sukses", f"Riwayat berhasil diekspor:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor riwayat: {e}")

    def tampilkan_tren_cabang(self):
        """Menampilkan grafik tren over-limit untuk cabang yang difilter pada
        combobox, atau rata-rata seluruh cabang jika 'Semua Cabang' dipilih."""
        pilihan = self.cmb_filter_cabang.get()
        cabang = None if pilihan in ("", "(Semua Cabang)") else pilihan
        df = self._query_riwayat(cabang)

        self.ax_tren.clear()
        if df.empty:
            self.ax_tren.set_title("Belum ada riwayat untuk ditampilkan")
        elif cabang:
            self.ax_tren.plot(df["tanggal_input"], df["over_limit"], marker="o")
            self.ax_tren.set_title(f"Tren Over-Limit — {cabang}")
            self.ax_tren.tick_params(axis="x", rotation=45)
        else:
            for cab, grup in df.groupby("cabang"):
                self.ax_tren.plot(grup["tanggal_input"], grup["over_limit"], marker="o", label=cab)
            self.ax_tren.set_title("Tren Over-Limit — Semua Cabang")
            self.ax_tren.legend(fontsize=7)
            self.ax_tren.tick_params(axis="x", rotation=45)
        self.ax_tren.set_ylabel("Over-Limit")
        self.fig_tren.tight_layout()
        self.canvas_tren.draw()

    def _ambil_ambang_batas(self):
        """Membaca nilai ambang batas over-limit (%) dari sidebar; default 20% jika input tidak valid."""
        try:
            return float(self.ent_ambang_batas.get())
        except (ValueError, AttributeError):
            return 20.0

    def _refresh_tree_preview(self):
        """Mengisi ulang Treeview Preview Data, menandai baris 'kritis' (merah) jika
        persentase over-limit melebihi ambang batas yang diatur pengguna di sidebar."""
        ambang = self._ambil_ambang_batas()
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, row in self.df_current.iterrows():
            tag = "kritis" if row["Over %"] >= ambang else ""
            self.tree.insert("", END, values=list(row), tags=(tag,) if tag else ())

    def _isi_teks(self, widget, isi):
        """Mengisi widget Text read-only dengan konten baru."""
        widget.config(state="normal")
        widget.delete("1.0", END)
        widget.insert(END, isi)
        widget.config(state="disabled")

    def upload_dataset_pagu(self):
        """
        Memuat dataset pagu kas harian, melatih/menjalankan MA3 & LSTM, lalu
        menampilkan:
        - Dashboard prediksi pagu buka untuk hari berikutnya (tab Prediksi Pagu Kas)
        - Rincian langkah perhitungan dan grafik evaluasi historis (tab Preview Perhitungan)
        """
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            df = pred.muat_dataset(path)
            series = df[pred.TARGET_COL].to_numpy(dtype=float)
            train, val, test = pred.split_kronologis(series)
            tanggal_terakhir = df["tanggal"].iloc[-1] if "tanggal" in df.columns else None
            tanggal_besok = tanggal_terakhir + pd.Timedelta(days=1) if tanggal_terakhir is not None else None

            # --- Evaluasi historis MA3 (test set) ---
            ma3_eval = pred.prediksi_ma3(np.concatenate([train, val]), test)
            valid_mask = ~np.isnan(ma3_eval)
            metrik_ma3 = pred.hitung_metrik(test[valid_mask], ma3_eval[valid_mask])

            self.ax_ma3.clear()
            self.ax_ma3.plot(test[valid_mask], label="Aktual")
            self.ax_ma3.plot(ma3_eval[valid_mask], label="Prediksi MA3")
            self.ax_ma3.set_xlabel("Hari ke-")
            self.ax_ma3.set_ylabel("Pagu Buka (Rp)")
            self.ax_ma3.set_title("Aktual vs MA3")
            self.ax_ma3.legend(fontsize=8)

            self.tree_metrik.delete(*self.tree_metrik.get_children())
            self.tree_metrik.insert("", END, values=(
                "MA3", f"{metrik_ma3['MAE']:,.0f}", f"{metrik_ma3['RMSE']:,.0f}",
                f"{metrik_ma3['MAPE (%)']:.2f}", f"{metrik_ma3['R2']:.3f}",
            ))

            # --- Prediksi besok: MA3 (model utama) ---
            pred_besok_ma3, last3 = pred.prediksi_besok_ma3(series)
            label_tanggal = f"untuk {tanggal_besok.date()}" if tanggal_besok is not None else "untuk hari berikutnya"
            self.lbl_tanggal_prediksi.config(text=f"Prediksi Pagu Buka {label_tanggal}")
            self.lbl_pred_ma3.config(text=f"Rp {pred_besok_ma3:,.0f}")

            teks_ma3 = (
                "Rumus: MA3 = (nilai_1 + nilai_2 + nilai_3) / 3\n\n"
                "3 nilai aktual terakhir yang dipakai:\n"
            )
            for i, v in enumerate(last3, start=1):
                teks_ma3 += f"  nilai_{i} = Rp {v:,.0f}\n"
            teks_ma3 += (
                f"\nPerhitungan: ({last3[0]:,.0f} + {last3[1]:,.0f} + {last3[2]:,.0f}) / 3"
                f"\nHasil prediksi besok (MA3) = Rp {pred_besok_ma3:,.0f}"
            )
            self._isi_teks(self.txt_perhitungan_ma3, teks_ma3)

            # --- LSTM (jika TensorFlow tersedia) ---
            if pred.TF_AVAILABLE:
                self.lstm_model, self.lstm_scaler, history = pred.latih_lstm(train, val)
                self.lstm_history = history
                context = np.concatenate([train, val])

                # Evaluasi historis LSTM (test set)
                lstm_eval = pred.prediksi_lstm(self.lstm_model, self.lstm_scaler, context, test)
                test_eval = test[pred.WINDOW_DEFAULT:] if len(test) > pred.WINDOW_DEFAULT else test
                metrik_lstm = pred.hitung_metrik(test_eval[:len(lstm_eval)], lstm_eval[:len(test_eval)])

                self.ax_lstm.clear()
                self.ax_lstm.plot(test_eval[:len(lstm_eval)], label="Aktual")
                self.ax_lstm.plot(lstm_eval[:len(test_eval)], label="Prediksi LSTM")
                self.ax_lstm.set_xlabel("Hari ke-")
                self.ax_lstm.set_ylabel("Pagu Buka (Rp)")
                self.ax_lstm.set_title("Aktual vs LSTM")
                self.ax_lstm.legend(fontsize=8)

                self.tree_metrik.insert("", END, values=(
                    "LSTM", f"{metrik_lstm['MAE']:,.0f}", f"{metrik_lstm['RMSE']:,.0f}",
                    f"{metrik_lstm['MAPE (%)']:.2f}", f"{metrik_lstm['R2']:.3f}",
                ))

                # Prediksi besok: LSTM (pembanding)
                pred_besok_lstm, window_vals, window_scaled, yhat_scaled = pred.prediksi_besok_lstm(
                    self.lstm_model, self.lstm_scaler, series
                )
                self.lbl_pred_lstm.config(text=f"Rp {pred_besok_lstm:,.0f}")

                teks_lstm = (
                    f"Scaler MinMaxScaler (fit hanya pada data train):\n"
                    f"  min_train = Rp {self.lstm_scaler.data_min_[0]:,.0f}\n"
                    f"  max_train = Rp {self.lstm_scaler.data_max_[0]:,.0f}\n\n"
                    f"Sequence {pred.WINDOW_DEFAULT} nilai aktual terakhir (input model):\n"
                )
                for i, (raw, sc) in enumerate(zip(window_vals, window_scaled), start=1):
                    teks_lstm += f"  hari_{i}: raw = Rp {raw:,.0f}  ->  normalized = {sc:.6f}\n"
                teks_lstm += (
                    f"\nOutput model (skala ternormalisasi) = {yhat_scaled:.6f}\n"
                    f"Inverse transform ke Rupiah = Rp {pred_besok_lstm:,.0f}"
                )
                self._isi_teks(self.txt_perhitungan_lstm, teks_lstm)
            else:
                self.ax_lstm.clear()
                self.ax_lstm.set_title("LSTM tidak aktif (TensorFlow tidak tersedia)")
                self.lbl_pred_lstm.config(text="Tidak aktif")
                self._isi_teks(
                    self.txt_perhitungan_lstm,
                    "LSTM tidak aktif — TensorFlow tidak tersedia di environment ini.",
                )

            self.fig_pred.tight_layout()
            self.canvas_pred.draw()

            messagebox.showinfo("Sukses", f"Prediksi selesai untuk {len(df)} baris data.")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menjalankan prediksi: {e}")

    def bersihkan_angka(self, nilai_raw):
        try:
            if isinstance(nilai_raw, (int, float)): return float(nilai_raw)
            text = str(nilai_raw).upper().replace("IDR", "").replace("RP", "").replace(" ", "").strip()
            if "." in text and "," in text: text = text.replace(".", "").replace(",", ".")
            elif "." in text: text = text.replace(".", "")
            elif "," in text: text = text.replace(",", ".")
            return float(text)
        except: return 0.0

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path: return

        try:
            xls = pd.ExcelFile(path)
            data_over = []
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                curr = "USD" if "USD" in sheet.upper() else "IDR"
                for index, row in df.iterrows():
                    if len(row) < 4: continue
                    cabang = str(row[1])
                    if pd.isna(cabang) or "TOTAL" in cabang or "NAMA" in cabang.upper() or "KCU" in cabang: continue
                    pagu = self.bersihkan_angka(row[2])
                    saldo = self.bersihkan_angka(row[3])
                    over = saldo - pagu
                    if over > 0:
                        data_over.append([cabang, curr, saldo, pagu, over])
            
            self.df_current = pd.DataFrame(data_over, columns=["Cabang", "Mata Uang", "Saldo", "Pagu", "Over"])
            self.df_current["Over %"] = (self.df_current["Over"] / self.df_current["Pagu"] * 100).round(2)

            self._refresh_tree_preview()
            
            messagebox.showinfo("Sukses", f"Berhasil memuat {len(self.df_current)} data.")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file: {e}")

    def save_data(self):
        if self.df_current.empty:
            messagebox.showwarning("Peringatan", "Data kosong!")
            return
        
        no_surat = self.ent_no_surat.get()
        if not no_surat:
            messagebox.showwarning("Peringatan", "Isi nomor surat!")
            return

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        tgl = datetime.now().strftime("%Y-%m-%d")
        
        for _, row in self.df_current.iterrows():
            c.execute('''INSERT INTO riwayat_over (tanggal_input, no_surat, cabang, mata_uang, saldo, pagu, over_limit)
                         VALUES (?, ?, ?, ?, ?, ?, ?)''', 
                      (tgl, no_surat, row['Cabang'], row['Mata Uang'], row['Saldo'], row['Pagu'], row['Over']))
        conn.commit()
        conn.close()
        if hasattr(self, "tree_riwayat"):
            self.muat_riwayat()
        messagebox.showinfo("Sukses", "Data berhasil disimpan ke database!")

    def cetak_surat(self):
        """Membuat surat PDF 'Cover Asuransi CIS Saldo Kas IDR dan Valas KC/KCP/KK'
        dari data over-limit yang sedang dimuat, lalu menyimpannya ke file PDF."""
        if self.df_current.empty:
            messagebox.showwarning("Peringatan", "Data kosong! Upload Excel data over-limit terlebih dahulu.")
            return

        no_surat = self.ent_no_surat.get()
        if not no_surat:
            messagebox.showwarning("Peringatan", "Isi nomor surat!")
            return

        nama_manager = self.ent_manager.get() or "Hasbiallah"

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"Surat_Over_Limit_{no_surat.replace('/', '-')}.pdf",
        )
        if not output_path:
            return

        try:
            rows = [
                {
                    "cabang": row["Cabang"],
                    "mata_uang": row["Mata Uang"],
                    "saldo": row["Saldo"],
                    "pagu": row["Pagu"],
                    "over": row["Over"],
                }
                for _, row in self.df_current.iterrows()
            ]
            jenis_surat = self.cmb_jenis_surat.get() or None
            suratgen.buat_surat_pdf(
                output_path, no_surat, nama_manager, datetime.now(), rows, jenis_surat=jenis_surat,
            )
            messagebox.showinfo("Sukses", f"Surat berhasil dibuat:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat surat PDF: {e}")

if __name__ == "__main__":
    app = AppBNI()
    app.mainloop()
