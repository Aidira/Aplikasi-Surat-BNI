import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import Colors, ThemeDefinition
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

# --- PALET WARNA KORPORAT BNI, gaya dashboard gelap (oranye & tosca dari logo.png) ---
WARNA_BNI_ORANGE = "#F15A23"
WARNA_BNI_TOSCA = "#00838F"
WARNA_BNI_GELAP = "#10171F"
WARNA_BNI_PANEL = "#1A232C"
WARNA_BNI_GARIS = "#2B3742"

BNI_THEME = ThemeDefinition(
    name="bni",
    themetype=DARK,
    colors=Colors(
        primary=WARNA_BNI_ORANGE,
        secondary=WARNA_BNI_TOSCA,
        success="#2ECC71",
        info="#3DAFD0",
        warning="#F2A104",
        danger="#E74C3C",
        light=WARNA_BNI_PANEL,
        dark=WARNA_BNI_GELAP,
        bg=WARNA_BNI_GELAP,
        fg="#E8ECEF",
        selectbg=WARNA_BNI_ORANGE,
        selectfg="#FFFFFF",
        border=WARNA_BNI_GARIS,
        inputfg="#E8ECEF",
        inputbg=WARNA_BNI_PANEL,
        active=WARNA_BNI_ORANGE,
    ),
)

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
    c.execute('''
        CREATE TABLE IF NOT EXISTS supply_remise (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal DATE,
            outlet TEXT,
            jenis TEXT,
            mata_uang TEXT,
            nominal REAL,
            keterangan TEXT
        )
    ''')
    conn.commit()
    conn.close()

# --- LOGIKA APLIKASI ---
class AppBNI(ttk.Window):
    def __init__(self):
        super().__init__(title="BNI - Manajemen Over-Limit Kas KC/KCP/KK")
        self.style.register_theme(BNI_THEME)
        self.style.theme_use("bni")
        self.geometry("1200x760")
        self.minsize(1000, 650)
        init_db()

        self.df_current = pd.DataFrame()
        self._setup_styles()
        self.build_header()
        self.create_widgets()
        self.build_statusbar()

    def _setup_styles(self):
        """Menata ulang gaya widget global (font, padding, tinggi baris tabel)
        agar tampilan lebih rapi dan konsisten di seluruh aplikasi."""
        style = self.style
        style.configure(".", font=("Helvetica", 10))
        style.configure("TButton", font=("Helvetica", 10, "bold"), padding=(10, 9))
        style.configure("TLabelframe.Label", font=("Helvetica", 10, "bold"))
        style.configure("TLabelframe", borderwidth=1)
        style.configure(
            "Treeview", font=("Helvetica", 9), rowheight=26,
            fieldbackground=WARNA_BNI_PANEL, borderwidth=0,
        )
        style.configure(
            "Treeview.Heading", font=("Helvetica", 9, "bold"),
            padding=(6, 8), foreground="#FFFFFF", background=WARNA_BNI_GARIS,
        )
        style.map("Treeview", background=[("selected", WARNA_BNI_TOSCA)])
        style.configure("TNotebook.Tab", font=("Helvetica", 10, "bold"), padding=(16, 8))
        style.configure("TEntry", padding=(6, 6))
        style.configure("TCombobox", padding=(6, 6))

    def build_header(self):
        """Membangun banner header korporat (logo BNI + judul aplikasi) di
        bagian atas jendela, agar tampilan terasa seperti aplikasi perbankan resmi."""
        header = ttk.Frame(self, bootstyle="secondary")
        header.pack(side=TOP, fill=X)

        isi = ttk.Frame(header, bootstyle="secondary", padding=(20, 12))
        isi.pack(fill=X)

        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
        if os.path.exists(logo_path):
            img = Image.open(logo_path)
            ratio = img.height / img.width
            img = img.resize((110, int(110 * ratio)))
            self._logo_img = ImageTk.PhotoImage(img)
            ttk.Label(isi, image=self._logo_img, bootstyle="inverse-secondary").pack(side=LEFT, padx=(0, 15))

        teks_frame = ttk.Frame(isi, bootstyle="secondary")
        teks_frame.pack(side=LEFT, fill=Y)
        ttk.Label(
            teks_frame, text="Manajemen Over-Limit Kas KC/KCP/KK",
            font=("Helvetica", 16, "bold"), bootstyle="inverse-secondary",
        ).pack(anchor=W)
        ttk.Label(
            teks_frame, text="BNI Kantor Cabang Tebet — Prototipe Internal",
            font=("Helvetica", 9), bootstyle="inverse-secondary",
        ).pack(anchor=W)

        ttk.Frame(self, bootstyle="primary", height=3).pack(side=TOP, fill=X)

    def build_statusbar(self):
        """Status bar di bagian bawah jendela menampilkan ringkasan singkat data yang dimuat."""
        ttk.Separator(self).pack(side=BOTTOM, fill=X)
        bar = ttk.Frame(self, bootstyle="light", padding=(18, 7))
        bar.pack(side=BOTTOM, fill=X)
        self.lbl_status = ttk.Label(
            bar, text="Siap. Belum ada data dimuat.", font=("Helvetica", 9), bootstyle="secondary",
        )
        self.lbl_status.pack(side=LEFT)

    def _set_status(self, teks):
        if hasattr(self, "lbl_status"):
            self.lbl_status.config(text=teks)

    @staticmethod
    def _gaya_chart_gelap(fig, *axes):
        """Menerapkan palet gelap (latar panel, teks & garis terang) pada
        figure/axes matplotlib agar konsisten dengan tema dashboard gelap,
        alih-alih latar putih default yang kontras secara mencolok."""
        fig.patch.set_facecolor(WARNA_BNI_PANEL)
        for ax in axes:
            ax.set_facecolor(WARNA_BNI_PANEL)
            ax.tick_params(colors="#E8ECEF", labelsize=8)
            ax.title.set_color("#E8ECEF")
            ax.xaxis.label.set_color("#E8ECEF")
            ax.yaxis.label.set_color("#E8ECEF")
            for spine in ax.spines.values():
                spine.set_color(WARNA_BNI_GARIS)
            legenda = ax.get_legend()
            if legenda:
                legenda.get_frame().set_facecolor(WARNA_BNI_PANEL)
                for teks in legenda.get_texts():
                    teks.set_color("#E8ECEF")

    def create_widgets(self):
        main_area = ttk.Frame(self)
        main_area.pack(side=TOP, fill=BOTH, expand=YES)

        # Sidebar
        sidebar = ttk.Frame(main_area, bootstyle="light", width=280, padding=(18, 20))
        sidebar.pack(side=LEFT, fill=Y)

        ttk.Label(
            sidebar, text="PENGATURAN SURAT", font=("Helvetica", 12, "bold"),
            bootstyle="primary",
        ).pack(anchor=W, pady=(0, 16))

        grup_surat = ttk.Labelframe(sidebar, text="Data Surat", padding=14, bootstyle="secondary")
        grup_surat.pack(fill=X, pady=(0, 14))

        ttk.Label(grup_surat, text="Nomor Surat", font=("Helvetica", 9), bootstyle="secondary").pack(anchor=W)
        self.ent_no_surat = ttk.Entry(grup_surat)
        self.ent_no_surat.pack(fill=X, pady=(3, 10))

        ttk.Label(grup_surat, text="Nama Manager", font=("Helvetica", 9), bootstyle="secondary").pack(anchor=W)
        self.ent_manager = ttk.Entry(grup_surat)
        self.ent_manager.insert(0, "Hasbiallah")
        self.ent_manager.pack(fill=X, pady=(3, 10))

        ttk.Label(grup_surat, text="Jenis Surat", font=("Helvetica", 9), bootstyle="secondary").pack(anchor=W)
        self.cmb_jenis_surat = ttk.Combobox(
            grup_surat, values=list(suratgen.TEMPLATE_SURAT.keys()), state="readonly",
        )
        self.cmb_jenis_surat.current(0)
        self.cmb_jenis_surat.pack(fill=X, pady=(3, 0))

        grup_ambang = ttk.Labelframe(sidebar, text="Ambang Batas Peringatan", padding=14, bootstyle="secondary")
        grup_ambang.pack(fill=X, pady=(0, 16))

        ttk.Label(grup_ambang, text="Over-Limit (% dari Pagu)", font=("Helvetica", 9), bootstyle="secondary").pack(anchor=W)
        self.ent_ambang_batas = ttk.Entry(grup_ambang)
        self.ent_ambang_batas.insert(0, "20")
        self.ent_ambang_batas.pack(fill=X, pady=(3, 0))
        self.ent_ambang_batas.bind("<Return>", lambda e: self._refresh_tree_preview())
        self.ent_ambang_batas.bind("<FocusOut>", lambda e: self._refresh_tree_preview())

        ttk.Separator(sidebar).pack(fill=X, pady=(0, 16))

        self.btn_upload = ttk.Button(sidebar, text="Upload Excel", bootstyle="info", command=self.load_excel)
        self.btn_upload.pack(fill=X, pady=(0, 10))

        self.btn_save = ttk.Button(sidebar, text="Simpan ke Database", bootstyle="success", command=self.save_data)
        self.btn_save.pack(fill=X, pady=(0, 10))

        self.btn_cetak = ttk.Button(sidebar, text="Cetak Surat (PDF)", bootstyle="primary", command=self.cetak_surat)
        self.btn_cetak.pack(fill=X)

        # Main Area (Tabs)
        self.notebook = ttk.Notebook(main_area, bootstyle="secondary")
        self.notebook.pack(side=LEFT, fill=BOTH, expand=YES, padx=(14, 16), pady=16)

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
        self.tree.tag_configure("kritis", background="#5A1F1F", foreground="#FF6B6B")
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

        # Tab 5: Supply & Remise — pencatatan pengisian (supply) dan penarikan (remise) kas outlet
        self.tab_supply_remise = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_supply_remise, text="Supply & Remise")
        self.build_tab_supply_remise()

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
        self._gaya_chart_gelap(self.fig_pred, self.ax_ma3, self.ax_lstm)
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
        self._gaya_chart_gelap(self.fig_tren, self.ax_tren)
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
        self._gaya_chart_gelap(self.fig_tren, self.ax_tren)
        self.fig_tren.tight_layout()
        self.canvas_tren.draw()

    def build_tab_supply_remise(self):
        """Membangun tab Supply & Remise: form input pencatatan pengisian (supply)
        kas dari KC ke outlet dan penarikan kelebihan kas (remise) dari outlet ke
        KC, beserta tabel riwayat yang bisa difilter per outlet."""
        form = ttk.Labelframe(self.tab_supply_remise, text="Input Transaksi Supply / Remise", padding=12, bootstyle="secondary")
        form.pack(fill=X, pady=(0, 12))

        baris1 = ttk.Frame(form)
        baris1.pack(fill=X, pady=(0, 8))

        ttk.Label(baris1, text="Tanggal (YYYY-MM-DD):").pack(side=LEFT, padx=(0, 5))
        self.ent_sr_tanggal = ttk.Entry(baris1, width=14)
        self.ent_sr_tanggal.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.ent_sr_tanggal.pack(side=LEFT, padx=(0, 15))

        ttk.Label(baris1, text="Outlet:").pack(side=LEFT, padx=(0, 5))
        self.ent_sr_outlet = ttk.Entry(baris1, width=20)
        self.ent_sr_outlet.pack(side=LEFT, padx=(0, 15))

        ttk.Label(baris1, text="Jenis:").pack(side=LEFT, padx=(0, 5))
        self.cmb_sr_jenis = ttk.Combobox(baris1, values=["Supply", "Remise"], state="readonly", width=10)
        self.cmb_sr_jenis.current(0)
        self.cmb_sr_jenis.pack(side=LEFT)

        baris2 = ttk.Frame(form)
        baris2.pack(fill=X, pady=(0, 8))

        ttk.Label(baris2, text="Mata Uang:").pack(side=LEFT, padx=(0, 5))
        self.cmb_sr_mata_uang = ttk.Combobox(baris2, values=["IDR", "USD"], state="readonly", width=8)
        self.cmb_sr_mata_uang.current(0)
        self.cmb_sr_mata_uang.pack(side=LEFT, padx=(0, 15))

        ttk.Label(baris2, text="Nominal:").pack(side=LEFT, padx=(0, 5))
        self.ent_sr_nominal = ttk.Entry(baris2, width=20)
        self.ent_sr_nominal.pack(side=LEFT)

        baris3 = ttk.Frame(form)
        baris3.pack(fill=X)

        ttk.Label(baris3, text="Keterangan:").pack(side=LEFT, padx=(0, 5))
        self.ent_sr_keterangan = ttk.Entry(baris3, width=40)
        self.ent_sr_keterangan.pack(side=LEFT, padx=(0, 15), fill=X, expand=YES)

        ttk.Button(baris3, text="Simpan Transaksi", bootstyle="success",
                   command=self.simpan_supply_remise).pack(side=LEFT)

        top = ttk.Frame(self.tab_supply_remise)
        top.pack(fill=X, pady=(0, 8))

        ttk.Label(top, text="Filter Outlet:").pack(side=LEFT, padx=(0, 5))
        self.cmb_filter_outlet_sr = ttk.Combobox(top, state="readonly", width=25)
        self.cmb_filter_outlet_sr.pack(side=LEFT, padx=(0, 10))

        ttk.Button(top, text="Muat Riwayat", bootstyle="info",
                   command=self.muat_riwayat_supply_remise).pack(side=LEFT, padx=5)

        self.tree_supply_remise = ttk.Treeview(
            self.tab_supply_remise,
            columns=("tanggal", "outlet", "jenis", "mata_uang", "nominal", "keterangan"),
            show="headings", height=12,
        )
        label_sr = {
            "tanggal": "Tanggal", "outlet": "Outlet", "jenis": "Jenis",
            "mata_uang": "Mata Uang", "nominal": "Nominal", "keterangan": "Keterangan",
        }
        for col in self.tree_supply_remise["columns"]:
            self.tree_supply_remise.heading(col, text=label_sr[col])
            self.tree_supply_remise.column(col, width=130, anchor=CENTER)
        self.tree_supply_remise.tag_configure("remise", background="#3A2A1A", foreground="#F2A104")
        self.tree_supply_remise.tag_configure("supply", background="#1A3A2A", foreground="#2ECC71")
        self.tree_supply_remise.pack(fill=BOTH, expand=YES)

        self.muat_riwayat_supply_remise()

    def _query_supply_remise(self, outlet=None):
        """Mengambil data supply_remise dari database, opsional difilter per outlet."""
        conn = sqlite3.connect(DB_NAME)
        if outlet:
            df = pd.read_sql_query(
                "SELECT * FROM supply_remise WHERE outlet = ? ORDER BY tanggal", conn, params=(outlet,),
            )
        else:
            df = pd.read_sql_query("SELECT * FROM supply_remise ORDER BY tanggal", conn)
        conn.close()
        return df

    def simpan_supply_remise(self):
        """Menyimpan satu transaksi supply/remise kas outlet dari form ke database."""
        outlet = self.ent_sr_outlet.get().strip()
        if not outlet:
            messagebox.showwarning("Peringatan", "Isi nama outlet!")
            return
        nominal = self.bersihkan_angka(self.ent_sr_nominal.get())
        if nominal <= 0:
            messagebox.showwarning("Peringatan", "Isi nominal yang valid (lebih dari 0)!")
            return

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute(
            "INSERT INTO supply_remise (tanggal, outlet, jenis, mata_uang, nominal, keterangan) VALUES (?, ?, ?, ?, ?, ?)",
            (
                self.ent_sr_tanggal.get().strip(), outlet, self.cmb_sr_jenis.get(),
                self.cmb_sr_mata_uang.get(), nominal, self.ent_sr_keterangan.get().strip(),
            ),
        )
        conn.commit()
        conn.close()

        self.ent_sr_outlet.delete(0, END)
        self.ent_sr_nominal.delete(0, END)
        self.ent_sr_keterangan.delete(0, END)

        self._set_status(f"Transaksi {self.cmb_sr_jenis.get().lower()} kas outlet {outlet} tersimpan.")
        messagebox.showinfo("Sukses", "Transaksi supply/remise berhasil disimpan!")
        self.muat_riwayat_supply_remise()

    def muat_riwayat_supply_remise(self):
        """Memuat ulang tabel riwayat supply/remise sesuai filter outlet yang dipilih,
        dan memperbarui daftar pilihan outlet pada combobox filter."""
        df_semua = self._query_supply_remise()
        daftar_outlet = sorted(df_semua["outlet"].unique().tolist()) if not df_semua.empty else []
        self.cmb_filter_outlet_sr["values"] = ["(Semua Outlet)"] + daftar_outlet
        if not self.cmb_filter_outlet_sr.get():
            self.cmb_filter_outlet_sr.current(0)

        pilihan = self.cmb_filter_outlet_sr.get()
        outlet = None if pilihan in ("", "(Semua Outlet)") else pilihan
        df = self._query_supply_remise(outlet)

        for i in self.tree_supply_remise.get_children():
            self.tree_supply_remise.delete(i)
        for _, row in df.iterrows():
            tag = "supply" if row["jenis"] == "Supply" else "remise"
            self.tree_supply_remise.insert("", END, values=(
                row["tanggal"], row["outlet"], row["jenis"], row["mata_uang"],
                f"{row['nominal']:,.0f}", row["keterangan"],
            ), tags=(tag,))

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

            self._gaya_chart_gelap(self.fig_pred, self.ax_ma3, self.ax_lstm)
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

            self._set_status(f"{len(self.df_current)} data over-limit termuat dari {os.path.basename(path)}.")
            messagebox.showinfo("Sukses", f"Berhasil memuat {len(self.df_current)} data.")
        except Exception as e:
            self._set_status("Gagal memuat file Excel.")
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
        self._set_status(f"Data tersimpan ke database dengan nomor surat {no_surat}.")
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
            self._set_status(f"Surat PDF berhasil dibuat: {os.path.basename(output_path)}")
            messagebox.showinfo("Sukses", f"Surat berhasil dibuat:\n{output_path}")
        except Exception as e:
            self._set_status("Gagal membuat surat PDF.")
            messagebox.showerror("Error", f"Gagal membuat surat PDF: {e}")

if __name__ == "__main__":
    app = AppBNI()
    app.mainloop()
