import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import sqlite3
import os
from datetime import datetime
from PIL import Image, ImageTk

# --- KONFIGURASI DATABASE ---
DB_NAME = "database_bni.db"

# --- KONFIGURASI FORM AMBIL/BUANG MODAL (mengikuti slip "Uang Tunai dari/ke Kasir") ---
DENOM_KERTAS = [100000, 50000, 20000, 10000, 5000, 2000, 1000, 500, 100]
DENOM_LOGAM = [1000, 500, 200, 100, 50, 25, 10, 5]

JUDUL_MODAL = {
    "AMBIL": "Uang Tunai dari Kasir",
    "BUANG": "Uang Tunai ke Kasir",
}

REKENING_MODAL = {
    "AMBIL": {"tujuan": "KAS TELLER", "asal": "KAS", "diberikan": "Kasir", "diterima": "Teller"},
    "BUANG": {"tujuan": "KAS", "asal": "KAS TELLER", "diberikan": "Teller", "diterima": "Kasir"},
}

def format_rupiah(value):
    try:
        value = int(round(float(value)))
    except (TypeError, ValueError):
        value = 0
    return "Rp {:,.0f}".format(value).replace(",", ".")

def angka_ke_terbilang(n):
    satuan = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan"]

    def eja(n):
        if n < 10:
            return satuan[n]
        if n < 20:
            return "sebelas" if n == 11 else satuan[n - 10] + " belas"
        if n < 100:
            sisa = eja(n % 10) if n % 10 else ""
            return (satuan[n // 10] + " puluh " + sisa).strip()
        if n < 200:
            sisa = eja(n - 100) if n > 100 else ""
            return ("seratus " + sisa).strip()
        if n < 1000:
            sisa = eja(n % 100) if n % 100 else ""
            return (satuan[n // 100] + " ratus " + sisa).strip()
        if n < 2000:
            sisa = eja(n - 1000) if n > 1000 else ""
            return ("seribu " + sisa).strip()
        if n < 1000000:
            sisa = eja(n % 1000) if n % 1000 else ""
            return (eja(n // 1000) + " ribu " + sisa).strip()
        if n < 1000000000:
            sisa = eja(n % 1000000) if n % 1000000 else ""
            return (eja(n // 1000000) + " juta " + sisa).strip()
        if n < 1000000000000:
            sisa = eja(n % 1000000000) if n % 1000000000 else ""
            return (eja(n // 1000000000) + " miliar " + sisa).strip()
        sisa = eja(n % 1000000000000) if n % 1000000000000 else ""
        return (eja(n // 1000000000000) + " triliun " + sisa).strip()

    n = int(round(float(n))) if n else 0
    if n <= 0:
        return "nol rupiah"
    hasil = eja(n)
    return hasil[0].upper() + hasil[1:] + " rupiah"

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
        CREATE TABLE IF NOT EXISTS riwayat_modal (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            jenis TEXT,
            tanggal TEXT,
            jam TEXT,
            no_rekening_asal TEXT,
            nama_rekening_asal TEXT,
            no_rekening_tujuan TEXT,
            nama_rekening_tujuan TEXT,
            terbilang TEXT,
            rincian_kertas REAL,
            rincian_logam REAL,
            rincian_valas REAL,
            total REAL,
            diberikan_oleh TEXT,
            diterima_oleh TEXT,
            dibuat_pada TEXT
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
        self.modal_forms = {}
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

        self.btn_upload = ttk.Button(sidebar, text="Upload Excel", bootstyle="info", command=self.load_excel)
        self.btn_upload.pack(fill=X, pady=20)

        self.btn_save = ttk.Button(sidebar, text="Simpan ke DB", bootstyle="success", command=self.save_data)
        self.btn_save.pack(fill=X, pady=5)

        # Main Area (Tabs)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)

        # Tab 1: Data Editor / Preview
        self.tab_preview = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_preview, text="Preview Data")
        
        self.tree = ttk.Treeview(self.tab_preview, columns=("cabang", "mata_uang", "saldo", "pagu", "over"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, width=150, anchor=CENTER)
        self.tree.pack(fill=BOTH, expand=YES)

        # Tab 2: Ambil Modal (Uang Tunai dari Kasir)
        self.tab_ambil = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_ambil, text="Ambil Modal")
        self.build_form_modal(self.make_scrollable(self.tab_ambil), "AMBIL")

        # Tab 3: Buang Modal (Uang Tunai ke Kasir)
        self.tab_buang = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_buang, text="Buang Modal")
        self.build_form_modal(self.make_scrollable(self.tab_buang), "BUANG")

        # Tab 4: Riwayat Modal
        self.tab_riwayat_modal = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_riwayat_modal, text="Riwayat Modal")

        ttk.Button(self.tab_riwayat_modal, text="Refresh", bootstyle="info",
                   command=self.load_riwayat_modal).pack(anchor=W, pady=(0, 5))

        cols_modal = ("jenis", "tanggal", "jam", "rek_asal", "rek_tujuan", "total", "diberikan_oleh", "diterima_oleh")
        self.tree_modal = ttk.Treeview(self.tab_riwayat_modal, columns=cols_modal, show="headings")
        for col in cols_modal:
            self.tree_modal.heading(col, text=col.upper())
            self.tree_modal.column(col, width=130, anchor=CENTER)
        self.tree_modal.pack(fill=BOTH, expand=YES)
        self.load_riwayat_modal()

    def make_scrollable(self, parent):
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas, padding=15)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor=NW)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)
        return inner

    def build_form_modal(self, parent, jenis):
        cfg = REKENING_MODAL[jenis]
        w = {}
        self.modal_forms[jenis] = w

        warna = "danger" if jenis == "AMBIL" else "info"
        ttk.Label(parent, text=JUDUL_MODAL[jenis], font=("Helvetica", 14, "bold"), bootstyle=warna).pack(anchor=W, pady=(0, 10))

        top = ttk.Frame(parent)
        top.pack(fill=X, pady=5)

        ttk.Label(top, text="Tanggal:").grid(row=0, column=0, sticky=W, padx=5, pady=3)
        w["ent_tanggal"] = ttk.Entry(top, width=15)
        w["ent_tanggal"].insert(0, datetime.now().strftime("%Y-%m-%d"))
        w["ent_tanggal"].grid(row=0, column=1, sticky=W, padx=5, pady=3)

        ttk.Label(top, text="Jam:").grid(row=0, column=2, sticky=W, padx=5, pady=3)
        w["ent_jam"] = ttk.Entry(top, width=10)
        w["ent_jam"].insert(0, datetime.now().strftime("%H:%M"))
        w["ent_jam"].grid(row=0, column=3, sticky=W, padx=5, pady=3)

        ttk.Label(top, text="No. Rek. Tujuan:").grid(row=1, column=0, sticky=W, padx=5, pady=3)
        w["ent_rek_tujuan"] = ttk.Entry(top, width=15)
        w["ent_rek_tujuan"].grid(row=1, column=1, sticky=W, padx=5, pady=3)
        ttk.Label(top, text=f"Nama Rekening: {cfg['tujuan']}").grid(row=1, column=2, columnspan=2, sticky=W, padx=5, pady=3)

        ttk.Label(top, text="No. Rek. Asal:").grid(row=2, column=0, sticky=W, padx=5, pady=3)
        w["ent_rek_asal"] = ttk.Entry(top, width=15)
        w["ent_rek_asal"].grid(row=2, column=1, sticky=W, padx=5, pady=3)
        ttk.Label(top, text=f"Nama Rekening: {cfg['asal']}").grid(row=2, column=2, columnspan=2, sticky=W, padx=5, pady=3)

        rincian = ttk.Labelframe(parent, text="Perincian (terdiri dari)", padding=10)
        rincian.pack(fill=X, pady=10)

        kertas_frame = ttk.Frame(rincian)
        kertas_frame.grid(row=0, column=0, sticky=N, padx=20)
        ttk.Label(kertas_frame, text="Uang Kertas, Rupiah", font=("Helvetica", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky=W, pady=(0, 5))
        w["kertas_entries"] = []
        for i, denom in enumerate(DENOM_KERTAS, start=1):
            ttk.Label(kertas_frame, text="x " + "{:,}".format(denom).replace(",", ".")).grid(row=i, column=0, sticky=W)
            e = ttk.Entry(kertas_frame, width=6)
            e.grid(row=i, column=1, padx=5)
            e.bind("<KeyRelease>", lambda ev, j=jenis: self.recalc_modal(j))
            lbl = ttk.Label(kertas_frame, text="Rp 0", width=15, anchor=E)
            lbl.grid(row=i, column=2, sticky=E)
            w["kertas_entries"].append((denom, e, lbl))
        w["lbl_sub_kertas"] = ttk.Label(kertas_frame, text="Sub Total: Rp 0", font=("Helvetica", 9, "bold"))
        w["lbl_sub_kertas"].grid(row=len(DENOM_KERTAS) + 1, column=0, columnspan=3, sticky=W, pady=(5, 0))

        logam_frame = ttk.Frame(rincian)
        logam_frame.grid(row=0, column=1, sticky=N, padx=20)
        ttk.Label(logam_frame, text="Uang Logam, Rupiah", font=("Helvetica", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky=W, pady=(0, 5))
        w["logam_entries"] = []
        for i, denom in enumerate(DENOM_LOGAM, start=1):
            ttk.Label(logam_frame, text="x " + "{:,}".format(denom).replace(",", ".")).grid(row=i, column=0, sticky=W)
            e = ttk.Entry(logam_frame, width=6)
            e.grid(row=i, column=1, padx=5)
            e.bind("<KeyRelease>", lambda ev, j=jenis: self.recalc_modal(j))
            lbl = ttk.Label(logam_frame, text="Rp 0", width=15, anchor=E)
            lbl.grid(row=i, column=2, sticky=E)
            w["logam_entries"].append((denom, e, lbl))
        w["lbl_sub_logam"] = ttk.Label(logam_frame, text="Sub Total: Rp 0", font=("Helvetica", 9, "bold"))
        w["lbl_sub_logam"].grid(row=len(DENOM_LOGAM) + 1, column=0, columnspan=3, sticky=W, pady=(5, 0))

        valas_frame = ttk.Frame(rincian)
        valas_frame.grid(row=1, column=0, columnspan=2, sticky=W, pady=10)
        ttk.Label(valas_frame, text="Uang Kertas Valas (Sub Total):").pack(side=LEFT, padx=(0, 5))
        w["ent_valas"] = ttk.Entry(valas_frame, width=15)
        w["ent_valas"].insert(0, "0")
        w["ent_valas"].pack(side=LEFT)
        w["ent_valas"].bind("<KeyRelease>", lambda ev, j=jenis: self.recalc_modal(j))

        w["lbl_total"] = ttk.Label(parent, text="TOTAL: Rp 0", font=("Helvetica", 12, "bold"), bootstyle=warna)
        w["lbl_total"].pack(anchor=E, pady=5)

        bottom = ttk.Frame(parent)
        bottom.pack(fill=X, pady=5)

        ttk.Label(bottom, text="Jumlah:").grid(row=0, column=0, sticky=W, padx=5, pady=3)
        w["ent_jumlah"] = ttk.Entry(bottom, width=20, state="readonly")
        w["ent_jumlah"].grid(row=0, column=1, sticky=W, padx=5, pady=3)

        ttk.Label(bottom, text="Terbilang:").grid(row=1, column=0, sticky=W, padx=5, pady=3)
        w["ent_terbilang"] = ttk.Entry(bottom, width=70, state="readonly")
        w["ent_terbilang"].grid(row=1, column=1, columnspan=3, sticky=W, padx=5, pady=3)

        ttk.Label(bottom, text=f"Diberikan oleh ({cfg['diberikan']}):").grid(row=2, column=0, sticky=W, padx=5, pady=3)
        w["ent_diberikan"] = ttk.Entry(bottom, width=25)
        w["ent_diberikan"].grid(row=2, column=1, sticky=W, padx=5, pady=3)

        ttk.Label(bottom, text=f"Diterima oleh ({cfg['diterima']}):").grid(row=2, column=2, sticky=W, padx=5, pady=3)
        w["ent_diterima"] = ttk.Entry(bottom, width=25)
        w["ent_diterima"].grid(row=2, column=3, sticky=W, padx=5, pady=3)

        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=X, pady=10)
        ttk.Button(btn_frame, text="Simpan Transaksi", bootstyle="success",
                   command=lambda j=jenis: self.save_modal(j)).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Reset Form", bootstyle="secondary",
                   command=lambda j=jenis: self.reset_modal(j)).pack(side=LEFT, padx=5)

        w["nama_tujuan"] = cfg["tujuan"]
        w["nama_asal"] = cfg["asal"]

        self.recalc_modal(jenis)

    def _safe_int(self, val):
        try:
            return int(val)
        except (TypeError, ValueError):
            return 0

    def _safe_float(self, val):
        try:
            return float(val)
        except (TypeError, ValueError):
            return 0.0

    def recalc_modal(self, jenis):
        w = self.modal_forms[jenis]

        sub_kertas = 0
        for denom, entry, lbl in w["kertas_entries"]:
            amount = self._safe_int(entry.get()) * denom
            lbl.config(text=format_rupiah(amount))
            sub_kertas += amount

        sub_logam = 0
        for denom, entry, lbl in w["logam_entries"]:
            amount = self._safe_int(entry.get()) * denom
            lbl.config(text=format_rupiah(amount))
            sub_logam += amount

        valas = self._safe_float(w["ent_valas"].get())
        total = sub_kertas + sub_logam + valas

        w["lbl_sub_kertas"].config(text=f"Sub Total: {format_rupiah(sub_kertas)}")
        w["lbl_sub_logam"].config(text=f"Sub Total: {format_rupiah(sub_logam)}")
        w["lbl_total"].config(text=f"TOTAL: {format_rupiah(total)}")

        for key, value in (("ent_jumlah", format_rupiah(total)), ("ent_terbilang", angka_ke_terbilang(total))):
            w[key].config(state="normal")
            w[key].delete(0, END)
            w[key].insert(0, value)
            w[key].config(state="readonly")

        return sub_kertas, sub_logam, valas, total

    def save_modal(self, jenis):
        w = self.modal_forms[jenis]
        sub_kertas, sub_logam, valas, total = self.recalc_modal(jenis)

        if total <= 0:
            messagebox.showwarning("Peringatan", "Jumlah transaksi masih kosong/nol!")
            return
        if not w["ent_diberikan"].get().strip() or not w["ent_diterima"].get().strip():
            messagebox.showwarning("Peringatan", "Lengkapi nama yang menyerahkan dan menerima!")
            return

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''INSERT INTO riwayat_modal
                     (jenis, tanggal, jam, no_rekening_asal, nama_rekening_asal,
                      no_rekening_tujuan, nama_rekening_tujuan, terbilang,
                      rincian_kertas, rincian_logam, rincian_valas, total,
                      diberikan_oleh, diterima_oleh, dibuat_pada)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (jenis, w["ent_tanggal"].get(), w["ent_jam"].get(),
                   w["ent_rek_asal"].get(), w["nama_asal"],
                   w["ent_rek_tujuan"].get(), w["nama_tujuan"],
                   w["ent_terbilang"].get(),
                   sub_kertas, sub_logam, valas, total,
                   w["ent_diberikan"].get(), w["ent_diterima"].get(),
                   datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()

        messagebox.showinfo("Sukses", f"Transaksi {JUDUL_MODAL[jenis]} berhasil disimpan!")
        self.reset_modal(jenis)
        self.load_riwayat_modal()

    def reset_modal(self, jenis):
        w = self.modal_forms[jenis]
        for _, entry, _ in w["kertas_entries"] + w["logam_entries"]:
            entry.delete(0, END)
        w["ent_valas"].delete(0, END)
        w["ent_valas"].insert(0, "0")
        w["ent_rek_asal"].delete(0, END)
        w["ent_rek_tujuan"].delete(0, END)
        w["ent_diberikan"].delete(0, END)
        w["ent_diterima"].delete(0, END)
        w["ent_tanggal"].delete(0, END)
        w["ent_tanggal"].insert(0, datetime.now().strftime("%Y-%m-%d"))
        w["ent_jam"].delete(0, END)
        w["ent_jam"].insert(0, datetime.now().strftime("%H:%M"))
        self.recalc_modal(jenis)

    def load_riwayat_modal(self):
        for i in self.tree_modal.get_children():
            self.tree_modal.delete(i)
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''SELECT jenis, tanggal, jam, no_rekening_asal, no_rekening_tujuan,
                            total, diberikan_oleh, diterima_oleh
                     FROM riwayat_modal ORDER BY id DESC''')
        rows = c.fetchall()
        conn.close()
        for jenis, tanggal, jam, rek_asal, rek_tujuan, total, diberikan, diterima in rows:
            self.tree_modal.insert("", END, values=(jenis, tanggal, jam, rek_asal, rek_tujuan,
                                                      format_rupiah(total), diberikan, diterima))

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
            
            # Refresh Treeview
            for i in self.tree.get_children(): self.tree.delete(i)
            for _, row in self.df_current.iterrows():
                self.tree.insert("", END, values=list(row))
            
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
        messagebox.showinfo("Sukses", "Data berhasil disimpan ke database!")

if __name__ == "__main__":
    app = AppBNI()
    app.mainloop()
