import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
import pandas as pd
import sqlite3
import os
from datetime import datetime
from PIL import Image, ImageTk

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


class StatCard(ttk.Frame):
    def __init__(self, parent, title, value="0", color="primary", **kwargs):
        super().__init__(parent, bootstyle=color, padding=(16, 12), **kwargs)
        ttk.Label(self, text=title, font=("Helvetica", 9), bootstyle=f"inverse-{color}").pack(anchor=W)
        self.val_label = ttk.Label(self, text=value, font=("Helvetica", 20, "bold"), bootstyle=f"inverse-{color}")
        self.val_label.pack(anchor=W, pady=(2, 0))

    def update(self, value):
        self.val_label.config(text=value)


class AppBNI(ttk.Window):
    def __init__(self):
        super().__init__(themename="cosmo", title="BNI Asuransi — Over Limit Dashboard")
        self.geometry("1280x760")
        self.minsize(1100, 680)
        init_db()

        self.df_current = pd.DataFrame()
        self._build_ui()

    def _build_ui(self):
        # ── Top navbar ──────────────────────────────────────────────────────
        navbar = ttk.Frame(self, bootstyle="primary", padding=(20, 10))
        navbar.pack(fill=X)

        # Logo
        logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
        if os.path.exists(logo_path):
            img = Image.open(logo_path).resize((120, 39), Image.LANCZOS)
            self._logo = ImageTk.PhotoImage(img)
            ttk.Label(navbar, image=self._logo, bootstyle="inverse-primary").pack(side=LEFT)
        else:
            ttk.Label(navbar, text="BNI", font=("Helvetica", 18, "bold"),
                      bootstyle="inverse-primary").pack(side=LEFT)

        ttk.Label(navbar, text="Over Limit Dashboard",
                  font=("Helvetica", 14), bootstyle="inverse-primary").pack(side=LEFT, padx=16)

        self.lbl_clock = ttk.Label(navbar, text="", font=("Helvetica", 10),
                                   bootstyle="inverse-primary")
        self.lbl_clock.pack(side=RIGHT)
        self._tick()

        # ── Body ────────────────────────────────────────────────────────────
        body = ttk.Frame(self)
        body.pack(fill=BOTH, expand=YES)

        self._build_sidebar(body)
        self._build_main(body)

        # ── Status bar ──────────────────────────────────────────────────────
        statusbar = ttk.Frame(self, bootstyle="secondary", padding=(12, 4))
        statusbar.pack(fill=X, side=BOTTOM)
        self.lbl_status = ttk.Label(statusbar, text="Siap. Silakan upload file Excel.",
                                    font=("Helvetica", 9), bootstyle="inverse-secondary")
        self.lbl_status.pack(side=LEFT)

    def _build_sidebar(self, parent):
        sidebar = ttk.Frame(parent, bootstyle="light", padding=20, width=260)
        sidebar.pack(side=LEFT, fill=Y)
        sidebar.pack_propagate(False)

        # Section: Input
        ttk.Label(sidebar, text="INPUT SURAT", font=("Helvetica", 10, "bold"),
                  bootstyle="secondary").pack(anchor=W, pady=(0, 8))

        ttk.Label(sidebar, text="Nomor Surat", font=("Helvetica", 9)).pack(anchor=W)
        self.ent_no_surat = ttk.Entry(sidebar, font=("Helvetica", 10))
        self.ent_no_surat.pack(fill=X, pady=(2, 12))

        ttk.Label(sidebar, text="Nama Manager", font=("Helvetica", 9)).pack(anchor=W)
        self.ent_manager = ttk.Entry(sidebar, font=("Helvetica", 10))
        self.ent_manager.insert(0, "Hasbiallah")
        self.ent_manager.pack(fill=X, pady=(2, 20))

        ttk.Separator(sidebar).pack(fill=X, pady=4)

        # Section: Actions
        ttk.Label(sidebar, text="AKSI", font=("Helvetica", 10, "bold"),
                  bootstyle="secondary").pack(anchor=W, pady=(12, 8))

        self.btn_upload = ttk.Button(sidebar, text="  Upload File Excel",
                                     bootstyle="info-outline", command=self.load_excel,
                                     width=22)
        self.btn_upload.pack(fill=X, pady=4)

        self.btn_save = ttk.Button(sidebar, text="  Simpan ke Database",
                                   bootstyle="success", command=self.save_data,
                                   width=22)
        self.btn_save.pack(fill=X, pady=4)

        self.btn_riwayat = ttk.Button(sidebar, text="  Lihat Riwayat",
                                      bootstyle="secondary-outline", command=self.load_riwayat,
                                      width=22)
        self.btn_riwayat.pack(fill=X, pady=4)

        ttk.Separator(sidebar).pack(fill=X, pady=16)

        # Stat cards
        ttk.Label(sidebar, text="RINGKASAN", font=("Helvetica", 10, "bold"),
                  bootstyle="secondary").pack(anchor=W, pady=(0, 8))

        self.card_total = StatCard(sidebar, "Total Cabang Over", "0", "danger")
        self.card_total.pack(fill=X, pady=4)

        self.card_idr = StatCard(sidebar, "Over IDR (cabang)", "0", "warning")
        self.card_idr.pack(fill=X, pady=4)

        self.card_usd = StatCard(sidebar, "Over USD (cabang)", "0", "info")
        self.card_usd.pack(fill=X, pady=4)

    def _build_main(self, parent):
        main = ttk.Frame(parent, padding=(16, 16, 16, 8))
        main.pack(side=LEFT, fill=BOTH, expand=YES)

        self.notebook = ttk.Notebook(main, bootstyle="primary")
        self.notebook.pack(fill=BOTH, expand=YES)

        # ── Tab 1: Preview ───────────────────────────────────────────────────
        tab1 = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab1, text="  Preview Data  ")

        # Search bar
        search_frame = ttk.Frame(tab1)
        search_frame.pack(fill=X, pady=(0, 8))
        ttk.Label(search_frame, text="Cari:", font=("Helvetica", 9)).pack(side=LEFT, padx=(0, 4))
        self.ent_search = ttk.Entry(search_frame, width=30)
        self.ent_search.pack(side=LEFT)
        self.ent_search.bind("<KeyRelease>", self._filter_tree)
        ttk.Button(search_frame, text="Reset", bootstyle="secondary-outline",
                   command=self._reset_filter, width=8).pack(side=LEFT, padx=6)

        # Treeview
        cols = ("no", "cabang", "mata_uang", "saldo", "pagu", "over")
        self.tree = ttk.Treeview(tab1, columns=cols, show="headings",
                                 bootstyle="primary", selectmode="browse")

        headers = {"no": ("#", 40), "cabang": ("Cabang", 220), "mata_uang": ("Mata Uang", 90),
                   "saldo": ("Saldo", 150), "pagu": ("Pagu", 150), "over": ("Over Limit", 150)}
        for col, (label, width) in headers.items():
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width,
                             anchor=CENTER if col in ("no", "mata_uang") else E if col != "cabang" else W)

        self.tree.tag_configure("odd", background="#f8f9fa")
        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("usd", foreground="#0d6efd")

        vsb = ttk.Scrollbar(tab1, orient=VERTICAL, command=self.tree.yview, bootstyle="primary-round")
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=RIGHT, fill=Y)
        self.tree.pack(fill=BOTH, expand=YES)

        # ── Tab 2: Riwayat ───────────────────────────────────────────────────
        tab2 = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab2, text="  Riwayat DB  ")

        cols2 = ("id", "tanggal", "no_surat", "cabang", "mata_uang", "saldo", "pagu", "over")
        self.tree2 = ttk.Treeview(tab2, columns=cols2, show="headings",
                                  bootstyle="secondary", selectmode="browse")
        headers2 = {"id": ("ID", 40), "tanggal": ("Tanggal", 100), "no_surat": ("No. Surat", 140),
                    "cabang": ("Cabang", 180), "mata_uang": ("Mata Uang", 80),
                    "saldo": ("Saldo", 130), "pagu": ("Pagu", 130), "over": ("Over", 130)}
        for col, (label, width) in headers2.items():
            self.tree2.heading(col, text=label)
            self.tree2.column(col, width=width,
                              anchor=CENTER if col in ("id", "tanggal", "mata_uang") else E if col not in ("no_surat", "cabang") else W)

        vsb2 = ttk.Scrollbar(tab2, orient=VERTICAL, command=self.tree2.yview, bootstyle="secondary-round")
        self.tree2.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side=RIGHT, fill=Y)
        self.tree2.pack(fill=BOTH, expand=YES)

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _tick(self):
        self.lbl_clock.config(text=datetime.now().strftime("%A, %d %B %Y  |  %H:%M:%S"))
        self.after(1000, self._tick)

    def _fmt(self, val):
        try:
            return f"{float(val):,.2f}"
        except:
            return str(val)

    def _set_status(self, msg):
        self.lbl_status.config(text=msg)

    def _populate_tree(self, df):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, (_, row) in enumerate(df.iterrows()):
            tag = ("usd" if row["Mata Uang"] == "USD" else "") + ("odd" if idx % 2 else "even")
            self.tree.insert("", END, values=(
                idx + 1,
                row["Cabang"],
                row["Mata Uang"],
                self._fmt(row["Saldo"]),
                self._fmt(row["Pagu"]),
                self._fmt(row["Over"]),
            ), tags=(tag,))

    def _filter_tree(self, event=None):
        q = self.ent_search.get().lower()
        if not self.df_current.empty:
            filtered = self.df_current[self.df_current["Cabang"].str.lower().str.contains(q)]
            self._populate_tree(filtered)

    def _reset_filter(self):
        self.ent_search.delete(0, END)
        self._populate_tree(self.df_current)

    def _update_cards(self):
        total = len(self.df_current)
        idr = len(self.df_current[self.df_current["Mata Uang"] == "IDR"]) if not self.df_current.empty else 0
        usd = len(self.df_current[self.df_current["Mata Uang"] == "USD"]) if not self.df_current.empty else 0
        self.card_total.update(str(total))
        self.card_idr.update(str(idr))
        self.card_usd.update(str(usd))

    def bersihkan_angka(self, nilai_raw):
        try:
            if isinstance(nilai_raw, (int, float)):
                return float(nilai_raw)
            text = str(nilai_raw).upper().replace("IDR", "").replace("RP", "").replace(" ", "").strip()
            if "." in text and "," in text:
                text = text.replace(".", "").replace(",", ".")
            elif "." in text:
                text = text.replace(".", "")
            elif "," in text:
                text = text.replace(",", ".")
            return float(text)
        except:
            return 0.0

    # ── Commands ────────────────────────────────────────────────────────────

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return

        try:
            self._set_status("Membaca file Excel...")
            self.update_idletasks()
            xls = pd.ExcelFile(path)
            data_over = []
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                curr = "USD" if "USD" in sheet.upper() else "IDR"
                for _, row in df.iterrows():
                    if len(row) < 4:
                        continue
                    cabang = str(row[1])
                    if pd.isna(cabang) or "TOTAL" in cabang or "NAMA" in cabang.upper() or "KCU" in cabang:
                        continue
                    pagu = self.bersihkan_angka(row[2])
                    saldo = self.bersihkan_angka(row[3])
                    over = saldo - pagu
                    if over > 0:
                        data_over.append([cabang, curr, saldo, pagu, over])

            self.df_current = pd.DataFrame(data_over, columns=["Cabang", "Mata Uang", "Saldo", "Pagu", "Over"])
            self._populate_tree(self.df_current)
            self._update_cards()
            self._set_status(f"Berhasil memuat {len(self.df_current)} data over limit dari '{os.path.basename(path)}'.")
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file:\n{e}")
            self._set_status("Gagal membaca file.")

    def save_data(self):
        if self.df_current.empty:
            messagebox.showwarning("Peringatan", "Tidak ada data untuk disimpan.")
            return
        no_surat = self.ent_no_surat.get().strip()
        if not no_surat:
            messagebox.showwarning("Peringatan", "Isi nomor surat terlebih dahulu.")
            self.ent_no_surat.focus()
            return

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        tgl = datetime.now().strftime("%Y-%m-%d")
        for _, row in self.df_current.iterrows():
            c.execute(
                "INSERT INTO riwayat_over (tanggal_input, no_surat, cabang, mata_uang, saldo, pagu, over_limit) VALUES (?,?,?,?,?,?,?)",
                (tgl, no_surat, row["Cabang"], row["Mata Uang"], row["Saldo"], row["Pagu"], row["Over"])
            )
        conn.commit()
        conn.close()
        self._set_status(f"Data berhasil disimpan — No. Surat: {no_surat}, {len(self.df_current)} cabang.")
        messagebox.showinfo("Sukses", f"{len(self.df_current)} data berhasil disimpan ke database.")

    def load_riwayat(self):
        conn = sqlite3.connect(DB_NAME)
        rows = conn.execute("SELECT id, tanggal_input, no_surat, cabang, mata_uang, saldo, pagu, over_limit FROM riwayat_over ORDER BY id DESC").fetchall()
        conn.close()

        for i in self.tree2.get_children():
            self.tree2.delete(i)
        for idx, row in enumerate(rows):
            tag = "odd" if idx % 2 else "even"
            self.tree2.insert("", END, values=(
                row[0], row[1], row[2], row[3], row[4],
                self._fmt(row[5]), self._fmt(row[6]), self._fmt(row[7])
            ), tags=(tag,))
        self.tree2.tag_configure("odd", background="#f8f9fa")
        self.tree2.tag_configure("even", background="#ffffff")

        self.notebook.select(1)
        self._set_status(f"Riwayat dimuat — {len(rows)} record ditemukan.")


if __name__ == "__main__":
    app = AppBNI()
    app.mainloop()
