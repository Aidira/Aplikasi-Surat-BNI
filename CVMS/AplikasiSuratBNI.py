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
