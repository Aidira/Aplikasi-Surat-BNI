import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import webbrowser
import pathlib
import platform
import csv
import json
from datetime import datetime

# --- KONFIGURASI FILE ---
CONFIG_FILE = "config_bni.json"
HISTORY_FILE = "riwayat_cetak.csv"

# --- TEMPLATE SURAT (HTML) DENGAN SUMMARY ---
TEMPLATE_HTML = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <title>Cetak Surat BNI</title>
    <style>
        body {{ font-family: "Times New Roman", Times, serif; font-size: 12pt; color: #000; padding: 40px; }}
        table.surat {{ width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 11pt; }}
        table.surat, table.surat th, table.surat td {{ border: 1px solid black; }}
        table.surat th {{ text-align: center; padding: 5px; font-weight: bold; background-color: #fff; }}
        table.surat td {{ padding: 4px 6px; }}
        table.surat td.kanan {{ text-align: right; }}
        table.surat td.tengah {{ text-align: center; }}
        .bold {{ font-weight: bold; }}
        .summary-box {{ margin-top: 20px; border: 1px solid #000; padding: 10px; width: 60%; font-size: 11pt; }}
        .signature {{ margin-top: 40px; page-break-inside: avoid; }}
        .ttd-space {{ height: 80px; }}
        @media print {{ .no-print {{ display: none !important; }} }}
    </style>
</head>
<body>
    <div class="no-print" style="text-align:center; margin-bottom:20px; padding:10px; background:#f0f0f0; border:1px solid #ccc;">
        <button onclick="window.print()" style="font-size:16px; padding:10px 20px; cursor:pointer; font-weight:bold;">🖨️ KLIK DISINI UNTUK PRINT / SAVE PDF</button>
    </div>

    <p>Jakarta, {tgl_surat}</p>
    <p>No. Surat : TEB/3.2/{no_surat}</p>
    <br>
    
    <p class="bold">Kepada<br>PT. Asuransi TRI PAKARTA<br>
    <span style="font-weight:normal">Kantor Cabang Jakarta Selatan<br>
    Komplek Sentra Arteri Mas<br>Jl. Sultan Iskandar Muda No. 10B<br>Jaksel 12240</span></p>

    <p class="bold">UP. Ibu Siska (Fax. 021-7293312 / 75917755 / 7394748)</p>
    <p class="bold">Hal : Cover Asuransi CIS Saldo Kas IDR dan Valas KC/KCP/KK</p>

    <p style="text-align:justify; line-height:1.5;">Menunjuk perihal pokok surat tersebut diatas, dengan ini kami sampaikan adanya kelebihan pagu kas (over limit) IDR dan Valas di lingkungan BNI KC Tebet, dengan perincian sbb :</p>

    <table class="surat">
        <thead>
            <tr>
                <th width="5%">NO</th>
                <th>KCU/KCP/KK</th>
                <th>Saldo</th>
                <th>Open (Pagu)</th>
                <th>Over (Selisih)</th>
            </tr>
        </thead>
        <tbody>
            {rows}
        </tbody>
    </table>

    <div class="summary-box">
        <b>RINGKASAN TOTAL OVER LIMIT:</b><br>
        Total IDR : {total_idr}<br>
        Total Valas : {total_valas}
    </div>

    <div id="infoText">
        <p style="text-align:justify; line-height:1.5;">Saldo tersebut telah melebihi cover asuransi cash in save pada open cover Saudara, dengan ini kami laporkan via faksimili/email, agar kelebihan saldo tersebut dapat Saudara tutup dengan asuransi Cash In Save.</p>
    </div>

    <p>Demikianlah untuk dimaklumi, atas perhatian dan kerjasama Saudara kami ucapkan terima kasih.</p>

    <div class="signature">
        <p>PT. Bank Negara Indonesia (Persero) Tbk<br>Kantor Cabang Tebet</p>
        <div class="ttd-space"></div>
        <p><strong><u>{nama_manager}</u></strong><br>{jabatan_manager}</p>
    </div>
</body>
</html>
"""

class AplikasiSurat(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Aplikasi Cetak Surat BNI (Ultimate)")
        self.geometry("600x750") # Diperbesar untuk tabel preview
        self.resizable(True, True)
        
        # --- SETTING WARNA (Fix Mac Dark Mode) ---
        self.bg_color = "#f4f4f4"
        self.fg_color = "black"
        self.input_bg = "white"
        
        self.configure(bg=self.bg_color)
        self.file_path = ""
        self.data_over_limit = []
        
        # Load Config (Nama Manager Terakhir)
        self.default_manager = "Hasbiallah"
        self.load_config()

        self.create_widgets()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    self.default_manager = data.get("manager", "Hasbiallah")
            except:
                pass

    def save_config(self):
        data = {"manager": self.entry_nama.get()}
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(data, f)
        except:
            pass

    def log_history(self, total_cabang, total_idr, total_valas):
        # Simpan log ke CSV
        file_exists = os.path.isfile(HISTORY_FILE)
        try:
            with open(HISTORY_FILE, mode='a', newline='') as file:
                writer = csv.writer(file)
                if not file_exists:
                    writer.writerow(["Tanggal", "Jam", "No Surat", "Manager", "Jml Cabang", "Total IDR", "Total Valas"])
                
                now = datetime.now()
                writer.writerow([
                    now.strftime("%Y-%m-%d"),
                    now.strftime("%H:%M:%S"),
                    self.entry_nomor.get(),
                    self.entry_nama.get(),
                    total_cabang,
                    total_idr,
                    total_valas
                ])
        except Exception as e:
            print(f"Gagal simpan log: {e}")

    def create_widgets(self):
        # Header
        lbl_judul = tk.Label(self, text="GENERATOR SURAT ASURANSI", 
                           font=("Arial", 16, "bold"), bg=self.bg_color, fg="#005E6A")
        lbl_judul.pack(pady=15)

        frame = tk.Frame(self, bg=self.bg_color)
        frame.pack(pady=5, padx=20, fill="x")

        # --- 1. PILIH FILE ---
        tk.Label(frame, text="1. File Excel Laporan (.xlsx / .xls):", 
                 bg=self.bg_color, fg=self.fg_color, anchor="w", font=("Arial", 10, "bold")).pack(fill="x")
        
        self.btn_file = tk.Button(frame, text="Pilih File Excel", command=self.pilih_file, 
                                  bg="white", fg="black") 
        self.btn_file.pack(fill="x", pady=2)
        
        self.lbl_file_status = tk.Label(frame, text="Belum ada file dipilih", 
                                        bg=self.bg_color, fg="red", font=("Arial", 9))
        self.lbl_file_status.pack(pady=(0, 10))

        # --- 2. NOMOR SURAT ---
        tk.Label(frame, text="2. Nomor Surat (4 digit terakhir):", 
                 bg=self.bg_color, fg=self.fg_color, anchor="w", font=("Arial", 10, "bold")).pack(fill="x")
        
        frame_no = tk.Frame(frame, bg=self.bg_color)
        frame_no.pack(fill="x", pady=2)
        
        tk.Label(frame_no, text="TEB/3.2/", bg="#e0e0e0", fg="black", padx=5, relief="solid").pack(side="left", fill="y")
        self.entry_nomor = tk.Entry(frame_no, bg=self.input_bg, fg=self.fg_color, insertbackground="black")
        self.entry_nomor.pack(side="left", fill="x", expand=True, padx=(5,0))

        # --- 3. PENANDA TANGAN ---
        tk.Label(frame, text="3. Penanda Tangan:", 
                 bg=self.bg_color, fg=self.fg_color, anchor="w", font=("Arial", 10, "bold")).pack(fill="x", pady=(10, 0))
        
        self.entry_nama = tk.Entry(frame, bg=self.input_bg, fg=self.fg_color, insertbackground="black")
        self.entry_nama.insert(0, self.default_manager)
        self.entry_nama.pack(fill="x", pady=2)

        self.var_pgs = tk.BooleanVar()
        self.chk_pgs = tk.Checkbutton(frame, text="Pgs. (Pejabat Pengganti)", variable=self.var_pgs, 
                                      bg=self.bg_color, fg=self.fg_color, anchor="w")
        self.chk_pgs.pack(fill="x")

        # --- TOMBOL ANALISA ---
        tk.Button(frame, text="🔍 CEK DATA DULU (PREVIEW)", command=self.analisa_data, 
                  bg="#FFC107", fg="black", font=("Arial", 10, "bold")).pack(fill="x", pady=15)

        # --- TABEL PREVIEW ---
        tk.Label(self, text="Preview Data Over Limit:", bg=self.bg_color, fg=self.fg_color).pack(padx=20, anchor="w")
        
        columns = ('cabang', 'saldo', 'pagu', 'over')
        self.tree = ttk.Treeview(self, columns=columns, show='headings', height=8)
        self.tree.heading('cabang', text='Cabang')
        self.tree.heading('saldo', text='Saldo')
        self.tree.heading('pagu', text='Pagu')
        self.tree.heading('over', text='Over (Selisih)')
        
        self.tree.column('cabang', width=150)
        self.tree.column('saldo', width=120, anchor='e')
        self.tree.column('pagu', width=120, anchor='e')
        self.tree.column('over', width=120, anchor='e')
        
        self.tree.pack(padx=20, fill="both", expand=True)

        # --- TOMBOL CETAK ---
        btn_proses = tk.Button(self, text="🖨️ CETAK SURAT SEKARANG", command=self.cetak_surat, 
                               bg="#005E6A", fg="white", font=("Arial", 12, "bold"), height=2)
        if platform.system() == "Darwin": 
             btn_proses.config(fg="black", bg="#e0e0e0")
        btn_proses.pack(fill="x", padx=20, pady=20)

    def pilih_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.file_path = file
            self.lbl_file_status.config(text=f"File: {os.path.basename(file)}", fg="#008000")
            # Bersihkan tabel jika ganti file
            for i in self.tree.get_children():
                self.tree.delete(i)

    def format_uang(self, nilai, currency="IDR"):
        try:
            return "{} {:,.0f}".format(currency, float(nilai)).replace(",", ".")
        except:
            return "0"

    def bersihkan_angka(self, nilai_raw):
        try:
            if isinstance(nilai_raw, (int, float)):
                return float(nilai_raw)
            text = str(nilai_raw).upper().replace("IDR", "").replace("RP", "").replace(" ", "").strip()
            if "." in text and "," in text: text = text.replace(".", "").replace(",", ".")
            elif "." in text: text = text.replace(".", "")
            elif "," in text: text = text.replace(",", ".")
            return float(text)
        except:
            return 0.0

    def analisa_data(self):
        if not self.file_path:
            messagebox.showwarning("Peringatan", "Pilih file Excel dulu!")
            return
        
        try:
            # Otomatis deteksi engine buat .xls atau .xlsx
            xls = pd.ExcelFile(self.file_path)
            self.data_over_limit = []
            
            # Reset Tabel
            for i in self.tree.get_children():
                self.tree.delete(i)

            count_over = 0

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                curr = "USD" if "USD" in sheet_name.upper() else "IDR"

                for index, row in df.iterrows():
                    if len(row) < 4: continue
                    cabang = str(row[1])
                    if pd.isna(cabang) or "KCU" in cabang or "TOTAL" in cabang or "NAMA" in cabang.upper() or cabang.strip() == "":
                        continue

                    pagu = self.bersihkan_angka(row[2])
                    saldo = self.bersihkan_angka(row[3])

                    if pagu > 0 and saldo > pagu:
                        selisih = saldo - pagu
                        self.data_over_limit.append({
                            "cabang": cabang,
                            "saldo_fmt": self.format_uang(saldo, curr),
                            "pagu_fmt": self.format_uang(pagu, curr),
                            "over_fmt": self.format_uang(selisih, curr),
                            "raw_over": selisih,
                            "currency": curr
                        })
                        # Masukkan ke Tabel Preview
                        self.tree.insert('', 'end', values=(
                            "BNI " + cabang,
                            self.format_uang(saldo, curr),
                            self.format_uang(pagu, curr),
                            self.format_uang(selisih, curr)
                        ))
                        count_over += 1
            
            if count_over == 0:
                messagebox.showinfo("Info", "Aman! Tidak ada cabang yang Over Limit.")
            else:
                messagebox.showinfo("Selesai", f"Ditemukan {count_over} cabang Over Limit.\nSilakan cek tabel preview.")

        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file:\n{str(e)}")

    def cetak_surat(self):
        if not self.data_over_limit:
            messagebox.showwarning("Peringatan", "Data kosong! Klik tombol 'Cek Data' dulu atau pastikan ada data over limit.")
            return

        no_surat_input = self.entry_nomor.get()
        if not no_surat_input:
            messagebox.showwarning("Peringatan", "Nomor Surat wajib diisi!")
            return

        # Simpan Config Manager
        self.save_config()

        jabatan = "Branch Service Manager"
        if self.var_pgs.get():
            jabatan = "Pgs. Branch Service Manager"

        # Hitung Total Summary
        sum_idr = sum(d['raw_over'] for d in self.data_over_limit if d['currency'] == 'IDR')
        sum_usd = sum(d['raw_over'] for d in self.data_over_limit if d['currency'] == 'USD')
        
        str_sum_idr = self.format_uang(sum_idr, "IDR") if sum_idr > 0 else "-"
        str_sum_usd = self.format_uang(sum_usd, "USD") if sum_usd > 0 else "-"

        # Buat HTML Rows
        html_rows = ""
        for i, d in enumerate(self.data_over_limit):
            html_rows += f"<tr><td class='tengah'>{i+1}</td><td>BNI {d['cabang']}</td><td class='kanan'>{d['saldo_fmt']}</td><td class='kanan'>{d['pagu_fmt']}</td><td class='kanan'>{d['over_fmt']}</td></tr>"

        tgl = datetime.now().strftime("%d %B %Y")
        
        # Render HTML
        html = TEMPLATE_HTML.format(
            tgl_surat=tgl, 
            no_surat=no_surat_input, 
            nama_manager=self.entry_nama.get(),
            jabatan_manager=jabatan,
            rows=html_rows,
            total_idr=str_sum_idr,
            total_valas=str_sum_usd
        )

        try:
            path_output = os.path.abspath("Surat_Cetak.html")
            with open(path_output, "w", encoding="utf-8") as f:
                f.write(html)
            
            # Simpan Log
            self.log_history(len(self.data_over_limit), str_sum_idr, str_sum_usd)

            output_uri = pathlib.Path(path_output).as_uri()
            webbrowser.open(output_uri)
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat surat: {str(e)}")

if __name__ == "__main__":
    app = AplikasiSurat()
    app.mainloop()
