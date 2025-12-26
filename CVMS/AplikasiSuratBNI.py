import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import webbrowser
import pathlib
import platform
from datetime import datetime

# --- TEMPLATE SURAT (HTML) ---
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
        .content-text {{ line-height: 1.5; text-align: justify; }}
        .signature {{ margin-top: 40px; page-break-inside: avoid; }}
        .ttd-space {{ height: 80px; }}
        @media print {{ .no-print {{ display: none !important; }} }}
    </style>
</head>
<body>
    <div class="no-print" style="text-align:center; margin-bottom:20px; padding:10px; background:#f0f0f0; border:1px solid #ccc;">
        <button onclick="window.print()" style="font-size:16px; padding:10px 20px; cursor:pointer; font-weight:bold;">🖨️ KLIK DISINI UNTUK PRINT</button>
    </div>

    <p>Jakarta, {tgl_surat}</p>
    <p>No. Surat : TEB/3.2/{no_surat}</p>
    <br>
    
    <p class="bold">Kepada<br>PT. Asuransi TRI PAKARTA<br>
    <span style="font-weight:normal">Kantor Cabang Jakarta Selatan<br>
    Komplek Sentra Arteri Mas<br>Jl. Sultan Iskandar Muda No. 10B<br>Jaksel 12240</span></p>

    <p class="bold">UP. Ibu Siska (Fax. 021-7293312 / 75917755 / 7394748)</p>
    <p class="bold">Hal : Cover Asuransi CIS Saldo Kas IDR dan Valas KC/KCP/KK</p>

    <p class="content-text">Menunjuk perihal pokok surat tersebut diatas, dengan ini kami sampaikan adanya kelebihan pagu kas (over limit) IDR dan Valas di lingkungan BNI KC Tebet, dengan perincian sbb :</p>

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

    <div id="infoText">
        <p class="content-text">Saldo tersebut telah melebihi cover asuransi cash in save pada open cover Saudara, dengan ini kami laporkan via faksimili/email, agar kelebihan saldo tersebut dapat Saudara tutup dengan asuransi Cash In Save.</p>
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
        # Judul Window
        self.title("Aplikasi Cetak Surat BNI (Universal)")
        self.geometry("500x580")
        self.resizable(False, False)
        
        # --- SETTING WARNA (PENTING BUAT MAC & WINDOWS) ---
        # Kita paksa warna Background (bg) jadi Putih/Abu terang
        # Dan Foreground (fg) atau Teks jadi Hitam.
        self.bg_color = "#f4f4f4"
        self.fg_color = "black"
        self.input_bg = "white"
        
        self.configure(bg=self.bg_color)
        self.file_path = ""
        self.create_widgets()

    def create_widgets(self):
        # Header
        lbl_judul = tk.Label(self, text="GENERATOR SURAT ASURANSI", 
                           font=("Arial", 14, "bold"), bg=self.bg_color, fg="#005E6A")
        lbl_judul.pack(pady=15)

        # Container
        frame = tk.Frame(self, bg=self.bg_color)
        frame.pack(pady=5, padx=20, fill="x")

        # --- 1. PILIH FILE ---
        tk.Label(frame, text="1. File Excel Laporan (.xlsx):", 
                 bg=self.bg_color, fg=self.fg_color, anchor="w", font=("Arial", 10, "bold")).pack(fill="x", pady=(5, 2))
        
        self.btn_file = tk.Button(frame, text="Pilih File Excel", command=self.pilih_file, 
                                  bg="white", fg="black") # Tombol putih teks hitam
        self.btn_file.pack(fill="x", pady=(0, 2))
        
        self.lbl_file_status = tk.Label(frame, text="Belum ada file dipilih", 
                                        bg=self.bg_color, fg="red", font=("Arial", 9))
        self.lbl_file_status.pack(pady=(0, 10))

        # --- 2. NOMOR SURAT ---
        tk.Label(frame, text="2. Nomor Surat (4 digit terakhir):", 
                 bg=self.bg_color, fg=self.fg_color, anchor="w", font=("Arial", 10, "bold")).pack(fill="x", pady=(5, 2))
        
        frame_no = tk.Frame(frame, bg=self.bg_color)
        frame_no.pack(fill="x")
        
        lbl_prefix = tk.Label(frame_no, text="TEB/3.2/", 
                              bg="#e0e0e0", fg="black", padx=5, borderwidth=1, relief="solid")
        lbl_prefix.pack(side="left", fill="y")
        
        # Input: Paksa background putih, teks hitam, kursor hitam
        self.entry_nomor = tk.Entry(frame_no, bg=self.input_bg, fg=self.fg_color, insertbackground="black")
        self.entry_nomor.pack(side="left", fill="x", expand=True, padx=(5,0))

        # --- 3. PENANDA TANGAN ---
        tk.Label(frame, text="3. Penanda Tangan:", 
                 bg=self.bg_color, fg=self.fg_color, anchor="w", font=("Arial", 10, "bold")).pack(fill="x", pady=(15, 2))
        
        tk.Label(frame, text="Nama Manager:", 
                 bg=self.bg_color, fg=self.fg_color, anchor="w", font=("Arial", 9)).pack(fill="x")
        
        self.entry_nama = tk.Entry(frame, bg=self.input_bg, fg=self.fg_color, insertbackground="black")
        self.entry_nama.insert(0, "Hasbiallah")
        self.entry_nama.pack(fill="x", pady=(0, 5))

        # Checkbox
        self.var_pgs = tk.BooleanVar()
        self.chk_pgs = tk.Checkbutton(frame, text="Pgs. (Pejabat Pengganti)", variable=self.var_pgs, 
                                      bg=self.bg_color, fg=self.fg_color, anchor="w")
        self.chk_pgs.pack(fill="x")

        # Tombol Proses
        tk.Frame(self, height=20, bg=self.bg_color).pack() 
        btn_proses = tk.Button(self, text="PROSES & CETAK", command=self.proses_data, 
                               bg="#005E6A", fg="white", font=("Arial", 11, "bold"), height=2)
        # Fix tombol mac kadang tulisan putih di background putih
        if platform.system() == "Darwin": # Jika Mac
             btn_proses.config(fg="black", bg="#e0e0e0") # Ubah jadi standar mac

        btn_proses.pack(fill="x", padx=20, pady=10)

    def pilih_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.file_path = file
            self.lbl_file_status.config(text=f"File: {os.path.basename(file)}", fg="#008000")

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

    def proses_data(self):
        if not self.file_path:
            messagebox.showwarning("Peringatan", "Pilih file Excel dulu!")
            return
        
        no_surat_input = self.entry_nomor.get()
        if not no_surat_input:
            messagebox.showwarning("Peringatan", "Nomor Surat wajib diisi!")
            return

        jabatan = "Branch Service Manager"
        if self.var_pgs.get():
            jabatan = "Pgs. Branch Service Manager"

        try:
            xls = pd.ExcelFile(self.file_path)
            data_over = []

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
                        data_over.append({
                            "cabang": cabang,
                            "saldo": self.format_uang(saldo, curr),
                            "pagu": self.format_uang(pagu, curr),
                            "over": self.format_uang(saldo - pagu, curr)
                        })

            html_rows = ""
            if not data_over:
                html_rows = "<tr><td colspan='5' class='tengah' style='padding:20px'><b>Aman! Tidak ada cabang Over Limit.</b></td></tr>"
            else:
                for i, d in enumerate(data_over):
                    html_rows += f"<tr><td class='tengah'>{i+1}</td><td>BNI {d['cabang']}</td><td class='kanan'>{d['saldo']}</td><td class='kanan'>{d['pagu']}</td><td class='kanan'>{d['over']}</td></tr>"

            tgl = datetime.now().strftime("%d %B %Y")
            html = TEMPLATE_HTML.format(
                tgl_surat=tgl, 
                no_surat=no_surat_input, 
                nama_manager=self.entry_nama.get(),
                jabatan_manager=jabatan,
                rows=html_rows
            )

            # --- BAGIAN UNIVERSAL PATH (Mac/Win) ---
            path_output = os.path.abspath("Surat_Cetak.html")
            with open(path_output, "w", encoding="utf-8") as f:
                f.write(html)
            
            # Membuka browser dengan cara paling aman
            output_uri = pathlib.Path(path_output).as_uri()
            webbrowser.open(output_uri)

        except Exception as e:
            messagebox.showerror("Error Program", f"Terjadi kesalahan:\n{str(e)}")

if __name__ == "__main__":
    app = AplikasiSurat()
    app.mainloop()