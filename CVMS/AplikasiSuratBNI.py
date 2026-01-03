import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import os
import webbrowser
import pathlib
import platform
import csv
import json
from datetime import datetime

# --- KONFIGURASI ---
NAMA_FILE_ICON = "logo_bni.png" 
CONFIG_FILE = "config_bni.json"
HISTORY_FILE = "riwayat_cetak.csv"

# --- TEMPLATE SURAT HTML (Sama seperti sebelumnya) ---
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

class AplikasiSurat(ttk.Window):
    def __init__(self):
        # Gunakan theme "cosmo" (putih bersih, biru modern) atau "superhero" (gelap)
        super().__init__(themename="cosmo") 
        self.title("BNI Insurance Generator")
        self.geometry("600x800")
        self.resizable(True, True)
        
        # Setting Icon
        try:
            if os.path.exists(NAMA_FILE_ICON):
                img_data = Image.open(NAMA_FILE_ICON)
                self.icon_img = ImageTk.PhotoImage(img_data)
                self.iconphoto(True, self.icon_img)
        except:
            pass
        
        self.file_path = ""
        self.data_over_limit = []
        
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
        file_exists = os.path.isfile(HISTORY_FILE)
        try:
            with open(HISTORY_FILE, mode='a', newline='') as file:
                writer = csv.writer(file)
                if not file_exists:
                    writer.writerow(["Tanggal", "Jam", "No Surat", "Manager", "Jml Cabang", "Total IDR", "Total Valas"])
                now = datetime.now()
                writer.writerow([
                    now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"),
                    self.entry_nomor.get(), self.entry_nama.get(),
                    total_cabang, total_idr, total_valas
                ])
        except Exception as e:
            print(f"Log error: {e}")

    def create_widgets(self):
        # Container Utama (Padding agar tidak mepet pinggir)
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=YES)

        # --- HEADER ---
        lbl_judul = ttk.Label(main_frame, text="GENERATOR SURAT ASURANSI", 
                              font=("Helvetica", 16, "bold"), bootstyle="primary")
        lbl_judul.pack(pady=(0, 20))

        # --- SEKSI 1: INPUT DATA ---
        info_frame = ttk.Labelframe(main_frame, text=" Data Laporan ", padding=15, bootstyle="info")
        info_frame.pack(fill=X, pady=5)

        ttk.Label(info_frame, text="File Laporan Excel:", font=("Arial", 9, "bold")).pack(anchor=W)
        
        # Tombol & Status File
        file_frame = ttk.Frame(info_frame)
        file_frame.pack(fill=X, pady=5)
        
        self.btn_file = ttk.Button(file_frame, text="📂 Pilih File", command=self.pilih_file, bootstyle="secondary")
        self.btn_file.pack(side=LEFT, padx=(0, 10))
        
        self.lbl_file_status = ttk.Label(file_frame, text="Belum ada file", foreground="red")
        self.lbl_file_status.pack(side=LEFT)

        # --- SEKSI 2: FORM SURAT ---
        form_frame = ttk.Labelframe(main_frame, text=" Detail Surat ", padding=15, bootstyle="info")
        form_frame.pack(fill=X, pady=15)

        # Nomor Surat
        ttk.Label(form_frame, text="Nomor Surat (4 digit):").pack(anchor=W)
        input_no_frame = ttk.Frame(form_frame)
        input_no_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(input_no_frame, text="TEB/3.2/", bootstyle="inverse-secondary", padding=5).pack(side=LEFT)
        self.entry_nomor = ttk.Entry(input_no_frame)
        self.entry_nomor.pack(side=LEFT, fill=X, expand=YES, padx=(5,0))

        # Manager
        ttk.Label(form_frame, text="Nama Manager:").pack(anchor=W)
        self.entry_nama = ttk.Entry(form_frame)
        self.entry_nama.insert(0, self.default_manager)
        self.entry_nama.pack(fill=X, pady=(0, 5))

        self.var_pgs = ttk.BooleanVar()
        self.chk_pgs = ttk.Checkbutton(form_frame, text="Pgs. (Pejabat Pengganti)", variable=self.var_pgs, bootstyle="round-toggle")
        self.chk_pgs.pack(anchor=W, pady=5)

        # --- SEKSI 3: AKSI & PREVIEW ---
        aksi_frame = ttk.Frame(main_frame)
        aksi_frame.pack(fill=BOTH, expand=YES, pady=10)

        # Tombol Preview
        ttk.Button(aksi_frame, text="🔍 SCAN & PREVIEW DATA", command=self.analisa_data, bootstyle="warning-outline").pack(fill=X, pady=5)

        # Tabel Modern
        columns = ('cabang', 'saldo', 'pagu', 'over')
        self.tree = ttk.Treeview(aksi_frame, columns=columns, show='headings', height=6, bootstyle="info")
        
        self.tree.heading('cabang', text='Cabang')
        self.tree.heading('saldo', text='Saldo')
        self.tree.heading('pagu', text='Pagu')
        self.tree.heading('over', text='Over Limit')
        
        self.tree.column('cabang', width=120)
        self.tree.column('saldo', width=100, anchor=E)
        self.tree.column('pagu', width=100, anchor=E)
        self.tree.column('over', width=100, anchor=E)
        
        self.tree.pack(fill=BOTH, expand=YES, pady=5)

        # Tombol Cetak (Besar)
        ttk.Button(aksi_frame, text="🖨️ CETAK SURAT SEKARANG", command=self.cetak_surat, bootstyle="success", width=30).pack(fill=X, pady=10)

    # --- LOGIKA PROGRAM (SAMA SEPERTI SEBELUMNYA) ---
    def pilih_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.file_path = file
            self.lbl_file_status.config(text=os.path.basename(file), foreground="green")
            for i in self.tree.get_children(): self.tree.delete(i)

    def format_uang(self, nilai, currency="IDR"):
        try: return "{} {:,.0f}".format(currency, float(nilai)).replace(",", ".")
        except: return "0"

    def bersihkan_angka(self, nilai_raw):
        try:
            if isinstance(nilai_raw, (int, float)): return float(nilai_raw)
            text = str(nilai_raw).upper().replace("IDR", "").replace("RP", "").replace(" ", "").strip()
            if "." in text and "," in text: text = text.replace(".", "").replace(",", ".")
            elif "." in text: text = text.replace(".", "")
            elif "," in text: text = text.replace(",", ".")
            return float(text)
        except: return 0.0

    def analisa_data(self):
        if not self.file_path:
            messagebox.showwarning("Peringatan", "Pilih file Excel dulu!")
            return
        try:
            xls = pd.ExcelFile(self.file_path)
            self.data_over_limit = []
            for i in self.tree.get_children(): self.tree.delete(i)
            count_over = 0

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                curr = "USD" if "USD" in sheet_name.upper() else "IDR"
                for index, row in df.iterrows():
                    if len(row) < 4: continue
                    cabang = str(row[1])
                    if pd.isna(cabang) or "KCU" in cabang or "TOTAL" in cabang or "NAMA" in cabang.upper() or cabang.strip() == "": continue

                    pagu = self.bersihkan_angka(row[2])
                    saldo = self.bersihkan_angka(row[3])

                    if pagu > 0 and saldo > pagu:
                        selisih = saldo - pagu
                        self.data_over_limit.append({
                            "cabang": cabang, "saldo_fmt": self.format_uang(saldo, curr),
                            "pagu_fmt": self.format_uang(pagu, curr), "over_fmt": self.format_uang(selisih, curr),
                            "raw_over": selisih, "currency": curr
                        })
                        self.tree.insert('', 'end', values=("BNI " + cabang, self.format_uang(saldo, curr), self.format_uang(pagu, curr), self.format_uang(selisih, curr)))
                        count_over += 1
            
            if count_over == 0: messagebox.showinfo("Aman", "Tidak ada cabang Over Limit.")
            else: messagebox.showinfo("Selesai", f"Ditemukan {count_over} cabang Over Limit.")

        except Exception as e: messagebox.showerror("Error", f"Gagal baca file: {str(e)}")

    def cetak_surat(self):
        if not self.data_over_limit:
            messagebox.showwarning("Peringatan", "Data kosong!")
            return
        if not self.entry_nomor.get():
            messagebox.showwarning("Peringatan", "Nomor Surat wajib diisi!")
            return
        self.save_config()

        jabatan = "Pgs. Branch Service Manager" if self.var_pgs.get() else "Branch Service Manager"
        sum_idr = sum(d['raw_over'] for d in self.data_over_limit if d['currency'] == 'IDR')
        sum_usd = sum(d['raw_over'] for d in self.data_over_limit if d['currency'] == 'USD')
        
        html_rows = ""
        for i, d in enumerate(self.data_over_limit):
            html_rows += f"<tr><td class='tengah'>{i+1}</td><td>BNI {d['cabang']}</td><td class='kanan'>{d['saldo_fmt']}</td><td class='kanan'>{d['pagu_fmt']}</td><td class='kanan'>{d['over_fmt']}</td></tr>"

        tgl = datetime.now().strftime("%d %B %Y")
        html = TEMPLATE_HTML.format(
            tgl_surat=tgl, no_surat=self.entry_nomor.get(),
            nama_manager=self.entry_nama.get(), jabatan_manager=jabatan,
            rows=html_rows,
            total_idr=self.format_uang(sum_idr, "IDR") if sum_idr > 0 else "-",
            total_valas=self.format_uang(sum_usd, "USD") if sum_usd > 0 else "-"
        )

        try:
            path_output = os.path.abspath("Surat_Cetak.html")
            with open(path_output, "w", encoding="utf-8") as f: f.write(html)
            self.log_history(len(self.data_over_limit), self.format_uang(sum_idr), self.format_uang(sum_usd))
            webbrowser.open(pathlib.Path(path_output).as_uri())
        except Exception as e: messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = AplikasiSurat()
    app.mainloop()
