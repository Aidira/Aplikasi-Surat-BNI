import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import os
import webbrowser
import pathlib
import json
import sqlite3
import base64
import mimetypes
import urllib.request
from datetime import datetime

# --- KONFIGURASI ---
NAMA_FILE_ICON = "logo.png"
CONFIG_FILE = "config_bni.json"
DB_NAME = "database_bni.db"

# URL Logo BNI Resmi (Cadangan jika file lokal tidak ada)
URL_LOGO_ONLINE = "https://upload.wikimedia.org/wikipedia/id/thumb/5/55/BNI_logo.svg/500px-BNI_logo.svg.png"

# --- FITUR AUTO-DOWNLOAD LOGO ---
def cek_dan_download_logo():
    if not os.path.exists(NAMA_FILE_ICON):
        try:
            # Fake User-Agent agar tidak diblokir server Wikimedia
            opener = urllib.request.build_opener()
            opener.addheaders = [('User-agent', 'Mozilla/5.0')]
            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(URL_LOGO_ONLINE, NAMA_FILE_ICON)
        except: pass # Silent error jika gagal, nanti pakai teks saja

# Jalankan cek logo saat start
cek_dan_download_logo()

# --- TEMPLATE SURAT HTML (STANDAR BAKU) ---
TEMPLATE_HTML = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <title>Asuransi BNI</title>
    <style>
        body {{ font-family: "Times New Roman", Times, serif; font-size: 12pt; color: #000; padding: 40px; position: relative; min-height: 100vh; line-height: 1.3; }}
        .surat-header {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; }}
        .surat-alamat {{ width: 65%; }}
        .surat-logo-container {{ width: 30%; text-align: right; }}
        .surat-logo-img {{ max-width: 180px; height: auto; }}
        table {{ width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 11pt; }}
        th, td {{ border: 1px solid black; padding: 4px 8px; vertical-align: middle; }}
        th {{ text-align: center; font-weight: normal; background-color: #f2f2f2; height: 40px; }}
        .kanan {{ text-align: right; }}
        .tengah {{ text-align: center; }}
        p {{ margin-bottom: 5px; margin-top: 5px; }}
        .bold {{ font-weight: bold; }}
        .justify {{ text-align: justify; }}
        .signature {{ margin-top: 40px; page-break-inside: avoid; }}
        .ttd-space {{ height: 80px; }}
        .footer-small {{ position: absolute; bottom: 20px; right: 40px; font-size: 8pt; text-align: right; color: #444; line-height: 1.2; }}
        @media print {{ .no-print {{ display: none !important; }} }}
    </style>
</head>
<body>
    <div class="no-print" style="text-align:center; margin-bottom:20px; padding:10px; background:#f0f0f0; border:1px solid #ccc;">
        <button onclick="window.print()" style="font-size:16px; padding:10px 20px; cursor:pointer; font-weight:bold;">🖨️ KLIK DISINI UNTUK PRINT / SAVE PDF</button>
    </div>

    <div class="surat-header">
        <div class="surat-alamat">
            <p>Jakarta, {tgl_surat}</p>
            <p>No. Surat : TEB/3.2/{no_surat}</p>
            <br>
            <p class="bold">Kepada</p>
            <p class="bold">PT. Asuransi TRI PAKARTA</p>
            <p>Kantor Cabang Jakarta Selatan<br>Komplek Sentra Arteri Mas<br>Jl. Sultan Iskandar Muda No. 10B<br>Jaksel 12240</p>
            <br>
            <p><span class="bold">UP. Ibu Siska</span> &nbsp;&nbsp; <i>Fax. 021-7293312 / 75917755 / 7394748</i></p>
            <p><span class="bold">Hal &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: Cover Asuransi CIS Saldo Kas IDR dan Valas KC/KCP/KK</span></p>
        </div>
        <div class="surat-logo-container">{logo_tag}</div>
    </div>
    <br>

    <p class="justify">Menunjuk perihal pokok surat tersebut diatas, dengan ini kami sampaikan adanya kelebihan pagu kas (over limit) IDR dan Valas di KCP/KK di lingkungan BNI KC Tebet, dengan perincian sbb :</p>

    <table>
        <thead>
            <tr>
                <th width="5%">NO</th>
                <th>KCU/KCP/KK</th>
                <th>Saldo (idr/usd)</th>
                <th>Open (idr/usd)</th>
                <th>Over(idr/usd)</th>
            </tr>
        </thead>
        <tbody>{rows}</tbody>
    </table>

    <div id="infoText">
        <p class="justify">Saldo tersebut telah melebihi cover asuransi cash in save pada open cover Saudara, dengan ini kami laporkan via faksimili/email, agar kelebihan saldo tersebut dapat Saudara tutup dengan asuransi Cash In Save.</p>
        <p class="justify">Demikianlah untuk dimaklumi, atas perhatian dan kerjasama Saudara kami ucapkan terima kasih.</p>
    </div>

    <div class="signature">
        <p>PT. Bank Negara Indonesia (Persero) Tbk<br>Kantor Cabang Tebet</p>
        <div class="ttd-space"></div>
        <p><span class="bold" style="text-decoration: underline;">{nama_manager}</span><br>{jabatan_manager}</p>
    </div>

    <div class="footer-small">
        PT Bank Negara Indonesia (Persero) Tbk<br>Kantor Cabang Utama Tebet<br>Jl. Prof. Supomo SH No. 25, Tebet<br>Jakarta Selatan 12810, Indonesia<br>www.bni.co.id
    </div>
</body>
</html>
"""

# --- DATABASE SETUP ---
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

# --- UTILITIES ---
def get_base64_image(file_path):
    try:
        if not os.path.exists(file_path): return None
        with open(file_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            mime_type, _ = mimetypes.guess_type(file_path)
            return f"data:{mime_type};base64,{encoded_string}"
    except: return None

class AplikasiSurat(ttk.Window):
    def __init__(self):
        # Menggunakan tema 'cosmo' (biru modern) atau 'flatly' (hijau)
        super().__init__(themename="cosmo") 
        self.title("BNI Insurance Generator (Desktop System)")
        self.geometry("1000x750")
        
        init_db()

        # Icon Window
        try:
            if os.path.exists(NAMA_FILE_ICON):
                img = Image.open(NAMA_FILE_ICON)
                self.iconphoto(True, ImageTk.PhotoImage(img))
        except: pass
        
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
            except: pass

    def save_config(self):
        data = {"manager": self.entry_nama.get()}
        try:
            with open(CONFIG_FILE, 'w') as f: json.dump(data, f)
        except: pass

    # --- SIMPAN KE DATABASE ---
    def simpan_ke_db(self, no_surat):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        tgl_sekarang = datetime.now().strftime("%Y-%m-%d")
        
        # Hapus data lama jika nomor surat sama di hari yg sama (replace)
        c.execute("DELETE FROM riwayat_over WHERE no_surat = ? AND tanggal_input = ?", (no_surat, tgl_sekarang))
        
        for d in self.data_over_limit:
            c.execute('''
                INSERT INTO riwayat_over (tanggal_input, no_surat, cabang, mata_uang, saldo, pagu, over_limit)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (tgl_sekarang, no_surat, d['cabang'], d['currency'], d['raw_saldo'], d['raw_pagu'], d['raw_over']))
        
        conn.commit()
        conn.close()
        self.refresh_history()

    def create_widgets(self):
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill=BOTH, expand=YES, padx=10, pady=10)

        # TAB 1: OPERASIONAL
        self.tab_ops = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(self.tab_ops, text="📝 Operasional Surat")
        self.setup_tab_ops()

        # TAB 2: RIWAYAT
        self.tab_hist = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(self.tab_hist, text="🗄️ Database Riwayat")
        self.setup_tab_hist()

    def setup_tab_ops(self):
        # HEADER
        header_frame = ttk.Frame(self.tab_ops)
        header_frame.pack(fill=X, pady=5)
        
        # Logo di dalam aplikasi
        try:
            if os.path.exists(NAMA_FILE_ICON):
                img = Image.open(NAMA_FILE_ICON).resize((60, 20)) # Resize proporsional
                self.logo_tk = ImageTk.PhotoImage(img)
                ttk.Label(header_frame, image=self.logo_tk).pack(side=RIGHT, padx=10)
        except: pass

        ttk.Label(header_frame, text="SISTEM KONTROL KAS (KCU TEBET)", font=("Helvetica", 16, "bold"), bootstyle="primary").pack(side=LEFT, pady=10)

        # INPUT FRAME
        input_frame = ttk.Labelframe(self.tab_ops, text=" Data Input ", padding=10, bootstyle="info")
        input_frame.pack(fill=X, pady=5)

        # File Chooser
        f_file = ttk.Frame(input_frame)
        f_file.pack(fill=X, pady=2)
        self.btn_file = ttk.Button(f_file, text="📂 Upload Excel", command=self.pilih_file, bootstyle="secondary")
        self.btn_file.pack(side=LEFT, padx=(0, 10))
        self.lbl_file_status = ttk.Label(f_file, text="Belum ada file dipilih", foreground="red")
        self.lbl_file_status.pack(side=LEFT)

        # Form Manager & Surat
        f_form = ttk.Frame(input_frame)
        f_form.pack(fill=X, pady=5)
        
        # Kiri: Surat
        f_left = ttk.Frame(f_form)
        f_left.pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))
        ttk.Label(f_left, text="Nomor Surat (4 digit):").pack(anchor=W)
        self.entry_nomor = ttk.Entry(f_left)
        self.entry_nomor.pack(fill=X)

        # Kanan: Manager
        f_right = ttk.Frame(f_form)
        f_right.pack(side=LEFT, fill=X, expand=YES)
        ttk.Label(f_right, text="Nama Manager:").pack(anchor=W)
        self.entry_nama = ttk.Entry(f_right)
        self.entry_nama.insert(0, self.default_manager)
        self.entry_nama.pack(fill=X)
        
        self.var_pgs = ttk.BooleanVar()
        ttk.Checkbutton(f_right, text="Pgs. (Pejabat Pengganti)", variable=self.var_pgs, bootstyle="round-toggle").pack(anchor=W, pady=2)

        # ACTION BUTTONS
        btn_frame = ttk.Frame(self.tab_ops)
        btn_frame.pack(fill=X, pady=10)
        ttk.Button(btn_frame, text="🔍 SCAN EXCEL", command=self.analisa_data, bootstyle="warning-outline", width=20).pack(side=LEFT, padx=(0,5))
        ttk.Button(btn_frame, text="🖨️ CETAK & SIMPAN DB", command=self.cetak_surat, bootstyle="success", width=25).pack(side=LEFT)

        # TABLE PREVIEW
        self.tree = ttk.Treeview(self.tab_ops, columns=('cabang', 'mata_uang', 'saldo', 'pagu', 'over'), show='headings', height=12, bootstyle="info")
        self.tree.heading('cabang', text='KCU/KCP/KK')
        self.tree.heading('mata_uang', text='Curr')
        self.tree.heading('saldo', text='Saldo (idr/usd)')
        self.tree.heading('pagu', text='Open (idr/usd)')
        self.tree.heading('over', text='Over(idr/usd)')
        
        self.tree.column('cabang', width=200)
        self.tree.column('mata_uang', width=50, anchor=CENTER)
        self.tree.column('saldo', width=120, anchor=E)
        self.tree.column('pagu', width=120, anchor=E)
        self.tree.column('over', width=120, anchor=E)
        
        self.tree.pack(fill=BOTH, expand=YES)

    def setup_tab_hist(self):
        ttk.Button(self.tab_hist, text="🔄 Refresh Data", command=self.refresh_history, bootstyle="info-outline").pack(anchor=W, pady=5)
        
        self.hist_tree = ttk.Treeview(self.tab_hist, columns=('tgl', 'no', 'cabang', 'curr', 'over'), show='headings', height=15)
        self.hist_tree.heading('tgl', text='Tanggal')
        self.hist_tree.heading('no', text='No Surat')
        self.hist_tree.heading('cabang', text='Cabang')
        self.hist_tree.heading('curr', text='Mata Uang')
        self.hist_tree.heading('over', text='Over Limit')
        
        self.hist_tree.column('tgl', width=100)
        self.hist_tree.column('no', width=80)
        self.hist_tree.column('cabang', width=150)
        self.hist_tree.column('curr', width=50, anchor=CENTER)
        self.hist_tree.column('over', width=100, anchor=E)
        
        self.hist_tree.pack(fill=BOTH, expand=YES)
        self.refresh_history()

    # --- LOGIC ---
    def refresh_history(self):
        for i in self.hist_tree.get_children(): self.hist_tree.delete(i)
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        try:
            c.execute("SELECT tanggal_input, no_surat, cabang, mata_uang, over_limit FROM riwayat_over ORDER BY id DESC")
            rows = c.fetchall()
            for r in rows:
                over_fmt = "{:,.0f}".format(r[4]).replace(",", ".")
                self.hist_tree.insert('', 'end', values=(r[0], r[1], r[2], r[3], over_fmt))
        except: pass
        conn.close()

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

    def pilih_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.file_path = file
            self.lbl_file_status.config(text=os.path.basename(file), foreground="green")
            for i in self.tree.get_children(): self.tree.delete(i)

    def analisa_data(self):
        if not self.file_path:
            messagebox.showwarning("Warning", "Pilih file Excel dulu!")
            return
        
        try:
            xls = pd.ExcelFile(self.file_path)
            self.data_over_limit = []
            for i in self.tree.get_children(): self.tree.delete(i)
            
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                curr = "USD" if "USD" in sheet_name.upper() else "IDR"
                
                for index, row in df.iterrows():
                    if len(row) < 4: continue
                    cabang = str(row[1])
                    
                    # LOGIKA AGAR KCU TEBET MASUK:
                    # Kita hanya melewati baris yang mengandung 'TOTAL' atau 'NAMA' (Header/Footer)
                    # Kita TIDAK memfilter 'KCU'
                    if pd.isna(cabang) or "TOTAL" in cabang or "NAMA" in cabang.upper(): continue
                    
                    pagu = self.bersihkan_angka(row[2])
                    saldo = self.bersihkan_angka(row[3])
                    
                    if pagu > 0 and saldo > pagu:
                        selisih = saldo - pagu
                        self.data_over_limit.append({
                            "cabang": cabang, 
                            "currency": curr,
                            "raw_saldo": saldo, "raw_pagu": pagu, "raw_over": selisih,
                            "saldo_fmt": self.format_uang(saldo, curr),
                            "pagu_fmt": self.format_uang(pagu, curr), 
                            "over_fmt": self.format_uang(selisih, curr)
                        })
                        self.tree.insert('', 'end', values=(cabang, curr, self.format_uang(saldo, ""), self.format_uang(pagu, ""), self.format_uang(selisih, "")))
            
            if not self.data_over_limit: messagebox.showinfo("Info", "Tidak ada cabang Over Limit.")
            else: messagebox.showinfo("Selesai", f"Ditemukan {len(self.data_over_limit)} cabang Over Limit.")

        except Exception as e: messagebox.showerror("Error", str(e))

    def cetak_surat(self):
        if not self.data_over_limit:
            messagebox.showwarning("Warning", "Scan data dulu!")
            return
        if not self.entry_nomor.get():
            messagebox.showwarning("Warning", "Nomor Surat wajib diisi!")
            return
        
        self.save_config()
        jabatan = "Pgs. Branch Service Manager" if self.var_pgs.get() else "Branch Service Manager"
        
        # Simpan ke Database
        self.simpan_ke_db(self.entry_nomor.get())

        # Generate HTML Rows
        html_rows = ""
        for i, d in enumerate(self.data_over_limit):
            html_rows += f"<tr><td class='tengah'>{i+1}</td><td>BNI {d['cabang']}</td><td class='kanan'>{d['saldo_fmt']}</td><td class='kanan'>{d['pagu_fmt']}</td><td class='kanan'>{d['over_fmt']}</td></tr>"

        base64_logo = get_base64_image(NAMA_FILE_ICON)
        logo_tag_html = f'<img src="{base64_logo}" class="surat-logo-img" alt="BNI Logo">' if base64_logo else ""

        html = TEMPLATE_HTML.format(
            tgl_surat=datetime.now().strftime("%d %B %Y"),
            no_surat=self.entry_nomor.get(),
            nama_manager=self.entry_nama.get(),
            jabatan_manager=jabatan,
            logo_tag=logo_tag_html,
            rows=html_rows
        )

        try:
            path_output = os.path.abspath("Surat_Cetak.html")
            with open(path_output, "w", encoding="utf-8") as f: f.write(html)
            webbrowser.open(pathlib.Path(path_output).as_uri())
            messagebox.showinfo("Sukses", "Surat berhasil dibuat & Data disimpan ke Database!")
        except Exception as e: messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = AplikasiSurat()
    app.mainloop()
