import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import os

DB_NAME = os.path.join(os.path.dirname(__file__), "database_bni.db")

st.set_page_config(
    page_title="BNI Over Limit Dashboard",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #f0f4f8; }
.metric-card {
    background: white;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-left: 5px solid;
    margin-bottom: 12px;
}
.card-danger  { border-color: #dc3545; }
.card-warning { border-color: #ffc107; }
.card-info    { border-color: #0d6efd; }
.card-title { font-size: 12px; color: #6c757d; font-weight: 600; text-transform: uppercase; letter-spacing: .5px; }
.card-value { font-size: 28px; font-weight: 700; color: #212529; margin-top: 4px; }
.navbar {
    background: linear-gradient(90deg, #003d82, #0066cc);
    padding: 16px 24px;
    border-radius: 12px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 16px;
}
.navbar h1 { color: white; font-size: 22px; margin: 0; }
.navbar p  { color: rgba(255,255,255,0.75); font-size: 13px; margin: 0; }
</style>
""", unsafe_allow_html=True)


# ── DB ───────────────────────────────────────────────────────────────────────
def init_db():
    conn = sqlite3.connect(DB_NAME)
    conn.execute('''CREATE TABLE IF NOT EXISTS riwayat_over (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tanggal_input DATE, no_surat TEXT, cabang TEXT,
        mata_uang TEXT, saldo REAL, pagu REAL, over_limit REAL
    )''')
    conn.commit()
    conn.close()

def save_to_db(df, no_surat):
    conn = sqlite3.connect(DB_NAME)
    tgl = datetime.now().strftime("%Y-%m-%d")
    for _, row in df.iterrows():
        conn.execute(
            "INSERT INTO riwayat_over (tanggal_input,no_surat,cabang,mata_uang,saldo,pagu,over_limit) VALUES (?,?,?,?,?,?,?)",
            (tgl, no_surat, row["Cabang"], row["Mata Uang"], row["Saldo"], row["Pagu"], row["Over"])
        )
    conn.commit()
    conn.close()

def load_riwayat():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query(
        "SELECT id, tanggal_input as Tanggal, no_surat as 'No. Surat', cabang as Cabang, mata_uang as 'Mata Uang', saldo as Saldo, pagu as Pagu, over_limit as 'Over Limit' FROM riwayat_over ORDER BY id DESC",
        conn
    )
    conn.close()
    return df

def bersihkan_angka(nilai_raw):
    try:
        if isinstance(nilai_raw, (int, float)):
            return float(nilai_raw)
        text = str(nilai_raw).upper().replace("IDR","").replace("RP","").replace(" ","").strip()
        if "." in text and "," in text:
            text = text.replace(".", "").replace(",", ".")
        elif "." in text:
            text = text.replace(".", "")
        elif "," in text:
            text = text.replace(",", ".")
        return float(text)
    except:
        return 0.0

def parse_excel(file):
    xls = pd.ExcelFile(file)
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
            pagu = bersihkan_angka(row[2])
            saldo = bersihkan_angka(row[3])
            over = saldo - pagu
            if over > 0:
                data_over.append([cabang, curr, saldo, pagu, over])
    return pd.DataFrame(data_over, columns=["Cabang", "Mata Uang", "Saldo", "Pagu", "Over"])


# ── Init ─────────────────────────────────────────────────────────────────────
init_db()

if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()

# ── Navbar ───────────────────────────────────────────────────────────────────
logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
col_logo, col_title = st.columns([1, 6])
with col_logo:
    if os.path.exists(logo_path):
        st.image(logo_path, width=130)
with col_title:
    st.markdown("""
    <div style='padding-top:10px'>
        <h2 style='margin:0;color:#003d82;'>BNI Asuransi — Over Limit Dashboard</h2>
        <p style='margin:0;color:#6c757d;font-size:13px;'>Monitoring cabang dengan saldo melebihi pagu</p>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Pengaturan")

    no_surat = st.text_input("Nomor Surat", placeholder="Contoh: S-001/2024")
    nama_manager = st.text_input("Nama Manager", value="Hasbiallah")

    st.markdown("---")
    st.markdown("### 📂 Upload File")
    uploaded = st.file_uploader("Pilih file Excel (.xlsx / .xls)",
                                type=["xlsx", "xls"], label_visibility="collapsed")

    if uploaded:
        with st.spinner("Membaca file..."):
            try:
                st.session_state.df = parse_excel(uploaded)
                st.success(f"{len(st.session_state.df)} data berhasil dimuat")
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")

    st.markdown("---")
    st.markdown("### 💾 Simpan")
    if st.button("Simpan ke Database", type="primary", use_container_width=True):
        if st.session_state.df.empty:
            st.warning("Tidak ada data untuk disimpan.")
        elif not no_surat:
            st.warning("Isi nomor surat terlebih dahulu.")
        else:
            save_to_db(st.session_state.df, no_surat)
            st.success("Data berhasil disimpan!")

    st.markdown(f"<p style='color:#aaa;font-size:11px;margin-top:40px;'>{datetime.now().strftime('%d %B %Y, %H:%M')}</p>", unsafe_allow_html=True)


# ── Main content ──────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📊 Preview Data", "🗂️ Riwayat Database"])

with tab1:
    df = st.session_state.df

    # Stat cards
    total = len(df)
    idr_count = len(df[df["Mata Uang"] == "IDR"]) if not df.empty else 0
    usd_count = len(df[df["Mata Uang"] == "USD"]) if not df.empty else 0
    total_over = df["Over"].sum() if not df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Cabang Over", total)
    c2.metric("Cabang Over IDR", idr_count)
    c3.metric("Cabang Over USD", usd_count)
    c4.metric("Total Over IDR", f"{total_over:,.0f}" if idr_count > 0 else "—")

    st.markdown("---")

    if df.empty:
        st.info("Belum ada data. Silakan upload file Excel di sidebar.")
    else:
        # Filter
        col_f1, col_f2 = st.columns([3, 1])
        with col_f1:
            search = st.text_input("🔍 Cari cabang...", label_visibility="collapsed",
                                   placeholder="Cari nama cabang...")
        with col_f2:
            filter_curr = st.selectbox("Mata Uang", ["Semua", "IDR", "USD"],
                                       label_visibility="collapsed")

        filtered = df.copy()
        if search:
            filtered = filtered[filtered["Cabang"].str.lower().str.contains(search.lower())]
        if filter_curr != "Semua":
            filtered = filtered[filtered["Mata Uang"] == filter_curr]

        # Format angka
        display = filtered.copy().reset_index(drop=True)
        display.index += 1
        for col in ["Saldo", "Pagu", "Over"]:
            display[col] = display[col].apply(lambda x: f"{x:,.2f}")

        st.dataframe(display, use_container_width=True, height=420)
        st.caption(f"Menampilkan {len(filtered)} dari {total} data")

with tab2:
    st.markdown("#### Riwayat Data Tersimpan")
    riwayat = load_riwayat()
    if riwayat.empty:
        st.info("Belum ada data tersimpan di database.")
    else:
        search_r = st.text_input("🔍 Cari riwayat...", placeholder="Cari no. surat atau cabang...",
                                 key="search_riwayat")
        if search_r:
            riwayat = riwayat[
                riwayat["Cabang"].str.lower().str.contains(search_r.lower()) |
                riwayat["No. Surat"].str.lower().str.contains(search_r.lower())
            ]
        for col in ["Saldo", "Pagu", "Over Limit"]:
            riwayat[col] = riwayat[col].apply(lambda x: f"{x:,.2f}")
        st.dataframe(riwayat, use_container_width=True, height=460)
        st.caption(f"{len(riwayat)} record ditemukan")
