"""
Modul pembuat surat PDF "Cover Asuransi CIS Saldo Kas IDR dan Valas KC/KCP/KK".

Mereplikasi tata letak surat resmi BNI KC Tebet ke PT. Asuransi Tri Pakarta:
kop logo BNI, nomor surat, tanggal, alamat tujuan, tabel rincian over-limit
per outlet, paragraf penutup, dan blok tanda tangan Branch Service Manager.
"""

import os

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")

NAMA_BULAN = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]

TUJUAN = [
    "PT. Asuransi TRI PAKARTA",
    "Kantor Cabang Jakarta Selatan",
    "Komplek Sentra Arteri Mas",
    "Jl. Sultan Iskandar Muda No. 10B",
    "Jaksel 12240",
]

FOOTER_LINES = [
    "PT Bank Negara Indonesia (Persero) Tbk",
    "Kantor Cabang Tebet",
    "Jl. Prof. Supomo, SH No.25, Tebet",
    "Jakarta Selatan 12810, Indonesia",
    "www.bni.co.id",
]


def format_tanggal_indonesia(tanggal):
    """Memformat objek date/datetime menjadi 'Jakarta, DD Bulan YYYY'."""
    return f"Jakarta, {tanggal.day:02d} {NAMA_BULAN[tanggal.month]} {tanggal.year}"


def format_nilai(nilai, mata_uang):
    """Memformat nilai numerik dengan pemisah ribuan titik dan kode mata uang, mis. 'IDR 8.392.959.700'."""
    teks = f"{abs(nilai):,.0f}".replace(",", ".")
    return f"{mata_uang.upper()} {teks}"


def buat_surat_pdf(output_path, no_surat, nama_manager, tanggal, rows):
    """
    Membuat surat PDF "Cover Asuransi CIS Saldo Kas IDR dan Valas KC/KCP/KK".

    `rows` adalah list of dict dengan key: cabang, mata_uang, saldo, pagu, over.
    """
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
        leftMargin=2 * cm, rightMargin=2 * cm,
    )
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("normal", parent=styles["Normal"], fontName="Helvetica", fontSize=10, leading=14)
    bold = ParagraphStyle("bold", parent=normal, fontName="Helvetica-Bold")

    elemen = []

    # --- Header: tanggal kiri, logo kanan ---
    logo = Image(LOGO_PATH, width=4 * cm, height=4 * cm * 417 / 1280)
    header_tbl = Table(
        [[Paragraph(format_tanggal_indonesia(tanggal), normal), logo]],
        colWidths=[11 * cm, 5 * cm],
    )
    header_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
    ]))
    elemen.append(header_tbl)
    elemen.append(Spacer(1, 0.8 * cm))

    # --- Nomor surat & tujuan ---
    elemen.append(Paragraph(f"No.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: TEB/3.2/{no_surat}", normal))
    elemen.append(Paragraph("S", normal))
    elemen.append(Paragraph("Kepada", normal))
    elemen.append(Spacer(1, 0.3 * cm))
    for i, baris in enumerate(TUJUAN):
        style = bold if i == 0 else normal
        elemen.append(Paragraph(baris, style))
    elemen.append(Spacer(1, 0.4 * cm))

    elemen.append(Paragraph(
        "<i>UP.Ibu.Siska&nbsp;&nbsp;Fax.021-7293312 / 75917755 / 7394748</i>", normal,
    ))
    elemen.append(Paragraph(
        "<b>Hal&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: Cover Asuransi CIS Saldo Kas IDR dan Valas KC/KCP/KK</b>",
        normal,
    ))
    elemen.append(Spacer(1, 0.4 * cm))

    elemen.append(Paragraph(
        "Menunjuk perihal pokok surat tersebut diatas, dengan ini kami sampaikan adanya "
        "kelebihan pagu kas (over limit) IDR dan Valas di KCP/KK di lingkungan BNI KC Tebet, "
        "dengan perincian sbb :",
        normal,
    ))
    elemen.append(Spacer(1, 0.4 * cm))

    # --- Tabel rincian over-limit ---
    header = ["NO", "KCU/KCP/KK", "Saldo (idr/usd)", "Open (idr/usd)", "Over(idr/usd)"]
    data = [header]
    for i, r in enumerate(rows, start=1):
        data.append([
            str(i),
            r["cabang"],
            format_nilai(r["saldo"], r["mata_uang"]),
            format_nilai(r["pagu"], r["mata_uang"]),
            format_nilai(r["over"], r["mata_uang"]),
        ])

    tabel = Table(data, colWidths=[1.5 * cm, 4 * cm, 4 * cm, 4 * cm, 3.5 * cm])
    tabel.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elemen.append(tabel)
    elemen.append(Spacer(1, 0.5 * cm))

    elemen.append(Paragraph(
        "Saldo tersebut telah melebihi cover asuransi cash in save pada open cover Saudara, "
        "dengan ini kami laporkan via faksimili/email, agar kelebihan saldo tersebut dapat "
        "Saudara tutup dengan asuransi Cash In Save.",
        normal,
    ))
    elemen.append(Spacer(1, 0.3 * cm))
    elemen.append(Paragraph(
        "Demikianlah untuk dimaklumi, atas perhatian dan kerjasama Saudara kami ucapkan terima kasih.",
        normal,
    ))
    elemen.append(Spacer(1, 0.8 * cm))

    # --- Blok tanda tangan ---
    elemen.append(Paragraph("PT.Bank Negara Indonesia (Persero) Tbk", normal))
    elemen.append(Paragraph("Kantor Cabang Tebet", normal))
    elemen.append(Spacer(1, 1.6 * cm))
    elemen.append(Paragraph(f"<u>{nama_manager}</u>", normal))
    elemen.append(Paragraph("Branch Service Manager", normal))

    def gambar_footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        y = 1.3 * cm
        for line in reversed(FOOTER_LINES):
            canvas.drawRightString(A4[0] - 2 * cm, y, line)
            y += 0.32 * cm
        canvas.restoreState()

    doc.build(elemen, onFirstPage=gambar_footer, onLaterPages=gambar_footer)
    return output_path
