import io
import math
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="AI Smart HPP Calculator",
    page_icon="🧮",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------
# Helper functions
# ---------------------------
def rupiah(value):
    return f"Rp{value:,.0f}".replace(",", ".")


def format_quantity(value, decimals=4):
    """Tampilkan angka seperlunya tanpa deretan nol yang tidak diperlukan."""
    if pd.isna(value):
        return ""
    text = f"{float(value):,.{decimals}f}"
    return text.rstrip("0").rstrip(".") if "." in text else text


def style_table_alignment(df):
    """Kolom nominal rupiah rata kanan; seluruh kolom selain rupiah rata tengah."""
    price_keywords = ("harga", "biaya", "hpp", "rupiah", "(rp", "nilai addcost", "nilai pajak")
    price_columns = [
        col for col in df.columns
        if any(keyword in str(col).lower() for keyword in price_keywords)
    ]
    other_columns = [col for col in df.columns if col not in price_columns]

    styler = df.style
    if other_columns:
        styler = styler.set_properties(subset=other_columns, **{"text-align": "center"})
    if price_columns:
        styler = styler.set_properties(subset=price_columns, **{"text-align": "right"})

    # Rapikan angka kuantitas dan persentase agar tidak menampilkan 6 angka desimal.
    formats = {}
    for col in df.columns:
        name = str(col).lower()
        if col in price_columns:
            continue
        if "persentase" in name or "(%)" in name or "target hpp" in name or "hpp aktual" in name:
            formats[col] = lambda x: format_quantity(x, 2) if isinstance(x, (int, float)) else x
        elif pd.api.types.is_numeric_dtype(df[col]):
            formats[col] = lambda x: format_quantity(x, 4) if isinstance(x, (int, float)) else x
    if formats:
        styler = styler.format(formats, na_rep="")
    return styler


def round_up(value, increment=500):
    return int(math.ceil(value / increment) * increment)


def classify_hpp(actual_pct, target_pct):
    """Forward Chaining: evaluasi berurutan berdasarkan fakta HPP aktual."""
    b1 = target_pct * 0.75
    b2 = target_pct
    b3 = target_pct * 1.25

    if actual_pct <= b1:
        return {
            "rule": "R1",
            "status": "HPP sangat efisien",
            "recommendation": "Biaya berada jauh di bawah target. Pertahankan efisiensi, sambil memastikan kualitas dan ukuran produk tetap sesuai standar.",
            "condition": f"HPP aktual ≤ {b1:.2f}% (75% dari target)",
            "level": "Efisien",
        }
    elif actual_pct <= b2:
        return {
            "rule": "R2",
            "status": "HPP sesuai target",
            "recommendation": "Pertahankan pengendalian biaya dan lakukan pemantauan harga bahan/komponen secara berkala.",
            "condition": f"{b1:.2f}% < HPP aktual ≤ {b2:.2f}%",
            "level": "Sesuai target",
        }
    elif actual_pct <= b3:
        return {
            "rule": "R3",
            "status": "HPP perlu dievaluasi",
            "recommendation": "Periksa komponen biaya terbesar, penggunaan bahan, pemasok, dan kemungkinan pemborosan. Evaluasi harga jual bila diperlukan.",
            "condition": f"{b2:.2f}% < HPP aktual ≤ {b3:.2f}%",
            "level": "Perlu evaluasi",
        }
    return {
        "rule": "R4",
        "status": "HPP terlalu tinggi",
        "recommendation": "Lakukan evaluasi biaya secara menyeluruh. Pertimbangkan negosiasi pemasok, efisiensi pemakaian, perubahan spesifikasi, atau penyesuaian harga jual.",
        "condition": f"HPP aktual > {b3:.2f}%",
        "level": "Tinggi",
    }


def make_excel(data, materials, costs, inference):
    output = io.BytesIO()
    summary = pd.DataFrame(
        [{"Parameter": k, "Nilai": v} for k, v in data.items()]
    )
    material_columns = [
        "Komponen", "Harga beli (Rp)", "Isi kemasan", "Satuan isi kemasan",
        "Pemakaian per unit", "Satuan pemakaian", "Pemakaian (dikonversi)",
        "Biaya per unit (Rp)"
    ]
    material_df = pd.DataFrame(materials)[material_columns]
    cost_df = pd.DataFrame(costs)
    inference_df = pd.DataFrame(
        [{"Tahap": "Fakta awal", "Keterangan": f"HPP aktual {data['HPP Aktual (%)']:.2f}% dan target HPP {data['Target HPP (%)']:.2f}%"},
         {"Tahap": "Aturan aktif", "Keterangan": f"{inference['rule']} — {inference['condition']}"},
         {"Tahap": "Kesimpulan", "Keterangan": inference["status"]},
         {"Tahap": "Rekomendasi", "Keterangan": inference["recommendation"]}]
    )
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Ringkasan HPP")
        material_df.to_excel(writer, index=False, sheet_name="Rincian Komponen")
        cost_df.to_excel(writer, index=False, sheet_name="Biaya Tambahan")
        inference_df.to_excel(writer, index=False, sheet_name="Forward Chaining")
    output.seek(0)
    return output.getvalue()


def make_pdf(data, materials, costs, inference):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.units import mm
    from xml.sax.saxutils import escape
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    )

    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output, pagesize=A4,
        rightMargin=18 * mm, leftMargin=18 * mm,
        topMargin=16 * mm, bottomMargin=16 * mm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="AppTitle", parent=styles["Title"],
        alignment=TA_CENTER, fontSize=17, leading=21, spaceAfter=5
    ))
    story = [
        Paragraph("AI Smart HPP Calculator", styles["AppTitle"]),
        Paragraph("Laporan Perhitungan Harga Pokok Produksi", styles["Heading2"]),
        Paragraph(f"Tanggal laporan: {datetime.now().strftime('%d-%m-%Y %H:%M')}", styles["Normal"]),
        Spacer(1, 8),
    ]

    summary_rows = [["Parameter", "Nilai"]] + [[str(k), str(v)] for k, v in data.items()]
    t = Table(summary_rows, colWidths=[65 * mm, 105 * mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#183B56")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    story += [Paragraph("1. Ringkasan Perhitungan", styles["Heading2"]), t, Spacer(1, 10)]

    mat_rows = [[
        "Komponen", "Harga beli (Rp)", "Isi kemasan",
        "Pemakaian/unit", "Biaya/unit (Rp)"
    ]]
    for m in materials:
        mat_rows.append([
            m["Komponen"],
            f'{m["Harga beli (Rp)"]:,.0f}',
            f'{m["Isi kemasan"]:g} {m["Satuan isi kemasan"]}',
            f'{m["Pemakaian per unit"]:g} {m["Satuan pemakaian"]}',
            f'{m["Biaya per unit (Rp)"]:,.0f}'
        ])
    mt = Table(
        mat_rows,
        colWidths=[42 * mm, 30 * mm, 32 * mm, 32 * mm, 34 * mm],
        repeatRows=1
    )
    mt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#183B56")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    story += [Paragraph("2. Rincian Bahan/Komponen", styles["Heading2"]), mt, Spacer(1, 10)]

    cost_rows = [["Biaya tambahan", "Biaya per unit (Rp)"]]
    cost_rows += [[c["Jenis biaya"], f'{c["Biaya per unit (Rp)"]:,.0f}'] for c in costs]
    ct = Table(cost_rows, colWidths=[100 * mm, 70 * mm], repeatRows=1)
    ct.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#183B56")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    story += [Paragraph("3. Biaya Tambahan", styles["Heading2"]), ct, Spacer(1, 10)]
    story += [
        Paragraph("4. Hasil Analisis HPP", styles["Heading2"]),
        Paragraph(f"Aturan aktif: <b>{escape(str(inference['rule']))}</b>", styles["Normal"]),
        Paragraph(f"Kondisi: {escape(str(inference['condition']))}", styles["Normal"]),
        Paragraph(f"Kesimpulan: <b>{escape(str(inference['status']))}</b>", styles["Normal"]),
        Paragraph(f"Rekomendasi: {escape(str(inference['recommendation']))}", styles["Normal"]),
        Spacer(1, 10),
        Paragraph("Catatan: hasil merupakan estimasi berdasarkan data dan asumsi biaya yang dimasukkan pengguna.", styles["Italic"]),
    ]
    doc.build(story)
    output.seek(0)
    return output.getvalue()


def make_multi_excel(products):
    output = io.BytesIO()
    summary_rows, material_rows, cost_rows, analysis_rows = [], [], [], []
    for no, item in enumerate(products, 1):
        d = item["data"]
        summary_rows.append({"No": no, **d})
        for m in item["materials"]:
            material_rows.append({"No": no, "Nama produk": d["Nama produk"], **{k:v for k,v in m.items() if k != "Satuan valid"}})
        for c in item["costs"]:
            cost_rows.append({"No": no, "Nama produk": d["Nama produk"], **c})
        inf = item["inference"]
        analysis_rows.append({"No": no, "Nama produk": d["Nama produk"], "Aturan": inf["rule"], "Kondisi": inf["condition"], "Kesimpulan": inf["status"], "Rekomendasi": inf["recommendation"]})
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="Rekap Semua Produk")
        pd.DataFrame(material_rows).to_excel(writer, index=False, sheet_name="Rincian Komponen")
        pd.DataFrame(cost_rows).to_excel(writer, index=False, sheet_name="Addcost Pajak Kemasan")
        pd.DataFrame(analysis_rows).to_excel(writer, index=False, sheet_name="Hasil Analisis")
    output.seek(0)
    return output.getvalue()


def make_multi_pdf(products):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.units import mm
    from xml.sax.saxutils import escape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, rightMargin=15*mm, leftMargin=15*mm, topMargin=14*mm, bottomMargin=14*mm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="MultiTitle", parent=styles["Title"], alignment=TA_CENTER, fontSize=16, leading=20))
    story = [Paragraph("AI Smart HPP Calculator", styles["MultiTitle"]), Paragraph("Rekap Laporan Perhitungan HPP Seluruh Produk", styles["Heading2"]), Paragraph(f"Tanggal laporan: {datetime.now().strftime('%d-%m-%Y %H:%M')}", styles["Normal"]), Spacer(1, 8)]
    recap = [["No", "Produk", "Total HPP/unit", "HPP Aktual", "Harga Jual", "Rekomendasi", "Hasil"]]
    for no, item in enumerate(products, 1):
        d=item["data"]; inf=item["inference"]
        recap.append([str(no), escape(str(d["Nama produk"])), d["Total HPP per unit (Rp)"], f'{d["HPP Aktual (%)"]:.2f}%', d["Harga jual saat ini (Rp)"], d["Harga jual rekomendasi (Rp)"], escape(inf["status"])])
    t=Table(recap, colWidths=[9*mm, 34*mm, 29*mm, 19*mm, 25*mm, 28*mm, 34*mm], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#183B56")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),7),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("PADDING",(0,0),(-1,-1),4)]))
    story += [Paragraph("Ringkasan Seluruh Produk", styles["Heading2"]), t]
    for no, item in enumerate(products, 1):
        d=item["data"]; inf=item["inference"]
        story += [PageBreak(), Paragraph(f'{no}. {escape(str(d["Nama produk"]))}', styles["Heading1"])]
        rows=[["Parameter", "Nilai"]]+[[str(k), (f'{v:.2f}%' if k=="HPP Aktual (%)" else str(v))] for k,v in d.items()]
        stbl=Table(rows, colWidths=[65*mm, 105*mm], repeatRows=1)
        stbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#183B56")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.4,colors.grey),("PADDING",(0,0),(-1,-1),5)]))
        story += [Paragraph("Ringkasan Perhitungan", styles["Heading2"]), stbl, Spacer(1,8)]
        mr=[["Komponen", "Harga beli", "Isi kemasan", "Pemakaian/unit", "Biaya/unit"]]
        for m in item["materials"]:
            mr.append([escape(str(m["Komponen"])), f'{m["Harga beli (Rp)"]:,.0f}', f'{m["Isi kemasan"]:g} {m["Satuan isi kemasan"]}', f'{m["Pemakaian per unit"]:g} {m["Satuan pemakaian"]}', f'{m["Biaya per unit (Rp)"]:,.0f}'])
        mt=Table(mr, colWidths=[43*mm,27*mm,32*mm,35*mm,31*mm], repeatRows=1)
        mt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#168B89")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),8),("PADDING",(0,0),(-1,-1),4)]))
        story += [Paragraph("Rincian Bahan/Komponen", styles["Heading2"]), mt, Spacer(1,8), Paragraph("Biaya Tambahan", styles["Heading2"])]
        cr=[["Jenis biaya", "Nilai per unit"]]+[[c["Jenis biaya"], f'{c["Biaya per unit (Rp)"]:,.0f}'] for c in item["costs"]]
        ct=Table(cr, colWidths=[90*mm,80*mm]); ct.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#168B89")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.4,colors.grey),("PADDING",(0,0),(-1,-1),5)]))
        story += [ct, Spacer(1,8), Paragraph("Hasil Analisis", styles["Heading2"]), Paragraph(f'Kesimpulan: <b>{escape(inf["status"])}</b>', styles["Normal"]), Paragraph(f'Kriteria: {escape(inf["condition"])}', styles["Normal"]), Paragraph(f'Rekomendasi: {escape(inf["recommendation"])}', styles["Normal"])]
    doc.build(story); output.seek(0); return output.getvalue()


# ---------------------------
# Styling
# ---------------------------
st.markdown("""
<style>
/* Latar belakang lembut dengan aksen gradasi */
.stApp {
    background:
      radial-gradient(circle at 8% 4%, rgba(91, 141, 239, .13), transparent 27%),
      radial-gradient(circle at 92% 16%, rgba(28, 190, 170, .12), transparent 25%),
      linear-gradient(135deg, #f5f8ff 0%, #f8fbff 48%, #f1faf8 100%);
}
.block-container {padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1320px;}
/* Header utama */
.hero-wrap {
    display: flex; align-items: center; justify-content: flex-start; gap: 26px;
    padding: 27px 32px; margin: 4px 0 22px 0; border-radius: 22px;
    background: linear-gradient(115deg, #142b4a 0%, #245b83 58%, #168b89 100%);
    box-shadow: 0 14px 34px rgba(24, 59, 86, .18);
    color: white; overflow: hidden; position: relative;
}
.hero-title {font-size: clamp(1.7rem, 3vw, 2.45rem); font-weight: 800; letter-spacing: -.7px; line-height: 1.15; color: #fff;}
.hero-subtitle {font-size: 1.05rem; color: #b9f0e7; margin-top: 7px; font-weight: 700;}
.hero-logo-wrap {width: 148px; min-width: 148px; height: 148px; padding: 8px; border-radius: 18px; background: #fff; display:flex; align-items:center; justify-content:center; box-shadow: 0 5px 18px rgba(0,0,0,.16);}
.unm-logo {width: 100%; height: 100%; object-fit: contain; border-radius: 10px;}
.hero-content {min-width: 0;}
.hero-details {font-size: .91rem; line-height: 1.65; color: rgba(255,255,255,.91); margin-top: 13px;}
.hero-university {font-weight: 800; color: #fff; font-size: 1rem; margin-top: 4px;}
.intro-text {color: #53657b; font-size: .98rem; margin: 0 0 25px 2px;}
/* Sidebar */
section[data-testid="stSidebar"] {background: linear-gradient(180deg, #eaf1fb 0%, #f5fafb 100%); border-right: 1px solid #dce6f0;}
section[data-testid="stSidebar"] h2 {color: #193a5b;}
/* Judul bagian dan panel input */
h1, h2, h3 {color: #183653;}
.stHeading h2 {font-weight: 750; letter-spacing: -.35px;}
div[data-testid="stVerticalBlockBorderWrapper"] {background: rgba(255,255,255,.82); border: 1px solid #dce6f0; border-radius: 16px; box-shadow: 0 5px 18px rgba(35,65,95,.045);}
div[data-testid="stMetric"] {background: linear-gradient(145deg, #ffffff, #f1f7fc); border: 1px solid #d8e5f0; padding: 15px 16px; border-radius: 15px; box-shadow: 0 5px 15px rgba(35,65,95,.06);}
/* Input */
div[data-baseweb="input"] > div, div[data-baseweb="select"] > div, div[data-baseweb="textarea"] > div {background: #fff; border-radius: 10px;}
/* Tombol */
div.stButton > button, div.stDownloadButton > button {border: 0; border-radius: 11px; font-weight: 700; padding: .58rem 1.1rem; background: linear-gradient(100deg, #245b83, #168b89); color: #fff; box-shadow: 0 5px 14px rgba(36,91,131,.16); transition: transform .15s ease, box-shadow .15s ease;}
div.stButton > button:hover, div.stDownloadButton > button:hover {transform: translateY(-1px); box-shadow: 0 8px 18px rgba(36,91,131,.23); color: #fff; border: 0;}
hr {border-color: #dce6ef;}
.small-note {font-size: .88rem; color: #64748b;}
@media (max-width: 700px) {
  .hero-wrap {padding: 22px 20px;}
  .hero-wrap {align-items: flex-start; gap: 16px;}
  .hero-logo-wrap {width: 88px; min-width: 88px; height: 88px; padding: 5px;}
  .hero-details {font-size: .78rem;}
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    f"""
<div class="hero-wrap">
  <div class="hero-logo-wrap">
    <img class="unm-logo" src="data:image/png;base64,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Xyaz4e8wOrPhz1U4WkVhviQQIAHrBqcCYN26YclIgAQUAUdyHknsMpUNv2LB/yQQSgJUAEJJk2mRAAmEnEBCgiZJSTlCni4TJIFIELByHlQArFw7LBsJkAAJkAAJhIkAFYAwgWWyJEACJEACdidgbfmpAFi7flg6EiABEggpAcfJvSFNj4nFLgEqALFbdyw5CRHF/l0AABAASURBVJAACQRO4PzpwOMwRlAErB6JCoDVa4jlIwESIAESIIEwEKACEAaoTJIESIAESMDuBKwvPxUA69cRS0gCJEACJEACISdABSDkSJkgCZAACZCA3QnEgvxUAGKhllhGEiABEiABEggxASoAIQbK5EiABEiABOxOIDbkpwIQG/XEUpIACZAACZBASAlQAQgpTiZGAiRAAiRgdwKxIj8VgFipKZaTBEiABEiABEJIgApACGEyKRIgARIgAbsTiB35qQDETl2xpCRAAiSQbQKOE/wWQLYhxkkCVADipCIpBgmQAAn4RSD9nF/BGCg4ArEUiwpALNUWy0oCJEACJEACISJABSBEIJkMCZAACZCA3QnElvxUAGKrvlhaEiABEiABEggJASoAIcHIREiABEiABOxOINbkpwIQazXG8pIACZAACZBACAhQAQgBRCZBAqEkkJHhEMfRnaFMkmmRAAmEnUDsZUAFIPbqjCWOcwIJCZr8+++JOJeS4pEACUSbABWAaNcA8ycBEiABEoh5ArEoABWAWKw1lpkESIAESIAEskmACkA2ATI6CZAACZCA3QnEpvxUAGKz3lhqEiABEgiKgOP0kaDiMVL8EaACEH91SolIgARIwJzA2aPmfvQJikCsRqICEKs1x3KTAAmQAAmQQDYIUAHIBjxGJQESIAESsDuB2JWfCkDs1h1LTgIkQAIkQAJBE6ACEDQ6RiQBEiABErA7gViWnwpALNcey04CJEACJEACQRKgAhAkOEYjARIgARKwO4HYlp8KQGzXH0tPAiRAAiRAAkERoAIQFDZGIgESIAESsDuBWJefCkCs1yDLTwIkQAIkQAJBEKACEAQ0RiEBEiABErA7gdiXnwpA7NchJSABEiABvwg4zp0QyTjvV1gGin8CVADiv44pIQmQAAlkEjhzTBxUADJZZPNvPESnAhAPtUgZSIAESIAESCBAAlQAAgTG4CRAAiRAAnYnEB/yUwGIj3qkFCRA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" alt="Logo Universitas Nusa Mandiri">
  </div>
  <div class="hero-content">
    <div class="hero-title">AI Smart HPP Calculator</div>
    <div class="hero-subtitle">Tugas Akhir Kecerdasan Buatan — Kelompok 9</div>
    <div class="hero-details">
      <div><b>Kelas:</b> 12.7C.12</div>
      <div><b>Anggota Kelompok:</b></div>
      <div>Sandro Arnesto (12260084) &nbsp; • &nbsp; Syauqi (12260140) &nbsp; • &nbsp; Fakhri Husaini (12260120)</div>
      <div class="hero-university">Universitas Nusa Mandiri</div>
    </div>
  </div>
</div>
<p class="intro-text">Masukkan informasi produk dan komponen biaya. Aplikasi akan menghitung HPP per unit, membandingkannya dengan target, lalu menampilkan rekomendasi.</p>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Panduan Penggunaan")
    st.markdown(
        "1. Isi informasi produk.\n"
        "2. Masukkan bahan/komponen biaya.\n"
        "3. Masukkan Addcost, Pajak, dan Kemasan.\n"
        "4. Klik **Hitung HPP dan Analisis**, lalu simpan produk ke rekap.\n"
        "5. Tambahkan produk berikutnya dan unduh rekap Excel/PDF."
    )
    st.divider()
    st.caption("Basis perhitungan: seluruh biaya dikonversi menjadi biaya per satu unit produk.")

# ---------------------------
# Inputs and multi-product workflow
# ---------------------------
if "saved_products" not in st.session_state:
    st.session_state.saved_products = []
if "product_index" not in st.session_state:
    st.session_state.product_index = 1
if "hpp_result" not in st.session_state:
    st.session_state.hpp_result = None

idx = st.session_state.product_index
st.header(f"1. Informasi Produk — Produk {idx}")
c1, c2, c3 = st.columns([2, 1, 1])
with c1:
    nama_produk = st.text_input("Nama produk", placeholder="Contoh: Chicken Rice Bowl", key=f"name_{idx}")
with c2:
    kategori = st.selectbox("Kategori", ["Makanan & Minuman", "Retail", "Manufaktur", "Jasa", "Lainnya"], key=f"category_{idx}")
with c3:
    satuan_produk = st.text_input("Satuan produk", value="unit", key=f"unit_{idx}", help="Contoh: porsi, pcs, botol, paket, jam")
c4, c5, c6 = st.columns(3)
with c4:
    target_hpp = st.number_input("Target HPP (%)", min_value=1.0, max_value=100.0, value=40.0, step=1.0, key=f"target_{idx}")
with c5:
    harga_jual = st.number_input("Harga jual saat ini (Rp/unit)", min_value=0, value=30000, step=500, key=f"price_{idx}")
with c6:
    pembulatan = st.selectbox("Pembulatan harga rekomendasi", [100, 500, 1000], index=1, format_func=lambda x: f"Ke atas Rp{x:,}".replace(",", "."), key=f"round_{idx}")

st.divider()
st.header("2. Bahan atau Komponen Biaya")
st.caption("Masukkan harga pembelian, isi kemasan, dan pemakaian untuk satu unit produk. Satuan massa/volume dapat dikonversi otomatis.")
jumlah_bahan = st.number_input("Jumlah bahan/komponen", min_value=1, max_value=30, value=3, step=1, key=f"count_{idx}")
materials = []
for i in range(int(jumlah_bahan)):
    with st.container(border=True):
        st.markdown(f"**Komponen {i + 1}**")
        a,b,c,d=st.columns([2,1.2,1,1])
        with a: comp_name=st.text_input("Nama komponen", placeholder="Contoh: Beras", key=f"comp_{idx}_{i}")
        with b: purchase_price=st.number_input("Harga beli (Rp)", min_value=0, value=15000, step=500, key=f"purchase_{idx}_{i}")
        with c: pack_size=st.number_input("Isi kemasan", min_value=0.0, value=1000.0, step=10.0, key=f"pack_{idx}_{i}")
        with d: usage=st.number_input("Pemakaian/unit", min_value=0.0, value=100.0, step=1.0, key=f"usage_{idx}_{i}")
        unit_options={"mg":("massa",0.001),"gram":("massa",1.0),"kg":("massa",1000.0),"ml":("volume",1.0),"liter":("volume",1000.0),"pcs":("jumlah",1.0)}
        u1,u2=st.columns(2)
        with u1: pack_unit=st.selectbox("Satuan isi kemasan",list(unit_options),index=1,key=f"pack_unit_{idx}_{i}")
        with u2: usage_unit=st.selectbox("Satuan pemakaian",list(unit_options),index=1,key=f"usage_unit_{idx}_{i}")
        same_dimension=unit_options[pack_unit][0]==unit_options[usage_unit][0]
        if same_dimension and pack_size>0:
            usage_in_pack_unit=usage*unit_options[usage_unit][1]/unit_options[pack_unit][1]
            cost=purchase_price/pack_size*usage_in_pack_unit
            st.caption(f"Biaya komponen per {satuan_produk or 'unit'}: **{rupiah(cost)}** (pemakaian setara {usage_in_pack_unit:g} {pack_unit})")
        else:
            usage_in_pack_unit=0; cost=0
            st.warning("Satuan isi kemasan dan pemakaian harus sejenis.")
        materials.append({"Komponen":comp_name.strip(),"Harga beli (Rp)":float(purchase_price),"Isi kemasan":float(pack_size),"Satuan isi kemasan":pack_unit,"Pemakaian per unit":float(usage),"Satuan pemakaian":usage_unit,"Pemakaian (dikonversi)":float(usage_in_pack_unit),"Biaya per unit (Rp)":float(cost),"Satuan valid":bool(same_dimension)})

st.divider()
st.header("3. Biaya Tambahan")
st.caption("Addcost dan pajak dihitung dari total biaya bahan/komponen. Kemasan ditambahkan sebagai nominal per unit.")
a1,a2,a3=st.columns(3)
with a1: addcost_pct=st.number_input("Addcost (%)",min_value=0.0,max_value=1000.0,value=0.0,step=1.0,key=f"addcost_{idx}")
with a2: pajak_pct=st.number_input("Pajak (%)",min_value=0.0,max_value=1000.0,value=0.0,step=1.0,key=f"tax_{idx}")
with a3: kemasan_rp=st.number_input("Kemasan (Rp/unit)",min_value=0,value=0,step=500,key=f"packcost_{idx}")

calculate=st.button("Hitung HPP dan Analisis",type="primary",use_container_width=True,key=f"calculate_{idx}")
if calculate:
    errors=[]
    if not nama_produk.strip(): errors.append("Nama produk wajib diisi.")
    if not satuan_produk.strip(): errors.append("Satuan produk wajib diisi.")
    if harga_jual<=0: errors.append("Harga jual saat ini harus lebih besar dari nol.")
    for n,m in enumerate(materials,1):
        if not m["Komponen"]: errors.append(f"Nama komponen pada baris {n} belum diisi.")
        if m["Harga beli (Rp)"]<=0: errors.append(f"Harga beli komponen {n} harus lebih besar dari nol.")
        if m["Isi kemasan"]<=0: errors.append(f"Isi kemasan komponen {n} harus lebih besar dari nol.")
        if m["Pemakaian per unit"]<=0: errors.append(f"Pemakaian komponen {n} harus lebih besar dari nol.")
        if not m["Satuan valid"]: errors.append(f"Satuan komponen {n} tidak sejenis.")
    if errors:
        st.error("Data belum dapat dihitung. Perbaiki hal berikut:")
        for error in errors: st.write(f"- {error}")
        st.session_state.hpp_result=None
    else:
        material_total=sum(m["Biaya per unit (Rp)"] for m in materials)
        addcost_value=material_total*addcost_pct/100
        tax_value=material_total*pajak_pct/100
        hpp_total=material_total+addcost_value+tax_value+float(kemasan_rp)
        if hpp_total<=0:
            st.error("Total HPP harus lebih besar dari nol."); st.session_state.hpp_result=None
        else:
            actual_pct=hpp_total/harga_jual*100
            recommended_raw=hpp_total/(target_hpp/100)
            recommended_price=round_up(recommended_raw,int(pembulatan))
            inference=classify_hpp(actual_pct,target_hpp)
            costs=[{"Jenis biaya":"Addcost", "Persentase (%)":addcost_pct,"Biaya per unit (Rp)":addcost_value},{"Jenis biaya":"Pajak","Persentase (%)":pajak_pct,"Biaya per unit (Rp)":tax_value},{"Jenis biaya":"Kemasan","Persentase (%)":None,"Biaya per unit (Rp)":float(kemasan_rp)}]
            report_data={"Nama produk":nama_produk.strip(),"Kategori":kategori,"Satuan produk":satuan_produk.strip(),"Target HPP (%)":target_hpp,"Harga jual saat ini (Rp)":rupiah(harga_jual),"Total biaya bahan/komponen (Rp)":rupiah(material_total),"Addcost (%)":addcost_pct,"Nilai Addcost (Rp)":rupiah(addcost_value),"Pajak (%)":pajak_pct,"Nilai Pajak (Rp)":rupiah(tax_value),"Kemasan (Rp/unit)":rupiah(kemasan_rp),"Total HPP per unit (Rp)":rupiah(hpp_total),"HPP Aktual (%)":actual_pct,"Harga jual rekomendasi (Rp)":rupiah(recommended_price)}
            new_result={"product_id":idx,"data":report_data,"materials":materials,"costs":costs,"inference":inference,"hpp_total":hpp_total,"material_total":material_total,"extra_total":addcost_value+tax_value+float(kemasan_rp),"actual_pct":actual_pct,"recommended_raw":recommended_raw,"recommended_price":recommended_price,"target_hpp":target_hpp,"nama_produk":nama_produk.strip(),"satuan_produk":satuan_produk.strip()}
            st.session_state.hpp_result=new_result

            # Simpan otomatis setiap kali perhitungan berhasil.
            # Jika produk pada nomor yang sama dihitung ulang, perbarui datanya
            # agar tidak terjadi duplikasi.
            saved_list=st.session_state.saved_products
            existing_index=next((j for j,item in enumerate(saved_list) if item.get("product_id")==idx), None)
            if existing_index is None:
                saved_list.append(new_result.copy())
            else:
                saved_list[existing_index]=new_result.copy()

# Display current calculation and save it into the multi-product recap.
result=st.session_state.get("hpp_result")
if result:
    st.divider(); st.header("4. Hasil Perhitungan")
    m1,m2,m3,m4=st.columns(4)
    m1.metric("Total HPP per unit",rupiah(result["hpp_total"]))
    m2.metric("HPP Aktual",f'{result["actual_pct"]:.2f}%')
    m3.metric("Target HPP",f'{result["target_hpp"]:.2f}%')
    m4.metric("Harga jual rekomendasi",rupiah(result["recommended_price"]))
    st.info(f'Harga rekomendasi dihitung dari HPP per unit ÷ target HPP. Nilai sebelum pembulatan: {rupiah(result["recommended_raw"])}; dibulatkan ke atas menjadi {rupiah(result["recommended_price"])}.')
    st.subheader("5. Rincian Perhitungan Biaya")
    material_display=pd.DataFrame(result["materials"]).drop(columns=["Satuan valid"])
    material_display["Harga beli (Rp)"]=material_display["Harga beli (Rp)"].map(rupiah); material_display["Biaya per unit (Rp)"]=material_display["Biaya per unit (Rp)"].map(rupiah)
    st.dataframe(style_table_alignment(material_display),use_container_width=True,hide_index=True)
    cost_display=pd.DataFrame(result["costs"]); cost_display["Biaya per unit (Rp)"]=cost_display["Biaya per unit (Rp)"].map(rupiah)
    st.dataframe(style_table_alignment(cost_display),use_container_width=True,hide_index=True)
    st.write(f'Total bahan/komponen: **{rupiah(result["material_total"])}**')
    st.write(f'Total biaya tambahan: **{rupiah(result["extra_total"])}**')
    st.write(f'Total HPP: **{rupiah(result["hpp_total"])}**')
    st.divider(); st.header("6. Hasil Analisis HPP")
    inf=result["inference"]; st.success(f'Hasil analisis: {inf["status"]}')
    st.markdown("**Data Perhitungan**")
    st.write(f'- HPP aktual = {result["actual_pct"]:.2f}%'); st.write(f'- Target HPP = {result["target_hpp"]:.2f}%')
    st.markdown("**Proses Analisis**")
    st.write(f'Kriteria yang terpenuhi: **{inf["condition"]}**.'); st.write(f'Kesimpulan: **{inf["status"]}**.'); st.write(f'Rekomendasi: {inf["recommendation"]}')
    st.success(f'Produk **{result["nama_produk"]}** sudah tersimpan otomatis ke rekap.')

# Recap all products saved during this app session.
st.divider(); st.header("7. Rekap Perhitungan Produk")
saved=st.session_state.saved_products
if saved:
    recap=[]
    for no,item in enumerate(saved,1):
        d=item["data"]
        recap.append({"No":no,"Nama produk":d["Nama produk"],"Kategori":d["Kategori"],"Total HPP/unit":d["Total HPP per unit (Rp)"],"HPP Aktual":f'{d["HPP Aktual (%)"]:.2f}%',"Harga jual":d["Harga jual saat ini (Rp)"],"Harga rekomendasi":d["Harga jual rekomendasi (Rp)"]})
    st.dataframe(style_table_alignment(pd.DataFrame(recap)),use_container_width=True,hide_index=True)
    st.caption(f"Total produk tersimpan: {len(saved)}. Data tersimpan selama sesi aplikasi ini masih berjalan.")
    b1,b2=st.columns(2)
    with b1:
        try:
            st.download_button("📊 Unduh Rekap Semua Produk (Excel)",data=make_multi_excel(saved),file_name="Rekap_HPP_Seluruh_Produk.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        except Exception as exc: st.warning(f"Ekspor Excel belum tersedia: {exc}")
    with b2:
        try:
            st.download_button("📄 Unduh Rekap Semua Produk (PDF)",data=make_multi_pdf(saved),file_name="Rekap_HPP_Seluruh_Produk.pdf",mime="application/pdf",use_container_width=True)
        except ImportError: st.warning("Paket reportlab belum terpasang. Jalankan: pip install reportlab")
        except Exception as exc: st.warning(f"Ekspor PDF belum tersedia: {exc}")
    if st.button("Tambah Produk Berikutnya",type="primary",use_container_width=True):
        st.session_state.product_index += 1
        st.session_state.hpp_result=None
        st.rerun()
    if st.button("Hapus Seluruh Rekap",type="secondary"):
        st.session_state.saved_products=[]; st.session_state.hpp_result=None; st.session_state.product_index=1; st.rerun()
else:
    st.info("Belum ada produk yang disimpan. Hitung produk pertama, lalu klik 'Simpan Produk ke Rekap'.")

st.divider()
st.caption("AI Smart HPP Calculator | Hasil perhitungan bergantung pada kelengkapan dan ketepatan data yang dimasukkan.")
