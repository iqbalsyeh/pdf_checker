import os
import pytesseract
from collections import defaultdict
import pandas as pd
import time
import streamlit as st
from PIL import Image, ImageOps, ImageFilter
import fitz  # PyMuPDF
from io import BytesIO

# Kata kunci dokumen
keywords = [
    'SURAT PERINTAH MEMBAYAR', 
    'SURAT PERINTAH PEMBAYARAN', 
    'DAFTAR SP2D SATKER',
    'SURAT PERMINTAAN PEMBAYARAN', 
    'SURAT TUGAS',
    'SURAT PERJALANAN DINAS', 
    'BERITA ACARA SERAH TERIMA', 
    'BERITA ACARA PEMBAYARAN'
]

# Preprocessing gambar
def preprocess_image(image):
    gray = ImageOps.grayscale(image)
    sharpened = gray.filter(ImageFilter.SHARPEN)
    bw = sharpened.point(lambda x: 0 if x < 180 else 255, '1')
    return bw

# Proses satu file PDF
def process_pdf_from_bytes(file_bytes, progress_bar=None, idx=0, total=1):
    results = defaultdict(bool)
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        pages = len(doc)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            processed_image = preprocess_image(img)
            text = pytesseract.image_to_string(processed_image)
            for keyword in keywords:
                if keyword in text:
                    results[keyword] = True
            if progress_bar:
                progress = ((idx + i / pages) / total)
                progress_bar.progress(min(progress, 1.0))
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
    return results

# ========== Tampilan Web ==========
st.set_page_config(page_title="PDF Checker", layout="wide")
st.markdown("<h1 style='text-align: center; color: darkblue;'>📄 Pemeriksa Kelengkapan Dokumen PDF Scan</h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📤 Unggah file PDF (boleh lebih dari satu)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    start = time.time()
    st.info("⏳ Memproses dokumen... Mohon tunggu.")
    progress_bar = st.progress(0)
    summary = {}

    for idx, uploaded_file in enumerate(uploaded_files):
        file_bytes = uploaded_file.read()
        result = process_pdf_from_bytes(file_bytes, progress_bar, idx, len(uploaded_files))
        summary[uploaded_file.name] = result

    progress_bar.progress(1.0)
    st.success("✅ Pemeriksaan selesai!")

    # ===== Ringkasan dan Tabel =====
    data = []
    jumlah_lengkap = 0
    jumlah_tidak_lengkap = 0

    for filename, result in summary.items():
        row = {'Nama File': filename}
        lengkap = True
        for keyword in keywords:
            ada = result[keyword]
            row[keyword] = '✅' if ada else '❌'
            if not ada:
                lengkap = False
        row['Status Dokumen'] = 'Lengkap' if lengkap else 'Tidak Lengkap'
        if lengkap:
            jumlah_lengkap += 1
        else:
            jumlah_tidak_lengkap += 1
        data.append(row)

    df = pd.DataFrame(data)

    st.subheader("📊 Hasil Pemeriksaan")
    st.dataframe(df.style.applymap(
        lambda val: 'color: green;' if val == '✅' else ('color: red;' if val == '❌' else ''),
        subset=keywords
    ))

    rekap_df = pd.DataFrame({
        'Keterangan': ['Jumlah Dokumen', 'Dokumen Lengkap', 'Dokumen Tidak Lengkap'],
        'Jumlah': [len(uploaded_files), jumlah_lengkap, jumlah_tidak_lengkap]
    })

    st.subheader("📋 Ringkasan")
    st.table(rekap_df)

    # Simpan Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Detail Pemeriksaan', index=False)
        rekap_df.to_excel(writer, sheet_name='Rekapitulasi', index=False)
    output.seek(0)

    st.download_button(
        label="📥 Unduh Hasil Excel",
        data=output,
        file_name="hasil_pemeriksaan_dokumen.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    durasi = time.time() - start
    st.caption(f"⏱️ Waktu proses: {durasi:.2f} detik")

# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>🔍 Aplikasi dibuat untuk membantu analisis dokumen pemeriksaan secara cepat dan efisien</p>", unsafe_allow_html=True)
