import os
import pytesseract
from collections import defaultdict
import pandas as pd
import time
import streamlit as st

from PIL import Image, ImageOps, ImageFilter
import fitz  # PyMuPDF

# === Konfigurasi awal ===
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# === Kata kunci dokumen ===
keywords = [
    'SURAT PERINTAH MEMBAYAR', 
    'SURAT PERINTAH PEMBAYARAN', 
    'DAFTAR SP2D SATKER',
    'SURAT PERMINTAAN PEMBAYARAN', 
    'SURAT TUGAS',
    'SURAT PERJALANAN DINAS', 'BERITA ACARA SERAH TERIMA', 'BERITA ACARA PEMBAYARAN'
]

# === Preprocessing gambar ===
def preprocess_image(image):
    gray = ImageOps.grayscale(image)
    sharpened = gray.filter(ImageFilter.SHARPEN)
    bw = sharpened.point(lambda x: 0 if x < 180 else 255, '1')
    return bw

# === Proses satu file PDF (dari bytes) ===
def process_pdf_from_bytes(file_bytes):
    results = defaultdict(bool)
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            processed_image = preprocess_image(img)
            text = pytesseract.image_to_string(processed_image)
            for keyword in keywords:
                if keyword in text:
                    results[keyword] = True
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
    return results

# === Tampilan Web ===
st.title("📄 Analisis Kelengkapan Dokumen PDF Scan")
uploaded_files = st.file_uploader("Unggah file PDF (lebih dari satu diperbolehkan)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    start = time.time()
    summary = {}
    for uploaded_file in uploaded_files:
        file_bytes = uploaded_file.read()
        result = process_pdf_from_bytes(file_bytes)
        summary[uploaded_file.name] = result

    # === Siapkan data untuk tabel dan Excel ===
    data = []
    jumlah_lengkap = 0
    jumlah_tidak_lengkap = 0

    for filename, result in summary.items():
        row = {'Nama File': filename}
        lengkap = True
        for keyword in keywords:
            ada = result[keyword]
            row[keyword] = 'ADA' if ada else 'TIDAK ADA'
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
    st.dataframe(df)

    # === Tampilkan ringkasan
    st.success(f"✔️ Dokumen lengkap: {jumlah_lengkap}")
    st.warning(f"❌ Dokumen tidak lengkap: {jumlah_tidak_lengkap}")

    rekap_df = pd.DataFrame({
        'Keterangan': ['Jumlah Dokumen', 'Dokumen Lengkap', 'Dokumen Tidak Lengkap'],
        'Jumlah': [len(uploaded_files), jumlah_lengkap, jumlah_tidak_lengkap]
    })

    # === Simpan ke Excel (in memory)
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Detail Pemeriksaan', index=False)
        rekap_df.to_excel(writer, sheet_name='Rekapitulasi', index=False)
    output.seek(0)

    # === Tombol download hasil
    st.download_button(
        label="📥 Unduh Hasil Excel",
        data=output,
        file_name="hasil_pemeriksaan_dokumen.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.caption(f"⏱️ Waktu proses: {time.time() - start:.2f} detik")
