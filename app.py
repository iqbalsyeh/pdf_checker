import os
import pytesseract
from collections import defaultdict
import pandas as pd
import time
import streamlit as st

from PIL import Image, ImageOps, ImageFilter
import fitz  # PyMuPDF

# === Kata kunci dokumen ===
keywords = [
    'SURAT PERINTAH MEMBAYAR', 
    'SURAT PERINTAH PEMBAYARAN', 
    'DAFTAR SP2D SATKER',
    'SURAT PERMINTAAN PEMBAYARAN', 
    'MEMBERI TUGAS', 'MENUGASKAN', 
    'SURAT PERJALANAN DINAS', 'BERITA ACARA SERAH TERIMA', 'BERITA ACARA PEMBAYARAN'
]

# === Preprocessing gambar ===
def preprocess_image(image):
    gray = ImageOps.grayscale(image)
    sharpened = gray.filter(ImageFilter.SHARPEN)
    bw = sharpened.point(lambda x: 0 if x < 180 else 255, '1')
    return bw

# === Proses satu file PDF (dari bytes) dengan progress bar dan estimasi waktu ===
def process_pdf_from_bytes(file_bytes, file_index, total_files):
    results = defaultdict(bool)
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        total_pages = len(doc)

        progress_text = st.empty()
        progress_bar = st.progress(0)

        start_time = time.time()

        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            processed_image = preprocess_image(img)
            text = pytesseract.image_to_string(processed_image)

            for keyword in keywords:
                if keyword in text:
                    results[keyword] = True

            # Estimasi waktu
            elapsed = time.time() - start_time
            avg_time = elapsed / (i + 1)
            estimated_total = avg_time * total_pages
            remaining = estimated_total - elapsed

            # Update progress
            progress = (i + 1) / total_pages
            progress_bar.progress(progress)
            progress_text.text(
                f"📄 File {file_index + 1}/{total_files} - Halaman {i + 1}/{total_pages} "
                f"(estimasi sisa: {remaining:.1f} detik)"
            )

        progress_bar.empty()
        progress_text.text(f"✅ Selesai memproses file {file_index + 1}/{total_files}")
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
    return results

# === Tampilan Web ===
st.title("📄 Analisis Kelengkapan Dokumen PDF Scan")
uploaded_files = st.file_uploader("Unggah file PDF (lebih dari satu diperbolehkan)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    start = time.time()
    summary = {}
    for i, uploaded_file in enumerate(uploaded_files):
        file_bytes = uploaded_file.read()
        result = process_pdf_from_bytes(file_bytes, i, len(uploaded_files))
        summary[uploaded_file.name] = result

    # === Tabel dan rekap ===
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

    st.success(f"✔️ Dokumen lengkap: {jumlah_lengkap}")
    st.warning(f"❌ Dokumen tidak lengkap: {jumlah_tidak_lengkap}")

    rekap_df = pd.DataFrame({
        'Keterangan': ['Jumlah Dokumen', 'Dokumen Lengkap', 'Dokumen Tidak Lengkap'],
        'Jumlah': [len(uploaded_files), jumlah_lengkap, jumlah_tidak_lengkap]
    })

    # === Excel output
    from io import BytesIO
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

    st.caption(f"⏱️ Waktu proses: {time.time() - start:.2f} detik")
