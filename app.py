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
    'MENUGASKAN', 'MEMBERI TUGAS',
    'SURAT PERJALANAN DINAS',
    'BERITA ACARA SERAH TERIMA',
    'BERITA ACARA PEMBAYARAN'
]

ocr_cache = {}  # Global cache OCR

# Preprocessing gambar
def preprocess_image(image):
    gray = ImageOps.grayscale(image)
    sharpened = gray.filter(ImageFilter.SHARPEN)
    bw = sharpened.point(lambda x: 0 if x < 180 else 255, '1')
    return bw

# Proses satu file PDF
def process_pdf_from_bytes(file_bytes, progress_bar=None, idx=0, total_files=1,
                           status_area=None, est_time_area=None,
                           total_pages_all=1, start_time=0, filename="",
                           dpi_setting=100):
    results = defaultdict(bool)
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        pages = len(doc)
        for i, page in enumerate(doc):
            key = f"{filename}_page_{i}"
            if key in ocr_cache:
                text = ocr_cache[key]
            else:
                pix = page.get_pixmap(dpi=dpi_setting)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                processed_image = preprocess_image(img)
                text = pytesseract.image_to_string(processed_image)
                ocr_cache[key] = text

            for keyword in keywords:
                if keyword in text:
                    results[keyword] = True

            current_page = sum([len(fitz.open(stream=f['data'], filetype="pdf")) for f in uploaded_files[:idx]]) + i + 1
            elapsed = time.time() - start_time
            progress = current_page / total_pages_all
            est_total = elapsed / progress if progress > 0 else 0
            remaining = est_total - elapsed

            if progress_bar:
                progress_bar.progress(min(progress, 1.0))
            if status_area:
                status_area.markdown(f"📄 Memeriksa **{filename}**, halaman **{i + 1}/{pages}**")
            if est_time_area:
                est_time_area.markdown(f"⏳ Estimasi selesai: **{remaining:.1f} detik lagi**")
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
    return results

# ========== Tampilan Web ==========
st.set_page_config(page_title="PDF Checker", layout="wide")

# Logo dan Judul
col1, col2 = st.columns([1, 5])
with col1:
    st.image("logo_bpk.png", width=100)
with col2:
    st.markdown("<h1 style='color: darkblue;'>📄 Pemeriksa Kelengkapan Dokumen PDF Scan</h1>", unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# Opsi Cepat
fast_mode = st.sidebar.checkbox("⚡ Mode cepat (turunkan kualitas scan)", value=True)
dpi_setting = 100 if fast_mode else 200

# Upload File
uploaded_streams = st.file_uploader("📤 Unggah file PDF (boleh lebih dari satu)", type="pdf", accept_multiple_files=True)

if uploaded_streams:
    start = time.time()
    st.info("⏳ Memproses dokumen... Mohon tunggu.")

    uploaded_files = []
    total_pages = 0
    for f in uploaded_streams:
        f_bytes = f.read()
        try:
            doc = fitz.open(stream=f_bytes, filetype="pdf")
            total_pages += len(doc)
        except:
            continue
        uploaded_files.append({'name': f.name, 'data': f_bytes})

    progress_bar = st.progress(0)
    status_area = st.empty()
    est_time_area = st.empty()
    summary = {}

    for idx, file in enumerate(uploaded_files):
        result = process_pdf_from_bytes(
            file['data'],
            progress_bar,
            idx,
            len(uploaded_files),
            status_area,
            est_time_area,
            total_pages,
            start_time=start,
            filename=file['name'],
            dpi_setting=dpi_setting
        )
        summary[file['name']] = result

    progress_bar.progress(1.0)
    status_area.markdown("✅ Pemeriksaan dokumen selesai!")
    est_time_area.empty()

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
