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

# Preprocessing gambar
def preprocess_image(image):
    gray = ImageOps.grayscale(image)
    sharpened = gray.filter(ImageFilter.SHARPEN)
    bw = sharpened.point(lambda x: 0 if x < 180 else 255, '1')
    return bw

# Proses satu file PDF
def process_pdf_from_bytes(file_bytes, filename, progress_bar, idx, total, waktu_mulai_global, total_halaman_seluruhnya, halaman_diproses, info_area, caption_area):
    results = defaultdict(bool)
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        pages = len(doc)
        for i, page in enumerate(doc):
            # Update info proses
            info_area.info(f"🔍 Memeriksa file: **{filename}** | Halaman: {i+1} dari {pages}")

            # Proses halaman
            pix = page.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            processed_image = preprocess_image(img)
            text = pytesseract.image_to_string(processed_image)

            for keyword in keywords:
                if keyword in text:
                    results[keyword] = True

            # Update progress
            halaman_diproses += 1
            elapsed = time.time() - waktu_mulai_global
            rata2_per_halaman = elapsed / halaman_diproses
            estimasi_sisa = rata2_per_halaman * (total_halaman_seluruhnya - halaman_diproses)

            progress = halaman_diproses / total_halaman_seluruhnya
            progress_bar.progress(min(progress, 1.0))
            caption_area.caption(f"📄 Halaman {halaman_diproses} dari {total_halaman_seluruhnya} | ⏳ Estimasi sisa: {estimasi_sisa:.1f} detik")
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
    return results, halaman_diproses

# ========== Tampilan Web ==========
st.set_page_config(page_title="PDF Checker", layout="wide")
st.markdown("<h1 style='text-align: center; color: darkblue;'>📄 Pemeriksa Kelengkapan Dokumen PDF Scan</h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📤 Unggah file PDF (boleh lebih dari satu)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    waktu_mulai_global = time.time()
    st.info("⏳ Menghitung total halaman...")
    
    # Hitung total halaman
    total_halaman = 0
    file_byte_cache = {}  # Cache file agar bisa dibaca dua kali
    for file in uploaded_files:
        bytes_ = file.read()
        file_byte_cache[file.name] = bytes_
        try:
            doc = fitz.open(stream=bytes_, filetype="pdf")
            total_halaman += len(doc)
        except:
            pass

    st.success(f"📚 Total halaman seluruh dokumen: {total_halaman}")
    st.markdown("<hr>", unsafe_allow_html=True)

    # Inisialisasi area
    progress_bar = st.progress(0)
    info_area = st.empty()
    caption_area = st.empty()
    summary = {}
    halaman_diproses = 0

    # Mulai proses file
    for idx, uploaded_file in enumerate(uploaded_files):
        file_bytes = file_byte_cache[uploaded_file.name]
        result, halaman_diproses = process_pdf_from_bytes(
            file_bytes, uploaded_file.name,
            progress_bar, idx, len(uploaded_files),
            waktu_mulai_global, total_halaman, halaman_diproses,
            info_area, caption_area
        )
        summary[uploaded_file.name] = result

    progress_bar.progress(1.0)
    info_area.success("✅ Pemeriksaan selesai!")
    caption_area.caption("✅ Semua halaman telah diperiksa.")

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

    durasi = time.time() - waktu_mulai_global
    st.caption(f"⏱️ Total waktu proses: {durasi:.2f} detik")

# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>🔍 Aplikasi dibuat untuk membantu analisis dokumen pemeriksaan secara cepat dan efisien</p>", unsafe_allow_html=True)
