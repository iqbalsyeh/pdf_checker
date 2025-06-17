# app.py

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

ocr_cache = {}

def preprocess_image(image):
    gray = ImageOps.grayscale(image)
    sharpened = gray.filter(ImageFilter.SHARPEN)
    bw = sharpened.point(lambda x: 0 if x < 180 else 255, '1')
    return bw

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

            current_page = idx + i + 1
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

def load_files_from_folder(folder_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    files = []
    for filename in pdf_files:
        file_path = os.path.join(folder_path, filename)
        with open(file_path, 'rb') as f:
            files.append({'name': filename, 'data': f.read()})
    return files

# ===== Tampilan Streamlit =====
st.set_page_config(page_title="PDF Checker Otomatis", layout="wide")

st.title("📄 Pemeriksaan Otomatis Dokumen PDF dari Power Automate")
st.markdown("<hr>", unsafe_allow_html=True)

fast_mode = st.sidebar.checkbox("⚡ Mode cepat (turunkan kualitas scan)", value=True)
dpi_setting = 100 if fast_mode else 200

# Folder sumber file (otomatis masuk)
FOLDER_PATH = "inbox"
uploaded_files = load_files_from_folder(FOLDER_PATH)

if not uploaded_files:
    st.warning("📂 Belum ada file masuk dari Power Automate.")
else:
    start = time.time()
    st.info("⏳ Memproses dokumen...")

    progress_bar = st.progress(0)
    status_area = st.empty()
    est_time_area = st.empty()
    summary = {}

    total_pages = 0
    for f in uploaded_files:
        try:
            doc = fitz.open(stream=f['data'], filetype="pdf")
            total_pages += len(doc)
        except:
            continue

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

    data = []
    lengkap_count, tidak_lengkap_count = 0, 0
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
            lengkap_count += 1
        else:
            tidak_lengkap_count += 1
        data.append(row)

    df = pd.DataFrame(data)

    st.subheader("📊 Hasil Pemeriksaan")
    st.dataframe(df.style.applymap(
        lambda val: 'color: green;' if val == '✅' else ('color: red;' if val == '❌' else ''),
        subset=keywords
    ))

    rekap_df = pd.DataFrame({
        'Keterangan': ['Jumlah Dokumen', 'Dokumen Lengkap', 'Dokumen Tidak Lengkap'],
        'Jumlah': [len(uploaded_files), lengkap_count, tidak_lengkap_count]
    })

    st.subheader("📋 Ringkasan")
    st.table(rekap_df)

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

st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>🔍 Pemeriksaan otomatis dokumen PDF via OCR dan Power Automate</p>", unsafe_allow_html=True)
