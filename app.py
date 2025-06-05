import os
import pytesseract
from collections import defaultdict
import pandas as pd
import time
import streamlit as st
from PIL import Image, ImageOps
import fitz  # PyMuPDF
from io import BytesIO
from cryptography.fernet import Fernet
from multiprocessing import Pool

# ======== Enkripsi Setup ==========
KEY = os.environ.get("FERNET_KEY") or Fernet.generate_key()
cipher = Fernet(KEY)

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

def preprocess_image(image):
    gray = ImageOps.grayscale(image)
    bw = gray.point(lambda x: 0 if x < 180 else 255, '1')
    return bw

def ocr_on_page(page):
    pix = page.get_pixmap(dpi=150)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    processed_image = preprocess_image(img)
    text = pytesseract.image_to_string(processed_image, config='--psm 6')
    found_keywords = {k: (k in text) for k in keywords}
    return found_keywords

def process_pdf_from_bytes_parallel(file_bytes):
    results = defaultdict(bool)
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    with Pool() as pool:
        results_per_page = pool.map(ocr_on_page, doc)
    for page_result in results_per_page:
        for k, v in page_result.items():
            if v:
                results[k] = True
    return results

def process_pdf_from_bytes(file_bytes):
    # fallback single process jika multiprocessing error
    try:
        return process_pdf_from_bytes_parallel(file_bytes)
    except Exception:
        results = defaultdict(bool)
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            processed_image = preprocess_image(img)
            text = pytesseract.image_to_string(processed_image, config='--psm 6')
            for keyword in keywords:
                if keyword in text:
                    results[keyword] = True
        return results

# Streamlit UI
st.set_page_config(page_title="PDF Checker Aman dan Cepat", layout="wide")
col1, col2 = st.columns([1, 5])
with col1:
    st.image("logo_bpk.png", width=100)
with col2:
    st.markdown("<h1 style='color: darkblue;'>📄 Pemeriksa Kelengkapan Dokumen - Versi Cepat</h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

uploaded_streams = st.file_uploader("📤 Unggah file PDF", type="pdf", accept_multiple_files=True)

if uploaded_streams:
    start = time.time()
    st.info("🔐 Dokumen Anda sedang diproses secara aman dan cepat...")

    uploaded_files = []
    total_pages = 0
    total_pages_per_file = []

    for f in uploaded_streams:
        f_bytes = f.read()
        try:
            doc = fitz.open(stream=f_bytes, filetype="pdf")
            pages = len(doc)
            total_pages += pages
            total_pages_per_file.append(pages)
        except Exception:
            continue
        encrypted = cipher.encrypt(f_bytes)
        uploaded_files.append({'name': f.name, 'data': encrypted})

    progress_bar = st.progress(0)
    status_area = st.empty()
    est_time_area = st.empty()
    summary = {}

    for idx, file in enumerate(uploaded_files):
        file_bytes = cipher.decrypt(file['data'])
        status_area.markdown(f"📄 Memeriksa **{file['name']}** ...")
        result = process_pdf_from_bytes(file_bytes)
        summary[file['name']] = result
        progress_bar.progress((idx + 1) / len(uploaded_files))

    progress_bar.progress(1.0)
    status_area.markdown("✅ Pemeriksaan dokumen selesai!")
    est_time_area.empty()

    data, lengkap, tidak_lengkap = [], 0, 0
    for filename, result in summary.items():
        row = {'Nama File': filename}
        is_complete = True
        for keyword in keywords:
            ada = result[keyword]
            row[keyword] = '✅' if ada else '❌'
            if not ada:
                is_complete = False
        row['Status Dokumen'] = 'Lengkap' if is_complete else 'Tidak Lengkap'
        if is_complete:
            lengkap += 1
        else:
            tidak_lengkap += 1
        data.append(row)

    df = pd.DataFrame(data)
    st.subheader("📊 Hasil Pemeriksaan")
    st.dataframe(df.style.applymap(
        lambda val: 'color: green;' if val == '✅' else ('color: red;' if val == '❌' else ''),
        subset=keywords
    ))

    rekap_df = pd.DataFrame({
        'Keterangan': ['Jumlah Dokumen', 'Dokumen Lengkap', 'Dokumen Tidak Lengkap'],
        'Jumlah': [len(uploaded_files), lengkap, tidak_lengkap]
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
st.markdown("<p style='text-align: center; color: gray;'>🔐 Dokumen dienkripsi untuk keamanan dan pemrosesan cepat.</p>", unsafe_allow_html=True)
