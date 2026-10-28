import streamlit as st
import pandas as pd
from modules.cleaner import bersihkan_master_mb, bersihkan_master_pibc, bersihkan_master_pg
from modules.processor import (
    siapkan_kolom_standar, 
    jalankan_proses_crosscheck, 
    buat_excel_hasil_keseluruhan,
    buat_excel_rekap_audit,
    buat_excel_investigasi_pibc,
    buat_excel_investigasi_pg
)
import io
import base64

# Konfigurasi Halaman (Harus selalu di paling atas)
st.set_page_config(
    page_title="Data Control Dyeing (DCD)",
    page_icon="🖥️",
    layout="wide"
)

# Fungsi untuk memuat gambar background lokal agar terbaca oleh CSS
def get_base64_of_bin_file(bin_file):
    try:
        with open(bin_file, 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except:
        return ""

# Membaca file bg_datacenter.png secara lokal
bin_str = get_base64_of_bin_file('bg_datacenter.png')
bg_css = f"data:image/png;base64,{bin_str}" if bin_str else ""

# ==========================================
# INJEKSI TEMA UI/UX: STABIL & BERSIH TANPA KORBAN TOMBOL
# ==========================================
st.markdown(f"""
<style>
    /* MENGGUNAKAN BACKGROUND DATACENTER LOKAL */
    [data-testid="stAppViewContainer"] {{
        background-image: 
            linear-gradient(rgba(5, 5, 20, 0.75), rgba(10, 5, 25, 0.85)),
            url("{bg_css}") !important;
        background-size: cover !important;
        background-position: center !important;
        background-repeat: no-repeat !important;
        background-attachment: fixed !important;
    }}
    
    /* Menghilangkan background bawaan Streamlit agar transparan */
    .stApp, [data-testid="stHeader"] {{
        background: transparent !important;
    }}
    
    /* MODIFIKASI SIDEBAR GLASSMORPHISM 3D */
    [data-testid="stSidebar"] {{
        background: rgba(10, 10, 25, 0.65) !important;
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
        border-right: 1px solid rgba(0, 243, 255, 0.3) !important;
        box-shadow: 
            5px 0px 30px rgba(127, 0, 255, 0.25),
            inset -1px 0px 0px 0px rgba(255, 255, 255, 0.1);
    }}

    /* Styling label sidebar tanpa mengubah font ikon Streamlit */
    [data-testid="stSidebar"] label {{
        color: #00f3ff !important;
        font-family: 'Courier New', Courier, monospace;
        letter-spacing: 1px;
    }}

    /* Jangan mengubah font seluruh span/header sidebar.
       Streamlit dapat memakai elemen tersebut untuk merender ikon. */

    /* Kotak Pilihan Menu (Selectbox) di Sidebar agar Senada */
    [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"] > div {{
        background-color: rgba(15, 15, 35, 0.8) !important;
        border: 1px solid rgba(0, 243, 255, 0.4) !important;
        border-radius: 8px !important;
        color: #ffffff !important;
        box-shadow: inset 0 0 10px rgba(0, 243, 255, 0.1);
    }}

    [data-testid="stSidebar"] [data-testid="stSelectbox"] svg {{
        fill: #00f3ff !important;
    }}
    
    /* Warna Teks Umum */
    p, span, label, .stMarkdown, .stText {{
        color: #c9c9e8 !important;
    }}

    /* Mengurangi jarak kosong di bagian atas */
    [data-testid="block-container"] {{
