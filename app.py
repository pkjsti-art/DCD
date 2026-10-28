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
# INJEKSI TEMA UI/UX: CYBER-SIDEBAR & GLASSMORPHISM 3D
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
    
    /* 
       ======================================
       MODIFIKASI SIDEBAR Keren & Selaras (Glassmorphism 3D)
       ====================================== 
    */
    [data-testid="stSidebar"] {{
        background: rgba(10, 10, 25, 0.65) !important;
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
        border-right: 1px solid rgba(0, 243, 255, 0.3) !important;
        box-shadow: 
            5px 0px 30px rgba(127, 0, 255, 0.25),
            inset -1px 0px 0px 0px rgba(255, 255, 255, 0.1);
    }}

    /* Styling Teks di dalam Sidebar */
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {{
        color: #00f3ff !important;
        font-family: 'Courier New', Courier, monospace;
        letter-spacing: 1px;
    }}

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
        padding-top: 2rem !important;
    }}

    /* 
       ======================================
       EFEK KACA 3D (GLASSMORPHISM HEADER CARD)
       ====================================== 
    */
    .glass-header-card {{
        background: rgba(15, 15, 35, 0.45);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border-radius: 16px;
        border: 1px solid rgba(0, 243, 255, 0.25);
        padding: 25px 20px;
        margin-bottom: 35px;
        box-shadow: 
            0 8px 32px 0 rgba(0, 0, 0, 0.5),
            inset 0 1px 1px 0 rgba(255, 255, 255, 0.2),
            0 0 20px rgba(0, 243, 255, 0.15);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        width: 100%;
    }}

    /* 
       ======================================
       TOMBOL KENDALI (INDUSTRIAL NEON)
       ====================================== 
    */
    .stButton > button, [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(45deg, #0d0d26, #1a0033) !important;
        color: #00f3ff !important;
        border: 1px solid #00f3ff !important;
        border-radius: 4px !important; 
        padding: 10px 24px !important;
        font-weight: 900 !important;
        letter-spacing: 1px;
        text-transform: uppercase;
        box-shadow: 0 0 10px rgba(0, 243, 255, 0.3) inset, 0 0 10px rgba(0, 243, 255, 0.3) !important;
        transition: all 0.3s ease-in-out;
    }}
    
    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {{
        background: linear-gradient(45deg, #0033aa, #7f00ff) !important;
        border: 1px solid #00f3ff !important;
        box-shadow: 0 0 20px rgba(0, 243, 255, 0.8), 0 0 30px rgba(127, 0, 255, 0.6) !important;
        color: #ffffff !important;
    }}

    /* 
       ======================================
       AREA UPLOAD FILE (HUD SCANNER)
       ====================================== 
    */
    [data-testid="stFileUploadDropzone"] {{
        background-color: rgba(5, 5, 20, 0.75) !important;
        border: 1px dashed #7f00ff !important;
        border-radius: 4px !important;
        box-shadow: 0 0 15px rgba(127, 0, 255, 0.1) inset;
        backdrop-filter: blur(5px);
        transition: all 0.3s ease;
    }}
    
    [data-testid="stFileUploadDropzone"]:hover {{
        background-color: rgba(15, 5, 40, 0.9) !important;
        border: 1px dashed #00f3ff !important;
        box-shadow: 0 0 20px rgba(0, 243, 255, 0.3) inset, 0 0 15px rgba(0, 243, 255, 0.3);
    }}
    
    [data-testid="stFileUploadDropzone"] * {{
        color: #00f3ff !important;
    }}

    /* Garis pembatas (Divider) */
    hr {{
        border-top: 1px solid #7f00ff !important;
        box-shadow: 0px 0px 10px #7f00ff;
    }}

    /* Tabel Dataframe (Monitor Display) */
    [data-testid="stDataFrame"] {{
        border: 1px solid #00f3ff;
        box-shadow: 0 0 15px rgba(0, 243, 255, 0.2);
        border-radius: 4px;
        background-color: rgba(0,0,0,0.7);
    }}
</style>
""", unsafe_allow_html=True)
# ==========================================

# HEADER DENGAN FRAME KACA 3D & PENYESUAIAN FONT-SIZE SUBTITLE AGAR PRESISI
st.markdown("""
    <div class="glass-header-card">
        <h1 style="color: #ffffff; font-family: 'Courier New', Courier, monospace; font-weight: 900; letter-spacing: 3px; font-size: 2.5rem; text-transform: uppercase; margin: 0; padding: 0; text-shadow: 0 0 10px #00f3ff, 0 0 20px #00f3ff, 0 0 40px #7f00ff, 0 0 80px #7f00ff; text-align: center;">DATA CONTROL DYEING (DCD)</h1>
        <p style="color: #b3b3ff; font-weight: bold; letter-spacing: 1.2px; font-size: 1.18rem; margin-top: 12px; margin-bottom: 0; padding: 0; text-align: center;">Sistem pengecekan dan rekonsiliasi data produksi secara otomatis dan terintegrasi.</p>
    </div>
""", unsafe_allow_html=True)

# 1. Inisialisasi Session State agar data tetap tersimpan (STAY) meskipun tombol unduh diklik
if 'df_hasil_crosscheck' not in st.session_state:
    st.session_state['df_hasil_crosscheck'] = None
if 'df_pibc_bersih' not in st.session_state:
    st.session_state['df_pibc_bersih'] = None
if 'df_pg_bersih' not in st.session_state:
    st.session_state['Tulisan **`keyboard_double`** tersebut muncul karena itu adalah nama *icon* bawaan dari Material Design / Streamlit yang terpanggil di dalam kode `st.sidebar.header("Navigasi Menu")
