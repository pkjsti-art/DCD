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

# Konfigurasi Halaman (Harus selalu di paling atas)
st.set_page_config(
    page_title="Data Control Dyeing (DCD)",
    page_icon="🧪",
    layout="wide"
)

# ==========================================
# INJEKSI TEMA UI/UX: CYBERPUNK COLORFUL & CENTERED GLOW
# ==========================================
st.markdown("""
<style>
    /* Latar Belakang Rame (Ada efek semburat cahaya Pink & Cyan) */
    .stApp {
        background-color: #05050A;
        background-image: 
            radial-gradient(circle at 10% 20%, rgba(255, 0, 128, 0.15) 0%, transparent 40%),
            radial-gradient(circle at 90% 80%, rgba(0, 243, 255, 0.15) 0%, transparent 40%),
            linear-gradient(180deg, #05050A 0%, #0a0a14 100%);
        background-attachment: fixed;
    }
    
    /* Modifikasi Sidebar (Transparan dengan Glow) */
    [data-testid="stSidebar"] {
        background-color: rgba(5, 5, 10, 0.8) !important;
        backdrop-filter: blur(10px);
        border-right: 2px solid #ff007f !important;
        box-shadow: 5px 0px 20px rgba(255, 0, 127, 0.4);
    }
    
    /* Warna Teks Umum */
    p, span, label, .stMarkdown {
        color: #d1d1e0 !important;
    }

    /* 
       ======================================
       JUDUL UTAMA (CENTERED & SUPER GLOWING) 
       ====================================== 
    */
    h1 {
        text-align: center !important;
        width: 100% !important;
        color: #ffffff !important;
        font-family: 'Courier New', Courier, monospace;
        font-weight: 900 !important;
        letter-spacing: 4px;
        text-transform: uppercase;
        /* Efek Menyala (Glow) Berlapis */
        text-shadow: 
            0 0 10px #ff007f, 
            0 0 20px #ff007f, 
            0 0 40px #ff007f, 
            0 0 80px #bf00ff !important;
        margin-bottom: 10px;
    }
    
    /* Subjudul / H2 & H3 */
    h2, h3 {
        color: #00f3ff !important;
        text-shadow: 0 0 10px rgba(0, 243, 255, 0.8);
        border-bottom: 1px solid rgba(0, 243, 255, 0.3);
        padding-bottom: 5px;
    }
    
    /* Posisi Deskripsi di bawah judul utama ke tengah */
    [data-testid="stMarkdownContainer"] > p:first-child {
        text-align: center !important;
        color: #00f3ff !important;
        font-weight: bold;
        letter-spacing: 1px;
    }

    /* 
       ======================================
       TOMBOL COLORFUL (GRADIENT & GLOW)
       ====================================== 
    */
    .stButton > button, [data-testid="stDownloadButton"] > button {
        background: linear-gradient(45deg, #ff007f, #7f00ff) !important;
        color: #ffffff !important;
        border: 2px solid #ff007f !important;
        border-radius: 8px !important;
        padding: 10px 24px !important;
        font-weight: 900 !important;
        letter-spacing: 1px;
        box-shadow: 0 0 15px rgba(255, 0, 127, 0.5) !important;
        transition: all 0.3s ease-in-out;
    }
    
    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        background: linear-gradient(45deg, #00f3ff, #007fff) !important;
        border: 2px solid #00f3ff !important;
        box-shadow: 0 0 25px rgba(0, 243, 255, 0.8), 0 0 40px rgba(0, 127, 255, 0.5) !important;
        transform: scale(1.05);
        color: #ffffff !important;
    }

    /* 
       ======================================
       AREA UPLOAD FILE (CYBERPUNK BOX)
       ====================================== 
    */
    [data-testid="stFileUploadDropzone"] {
        background-color: rgba(10, 10, 25, 0.6) !important;
        border: 2px dashed #00f3ff !important;
        border-radius: 12px !important;
        box-shadow: 0 0 15px rgba(0, 243, 255, 0.2) inset, 0 0 15px rgba(0, 243, 255, 0.2);
        transition: all 0.3s ease;
    }
    
    [data-testid="stFileUploadDropzone"]:hover {
        background-color: rgba(30, 10, 40, 0.8) !important;
        border: 2px dashed #ff007f !important;
        box-shadow: 0 0 25px rgba(255, 0, 127, 0.5) inset, 0 0 25px rgba(255, 0, 127, 0.5);
    }
    
    [data-testid="stFileUploadDropzone"] * {
        color: #00f3ff !important;
    }

    /* Garis pembatas (Divider) */
    hr {
        border-top: 2px solid #bf00ff !important;
        box-shadow: 0px 0px 15px #bf00ff;
    }

    /* Tabel Dataframe */
    [data-testid="stDataFrame"] {
        border: 2px solid #7f00ff;
        box-shadow: 0 0 20px rgba(127, 0, 255, 0.4);
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)
# ==========================================

st.title("DATA CONTROL DYEING (DCD)")
st.markdown("Sistem pengecekan dan rekonsiliasi data produksi secara otomatis dan terintegrasi.")

# 1. Inisialisasi Session State agar data tetap tersimpan (STAY) meskipun tombol unduh diklik
if 'df_hasil_crosscheck' not in st.session_state:
    st.session_state['df_hasil_crosscheck'] = None
if 'df_pibc_bersih' not in st.session_state:
    st.session_state['df_pibc_bersih'] = None
if 'df_pg_bersih' not in st.session_state:
    st.session_state['df_pg_bersih'] = None

st.sidebar.header("Navigasi Menu")
menu = st.sidebar.selectbox("Pilih Menu", ["Cross Check Data"])

if menu == "Cross Check Data":
    st.subheader("Upload Dokumen Master Data")
    
    col1, col2 = st.columns(2)
    with col1:
        uploaded_mb = st.file_uploader("Upload File Master MB (Bisa Pilih Banyak)", type=["xlsx", "xls"], accept_multiple_files=True)
        uploaded_pibc = st.file_uploader("Upload File Master PIBC / Dyelot", type=["xlsx", "xls"])
    with col2:
        uploaded_pg = st.file_uploader("Upload File Master PG / Pindah Gudang (Bisa Pilih Banyak)", type=["xlsx", "xls"], accept_multiple_files=True)
        
    if st.button("Proses Data Sekarang", type="primary"):
        if uploaded_mb and uploaded_pibc:
            with st.spinner("Sedang membersihkan dan memproses crosscheck data..."):
                # 1. Bersihkan Data Master
                df_pibc_bersih = bersihkan_master_pibc(uploaded_pibc)
                
                df_pg_bersih = None
                if uploaded_pg:
                    df_pg_bersih = bersihkan_master_pg(uploaded_pg) 
                
                # 2. Proses dan Gabungkan File Master MB
                df_gabung_list = []
                for file_mb in uploaded_mb:
                    df_bersih = bersihkan_master_mb(file_mb)
                    df_siap = siapkan_kolom_standar(df_bersih)
                    df_gabung_list.append(df_siap)
                
                df_final_mb = pd.concat(df_gabung_list, ignore_index=True)
                
                # 3. Jalankan Eksekusi Logika Utama
                df_hasil = jalankan_proses_crosscheck(df_final_mb, df_pibc_bersih, df_pg_bersih)
                
                # SIMPAN HASIL KE SESSION STATE
                st.session_state['df_hasil_crosscheck'] = df_hasil
                st.session_state['df_pibc_bersih'] = df_pibc_bersih
                st.session_state['df_pg_bersih'] = df_pg_bersih
                
                st.success("Proses pengecekan selesai sepenuhnya!")
        else:
            st.warning("Mohon upload setidaknya file Master MB dan Master PIBC terlebih dahulu.")

    # 2. TAMPILAN STAY: Mengambil data dari Session State
    if st.session_state['df_hasil_crosscheck'] is not None:
        df_hasil = st.session_state['df_hasil_crosscheck']
        df_pibc_b = st.session_state['df_pibc_bersih']
        df_pg_b = st.session_state['df_pg_bersih']

        st.markdown("---")
        st.subheader("Tabel Hasil Cross Check Data")
        st.dataframe(df_hasil, use_container_width=True)
        
        st.subheader("Pusat Unduhan (Download Center)")
        col_dl1, col_dl2 = st.columns(2)
        col_dl3, col_dl4 = st.columns(2)
        
        # Generator File Unduhan 1: Keseluruhan Hasil
        data_all = buat_excel_hasil_keseluruhan(df_hasil)
        with col_dl1:
            st.download_button(
                label="📥 Unduh Hasil Cross Check Keseluruhan",
                data=data_all,
                file_name="Hasil_Crosscheck_DCD_Lengkap.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_all"
            )
        
        # Generator File Unduhan 2: Laporan Audit 4 Sheet Terstruktur & Berwarna
        data_audit = buat_excel_rekap_audit(df_hasil)
        with col_dl2:
            st.download_button(
                label="📥 Unduh Laporan Audit & Temuan (4 Sheet)",
                data=data_audit,
                file_name="Rekap_Audit_Temuan_DCD.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_audit"
            )

        # Generator File Unduhan 3: Laporan Investigasi Anomali PIBC
        data_pibc_bodong = buat_excel_investigasi_pibc(df_hasil, df_pibc_b)
        with col_dl3:
            st.download_button(
                label="📥 Unduh Laporan Investigasi Anomali PIBC",
                data=data_pibc_bodong,
                file_name="Laporan_Investigasi_Anomali_PIBC.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_pibc_missing"
            )

        # Generator File Unduhan 4: Laporan Investigasi Anomali PG
        data_pg_bodong = buat_excel_investigasi_pg(df_hasil, df_pg_b)
        with col_dl4:
            st.download_button(
                label="📥 Unduh Laporan Investigasi Anomali PG",
                data=data_pg_bodong,
                file_name="Laporan_Investigasi_Anomali_PG.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_pg_missing"
            )
