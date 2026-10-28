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
# INJEKSI TEMA UI/UX CYBERPUNK / NEON
# ==========================================
st.markdown("""
<style>
    /* Latar belakang utama aplikasi (Sangat Gelap) */
    .stApp {
        background-color: #05050A;
    }
    
    /* Latar belakang dan efek glow pada Sidebar */
    [data-testid="stSidebar"] {
        background-color: #020204 !important;
        border-right: 1px solid #BF00FF;
        box-shadow: 2px 0px 15px rgba(191, 0, 255, 0.4);
    }
    
    /* Modifikasi teks Header (H1, H2, H3) - Biru Cyan Bercahaya */
    h1, h2, h3 {
        color: #00F3FF !important;
        text-shadow: 0px 0px 10px rgba(0, 243, 255, 0.6);
        font-family: 'Courier New', Courier, monospace;
        letter-spacing: 1px;
    }
    
    /* Modifikasi semua Tombol dengan efek Neon Purple Border & Glow */
    .stButton > button, [data-testid="stDownloadButton"] > button {
        background-color: transparent !important;
        color: #BF00FF !important;
        border: 2px solid #BF00FF !important;
        border-radius: 8px;
        box-shadow: 0 0 10px rgba(191, 0, 255, 0.4) inset, 0 0 10px rgba(191, 0, 255, 0.4);
        transition: all 0.3s ease;
        font-weight: bold;
        letter-spacing: 0.5px;
    }
    
    /* Efek saat kursor diarahkan ke tombol (Hover) */
    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        background-color: #BF00FF !important;
        color: #FFFFFF !important;
        box-shadow: 0 0 20px rgba(191, 0, 255, 0.8) inset, 0 0 20px rgba(191, 0, 255, 0.8);
        border: 2px solid #BF00FF !important;
    }
    
    /* Modifikasi Area Drag & Drop File Uploader */
    [data-testid="stFileUploadDropzone"] {
        background-color: #0A0A14 !important;
        border: 2px dashed #00F3FF !important;
        box-shadow: 0 0 15px rgba(0, 243, 255, 0.2);
        border-radius: 10px;
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        background-color: #0F0F1A !important;
        box-shadow: 0 0 25px rgba(0, 243, 255, 0.4);
    }
    [data-testid="stFileUploadDropzone"] * {
        color: #00F3FF !important;
    }
    
    /* Modifikasi garis pemisah (Divider) */
    hr {
        border-top: 1px solid #BF00FF !important;
        box-shadow: 0px 1px 10px rgba(191, 0, 255, 0.5);
    }
    
    /* Warna teks standar (Paragraf, Label, dll) */
    p, span, label, .stMarkdown {
        color: #B0B0D0 !important;
    }
    
    /* Efek pendaran pada tabel Dataframe Streamlit */
    [data-testid="stDataFrame"] {
        border: 1px solid #BF00FF;
        box-shadow: 0 0 15px rgba(191, 0, 255, 0.2);
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
