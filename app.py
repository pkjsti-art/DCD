import streamlit as st
import pandas as pd
from modules.cleaner import bersihkan_master_mb, bersihkan_master_pibc, bersihkan_master_pg
from modules.processor import (
    siapkan_kolom_standar, 
    jalankan_proses_crosscheck, 
    buat_excel_incorrect_multi_sheet,
    buat_excel_pibc_tidak_di_mb,
    buat_excel_pg_tidak_di_mb
)
import io

st.set_page_config(
    page_title="Data Control Dyeing (DCD)",
    page_icon="🧪",
    layout="wide"
)

st.title("DATA CONTROL DYEING (DCD)")
st.markdown("Sistem pengecekan dan rekonsiliasi data produksi secara otomatis dan terintegrasi.")

# Inisialisasi Session State
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
                df_pibc_bersih = bersihkan_master_pibc(uploaded_pibc)
                
                df_pg_bersih = None
                if uploaded_pg:
                    df_pg_list = [bersihkan_master_pg(f) for f in uploaded_pg]
                    df_pg_bersih = pd.concat(df_pg_list, ignore_index=True)
                
                df_gabung_list = []
                for file_mb in uploaded_mb:
                    df_bersih = bersihkan_master_mb(file_mb)
                    df_siap = siapkan_kolom_standar(df_bersih)
                    df_gabung_list.append(df_siap)
                
                df_final_mb = pd.concat(df_gabung_list, ignore_index=True)
                df_hasil = jalankan_proses_crosscheck(df_final_mb, df_pibc_bersih, df_pg_bersih)
                
                # Simpan ke Session State
                st.session_state['df_hasil_crosscheck'] = df_hasil
                st.session_state['df_pibc_bersih'] = df_pibc_bersih
                st.session_state['df_pg_bersih'] = df_pg_bersih
                
                st.success("Proses pengecekan selesai sepenuhnya!")
        else:
            st.warning("Mohon upload setidaknya file Master MB dan Master PIBC terlebih dahulu.")

    # Tampilkan Hasil jika Data Ada di Session State
    if st.session_state['df_hasil_crosscheck'] is not None:
        df_hasil = st.session_state['df_hasil_crosscheck']
        df_pibc_b = st.session_state['df_pibc_bersih']
        df_pg_b = st.session_state['df_pg_bersih']

        st.subheader("Tabel Hasil Cross Check Data")
        st.dataframe(df_hasil, use_container_width=True)
        
        st.subheader("Pusat Unduhan (Download Center)")
        col_dl1, col_dl2 = st.columns(2)
        col_dl3, col_dl4 = st.columns(2)
        
        output_all = io.BytesIO()
        with pd.ExcelWriter(output_all, engine='openpyxl') as writer:
            df_hasil.to_excel(writer, index=False, sheet_name='Hasil Crosscheck')
        data_all = output_all.getvalue()
        
        with col_dl1:
            st.download_button(
                label="📥 Unduh Hasil Cross Check Keseluruhan",
                data=data_all,
                file_name="Hasil_Crosscheck_DCD_Lengkap.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        data_incorrect = buat_excel_incorrect_multi_sheet(df_hasil)
        with col_dl2:
            st.download_button(
                label="📥 Unduh Rekap Data INCORRECT (3 Sheet)",
                data=data_incorrect,
                file_name="Rekap_Incorrect_DCD.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        data_pibc_tdk_ada = buat_excel_pibc_tidak_di_mb(df_hasil, df_pibc_b)
        with col_dl3:
            st.download_button(
                label="📥 Unduh Hasil PIBC yang Tidak Ada di MB",
                data=data_pibc_tdk_ada,
                file_name="PIBC_Tidak_Ada_di_MB.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        data_pg_tdk_ada = buat_excel_pg_tidak_di_mb(df_hasil, df_pg_b)
        with col_dl4:
            st.download_button(
                label="📥 Unduh Hasil PG yang Tidak Ada di MB",
                data=data_pg_tdk_ada,
                file_name="PG_Tidak_Ada_di_MB.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
