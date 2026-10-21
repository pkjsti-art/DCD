import streamlit as st
import pandas as pd
from modules.cleaner import bersihkan_master_mb, bersihkan_master_pibc, bersihkan_master_pg
from modules.processor import siapkan_kolom_standar, jalankan_proses_crosscheck, buat_excel_incorrect_multi_sheet
import io

# Konfigurasi Tampilan Halaman Web
st.set_page_config(
    page_title="Data Control Dyeing (DCD)",
    page_icon="🧪",
    layout="wide"
)

st.title("DATA CONTROL DYEING (DCD)")
st.markdown("Sistem pengecekan dan rekonsiliasi data produksi secara otomatis dan terintegrasi.")

# Menu Sidebar
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
            with st.spinner("Sedang membersihkan, menyiapkan, dan memproses crosscheck data secara menyeluruh..."):
                # 1. Bersihkan Master PIBC
                df_pibc_bersih = bersihkan_master_pibc(uploaded_pibc)
                
                # 2. Bersihkan Master PG (jika ada)
                df_pg_bersih = None
                if uploaded_pg:
                    df_pg_list = [bersihkan_master_pg(f) for f in uploaded_pg]
                    df_pg_bersih = pd.concat(df_pg_list, ignore_index=True)
                
                # 3. Proses Gabung & Bersihkan Master MB dari semua gudang
                df_gabung_list = []
                for file_mb in uploaded_mb:
                    df_bersih = bersihkan_master_mb(file_mb)
                    df_siap = siapkan_kolom_standar(df_bersih)
                    df_gabung_list.append(df_siap)
                
                df_final_mb = pd.concat(df_gabung_list, ignore_index=True)
                
                # 4. Jalankan Core Crosscheck Logic End-to-End
                df_hasil_crosscheck = jalankan_proses_crosscheck(df_final_mb, df_pibc_bersih, df_pg_bersih)
                
                st.success("Proses pengecekan dan crosscheck data selesai sepenuhnya!")
                
                # 5. Tampilkan Tabel Hasil Interaktif
                st.subheader("Tabel Hasil Cross Check Data")
                st.dataframe(df_hasil_crosscheck, use_container_width=True)
                
                # 6. Pusat Unduhan (Download Center - 4 Jenis Output File Excel)
                st.subheader("Pusat Unduhan (Download Center)")
                col_dl1, col_dl2 = st.columns(2)
                
                # Unduhan File 1: Keseluruhan Hasil Crosscheck
                output_all = io.BytesIO()
                with pd.ExcelWriter(output_all, engine='openpyxl') as writer:
                    df_hasil_crosscheck.to_excel(writer, index=False, sheet_name='Hasil Crosscheck')
                data_all = output_all.getvalue()
                
                with col_dl1:
                    st.download_button(
                        label="📥 Unduh Hasil Cross Check Keseluruhan",
                        data=data_all,
                        file_name="Hasil_Crosscheck_DCD_Lengkap.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Unduhan File 2: Rekap Incorrect 3 Sheet
                data_incorrect = buat_excel_incorrect_multi_sheet(df_hasil_crosscheck)
                with col_dl2:
                    st.download_button(
                        label="📥 Unduh Rekap Data INCORRECT (3 Sheet)",
                        data=data_incorrect,
                        file_name="Rekap_Incorrect_DCD.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.warning("Mohon upload setidaknya file Master MB dan Master PIBC terlebih dahulu.")
