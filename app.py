import streamlit as st
import pandas as pd
from modules.cleaner import bersihkan_master_mb, bersihkan_master_pibc, bersihkan_master_pg
from modules.processor import siapkan_kolom_standar

# Konfigurasi Tampilan Halaman Web
st.set_page_config(
    page_title="Data Control Dyeing (DCD)",
    page_icon="🧪",
    layout="wide"
)

st.title("DATA CONTROL DYEING (DCD)")
st.markdown("Sistem pengecekan dan rekonsiliasi data produksi secara otomatis.")

# Menu Sidebar
st.sidebar.header("Navigasi")
menu = st.sidebar.selectbox("Pilih Menu", ["Cross Check Data"])

if menu == "Cross Check Data":
    st.subheader("Upload File Master Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        uploaded_mb = st.file_uploader("Upload File Master MB", type=["xlsx", "xls"], accept_multiple_files=True)
        uploaded_pibc = st.file_uploader("Upload File Master PIBC / Dyelot", type=["xlsx", "xls"])
        
    with col2:
        uploaded_pg = st.file_uploader("Upload File Master PG (Pindah Gudang)", type=["xlsx", "xls"], accept_multiple_files=True)
        
    if st.button("Proses Data", type="primary"):
        if uploaded_mb and uploaded_pibc:
            with st.spinner("Sedang membersihkan dan memproses data..."):
                df_gabung_list = []
                for file_mb in uploaded_mb:
                    df_bersih = bersihkan_master_mb(file_mb)
                    df_siap = siapkan_kolom_standar(df_bersih)
                    df_gabung_list.append(df_siap)
                
                df_final_mb = pd.concat(df_gabung_list, ignore_index=True)
                
                st.success("Data berhasil dibersihkan dan disiapkan!")
                st.dataframe(df_final_mb.head(50), use_container_width=True)
        else:
            st.warning("Mohon upload setidaknya file Master MB dan Master PIBC terlebih dahulu.")
