import streamlit as st
import pandas as pd
from modules.cleaner import (
    KOLOM_LOG_DUPLIKAT,
    bersihkan_master_mb,
    deduplikasi_lintas_file_mb,
    bersihkan_master_pibc,
    bersihkan_master_pg,
    bersihkan_master_mesin
)
from modules.processor import (
    siapkan_kolom_standar, 
    jalankan_proses_crosscheck, 
    buat_excel_hasil_keseluruhan,
    buat_excel_rekap_audit,
    buat_excel_investigasi_pibc,
    buat_excel_investigasi_pg,
    buat_zip_semua_laporan,
    get_hari_ini_wib
)
import base64

st.set_page_config(
    page_title="Data Control Dyeing (DCD)",
    page_icon="🧪",
    layout="wide"
)

def get_base64_of_bin_file(bin_file):
    try:
        with open(bin_file, 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except:
        return ""

bin_str = get_base64_of_bin_file('bg_datacenter.png')
bg_css = f"data:image/png;base64,{bin_str}" if bin_str else ""

st.markdown(f"""
<style>
    [data-testid="stAppViewContainer"] {{
        background-image: 
            linear-gradient(rgba(5, 5, 20, 0.75), rgba(10, 5, 25, 0.85)),
            url("{bg_css}") !important;
        background-size: cover !important;
        background-position: center !important;
        background-repeat: no-repeat !important;
        background-attachment: fixed !important;
    }}
    .stApp, [data-testid="stHeader"] {{
        background: transparent !important;
    }}
    [data-testid="stSidebar"] {{
        background: rgba(10, 10, 25, 0.65) !important;
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
        border-right: 1px solid rgba(0, 243, 255, 0.3) !important;
        box-shadow: 
            5px 0px 30px rgba(127, 0, 255, 0.25),
            inset -1px 0px 0px 0px rgba(255, 255, 255, 0.1);
    }}
    [data-testid="stSidebar"] header {{
        display: flex;
        align-items: center;
        justify-content: flex-end;
        padding-right: 10px;
    }}
    [data-testid="stSidebar"] header span:not(:has(svg)) {{
        display: none !important;
    }}
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {{
        color: #00f3ff !important;
        font-family: 'Courier New', Courier, monospace;
        letter-spacing: 1px;
    }}
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
    p, span, label, .stMarkdown, .stText {{
        color: #c9c9e8 !important;
    }}
    [data-testid="block-container"] {{
        padding-top: 2rem !important;
    }}
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
    hr {{
        border-top: 1px solid #7f00ff !important;
        box-shadow: 0px 0px 10px #7f00ff;
    }}
    [data-testid="stDataFrame"] {{
        border: 1px solid #00f3ff;
        box-shadow: 0 0 15px rgba(0, 243, 255, 0.2);
        border-radius: 4px;
        background-color: rgba(0,0,0,0.7);
    }}
</style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="glass-header-card">
        <h1 style="color: #ffffff; font-family: 'Courier New', Courier, monospace; font-weight: 900; letter-spacing: 3px; font-size: 2.5rem; text-transform: uppercase; margin: 0; padding: 0; text-shadow: 0 0 10px #00f3ff, 0 0 20px #00f3ff, 0 0 40px #7f00ff, 0 0 80px #7f00ff; text-align: center;">DATA CONTROL DYEING (DCD)</h1>
        <p style="color: #b3b3ff; font-weight: bold; letter-spacing: 1.2px; font-size: 1.18rem; margin-top: 12px; margin-bottom: 0; padding: 0; text-align: center;">Sistem pengecekan dan rekonsiliasi data produksi secara otomatis dan terintegrasi.</p>
    </div>
""", unsafe_allow_html=True)

for state_key in [
    'df_hasil_crosscheck', 'df_pibc_bersih', 'df_pg_bersih', 'df_mesin_bersih',
    'df_log_duplikat', 'peringatan_duplikasi', 'paket_unduhan', 'stempel_tanggal'
]:
    if state_key not in st.session_state:
        st.session_state[state_key] = None

st.sidebar.markdown("### ⚡ Navigasi Menu")
menu = st.sidebar.selectbox("Pilih Menu", ["Cross Check Data"])

if menu == "Cross Check Data":
    st.subheader("Upload Dokumen Master Data")
    
    col1, col2 = st.columns(2)
    with col1:
        uploaded_mb = st.file_uploader("Upload File Master MB (Bisa Pilih Banyak)", type=["xlsx", "xls"], accept_multiple_files=True)
        uploaded_pibc = st.file_uploader("Upload File Master Memphis (PIBC / Dyelot)", type=["xlsx", "xls"])
    with col2:
        uploaded_pg = st.file_uploader("Upload File Master PG / Pindah Gudang (Bisa Pilih Banyak)", type=["xlsx", "xls"], accept_multiple_files=True)
        uploaded_mesin = st.file_uploader("Upload File Master Data Mesin (Opsional)", type=["xlsx", "xls"])
        
    if st.button("Proses Data Sekarang", type="primary"):
        if uploaded_mb and uploaded_pibc:
            with st.spinner("Sedang membersihkan, mengecek duplikat, memproses audit 3 arah, dan menyiapkan laporan Excel..."):
                peringatan_list = []
                log_parts = []

                df_pibc_bersih = bersihkan_master_pibc(uploaded_pibc)
                
                df_pg_bersih = None
                if uploaded_pg:
                    df_pg_bersih, df_log_pg = bersihkan_master_pg(uploaded_pg, return_log=True)
                    if not df_log_pg.empty:
                        log_parts.append(df_log_pg)

                df_mesin_bersih = None
                if uploaded_mesin:
                    df_mesin_bersih = bersihkan_master_mesin(uploaded_mesin)
                
                list_tuple_mb = []
                for idx_m, file_mb in enumerate(uploaded_mb):
                    fname_mb = getattr(file_mb, 'name', f'File_MB_{idx_m + 1}')
                    df_bersih = bersihkan_master_mb(file_mb)
                    df_siap = siapkan_kolom_standar(df_bersih)
                    list_tuple_mb.append((fname_mb, df_siap))

                df_final_mb, df_log_mb, duplikat_mb_map = deduplikasi_lintas_file_mb(list_tuple_mb)
                if not df_log_mb.empty:
                    log_parts.append(df_log_mb)

                df_log_duplikat = (
                    pd.concat(log_parts, ignore_index=True)
                    if log_parts else
                    pd.DataFrame(columns=KOLOM_LOG_DUPLIKAT)
                )

                if not df_log_duplikat.empty:
                    for sumber in ['Master PG', 'Master MB']:
                        sub_s = df_log_duplikat[df_log_duplikat['Sumber Master'] == sumber]
                        if sub_s.empty:
                            continue

                        sub_lintas = sub_s[sub_s['Jenis Duplikat'].str.contains('Lintas File', na=False)]
                        if not sub_lintas.empty:
                            jml_dok = len(sub_lintas)
                            jml_beda = int(sub_lintas['Jenis Duplikat'].str.contains('Isi Berbeda', na=False).sum())
                            info_beda = f" ({jml_beda} dokumen terdeteksi isi berbeda antar file)" if jml_beda > 0 else ""
                            peringatan_list.append(
                                f"{sumber}: {jml_dok} kode transaksi duplikat lintas file (gudang sama) otomatis dibersihkan dengan mempertahankan versi file terakhir{info_beda}."
                            )

                        sub_blok_c = sub_s[sub_s['Jenis Duplikat'].str.contains('Dalam 1 File', na=False)]
                        if not sub_blok_c.empty:
                            peringatan_list.append(
                                f"{sumber}: {len(sub_blok_c)} blok transaksi identik berulang (sama Kode MB dalam 1 file) otomatis dibersihkan."
                            )

                        sub_beda_a = sub_s[sub_s['Jenis Duplikat'].str.contains('Beda Kode', na=False)]
                        if not sub_beda_a.empty:
                            peringatan_list.append(
                                f"{sumber}: Terdeteksi {len(sub_beda_a)} Kode MB dengan seluruh isi kelompok transaksi identik (beda Kode). Data TIDAK dihapus otomatis — cek kolom Info dan Indikasi serta Log Duplikat."
                            )
                
                df_hasil = jalankan_proses_crosscheck(
                    df_final_mb, df_pibc_bersih, df_pg_bersih, df_mesin_bersih, duplikat_mb_map
                )

                tgl_str = get_hari_ini_wib().strftime("%Y%m%d")
                nama_f1 = f"Hasil_Crosscheck_DCD_Lengkap_{tgl_str}.xlsx"
                nama_f2 = f"Rekap_Audit_Temuan_DCD_{tgl_str}.xlsx"
                nama_f3 = f"Laporan_Investigasi_Anomali_PIBC_{tgl_str}.xlsx"
                nama_f4 = f"Laporan_Investigasi_Anomali_PG_{tgl_str}.xlsx"

                bytes_all = buat_excel_hasil_keseluruhan(df_hasil, df_pibc_bersih, df_pg_bersih)
                bytes_audit = buat_excel_rekap_audit(df_hasil, df_pibc_bersih, df_pg_bersih, df_log_duplikat)
                bytes_pibc = buat_excel_investigasi_pibc(df_hasil, df_pibc_bersih, df_pg_bersih)
                bytes_pg = buat_excel_investigasi_pg(df_hasil, df_pg_bersih)
                bytes_zip = buat_zip_semua_laporan({
                    nama_f1: bytes_all,
                    nama_f2: bytes_audit,
                    nama_f3: bytes_pibc,
                    nama_f4: bytes_pg
                })
                
                st.session_state['df_hasil_crosscheck'] = df_hasil
                st.session_state['df_pibc_bersih'] = df_pibc_bersih
                st.session_state['df_pg_bersih'] = df_pg_bersih
                st.session_state['df_mesin_bersih'] = df_mesin_bersih
                st.session_state['df_log_duplikat'] = df_log_duplikat
                st.session_state['peringatan_duplikasi'] = peringatan_list
                st.session_state['stempel_tanggal'] = tgl_str
                st.session_state['paket_unduhan'] = {
                    'all': (nama_f1, bytes_all),
                    'audit': (nama_f2, bytes_audit),
                    'pibc': (nama_f3, bytes_pibc),
                    'pg': (nama_f4, bytes_pg),
                    'zip': (f"Paket_Laporan_Audit_DCD_{tgl_str}.zip", bytes_zip)
                }
                
                st.success("Proses pengecekan selesai sepenuhnya!")
        else:
            st.warning("Mohon upload setidaknya file Master MB dan Master Memphis (PIBC) terlebih dahulu.")

    if st.session_state['df_hasil_crosscheck'] is not None:
        df_hasil = st.session_state['df_hasil_crosscheck']
        df_log_dup = st.session_state.get('df_log_duplikat')
        paket = st.session_state['paket_unduhan']
        peringatan = st.session_state.get('peringatan_duplikasi') or []

        for msg in peringatan:
            st.warning(msg)

        if df_log_dup is not None and not df_log_dup.empty:
            with st.expander(f"🔍 Lihat Rincian Log Duplikat ({len(df_log_dup)} dokumen terdeteksi)"):
                st.dataframe(df_log_dup, use_container_width=True)

        st.markdown("---")
        st.subheader("Tabel Hasil Cross Check Data")
        st.dataframe(df_hasil, use_container_width=True)
        
        st.subheader("Pusat Unduhan (Download Center)")
        col_dl1, col_dl2 = st.columns(2)
        col_dl3, col_dl4 = st.columns(2)
        
        with col_dl1:
            st.download_button(
                label="📥 Unduh Hasil Cross Check Keseluruhan",
                data=paket['all'][1],
                file_name=paket['all'][0],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_all"
            )
        
        with col_dl2:
            st.download_button(
                label="📥 Unduh Laporan Rekap Audit & KPI",
                data=paket['audit'][1],
                file_name=paket['audit'][0],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_audit"
            )

        with col_dl3:
            st.download_button(
                label="📥 Unduh Laporan Investigasi Anomali Memphis",
                data=paket['pibc'][1],
                file_name=paket['pibc'][0],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_pibc_missing"
            )

        with col_dl4:
            st.download_button(
                label="📥 Unduh Laporan Investigasi Anomali PG",
                data=paket['pg'][1],
                file_name=paket['pg'][0],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_pg_missing"
            )

        st.markdown("")
        st.download_button(
            label="📦 Unduh Semua Laporan Sekaligus (.ZIP)",
            data=paket['zip'][1],
            file_name=paket['zip'][0],
            mime="application/zip",
            key="btn_download_zip",
            use_container_width=True
        )
