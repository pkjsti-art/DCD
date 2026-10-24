# Generator File Unduhan 1: Keseluruhan Hasil (MEMANGGIL FUNGSI WARNA BARU)
        data_all = buat_excel_hasil_keseluruhan(df_hasil)
        
        with col_dl1:
            st.download_button(
                label="📥 Unduh Hasil Cross Check Keseluruhan",
                data=data_all,
                file_name="Hasil_Crosscheck_DCD_Lengkap.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_all"
            )
