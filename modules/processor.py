def buat_excel_rekap_audit(df_hasil):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # 1. Sheet Laporan Ringkas Admin (Surat Tilang)
        data_ringkas = []
        for idx, row in df_hasil.iterrows():
            cq = str(row.get('Crosscheck Qty', '')).upper()
            cv = str(row.get('Cek Variant', '')).upper()
            
            is_qty_err = ('INCORRECT' in cq or 'SELISIH' in cq) and cq != 'MEMANG SELISIH'
            is_var_err = cv != '' and cv != 'CLEAN'
            
            if is_qty_err or is_var_err:
                item_bermasalah = f"{row.get('Kode Item', '')} - {row.get('Nama Item', '')}"
                if is_qty_err and is_var_err:
                    jenis = "Dobel Error (Varian & Qty)"
                    detail = f"Varian: {cv}. Qty: {cq}"
                elif is_qty_err:
                    jenis = "Selisih Kuantitas"
                    detail = cq
                else:
                    jenis = "Pelanggaran SOP Varian"
                    detail = cv
                    
                data_ringkas.append({
                    'Gudang': row.get('Gudang', ''),
                    'Kode Dokumen MB': row.get('Kode', ''),
                    'Kode PIBC Terdeteksi': row.get('Ext_PIBC', ''),
                    'Item Bermasalah': item_bermasalah,
                    'Keterangan': row.get('Keterangan', ''),
                    'Keterangan Lain': row.get('Keterangan Lain', ''),
                    'Qty Target Standar (GR)': row.get('Qty Target Standar (GR)', ''),
                    'Qty BB Standar (GR)': row.get('Qty BB Standar (GR)', ''),
                    'Qty PIBC': row.get('Qty PIBC', ''),
                    'Qty PG': row.get('Qty PG', ''),
                    'Jenis Temuan': jenis,
                    'Detail Kesalahan & Seharusnya': detail,
                    'Catatan Auditor': ''
                })
        
        df_ringkas = pd.DataFrame(data_ringkas)
        if not df_ringkas.empty:
            df_ringkas.to_excel(writer, index=False, sheet_name='Laporan Ringkas Admin')
        else:
            pd.DataFrame({'Info': ['Tidak ada temuan error/selisih']}).to_excel(writer, index=False, sheet_name='Laporan Ringkas Admin')
            
        # 2. Sheet Error Qty (Full Grup)
        mask_qty = df_hasil['Crosscheck Qty'].astype(str).str.upper().apply(lambda x: ('INCORRECT' in x or 'SELISIH' in x) and x != 'MEMANG SELISIH')
        kode_qty = df_hasil[mask_qty]['Kode'].unique() if 'Kode' in df_hasil.columns else []
        if len(kode_qty) > 0:
            df_qty_full = df_hasil[df_hasil['Kode'].isin(kode_qty)]
            df_qty_full.to_excel(writer, index=False, sheet_name='Error Qty (Full Grup)')
        else:
            pd.DataFrame({'Info': ['Tidak ada temuan']}).to_excel(writer, index=False, sheet_name='Error Qty (Full Grup)')
            
        # 3. Sheet Error Varian (Full Grup)
        mask_var = df_hasil['Cek Variant'].astype(str).str.upper().apply(lambda x: x != '' and x != 'CLEAN')
        kode_var = df_hasil[mask_var]['Kode'].unique() if 'Kode' in df_hasil.columns else []
        if len(kode_var) > 0:
            df_var_full = df_hasil[df_hasil['Kode'].isin(kode_var)]
            df_var_full.to_excel(writer, index=False, sheet_name='Error Varian (Full Grup)')
        else:
            pd.DataFrame({'Info': ['Tidak ada temuan']}).to_excel(writer, index=False, sheet_name='Error Varian (Full Grup)')
            
        # 4. Sheet Memang Selisih (Full Grup)
        mask_ms = df_hasil['Crosscheck Qty'].astype(str).str.upper() == 'MEMANG SELISIH'
        kode_ms = df_hasil[mask_ms]['Kode'].unique() if 'Kode' in df_hasil.columns else []
        if len(kode_ms) > 0:
            df_ms_full = df_hasil[df_hasil['Kode'].isin(kode_ms)]
            df_ms_full.to_excel(writer, index=False, sheet_name='Memang Selisih (Full Grup)')
        else:
            pd.DataFrame({'Info': ['Tidak ada temuan']}).to_excel(writer, index=False, sheet_name='Memang Selisih (Full Grup)')
            
    return output.getvalue()
