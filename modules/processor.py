import pandas as pd
import re
import io

def siapkan_kolom_standar(df_mb):
    """
    Menambahkan kolom baru dan mengonversi satuan ke GR sesuai urutan rancang bangun.
    """
    # 1. Konversi Qty Target ke Satuan GR
    if 'Target Unit' in df_mb.columns and 'Target Qty' in df_mb.columns:
        def konversi_target_gr(row):
            unit = str(row['Target Unit']).strip().upper()
            try:
                qty_angka = float(row['Target Qty'])
            except:
                qty_angka = 0.0
            return qty_angka if unit == 'GR' else qty_angka * 1000
        df_mb['Qty Target Standar (GR)'] = df_mb.apply(konversi_target_gr, axis=1)
    else:
        df_mb['Qty Target Standar (GR)'] = 0.0

    # 2. Konversi Qty BB ke Satuan GR
    if 'Unit' in df_mb.columns and 'Qty' in df_mb.columns:
        def konversi_bb_gr(row):
            unit = str(row['Unit']).strip().upper()
            try:
                qty_angka = float(row['Qty'])
            except:
                qty_angka = 0.0
            return qty_angka if unit == 'GR' else qty_angka * 1000
        df_mb['Qty BB Standar (GR)'] = df_mb.apply(konversi_bb_gr, axis=1)
    else:
        df_mb['Qty BB Standar (GR)'] = 0.0

    # Menambahkan kolom penampung hasil crosscheck sesuai urutan rancang bangun
    df_mb['Crosscheck Qty'] = 'CORRECT'
    df_mb['Qty PIBC'] = 0.0
    df_mb['Qty PG'] = 0.0
    df_mb['Selisih (GR)'] = 0.0
    df_mb['Status'] = 'COMPLATE'
    df_mb['Status PIBC'] = 'KODE PIBC ADA'
    df_mb['Rincian Benang'] = ''
    df_mb['Cek Variant'] = 'CLEAN'

    return df_mb


def ekstrak_kode_pibc(teks, gudang=""):
    """
    Mengekstrak kode PIBC dengan dukungan Regex fleksibel.
    Untuk Gudang Mesin Dyeing pada kasus TO (Reprocess), mengambil kode terakhir di dalam () atau setelah 'TO'.
    """
    if pd.isna(teks):
        return None
    teks_str = str(teks)
    
    # Khusus Case TO / Reprocess di Gudang Mesin Dyeing
    if 'DYEING' in gudang.upper() and ('REPROCES' in teks_str.upper() or 'TO' in teks_str.upper()):
        # Cari di dalam tanda kurung ()
        match_kurung = re.search(r'\((PIBC-[A-Za-z0-9#\-\.]+)\)', teks_str, re.IGNORECASE)
        if match_kurung:
            return match_kurung.group(1).strip().rstrip('.')
        
        # Cari setelah kata TO
        match_to = re.search(r'TO\s+([A-Za-z0-9\-\#]+)', teks_str, re.IGNORECASE)
        if match_to:
            return match_to.group(1).strip().rstrip('.')

    # Pencarian standar kode PIBC
    pola = r'(PIBC-[A-Za-z0-9#\-\.]+)'
    hasil = re.search(pola, teks_str, re.IGNORECASE)
    if hasil:
        return hasil.group(1).strip().rstrip('.')
    return None


def jalankan_proses_crosscheck(df_mb, df_pibc, df_pg):
    """
    Logika utama crosscheck, rekonsiliasi, pengecekan benang, chemical, variant, dan keyword khusus.
    """
    baris_diabaikan = [
        "Total Material (Details)", "Total Pekerjaan (Works)", 
        "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor"
    ]
    
    keyword_selisih = ['AVALAN', 'CROCHET', 'KOR', 'KOR ROMBE', 'KOLONG', 'REEBOK', 'ROLL', 'TALI', 'KUR', 'AVL', 'CONS', 'HTC']

    for idx, row in df_mb.iterrows():
        kode_item = str(row.get('Kode Item', '')).strip()
        
        # Hiraukan baris total/overhead
        if kode_item in baris_diabaikan:
            df_mb.at[idx, 'Crosscheck Qty'] = 'DIHIRAUKAN'
            continue
            
        gudang = str(row.get('Gudang', '')).strip().upper()
        nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
        keterangan = str(row.get('Keterangan', ''))
        keterangan_lain = str(row.get('Keterangan Lain', ''))
        teks_gabungan = keterangan + " " + keterangan_lain
        
        # Cek Keyword "Memang Selisih"
        ada_keyword_selisih = any(kw in nama_brg_jdi or kw in teks_gabungan.upper() for kw in keyword_selisih)
        if ada_keyword_selisih:
            df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH'

        # Ekstrak Kode PIBC
        kode_pibc = ekstrak_kode_pibc(teks_gabungan, gudang)
        if not kode_pibc:
            df_mb.at[idx, 'Status PIBC'] = 'KODE PIBC TIDAK ADA'

        # 1. Pengecekan Benang
        is_benang_softcone = 'PRODUKSI SOFTCONE' in gudang
        is_benang_standar = kode_item.startswith(('TWP', 'MWP', 'TBM', 'MBB'))
        
        if is_benang_softcone or is_benang_standar:
            qty_target = row.get('Qty Target Standar (GR)', 0.0)
            qty_bb = row.get('Qty BB Standar (GR)', 0.0)
            selisih = qty_bb - qty_target
            df_mb.at[idx, 'Selisih (GR)'] = selisih
            
            if abs(selisih) > 0.01:
                if not ada_keyword_selisih:
                    df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
            else:
                df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
            df_mb.at[idx, 'Rincian Benang'] = kode_item

        # Aturan Gudang Dyeing STI (Non-benang TBB)
        elif 'DYEING STI' in gudang:
            if kode_item != 'TBB':
                df_mb.at[idx, 'Rincian Benang'] = 'BUKAN BENANG'

        # 2. Pengecekan Chemical & Dyestuff (Mesin Dyeing & Lab)
        elif 'DYEING' in gudang or 'LAB' in gudang:
            if df_pibc is not None and not df_pibc.empty and kode_pibc:
                # Cocokkan dengan Master PIBC
                match_pibc = df_pibc[df_pibc.astype(str).apply(lambda x: x.str.contains(kode_pibc, case=False, na=False)).any(axis=1)]
                
                if not match_pibc.empty:
                    qty_actual = float(match_pibc.iloc[0].get('Actual', 0.0) or 0.0)
                    qty_bb = row.get('Qty BB Standar (GR)', 0.0)
                    df_mb.at[idx, 'Qty PIBC'] = qty_actual
                    
                    selisih = qty_bb - qty_actual
                    df_mb.at[idx, 'Selisih (GR)'] = selisih
                    
                    if abs(selisih) > 0.01:
                        if not ada_keyword_selisih:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                    else:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                elif df_pg is not None and not df_pg.empty:
                    # Fallback ke Master PG
                    match_pg = df_pg[df_pg.astype(str).apply(lambda x: x.str.contains(kode_pibc, case=False, na=False)).any(axis=1)]
                    if not match_pg.empty:
                        qty_masuk = float(match_pg.iloc[0].get('MASUK', 0.0) or 0.0)
                        qty_bb = row.get('Qty BB Standar (GR)', 0.0)
                        df_mb.at[idx, 'Qty PG'] = qty_masuk
                        selisih = qty_bb - qty_masuk
                        df_mb.at[idx, 'Selisih (GR)'] = selisih
                        if abs(selisih) > 0.01 and not ada_keyword_selisih:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                    else:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'
                else:
                    df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'
            else:
                df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'

        # Cek Variant (Gudang Mesin Dyeing & Lab)
        kode_bj = str(row.get('Kode Barang Jadi', '')).strip()
        if 'MESIN DYEING' in gudang:
            if not kode_item.startswith(('TWP.050', 'TWP.100', 'TWP.051')) and kode_bj not in ['TWP.100', 'TWP.400', 'TWP.051', 'TWP.666', 'TWP.901']:
                df_mb.at[idx, 'Cek Variant'] = kode_bj
        elif 'LAB' in gudang:
            if not kode_item.startswith(('TWP.050', 'TWP.051', 'TWP.999', 'TWP.100')) and kode_bj not in ['TWP.999', 'TWP.100', 'TWP.400', 'TWP.666', 'TWP.051']:
                df_mb.at[idx, 'Cek Variant'] = kode_bj

    return df_mb


def buat_excel_incorrect_multi_sheet(df_hasil):
    """
    Membuat file Excel unduhan Incorrect yang berisi 3 Sheet:
    1. INCORRECT (Semua data grup MB yang memiliki status Incorrect)
    2. INCORRECT Memang Selisih (Grup MB dengan keyword khusus selisih)
    3. INCORRECT Benang (Grup MB yang salah pada pengecekan benang)
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Sheet Incorrect Umum
        df_inc = df_hasil[df_hasil['Crosscheck Qty'] == 'INCORRECT']
        kode_grup_inc = df_inc['Kode'].unique() if 'Kode' in df_inc.columns else []
        df_full_inc = df_hasil[df_hasil['Kode'].isin(kode_grup_inc)] if len(kode_grup_inc) > 0 else df_inc
        df_full_inc.to_excel(writer, index=False, sheet_name='INCORRECT')

        # 2. Sheet Memang Selisih
        df_selisih = df_hasil[df_hasil['Crosscheck Qty'] == 'MEMANG SELISIH']
        kode_grup_selisih = df_selisih['Kode'].unique() if 'Kode' in df_selisih.columns else []
        df_full_selisih = df_hasil[df_hasil['Kode'].isin(kode_grup_selisih)] if len(kode_grup_selisih) > 0 else df_selisih
        df_full_selisih.to_excel(writer, index=False, sheet_name='INCORRECT Memang Selisih')

        # 3. Sheet Incorrect Benang
        df_benang_inc = df_hasil[(df_hasil['Rincian Benang'] != '') & (df_hasil['Crosscheck Qty'] == 'INCORRECT')]
        kode_grup_benang = df_benang_inc['Kode'].unique() if 'Kode' in df_benang_inc.columns else []
        df_full_benang = df_hasil[df_hasil['Kode'].isin(kode_grup_benang)] if len(kode_grup_benang) > 0 else df_benang_inc
        df_full_benang.to_excel(writer, index=False, sheet_name='INCORRECT Benang')

    return output.getvalue()
