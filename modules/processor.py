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

    # Menambahkan kolom penampung hasil crosscheck
    df_mb['Crosscheck Qty'] = 'CORRECT'
    df_mb['Qty PIBC'] = 0.0
    df_mb['Qty PG'] = 0.0
    df_mb['Selisih (GR)'] = 0.0
    df_mb['Status'] = 'COMPLATE'
    df_mb['Status PIBC'] = 'KODE PIBC ADA'
    df_mb['Rincian Benang'] = ''
    df_mb['Cek Variant'] = 'CLEAN'

    return df_mb


def ekstrak_kode_pibc_serries(seri_teks, seri_gudang):
    """
    Ekstraksi kode PIBC berbasis vektor (cepat untuk ribuan baris sekaligus).
    """
    def parse_satu_teks(teks, gudang):
        if pd.isna(teks):
            return None
        teks_str = str(teks)
        
        # Kasus TO / Reprocess khusus Mesin Dyeing
        if 'DYEING' in str(gudang).upper() and ('REPROCES' in teks_str.upper() or 'TO' in teks_str.upper()):
            match_kurung = re.search(r'\((PIBC-[A-Za-z0-9#\-\.]+)\)', teks_str, re.IGNORECASE)
            if match_kurung:
                return match_kurung.group(1).strip().rstrip('.')
            match_to = re.search(r'TO\s+([A-Za-z0-9\-\#]+)', teks_str, re.IGNORECASE)
            if match_to:
                return match_to.group(1).strip().rstrip('.')

        pola = r'(PIBC-[A-Za-z0-9#\-\.]+)'
        hasil = re.search(pola, teks_str, re.IGNORECASE)
        if hasil:
            return hasil.group(1).strip().rstrip('.')
        return None

    return [parse_satu_teks(t, g) for t, g in zip(seri_teks, seri_gudang)]


def jalankan_proses_crosscheck(df_mb, df_pibc, df_pg):
    """
    Logika crosscheck kilat menggunakan operasi Pandas yang dioptimalkan.
    """
    baris_diabaikan = [
        "Total Material (Details)", "Total Pekerjaan (Works)", 
        "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor"
    ]
    keyword_selisih = ['AVALAN', 'CROCHET', 'KOR', 'KOR ROMBE', 'KOLONG', 'REEBOK', 'ROLL', 'TALI', 'KUR', 'AVL', 'CONS', 'HTC']

    # Gabung kolom keterangan untuk pencarian
    ket = df_mb['Keterangan'].astype(str) if 'Keterangan' in df_mb.columns else ""
    ket_lain = df_mb['Keterangan Lain'].astype(str) if 'Keterangan Lain' in df_mb.columns else ""
    gudang_seri = df_mb['Gudang'].astype(str) if 'Gudang' in df_mb.columns else ""
    
    df_mb['Te Gabungan'] = ket + " " + ket_lain
    df_mb['Ext_PIBC'] = ekstrak_kode_pibc_serries(df_mb['Te Gabungan'], gudang_seri)

    # Tandai status PIBC ada / tidak ada secara massal
    df_mb.loc[df_mb['Ext_PIBC'].isna(), 'Status PIBC'] = 'KODE PIBC TIDAK ADA'

    # Mapping ringkas Master PIBC jika ada kolom Dyelot 1 dan Actual
    pibc_dict = {}
    if df_pibc is not None and not df_pibc.empty:
        # Cari kolom yang mirip dyelot / actual
        kolom_dyelot = next((c for c in df_pibc.columns if 'DYELOT' in c.upper() or 'PIBC' in c.upper()), None)
        kolom_actual = next((c for c in df_pibc.columns if 'ACTUAL' in c.upper() or 'QTY' in c.upper()), None)
        if kolom_dyelot and kolom_actual:
            pibc_dict = pd.Series(df_pibc[kolom_actual].values, index=df_pibc[kolom_dyelot].astype(str).str.strip()).to_dict()

    for idx, row in df_mb.iterrows():
        kode_item = str(row.get('Kode Item', '')).strip()
        if kode_item in baris_diabaikan:
            df_mb.at[idx, 'Crosscheck Qty'] = 'DIHIRAUKAN'
            continue
            
        gudang = str(row.get('Gudang', '')).strip().upper()
        nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
        teks_gab = str(row.get('Te Gabungan', ''))
        
        ada_keyword_selisih = any(kw in nama_brg_jdi or kw in teks_gab.upper() for kw in keyword_selisih)
        if ada_keyword_selisih:
            df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH'

        # Pengecekan Benang
        if 'PRODUKSI SOFTCONE' in gudang or kode_item.startswith(('TWP', 'MWP', 'TBM', 'MBB')):
            qty_target = row.get('Qty Target Standar (GR)', 0.0)
            qty_bb = row.get('Qty BB Standar (GR)', 0.0)
            selisih = qty_bb - qty_target
            df_mb.at[idx, 'Selisih (GR)'] = selisih
            if abs(selisih) > 0.01 and not ada_keyword_selisih:
                df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
            else:
                df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
            df_mb.at[idx, 'Rincian Benang'] = kode_item
        elif 'DYEING STI' in gudang:
            if kode_item != 'TBB':
                df_mb.at[idx, 'Rincian Benang'] = 'BUKAN BENANG'
        elif 'DYEING' in gudang or 'LAB' in gudang:
            kode_pbc = row.get('Ext_PIBC')
            if pibc_dict and kode_pbc and kode_pbc in pibc_dict:
                qty_actual = float(pibc_dict[kode_pbc] or 0.0)
                qty_bb = row.get('Qty BB Standar (GR)', 0.0)
                df_mb.at[idx, 'Qty PIBC'] = qty_actual
                selisih = qty_bb - qty_actual
                df_mb.at[idx, 'Selisih (GR)'] = selisih
                if abs(selisih) > 0.01 and not ada_keyword_selisih:
                    df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                else:
                    df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
            else:
                df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'

        # Cek Variant
        kode_bj = str(row.get('Kode Barang Jadi', '')).strip()
        if 'MESIN DYEING' in gudang:
            if not kode_item.startswith(('TWP.050', 'TWP.100', 'TWP.051')) and kode_bj not in ['TWP.100', 'TWP.400', 'TWP.051', 'TWP.666', 'TWP.901']:
                df_mb.at[idx, 'Cek Variant'] = kode_bj
        elif 'LAB' in gudang:
            if not kode_item.startswith(('TWP.050', 'TWP.051', 'TWP.999', 'TWP.100')) and kode_bj not in ['TWP.999', 'TWP.100', 'TWP.400', 'TWP.666', 'TWP.051']:
                df_mb.at[idx, 'Cek Variant'] = kode_bj

    # Hapus kolom bantu sementara
    if 'Te Gabungan' in df_mb.columns:
        df_mb = df_mb.drop(columns=['Te Gabungan', 'Ext_PIBC'])

    return df_mb


def buat_excel_incorrect_multi_sheet(df_hasil):
    """
    Membuat file Excel unduhan Incorrect multi-sheet.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_inc = df_hasil[df_hasil['Crosscheck Qty'] == 'INCORRECT']
        kode_grup_inc = df_inc['Kode'].unique() if 'Kode' in df_inc.columns else []
        df_full_inc = df_hasil[df_hasil['Kode'].isin(kode_grup_inc)] if len(kode_grup_inc) > 0 else df_inc
        df_full_inc.to_excel(writer, index=False, sheet_name='INCORRECT')

        df_selisih = df_hasil[df_hasil['Crosscheck Qty'] == 'MEMANG SELISIH']
        kode_grup_selisih = df_selisih['Kode'].unique() if 'Kode' in df_selisih.columns else []
        df_full_selisih = df_hasil[df_hasil['Kode'].isin(kode_grup_selisih)] if len(kode_grup_selisih) > 0 else df_selisih
        df_full_selisih.to_excel(writer, index=False, sheet_name='INCORRECT Memang Selisih')

        df_benang_inc = df_hasil[(df_hasil['Rincian Benang'] != '') & (df_hasil['Crosscheck Qty'] == 'INCORRECT')]
        kode_grup_benang = df_benang_inc['Kode'].unique() if 'Kode' in df_benang_inc.columns else []
        df_full_benang = df_hasil[df_hasil['Kode'].isin(kode_grup_benang)] if len(kode_grup_benang) > 0 else df_benang_inc
        df_full_benang.to_excel(writer, index=False, sheet_name='INCORRECT Benang')

    return output.getvalue()
