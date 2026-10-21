import pandas as pd
import re
import io

def cari_kolom(df, daftar_kata_kunci):
    for col in df.columns:
        for kw in daftar_kata_kunci:
            if kw.lower() in str(col).lower():
                return col
    return None

def siapkan_kolom_standar(df_mb):
    """
    Menyisipkan Qty Target Standar (GR) tepat di sebelah kanan Target Unit,
    dan Qty BB Standar (GR) tepat di sebelah kanan Qty/Unit.
    """
    col_target_qty = cari_kolom(df_mb, ['target qty', 'target_qty', 'qty target'])
    col_target_unit = cari_kolom(df_mb, ['target unit', 'target_unit', 'unit target'])
    
    col_bb_qty = cari_kolom(df_mb, ['qty', 'bb qty', 'quantity'])
    col_bb_unit = cari_kolom(df_mb, ['unit', 'satuan'])

    # 1. Hitung Qty Target Standar (GR)
    if col_target_qty and col_target_unit:
        def calc_target_gr(r):
            unit = str(r[col_target_unit]).strip().upper()
            try:
                val = float(r[col_target_qty])
            except:
                val = 0.0
            return val if unit == 'GR' else val * 1000.0
        seri_target_gr = df_mb.apply(calc_target_gr, axis=1)
    else:
        seri_target_gr = pd.Series([0.0]*len(df_mb))

    # 2. Hitung Qty BB Standar (GR)
    if col_bb_qty and col_bb_unit:
        def calc_bb_gr(r):
            unit = str(r[col_bb_unit]).strip().upper()
            try:
                val = float(r[col_bb_qty])
            except:
                val = 0.0
            return val if unit == 'GR' else val * 1000.0
        seri_bb_gr = df_mb.apply(calc_bb_gr, axis=1)
    else:
        seri_bb_gr = pd.Series([0.0]*len(df_mb))

    # Sisipkan kolom sesuai tata letak presisi
    idx_target = list(df_mb.columns).index(col_target_unit) + 1 if col_target_unit in df_mb.columns else len(df_mb.columns)
    df_mb.insert(idx_target, 'Qty Target Standar (GR)', seri_target_gr)

    idx_bb = list(df_mb.columns).index(col_bb_qty) + 1 if col_bb_qty in df_mb.columns else len(df_mb.columns)
    df_mb.insert(idx_bb, 'Qty BB Standar (GR)', seri_bb_gr)

    # Tambahkan kolom penampung dengan tipe object agar fleksibel diisi string/angka
    for col_baru in ['Crosscheck Qty', 'Qty PIBC', 'Qty PG', 'Selisih (GR)', 'Status Cek', 'Status PIBC', 'Rincian Benang', 'Cek Variant']:
        df_mb[col_baru] = ''
        df_mb[col_baru] = df_mb[col_baru].astype(object)

    df_mb['Status PIBC'] = 'KODE PIBC ADA'

    return df_mb


def ekstrak_kode_pibc(teks, gudang=""):
    if pd.isna(teks):
        return None
    teks_str = str(teks)
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


def jalankan_proses_crosscheck(df_mb, df_pibc, df_pg):
    baris_diabaikan = [
        "Total Material (Details)", "Total Pekerjaan (Works)", 
        "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor"
    ]
    keyword_selisih = ['AVALAN', 'CROCHET', 'KOR', 'KOR ROMBE', 'KOLONG', 'REEBOK', 'ROLL', 'TALI', 'KUR', 'AVL', 'CONS', 'HTC']

    df_mb['Qty PIBC'] = df_mb['Qty PIBC'].astype(object)
    df_mb['Qty PG'] = df_mb['Qty PG'].astype(object)
    df_mb['Selisih (GR)'] = df_mb['Selisih (GR)'].astype(object)
    df_mb['Crosscheck Qty'] = df_mb['Crosscheck Qty'].astype(object)

    pibc_lookup = {}
    if df_pibc is not None and not df_pibc.empty:
        col_dyelot = cari_kolom(df_pibc, ['dyelot', 'pibc'])
        col_item = cari_kolom(df_pibc, ['obat', 'kode obat', 'item'])
        col_act = cari_kolom(df_pibc, ['actual', 'qty'])
        if col_dyelot and col_item and col_act:
            for _, r in df_pibc.iterrows():
                key = (str(r[col_dyelot]).strip().upper(), str(r[col_item]).strip().upper())
                try:
                    pibc_lookup[key] = float(r[col_act])
                except:
                    pibc_lookup[key] = 0.0

    pg_lookup = {}
    if df_pg is not None and not df_pg.empty:
        col_ket = cari_kolom(df_pg, ['keterangan', 'det', 'pibc'])
        col_var = cari_kolom(df_pg, ['variant', 'item', 'nama'])
        col_masuk = cari_kolom(df_pg, ['masuk', 'qty'])
        if col_ket and col_var and col_masuk:
            for _, r in df_pg.iterrows():
                key = (str(r[col_ket]).strip().upper(), str(r[col_var]).strip().upper())
                try:
                    pg_lookup[key] = float(r[col_masuk])
                except:
                    pg_lookup[key] = 0.0

    grup_kode = df_mb.groupby('Kode', sort=False) if 'Kode' in df_mb.columns else [(None, df_mb)]

    for k_val, df_sub in grup_kode:
        is_first_row = True
        status_cek_list = []
        
        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            
            if kode_item in baris_diabaikan:
                df_mb.at[idx, 'Crosscheck Qty'] = ''
                df_mb.at[idx, 'Qty PIBC'] = ''
                df_mb.at[idx, 'Qty PG'] = ''
                df_mb.at[idx, 'Selisih (GR)'] = ''
                continue
                
            gudang = str(row.get('Gudang', '')).strip().upper()
            nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
            ket = str(row.get('Keterangan', ''))
            ket_lain = str(row.get('Keterangan Lain', ''))
            teks_gab = ket + " " + ket_lain
            
            kode_pibc = ekstrak_kode_pibc(teks_gab, gudang)
            df_mb.at[idx, 'Ext_PIBC'] = kode_pibc
            if not kode_pibc:
                df_mb.at[idx, 'Status PIBC'] = 'KODE PIBC TIDAK ADA'

            ada_keyword_selisih = any(kw in nama_brg_jdi or kw in teks_gab.upper() for kw in keyword_selisih)
            
            if 'PRODUKSI SOFTCONE' in gudang or kode_item.startswith(('TWP', 'MWP', 'TBM', 'MBB')):
                qty_target = row.get('Qty Target Standar (GR)', 0.0)
                qty_bb = row.get('Qty BB Standar (GR)', 0.0)
                selisih = qty_bb - qty_target
                df_mb.at[idx, 'Selisih (GR)'] = selisih
                
                if ada_keyword_selisih:
                    df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH'
                elif abs(selisih) > 0.01:
                    df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                else:
                    df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                df_mb.at[idx, 'Rincian Benang'] = kode_item

            elif 'DYEING STI' in gudang:
                if kode_item != 'TBB':
                    df_mb.at[idx, 'Rincian Benang'] = 'BUKAN BENANG'

            elif 'DYEING' in gudang or 'LAB' in gudang:
                qty_bb = row.get('Qty BB Standar (GR)', 0.0)
                matched = False
                
                if kode_pibc:
                    key_pibc = (kode_pibc.upper(), kode_item.upper())
                    if key_pibc in pibc_lookup:
                        qty_act = pibc_lookup[key_pibc]
                        df_mb.at[idx, 'Qty PIBC'] = qty_act
                        selisih = qty_bb - qty_act
                        df_mb.at[idx, 'Selisih (GR)'] = selisih
                        matched = True
                        if ada_keyword_selisih:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH'
                        elif abs(selisih) > 0.01:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                        else:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else:
                        key_pg = (kode_pibc.upper(), kode_item.upper())
                        if key_pg in pg_lookup:
                            qty_masuk = pg_lookup[key_pg]
                            df_mb.at[idx, 'Qty PG'] = qty_masuk
                            selisih = qty_bb - qty_masuk
                            df_mb.at[idx, 'Selisih (GR)'] = selisih
                            matched = True
                            if ada_keyword_selisih:
                                df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH'
                            elif abs(selisih) > 0.01:
                                df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                            else:
                                df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                
                if not matched:
                    df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'
                    status_cek_list.append(f"Kelebihan [{kode_item}]")

            if is_first_row:
                kode_bj = str(row.get('Kode Barang Jadi', '')).strip()
                if 'MESIN DYEING' in gudang:
                    if not kode_item.startswith(('TWP.050', 'TWP.100', 'TWP.051')) and kode_bj not in ['TWP.100', 'TWP.400', 'TWP.051', 'TWP.666', 'TWP.901']:
                        df_mb.at[idx, 'Cek Variant'] = kode_bj
                    else:
                        df_mb.at[idx, 'Cek Variant'] = 'CLEAN'
                elif 'LAB' in gudang:
                    if not kode_item.startswith(('TWP.050', 'TWP.051', 'TWP.999', 'TWP.100')) and kode_bj not in ['TWP.999', 'TWP.100', 'TWP.400', 'TWP.666', 'TWP.051']:
                        df_mb.at[idx, 'Cek Variant'] = kode_bj
                    else:
                        df_mb.at[idx, 'Cek Variant'] = 'CLEAN'

            is_first_row = False

        idx_first = df_sub.index[0]
        if status_cek_list:
            df_mb.at[idx_first, 'Status Cek'] = ", ".join(list(set(status_cek_list)))
        else:
            df_mb.at[idx_first, 'Status Cek'] = 'COMPLATE'

    return df_mb


def buat_excel_incorrect_multi_sheet(df_hasil):
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


def buat_excel_pibc_tidak_di_mb(df_hasil_mb, df_pibc):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if df_pibc is not None and not df_pibc.empty:
            col_dyelot = cari_kolom(df_pibc, ['dyelot', 'pibc'])
            pibc_di_mb = set(df_hasil_mb['Ext_PIBC'].dropna().astype(str).str.strip().unique()) if 'Ext_PIBC' in df_hasil_mb.columns else set()
            if col_dyelot:
                df_tidak_ada = df_pibc[~df_pibc[col_dyelot].astype(str).str.strip().isin(pibc_di_mb)]
                df_tidak_ada.to_excel(writer, index=False, sheet_name='PIBC Tidak Ada di MB')
            else:
                df_pibc.to_excel(writer, index=False, sheet_name='PIBC Tidak Ada di MB')
        else:
            pd.DataFrame({'Info': ['Master PIBC kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()


def buat_excel_pg_tidak_di_mb(df_hasil_mb, df_pg):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if df_pg is not None and not df_pg.empty:
            col_ket = cari_kolom(df_pg, ['keterangan', 'det', 'pibc'])
            pibc_di_mb = set(df_hasil_mb['Ext_PIBC'].dropna().astype(str).str.strip().unique()) if 'Ext_PIBC' in df_hasil_mb.columns else set()
            if col_ket:
                df_tidak_ada = df_pg[~df_pg[col_ket].astype(str).str.strip().isin(pibc_di_mb)]
                df_tidak_ada.to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
            else:
                df_pg.to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
        else:
            pd.DataFrame({'Info': ['Master PG kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()
