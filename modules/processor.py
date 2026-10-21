import pandas as pd
import re
import io

# Daftar prefix item benang resmi yang valid untuk Softcone
PREFIX_BENANG_VALID = ('TBB', 'MBB', 'TBM', 'TWP', 'MWP')

# Daftar baris total/overhead/pekerjaan yang wajib dihiraukan
BARIS_DIABAIKAN = [
    "Total Material (Details)", "Total Pekerjaan (Works)", 
    "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor",
    "Kode Pekerjaan", "Nama Pekerjaan"
]

def cari_kolom_akurat(df, daftar_kemungkinan):
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for kem in daftar_kemungkinan:
            if kem.lower() in col_clean:
                return col
    return None

def bersihkan_angka_indo(val):
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if not val_str or val_str.lower() == 'nan':
        return 0.0
    val_str = val_str.replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

def siapkan_kolom_standar(df_mb):
    col_target_qty = None
    col_target_unit = None
    col_qty = None
    col_unit = None
    col_kode_item = None

    for col in df_mb.columns:
        c_lower = str(col).strip().lower()
        if 'target' in c_lower and ('qty' in c_lower or 'quantity' in c_lower):
            col_target_qty = col
        elif 'target' in c_lower and 'unit' in c_lower:
            col_target_unit = col
        elif c_lower in ['qty', 'quantity'] and 'target' not in c_lower and 'bb' not in c_lower:
            col_qty = col
        elif c_lower in ['unit', 'satuan'] and 'target' not in c_lower:
            col_unit = col
        elif 'kode item' in c_lower or c_lower == 'kode_item':
            col_kode_item = col

    if not col_target_qty:
        col_target_qty = cari_kolom_akurat(df_mb, ['target qty', 'target_qty'])
    if not col_target_unit:
        col_target_unit = cari_kolom_akurat(df_mb, ['target unit', 'target_unit'])
    if not col_qty:
        col_qty = cari_kolom_akurat(df_mb, ['qty'])
    if not col_unit:
        col_unit = cari_kolom_akurat(df_mb, ['unit', 'satuan'])
    if not col_kode_item:
        col_kode_item = cari_kolom_akurat(df_mb, ['kode item', 'kode_item'])

    seri_target_gr = pd.Series([''] * len(df_mb), dtype=object)
    seri_bb_gr = pd.Series([''] * len(df_mb), dtype=object)

    for idx, r in df_mb.iterrows():
        kode_item = str(r[col_kode_item]).strip() if col_kode_item and pd.notna(r[col_kode_item]) else ''
        gudang = str(r.get('Gudang', '')).strip().upper()
        
        is_diabaikan = (
            kode_item in BARIS_DIABAIKAN or
            'PEKERJAAN' in kode_item.upper() or
            ('SOFTCONE' in gudang and not kode_item.startswith(PREFIX_BENANG_VALID))
        )

        if not is_diabaikan:
            # 1. Qty Target Standar (GR)
            if col_target_qty:
                val_raw = r[col_target_qty]
                if pd.notna(val_raw) and str(val_raw).strip() != '' and str(val_raw).strip().lower() != 'nan':
                    val = bersihkan_angka_indo(val_raw)
                    unit = str(r[col_target_unit]).strip().upper() if col_target_unit in df_mb.columns and pd.notna(r[col_target_unit]) else 'GR'
                    seri_target_gr[idx] = val if unit == 'GR' else val * 1000.0

            # 2. Qty BB Standar (GR)
            if col_qty:
                val_raw = r[col_qty]
                if pd.notna(val_raw) and str(val_raw).strip() != '' and str(val_raw).strip().lower() != 'nan':
                    val = bersihkan_angka_indo(val_raw)
                    unit = str(r[col_unit]).strip().upper() if col_unit in df_mb.columns and pd.notna(r[col_unit]) else 'GR'
                    seri_bb_gr[idx] = val if unit == 'GR' else val * 1000.0

    if col_target_unit in df_mb.columns:
        idx_target = list(df_mb.columns).index(col_target_unit) + 1
        df_mb.insert(idx_target, 'Qty Target Standar (GR)', seri_target_gr)
    elif col_target_qty in df_mb.columns:
        idx_target = list(df_mb.columns).index(col_target_qty) + 1
        df_mb.insert(idx_target, 'Qty Target Standar (GR)', seri_target_gr)
    else:
        df_mb['Qty Target Standar (GR)'] = seri_target_gr

    if col_qty in df_mb.columns:
        idx_bb = list(df_mb.columns).index(col_qty) + 1
        df_mb.insert(idx_bb, 'Qty BB Standar (GR)', seri_bb_gr)
    else:
        df_mb['Qty BB Standar (GR)'] = seri_bb_gr

    for col_baru in ['Crosscheck Qty', 'Qty PIBC', 'Qty PG', 'Selisih (GR)', 'Status Cek', 'Status PIBC', 'Rincian Benang', 'Cek Variant']:
        if col_baru not in df_mb.columns:
            df_mb[col_baru] = ''
        df_mb[col_baru] = df_mb[col_baru].astype(object)

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
    keyword_selisih = ['AVALAN', 'CROCHET', 'KOR', 'KOR ROMBE', 'KOLONG', 'REEBOK', 'ROLL', 'TALI', 'KUR', 'AVL', 'CONS', 'HTC']

    pibc_lookup = {}
    if df_pibc is not None and not df_pibc.empty:
        col_dyelot = cari_kolom_akurat(df_pibc, ['dyelot', 'pibc'])
        col_item = cari_kolom_akurat(df_pibc, ['obat', 'item', 'nama'])
        col_act = cari_kolom_akurat(df_pibc, ['actual', 'qty'])
        if col_dyelot and col_item and col_act:
            for _, r in df_pibc.iterrows():
                k_dyelot = str(r[col_dyelot]).strip().upper()
                k_item = str(r[col_item]).strip().upper()
                val_act = bersihkan_angka_indo(r[col_act])
                pibc_lookup[(k_dyelot, k_item)] = val_act
                pibc_lookup[(k_dyelot.replace('PIBC-', ''), k_item)] = val_act

    pg_lookup = {}
    if df_pg is not None and not df_pg.empty:
        col_ket = cari_kolom_akurat(df_pg, ['keterangan', 'det', 'pibc'])
        col_var = cari_kolom_akurat(df_pg, ['variant', 'item', 'nama'])
        col_masuk = cari_kolom_akurat(df_pg, ['masuk', 'qty'])
        if col_ket and col_var and col_masuk:
            for _, r in df_pg.iterrows():
                k_ket = str(r[col_ket]).strip().upper()
                k_var = str(r[col_var]).strip().upper()
                val_msk = bersihkan_angka_indo(r[col_masuk])
                pg_lookup[(k_ket, k_var)] = val_msk

    grup_kode = df_mb.groupby('Kode', sort=False) if 'Kode' in df_mb.columns else [(None, df_mb)]

    for k_val, df_sub in grup_kode:
        grup_punya_pibc = False
        kode_pibc_grup = None
        target_gr_utama_grup = 0.0
        total_bb_softcone_grup = 0.0
        ada_item_non_tbb_sti = False
        
        # 1. Scan Awal Kelompok
        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            gudang = str(row.get('Gudang', '')).strip().upper()
            ket = str(row.get('Keterangan', ''))
            ket_lain = str(row.get('Keterangan Lain', ''))
            teks_gab = ket + " " + ket_lain
            
            if not kode_pibc_grup:
                kp = ekstrak_kode_pibc(teks_gab, gudang)
                if kp:
                    grup_punya_pibc = True
                    kode_pibc_grup = kp

            val_target = row.get('Qty Target Standar (GR)')
            if isinstance(val_target, (int, float)) and val_target > 0 and target_gr_utama_grup == 0.0:
                target_gr_utama_grup = float(val_target)

            if 'SOFTCONE' in gudang and kode_item.startswith(PREFIX_BENANG_VALID):
                val_bb = row.get('Qty BB Standar (GR)')
                if isinstance(val_bb, (int, float)):
                    total_bb_softcone_grup += float(val_bb)

            if 'DYEING STI' in gudang:
                is_diabaikan_check = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper())
                if not is_diabaikan_check and kode_item and not kode_item.startswith('TBB'):
                    ada_item_non_tbb_sti = True

        status_pibc_grup = 'KODE PIBC ADA' if grup_punya_pibc else 'KODE PIBC TIDAK ADA'
        is_first_row = True
        softcone_status_rendered = False
        status_cek_list = []

        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            gudang = str(row.get('Gudang', '')).strip().upper()

            # BARIS DIHIRAUKAN (Total, Pekerjaan, T.AV di Softcone, dll.)
            is_diabaikan = (
                kode_item in BARIS_DIABAIKAN or
                'PEKERJAAN' in kode_item.upper() or
                ('SOFTCONE' in gudang and not kode_item.startswith(PREFIX_BENANG_VALID))
            )

            if is_diabaikan:
                df_mb.at[idx, 'Crosscheck Qty'] = ''
                df_mb.at[idx, 'Qty PIBC'] = ''
                df_mb.at[idx, 'Qty PG'] = ''
                df_mb.at[idx, 'Selisih (GR)'] = ''
                df_mb.at[idx, 'Qty Target Standar (GR)'] = ''
                df_mb.at[idx, 'Qty BB Standar (GR)'] = ''
                df_mb.at[idx, 'Status PIBC'] = ''
                df_mb.at[idx, 'Ext_PIBC'] = ''
                continue

            # GUDANG DYEING STI KOSONG JIKA HANYA TBB
            if 'DYEING STI' in gudang and not ada_item_non_tbb_sti:
                df_mb.at[idx, 'Status PIBC'] = ''
                df_mb.at[idx, 'Ext_PIBC'] = ''
                df_mb.at[idx, 'Crosscheck Qty'] = ''
                df_mb.at[idx, 'Rincian Benang'] = ''
                df_mb.at[idx, 'Qty PIBC'] = ''
                df_mb.at[idx, 'Qty PG'] = ''
                df_mb.at[idx, 'Selisih (GR)'] = ''
                df_mb.at[idx, 'Cek Variant'] = ''
                is_first_row = False
                continue

            # PENGECEKAN NORMAL
            df_mb.at[idx, 'Status PIBC'] = status_pibc_grup
            df_mb.at[idx, 'Ext_PIBC'] = kode_pibc_grup

            nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
            ket = str(row.get('Keterangan', ''))
            ket_lain = str(row.get('Keterangan Lain', ''))
            teks_gab = ket + " " + ket_lain
            
            ada_keyword_selisih = any(kw in nama_brg_jdi or kw in teks_gab.upper() for kw in keyword_selisih)

            if 'DYEING STI' in gudang and not kode_item.startswith('TBB'):
                df_mb.at[idx, 'Rincian Benang'] = 'BUKAN BENANG'
            
            # A. LOGIKA PRODUKSI SOFTCONE
            if 'PRODUKSI SOFTCONE' in gudang or (gudang == '' and kode_item.startswith(PREFIX_BENANG_VALID)):
                df_mb.at[idx, 'Rincian Benang'] = kode_item
                val_bb = row.get('Qty BB Standar (GR)')
                qty_bb_baris = float(val_bb) if isinstance(val_bb, (int, float)) else 0.0
                
                selisih_baris = qty_bb_baris - target_gr_utama_grup
                df_mb.at[idx, 'Selisih (GR)'] = selisih_baris

                if not softcone_status_rendered:
                    selisih_total = total_bb_softcone_grup - target_gr_utama_grup
                    if abs(selisih_total) <= 0.01:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else:
                        if ada_keyword_selisih:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH'
                        else:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                    softcone_status_rendered = True
                else:
                    df_mb.at[idx, 'Crosscheck Qty'] = ''

            # B. LOGIKA GUDANG MESIN DYEING & LAB
            else:
                val_bb = row.get('Qty BB Standar (GR)')
                qty_bb = float(val_bb) if isinstance(val_bb, (int, float)) else 0.0
                
                # B1. BENANG DI MESIN DYEING
                if 'MESIN DYEING' in gudang and kode_item.startswith(PREFIX_BENANG_VALID):
                    df_mb.at[idx, 'Rincian Benang'] = kode_item
                    selisih = qty_bb - target_gr_utama_grup
                    df_mb.at[idx, 'Selisih (GR)'] = selisih
                    if abs(selisih) <= 0.01:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else:
                        if ada_keyword_selisih:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH'
                        else:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'

                # B2. CHEMICAL & DYESTUFF
                else:
                    # ATURAN BARU: JIKA KODE PIBC TIDAK ADA -> HENTIKAN CEK CHEMICAL & DYESTUFF (KOSONGKAN)
                    if not kode_pibc_grup:
                        df_mb.at[idx, 'Qty PIBC'] = ''
                        df_mb.at[idx, 'Qty PG'] = ''
                        df_mb.at[idx, 'Selisih (GR)'] = ''
                        df_mb.at[idx, 'Crosscheck Qty'] = ''
                    else:
                        matched = False
                        key_pibc = (kode_pibc_grup.upper(), kode_item.upper())
                        key_pibc_alt = (kode_pibc_grup.upper().replace('PIBC-', ''), kode_item.upper())
                        
                        if key_pibc in pibc_lookup:
                            qty_act = pibc_lookup[key_pibc]
                            df_mb.at[idx, 'Qty PIBC'] = qty_act
                            selisih = qty_bb - qty_act
                            df_mb.at[idx, 'Selisih (GR)'] = selisih
                            matched = True
                        elif key_pibc_alt in pibc_lookup:
                            qty_act = pibc_lookup[key_pibc_alt]
                            df_mb.at[idx, 'Qty PIBC'] = qty_act
                            selisih = qty_bb - qty_act
                            df_mb.at[idx, 'Selisih (GR)'] = selisih
                            matched = True
                        else:
                            key_pg = (kode_pibc_grup.upper(), kode_item.upper())
                            if key_pg in pg_lookup:
                                qty_masuk = pg_lookup[key_pg]
                                df_mb.at[idx, 'Qty PG'] = qty_masuk
                                selisih = qty_bb - qty_masuk
                                df_mb.at[idx, 'Selisih (GR)'] = selisih
                                matched = True

                        if matched:
                            if abs(selisih) <= 0.01:
                                df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                            else:
                                df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                        else:
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
        gudang_first = str(df_mb.loc[idx_first].get('Gudang', '')).strip().upper()
        if 'DYEING STI' in gudang_first and not ada_item_non_tbb_sti:
            df_mb.at[idx_first, 'Status Cek'] = ''
        else:
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
            col_dyelot = cari_kolom_akurat(df_pibc, ['dyelot', 'pibc'])
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
            col_ket = cari_kolom_akurat(df_pg, ['keterangan', 'det', 'pibc'])
            pibc_di_mb = set(df_hasil_mb['Ext_PIBC'].dropna().astype(str).str.strip().unique()) if 'Ext_PIBC' in df_hasil_mb.columns else set()
            if col_ket:
                df_tidak_ada = df_pg[~df_pg[col_ket].astype(str).str.strip().isin(pibc_di_mb)]
                df_tidak_ada.to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
            else:
                df_pg.to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
        else:
            pd.DataFrame({'Info': ['Master PG kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()
