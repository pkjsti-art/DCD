import pandas as pd
import re
import io

BARIS_DIABAIKAN = [
    "Total Material (Details)", "Total Pekerjaan (Works)", 
    "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor",
    "Kode Pekerjaan", "Nama Pekerjaan"
]

def sanitize_key(k):
    if not k: return ""
    k = re.sub(r'[^A-Z0-9\.\-]', '', str(k).upper())
    return k.replace('PIBC-', '')

def normalisasi_gudang(nama_gudang):
    g = str(nama_gudang).strip().lower()
    if 'lab' in g or 'rnd' in g: return 'Gudang Lab & RnD'
    if 'mesin dyeing' in g or ('dyeing' in g and 'sti' not in g): return 'Gudang Mesin Dyeing'
    if 'dyeing sti' in g or 'sti' in g: return 'Gudang Dyeing STI'
    if 'softcone' in g: return 'Gudang Softcone'
    return str(nama_gudang).strip().title()

def cari_kolom_akurat(df, daftar_kemungkinan):
    for kem in daftar_kemungkinan:
        kem_clean = re.sub(r'[^A-Z0-9]', '', kem.upper())
        for col in df.columns:
            col_clean = re.sub(r'[^A-Z0-9]', '', str(col).upper())
            if kem_clean in col_clean: return col
    return None

def bersihkan_angka_indo(val):
    if pd.isna(val) or val is None: return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('.', '').replace(',', '.')
    try: return round(float(val_str), 3)
    except: return 0.0

def is_item_benang(kode_item, gudang=""):
    k_upper = str(kode_item).strip().upper()
    g_norm = normalisasi_gudang(gudang)
    if k_upper.startswith(('TWP', 'MWP', 'TBM', 'MBB')): return True
    if 'Mesin Dyeing' in g_norm or 'Lab & RnD' in g_norm or 'STI' in g_norm: return False
    if k_upper.startswith('TBB'): return True
    return False

def ekstrak_kode_pibc(teks, gudang=""):
    if pd.isna(teks): return None
    teks_str = str(teks)
    pola = r'(PIBC-[A-Za-z0-9#\-\.]+)'
    hasil = re.search(pola, teks_str, re.IGNORECASE)
    if hasil: return hasil.group(1).strip().rstrip('.')
    match_angka = re.search(r'([0-9]{8,}(?:#[A-Za-z0-9]+)?)', teks_str)
    if match_angka: return match_angka.group(1).strip()
    return None

def siapkan_kolom_standar(df_mb):
    col_gudang_mb = cari_kolom_akurat(df_mb, ['GUDANG'])
    if col_gudang_mb and col_gudang_mb != 'Gudang': df_mb.rename(columns={col_gudang_mb: 'Gudang'}, inplace=True)
    elif not col_gudang_mb: df_mb['Gudang'] = ''

    col_target_qty, col_target_unit, col_qty, col_unit, col_kode_item = None, None, None, None, None

    for col in df_mb.columns:
        c_lower = str(col).strip().lower()
        if 'target' in c_lower and ('qty' in c_lower or 'quantity' in c_lower): col_target_qty = col
        elif 'target' in c_lower and 'unit' in c_lower: col_target_unit = col
        elif ('qty' in c_lower or 'quantity' in c_lower) and 'target' not in c_lower and 'bb' not in c_lower: col_qty = col
        elif ('unit' in c_lower or 'satuan' in c_lower) and 'target' not in c_lower: col_unit = col
        elif 'kode item' in c_lower or c_lower == 'kode_item': col_kode_item = col

    if not col_kode_item: col_kode_item = cari_kolom_akurat(df_mb, ['KODEITEM'])

    seri_target_gr = pd.Series([''] * len(df_mb), dtype=object)
    seri_bb_gr = pd.Series([''] * len(df_mb), dtype=object)

    for idx, r in df_mb.iterrows():
        kode_item = str(r[col_kode_item]).strip() if col_kode_item and pd.notna(r[col_kode_item]) else ''
        gudang = str(r.get('Gudang', ''))
        g_norm = normalisasi_gudang(gudang)
        is_benang = is_item_benang(kode_item, gudang)
        
        is_diabaikan = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper() or ('Softcone' in g_norm and not is_benang))

        if not is_diabaikan:
            if col_target_qty and pd.notna(r[col_target_qty]) and str(r[col_target_qty]).strip().lower() != 'nan':
                val = bersihkan_angka_indo(r[col_target_qty])
                unit = str(r[col_target_unit]).strip().upper() if col_target_unit else 'GR'
                seri_target_gr[idx] = round(val if unit == 'GR' else val * 1000.0, 3)
            if col_qty and pd.notna(r[col_qty]) and str(r[col_qty]).strip().lower() != 'nan':
                val = bersihkan_angka_indo(r[col_qty])
                unit = str(r[col_unit]).strip().upper() if col_unit else 'GR'
                seri_bb_gr[idx] = round(val if unit == 'GR' else val * 1000.0, 3)

    idx_target = list(df_mb.columns).index(col_target_unit if col_target_unit else col_target_qty) + 1 if (col_target_unit or col_target_qty) else len(df_mb.columns)
    df_mb.insert(idx_target, 'Qty Target Standar (GR)', seri_target_gr)
    
    idx_bb = list(df_mb.columns).index(col_qty) + 1 if col_qty else len(df_mb.columns)
    df_mb.insert(idx_bb, 'Qty BB Standar (GR)', seri_bb_gr)

    for col_baru in ['Qty PIBC', 'Qty PG', 'Selisih Benang (GR)', 'Selisih PIBC (GR)', 'Selisih PG (GR)', 'Crosscheck Qty', 'Info Celup Gabung', 'Status Cek', 'Status PIBC', 'Rincian Benang', 'Cek Variant', 'Ext_PIBC']:
        if col_baru not in df_mb.columns: df_mb[col_baru] = ''
        df_mb[col_baru] = df_mb[col_baru].astype(object)

    return df_mb

def jalankan_proses_crosscheck(df_mb, df_pibc, df_pg):
    keyword_selisih = ['AVALAN', 'CROCHET', 'KOR', 'KOR ROMBE', 'KOLONG', 'REEBOK', 'ROLL', 'TALI', 'KUR', 'AVL', 'CONS', 'HTC']

    # 1. LOOKUP MASTER PIBC
    pibc_lookup = {}
    if df_pibc is not None and not df_pibc.empty:
        col_dyelot = cari_kolom_akurat(df_pibc, ['DYELOT', 'PIBC'])
        col_act = cari_kolom_akurat(df_pibc, ['ACTUAL', 'QTY', 'JUMLAH'])
        col_kode_obat = cari_kolom_akurat(df_pibc, ['KODEOBAT', 'KODEITEM'])
        col_nama_obat = cari_kolom_akurat(df_pibc, ['NAMA', 'ITEM', 'OBAT'])
        
        if col_dyelot and col_act:
            for _, r in df_pibc.iterrows():
                val_act = bersihkan_angka_indo(r[col_act])
                k_dyelot = sanitize_key(r[col_dyelot])
                k_kode = sanitize_key(r[col_kode_obat]) if col_kode_obat and pd.notna(r[col_kode_obat]) else ""
                k_nama = sanitize_key(r[col_nama_obat]) if col_nama_obat and pd.notna(r[col_nama_obat]) else ""
                
                if k_kode: pibc_lookup[(k_dyelot, k_kode)] = round(pibc_lookup.get((k_dyelot, k_kode), 0.0) + val_act, 3)
                if k_nama: pibc_lookup[(k_dyelot, k_nama)] = round(pibc_lookup.get((k_dyelot, k_nama), 0.0) + val_act, 3)

    # 2. LOOKUP MASTER PG
    pg_lookup = {}
    if df_pg is not None and not df_pg.empty:
        if 'PIBC_PG' in df_pg.columns and 'Qty_PG' in df_pg.columns:
            for _, r in df_pg.iterrows():
                val_msk = bersihkan_angka_indo(r['Qty_PG'])
                if val_msk <= 0: continue 
                
                raw_pibc = str(r['PIBC_PG'])
                if not raw_pibc or raw_pibc.lower() == 'nan': continue
                
                pibc_ext = ekstrak_kode_pibc(raw_pibc)
                if not pibc_ext: continue
                
                k_ket = sanitize_key(pibc_ext)
                k_kode = sanitize_key(r.get('Kode_Item_PG', ''))
                k_nama = sanitize_key(r.get('Nama_Item_PG', ''))
                
                if k_kode: pg_lookup[(k_ket, k_kode)] = round(pg_lookup.get((k_ket, k_kode), 0.0) + val_msk, 3)
                if k_nama: pg_lookup[(k_ket, k_nama)] = round(pg_lookup.get((k_ket, k_nama), 0.0) + val_msk, 3)

    # Ekstrak Ext_PIBC MB
    for idx, row in df_mb.iterrows():
        gudang = str(row.get('Gudang', ''))
        teks_gab = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
        kp = ekstrak_kode_pibc(teks_gab, gudang)
        if kp: df_mb.at[idx, 'Ext_PIBC'] = kp.strip().upper()

    grup_kode = df_mb.groupby('Kode', sort=False) if 'Kode' in df_mb.columns else [(None, df_mb)]

    # 3. PRE-PROCESSING CELUP GABUNG & AKUMULASI QTY OBAT PER TRANSAKSI
    pibc_to_mb = {}
    gabungan_qty_item = {}

    for idx, row in df_mb.iterrows():
        ket_gabung_cek = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
        is_teks_gabung = 'celup gabung' in ket_gabung_cek.lower() or 'gabung' in ket_gabung_cek.lower()
        kp = str(row.get('Ext_PIBC', '')).strip().upper()
        k_val = row.get('Kode', idx)
        
        if is_teks_gabung and kp and kp != '':
            if kp not in pibc_to_mb: pibc_to_mb[kp] = []
            if k_val not in pibc_to_mb[kp]: pibc_to_mb[kp].append(k_val)

            kode_item = str(row.get('Kode Item', '')).strip().upper()
            gudang = str(row.get('Gudang', ''))
            if not is_item_benang(kode_item, gudang) and kode_item not in BARIS_DIABAIKAN and 'PEKERJAAN' not in kode_item:
                val_bb = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                key_item = (kp, kode_item)
                gabungan_qty_item[key_item] = round(gabungan_qty_item.get(key_item, 0.0) + val_bb, 3)

    pibc_celup_gabung = {pibc for pibc, mbs in pibc_to_mb.items() if len(set(mbs)) > 1}

    # 4. PROSES UTAMA CROSSCHECK
    for k_val, df_sub in grup_kode:
        grup_punya_pibc = False
        kode_pibc_grup = None
        target_gr_utama_grup = 0.0
        total_bb_benang_grup = 0.0
        ada_item_non_tbb_sti = False
        rincian_benang_list = []
        item_qty_dalam_transaksi = {} 
        
        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            gudang = str(row.get('Gudang', ''))
            kp_baris = str(row.get('Ext_PIBC', '')).strip()
            
            if not kode_pibc_grup and kp_baris:
                grup_punya_pibc = True
                kode_pibc_grup = kp_baris

            val_target = row.get('Qty Target Standar (GR)')
            if isinstance(val_target, (int, float)) and val_target > 0 and target_gr_utama_grup == 0.0:
                target_gr_utama_grup = round(float(val_target), 3)

            if is_item_benang(kode_item, gudang):
                val_bb = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                total_bb_benang_grup = round(total_bb_benang_grup + val_bb, 3)
                if kode_item not in rincian_benang_list: rincian_benang_list.append(kode_item)
            else:
                if kode_item not in BARIS_DIABAIKAN and 'PEKERJAAN' not in kode_item.upper():
                    val_bb = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                    k_item_upper = kode_item.upper()
                    item_qty_dalam_transaksi[k_item_upper] = round(item_qty_dalam_transaksi.get(k_item_upper, 0.0) + val_bb, 3)

            if 'Dyeing STI' in normalisasi_gudang(gudang):
                if not (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper()) and kode_item and not kode_item.startswith('TBB'):
                    ada_item_non_tbb_sti = True

        rincian_benang_str = ", ".join(rincian_benang_list)
        is_celup_gabung = False
        pasangan_str = ""
        if kode_pibc_grup and kode_pibc_grup.upper() in pibc_celup_gabung:
            is_celup_gabung = True
            semua_mb_terkait = pibc_to_mb.get(kode_pibc_grup.upper(), [])
            pasangan_str = ", ".join([str(m) for m in set(semua_mb_terkait) if m != k_val])

        info_cg_base = f"Celup Gabung dgn: {pasangan_str}" if is_celup_gabung else ""
        is_first_row = True
        
        # Tracker untuk mencegah duplikasi cetak baris
        benang_status_rendered = False
        obat_rendered_set = set()
        status_cek_list = []

        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            nama_item = str(row.get('Nama Item', '')).strip()
            gudang = str(row.get('Gudang', ''))
            g_norm = normalisasi_gudang(gudang)
            benang_flag = is_item_benang(kode_item, gudang)
            kp_baris = str(row.get('Ext_PIBC', '')).strip().upper()

            is_diabaikan = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper() or ('Softcone' in g_norm and not benang_flag))

            df_mb.at[idx, 'Crosscheck Qty'] = ''
            df_mb.at[idx, 'Qty PIBC'] = ''
            df_mb.at[idx, 'Qty PG'] = ''
            df_mb.at[idx, 'Selisih Benang (GR)'] = ''
            df_mb.at[idx, 'Selisih PIBC (GR)'] = ''
            df_mb.at[idx, 'Selisih PG (GR)'] = ''
            df_mb.at[idx, 'Status PIBC'] = ''
            df_mb.at[idx, 'Info Celup Gabung'] = ''

            if is_diabaikan:
                df_mb.at[idx, 'Qty Target Standar (GR)'] = ''
                df_mb.at[idx, 'Qty BB Standar (GR)'] = ''
                df_mb.at[idx, 'Rincian Benang'] = rincian_benang_str if 'Dyeing STI' not in g_norm else ''
                continue

            if 'Dyeing STI' in g_norm and not ada_item_non_tbb_sti:
                df_mb.at[idx, 'Rincian Benang'] = ''
                df_mb.at[idx, 'Cek Variant'] = ''
                is_first_row = False
                continue

            if is_first_row:
                df_mb.at[idx, 'Status PIBC'] = 'KODE PIBC ADA' if grup_punya_pibc else 'KODE PIBC TIDAK ADA'

            df_mb.at[idx, 'Rincian Benang'] = 'BUKAN BENANG' if ('Dyeing STI' in g_norm and not kode_item.startswith('TBB')) else ('' if 'Dyeing STI' in g_norm else rincian_benang_str)

            nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
            teks_gab = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
            ada_keyword_selisih = any(kw in nama_brg_jdi or kw in teks_gab.upper() for kw in keyword_selisih)

            if benang_flag:
                # Hanya cetak di baris benang pertama saja
                if not benang_status_rendered:
                    selisih_total = round(total_bb_benang_grup - target_gr_utama_grup, 3)
                    df_mb.at[idx, 'Selisih Benang (GR)'] = selisih_total

                    if abs(selisih_total) <= 0.01: df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else: df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT' if ('Lab & RnD' in g_norm) else ('MEMANG SELISIH' if ada_keyword_selisih else 'INCORRECT')
                    benang_status_rendered = True
            else:
                k_item_upper = kode_item.upper()
                key_obat = (kp_baris, k_item_upper)
                
                # Hanya cetak jika item ini belum pernah diproses di transaksi ini
                if key_obat not in obat_rendered_set:
                    qty_bb_baris = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                    
                    if kp_baris in pibc_celup_gabung:
                        qty_bb_cek = round(gabungan_qty_item.get((kp_baris, k_item_upper), qty_bb_baris), 3)
                        df_mb.at[idx, 'Info Celup Gabung'] = f"{info_cg_base} (Total: {qty_bb_cek} GR)"
                    else:
                        qty_bb_cek = round(item_qty_dalam_transaksi.get(k_item_upper, qty_bb_baris), 3)

                    if kp_baris:
                        qty_pibc = None
                        qty_pg = None
                        
                        k_ket = sanitize_key(kp_baris)
                        k_kode = sanitize_key(kode_item)
                        k_nama = sanitize_key(nama_item)
                        
                        if (k_ket, k_kode) in pibc_lookup: qty_pibc = pibc_lookup[(k_ket, k_kode)]
                        elif (k_ket, k_nama) in pibc_lookup: qty_pibc = pibc_lookup[(k_ket, k_nama)]

                        if (k_ket, k_kode) in pg_lookup: qty_pg = pg_lookup[(k_ket, k_kode)]
                        elif (k_ket, k_nama) in pg_lookup: qty_pg = pg_lookup[(k_ket, k_nama)]

                        sel_pibc = None
                        sel_pg = None
                        
                        if qty_pibc is not None:
                            df_mb.at[idx, 'Qty PIBC'] = qty_pibc
                            sel_pibc = round(qty_bb_cek - qty_pibc, 3)
                            df_mb.at[idx, 'Selisih PIBC (GR)'] = sel_pibc
                            
                        if qty_pg is not None:
                            df_mb.at[idx, 'Qty PG'] = qty_pg
                            sel_pg = round(qty_bb_cek - qty_pg, 3)
                            df_mb.at[idx, 'Selisih PG (GR)'] = sel_pg

                        status_pibc = "TIDAK ADA" if qty_pibc is None else ("MATCH" if abs(sel_pibc) <= 0.01 else "SELISIH")
                        status_pg = "TIDAK ADA" if qty_pg is None else ("MATCH" if abs(sel_pg) <= 0.01 else "SELISIH")

                        if status_pibc == "TIDAK ADA" and status_pg == "TIDAK ADA":
                            df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'
                            status_cek_list.append(f"Kelebihan [{kode_item}]")
                        elif status_pibc == "MATCH" and status_pg == "MATCH": df_mb.at[idx, 'Crosscheck Qty'] = 'MATCH PIBC & PG'
                        elif status_pibc == "MATCH" and status_pg == "SELISIH": df_mb.at[idx, 'Crosscheck Qty'] = 'MATCH PIBC (PG SELISIH)'
                        elif status_pibc == "MATCH" and status_pg == "TIDAK ADA": df_mb.at[idx, 'Crosscheck Qty'] = 'MATCH PIBC (TIDAK ADA DI PG)'
                        elif status_pg == "MATCH" and status_pibc == "SELISIH": df_mb.at[idx, 'Crosscheck Qty'] = 'MATCH PG (PIBC SELISIH)'
                        elif status_pg == "MATCH" and status_pibc == "TIDAK ADA": df_mb.at[idx, 'Crosscheck Qty'] = 'MATCH PG (TIDAK ADA DI PIBC)'
                        elif status_pibc == "SELISIH" and status_pg == "SELISIH": df_mb.at[idx, 'Crosscheck Qty'] = 'SELISIH KEDUANYA'
                        elif status_pibc == "SELISIH" and status_pg == "TIDAK ADA": df_mb.at[idx, 'Crosscheck Qty'] = 'PIBC SELISIH (TIDAK ADA DI PG)'
                        elif status_pg == "SELISIH" and status_pibc == "TIDAK ADA": df_mb.at[idx, 'Crosscheck Qty'] = 'PG SELISIH (TIDAK ADA DI PIBC)'
                    
                    obat_rendered_set.add(key_obat)

            if is_first_row:
                kode_bj = str(row.get('Kode Barang Jadi', '')).strip()
                if 'Mesin Dyeing' in g_norm:
                    df_mb.at[idx, 'Cek Variant'] = kode_bj if not kode_item.startswith(('TWP.050', 'TWP.100', 'TWP.051')) and kode_bj not in ['TWP.100', 'TWP.400', 'TWP.051', 'TWP.666', 'TWP.901'] else 'CLEAN'
                elif 'Lab & RnD' in g_norm:
                    df_mb.at[idx, 'Cek Variant'] = kode_bj if not kode_item.startswith(('TWP.050', 'TWP.051', 'TWP.999', 'TWP.100')) and kode_bj not in ['TWP.999', 'TWP.100', 'TWP.400', 'TWP.666', 'TWP.051'] else 'CLEAN'

            is_first_row = False

        idx_first = df_sub.index[0]
        gudang_group = ""
        for i in df_sub.index:
            g = str(df_mb.loc[i].get('Gudang', '')).strip()
            if g != "":
                gudang_group = g
                break
        gudang_group_norm = normalisasi_gudang(gudang_group)
        is_softcone_group = 'Softcone' in gudang_group_norm or (gudang_group_norm == '' and str(df_mb.loc[idx_first].get('Kode Item', '')).strip().upper().startswith(('TWP', 'MWP', 'TBM', 'MBB')))

        if ('Dyeing STI' in gudang_group_norm and not ada_item_non_tbb_sti) or is_softcone_group:
            df_mb.at[idx_first, 'Status Cek'] = '' 
        else:
            df_mb.at[idx_first, 'Status Cek'] = ", ".join(list(set(status_cek_list))) if status_cek_list else 'COMPLATE'

    return df_mb

def buat_excel_incorrect_multi_sheet(df_hasil):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_inc = df_hasil[df_hasil['Crosscheck Qty'].astype(str).str.contains('INCORRECT|SELISIH|TIDAK DI TEMUKAN|TIDAK ADA', case=False, na=False)]
        kode_grup_inc = df_inc['Kode'].unique() if 'Kode' in df_inc.columns else []
        df_full_inc = df_hasil[df_hasil['Kode'].isin(kode_grup_inc)] if len(kode_grup_inc) > 0 else df_inc
        df_full_inc.to_excel(writer, index=False, sheet_name='INCORRECT')

        df_selisih = df_hasil[df_hasil['Crosscheck Qty'] == 'MEMANG SELISIH']
        kode_grup_selisih = df_selisih['Kode'].unique() if 'Kode' in df_selisih.columns else []
        df_full_selisih = df_hasil[df_hasil['Kode'].isin(kode_grup_selisih)] if len(kode_grup_selisih) > 0 else df_selisih
        df_full_selisih.to_excel(writer, index=False, sheet_name='INCORRECT Memang Selisih')

        df_benang_inc = df_hasil[(df_hasil['Rincian Benang'] != '') & (df_hasil['Crosscheck Qty'].astype(str).str.contains('INCORRECT|SELISIH', case=False, na=False))]
        kode_grup_benang = df_benang_inc['Kode'].unique() if 'Kode' in df_benang_inc.columns else []
        df_full_benang = df_hasil[df_hasil['Kode'].isin(kode_grup_benang)] if len(kode_grup_benang) > 0 else df_benang_inc
        df_full_benang.to_excel(writer, index=False, sheet_name='INCORRECT Benang')
        
    return output.getvalue()

def buat_excel_pibc_tidak_di_mb(df_hasil_mb, df_pibc):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if df_pibc is not None and not df_pibc.empty:
            col_dyelot = cari_kolom_akurat(df_pibc, ['DYELOT', 'PIBC'])
            pibc_di_mb = set(df_hasil_mb['Ext_PIBC'].dropna().astype(str).apply(sanitize_key).unique()) if 'Ext_PIBC' in df_hasil_mb.columns else set()
            if col_dyelot:
                df_pibc[~df_pibc[col_dyelot].astype(str).apply(sanitize_key).isin(pibc_di_mb)].to_excel(writer, index=False, sheet_name='PIBC Tidak Ada di MB')
            else: df_pibc.to_excel(writer, index=False, sheet_name='PIBC Tidak Ada di MB')
        else: pd.DataFrame({'Info': ['Master PIBC kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()

def buat_excel_pg_tidak_di_mb(df_hasil_mb, df_pg):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if df_pg is not None and not df_pg.empty:
            pibc_di_mb = set(df_hasil_mb['Ext_PIBC'].dropna().astype(str).apply(sanitize_key).unique()) if 'Ext_PIBC' in df_hasil_mb.columns else set()
            if 'PIBC_PG' in df_pg.columns:
                df_pg[~df_pg['PIBC_PG'].astype(str).apply(sanitize_key).isin(pibc_di_mb)].to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
            else: df_pg.to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
        else: pd.DataFrame({'Info': ['Master PG kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()
