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
    k = re.sub(r'[^A-Z0-9\.#\+]', '', str(k).upper())
    return k.replace('PIBC', '')

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

def filter_plus(teks):
    if pd.isna(teks): return ""
    teks = str(teks)
    teks = re.sub(r'((?:^|[,;/])\s*)\+', r'\g<1>PLUS_SAH', teks)
    teks = teks.replace('+', ' ')
    teks = teks.replace('PLUS_SAH', '+')
    return teks

POLA_PIBC = r'((?:\bTU[\-\s]+|\+[\s]*)?PIBC[\-\s]*[0-9]{6,}(?:#[A-Za-z0-9]+|-[A-Za-z0-9]+)?|[0-9]{8,}(?:#[A-Za-z0-9]+|-[A-Za-z0-9]+)?)'

def is_to_reprocess(teks, gudang):
    g_norm = normalisasi_gudang(gudang)
    if 'Mesin Dyeing' not in g_norm and 'Dyeing STI' not in g_norm:
        return False
    t_upper = str(teks).upper()
    keywords = ['REPROCESS', '+D', '+C', '+ D', '+ C']
    return any(kw in t_upper for kw in keywords)

def ekstrak_pibc_lab(teks):
    teks_str = str(teks)
    parts = re.split(r'\bREPROCESS\b', teks_str, flags=re.IGNORECASE)
    if len(parts) > 1:
        bagian_setelah = parts[-1]
        hasil = re.findall(POLA_PIBC, bagian_setelah, re.IGNORECASE)
        if hasil: return list(dict.fromkeys([x.strip().rstrip('.') for x in hasil]))
        bagian_sebelum = parts[0]
        hasil = re.findall(POLA_PIBC, bagian_sebelum, re.IGNORECASE)
        if hasil: return list(dict.fromkeys([x.strip().rstrip('.') for x in hasil]))
    return []

def ekstrak_pibc_to(teks):
    teks_str = str(teks)
    hasil = []
    chunks_kurung = re.findall(r'\([^)]*(?:\)|$)', teks_str)
    if chunks_kurung:
        for chunk in chunks_kurung:
            kodes = re.findall(POLA_PIBC, chunk, re.IGNORECASE)
            hasil.extend(kodes)
    if not hasil and re.search(r'\bTO\b', teks_str, re.IGNORECASE):
        parts = re.split(r'\bTO\b', teks_str, flags=re.IGNORECASE)
        if len(parts) > 1:
            bagian_setelah_to = parts[-1]
            kodes = re.findall(POLA_PIBC, bagian_setelah_to, re.IGNORECASE)
            hasil.extend(kodes)
    return list(dict.fromkeys([x.strip().rstrip('.') for x in hasil]))

def ekstrak_semua_pibc(teks):
    teks_str = str(teks)
    hasil = re.findall(POLA_PIBC, teks_str, re.IGNORECASE)
    if hasil: return list(dict.fromkeys([x.strip().rstrip('.') for x in hasil]))
    return []

def dapatkan_pibc_baris(teks, gudang):
    if not teks.strip(): return []
    g_norm = normalisasi_gudang(gudang)
    if 'Lab & RnD' in g_norm and 'REPROCESS' in str(teks).upper(): return ekstrak_pibc_lab(teks)
    if is_to_reprocess(teks, gudang): return ekstrak_pibc_to(teks)
    return ekstrak_semua_pibc(teks)

# =============================================================================
# LOGIKA SMART DETECTIVE VARIANT (Cek Semua Baris & Indikasi Keyword)
# =============================================================================
def cek_kombinasi_variant(ki_full, kb_full, is_to, teks_gab, gudang_norm):
    ki = str(ki_full).strip().upper()
    kb = str(kb_full).strip().upper()
    tg = str(teks_gab).upper()
    
    if not ki: return 'CLEAN'
    base_err = f"{ki_full} -> {kb_full}"
    
    if 'Mesin Dyeing' in gudang_norm:
        if not is_to:
            if ki.startswith(('TWP.050', 'TWP.051', 'MWP', 'TBM', 'TWP.400')) and kb.startswith('TWP.100'): return 'CLEAN'
            if ki.startswith('TWP.051') and kb.startswith('TWP.051'): return 'CLEAN'
            if ki.startswith(('TWP.100', 'TWP.90', 'TWP.400')) and kb.startswith('TWP.400'): return 'CLEAN'
            if ki.startswith(('TWP.100', 'TWP.90', 'TWP.051')) and kb.startswith('TWP.666'): return 'CLEAN'
            if ki.startswith('TWP.050') and kb.startswith('TWP.90'): return 'CLEAN'
            
            # Smart Detective lapis 2 (Cek jika ternyata ini aturan TO yang kurang keyword)
            if (ki.startswith('TWP.100') and kb.startswith('TWP.100')) or (ki.startswith('TWP.90') and kb.startswith('TWP.90')):
                return f"INDIKASI TO (Kurang Keyword Reprocess/+C/+D) : {base_err}"
            return base_err
        else:
            if ki.startswith('TWP.100') and kb.startswith('TWP.100'): return 'CLEAN'
            if ki.startswith('TWP.90') and kb.startswith('TWP.90'): return 'CLEAN'
            return base_err
            
    elif 'Lab & RnD' in gudang_norm:
        if not is_to:
            if ki.startswith(('TWP.050', 'MWP')) and kb.startswith('TWP.999'): return 'CLEAN'
            if ki.startswith('TWP.999') and kb.startswith('TWP.400'): return 'CLEAN'
            if ki.startswith('TWP.400') and kb.startswith('TWP.400'): return 'CLEAN'
            if ki.startswith(('TWP.999', 'TWP.051', 'TWP.100')) and kb.startswith('TWP.666'): return 'CLEAN'
            
            if ki.startswith('TWP.100') and kb.startswith('TWP.400'):
                if 'RND' in tg: return 'CLEAN'
                return f"KURANG KEYWORD 'RND' : {base_err}"
            
            # ATURAN BARU DITAMBAHKAN DI SINI
            if ki.startswith('TWP.050') and kb.startswith('TWP.100'):
                if 'RND' in tg: return 'CLEAN'
                return f"KURANG KEYWORD 'RND' : {base_err}"
                
            # Smart Detective lapis 2
            if ki.startswith('TWP.999') and kb.startswith('TWP.999'):
                if 'LAB' in tg: return f"INDIKASI TO (Kurang Keyword Reprocess/+C/+D) : {base_err}"
                return f"INDIKASI TO & LAB (Kurang Keyword) : {base_err}"
            if ki.startswith('TWP.100') and kb.startswith('TWP.100'):
                if 'RND' in tg: return f"INDIKASI TO (Kurang Keyword Reprocess/+C/+D) : {base_err}"
                return f"INDIKASI TO & RND (Kurang Keyword) : {base_err}"
            return base_err
        else:
            if ki.startswith('TWP.999') and kb.startswith('TWP.999'):
                if 'LAB' in tg: return 'CLEAN'
                return f"KURANG KEYWORD 'LAB' : {base_err}"
            if ki.startswith('TWP.100') and kb.startswith('TWP.100'):
                if 'RND' in tg: return 'CLEAN'
                return f"KURANG KEYWORD 'RND' : {base_err}"
            return base_err
            
    return 'CLEAN'

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

    df_mb['Ext_PIBC_List'] = ""
    df_mb['Ext_PIBC_List'] = df_mb['Ext_PIBC_List'].astype(object)

    return df_mb

def jalankan_proses_crosscheck(df_mb, df_pibc, df_pg):
    keyword_selisih = ['AVALAN', 'CROCHET', 'KOR', 'KOR ROMBE', 'KOLONG', 'REEBOK', 'ROLL', 'TALI', 'KUR', 'AVL', 'CONS', 'HTC']

    pibc_available_keys = set()
    pg_available_keys = set()
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
                
                if k_dyelot: pibc_available_keys.add(k_dyelot)
                if k_kode: pibc_lookup[(k_dyelot, k_kode)] = round(pibc_lookup.get((k_dyelot, k_kode), 0.0) + val_act, 3)
                if k_nama: pibc_lookup[(k_dyelot, k_nama)] = round(pibc_lookup.get((k_dyelot, k_nama), 0.0) + val_act, 3)

    pg_lookup = {}
    if df_pg is not None and not df_pg.empty:
        if 'PIBC_PG' in df_pg.columns and 'Qty_PG' in df_pg.columns:
            for _, r in df_pg.iterrows():
                val_msk = bersihkan_angka_indo(r['Qty_PG'])
                if val_msk <= 0: continue 
                raw_pibc = str(r['PIBC_PG'])
                if not raw_pibc or raw_pibc.lower() == 'nan': continue
                
                pibc_ext_list = ekstrak_semua_pibc(raw_pibc)
                for p_raw in pibc_ext_list:
                    k_ket = sanitize_key(p_raw)
                    if k_ket: pg_available_keys.add(k_ket)
                
                if not pibc_ext_list: continue
                k_ket_utama = sanitize_key(pibc_ext_list[0])
                k_kode = sanitize_key(r.get('Kode_Item_PG', ''))
                k_nama = sanitize_key(r.get('Nama_Item_PG', ''))
                
                if k_kode: pg_lookup[(k_ket_utama, k_kode)] = round(pg_lookup.get((k_ket_utama, k_kode), 0.0) + val_msk, 3)
                if k_nama: pg_lookup[(k_ket_utama, k_nama)] = round(pg_lookup.get((k_ket_utama, k_nama), 0.0) + val_msk, 3)

    for idx, row in df_mb.iterrows():
        gudang = str(row.get('Gudang', ''))
        ket1 = filter_plus(row.get('Keterangan', ''))
        ket2 = filter_plus(row.get('Keterangan Lain', ''))
        
        pibc_list1 = dapatkan_pibc_baris(ket1, gudang)
        pibc_list2 = dapatkan_pibc_baris(ket2, gudang)
        pibc_list = list(dict.fromkeys(pibc_list1 + pibc_list2))
        
        if pibc_list: 
            df_mb.at[idx, 'Ext_PIBC'] = ", ".join(pibc_list).upper()
            df_mb.at[idx, 'Ext_PIBC_List'] = pibc_list

    grup_kode = df_mb.groupby('Kode', sort=False) if 'Kode' in df_mb.columns else [(None, df_mb)]

    pibc_to_mb = {}
    gabungan_qty_item = {}

    for idx, row in df_mb.iterrows():
        ket_gabung_cek = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
        is_teks_gabung = 'celup gabung' in ket_gabung_cek.lower() or 'gabung' in ket_gabung_cek.lower()
        kp_list = row.get('Ext_PIBC_List')
        k_val = row.get('Kode', idx)
        
        kp_utama = sanitize_key(kp_list[0]) if isinstance(kp_list, list) and len(kp_list) > 0 else ""
        
        if is_teks_gabung and kp_utama != '':
            if kp_utama not in pibc_to_mb: pibc_to_mb[kp_utama] = []
            if k_val not in pibc_to_mb[kp_utama]: pibc_to_mb[kp_utama].append(k_val)

            kode_item = str(row.get('Kode Item', '')).strip().upper()
            gudang = str(row.get('Gudang', ''))
            if not is_item_benang(kode_item, gudang) and kode_item not in BARIS_DIABAIKAN and 'PEKERJAAN' not in kode_item:
                val_bb = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                key_item = (kp_utama, kode_item)
                gabungan_qty_item[key_item] = round(gabungan_qty_item.get(key_item, 0.0) + val_bb, 3)

    pibc_celup_gabung = {pibc for pibc, mbs in pibc_to_mb.items() if len(set(mbs)) > 1}

    for k_val, df_sub in grup_kode:
        grup_punya_pibc = False
        pibc_list_grup = []
        target_gr_utama_grup = 0.0
        total_bb_benang_grup = 0.0
        ada_item_non_tbb_sti = False
        rincian_benang_list = []
        item_qty_dalam_transaksi = {} 
        
        teks_gab_group = ""
        gudang_group = ""
        
        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            gudang_str = str(row.get('Gudang', ''))
            if not gudang_group and gudang_str.strip():
                gudang_group = gudang_str.strip()
            
            teks_gab_group += " " + str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
            
            if not grup_punya_pibc:
                kp_l = row.get('Ext_PIBC_List')
                if isinstance(kp_l, list) and len(kp_l) > 0:
                    grup_punya_pibc = True
                    pibc_list_grup = kp_l

            val_target = row.get('Qty Target Standar (GR)')
            if isinstance(val_target, (int, float)) and val_target > 0 and target_gr_utama_grup == 0.0:
                target_gr_utama_grup = round(float(val_target), 3)

            if is_item_benang(kode_item, gudang_str):
                val_bb = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                total_bb_benang_grup = round(total_bb_benang_grup + val_bb, 3)
                if kode_item not in rincian_benang_list: rincian_benang_list.append(kode_item)
            else:
                if kode_item not in BARIS_DIABAIKAN and 'PEKERJAAN' not in kode_item.upper():
                    val_bb = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                    k_item_upper = kode_item.upper()
                    item_qty_dalam_transaksi[k_item_upper] = round(item_qty_dalam_transaksi.get(k_item_upper, 0.0) + val_bb, 3)

            if 'Dyeing STI' in normalisasi_gudang(gudang_str):
                if not (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper()) and kode_item and not kode_item.startswith('TBB'):
                    ada_item_non_tbb_sti = True

        gudang_group_norm = normalisasi_gudang(gudang_group)
        is_to_group = is_to_reprocess(teks_gab_group, gudang_group_norm)
        
        rincian_benang_str = ", ".join(rincian_benang_list)
        is_celup_gabung = False
        pasangan_str = ""
        kp_grup_utama = sanitize_key(pibc_list_grup[0]) if len(pibc_list_grup) > 0 else ""
        
        if kp_grup_utama and kp_grup_utama in pibc_celup_gabung:
            is_celup_gabung = True
            semua_mb_terkait = pibc_to_mb.get(kp_grup_utama, [])
            pasangan_str = ", ".join([str(m) for m in set(semua_mb_terkait) if m != k_val])

        info_cg_base = f"Celup Gabung dgn: {pasangan_str}" if is_celup_gabung else ""
        is_first_row = True
        
        benang_status_rendered = False
        obat_rendered_set = set()
        status_cek_list = []
        error_variants = []

        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            nama_item = str(row.get('Nama Item', '')).strip()
            gudang = str(row.get('Gudang', ''))
            g_norm = normalisasi_gudang(gudang)
            benang_flag = is_item_benang(kode_item, gudang)

            is_diabaikan = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper() or ('Softcone' in g_norm and not benang_flag))

            if is_diabaikan:
                df_mb.at[idx, 'Qty Target Standar (GR)'] = ''
                df_mb.at[idx, 'Qty BB Standar (GR)'] = ''
                df_mb.at[idx, 'Rincian Benang'] = rincian_benang_str if 'Dyeing STI' not in g_norm else ''
                continue

            if 'Dyeing STI' in g_norm and not ada_item_non_tbb_sti:
                df_mb.at[idx, 'Rincian Benang'] = ''
                is_first_row = False
                continue

            # Kumpulkan Cek Variant semua baris benang
            if benang_flag and not is_diabaikan:
                if 'Mesin Dyeing' in g_norm or 'Lab & RnD' in g_norm:
                     res = cek_kombinasi_variant(kode_item, row.get('Kode Barang Jadi', ''), is_to_group, teks_gab_group, g_norm)
                     if res != 'CLEAN' and res not in error_variants:
                         error_variants.append(res)

            if is_first_row:
                if not grup_punya_pibc:
                    df_mb.at[idx, 'Status PIBC'] = 'Kode PIBC di MB Tidak Ada'
                else:
                    miss_pibc = [p for p in pibc_list_grup if sanitize_key(p) not in pibc_available_keys]
                    miss_pg = [p for p in pibc_list_grup if sanitize_key(p) not in pg_available_keys]
                    
                    if not miss_pibc and not miss_pg:
                        df_mb.at[idx, 'Status PIBC'] = 'Kode PIBC di MB Ada'
                    elif miss_pibc and miss_pg:
                        if set(miss_pibc) == set(pibc_list_grup) and set(miss_pg) == set(pibc_list_grup):
                            df_mb.at[idx, 'Status PIBC'] = 'Kode PIBC di MB Ada, di Master PIBC & PG Tidak Ada'
                        else:
                            gabung_miss = list(dict.fromkeys(miss_pibc + miss_pg))
                            df_mb.at[idx, 'Status PIBC'] = f'Kode PIBC di MB Ada (Namun [{", ".join(gabung_miss)}] di Master PIBC & PG Tidak Ada)'
                    elif miss_pibc:
                        if set(miss_pibc) == set(pibc_list_grup):
                            df_mb.at[idx, 'Status PIBC'] = 'Kode PIBC di MB Ada, di Master PIBC Tidak Ada'
                        else:
                            df_mb.at[idx, 'Status PIBC'] = f'Kode PIBC di MB Ada (Namun [{", ".join(miss_pibc)}] di Master PIBC Tidak Ada)'
                    elif miss_pg:
                        if set(miss_pg) == set(pibc_list_grup):
                            df_mb.at[idx, 'Status PIBC'] = 'Kode PIBC di MB Ada, di Master PG Tidak Ada'
                        else:
                            df_mb.at[idx, 'Status PIBC'] = f'Kode PIBC di MB Ada (Namun [{", ".join(miss_pg)}] di Master PG Tidak Ada)'

            df_mb.at[idx, 'Rincian Benang'] = 'BUKAN BENANG' if ('Dyeing STI' in g_norm and not kode_item.startswith('TBB')) else ('' if 'Dyeing STI' in g_norm else rincian_benang_str)

            nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
            tg = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
            ada_keyword_selisih = any(kw in nama_brg_jdi or kw in tg.upper() for kw in keyword_selisih)

            if benang_flag:
                if not benang_status_rendered:
                    selisih_total = round(total_bb_benang_grup - target_gr_utama_grup, 3)
                    df_mb.at[idx, 'Selisih Benang (GR)'] = selisih_total
                    if abs(selisih_total) <= 0.01: df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else: df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT' if ('Lab & RnD' in g_norm) else ('MEMANG SELISIH' if ada_keyword_selisih else 'INCORRECT')
                    benang_status_rendered = True
            else:
                k_item_upper = kode_item.upper()
                if k_item_upper not in obat_rendered_set:
                    qty_bb_baris = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                    
                    if kp_grup_utama in pibc_celup_gabung:
                        qty_bb_cek = round(gabungan_qty_item.get((kp_grup_utama, k_item_upper), qty_bb_baris), 3)
                        df_mb.at[idx, 'Info Celup Gabung'] = f"{info_cg_base} (Total: {qty_bb_cek} GR)"
                    else:
                        qty_bb_cek = round(item_qty_dalam_transaksi.get(k_item_upper, qty_bb_baris), 3)

                    if len(pibc_list_grup) > 0:
                        qty_pibc_total = 0.0
                        qty_pg_total = 0.0
                        found_in_pibc = False
                        found_in_pg = False
                        
                        k_kode = sanitize_key(kode_item)
                        k_nama = sanitize_key(nama_item)
                        
                        for pibc_raw in pibc_list_grup:
                            kp_san = sanitize_key(pibc_raw)
                            if (kp_san, k_kode) in pibc_lookup:
                                qty_pibc_total += pibc_lookup[(kp_san, k_kode)]
                                found_in_pibc = True
                            elif (kp_san, k_nama) in pibc_lookup:
                                qty_pibc_total += pibc_lookup[(kp_san, k_nama)]
                                found_in_pibc = True

                            if (kp_san, k_kode) in pg_lookup:
                                qty_pg_total += pg_lookup[(kp_san, k_kode)]
                                found_in_pg = True
                            elif (kp_san, k_nama) in pg_lookup:
                                qty_pg_total += pg_lookup[(kp_san, k_nama)]
                                found_in_pg = True

                        sel_pibc = None
                        sel_pg = None
                        
                        if found_in_pibc:
                            qty_pibc_total = round(qty_pibc_total, 3)
                            df_mb.at[idx, 'Qty PIBC'] = qty_pibc_total
                            sel_pibc = round(qty_bb_cek - qty_pibc_total, 3)
                            df_mb.at[idx, 'Selisih PIBC (GR)'] = sel_pibc
                        if found_in_pg:
                            qty_pg_total = round(qty_pg_total, 3)
                            df_mb.at[idx, 'Qty PG'] = qty_pg_total
                            sel_pg = round(qty_bb_cek - qty_pg_total, 3)
                            df_mb.at[idx, 'Selisih PG (GR)'] = sel_pg

                        status_pibc = "TIDAK ADA" if not found_in_pibc else ("MATCH" if abs(sel_pibc) <= 0.01 else "SELISIH")
                        status_pg = "TIDAK ADA" if not found_in_pg else ("MATCH" if abs(sel_pg) <= 0.01 else "SELISIH")

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
                    
                    obat_rendered_set.add(k_item_upper)

            is_first_row = False

        idx_first = df_sub.index[0]
        is_softcone_group = 'Softcone' in gudang_group_norm or (gudang_group_norm == '' and str(df_mb.loc[idx_first].get('Kode Item', '')).strip().upper().startswith(('TWP', 'MWP', 'TBM', 'MBB')))

        if ('Dyeing STI' in gudang_group_norm and not ada_item_non_tbb_sti) or is_softcone_group:
            df_mb.at[idx_first, 'Status Cek'] = '' 
        else:
            df_mb.at[idx_first, 'Status Cek'] = ", ".join(list(set(status_cek_list))) if status_cek_list else 'COMPLATE'
            
        if 'Mesin Dyeing' in gudang_group_norm or 'Lab & RnD' in gudang_group_norm:
            df_mb.at[idx_first, 'Cek Variant'] = ", ".join(error_variants) if error_variants else 'CLEAN'
            
    if 'Ext_PIBC_List' in df_mb.columns:
        df_mb = df_mb.drop(columns=['Ext_PIBC_List'])

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
            pibc_di_mb_expanded = set()
            for p_str in pibc_di_mb:
                for p in p_str.split(','): pibc_di_mb_expanded.add(p.strip())

            if col_dyelot:
                df_pibc[~df_pibc[col_dyelot].astype(str).apply(sanitize_key).isin(pibc_di_mb_expanded)].to_excel(writer, index=False, sheet_name='PIBC Tidak Ada di MB')
            else: df_pibc.to_excel(writer, index=False, sheet_name='PIBC Tidak Ada di MB')
        else: pd.DataFrame({'Info': ['Master PIBC kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()

def buat_excel_pg_tidak_di_mb(df_hasil_mb, df_pg):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if df_pg is not None and not df_pg.empty:
            pibc_di_mb = set(df_hasil_mb['Ext_PIBC'].dropna().astype(str).apply(sanitize_key).unique()) if 'Ext_PIBC' in df_hasil_mb.columns else set()
            pibc_di_mb_expanded = set()
            for p_str in pibc_di_mb:
                for p in p_str.split(','): pibc_di_mb_expanded.add(p.strip())
                    
            if 'PIBC_PG' in df_pg.columns:
                df_pg[~df_pg['PIBC_PG'].astype(str).apply(sanitize_key).isin(pibc_di_mb_expanded)].to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
            else: df_pg.to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
        else: pd.DataFrame({'Info': ['Master PG kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()
