import pandas as pd
import re
import io
from openpyxl.styles import PatternFill, Font, Alignment

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

def is_to_reprocess(teks_raw):
    t_upper = str(teks_raw).upper()
    if 'REPROCESS' in t_upper: return True
    if re.search(r'\+\s*[CD](?![A-Z])', t_upper): return True
    return False

def ekstrak_pibc_lab(teks_filtered):
    teks_str = str(teks_filtered)
    parts = re.split(r'\bREPROCESS\b', teks_str, flags=re.IGNORECASE)
    if len(parts) > 1:
        bagian_setelah = parts[-1]
        hasil = re.findall(POLA_PIBC, bagian_setelah, re.IGNORECASE)
        if hasil: return list(dict.fromkeys([x.strip().rstrip('.') for x in hasil]))
        bagian_sebelum = parts[0]
        hasil = re.findall(POLA_PIBC, bagian_sebelum, re.IGNORECASE)
        if hasil: return list(dict.fromkeys([x.strip().rstrip('.') for x in hasil]))
    return []

def ekstrak_pibc_to(teks_filtered):
    teks_str = str(teks_filtered)
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

def ekstrak_semua_pibc(teks_filtered):
    teks_str = str(teks_filtered)
    hasil = re.findall(POLA_PIBC, teks_str, re.IGNORECASE)
    if hasil: return list(dict.fromkeys([x.strip().rstrip('.') for x in hasil]))
    return []

def dapatkan_pibc_baris(teks_filtered, gudang, teks_raw):
    if not str(teks_filtered).strip() and not str(teks_raw).strip(): return []
    g_norm = normalisasi_gudang(gudang)
    if 'Lab & RnD' in g_norm and 'REPROCESS' in str(teks_raw).upper(): return ekstrak_pibc_lab(teks_filtered)
    if is_to_reprocess(teks_raw): return ekstrak_pibc_to(teks_filtered)
    return ekstrak_semua_pibc(teks_filtered)

def cek_kombinasi_variant(ki_full, kb_full, is_to, teks_gab, gudang_norm):
    ki = str(ki_full).strip().upper()
    kb = str(kb_full).strip().upper()
    tg = str(teks_gab).upper()
    
    if not ki: return 'CLEAN'
    base_err = f"{ki_full} -> {kb_full}"
    
    has_retur_dyeing = bool(re.search(r'\b(RETUR|BEKAS|BKS)\b', tg))
    has_retur_lab = bool(re.search(r'\bRETUR\b', tg))
    has_rnd = bool(re.search(r'\bRND\b', tg))
    has_lab = bool(re.search(r'\bLAB\b', tg))
    
    if ki.startswith('TWP.00') and kb.startswith('TWP.00'):
        return 'CLEAN'
    
    if 'Mesin Dyeing' in gudang_norm:
        if ki.startswith(('TWP.050', 'TWP.051', 'MWP', 'TBM', 'TWP.400')) and kb.startswith('TWP.100'): return 'CLEAN'
        if ki.startswith('TWP.051') and kb.startswith('TWP.051'): return 'CLEAN'
        if ki.startswith(('TWP.100', 'TWP.90', 'TWP.400')) and kb.startswith('TWP.400'): return 'CLEAN'
        if ki.startswith(('TWP.100', 'TWP.90', 'TWP.051', 'TWP.400')) and kb.startswith('TWP.666'): return 'CLEAN'
        if ki.startswith('TWP.050') and kb.startswith('TWP.90'): return 'CLEAN'
        
        if ki.startswith('TWP.100') and kb.startswith('TWP.100'):
            if has_retur_dyeing or is_to: return 'CLEAN'
            return f"KURANG KEYWORD PENJELAS (RETUR / BEKAS / BKS / TO) : {base_err}"
        
        if ki.startswith('TWP.90') and kb.startswith('TWP.90'):
            if has_retur_dyeing or is_to: return 'CLEAN'
            return f"KURANG KEYWORD PENJELAS (RETUR / BEKAS / BKS / TO) : {base_err}"
            
        return base_err
            
    elif 'Lab & RnD' in gudang_norm:
        if ki.startswith(('TWP.050', 'MWP', 'TBM')) and kb.startswith(('TWP.999', 'TWP.90')): return 'CLEAN'
        if ki.startswith(('TWP.999', 'TWP.400')) and kb.startswith('TWP.400'): return 'CLEAN'
        if ki.startswith(('TWP.999', 'TWP.051', 'TWP.100', 'TWP.400')) and kb.startswith('TWP.666'): return 'CLEAN'
        
        if ki.startswith('TWP.050') and kb.startswith('TWP.100'):
            if has_rnd: return 'CLEAN'
            return f"KURANG KEYWORD 'RND' : {base_err}"
            
        if ki.startswith('TWP.100') and kb.startswith('TWP.400'):
            if has_rnd: return 'CLEAN'
            return f"KURANG KEYWORD 'RND' : {base_err}"
            
        if ki.startswith('TWP.100') and kb.startswith('TWP.100'):
            if has_retur_lab or (is_to and has_rnd): return 'CLEAN'
            return f"KURANG KEYWORD 'RETUR' ATAU (TO + 'RND') : {base_err}"
            
        if ki.startswith('TWP.999') and kb.startswith('TWP.999'):
            if has_retur_lab or (is_to and has_lab): return 'CLEAN'
            return f"KURANG KEYWORD 'RETUR' ATAU (TO + 'LAB') : {base_err}"
            
        return base_err
            
    return 'CLEAN'

def siapkan_kolom_standar(df_mb):
    col_gudang_mb = cari_kolom_akurat(df_mb, ['GUDANG'])
    if col_gudang_mb and col_gudang_mb != 'Gudang': df_mb.rename(columns={col_gudang_mb: 'Gudang'}, inplace=True)
    elif not col_gudang_mb: df_mb['Gudang'] = ''

    if 'Kode' in df_mb.columns:
        last_k = ""
        for idx in df_mb.index:
            val_k = str(df_mb.at[idx, 'Kode']).strip()
            if val_k.lower() in ['nan', 'none', 'nat', '']: df_mb.at[idx, 'Kode'] = last_k
            else: last_k = val_k

    last_kode_tracker = ""
    last_gudang, last_ket, last_ket_lain = "", "", ""
    
    for idx in df_mb.index:
        curr_kode = str(df_mb.at[idx, 'Kode']).strip() if 'Kode' in df_mb.columns else ""
        
        if curr_kode != last_kode_tracker:
            last_kode_tracker = curr_kode
            last_gudang = str(df_mb.at[idx, 'Gudang']).strip() if 'Gudang' in df_mb.columns else ""
            last_ket = str(df_mb.at[idx, 'Keterangan']).strip() if 'Keterangan' in df_mb.columns else ""
            last_ket_lain = str(df_mb.at[idx, 'Keterangan Lain']).strip() if 'Keterangan Lain' in df_mb.columns else ""
            
            if last_gudang.lower() in ['nan', 'none', 'nat']: last_gudang = ''
            if last_ket.lower() in ['nan', 'none', 'nat']: last_ket = ''
            if last_ket_lain.lower() in ['nan', 'none', 'nat']: last_ket_lain = ''
            
            if 'Gudang' in df_mb.columns: df_mb.at[idx, 'Gudang'] = last_gudang
            if 'Keterangan' in df_mb.columns: df_mb.at[idx, 'Keterangan'] = last_ket
            if 'Keterangan Lain' in df_mb.columns: df_mb.at[idx, 'Keterangan Lain'] = last_ket_lain
        else:
            curr_gudang = str(df_mb.at[idx, 'Gudang']).strip() if 'Gudang' in df_mb.columns else ""
            curr_ket = str(df_mb.at[idx, 'Keterangan']).strip() if 'Keterangan' in df_mb.columns else ""
            curr_ket_lain = str(df_mb.at[idx, 'Keterangan Lain']).strip() if 'Keterangan Lain' in df_mb.columns else ""
            
            if curr_gudang.lower() in ['nan', 'none', 'nat', '']: 
                if 'Gudang' in df_mb.columns: df_mb.at[idx, 'Gudang'] = last_gudang
            else: last_gudang = curr_gudang
            
            if curr_ket.lower() in ['nan', 'none', 'nat', '']:
                if 'Keterangan' in df_mb.columns: df_mb.at[idx, 'Keterangan'] = last_ket
            else: last_ket = curr_ket
            
            if curr_ket_lain.lower() in ['nan', 'none', 'nat', '']:
                if 'Keterangan Lain' in df_mb.columns: df_mb.at[idx, 'Keterangan Lain'] = last_ket_lain
            else: last_ket_lain = curr_ket_lain

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

    for col_baru in ['Selisih Benang (GR)', 'Qty PIBC', 'Selisih PIBC (GR)', 'Qty PG', 'Selisih PG (GR)', 'Crosscheck Qty', 'Info Celup Gabung', 'Status Cek', 'Status PIBC', 'Rincian Benang', 'Cek Variant', 'Ext_PIBC']:
        if col_baru not in df_mb.columns: df_mb[col_baru] = ''
        df_mb[col_baru] = df_mb[col_baru].astype(object)

    df_mb['Ext_PIBC_List'] = ""
    df_mb['Ext_PIBC_List'] = df_mb['Ext_PIBC_List'].astype(object)

    return df_mb

def jalankan_proses_crosscheck(df_mb, df_pibc, df_pg):
    keyword_selisih = ['AVALAN', 'CROCHET', 'KOR', 'KOR ROMBE', 'KOLONG', 'REEBOK', 'ROLL', 'TALI', 'KUR', 'AVL', 'CONS', 'CONE', 'CONES', 'HTC']

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
        ket1_raw = str(row.get('Keterangan', ''))
        ket2_raw = str(row.get('Keterangan Lain', ''))
        
        ket1 = filter_plus(ket1_raw)
        ket2 = filter_plus(ket2_raw)
        
        pibc_list1 = dapatkan_pibc_baris(ket1, gudang, ket1_raw)
        pibc_list2 = dapatkan_pibc_baris(ket2, gudang, ket2_raw)
        
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

        gudang_group_norm = normalisasi_gudang(gudang_group)
        is_to_group = is_to_reprocess(teks_gab_group)
        
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
            g_norm = gudang_group_norm
            benang_flag = is_item_benang(kode_item, g_norm)

            is_diabaikan = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper() or ('Softcone' in g_norm and not benang_flag))

            if is_diabaikan:
                df_mb.at[idx, 'Qty Target Standar (GR)'] = ''
                df_mb.at[idx, 'Qty BB Standar (GR)'] = ''
                df_mb.at[idx, 'Rincian Benang'] = '' if 'Dyeing STI' in g_norm else rincian_benang_str
                continue

            if 'Dyeing STI' in g_norm:
                if kode_item.upper().startswith('TBB'):
                    df_mb.at[idx, 'Selisih Benang (GR)'] = ''
                    df_mb.at[idx, 'Crosscheck Qty'] = ''
                else:
                    val_target_sti = bersihkan_angka_indo(row.get('Qty Target Standar (GR)'))
                    val_bb_sti = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                    selisih_sti = round(val_bb_sti - val_target_sti, 3)
                    df_mb.at[idx, 'Selisih Benang (GR)'] = selisih_sti
                    
                    nama_brg_jdi_sti = str(row.get('Nama Barang Jadi', '')).upper()
                    tg_sti = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
                    ada_keyword_selisih_sti = any(kw in nama_brg_jdi_sti or kw in tg_sti.upper() for kw in keyword_selisih)
                    
                    if abs(selisih_sti) <= 0.01:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH' if ada_keyword_selisih_sti else 'INCORRECT'
                
                df_mb.at[idx, 'Qty PIBC'] = ''
                df_mb.at[idx, 'Qty PG'] = ''
                df_mb.at[idx, 'Selisih PIBC (GR)'] = ''
                df_mb.at[idx, 'Selisih PG (GR)'] = ''
                df_mb.at[idx, 'Info Celup Gabung'] = ''
                df_mb.at[idx, 'Status Cek'] = ''
                df_mb.at[idx, 'Status PIBC'] = ''
                df_mb.at[idx, 'Rincian Benang'] = ''
                df_mb.at[idx, 'Cek Variant'] = ''
                df_mb.at[idx, 'Ext_PIBC'] = ''
                
                is_first_row = False
                continue

            if benang_flag:
                if 'Mesin Dyeing' in g_norm or 'Lab & RnD' in g_norm:
                     res = cek_kombinasi_variant(kode_item, row.get('Kode Barang Jadi', ''), is_to_group, teks_gab_group, g_norm)
                     if res != 'CLEAN' and res not in error_variants:
                         error_variants.append(res)

            if is_first_row:
                if 'Softcone' not in g_norm:
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

            df_mb.at[idx, 'Rincian Benang'] = rincian_benang_str

            nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
            tg = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
            ada_keyword_selisih = any(kw in nama_brg_jdi or kw in tg.upper() for kw in keyword_selisih)

            if benang_flag:
                if not benang_status_rendered:
                    selisih_total = round(total_bb_benang_grup - target_gr_utama_grup, 3)
                    df_mb.at[idx, 'Selisih Benang (GR)'] = selisih_total
                    
                    if abs(selisih_total) <= 0.01: 
                        df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else: 
                        df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT' if ('Lab & RnD' in g_norm) else ('MEMANG SELISIH' if ada_keyword_selisih else 'INCORRECT')
                    
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

        if is_softcone_group:
            df_mb.at[idx_first, 'Status Cek'] = ''
        else:
            df_mb.at[idx_first, 'Status Cek'] = ", ".join(list(set(status_cek_list))) if status_cek_list else 'COMPLATE'
            if 'Mesin Dyeing' in gudang_group_norm or 'Lab & RnD' in gudang_group_norm:
                df_mb.at[idx_first, 'Cek Variant'] = ", ".join(error_variants) if error_variants else 'CLEAN'
            
    if 'Ext_PIBC_List' in df_mb.columns:
        df_mb = df_mb.drop(columns=['Ext_PIBC_List'])

    return df_mb

def format_excel_sheet(ws, df, nama_kolom_kode='Kode'):
    # 1. Warna Header Merah D51C39 dengan teks Bold Putih
    header_fill = PatternFill(start_color='FFD51C39', end_color='FFD51C39', fill_type='solid')
    header_font = Font(color='FFFFFFFF', bold=True)
    
    # Kolom yang akan otomatis di-Wrap Text
    wrap_cols = ['Keterangan', 'Keterangan Lain', 'Info Celup Gabung', 'Status Cek', 'Rincian Benang', 'Cek Variant', 'Ext_PIBC', 'Kode PIBC Terdeteksi']
    
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        
        # Lebarkan kolom spesifik agar wrap text rapi
        if col_name in wrap_cols:
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 35

    # 2. Warna Baris Selang-Seling (Banded Rows) per grup Kode MB
    warna_hex = ['FFBDB2FF', 'FFFFCEE3', 'FFFDFFD2', 'FFCAEDFF']
    color_fills = [PatternFill(start_color=c, end_color=c, fill_type='solid') for c in warna_hex]
    
    wrap_alignment = Alignment(wrap_text=True, vertical='top')
    std_alignment = Alignment(vertical='top')
    
    try:
        idx_kode = list(df.columns).index(nama_kolom_kode) + 1
    except ValueError:
        return 
        
    current_kode = None
    color_index = -1
    
    for r_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column), start=2):
        cell_kode_val = ws.cell(row=r_idx, column=idx_kode).value
        if cell_kode_val != current_kode:
            current_kode = cell_kode_val
            color_index = (color_index + 1) % len(color_fills)
        
        fill_to_apply = color_fills[color_index]
        for c_idx, cell in enumerate(row, start=1):
            cell.fill = fill_to_apply
            
            # Wrap text untuk kolom yang panjang
            col_name = df.columns[c_idx - 1]
            if col_name in wrap_cols:
                cell.alignment = wrap_alignment
            else:
                cell.alignment = std_alignment

def buat_excel_hasil_keseluruhan(df_hasil):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_hasil.to_excel(writer, index=False, sheet_name='Hasil Crosscheck')
        format_excel_sheet(writer.sheets['Hasil Crosscheck'], df_hasil, 'Kode')
    return output.getvalue()

def buat_excel_rekap_audit(df_hasil):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
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
                    
                # SUSUNAN 16 KOLOM FINAL
                data_ringkas.append({
                    'Gudang': row.get('Gudang', ''),
                    'Kode': row.get('Kode', ''),
                    'Keterangan': row.get('Keterangan', ''),
                    'Keterangan Lain': row.get('Keterangan Lain', ''),
                    'Kode PIBC Terdeteksi': row.get('Ext_PIBC', ''),
                    'Item Bermasalah': item_bermasalah,
                    'Qty Target Standar (GR)': row.get('Qty Target Standar (GR)', ''),
                    'Qty BB Standar (GR)': row.get('Qty BB Standar (GR)', ''),
                    'Qty PIBC': row.get('Qty PIBC', ''),
                    'Qty PG': row.get('Qty PG', ''),
                    'Selisih Benang (GR)': row.get('Selisih Benang (GR)', ''),
                    'Selisih PIBC (GR)': row.get('Selisih PIBC (GR)', ''),
                    'Selisih PG (GR)': row.get('Selisih PG (GR)', ''),
                    'Jenis Temuan': jenis,
                    'Detail Kesalahan & Seharusnya': detail,
                    'Catatan Auditor': ''
                })
        
        df_ringkas = pd.DataFrame(data_ringkas)
        if not df_ringkas.empty:
            df_ringkas.to_excel(writer, index=False, sheet_name='Laporan Ringkas Admin')
            format_excel_sheet(writer.sheets['Laporan Ringkas Admin'], df_ringkas, 'Kode')
        else:
            pd.DataFrame({'Info': ['Tidak ada temuan error']}).to_excel(writer, index=False, sheet_name='Laporan Ringkas Admin')
            
        mask_qty = df_hasil['Crosscheck Qty'].astype(str).str.upper().apply(lambda x: ('INCORRECT' in x or 'SELISIH' in x) and x != 'MEMANG SELISIH')
        kode_qty = df_hasil[mask_qty]['Kode'].unique() if 'Kode' in df_hasil.columns else []
        if len(kode_qty) > 0:
            df_qty_full = df_hasil[df_hasil['Kode'].isin(kode_qty)]
            df_qty_full.to_excel(writer, index=False, sheet_name='Error Qty (Full Grup)')
            format_excel_sheet(writer.sheets['Error Qty (Full Grup)'], df_qty_full, 'Kode')
        else:
            pd.DataFrame({'Info': ['Tidak ada temuan']}).to_excel(writer, index=False, sheet_name='Error Qty (Full Grup)')
            
        mask_var = df_hasil['Cek Variant'].astype(str).str.upper().apply(lambda x: x != '' and x != 'CLEAN')
        kode_var = df_hasil[mask_var]['Kode'].unique() if 'Kode' in df_hasil.columns else []
        if len(kode_var) > 0:
            df_var_full = df_hasil[df_hasil['Kode'].isin(kode_var)]
            df_var_full.to_excel(writer, index=False, sheet_name='Error Varian (Full Grup)')
            format_excel_sheet(writer.sheets['Error Varian (Full Grup)'], df_var_full, 'Kode')
        else:
            pd.DataFrame({'Info': ['Tidak ada temuan']}).to_excel(writer, index=False, sheet_name='Error Varian (Full Grup)')
            
        mask_ms = df_hasil['Crosscheck Qty'].astype(str).str.upper() == 'MEMANG SELISIH'
        kode_ms = df_hasil[mask_ms]['Kode'].unique() if 'Kode' in df_hasil.columns else []
        if len(kode_ms) > 0:
            df_ms_full = df_hasil[df_hasil['Kode'].isin(kode_ms)]
            df_ms_full.to_excel(writer, index=False, sheet_name='Memang Selisih (Full Grup)')
            format_excel_sheet(writer.sheets['Memang Selisih (Full Grup)'], df_ms_full, 'Kode')
        else:
            pd.DataFrame({'Info': ['Tidak ada temuan']}).to_excel(writer, index=False, sheet_name='Memang Selisih (Full Grup)')
            
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
