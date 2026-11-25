import pandas as pd
import re
import io
import zipfile
from datetime import datetime, timezone, timedelta
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

BARIS_DIABAIKAN = [
    "Total Material (Details)", "Total Pekerjaan (Works)", 
    "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor",
    "Kode Pekerjaan", "Nama Pekerjaan"
]

THIN_BLACK = Side(border_style='thin', color='FF000000')
MED_BLACK = Side(border_style='medium', color='FF000000')
CELL_BORDER = Border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)
HEADER_BORDER = Border(left=THIN_BLACK, right=THIN_BLACK, top=MED_BLACK, bottom=MED_BLACK)

def get_hari_ini_wib():
    wib = timezone(timedelta(hours=7))
    return datetime.now(wib).date()

def sanitize_key(k):
    if not k: return ""
    k = re.sub(r'[^A-Z0-9\.#\+]', '', str(k).upper())
    return k.replace('PIBC', '')

def sanitize_mesin(m):
    if pd.isna(m) or not m: return ""
    s = re.sub(r'\.0$', '', str(m).strip().upper())
    return re.sub(r'[^A-Z0-9]', '', s)

def is_target_keyword(nama_item):
    nm = str(nama_item).upper()
    return ('BIOWHITE' in nm) or ('SYNERWHITE' in nm)

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
            if kem_clean == col_clean: return col
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
    
    hasil = re.findall(POLA_PIBC, teks_str, re.IGNORECASE)
    return list(dict.fromkeys([x.strip().rstrip('.') for x in hasil]))

def ekstrak_pibc_to(teks_filtered):
    teks_str = str(teks_filtered)
    hasil = []
    
    chunks_kurung = re.findall(r'\([^)]*(?:\)|$)', teks_str)
    for chunk in chunks_kurung:
        hasil.extend(re.findall(POLA_PIBC, chunk, re.IGNORECASE))
            
    if not hasil and re.search(r'\bTO\b', teks_str, re.IGNORECASE):
        parts = re.split(r'\bTO\b', teks_str, flags=re.IGNORECASE)
        if len(parts) > 1:
            hasil.extend(re.findall(POLA_PIBC, parts[-1], re.IGNORECASE))

    if not hasil and re.search(r'\bREPROCESS', teks_str, re.IGNORECASE):
        parts = re.split(r'\bREPROCESS', teks_str, flags=re.IGNORECASE)
        if len(parts) > 1:
            hasil.extend(re.findall(POLA_PIBC, parts[-1], re.IGNORECASE))
            
    if not hasil:
        hasil.extend(re.findall(POLA_PIBC, teks_str, re.IGNORECASE))
        
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
            if has_retur_lab or (is_to and has_lab) or (is_to and has_rnd): return 'CLEAN'
            return f"KURANG KEYWORD 'RETUR' ATAU (TO + 'LAB'/'RND') : {base_err}"
            
        return base_err
            
    return 'CLEAN'

def _buat_label_item(kode_raw, nama_raw):
    k = str(kode_raw).strip() if pd.notna(kode_raw) else ""
    n = str(nama_raw).strip() if pd.notna(nama_raw) else ""
    if k.lower() in ['nan', 'none', 'nat']: k = ""
    if n.lower() in ['nan', 'none', 'nat']: n = ""
    if k and n: return f"{k} - {n}"
    return k or n

def _item_ada_di_kumpulan(item, kumpulan_dict):
    k_kode = item.get('k_kode', '')
    k_nama = item.get('k_nama', '')
    for other in kumpulan_dict.values():
        if k_kode and other.get('k_kode') and k_kode == other['k_kode']:
            return True
        if k_nama and other.get('k_nama') and k_nama == other['k_nama']:
            return True
    return False

def bandingkan_set_item(items_kiri, items_kanan, label_kiri, cek_lebih=True):
    kurang = []
    for it in items_kanan.values():
        if not _item_ada_di_kumpulan(it, items_kiri):
            kurang.append(it['label'])
            
    lebih = []
    if cek_lebih:
        for it in items_kiri.values():
            if not _item_ada_di_kumpulan(it, items_kanan):
                lebih.append(it['label'])
                
    if not kurang and not lebih:
        return 'COMPLETE'
        
    bagian = []
    if kurang:
        bagian.append(f"Kurang di {label_kiri}:\n" + "\n".join([f"- {x}" for x in dict.fromkeys(kurang)]))
    if lebih:
        bagian.append(f"Lebih di {label_kiri}:\n" + "\n".join([f"- {x}" for x in dict.fromkeys(lebih)]))
    return "\n".join(bagian)

def is_cek_item_bermasalah(val):
    s = str(val).upper()
    return ('KURANG DI ' in s) or ('LEBIH DI ' in s)

def parse_tanggal_aman(val):
    if pd.isna(val) or val is None:
        return pd.NaT
    if isinstance(val, (pd.Timestamp, datetime)):
        return val
    if hasattr(val, 'year') and hasattr(val, 'month') and hasattr(val, 'day'):
        return pd.Timestamp(val)
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'nat']:
        return pd.NaT
    if re.match(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}', s):
        return pd.to_datetime(s, dayfirst=False, errors='coerce')
    if re.match(r'^\d{1,2}[-/]\d{1,2}[-/]\d{4}', s):
        return pd.to_datetime(s, dayfirst=True, errors='coerce')
    return pd.to_datetime(s, dayfirst=True, errors='coerce')

def hitung_masa_dan_rentang(df, col_tgl, nama_kolom_masa):
    df_out = df.copy()
    hari_ini = get_hari_ini_wib()
    
    if col_tgl and col_tgl in df_out.columns:
        if pd.api.types.is_datetime64_any_dtype(df_out[col_tgl]):
            tgl_parsed = df_out[col_tgl]
        else:
            tgl_parsed = df_out[col_tgl].apply(parse_tanggal_aman)
            
        masa_list = []
        rentang_list = []
        for dt in tgl_parsed:
            if pd.isna(dt):
                masa_list.append(pd.NA)
                rentang_list.append("")
            else:
                selisih = int((hari_ini - dt.date()).days)
                masa_list.append(selisih)
                if selisih <= 3:
                    rentang_list.append("0 - 3 Hari")
                elif selisih <= 7:
                    rentang_list.append("4 - 7 Hari")
                else:
                    rentang_list.append("> 7 Hari")
                    
        df_out[nama_kolom_masa] = pd.Series(masa_list, index=df_out.index, dtype="Int64")
        df_out['Rentang Hari'] = rentang_list
        df_out = df_out.sort_values(by=nama_kolom_masa, ascending=False, na_position='last').reset_index(drop=True)
    else:
        df_out[nama_kolom_masa] = pd.Series([pd.NA] * len(df_out), dtype="Int64")
        df_out['Rentang Hari'] = ""
        
    return df_out

def siapkan_kolom_standar(df_mb):
    col_gudang_mb = cari_kolom_akurat(df_mb, ['GUDANG'])
    if col_gudang_mb and col_gudang_mb != 'Gudang': df_mb.rename(columns={col_gudang_mb: 'Gudang'}, inplace=True)
    elif not col_gudang_mb: df_mb['Gudang'] = ''

    col_tgl_mb = cari_kolom_akurat(df_mb, ['TANGGAL', 'TGL', 'DATE'])
    if col_tgl_mb and col_tgl_mb != 'Tanggal':
        df_mb.rename(columns={col_tgl_mb: 'Tanggal'}, inplace=True)

    if 'Kode' in df_mb.columns:
        last_k = ""
        for idx in df_mb.index:
            val_k = str(df_mb.at[idx, 'Kode']).strip()
            if val_k.lower() in ['nan', 'none', 'nat', '']: df_mb.at[idx, 'Kode'] = last_k
            else: last_k = val_k

    if 'Tanggal' in df_mb.columns:
        df_mb['Tanggal'] = df_mb['Tanggal'].ffill()

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

    seri_target_gr = pd.Series([None] * len(df_mb), dtype=object)
    seri_bb_gr = pd.Series([None] * len(df_mb), dtype=object)

    for idx, r in df_mb.iterrows():
        kode_item = str(r[col_kode_item]).strip() if col_kode_item and pd.notna(r[col_kode_item]) else ''
        gudang = str(r.get('Gudang', ''))
        g_norm = normalisasi_gudang(gudang)
        is_benang = is_item_benang(kode_item, gudang)
        
        is_diabaikan = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper() or ('Softcone' in g_norm and not is_benang))

        if not is_diabaikan:
            if col_target_qty and pd.notna(r[col_target_qty]):
                raw_tgt = str(r[col_target_qty]).strip().lower()
                if raw_tgt not in ['', 'nan', 'none', 'nat']:
                    val = bersihkan_angka_indo(r[col_target_qty])
                    unit = str(r[col_target_unit]).strip().upper() if col_target_unit else 'GR'
                    seri_target_gr[idx] = round(val if unit == 'GR' else val * 1000.0, 3)
                    
            if col_qty and pd.notna(r[col_qty]):
                raw_qty = str(r[col_qty]).strip().lower()
                if raw_qty not in ['', 'nan', 'none', 'nat']:
                    val = bersihkan_angka_indo(r[col_qty])
                    unit = str(r[col_unit]).strip().upper() if col_unit else 'GR'
                    seri_bb_gr[idx] = round(val if unit == 'GR' else val * 1000.0, 3)

    idx_target = list(df_mb.columns).index(col_target_unit if col_target_unit else col_target_qty) + 1 if (col_target_unit or col_target_qty) else len(df_mb.columns)
    df_mb.insert(idx_target, 'Qty Target Standar (GR)', seri_target_gr)
    
    idx_bb = list(df_mb.columns).index(col_qty) + 1 if col_qty else len(df_mb.columns)
    df_mb.insert(idx_bb, 'Qty BB Standar (GR)', seri_bb_gr)

    for col_baru in ['Qty Memphis', 'Qty PG', 'Selisih Benang (GR)', 'Selisih Memphis (GR)', 'Selisih PG (GR)', 'Selisih PG vs Memphis (GR)']:
        if col_baru not in df_mb.columns: df_mb[col_baru] = None
        df_mb[col_baru] = df_mb[col_baru].astype(object)
        
    for col_baru in [
        'Crosscheck Qty', 'Info Mesin', 'Info Celup Gabung',
        'Cek Item (MB vs Memphis)', 'Cek Item (MB vs PG)', 'Cek Item (PG vs Memphis)',
        'Status PIBC', 'Kode Transaksi PG', 'Rincian Benang', 'Cek Variant', 'Ext_PIBC'
    ]:
        if col_baru not in df_mb.columns: df_mb[col_baru] = ''
        df_mb[col_baru] = df_mb[col_baru].astype(object)

    df_mb['Ext_PIBC_List'] = ""
    df_mb['Ext_PIBC_List'] = df_mb['Ext_PIBC_List'].astype(object)

    return df_mb

def jalankan_proses_crosscheck(df_mb, df_pibc, df_pg, df_mesin=None):
    keyword_selisih = ['AVALAN', 'CROCHET', 'KOR', 'KOR ROMBE', 'KOLONG', 'REEBOK', 'ROLL', 'TALI', 'KUR', 'AVL', 'CONS', 'CONE', 'CONES', 'HTC']

    # 1. Lookup Master Data Mesin
    mesin_lookup = {}
    has_master_mesin = False
    if df_mesin is not None and not df_mesin.empty:
        col_m = cari_kolom_akurat(df_mesin, ['MESIN', 'NAMAMESIN', 'KODEMESIN'])
        col_j = cari_kolom_akurat(df_mesin, ['JENISMESIN', 'JENIS', 'KATEGORI'])
        if col_m and col_j:
            has_master_mesin = True
            for _, r in df_mesin.iterrows():
                m_key = sanitize_mesin(r[col_m])
                j_val = str(r[col_j]).strip().upper()
                if m_key:
                    mesin_lookup[m_key] = j_val

    # 2. Indexing Master Memphis (Kunci Utama: No. Ticket)
    pibc_available_keys = set()
    pibc_lookup_act = {}
    pibc_lookup_tgt = {}
    pibc_items_by_key = {}
    pibc_mesin_map = {}
    
    if df_pibc is not None and not df_pibc.empty:
        col_ticket = cari_kolom_akurat(df_pibc, ['NOTICKET', 'DYELOT', 'PIBC'])
        col_act = cari_kolom_akurat(df_pibc, ['ACTUAL', 'QTY', 'JUMLAH'])
        col_tgt = cari_kolom_akurat(df_pibc, ['TARGET'])
        col_kode_obat = cari_kolom_akurat(df_pibc, ['KODEOBAT', 'KODEITEM'])
        col_nama_obat = cari_kolom_akurat(df_pibc, ['NAMAOBAT', 'NAMAITEM', 'NAMA', 'ITEM', 'OBAT'])
        col_mesin_mem = cari_kolom_akurat(df_pibc, ['MESIN'])
        
        if col_ticket and col_act:
            for _, r in df_pibc.iterrows():
                k_ticket = sanitize_key(r[col_ticket])
                if not k_ticket: continue
                pibc_available_keys.add(k_ticket)
                
                if col_mesin_mem and pd.notna(r[col_mesin_mem]):
                    m_raw = re.sub(r'\.0$', '', str(r[col_mesin_mem]).strip())
                    if m_raw and m_raw.lower() not in ['nan', 'none', 'nat']:
                        if k_ticket not in pibc_mesin_map: pibc_mesin_map[k_ticket] = []
                        if m_raw not in pibc_mesin_map[k_ticket]: pibc_mesin_map[k_ticket].append(m_raw)
                        
                val_act = bersihkan_angka_indo(r[col_act])
                val_tgt = bersihkan_angka_indo(r[col_tgt]) if col_tgt and pd.notna(r[col_tgt]) else val_act
                
                kode_raw = str(r[col_kode_obat]).strip() if col_kode_obat and pd.notna(r[col_kode_obat]) else ""
                nama_raw = str(r[col_nama_obat]).strip() if col_nama_obat and pd.notna(r[col_nama_obat]) else ""
                if kode_raw.lower() in ['nan', 'none', 'nat']: kode_raw = ""
                if nama_raw.lower() in ['nan', 'none', 'nat']: nama_raw = ""
                
                k_kode = sanitize_key(kode_raw)
                k_nama = sanitize_key(nama_raw)
                is_bw_sw = is_target_keyword(nama_raw)
                
                if val_act > 0:
                    if k_kode: pibc_lookup_act[(k_ticket, k_kode)] = round(pibc_lookup_act.get((k_ticket, k_kode), 0.0) + val_act, 3)
                    if k_nama: pibc_lookup_act[(k_ticket, k_nama)] = round(pibc_lookup_act.get((k_ticket, k_nama), 0.0) + val_act, 3)
                    
                if val_tgt > 0:
                    if k_kode: pibc_lookup_tgt[(k_ticket, k_kode)] = round(pibc_lookup_tgt.get((k_ticket, k_kode), 0.0) + val_tgt, 3)
                    if k_nama: pibc_lookup_tgt[(k_ticket, k_nama)] = round(pibc_lookup_tgt.get((k_ticket, k_nama), 0.0) + val_tgt, 3)

                if (val_act > 0) or (is_bw_sw and 0 < val_tgt < 100.0):
                    canon_k = k_kode or k_nama
                    if canon_k:
                        if k_ticket not in pibc_items_by_key: pibc_items_by_key[k_ticket] = {}
                        if canon_k not in pibc_items_by_key[k_ticket]:
                            pibc_items_by_key[k_ticket][canon_k] = {
                                'kode': kode_raw, 'nama': nama_raw, 'k_kode': k_kode, 'k_nama': k_nama,
                                'label': _buat_label_item(kode_raw, nama_raw)
                            }

    # 3. Indexing Master PG (Terkunci per Gudang)
    pg_available_keys_by_gudang = set()
    pg_lookup = {}
    pg_trans_lookup = {}
    pg_items_by_gudang_key = {}
    
    if df_pg is not None and not df_pg.empty:
        if 'PIBC_PG' in df_pg.columns and 'Qty_PG' in df_pg.columns:
            col_no_trans_pg = cari_kolom_akurat(df_pg, ['NOTRANSAKSI', 'KODETRANSAKSI'])
            for _, r in df_pg.iterrows():
                val_msk = bersihkan_angka_indo(r['Qty_PG'])
                if val_msk <= 0: continue 
                raw_pibc = str(r['PIBC_PG'])
                if not raw_pibc or raw_pibc.lower() == 'nan': continue
                
                g_norm_pg = normalisasi_gudang(r.get('Gudang_PG', ''))
                pibc_ext_list = ekstrak_semua_pibc(raw_pibc)
                for p_raw in pibc_ext_list:
                    k_ket = sanitize_key(p_raw)
                    if k_ket: pg_available_keys_by_gudang.add((g_norm_pg, k_ket))
                
                if not pibc_ext_list: continue
                k_ket_utama = sanitize_key(pibc_ext_list[0])
                
                kode_raw = str(r.get('Kode_Item_PG', '')).strip()
                nama_raw = str(r.get('Nama_Item_PG', '')).strip()
                if kode_raw.lower() in ['nan', 'none', 'nat']: kode_raw = ""
                if nama_raw.lower() in ['nan', 'none', 'nat']: nama_raw = ""
                
                k_kode = sanitize_key(kode_raw)
                k_nama = sanitize_key(nama_raw)
                no_trans = str(r[col_no_trans_pg]).strip() if col_no_trans_pg and pd.notna(r[col_no_trans_pg]) else ""
                if no_trans.lower() in ['nan', 'none', 'nat']: no_trans = ""
                
                if k_kode:
                    key_pg = (g_norm_pg, k_ket_utama, k_kode)
                    pg_lookup[key_pg] = round(pg_lookup.get(key_pg, 0.0) + val_msk, 3)
                    if no_trans:
                        if key_pg not in pg_trans_lookup: pg_trans_lookup[key_pg] = {}
                        pg_trans_lookup[key_pg][no_trans] = round(pg_trans_lookup[key_pg].get(no_trans, 0.0) + val_msk, 3)
                if k_nama:
                    key_pg_n = (g_norm_pg, k_ket_utama, k_nama)
                    pg_lookup[key_pg_n] = round(pg_lookup.get(key_pg_n, 0.0) + val_msk, 3)
                    if no_trans:
                        if key_pg_n not in pg_trans_lookup: pg_trans_lookup[key_pg_n] = {}
                        pg_trans_lookup[key_pg_n][no_trans] = round(pg_trans_lookup[key_pg_n].get(no_trans, 0.0) + val_msk, 3)
                        
                canon_k = k_kode or k_nama
                if canon_k:
                    gk = (g_norm_pg, k_ket_utama)
                    if gk not in pg_items_by_gudang_key: pg_items_by_gudang_key[gk] = {}
                    if canon_k not in pg_items_by_gudang_key[gk]:
                        pg_items_by_gudang_key[gk][canon_k] = {
                            'kode': kode_raw, 'nama': nama_raw, 'k_kode': k_kode, 'k_nama': k_nama,
                            'label': _buat_label_item(kode_raw, nama_raw)
                        }

    # 4. Ekstraksi PIBC per Baris MB
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

    # 5. Pre-Scan Celup Gabung & PIBC Ganda Tanpa Keyword Gabung
    pibc_to_mb = {}
    all_pibc_to_mb = {}
    gabungan_qty_item = {}
    gabungan_items_mb = {}

    for idx, row in df_mb.iterrows():
        gudang = str(row.get('Gudang', ''))
        g_norm = normalisasi_gudang(gudang)
        ket_gabung_cek = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
        is_teks_gabung = 'celup gabung' in ket_gabung_cek.lower() or 'gabung' in ket_gabung_cek.lower()
        kp_list = row.get('Ext_PIBC_List')
        k_val = row.get('Kode', idx)
        
        kp_utama = sanitize_key(kp_list[0]) if isinstance(kp_list, list) and len(kp_list) > 0 else ""
        
        if kp_utama != '' and ('Mesin Dyeing' in g_norm or 'Lab & RnD' in g_norm):
            gk_mb = (g_norm, kp_utama)
            if gk_mb not in all_pibc_to_mb: all_pibc_to_mb[gk_mb] = []
            if k_val not in all_pibc_to_mb[gk_mb]: all_pibc_to_mb[gk_mb].append(k_val)

        if is_teks_gabung and kp_utama != '':
            if kp_utama not in pibc_to_mb: pibc_to_mb[kp_utama] = []
            if k_val not in pibc_to_mb[kp_utama]: pibc_to_mb[kp_utama].append(k_val)

            kode_item = str(row.get('Kode Item', '')).strip()
            nama_item = str(row.get('Nama Item', '')).strip()
            k_item_up = kode_item.upper()
            if not is_item_benang(k_item_up, gudang) and kode_item not in BARIS_DIABAIKAN and 'PEKERJAAN' not in k_item_up:
                val_bb = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                key_item = (kp_utama, k_item_up)
                gabungan_qty_item[key_item] = round(gabungan_qty_item.get(key_item, 0.0) + val_bb, 3)
                
                k_k = sanitize_key(kode_item)
                k_n = sanitize_key(nama_item)
                canon_k = k_k or k_n
                if canon_k:
                    if kp_utama not in gabungan_items_mb: gabungan_items_mb[kp_utama] = {}
                    if canon_k not in gabungan_items_mb[kp_utama]:
                        gabungan_items_mb[kp_utama][canon_k] = {
                            'kode': kode_item, 'nama': nama_item, 'k_kode': k_k, 'k_nama': k_n,
                            'label': _buat_label_item(kode_item, nama_item)
                        }

    pibc_celup_gabung = {pibc for pibc, mbs in pibc_to_mb.items() if len(set(mbs)) > 1}

    # 6. Loop Utama per Transaksi MB
    for k_val, df_sub in grup_kode:
        grup_punya_pibc = False
        pibc_list_grup = []
        target_gr_utama_grup = 0.0
        total_bb_benang_grup = 0.0
        rincian_benang_list = []
        
        item_qty_dalam_transaksi = {} 
        items_mb_grup = {}
        
        teks_gab_group = ""
        gudang_group = ""
        
        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            nama_item = str(row.get('Nama Item', '')).strip()
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
                    
                    k_k = sanitize_key(kode_item)
                    k_n = sanitize_key(nama_item)
                    canon_k = k_k or k_n
                    if canon_k and canon_k not in items_mb_grup:
                        items_mb_grup[canon_k] = {
                            'kode': kode_item, 'nama': nama_item, 'k_kode': k_k, 'k_nama': k_n,
                            'label': _buat_label_item(kode_item, nama_item)
                        }

        gudang_group_norm = normalisasi_gudang(gudang_group)
        is_to_group = is_to_reprocess(teks_gab_group)
        is_gudang_audit = ('Mesin Dyeing' in gudang_group_norm) or ('Lab & RnD' in gudang_group_norm)
        
        rincian_benang_str = ", ".join(rincian_benang_list)
        is_celup_gabung = False
        pasangan_str = ""
        kp_grup_utama = sanitize_key(pibc_list_grup[0]) if len(pibc_list_grup) > 0 else ""
        
        if kp_grup_utama and kp_grup_utama in pibc_celup_gabung:
            is_celup_gabung = True
            semua_mb_terkait = pibc_to_mb.get(kp_grup_utama, [])
            pasangan_str = ", ".join([str(m) for m in set(semua_mb_terkait) if m != k_val])

        peringatan_ganda = ""
        if not is_celup_gabung and kp_grup_utama and is_gudang_audit:
            semua_mb_sama = all_pibc_to_mb.get((gudang_group_norm, kp_grup_utama), [])
            if len(set(semua_mb_sama)) > 1:
                pasangan_ganda = ", ".join([str(m) for m in set(semua_mb_sama) if m != k_val])
                peringatan_ganda = f"Peringatan: PIBC Dipakai >1 MB (Tanpa Keyword Gabung) dgn: {pasangan_ganda}"

        info_cg_base = f"Celup Gabung dgn: {pasangan_str}" if is_celup_gabung else peringatan_ganda

        # Identifikasi Mesin & Kategori (BESAR / KECIL / TIDAK TERDAFTAR)
        info_mesin_str = ""
        is_mesin_kecil = False
        if is_gudang_audit and len(pibc_list_grup) > 0:
            mesin_raw_list = []
            for p_raw in pibc_list_grup:
                for m_r in pibc_mesin_map.get(sanitize_key(p_raw), []):
                    if m_r not in mesin_raw_list:
                        mesin_raw_list.append(m_r)
                        
            if mesin_raw_list:
                if has_master_mesin:
                    label_m_list = []
                    jenis_list = []
                    for m_r in mesin_raw_list:
                        j_kat = mesin_lookup.get(sanitize_mesin(m_r))
                        if j_kat:
                            label_m_list.append(f"{m_r} ({j_kat})")
                            jenis_list.append(j_kat)
                        else:
                            label_m_list.append(f"{m_r} (TIDAK TERDAFTAR)")
                            jenis_list.append("TIDAK TERDAFTAR")
                    info_mesin_str = ", ".join(label_m_list)
                    is_mesin_kecil = (len(jenis_list) > 0) and all('KECIL' in j for j in jenis_list)
                else:
                    info_mesin_str = ", ".join(mesin_raw_list)

        is_first_row = True
        benang_status_rendered = False
        obat_rendered_set = set()
        error_variants = []

        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            nama_item = str(row.get('Nama Item', '')).strip()
            g_norm = gudang_group_norm
            benang_flag = is_item_benang(kode_item, g_norm)

            is_diabaikan = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper() or ('Softcone' in g_norm and not benang_flag))

            if is_diabaikan:
                df_mb.loc[idx, 'Qty Target Standar (GR)'] = None
                df_mb.loc[idx, 'Qty BB Standar (GR)'] = None
                df_mb.loc[idx, 'Rincian Benang'] = '' if 'Dyeing STI' in g_norm else rincian_benang_str
                continue

            if 'Dyeing STI' in g_norm:
                if kode_item.upper().startswith('TBB'):
                    df_mb.loc[idx, 'Selisih Benang (GR)'] = None
                    df_mb.loc[idx, 'Crosscheck Qty'] = ''
                else:
                    val_target_sti = bersihkan_angka_indo(row.get('Qty Target Standar (GR)'))
                    val_bb_sti = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                    selisih_sti = round(val_bb_sti - val_target_sti, 3)
                    
                    nama_brg_jdi_sti = str(row.get('Nama Barang Jadi', '')).upper()
                    tg_sti = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
                    ada_keyword_selisih_sti = any(kw in nama_brg_jdi_sti or kw in tg_sti.upper() for kw in keyword_selisih)
                    
                    if abs(selisih_sti) <= 0.009:
                        selisih_sti = 0.0
                        
                    df_mb.loc[idx, 'Selisih Benang (GR)'] = float(selisih_sti)
                    
                    if selisih_sti == 0.0:
                        df_mb.loc[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else:
                        df_mb.loc[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH' if ada_keyword_selisih_sti else 'INCORRECT'
                
                df_mb.loc[idx, 'Qty Memphis'] = None
                df_mb.loc[idx, 'Qty PG'] = None
                df_mb.loc[idx, 'Selisih Memphis (GR)'] = None
                df_mb.loc[idx, 'Selisih PG (GR)'] = None
                df_mb.loc[idx, 'Selisih PG vs Memphis (GR)'] = None
                df_mb.loc[idx, 'Info Mesin'] = ''
                df_mb.loc[idx, 'Info Celup Gabung'] = ''
                df_mb.loc[idx, 'Cek Item (MB vs Memphis)'] = ''
                df_mb.loc[idx, 'Cek Item (MB vs PG)'] = ''
                df_mb.loc[idx, 'Cek Item (PG vs Memphis)'] = ''
                df_mb.loc[idx, 'Status PIBC'] = ''
                df_mb.loc[idx, 'Kode Transaksi PG'] = ''
                df_mb.loc[idx, 'Rincian Benang'] = ''
                df_mb.loc[idx, 'Cek Variant'] = ''
                df_mb.loc[idx, 'Ext_PIBC'] = ''
                
                is_first_row = False
                continue

            if is_gudang_audit and info_mesin_str:
                df_mb.loc[idx, 'Info Mesin'] = info_mesin_str

            if benang_flag:
                if is_gudang_audit:
                     res = cek_kombinasi_variant(kode_item, row.get('Kode Barang Jadi', ''), is_to_group, teks_gab_group, g_norm)
                     if res != 'CLEAN' and res not in error_variants:
                         error_variants.append(res)

            if is_first_row:
                if 'Softcone' not in g_norm:
                    if not grup_punya_pibc:
                        df_mb.loc[idx, 'Status PIBC'] = 'Kode PIBC di MB Tidak Ada'
                    else:
                        miss_pibc = [p for p in pibc_list_grup if sanitize_key(p) not in pibc_available_keys]
                        miss_pg = [p for p in pibc_list_grup if (g_norm, sanitize_key(p)) not in pg_available_keys_by_gudang]
                        
                        if not miss_pibc and not miss_pg:
                            df_mb.loc[idx, 'Status PIBC'] = 'Kode PIBC di MB Ada'
                        elif miss_pibc and miss_pg:
                            if set(miss_pibc) == set(pibc_list_grup) and set(miss_pg) == set(pibc_list_grup):
                                df_mb.loc[idx, 'Status PIBC'] = 'Kode PIBC di MB Ada, di Master Memphis & PG Tidak Ada'
                            else:
                                gabung_miss = list(dict.fromkeys(miss_pibc + miss_pg))
                                df_mb.loc[idx, 'Status PIBC'] = f'Kode PIBC di MB Ada (Namun [{", ".join(gabung_miss)}] di Master Memphis & PG Tidak Ada)'
                        elif miss_pibc:
                            if set(miss_pibc) == set(pibc_list_grup):
                                df_mb.loc[idx, 'Status PIBC'] = 'Kode PIBC di MB Ada, di Master Memphis Tidak Ada'
                            else:
                                df_mb.loc[idx, 'Status PIBC'] = f'Kode PIBC di MB Ada (Namun [{", ".join(miss_pibc)}] di Master Memphis Tidak Ada)'
                        elif miss_pg:
                            if set(miss_pg) == set(pibc_list_grup):
                                df_mb.loc[idx, 'Status PIBC'] = 'Kode PIBC di MB Ada, di Master PG Tidak Ada'
                            else:
                                df_mb.loc[idx, 'Status PIBC'] = f'Kode PIBC di MB Ada (Namun [{", ".join(miss_pg)}] di Master PG Tidak Ada)'

            df_mb.loc[idx, 'Rincian Benang'] = rincian_benang_str

            nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
            tg = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
            ada_keyword_selisih = any(kw in nama_brg_jdi or kw in tg.upper() for kw in keyword_selisih)

            if benang_flag:
                if not benang_status_rendered:
                    selisih_total = round(total_bb_benang_grup - target_gr_utama_grup, 3)
                    
                    if abs(selisih_total) <= 0.009:
                        selisih_total = 0.0
                        
                    df_mb.loc[idx, 'Selisih Benang (GR)'] = float(selisih_total)
                    
                    if selisih_total == 0.0: 
                        df_mb.loc[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else: 
                        df_mb.loc[idx, 'Crosscheck Qty'] = 'INCORRECT' if ('Lab & RnD' in g_norm) else ('MEMANG SELISIH' if ada_keyword_selisih else 'INCORRECT')
                    
                    if peringatan_ganda:
                        df_mb.loc[idx, 'Info Celup Gabung'] = peringatan_ganda
                    
                    benang_status_rendered = True
            else:
                k_item_upper = kode_item.upper()
                if k_item_upper not in obat_rendered_set:
                    qty_bb_baris = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                    qty_bb_mb_sendiri = round(item_qty_dalam_transaksi.get(k_item_upper, qty_bb_baris), 3)
                    
                    if kp_grup_utama in pibc_celup_gabung:
                        qty_bb_cek = round(gabungan_qty_item.get((kp_grup_utama, k_item_upper), qty_bb_mb_sendiri), 3)
                        info_cg = f"{info_cg_base} (Total: {qty_bb_cek} GR)"
                        if qty_bb_mb_sendiri != qty_bb_baris:
                            info_cg += " | Akumulasi Multi-Baris"
                        df_mb.loc[idx, 'Info Celup Gabung'] = info_cg
                    else:
                        qty_bb_cek = qty_bb_mb_sendiri
                        if qty_bb_mb_sendiri != qty_bb_baris:
                            df_mb.loc[idx, 'Info Celup Gabung'] = f"{peringatan_ganda} | Akumulasi Multi-Baris" if peringatan_ganda else "Akumulasi Multi-Baris"
                        elif peringatan_ganda:
                            df_mb.loc[idx, 'Info Celup Gabung'] = peringatan_ganda

                    df_mb.loc[idx, 'Qty BB Standar (GR)'] = float(qty_bb_mb_sendiri)

                    if len(pibc_list_grup) > 0:
                        qty_pibc_total = 0.0
                        qty_pg_total = 0.0
                        found_in_pibc = False
                        found_in_pg = False
                        pg_trans_gabung = {}
                        
                        k_kode = sanitize_key(kode_item)
                        k_nama = sanitize_key(nama_item)
                        
                        use_target_memphis = is_target_keyword(nama_item) and (qty_bb_cek < 100.0)
                        lookup_memphis = pibc_lookup_tgt if use_target_memphis else pibc_lookup_act
                        
                        for pibc_raw in pibc_list_grup:
                            kp_san = sanitize_key(pibc_raw)
                            if (kp_san, k_kode) in lookup_memphis:
                                qty_pibc_total += lookup_memphis[(kp_san, k_kode)]
                                found_in_pibc = True
                            elif (kp_san, k_nama) in lookup_memphis:
                                qty_pibc_total += lookup_memphis[(kp_san, k_nama)]
                                found_in_pibc = True

                            pg_k_kode = (g_norm, kp_san, k_kode)
                            pg_k_nama = (g_norm, kp_san, k_nama)
                            if pg_k_kode in pg_lookup:
                                qty_pg_total += pg_lookup[pg_k_kode]
                                found_in_pg = True
                                for nt, q_nt in pg_trans_lookup.get(pg_k_kode, {}).items():
                                    pg_trans_gabung[nt] = round(pg_trans_gabung.get(nt, 0.0) + q_nt, 3)
                            elif pg_k_nama in pg_lookup:
                                qty_pg_total += pg_lookup[pg_k_nama]
                                found_in_pg = True
                                for nt, q_nt in pg_trans_lookup.get(pg_k_nama, {}).items():
                                    pg_trans_gabung[nt] = round(pg_trans_gabung.get(nt, 0.0) + q_nt, 3)

                        sel_pibc = None
                        sel_pg = None
                        sel_pg_memphis = None
                        
                        if found_in_pibc:
                            qty_pibc_total = round(qty_pibc_total, 3)
                            df_mb.loc[idx, 'Qty Memphis'] = float(qty_pibc_total)
                            sel_pibc = round(qty_pibc_total - qty_bb_cek, 3)
                            if abs(sel_pibc) <= 0.009: sel_pibc = 0.0
                            df_mb.loc[idx, 'Selisih Memphis (GR)'] = float(sel_pibc)
                            
                        if found_in_pg:
                            qty_pg_total = round(qty_pg_total, 3)
                            df_mb.loc[idx, 'Qty PG'] = float(qty_pg_total)
                            sel_pg = round(qty_pg_total - qty_bb_cek, 3)
                            if abs(sel_pg) <= 0.009: sel_pg = 0.0
                            df_mb.loc[idx, 'Selisih PG (GR)'] = float(sel_pg)
                            
                            if pg_trans_gabung:
                                trans_str_list = [f"{nt} ({q_nt:,.2f} GR)" for nt, q_nt in pg_trans_gabung.items()]
                                df_mb.loc[idx, 'Kode Transaksi PG'] = ", ".join(trans_str_list)

                        if found_in_pg and found_in_pibc:
                            sel_pg_memphis = round(qty_pg_total - qty_pibc_total, 3)
                            if abs(sel_pg_memphis) <= 0.009: sel_pg_memphis = 0.0
                            df_mb.loc[idx, 'Selisih PG vs Memphis (GR)'] = float(sel_pg_memphis)

                        status_memphis = "TIDAK ADA" if not found_in_pibc else ("MATCH" if sel_pibc == 0.0 else "SELISIH")
                        status_pg = "TIDAK ADA" if not found_in_pg else ("MATCH" if sel_pg == 0.0 else "SELISIH")

                        if status_memphis == "TIDAK ADA" and status_pg == "TIDAK ADA":
                            df_mb.loc[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'
                        elif status_memphis == "MATCH" and status_pg == "MATCH":
                            df_mb.loc[idx, 'Crosscheck Qty'] = 'MATCH MEMPHIS & PG'
                        elif status_memphis == "MATCH" and status_pg == "SELISIH":
                            df_mb.loc[idx, 'Crosscheck Qty'] = 'MATCH MEMPHIS (PG SELISIH)'
                        elif status_memphis == "MATCH" and status_pg == "TIDAK ADA":
                            df_mb.loc[idx, 'Crosscheck Qty'] = 'MATCH MEMPHIS (TIDAK ADA DI PG)'
                        elif status_pg == "MATCH" and status_memphis == "SELISIH":
                            df_mb.loc[idx, 'Crosscheck Qty'] = 'MATCH PG (MEMPHIS SELISIH)'
                        elif status_pg == "MATCH" and status_memphis == "TIDAK ADA":
                            df_mb.loc[idx, 'Crosscheck Qty'] = 'MATCH PG (MESIN KECIL)' if is_mesin_kecil else 'MATCH PG (TIDAK ADA DI MEMPHIS)'
                        elif status_memphis == "SELISIH" and status_pg == "SELISIH":
                            if sel_pg_memphis == 0.0:
                                df_mb.loc[idx, 'Crosscheck Qty'] = 'SELISIH DI MB (MEMPHIS & PG MATCH)'
                            else:
                                df_mb.loc[idx, 'Crosscheck Qty'] = 'SELISIH SEMUA (MB, MEMPHIS, PG BEDA)'
                        elif status_memphis == "SELISIH" and status_pg == "TIDAK ADA":
                            df_mb.loc[idx, 'Crosscheck Qty'] = 'MEMPHIS SELISIH (TIDAK ADA DI PG)'
                        elif status_pg == "SELISIH" and status_memphis == "TIDAK ADA":
                            df_mb.loc[idx, 'Crosscheck Qty'] = 'PG SELISIH (MESIN KECIL)' if is_mesin_kecil else 'PG SELISIH (TIDAK ADA DI MEMPHIS)'
                    
                    obat_rendered_set.add(k_item_upper)
                else:
                    df_mb.loc[idx, 'Qty BB Standar (GR)'] = None
                    df_mb.loc[idx, 'Qty Target Standar (GR)'] = None
                    df_mb.loc[idx, 'Crosscheck Qty'] = 'DIABAIKAN (SUDAH DIGABUNG KE BARIS ATAS)'
                    df_mb.loc[idx, 'Selisih Memphis (GR)'] = None
                    df_mb.loc[idx, 'Selisih PG (GR)'] = None
                    df_mb.loc[idx, 'Selisih PG vs Memphis (GR)'] = None
                    df_mb.loc[idx, 'Kode Transaksi PG'] = ''

            is_first_row = False

        # 7. Finalisasi Baris Pertama Grup (3 Kolom Cek Item & Cek Variant)
        idx_first = df_sub.index[0]
        if is_gudang_audit:
            df_mb.loc[idx_first, 'Cek Variant'] = ", ".join(error_variants) if error_variants else 'CLEAN'
            
            if not grup_punya_pibc:
                df_mb.loc[idx_first, 'Cek Item (MB vs Memphis)'] = 'KODE PIBC TIDAK ADA'
                df_mb.loc[idx_first, 'Cek Item (MB vs PG)'] = 'KODE PIBC TIDAK ADA'
                df_mb.loc[idx_first, 'Cek Item (PG vs Memphis)'] = 'KODE PIBC TIDAK ADA'
            else:
                has_memphis = any(sanitize_key(p) in pibc_available_keys for p in pibc_list_grup)
                has_pg = any((gudang_group_norm, sanitize_key(p)) in pg_available_keys_by_gudang for p in pibc_list_grup)
                
                items_mb_final = gabungan_items_mb.get(kp_grup_utama, items_mb_grup) if is_celup_gabung else items_mb_grup
                
                items_memphis_final = {}
                for p_raw in pibc_list_grup:
                    for ck, info_it in pibc_items_by_key.get(sanitize_key(p_raw), {}).items():
                        if ck not in items_memphis_final:
                            items_memphis_final[ck] = info_it
                            
                items_pg_final = {}
                for p_raw in pibc_list_grup:
                    for ck, info_it in pg_items_by_gudang_key.get((gudang_group_norm, sanitize_key(p_raw)), {}).items():
                        if ck not in items_pg_final:
                            items_pg_final[ck] = info_it

                if not has_memphis:
                    df_mb.loc[idx_first, 'Cek Item (MB vs Memphis)'] = 'KODE PIBC TIDAK ADA DI MEMPHIS'
                else:
                    df_mb.loc[idx_first, 'Cek Item (MB vs Memphis)'] = bandingkan_set_item(
                        items_mb_final, items_memphis_final, 'MB', cek_lebih=(not is_mesin_kecil)
                    )

                if not has_pg:
                    df_mb.loc[idx_first, 'Cek Item (MB vs PG)'] = 'KODE PIBC TIDAK ADA DI PG'
                else:
                    df_mb.loc[idx_first, 'Cek Item (MB vs PG)'] = bandingkan_set_item(
                        items_mb_final, items_pg_final, 'MB', cek_lebih=True
                    )

                if not has_memphis and not has_pg:
                    df_mb.loc[idx_first, 'Cek Item (PG vs Memphis)'] = 'KODE PIBC TIDAK ADA DI MEMPHIS & PG'
                elif not has_memphis:
                    df_mb.loc[idx_first, 'Cek Item (PG vs Memphis)'] = 'KODE PIBC TIDAK ADA DI MEMPHIS'
                elif not has_pg:
                    df_mb.loc[idx_first, 'Cek Item (PG vs Memphis)'] = 'KODE PIBC TIDAK ADA DI PG'
                else:
                    df_mb.loc[idx_first, 'Cek Item (PG vs Memphis)'] = bandingkan_set_item(
                        items_pg_final, items_memphis_final, 'PG', cek_lebih=(not is_mesin_kecil)
                    )
        else:
            df_mb.loc[idx_first, 'Cek Item (MB vs Memphis)'] = ''
            df_mb.loc[idx_first, 'Cek Item (MB vs PG)'] = ''
            df_mb.loc[idx_first, 'Cek Item (PG vs Memphis)'] = ''
            
    if 'Ext_PIBC_List' in df_mb.columns:
        df_mb = df_mb.drop(columns=['Ext_PIBC_List'])

    cols = list(df_mb.columns)
    group_cols = [
        'Qty Target Standar (GR)', 'Qty BB Standar (GR)', 'Qty Memphis', 'Qty PG',
        'Selisih Benang (GR)', 'Selisih Memphis (GR)', 'Selisih PG (GR)', 'Selisih PG vs Memphis (GR)'
    ]
    
    for c in group_cols:
        if c in cols: cols.remove(c)
    
    if 'Crosscheck Qty' in cols:
        idx_insert = cols.index('Crosscheck Qty')
    else:
        idx_insert = len(cols)
        
    for c in reversed(group_cols):
        cols.insert(idx_insert, c)

    tail_order = [
        'Info Mesin', 'Info Celup Gabung',
        'Cek Item (MB vs Memphis)', 'Cek Item (MB vs PG)', 'Cek Item (PG vs Memphis)',
        'Status PIBC', 'Kode Transaksi PG', 'Rincian Benang', 'Cek Variant', 'Ext_PIBC'
    ]
    for tc in tail_order:
        if tc in cols: cols.remove(tc)
    for tc in tail_order:
        if tc in df_mb.columns: cols.append(tc)
        
    df_mb = df_mb[cols]

    for c in group_cols:
        if c in df_mb.columns:
            df_mb[c] = pd.to_numeric(df_mb[c], errors='coerce')

    return df_mb

def format_excel_sheet(ws, df, nama_kolom_kode='Kode', filter_end_row=None):
    header_fill = PatternFill(start_color='FFD51C39', end_color='FFD51C39', fill_type='solid')
    header_font = Font(color='FFFFFFFF', bold=True)
    
    wrap_cols = [
        'Keterangan', 'Keterangan Lain', 'Keterangan_PG', 'PIBC_PG',
        'Info Celup Gabung', 'Info Mesin',
        'Cek Item (MB vs Memphis)', 'Cek Item (MB vs PG)', 'Cek Item (PG vs Memphis)',
        'Status PIBC', 'Kode Transaksi PG', 'Rincian Benang', 'Cek Variant', 'Ext_PIBC',
        'Kode PIBC Terdeteksi', 'Catatan Auditor', 'Detail Kesalahan & Seharusnya',
        'Item Bermasalah', 'Keterangan Asli', 'Status di PG', 'Status Temuan PG',
        'Jenis Duplikat', 'Tindakan Sistem', 'Nama File Asal', 'PIBC / Keterangan'
    ]
    
    def is_int_col(c_name):
        c = str(c_name).lower()
        return any(k in c for k in [
            'masa pg', 'masa pibc', 'total transaksi', 'transaksi clean', 'transaksi bermasalah',
            'kasus ', 'pending > 7 hari', 'total baris'
        ])

    def is_pct_col(c_name):
        return '%' in str(c_name) or 'persentase' in str(c_name).lower()

    def is_numeric_col(c_name):
        c = str(c_name).lower()
        if is_int_col(c_name) or is_pct_col(c_name):
            return False
        kw = ['qty', 'harga', 'total', 'target', 'selisih', 'actual', 'jumlah', 'amount']
        if 'unit' in c or 'satuan' in c or 'crosscheck' in c or 'status' in c or 'keterangan' in c or 'info' in c or 'transaksi' in c:
            return False
        return any(k in c for k in kw)

    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = HEADER_BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    warna_hex = ['FFBDB2FF', 'FFFFCEE3', 'FFFDFFD2', 'FFCAEDFF']
    color_fills = [PatternFill(start_color=c, end_color=c, fill_type='solid') for c in warna_hex]
    
    wrap_alignment = Alignment(wrap_text=True, vertical='top')
    std_alignment = Alignment(vertical='top')
    
    idx_kode = None
    if nama_kolom_kode and nama_kolom_kode in df.columns:
        idx_kode = list(df.columns).index(nama_kolom_kode) + 1
        
    current_kode = None
    color_index = -1
    
    for r_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column), start=2):
        if idx_kode is not None:
            cell_kode_val = ws.cell(row=r_idx, column=idx_kode).value
            if cell_kode_val != current_kode:
                current_kode = cell_kode_val
                color_index = (color_index + 1) % len(color_fills)
            fill_to_apply = color_fills[color_index]
        else:
            fill_to_apply = None

        for c_idx, cell in enumerate(row, start=1):
            if fill_to_apply is not None:
                cell.fill = fill_to_apply
            cell.border = CELL_BORDER
            col_name = df.columns[c_idx - 1]
            
            if col_name in wrap_cols:
                cell.alignment = wrap_alignment
            else:
                cell.alignment = std_alignment
                
            if cell.value is not None and str(cell.value).strip() != '':
                if str(cell.value).strip().lower() in ['nan', 'none', 'nat', '<na>']:
                    cell.value = None
                elif is_int_col(col_name):
                    try:
                        cell.value = int(round(float(cell.value)))
                        cell.number_format = '#,##0'
                    except:
                        pass
                elif is_pct_col(col_name):
                    try:
                        cell.value = float(cell.value)
                        cell.number_format = '0.00"%"'
                    except:
                        pass
                elif is_numeric_col(col_name):
                    try:
                        if isinstance(cell.value, str):
                            val_str = str(cell.value).strip().replace('.', '').replace(',', '.')
                            val_float = float(val_str)
                        else:
                            val_float = float(cell.value)
                        cell.value = val_float
                        cell.number_format = '#,##0.00'
                    except:
                        pass

    for col_idx, col_name in enumerate(df.columns, start=1):
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        max_len = len(str(col_name)) + 4
        for r_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=r_idx, column=col_idx).value
            if val is None:
                continue
            if isinstance(val, int):
                val_str = f"{val:,}"
            elif isinstance(val, float):
                val_str = f"{val:,.2f}"
            else:
                val_str = str(val)
            line_max = max((len(line) for line in val_str.split('\n')), default=0)
            if line_max > max_len:
                max_len = line_max
        if col_name in wrap_cols:
            final_width = min(max(max_len + 3, 14), 55)
        else:
            final_width = min(max(max_len + 4, 12), 45)
        ws.column_dimensions[col_letter].width = final_width

    ws.freeze_panes = 'A2'
    if ws.max_row >= 1 and ws.max_column >= 1:
        if filter_end_row:
            last_col_letter = ws.cell(row=1, column=ws.max_column).column_letter
            ws.auto_filter.ref = f"A1:{last_col_letter}{filter_end_row}"
        else:
            ws.auto_filter.ref = ws.dimensions

def format_excel_sheet_cepat(ws, df):
    """Formatting cepat berbasis Pandas dengan garis border tabel tegas untuk sheet Master."""
    header_fill = PatternFill(start_color='FFD51C39', end_color='FFD51C39', fill_type='solid')
    header_font = Font(color='FFFFFFFF', bold=True)
    
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = HEADER_BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
        col_letter = cell.column_letter
        if not df.empty:
            seri_len = df[col_name].dropna().astype(str).str.len()
            max_data_len = int(seri_len.max()) if not seri_len.empty else 0
        else:
            max_data_len = 0
        final_width = min(max(max(len(str(col_name)) + 4, max_data_len + 3), 12), 48)
        ws.column_dimensions[col_letter].width = final_width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = CELL_BORDER

    ws.freeze_panes = 'A2'
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

def _tulis_sheet_master_clean(writer, df_pibc, df_pg):
    if df_pibc is not None and not df_pibc.empty:
        df_pibc.to_excel(writer, index=False, sheet_name='Master Memphis (Clean)')
        format_excel_sheet_cepat(writer.sheets['Master Memphis (Clean)'], df_pibc)
    else:
        df_info = pd.DataFrame({'Info': ['Master Memphis kosong / belum diupload']})
        df_info.to_excel(writer, index=False, sheet_name='Master Memphis (Clean)')
        format_excel_sheet_cepat(writer.sheets['Master Memphis (Clean)'], df_info)

    if df_pg is not None and not df_pg.empty:
        df_pg.to_excel(writer, index=False, sheet_name='Master PG (Clean)')
        format_excel_sheet_cepat(writer.sheets['Master PG (Clean)'], df_pg)
    else:
        df_info = pd.DataFrame({'Info': ['Master PG kosong / belum diupload']})
        df_info.to_excel(writer, index=False, sheet_name='Master PG (Clean)')
        format_excel_sheet_cepat(writer.sheets['Master PG (Clean)'], df_info)

def _dapatkan_df_belum_diproses_pibc(df_hasil_mb, df_pibc):
    if df_pibc is None or df_pibc.empty:
        return pd.DataFrame()
    col_ticket = cari_kolom_akurat(df_pibc, ['NOTICKET', 'DYELOT', 'PIBC'])
    if not col_ticket:
        return pd.DataFrame()
    pibc_di_mb_expanded = set()
    for ext in df_hasil_mb['Ext_PIBC'].dropna():
        if str(ext).strip():
            for p in str(ext).split(','):
                pibc_di_mb_expanded.add(sanitize_key(p.strip()))
    df_belum = df_pibc[~df_pibc[col_ticket].astype(str).apply(sanitize_key).isin(pibc_di_mb_expanded)].copy()
    col_tgl = cari_kolom_akurat(df_belum, ['TANGGAL', 'TGL', 'DATE'])
    return hitung_masa_dan_rentang(df_belum, col_tgl, 'Masa PIBC')

def _dapatkan_df_belum_diproses_pg(df_hasil_mb, df_pg):
    if df_pg is None or df_pg.empty or 'PIBC_PG' not in df_pg.columns:
        return pd.DataFrame()
    mb_gudang_pibc = set()
    for _, r in df_hasil_mb.iterrows():
        ext = str(r.get('Ext_PIBC', '')).strip()
        if not ext: continue
        g_norm = normalisasi_gudang(r.get('Gudang', ''))
        for p in ext.split(','):
            kp = sanitize_key(p.strip())
            if kp: mb_gudang_pibc.add((g_norm, kp))

    gudang_audit_set = {'Gudang Mesin Dyeing', 'Gudang Lab & RnD'}
    mask_keep = []
    for _, r in df_pg.iterrows():
        g_norm_pg = normalisasi_gudang(r.get('Gudang_PG', ''))
        if g_norm_pg not in gudang_audit_set:
            mask_keep.append(False)
            continue
        ext_list = ekstrak_semua_pibc(str(r.get('PIBC_PG', '')))
        if not ext_list:
            mask_keep.append(False)
            continue
        sudah_mb = any((g_norm_pg, sanitize_key(p)) in mb_gudang_pibc for p in ext_list)
        mask_keep.append(not sudah_mb)

    df_belum = df_pg[mask_keep].copy()
    col_tgl = cari_kolom_akurat(df_belum, ['TANGGALPG', 'TANGGAL', 'TGL', 'DATE'])
    return hitung_masa_dan_rentang(df_belum, col_tgl, 'Masa PG')

def buat_excel_hasil_keseluruhan(df_hasil, df_pibc=None, df_pg=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_hasil.to_excel(writer, index=False, sheet_name='Hasil Crosscheck')
        format_excel_sheet(writer.sheets['Hasil Crosscheck'], df_hasil, 'Kode')
        _tulis_sheet_master_clean(writer, df_pibc, df_pg)
    return output.getvalue()

def buat_excel_rekap_audit(df_hasil, df_pibc=None, df_pg=None, df_log_duplikat=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        df_belum_pibc = _dapatkan_df_belum_diproses_pibc(df_hasil, df_pibc)
        df_belum_pg = _dapatkan_df_belum_diproses_pg(df_hasil, df_pg)

        pg_over7_per_gudang = {}
        if not df_belum_pg.empty and 'Masa PG' in df_belum_pg.columns:
            df_pg7 = df_belum_pg[df_belum_pg['Masa PG'].fillna(-1) > 7]
            for g_raw in df_pg7['Gudang_PG']:
                gn = normalisasi_gudang(g_raw)
                pg_over7_per_gudang[gn] = pg_over7_per_gudang.get(gn, 0) + 1

        data_ringkas = []
        kpi_stats = {}
        kode_kelengkapan_err = set()
        kode_qty_err = set()
        kode_var_err = set()
        kode_ms_set = set()

        grup_hasil = df_hasil.groupby('Kode', sort=False) if 'Kode' in df_hasil.columns else [(None, df_hasil)]

        for k_val, df_sub in grup_hasil:
            first_row = df_sub.iloc[0]
            gudang_norm = normalisasi_gudang(first_row.get('Gudang', ''))
            if gudang_norm not in kpi_stats:
                kpi_stats[gudang_norm] = {
                    'total': 0, 'clean': 0, 'bermasalah': 0,
                    'err_qty': 0, 'err_var': 0, 'err_kelengkapan': 0,
                    'pibc_kosong': 0, 'pibc_miss_pg': 0, 'mesin_tdk_terdaftar': 0
                }
            kpi_stats[gudang_norm]['total'] += 1

            c_mm = str(first_row.get('Cek Item (MB vs Memphis)', '')).strip()
            c_mp = str(first_row.get('Cek Item (MB vs PG)', '')).strip()
            c_pm = str(first_row.get('Cek Item (PG vs Memphis)', '')).strip()
            info_m = str(first_row.get('Info Mesin', '')).strip()

            pibc_kosong_di_mb = (c_mm == 'KODE PIBC TIDAK ADA')
            pibc_miss_memphis = (c_mm == 'KODE PIBC TIDAK ADA DI MEMPHIS')
            pibc_miss_pg = (c_mp == 'KODE PIBC TIDAK ADA DI PG')
            pibc_miss_both = pibc_miss_memphis and pibc_miss_pg

            ada_err_kelengkapan = is_cek_item_bermasalah(c_mm) or is_cek_item_bermasalah(c_mp) or is_cek_item_bermasalah(c_pm)

            if pibc_kosong_di_mb:
                kpi_stats[gudang_norm]['pibc_kosong'] += 1
            if pibc_miss_both or pibc_miss_pg:
                kpi_stats[gudang_norm]['pibc_miss_pg'] += 1
            if 'TIDAK TERDAFTAR' in info_m.upper():
                kpi_stats[gudang_norm]['mesin_tdk_terdaftar'] += 1

            ada_err_qty_grup = False
            ada_err_var_grup = False
            ada_item_nf_spesifik = False
            baris_ringkas_grup = []

            for _, row in df_sub.iterrows():
                cq = str(row.get('Crosscheck Qty', '')).upper()
                cv = str(row.get('Cek Variant', '')).upper()

                if cq == 'MEMANG SELISIH':
                    kode_ms_set.add(k_val)

                is_qty_err = ('INCORRECT' in cq or 'SELISIH' in cq) and cq != 'MEMANG SELISIH'
                is_var_err = cv != '' and cv != 'CLEAN'
                is_item_nf_row = ('ITEM TIDAK DI TEMUKAN' in cq) and (not pibc_miss_both) and (not pibc_miss_pg)

                if is_qty_err: ada_err_qty_grup = True
                if is_var_err: ada_err_var_grup = True
                if is_item_nf_row: ada_item_nf_spesifik = True

                if is_qty_err or is_var_err or is_item_nf_row:
                    item_bermasalah = f"{row.get('Kode Item', '')} - {row.get('Nama Item', '')}"
                    if is_qty_err and is_var_err:
                        jenis = "Dobel Error (Varian & Qty)"
                        detail = f"Varian: {cv}. Qty: {cq}"
                    elif is_item_nf_row and is_var_err:
                        jenis = "Dobel Error (Varian & Item Tidak Ditemukan)"
                        detail = f"Varian: {cv}. Item: {cq}"
                    elif is_qty_err:
                        jenis = "Selisih Kuantitas"
                        detail = cq
                    elif is_item_nf_row:
                        jenis = "Item Tidak Ditemukan"
                        detail = cq
                    else:
                        jenis = "Pelanggaran SOP Varian"
                        detail = cv
                    
                    info_cg = str(row.get('Info Celup Gabung', '')).strip()
                    if info_cg:
                        detail = f"{detail} | [INFO: {info_cg}]" if detail else f"[INFO: {info_cg}]"
                        
                    baris_ringkas_grup.append({
                        'Gudang': row.get('Gudang', ''),
                        'Kode': row.get('Kode', ''),
                        'Tanggal': row.get('Tanggal', ''),
                        'Info Mesin': row.get('Info Mesin', ''),
                        'Keterangan': row.get('Keterangan', ''),
                        'Keterangan Lain': row.get('Keterangan Lain', ''),
                        'Kode PIBC Terdeteksi': row.get('Ext_PIBC', ''),
                        'Item Bermasalah': item_bermasalah,
                        'Qty Target Standar (GR)': row.get('Qty Target Standar (GR)') if pd.notna(row.get('Qty Target Standar (GR)')) else None,
                        'Qty BB Standar (GR)': row.get('Qty BB Standar (GR)') if pd.notna(row.get('Qty BB Standar (GR)')) else None,
                        'Qty Memphis': row.get('Qty Memphis') if pd.notna(row.get('Qty Memphis')) else None,
                        'Qty PG': row.get('Qty PG') if pd.notna(row.get('Qty PG')) else None,
                        'Selisih Benang (GR)': row.get('Selisih Benang (GR)') if pd.notna(row.get('Selisih Benang (GR)')) else None,
                        'Selisih Memphis (GR)': row.get('Selisih Memphis (GR)') if pd.notna(row.get('Selisih Memphis (GR)')) else None,
                        'Selisih PG (GR)': row.get('Selisih PG (GR)') if pd.notna(row.get('Selisih PG (GR)')) else None,
                        'Selisih PG vs Memphis (GR)': row.get('Selisih PG vs Memphis (GR)') if pd.notna(row.get('Selisih PG vs Memphis (GR)')) else None,
                        'Jenis Temuan': jenis,
                        'Detail Kesalahan & Seharusnya': detail,
                        'Cek Item (MB vs Memphis)': c_mm,
                        'Cek Item (MB vs PG)': c_mp,
                        'Cek Item (PG vs Memphis)': c_pm,
                        'Kode Transaksi PG': row.get('Kode Transaksi PG', ''),
                        'Catatan Auditor': ''
                    })

            if ada_item_nf_spesifik:
                ada_err_kelengkapan = True

            if ada_err_kelengkapan:
                kode_kelengkapan_err.add(k_val)
                kpi_stats[gudang_norm]['err_kelengkapan'] += 1
            if ada_err_qty_grup:
                kode_qty_err.add(k_val)
                kpi_stats[gudang_norm]['err_qty'] += 1
            if ada_err_var_grup:
                kode_var_err.add(k_val)
                kpi_stats[gudang_norm]['err_var'] += 1

            if ada_err_qty_grup or ada_err_var_grup or ada_err_kelengkapan or pibc_miss_both or pibc_miss_pg:
                kpi_stats[gudang_norm]['bermasalah'] += 1
            else:
                kpi_stats[gudang_norm]['clean'] += 1

            if baris_ringkas_grup:
                data_ringkas.extend(baris_ringkas_grup)

            if pibc_miss_both:
                data_ringkas.append({
                    'Gudang': first_row.get('Gudang', ''),
                    'Kode': first_row.get('Kode', ''),
                    'Tanggal': first_row.get('Tanggal', ''),
                    'Info Mesin': info_m,
                    'Keterangan': first_row.get('Keterangan', ''),
                    'Keterangan Lain': first_row.get('Keterangan Lain', ''),
                    'Kode PIBC Terdeteksi': first_row.get('Ext_PIBC', ''),
                    'Item Bermasalah': '(Kode PIBC Tidak Ada di Master Memphis & PG)',
                    'Qty Target Standar (GR)': None,
                    'Qty BB Standar (GR)': None,
                    'Qty Memphis': None,
                    'Qty PG': None,
                    'Selisih Benang (GR)': None,
                    'Selisih Memphis (GR)': None,
                    'Selisih PG (GR)': None,
                    'Selisih PG vs Memphis (GR)': None,
                    'Jenis Temuan': 'PIBC Tidak Ada di Memphis & PG',
                    'Detail Kesalahan & Seharusnya': 'Kode PIBC tidak ditemukan di Master Memphis maupun Master PG',
                    'Cek Item (MB vs Memphis)': c_mm,
                    'Cek Item (MB vs PG)': c_mp,
                    'Cek Item (PG vs Memphis)': c_pm,
                    'Kode Transaksi PG': '',
                    'Catatan Auditor': ''
                })
            elif pibc_miss_pg:
                data_ringkas.append({
                    'Gudang': first_row.get('Gudang', ''),
                    'Kode': first_row.get('Kode', ''),
                    'Tanggal': first_row.get('Tanggal', ''),
                    'Info Mesin': info_m,
                    'Keterangan': first_row.get('Keterangan', ''),
                    'Keterangan Lain': first_row.get('Keterangan Lain', ''),
                    'Kode PIBC Terdeteksi': first_row.get('Ext_PIBC', ''),
                    'Item Bermasalah': '(Kode PIBC Tidak Ada di Master PG)',
                    'Qty Target Standar (GR)': None,
                    'Qty BB Standar (GR)': None,
                    'Qty Memphis': None,
                    'Qty PG': None,
                    'Selisih Benang (GR)': None,
                    'Selisih Memphis (GR)': None,
                    'Selisih PG (GR)': None,
                    'Selisih PG vs Memphis (GR)': None,
                    'Jenis Temuan': 'PIBC Tidak Ada di PG',
                    'Detail Kesalahan & Seharusnya': 'Kode PIBC tidak ditemukan di Master PG pada gudang terkait',
                    'Cek Item (MB vs Memphis)': c_mm,
                    'Cek Item (MB vs PG)': c_mp,
                    'Cek Item (PG vs Memphis)': c_pm,
                    'Kode Transaksi PG': '',
                    'Catatan Auditor': ''
                })
            elif (not baris_ringkas_grup) and ada_err_kelengkapan:
                sumber_err = []
                if is_cek_item_bermasalah(c_mm): sumber_err.append("MB vs Memphis")
                if is_cek_item_bermasalah(c_mp): sumber_err.append("MB vs PG")
                if is_cek_item_bermasalah(c_pm): sumber_err.append("PG vs Memphis")
                data_ringkas.append({
                    'Gudang': first_row.get('Gudang', ''),
                    'Kode': first_row.get('Kode', ''),
                    'Tanggal': first_row.get('Tanggal', ''),
                    'Info Mesin': info_m,
                    'Keterangan': first_row.get('Keterangan', ''),
                    'Keterangan Lain': first_row.get('Keterangan Lain', ''),
                    'Kode PIBC Terdeteksi': first_row.get('Ext_PIBC', ''),
                    'Item Bermasalah': '(Lihat Rincian di Kolom Cek Item)',
                    'Qty Target Standar (GR)': None,
                    'Qty BB Standar (GR)': None,
                    'Qty Memphis': None,
                    'Qty PG': None,
                    'Selisih Benang (GR)': None,
                    'Selisih Memphis (GR)': None,
                    'Selisih PG (GR)': None,
                    'Selisih PG vs Memphis (GR)': None,
                    'Jenis Temuan': 'Ketidaklengkapan Item',
                    'Detail Kesalahan & Seharusnya': f"Item tidak lengkap/berlebih pada: {', '.join(sumber_err)}",
                    'Cek Item (MB vs Memphis)': c_mm,
                    'Cek Item (MB vs PG)': c_mp,
                    'Cek Item (PG vs Memphis)': c_pm,
                    'Kode Transaksi PG': '',
                    'Catatan Auditor': ''
                })

        # --- Sheet 1: Ringkasan KPI ---
        kpi_rows = []
        tot_all = {k: 0 for k in ['total', 'clean', 'bermasalah', 'err_qty', 'err_var', 'err_kelengkapan', 'pibc_kosong', 'pibc_miss_pg', 'mesin_tdk_terdaftar', 'pg_over7']}
        for g_name, st_val in kpi_stats.items():
            pg7 = pg_over7_per_gudang.get(g_name, 0)
            pct = round((st_val['clean'] / st_val['total'] * 100.0), 2) if st_val['total'] > 0 else 0.0
            kpi_rows.append({
                'Gudang': g_name,
                'Total Transaksi MB': st_val['total'],
                'Transaksi Clean': st_val['clean'],
                'Transaksi Bermasalah': st_val['bermasalah'],
                'Persentase Clean (%)': pct,
                'Kasus Error Qty': st_val['err_qty'],
                'Kasus Error Varian': st_val['err_var'],
                'Kasus Ketidaklengkapan Item': st_val['err_kelengkapan'],
                'Kasus PIBC Tidak Ada di PG': st_val['pibc_miss_pg'],
                'Kasus PIBC Tidak Ada di MB': st_val['pibc_kosong'],
                'Kasus Mesin Tidak Terdaftar': st_val['mesin_tdk_terdaftar'],
                'PG Pending > 7 Hari': pg7
            })
            for key_t in ['total', 'clean', 'bermasalah', 'err_qty', 'err_var', 'err_kelengkapan', 'pibc_kosong', 'pibc_miss_pg', 'mesin_tdk_terdaftar']:
                tot_all[key_t] += st_val[key_t]
            tot_all['pg_over7'] += pg7

        pct_tot = round((tot_all['clean'] / tot_all['total'] * 100.0), 2) if tot_all['total'] > 0 else 0.0
        kpi_rows.append({
            'Gudang': 'TOTAL KESELURUHAN',
            'Total Transaksi MB': tot_all['total'],
            'Transaksi Clean': tot_all['clean'],
            'Transaksi Bermasalah': tot_all['bermasalah'],
            'Persentase Clean (%)': pct_tot,
            'Kasus Error Qty': tot_all['err_qty'],
            'Kasus Error Varian': tot_all['err_var'],
            'Kasus Ketidaklengkapan Item': tot_all['err_kelengkapan'],
            'Kasus PIBC Tidak Ada di PG': tot_all['pibc_miss_pg'],
            'Kasus PIBC Tidak Ada di MB': tot_all['pibc_kosong'],
            'Kasus Mesin Tidak Terdaftar': tot_all['mesin_tdk_terdaftar'],
            'PG Pending > 7 Hari': tot_all['pg_over7']
        })
        df_kpi = pd.DataFrame(kpi_rows)
        df_kpi.to_excel(writer, index=False, sheet_name='Ringkasan KPI')
        ws_kpi = writer.sheets['Ringkasan KPI']
        format_excel_sheet(ws_kpi, df_kpi, None, filter_end_row=len(df_kpi) + 1)

        start_r = len(df_kpi) + 4
        mem_tot = len(df_belum_pibc)
        mem_over7 = int((df_belum_pibc['Masa PIBC'].fillna(-1) > 7).sum()) if (not df_belum_pibc.empty and 'Masa PIBC' in df_belum_pibc.columns) else 0
        pg_tot = len(df_belum_pg)
        pg_over7_tot = int((df_belum_pg['Masa PG'].fillna(-1) > 7).sum()) if (not df_belum_pg.empty and 'Masa PG' in df_belum_pg.columns) else 0

        hdr_fill = PatternFill(start_color='FFD51C39', end_color='FFD51C39', fill_type='solid')
        hdr_font = Font(color='FFFFFFFF', bold=True)
        sub_headers = ["Ringkasan Dokumen Belum Diproses (Aging)", "Total Baris", "Pending > 7 Hari"]
        for c_i, h_text in enumerate(sub_headers, start=1):
            c = ws_kpi.cell(row=start_r, column=c_i, value=h_text)
            c.fill = hdr_fill
            c.font = hdr_font
            c.border = HEADER_BORDER
            c.alignment = Alignment(horizontal='center', vertical='center')

        sub_rows = [
            ("Master Memphis Belum Diproses", mem_tot, mem_over7),
            ("Master PG Belum Diproses", pg_tot, pg_over7_tot)
        ]
        for r_offset, (label_s, v_tot, v_over7) in enumerate(sub_rows, start=1):
            c1 = ws_kpi.cell(row=start_r + r_offset, column=1, value=label_s)
            c2 = ws_kpi.cell(row=start_r + r_offset, column=2, value=v_tot)
            c3 = ws_kpi.cell(row=start_r + r_offset, column=3, value=v_over7)
            c1.border = CELL_BORDER
            c2.border = CELL_BORDER
            c3.border = CELL_BORDER
            c2.number_format = '#,##0'
            c3.number_format = '#,##0'

        # --- Sheet 2: Laporan Ringkas Admin ---
        kolom_fix = [
            'Gudang', 'Kode', 'Tanggal', 'Info Mesin', 'Keterangan', 'Keterangan Lain',
            'Kode PIBC Terdeteksi', 'Item Bermasalah', 'Qty Target Standar (GR)',
            'Qty BB Standar (GR)', 'Qty Memphis', 'Qty PG', 'Selisih Benang (GR)',
            'Selisih Memphis (GR)', 'Selisih PG (GR)', 'Selisih PG vs Memphis (GR)',
            'Jenis Temuan', 'Detail Kesalahan & Seharusnya',
            'Cek Item (MB vs Memphis)', 'Cek Item (MB vs PG)', 'Cek Item (PG vs Memphis)',
            'Kode Transaksi PG', 'Catatan Auditor'
        ]
        
        df_ringkas = pd.DataFrame(data_ringkas)
        if not df_ringkas.empty:
            df_ringkas = df_ringkas[kolom_fix] 
            df_ringkas.to_excel(writer, index=False, sheet_name='Laporan Ringkas Admin')
            format_excel_sheet(writer.sheets['Laporan Ringkas Admin'], df_ringkas, 'Kode')
        else:
            df_empty_ringkas = pd.DataFrame(columns=kolom_fix)
            df_empty_ringkas.to_excel(writer, index=False, sheet_name='Laporan Ringkas Admin')
            format_excel_sheet(writer.sheets['Laporan Ringkas Admin'], df_empty_ringkas, None)
            
        # --- Sheet 3: Error Qty (Full Grup) ---
        if len(kode_qty_err) > 0:
            df_qty_full = df_hasil[df_hasil['Kode'].isin(kode_qty_err)]
            df_qty_full.to_excel(writer, index=False, sheet_name='Error Qty (Full Grup)')
            format_excel_sheet(writer.sheets['Error Qty (Full Grup)'], df_qty_full, 'Kode')
        else:
            df_info = pd.DataFrame({'Info': ['Tidak ada temuan']})
            df_info.to_excel(writer, index=False, sheet_name='Error Qty (Full Grup)')
            format_excel_sheet(writer.sheets['Error Qty (Full Grup)'], df_info, None)
            
        # --- Sheet 4: Error Varian (Full Grup) ---
        if len(kode_var_err) > 0:
            df_var_full = df_hasil[df_hasil['Kode'].isin(kode_var_err)]
            df_var_full.to_excel(writer, index=False, sheet_name='Error Varian (Full Grup)')
            format_excel_sheet(writer.sheets['Error Varian (Full Grup)'], df_var_full, 'Kode')
        else:
            df_info = pd.DataFrame({'Info': ['Tidak ada temuan']})
            df_info.to_excel(writer, index=False, sheet_name='Error Varian (Full Grup)')
            format_excel_sheet(writer.sheets['Error Varian (Full Grup)'], df_info, None)

        # --- Sheet 5: Error Kelengkapan (Full Grup) ---
        if len(kode_kelengkapan_err) > 0:
            df_kel_full = df_hasil[df_hasil['Kode'].isin(kode_kelengkapan_err)]
            df_kel_full.to_excel(writer, index=False, sheet_name='Error Kelengkapan (Full Grup)')
            format_excel_sheet(writer.sheets['Error Kelengkapan (Full Grup)'], df_kel_full, 'Kode')
        else:
            df_info = pd.DataFrame({'Info': ['Tidak ada temuan']})
            df_info.to_excel(writer, index=False, sheet_name='Error Kelengkapan (Full Grup)')
            format_excel_sheet(writer.sheets['Error Kelengkapan (Full Grup)'], df_info, None)
            
        # --- Sheet 6: Memang Selisih (Full Grup) ---
        if len(kode_ms_set) > 0:
            df_ms_full = df_hasil[df_hasil['Kode'].isin(kode_ms_set)]
            df_ms_full.to_excel(writer, index=False, sheet_name='Memang Selisih (Full Grup)')
            format_excel_sheet(writer.sheets['Memang Selisih (Full Grup)'], df_ms_full, 'Kode')
        else:
            df_info = pd.DataFrame({'Info': ['Tidak ada temuan']})
            df_info.to_excel(writer, index=False, sheet_name='Memang Selisih (Full Grup)')
            format_excel_sheet(writer.sheets['Memang Selisih (Full Grup)'], df_info, None)

        # --- Sheet 7: Log Duplikat (Jejak Audit Data Ganda) ---
        if df_log_duplikat is not None and not df_log_duplikat.empty:
            df_log_duplikat.to_excel(writer, index=False, sheet_name='Log Duplikat')
            format_excel_sheet(writer.sheets['Log Duplikat'], df_log_duplikat, 'No. Transaksi / Kode')
        else:
            df_info = pd.DataFrame({'Info': ['Tidak ditemukan data duplikat (lintas file maupun dalam 1 file)']})
            df_info.to_excel(writer, index=False, sheet_name='Log Duplikat')
            format_excel_sheet(writer.sheets['Log Duplikat'], df_info, None)

        # --- Sheet 8 & 9: Master Clean ---
        _tulis_sheet_master_clean(writer, df_pibc, df_pg)
            
    return output.getvalue()

def buat_excel_investigasi_pibc(df_hasil_mb, df_pibc, df_pg=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data_bodong = []
        pibc_di_master = set()
        col_ticket = None
        
        if df_pibc is not None and not df_pibc.empty:
            col_ticket = cari_kolom_akurat(df_pibc, ['NOTICKET', 'DYELOT', 'PIBC'])
            if col_ticket:
                pibc_di_master = set(df_pibc[col_ticket].dropna().astype(str).apply(sanitize_key))

        pg_by_gudang = set()
        if df_pg is not None and not df_pg.empty and 'PIBC_PG' in df_pg.columns:
            for _, r in df_pg.iterrows():
                g_pg = normalisasi_gudang(r.get('Gudang_PG', ''))
                for p in ekstrak_semua_pibc(str(r.get('PIBC_PG', ''))):
                    kp = sanitize_key(p)
                    if kp: pg_by_gudang.add((g_pg, kp))

        seen_kode_mb = set()
        for _, row in df_hasil_mb.iterrows():
            kode_mb = str(row.get('Kode', '')).strip()
            if kode_mb in seen_kode_mb:
                continue
            ext_pibc = str(row.get('Ext_PIBC', '')).strip()
            if ext_pibc:
                seen_kode_mb.add(kode_mb)
                g_mb = normalisasi_gudang(row.get('Gudang', ''))
                list_pibc = [p.strip() for p in ext_pibc.split(',') if p.strip()]
                bodong = [p for p in list_pibc if sanitize_key(p) not in pibc_di_master]
                if bodong:
                    ada_di_pg = [p for p in bodong if (g_mb, sanitize_key(p)) in pg_by_gudang]
                    if len(ada_di_pg) == len(bodong):
                        status_pg_str = 'Ada di PG (Hanya Kosong di Memphis)'
                    elif len(ada_di_pg) == 0:
                        status_pg_str = 'Tidak Ada di Memphis & PG'
                    else:
                        status_pg_str = f'Sebagian Ada di PG ({", ".join(ada_di_pg)})'

                    data_bodong.append({
                        'Gudang': row.get('Gudang', ''),
                        'Kode MB': kode_mb,
                        'Tanggal': row.get('Tanggal', ''),
                        'Kode PIBC Tidak Ditemukan': ", ".join(bodong),
                        'Status di PG': status_pg_str,
                        'Keterangan': row.get('Keterangan', ''),
                        'Keterangan Lain': row.get('Keterangan Lain', '')
                    })

        df_bodong = pd.DataFrame(data_bodong)
        if not df_bodong.empty:
            df_bodong.to_excel(writer, index=False, sheet_name='PIBC MB Tidak Ditemukan')
            format_excel_sheet(writer.sheets['PIBC MB Tidak Ditemukan'], df_bodong, 'Kode MB')
        else:
            df_info = pd.DataFrame({'Info': ['Semua PIBC di MB terdaftar valid di Master Memphis']})
            df_info.to_excel(writer, index=False, sheet_name='PIBC MB Tidak Ditemukan')
            format_excel_sheet(writer.sheets['PIBC MB Tidak Ditemukan'], df_info, None)
        
        if df_pibc is not None and not df_pibc.empty and col_ticket:
            df_belum = _dapatkan_df_belum_diproses_pibc(df_hasil_mb, df_pibc)
            if not df_belum.empty:
                df_belum.to_excel(writer, index=False, sheet_name='Master PIBC Belum Diproses')
                format_excel_sheet(writer.sheets['Master PIBC Belum Diproses'], df_belum, None)
            else:
                df_info = pd.DataFrame({'Info': ['Semua data Master Memphis sudah diproses di MB']})
                df_info.to_excel(writer, index=False, sheet_name='Master PIBC Belum Diproses')
                format_excel_sheet(writer.sheets['Master PIBC Belum Diproses'], df_info, None)
        else:
            df_info = pd.DataFrame({'Info': ['Master Memphis kosong/belum diupload']})
            df_info.to_excel(writer, index=False, sheet_name='Master PIBC Belum Diproses')
            format_excel_sheet(writer.sheets['Master PIBC Belum Diproses'], df_info, None)
            
    return output.getvalue()

def buat_excel_investigasi_pg(df_hasil_mb, df_pg):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data_bodong = []
        pg_by_gudang = set()
        pg_gudang_map = {}
        
        if df_pg is not None and not df_pg.empty and 'PIBC_PG' in df_pg.columns:
            for _, r in df_pg.iterrows():
                g_pg = normalisasi_gudang(r.get('Gudang_PG', ''))
                for p in ekstrak_semua_pibc(str(r.get('PIBC_PG', ''))):
                    kp = sanitize_key(p)
                    if kp:
                        pg_by_gudang.add((g_pg, kp))
                        if kp not in pg_gudang_map: pg_gudang_map[kp] = set()
                        pg_gudang_map[kp].add(g_pg)

        seen_kode_mb = set()
        for _, row in df_hasil_mb.iterrows():
            kode_mb = str(row.get('Kode', '')).strip()
            if kode_mb in seen_kode_mb:
                continue
            ext_pibc = str(row.get('Ext_PIBC', '')).strip()
            if ext_pibc:
                seen_kode_mb.add(kode_mb)
                g_mb = normalisasi_gudang(row.get('Gudang', ''))
                list_pibc = [p.strip() for p in ext_pibc.split(',') if p.strip()]
                bodong = [p for p in list_pibc if (g_mb, sanitize_key(p)) not in pg_by_gudang]
                if bodong:
                    gudang_lain_list = set()
                    for p in bodong:
                        kp = sanitize_key(p)
                        if kp in pg_gudang_map:
                            for gl in pg_gudang_map[kp]:
                                if gl != g_mb:
                                    gudang_lain_list.add(gl)
                    if gudang_lain_list:
                        status_temuan = f"BEDA GUDANG (Di PG tercatat di: {', '.join(sorted(gudang_lain_list))})"
                    else:
                        status_temuan = "TIDAK ADA DI PG SAMA SEKALI"

                    data_bodong.append({
                        'Gudang': row.get('Gudang', ''),
                        'Kode MB': kode_mb,
                        'Tanggal': row.get('Tanggal', ''),
                        'Kode PG Tidak Ditemukan': ", ".join(bodong),
                        'Status Temuan PG': status_temuan,
                        'Keterangan': row.get('Keterangan', ''),
                        'Keterangan Lain': row.get('Keterangan Lain', '')
                    })
                    
        df_bodong = pd.DataFrame(data_bodong)
        if not df_bodong.empty:
            df_bodong.to_excel(writer, index=False, sheet_name='PG MB Tidak Ditemukan')
            format_excel_sheet(writer.sheets['PG MB Tidak Ditemukan'], df_bodong, 'Kode MB')
        else:
            df_info = pd.DataFrame({'Info': ['Semua PIBC di MB terdaftar valid di Master PG sesuai gudangnya']})
            df_info.to_excel(writer, index=False, sheet_name='PG MB Tidak Ditemukan')
            format_excel_sheet(writer.sheets['PG MB Tidak Ditemukan'], df_info, None)
        
        if df_pg is not None and not df_pg.empty and 'PIBC_PG' in df_pg.columns:
            df_belum = _dapatkan_df_belum_diproses_pg(df_hasil_mb, df_pg)
            if not df_belum.empty:
                df_belum.to_excel(writer, index=False, sheet_name='Master PG Belum Diproses')
                format_excel_sheet(writer.sheets['Master PG Belum Diproses'], df_belum, None)
            else:
                df_info = pd.DataFrame({'Info': ['Semua data Master PG (Mesin Dyeing & Lab & RnD) sudah diproses di MB']})
                df_info.to_excel(writer, index=False, sheet_name='Master PG Belum Diproses')
                format_excel_sheet(writer.sheets['Master PG Belum Diproses'], df_info, None)
        else:
            df_info = pd.DataFrame({'Info': ['Master PG kosong/belum diupload']})
            df_info.to_excel(writer, index=False, sheet_name='Master PG Belum Diproses')
            format_excel_sheet(writer.sheets['Master PG Belum Diproses'], df_info, None)
            
    return output.getvalue()

def buat_zip_semua_laporan(dict_files):
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
        for nama_file, bytes_data in dict_files.items():
            zf.writestr(nama_file, bytes_data)
    return output.getvalue()
