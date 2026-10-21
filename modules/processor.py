import pandas as pd
import re
import io

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

def is_item_benang(kode_item, gudang=""):
    """
    Aturan Presisi Benang per Gudang:
    - Mesin Dyeing & LAB/RND: Benang HANYA TWP, MWP, TBM, MBB (TBB adalah Chemical).
    - Softcone / Lainnya: Benang mencakup TWP, MWP, TBM, MBB, dan TBB.
    """
    k_upper = str(kode_item).strip().upper()
    g_upper = str(gudang).strip().upper()
    
    prefix_utama = ('TWP', 'MWP', 'TBM', 'MBB')
    
    if k_upper.startswith(prefix_utama):
        return True
        
    if 'MESIN DYEING' in g_upper or 'LAB' in g_upper or 'RND' in g_upper:
        return False
        
    if k_upper.startswith('TBB'):
        return True
        
    return False

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
        
        is_benang = is_item_benang(kode_item, gudang)
        is_diabaikan = (
            kode_item in BARIS_DIABAIKAN or
            'PEKERJAAN' in kode_item.upper() or
            ('SOFTCONE' in gudang and not is_benang)
        )

        if not is_diabaikan:
            if col_target_qty:
                val_raw = r[col_target_qty]
                if pd.notna(val_raw) and str(val_raw).strip() != '' and str(val_raw).strip().lower() != 'nan':
                    val = bersihkan_angka_indo(val_raw)
                    unit = str(r[col_target_unit]).strip().upper() if col_target_unit in df_mb.columns and pd.notna(r[col_target_unit]) else 'GR'
                    seri_target_gr[idx] = val if unit == 'GR' else val * 1000.0

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

    # TATA LETAK KOLOM BARU SESUAI KONSEP TRANSPARANSI TOTAL
    kolom_tambahan = [
        'Qty PIBC', 'Qty PG', 
        'Selisih Benang (GR)', 'Selisih PIBC (GR)', 'Selisih PG (GR)',
        'Crosscheck Qty', 'Info Celup Gabung', 'Status Cek', 
        'Status PIBC', 'Rincian Benang', 'Cek Variant', 'Ext_PIBC'
    ]
    
    for col_baru in kolom_tambahan:
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

    # LOOKUP MASTER PIBC (Double Matching)
    pibc_lookup = {}
    if df_pibc is not None and not df_pibc.empty:
        col_dyelot = cari_kolom_akurat(df_pibc, ['dyelot', 'pibc'])
        col_act = cari_kolom_akurat(df_pibc, ['actual', 'qty', 'jumlah'])
        col_kode_obat = cari_kolom_akurat(df_pibc, ['kode obat', 'kode', 'kode_obat', 'kode item'])
        col_nama_obat = cari_kolom_akurat(df_pibc, ['obat', 'nama', 'item', 'deskripsi'])
        
        if col_dyelot and col_act:
            for _, r in df_pibc.iterrows():
                k_dyelot = str(r[col_dyelot]).strip().upper()
                val_act = bersihkan_angka_indo(r[col_act])
                
                if col_kode_obat and pd.notna(r[col_kode_obat]):
                    k_kode = str(r[col_kode_obat]).strip().upper()
                    if k_kode:
                        pibc_lookup[(k_dyelot, k_kode)] = pibc_lookup.get((k_dyelot, k_kode), 0.0) + val_act
                        pibc_lookup[(k_dyelot.replace('PIBC-', ''), k_kode)] = pibc_lookup.get((k_dyelot.replace('PIBC-', ''), k_kode), 0.0) + val_act
                
                if col_nama_obat and pd.notna(r[col_nama_obat]):
                    k_nama = str(r[col_nama_obat]).strip().upper()
                    if k_nama:
                        pibc_lookup[(k_dyelot, k_nama)] = pibc_lookup.get((k_dyelot, k_nama), 0.0) + val_act
                        pibc_lookup[(k_dyelot.replace('PIBC-', ''), k_nama)] = pibc_lookup.get((k_dyelot.replace('PIBC-', ''), k_nama), 0.0) + val_act

    # LOOKUP MASTER PG (Double Matching)
    pg_lookup = {}
    if df_pg is not None and not df_pg.empty:
        col_ket = cari_kolom_akurat(df_pg, ['keterangan', 'det', 'pibc'])
        col_masuk = cari_kolom_akurat(df_pg, ['masuk', 'qty', 'jumlah'])
        col_kode_var = cari_kolom_akurat(df_pg, ['kode variant', 'kode', 'kode_variant', 'kode item'])
        col_nama_var = cari_kolom_akurat(df_pg, ['nama variant', 'variant', 'nama', 'item'])
        
        if col_ket and col_masuk:
            for _, r in df_pg.iterrows():
                k_ket = str(r[col_ket]).strip().upper()
                val_msk = bersihkan_angka_indo(r[col_masuk])
                
                if col_kode_var and pd.notna(r[col_kode_var]):
                    k_kode = str(r[col_kode_var]).strip().upper()
                    if k_kode:
                        pg_lookup[(k_ket, k_kode)] = pg_lookup.get((k_ket, k_kode), 0.0) + val_msk
                        
                if col_nama_var and pd.notna(r[col_nama_var]):
                    k_nama = str(r[col_nama_var]).strip().upper()
                    if k_nama:
                        pg_lookup[(k_ket, k_nama)] = pg_lookup.get((k_ket, k_nama), 0.0) + val_msk

    grup_kode = df_mb.groupby('Kode', sort=False) if 'Kode' in df_mb.columns else [(None, df_mb)]

    # =========================================================================
    # PRE-PROCESSING: DETEKSI DAN AKUMULASI OTOMATIS CASE "CELUP GABUNG"
    # =========================================================================
    pibc_to_mb = {}
    gabungan_qty_item = {}

    for k_val, df_sub in grup_kode:
        kp = None
        for idx, r in df_sub.iterrows():
            gudang = str(r.get('Gudang', '')).strip().upper()
            teks_gab = str(r.get('Keterangan', '')) + " " + str(r.get('Keterangan Lain', ''))
            kp = ekstrak_kode_pibc(teks_gab, gudang)
            if kp: break
            
        if kp:
            kp_upper = kp.upper()
            if kp_upper not in pibc_to_mb:
                pibc_to_mb[kp_upper] = []
            if k_val not in pibc_to_mb[kp_upper]:
                pibc_to_mb[kp_upper].append(k_val)

            for idx, r in df_sub.iterrows():
                k_item = str(r.get('Kode Item', '')).strip().upper()
                gudang = str(r.get('Gudang', '')).strip().upper()
                benang_flag = is_item_benang(k_item, gudang)
                if not benang_flag and k_item not in BARIS_DIABAIKAN and 'PEKERJAAN' not in k_item:
                    val_bb = bersihkan_angka_indo(r.get('Qty BB Standar (GR)'))
                    key_item = (kp_upper, k_item)
                    gabungan_qty_item[key_item] = gabungan_qty_item.get(key_item, 0.0) + val_bb
    # =========================================================================

    for k_val, df_sub in grup_kode:
        grup_punya_pibc = False
        kode_pibc_grup = None
        target_gr_utama_grup = 0.0
        total_bb_benang_grup = 0.0
        ada_item_non_tbb_sti = False
        
        # 1. Scan Awal Kelompok
        rincian_benang_list = []
        
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

            if is_item_benang(kode_item, gudang):
                val_bb = row.get('Qty BB Standar (GR)')
                if isinstance(val_bb, (int, float)):
                    total_bb_benang_grup += float(val_bb)
                
                if kode_item not in rincian_benang_list:
                    rincian_benang_list.append(kode_item)

            if 'DYEING STI' in gudang:
                is_diabaikan_check = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper())
                if not is_diabaikan_check and kode_item and not kode_item.startswith('TBB'):
                    ada_item_non_tbb_sti = True

        rincian_benang_str = ", ".join(rincian_benang_list)

        # 2. Penentuan Status Celup Gabung untuk Kelompok Ini
        is_celup_gabung = False
        pasangan_str = ""
        if kode_pibc_grup:
            semua_mb = pibc_to_mb.get(kode_pibc_grup.upper(), [])
            if len(semua_mb) > 1:
                is_celup_gabung = True
                pasangan = [str(m) for m in semua_mb if m != k_val]
                pasangan_str = ", ".join(pasangan)

        info_cg_base = f"Celup Gabung dgn: {pasangan_str}" if is_celup_gabung else ""

        status_pibc_grup = 'KODE PIBC ADA' if grup_punya_pibc else 'KODE PIBC TIDAK ADA'
        is_first_row = True
        benang_status_rendered = False
        status_cek_list = []

        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            nama_item = str(row.get('Nama Item', '')).strip()
            gudang = str(row.get('Gudang', '')).strip().upper()

            benang_flag = is_item_benang(kode_item, gudang)

            is_diabaikan = (
                kode_item in BARIS_DIABAIKAN or
                'PEKERJAAN' in kode_item.upper() or
                ('SOFTCONE' in gudang and not benang_flag)
            )

            # Kosongkan semua kolom kalkulasi di awal
            df_mb.at[idx, 'Crosscheck Qty'] = ''
            df_mb.at[idx, 'Qty PIBC'] = ''
            df_mb.at[idx, 'Qty PG'] = ''
            df_mb.at[idx, 'Selisih Benang (GR)'] = ''
            df_mb.at[idx, 'Selisih PIBC (GR)'] = ''
            df_mb.at[idx, 'Selisih PG (GR)'] = ''
            df_mb.at[idx, 'Status PIBC'] = ''
            df_mb.at[idx, 'Ext_PIBC'] = ''
            df_mb.at[idx, 'Info Celup Gabung'] = ''

            if is_diabaikan:
                df_mb.at[idx, 'Qty Target Standar (GR)'] = ''
                df_mb.at[idx, 'Qty BB Standar (GR)'] = ''
                df_mb.at[idx, 'Rincian Benang'] = rincian_benang_str if 'DYEING STI' not in gudang else ''
                continue

            if 'DYEING STI' in gudang and not ada_item_non_tbb_sti:
                df_mb.at[idx, 'Rincian Benang'] = ''
                df_mb.at[idx, 'Cek Variant'] = ''
                is_first_row = False
                continue

            if is_first_row:
                df_mb.at[idx, 'Status PIBC'] = status_pibc_grup
                
            df_mb.at[idx, 'Ext_PIBC'] = kode_pibc_grup

            if 'DYEING STI' in gudang:
                if not kode_item.startswith('TBB'):
                    df_mb.at[idx, 'Rincian Benang'] = 'BUKAN BENANG'
                else:
                    df_mb.at[idx, 'Rincian Benang'] = ''
            else:
                df_mb.at[idx, 'Rincian Benang'] = rincian_benang_str

            nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
            ket = str(row.get('Keterangan', ''))
            ket_lain = str(row.get('Keterangan Lain', ''))
            teks_gab = ket + " " + ket_lain
            
            ada_keyword_selisih = any(kw in nama_brg_jdi or kw in teks_gab.upper() for kw in keyword_selisih)

            # A. JALUR PENGECEKAN BENANG
            if benang_flag:
                val_bb = row.get('Qty BB Standar (GR)')
                qty_bb_baris = float(val_bb) if isinstance(val_bb, (int, float)) else 0.0
                
                selisih_baris = qty_bb_baris - target_gr_utama_grup
                df_mb.at[idx, 'Selisih Benang (GR)'] = selisih_baris

                if not benang_status_rendered:
                    selisih_total = total_bb_benang_grup - target_gr_utama_grup
                    
                    if abs(selisih_total) <= 0.01:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else:
                        if 'LAB' in gudang or 'RND' in gudang:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                        else:
                            if ada_keyword_selisih:
                                df_mb.at[idx, 'Crosscheck Qty'] = 'MEMANG SELISIH'
                            else:
                                df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                    benang_status_rendered = True

            # B. JALUR PENGECEKAN CHEMICAL & DYESTUFF (Transparansi Total)
            else:
                val_bb = row.get('Qty BB Standar (GR)')
                qty_bb = float(val_bb) if isinstance(val_bb, (int, float)) else 0.0
                
                if is_celup_gabung and kode_pibc_grup:
                    qty_bb_cek = gabungan_qty_item.get((kode_pibc_grup.upper(), kode_item.upper()), qty_bb)
                    df_mb.at[idx, 'Info Celup Gabung'] = f"{info_cg_base} (Total: {qty_bb_cek} GR)"
                else:
                    qty_bb_cek = qty_bb

                if kode_pibc_grup:
                    qty_pibc = None
                    qty_pg = None
                    
                    # 1. Tarik Data PIBC
                    pibc_key_candidates = [
                        (kode_pibc_grup.upper(), kode_item.upper()),
                        (kode_pibc_grup.upper().replace('PIBC-', ''), kode_item.upper()),
                        (kode_pibc_grup.upper(), nama_item.upper()),
                        (kode_pibc_grup.upper().replace('PIBC-', ''), nama_item.upper())
                    ]
                    for k in pibc_key_candidates:
                        if k[1] and k in pibc_lookup:
                            qty_pibc = pibc_lookup[k]
                            break
                            
                    # 2. Tarik Data PG
                    pg_key_candidates = [
                        (kode_pibc_grup.upper(), kode_item.upper()),
                        (kode_pibc_grup.upper(), nama_item.upper())
                    ]
                    for k in pg_key_candidates:
                        if k[1] and k in pg_lookup:
                            qty_pg = pg_lookup[k]
                            break

                    # 3. Hitung Selisih Paralel
                    sel_pibc = None
                    sel_pg = None
                    
                    if qty_pibc is not None:
                        df_mb.at[idx, 'Qty PIBC'] = qty_pibc
                        sel_pibc = qty_bb_cek - qty_pibc
                        df_mb.at[idx, 'Selisih PIBC (GR)'] = sel_pibc
                        
                    if qty_pg is not None:
                        df_mb.at[idx, 'Qty PG'] = qty_pg
                        sel_pg = qty_bb_cek - qty_pg
                        df_mb.at[idx, 'Selisih PG (GR)'] = sel_pg

                    # 4. Tentukan Status Crosscheck
                    if qty_pibc is None and qty_pg is None:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'
                        status_cek_list.append(f"Kelebihan [{kode_item}]")
                    else:
                        match_pibc = (qty_pibc is not None and abs(sel_pibc) <= 0.01)
                        match_pg = (qty_pg is not None and abs(sel_pg) <= 0.01)
                        
                        if match_pibc and match_pg:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'MATCH SEMUA'
                        elif match_pibc and not match_pg:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'MATCH PIBC (PG Beda)'
                        elif match_pg and not match_pibc:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'MATCH PG (PIBC Beda)'
                        else:
                            df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT KEDUANYA'

            # PENGECEKAN VARIANT DI BARIS PERTAMA
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

        # STATUS CEK (Kelengkapan Item)
        idx_first = df_sub.index[0]
        
        gudang_group = ""
        for i in df_sub.index:
            g = str(df_mb.loc[i].get('Gudang', '')).strip().upper()
            if g != "":
                gudang_group = g
                break
        
        is_softcone_group = 'SOFTCONE' in gudang_group or (gudang_group == '' and str(df_mb.loc[idx_first].get('Kode Item', '')).strip().upper().startswith(PREFIX_BENANG_VALID))

        if ('DYEING STI' in gudang_group and not ada_item_non_tbb_sti) or is_softcone_group:
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
        
        # Penyesuaian filter berdasarkan status output yang baru
        df_inc = df_hasil[df_hasil['Crosscheck Qty'].astype(str).str.contains('INCORRECT|BEDA|TIDAK DI TEMUKAN', case=False, na=False)]
        kode_grup_inc = df_inc['Kode'].unique() if 'Kode' in df_inc.columns else []
        df_full_inc = df_hasil[df_hasil['Kode'].isin(kode_grup_inc)] if len(kode_grup_inc) > 0 else df_inc
        df_full_inc.to_excel(writer, index=False, sheet_name='INCORRECT')

        df_selisih = df_hasil[df_hasil['Crosscheck Qty'] == 'MEMANG SELISIH']
        kode_grup_selisih = df_selisih['Kode'].unique() if 'Kode' in df_selisih.columns else []
        df_full_selisih = df_hasil[df_hasil['Kode'].isin(kode_grup_selisih)] if len(kode_grup_selisih) > 0 else df_selisih
        df_full_selisih.to_excel(writer, index=False, sheet_name='INCORRECT Memang Selisih')

        df_benang_inc = df_hasil[(df_hasil['Rincian Benang'] != '') & (df_hasil['Crosscheck Qty'].astype(str).str.contains('INCORRECT|BEDA', case=False, na=False))]
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
