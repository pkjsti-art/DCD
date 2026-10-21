import pandas as pd
import re
import io

# Daftar baris total/overhead/pekerjaan yang wajib dihiraukan
BARIS_DIABAIKAN = [
    "Total Material (Details)", "Total Pekerjaan (Works)", 
    "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor",
    "Kode Pekerjaan", "Nama Pekerjaan"
]

def perbaiki_header_pg(df):
    """Fungsi khusus memotong baris sampah di Master PG agar Header naik ke atas"""
    if df.empty: return df
    for i in range(min(15, len(df))):
        baris = [str(x).strip().upper() for x in df.iloc[i].values]
        if 'MASUK' in baris and 'GUDANG' in baris:
            df.columns = baris
            return df.iloc[i+1:].reset_index(drop=True)
    return df

def perbaiki_header_pibc(df):
    """Fungsi khusus memotong baris sampah di Master PIBC"""
    if df.empty: return df
    for i in range(min(15, len(df))):
        baris_str = " ".join([str(x).upper() for x in df.iloc[i].values])
        if ('DYELOT' in baris_str or 'PIBC' in baris_str) and ('ACTUAL' in baris_str or 'QTY' in baris_str):
            df.columns = df.iloc[i].values
            return df.iloc[i+1:].reset_index(drop=True)
    return df

def cari_kolom_akurat(df, daftar_kemungkinan):
    for kem in daftar_kemungkinan:
        for col in df.columns:
            col_clean = str(col).strip().lower()
            if kem.lower() in col_clean:
                return col
    return None

def bersihkan_angka_indo(val):
    if pd.isna(val) or val is None: return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('.', '').replace(',', '.')
    try: return float(val_str)
    except: return 0.0

def is_item_benang(kode_item, gudang=""):
    k_upper = str(kode_item).strip().upper()
    g_upper = str(gudang).strip().upper()
    if k_upper.startswith(('TWP', 'MWP', 'TBM', 'MBB')): return True
    if 'MESIN DYEING' in g_upper or 'LAB' in g_upper or 'RND' in g_upper: return False
    if k_upper.startswith('TBB'): return True
    return False

def ekstrak_kode_pibc(teks, gudang=""):
    if pd.isna(teks): return None
    teks_str = str(teks)
    if 'DYEING' in str(gudang).upper() and ('REPROCES' in teks_str.upper() or 'TO' in teks_str.upper()):
        match_kurung = re.search(r'\((PIBC-[A-Za-z0-9#\-\.]+)\)', teks_str, re.IGNORECASE)
        if match_kurung: return match_kurung.group(1).strip().rstrip('.')
        match_to = re.search(r'TO\s+([A-Za-z0-9\-\#]+)', teks_str, re.IGNORECASE)
        if match_to: return match_to.group(1).strip().rstrip('.')
    pola = r'(PIBC-[A-Za-z0-9#\-\.]+)'
    hasil = re.search(pola, teks_str, re.IGNORECASE)
    if hasil: return hasil.group(1).strip().rstrip('.')
    return None

def siapkan_kolom_standar(df_mb):
    col_target_qty = cari_kolom_akurat(df_mb, ['target qty', 'target_qty'])
    col_target_unit = cari_kolom_akurat(df_mb, ['target unit', 'target_unit'])
    col_qty = cari_kolom_akurat(df_mb, ['qty', 'quantity'])
    col_unit = cari_kolom_akurat(df_mb, ['unit', 'satuan'])
    col_kode_item = cari_kolom_akurat(df_mb, ['kode item', 'kode_item'])

    seri_target_gr = pd.Series([''] * len(df_mb), dtype=object)
    seri_bb_gr = pd.Series([''] * len(df_mb), dtype=object)

    for idx, r in df_mb.iterrows():
        kode_item = str(r[col_kode_item]).strip() if col_kode_item and pd.notna(r[col_kode_item]) else ''
        gudang = str(r.get('Gudang', '')).strip().upper()
        is_benang = is_item_benang(kode_item, gudang)
        is_diabaikan = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper() or ('SOFTCONE' in gudang and not is_benang))

        if not is_diabaikan:
            if col_target_qty and pd.notna(r[col_target_qty]) and str(r[col_target_qty]).strip().lower() != 'nan':
                val = bersihkan_angka_indo(r[col_target_qty])
                unit = str(r[col_target_unit]).strip().upper() if col_target_unit else 'GR'
                seri_target_gr[idx] = val if unit == 'GR' else val * 1000.0

            if col_qty and pd.notna(r[col_qty]) and str(r[col_qty]).strip().lower() != 'nan':
                val = bersihkan_angka_indo(r[col_qty])
                unit = str(r[col_unit]).strip().upper() if col_unit else 'GR'
                seri_bb_gr[idx] = val if unit == 'GR' else val * 1000.0

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

    # -------------------------------------------------------------
    # 1. SETUP LOOKUP MASTER PIBC
    # -------------------------------------------------------------
    pibc_lookup = {}
    if df_pibc is not None and not df_pibc.empty:
        df_pibc = perbaiki_header_pibc(df_pibc.copy()) 
        col_dyelot = cari_kolom_akurat(df_pibc, ['dyelot', 'pibc'])
        col_act = cari_kolom_akurat(df_pibc, ['actual', 'qty', 'jumlah'])
        col_kode_obat = cari_kolom_akurat(df_pibc, ['kode obat', 'kode'])
        col_nama_obat = cari_kolom_akurat(df_pibc, ['obat', 'nama', 'item'])
        
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

    # -------------------------------------------------------------
    # 2. SETUP LOOKUP MASTER PG (STRICT KONSEP GUDANG & KETERANGAN DET)
    # -------------------------------------------------------------
    pg_lookup = {}
    if df_pg is not None and not df_pg.empty:
        df_pg = perbaiki_header_pg(df_pg.copy()) 
        
        # Cari kolom secara eksplisit sesuai aturan
        col_gudang = cari_kolom_akurat(df_pg, ['GUDANG'])
        col_masuk = cari_kolom_akurat(df_pg, ['MASUK'])
        col_ket_det = cari_kolom_akurat(df_pg, ['KETERANGAN DET', 'KETERANGAN_DET'])
        col_kode_var = cari_kolom_akurat(df_pg, ['KODE VARIANT', 'KODE ITEM'])
        col_nama_var = cari_kolom_akurat(df_pg, ['NAMA VARIANT', 'NAMA ITEM'])
        
        if col_gudang and col_ket_det and col_masuk:
            for _, r in df_pg.iterrows():
                val_msk = bersihkan_angka_indo(r[col_masuk])
                if val_msk <= 0: continue # Aturan: MASUK tidak boleh 0
                
                gudang_pg = str(r[col_gudang]).strip().upper()
                
                # Identifikasi Kategori Gudang di PG
                kat_gudang_pg = ""
                if "LAB" in gudang_pg or "RND" in gudang_pg: kat_gudang_pg = "LAB"
                elif "MESIN DYEING" in gudang_pg: kat_gudang_pg = "MESIN DYEING"
                
                if not kat_gudang_pg: continue # Skip jika bukan gudang yang kita proses
                
                # Ekstrak PIBC mutlak dari KETERANGAN DET.
                raw_pibc = str(r[col_ket_det])
                if not raw_pibc or raw_pibc.lower() == 'nan': continue
                
                pibc_ext = ekstrak_kode_pibc(raw_pibc)
                k_ket = pibc_ext.upper() if pibc_ext else raw_pibc.strip().upper()
                k_ket_clean = k_ket.replace('PIBC-', '')
                
                k_kode = str(r[col_kode_var]).strip().upper() if col_kode_var and pd.notna(r[col_kode_var]) else ""
                k_nama = str(r[col_nama_var]).strip().upper() if col_nama_var and pd.notna(r[col_nama_var]) else ""
                
                # Simpan lookup dengan format tuple 3-kunci: (Kategori Gudang, PIBC, Item)
                if k_kode:
                    pg_lookup[(kat_gudang_pg, k_ket, k_kode)] = pg_lookup.get((kat_gudang_pg, k_ket, k_kode), 0.0) + val_msk
                    pg_lookup[(kat_gudang_pg, k_ket_clean, k_kode)] = pg_lookup.get((kat_gudang_pg, k_ket_clean, k_kode), 0.0) + val_msk
                if k_nama:
                    pg_lookup[(kat_gudang_pg, k_ket, k_nama)] = pg_lookup.get((kat_gudang_pg, k_ket, k_nama), 0.0) + val_msk
                    pg_lookup[(kat_gudang_pg, k_ket_clean, k_nama)] = pg_lookup.get((kat_gudang_pg, k_ket_clean, k_nama), 0.0) + val_msk

    grup_kode = df_mb.groupby('Kode', sort=False) if 'Kode' in df_mb.columns else [(None, df_mb)]

    # -------------------------------------------------------------
    # PRE-PROCESSING: DETEKSI CELUP GABUNG OTOMATIS
    # -------------------------------------------------------------
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
            if kp_upper not in pibc_to_mb: pibc_to_mb[kp_upper] = []
            if k_val not in pibc_to_mb[kp_upper]: pibc_to_mb[kp_upper].append(k_val)

            for idx, r in df_sub.iterrows():
                k_item = str(r.get('Kode Item', '')).strip().upper()
                gudang = str(r.get('Gudang', '')).strip().upper()
                if not is_item_benang(k_item, gudang) and k_item not in BARIS_DIABAIKAN and 'PEKERJAAN' not in k_item:
                    val_bb = bersihkan_angka_indo(r.get('Qty BB Standar (GR)'))
                    gabungan_qty_item[(kp_upper, k_item)] = gabungan_qty_item.get((kp_upper, k_item), 0.0) + val_bb

    # -------------------------------------------------------------
    # PROSES CROSSCHECK UTAMA MB
    # -------------------------------------------------------------
    for k_val, df_sub in grup_kode:
        grup_punya_pibc = False
        kode_pibc_grup = None
        target_gr_utama_grup = 0.0
        total_bb_benang_grup = 0.0
        ada_item_non_tbb_sti = False
        rincian_benang_list = []
        
        # Identifikasi Kategori Gudang dominan di kelompok MB ini
        gudang_group = ""
        for i in df_sub.index:
            g = str(df_mb.loc[i].get('Gudang', '')).strip().upper()
            if g != "":
                gudang_group = g
                break
                
        kat_gudang_mb = ""
        if "LAB" in gudang_group or "RND" in gudang_group: kat_gudang_mb = "LAB"
        elif "MESIN DYEING" in gudang_group: kat_gudang_mb = "MESIN DYEING"
        
        # 1. Scan Awal Kelompok
        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            gudang = str(row.get('Gudang', '')).strip().upper()
            teks_gab = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
            
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
                if isinstance(val_bb, (int, float)): total_bb_benang_grup += float(val_bb)
                if kode_item not in rincian_benang_list: rincian_benang_list.append(kode_item)

            if 'DYEING STI' in gudang:
                if not (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper()) and kode_item and not kode_item.startswith('TBB'):
                    ada_item_non_tbb_sti = True

        rincian_benang_str = ", ".join(rincian_benang_list)
        is_celup_gabung = False
        pasangan_str = ""
        if kode_pibc_grup:
            semua_mb = pibc_to_mb.get(kode_pibc_grup.upper(), [])
            if len(semua_mb) > 1:
                is_celup_gabung = True
                pasangan_str = ", ".join([str(m) for m in semua_mb if m != k_val])

        info_cg_base = f"Celup Gabung dgn: {pasangan_str}" if is_celup_gabung else ""
        is_first_row = True
        benang_status_rendered = False
        status_cek_list = []

        for idx in df_sub.index:
            row = df_mb.loc[idx]
            kode_item = str(row.get('Kode Item', '')).strip()
            nama_item = str(row.get('Nama Item', '')).strip()
            gudang = str(row.get('Gudang', '')).strip().upper()
            benang_flag = is_item_benang(kode_item, gudang)

            is_diabaikan = (kode_item in BARIS_DIABAIKAN or 'PEKERJAAN' in kode_item.upper() or ('SOFTCONE' in gudang and not benang_flag))

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
                df_mb.at[idx, 'Status PIBC'] = 'KODE PIBC ADA' if grup_punya_pibc else 'KODE PIBC TIDAK ADA'
            df_mb.at[idx, 'Ext_PIBC'] = kode_pibc_grup

            df_mb.at[idx, 'Rincian Benang'] = 'BUKAN BENANG' if ('DYEING STI' in gudang and not kode_item.startswith('TBB')) else ('' if 'DYEING STI' in gudang else rincian_benang_str)

            nama_brg_jdi = str(row.get('Nama Barang Jadi', '')).upper()
            teks_gab = str(row.get('Keterangan', '')) + " " + str(row.get('Keterangan Lain', ''))
            ada_keyword_selisih = any(kw in nama_brg_jdi or kw in teks_gab.upper() for kw in keyword_selisih)

            # A. PENGECEKAN BENANG
            if benang_flag:
                qty_bb_baris = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                df_mb.at[idx, 'Selisih Benang (GR)'] = qty_bb_baris - target_gr_utama_grup

                if not benang_status_rendered:
                    selisih_total = total_bb_benang_grup - target_gr_utama_grup
                    if abs(selisih_total) <= 0.01: df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                    else: df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT' if ('LAB' in gudang or 'RND' in gudang) else ('MEMANG SELISIH' if ada_keyword_selisih else 'INCORRECT')
                    benang_status_rendered = True

            # B. PENGECEKAN OBAT (TRANSPARANSI TOTAL & STRICT MATCHING)
            else:
                qty_bb = bersihkan_angka_indo(row.get('Qty BB Standar (GR)'))
                qty_bb_cek = gabungan_qty_item.get((kode_pibc_grup.upper(), kode_item.upper()), qty_bb) if (is_celup_gabung and kode_pibc_grup) else qty_bb
                if is_celup_gabung and kode_pibc_grup: df_mb.at[idx, 'Info Celup Gabung'] = f"{info_cg_base} (Total: {qty_bb_cek} GR)"

                if kode_pibc_grup:
                    qty_pibc = None
                    qty_pg = None
                    
                    # Cek PIBC
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
                            
                    # Cek PG (Sesuai Konsep: Sertakan Identitas Gudang-nya)
                    if kat_gudang_mb: 
                        pg_key_candidates = [
                            (kat_gudang_mb, kode_pibc_grup.upper(), kode_item.upper()),
                            (kat_gudang_mb, kode_pibc_grup.upper().replace('PIBC-', ''), kode_item.upper()),
                            (kat_gudang_mb, kode_pibc_grup.upper(), nama_item.upper()),
                            (kat_gudang_mb, kode_pibc_grup.upper().replace('PIBC-', ''), nama_item.upper())
                        ]
                        for k in pg_key_candidates:
                            if k[2] and k in pg_lookup:
                                qty_pg = pg_lookup[k]
                                break

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

                    # PENENTUAN STATUS
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

            if is_first_row:
                kode_bj = str(row.get('Kode Barang Jadi', '')).strip()
                if 'MESIN DYEING' in gudang:
                    df_mb.at[idx, 'Cek Variant'] = kode_bj if not kode_item.startswith(('TWP.050', 'TWP.100', 'TWP.051')) and kode_bj not in ['TWP.100', 'TWP.400', 'TWP.051', 'TWP.666', 'TWP.901'] else 'CLEAN'
                elif 'LAB' in gudang:
                    df_mb.at[idx, 'Cek Variant'] = kode_bj if not kode_item.startswith(('TWP.050', 'TWP.051', 'TWP.999', 'TWP.100')) and kode_bj not in ['TWP.999', 'TWP.100', 'TWP.400', 'TWP.666', 'TWP.051'] else 'CLEAN'

            is_first_row = False

        # STATUS CEK
        idx_first = df_sub.index[0]
        is_softcone_group = 'SOFTCONE' in gudang_group or (gudang_group == '' and str(df_mb.loc[idx_first].get('Kode Item', '')).strip().upper().startswith(('TWP', 'MWP', 'TBM', 'MBB')))

        if ('DYEING STI' in gudang_group and not ada_item_non_tbb_sti) or is_softcone_group:
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
            df_pibc = perbaiki_header_pibc(df_pibc.copy())
            col_dyelot = cari_kolom_akurat(df_pibc, ['dyelot', 'pibc'])
            pibc_di_mb = set(df_hasil_mb['Ext_PIBC'].dropna().astype(str).str.strip().unique()) if 'Ext_PIBC' in df_hasil_mb.columns else set()
            if col_dyelot:
                df_pibc[~df_pibc[col_dyelot].astype(str).str.strip().isin(pibc_di_mb)].to_excel(writer, index=False, sheet_name='PIBC Tidak Ada di MB')
            else: df_pibc.to_excel(writer, index=False, sheet_name='PIBC Tidak Ada di MB')
        else: pd.DataFrame({'Info': ['Master PIBC kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()

def buat_excel_pg_tidak_di_mb(df_hasil_mb, df_pg):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if df_pg is not None and not df_pg.empty:
            df_pg = perbaiki_header_pg(df_pg.copy())
            col_ket = cari_kolom_akurat(df_pg, ['keterangan det', 'keterangan_det'])
            pibc_di_mb = set(df_hasil_mb['Ext_PIBC'].dropna().astype(str).str.strip().unique()) if 'Ext_PIBC' in df_hasil_mb.columns else set()
            if col_ket:
                df_pg[~df_pg[col_ket].astype(str).str.strip().isin(pibc_di_mb)].to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
            else: df_pg.to_excel(writer, index=False, sheet_name='PG Tidak Ada di MB')
        else: pd.DataFrame({'Info': ['Master PG kosong/belum diupload']}).to_excel(writer, index=False, sheet_name='Kosong')
    return output.getvalue()
