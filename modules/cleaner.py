import pandas as pd
import re

KOLOM_LOG_DUPLIKAT = [
    'Sumber Master', 'Jenis Duplikat', 'Tindakan Sistem',
    'No. Transaksi / Kode', 'Tanggal', 'Gudang',
    'File Dihapus', 'File Dipakai', 'Keterangan'
]

BARIS_DIABAIKAN_MB = {
    "Total Material (Details)", "Total Pekerjaan (Works)",
    "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor",
    "Kode Pekerjaan", "Nama Pekerjaan"
}

def normalisasi_gudang(nama_gudang):
    g = str(nama_gudang).strip().lower()
    if 'lab' in g or 'rnd' in g: return 'Gudang Lab & RnD'
    if 'mesin dyeing' in g or ('dyeing' in g and 'sti' not in g): return 'Gudang Mesin Dyeing'
    if 'dyeing sti' in g or 'sti' in g: return 'Gudang Dyeing STI'
    if 'softcone' in g: return 'Gudang Softcone'
    return str(nama_gudang).strip().title()

def bersihkan_angka_indo(val):
    if pd.isna(val) or val is None: return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('.', '').replace(',', '.')
    try: return float(val_str)
    except: return 0.0

def cari_kolom_akurat(df, daftar_kemungkinan):
    for kem in daftar_kemungkinan:
        kem_clean = re.sub(r'[^A-Z0-9]', '', kem.upper())
        for col in df.columns:
            col_clean = re.sub(r'[^A-Z0-9]', '', str(col).upper())
            if kem_clean == col_clean:
                return col
    for kem in daftar_kemungkinan:
        kem_clean = re.sub(r'[^A-Z0-9]', '', kem.upper())
        for col in df.columns:
            col_clean = re.sub(r'[^A-Z0-9]', '', str(col).upper())
            if kem_clean in col_clean:
                return col
    return None

def bersihkan_master_mb(file_excel):
    file_excel.seek(0)
    df_raw = pd.read_excel(file_excel, header=None)
    
    header_row_index = 0
    for idx, row in df_raw.iterrows():
        row_str_list = [str(val).strip() for val in row.values]
        if any(k in row_str_list for k in ['Kode', 'Kode Item', 'Target Qty', 'Kode Barang Jadi']):
            header_row_index = idx
            break
            
    file_excel.seek(0)
    df_mb = pd.read_excel(file_excel, header=header_row_index)
    df_mb.columns = [str(col).strip() for col in df_mb.columns]
    df_mb = df_mb.replace(r'^\s*$', pd.NA, regex=True)
    df_mb = df_mb.dropna(how='all').reset_index(drop=True)
    
    if len(df_mb) > 4:
        df_mb = df_mb.iloc[:-4].reset_index(drop=True)
        
    col_tgl = cari_kolom_akurat(df_mb, ['TANGGAL', 'TGL', 'DATE'])
    kolom_identitas = ['Gudang', 'Kode', 'Kode Barang Jadi', 'Nama Barang Jadi']
    if col_tgl and col_tgl not in kolom_identitas:
        kolom_identitas.append(col_tgl)

    for nama_kolom in kolom_identitas:
        if nama_kolom in df_mb.columns:
            df_mb[nama_kolom] = df_mb[nama_kolom].ffill()

    if 'Kode' in df_mb.columns:
        for col_ket in ['Keterangan', 'Keterangan Lain']:
            if col_ket in df_mb.columns:
                df_mb[col_ket] = df_mb[col_ket].replace(r'^\s*$', pd.NA, regex=True)
                df_mb[col_ket] = df_mb.groupby('Kode')[col_ket].ffill().fillna('')
            
    return df_mb

def _row_tuple_mb(r):
    ki = str(r.get('Kode Item', '')).strip().upper()
    ni = str(r.get('Nama Item', '')).strip().upper()
    q_bb = round(bersihkan_angka_indo(r.get('Qty BB Standar (GR)')), 3)
    q_tgt = round(bersihkan_angka_indo(r.get('Qty Target Standar (GR)')), 3)
    return (ki, ni, q_bb, q_tgt)

def _signature_dokumen_mb(df_sub, include_header=False):
    items = []
    for _, r in df_sub.iterrows():
        ki = str(r.get('Kode Item', '')).strip().upper()
        if ki in BARIS_DIABAIKAN_MB or 'PEKERJAAN' in ki:
            continue
        items.append(_row_tuple_mb(r))
    items_sorted = tuple(sorted(items))
    if not include_header:
        return items_sorted
    first = df_sub.iloc[0]
    g_norm = normalisasi_gudang(first.get('Gudang', ''))
    kbj = str(first.get('Kode Barang Jadi', '')).strip().upper()
    nbj = str(first.get('Nama Barang Jadi', '')).strip().upper()
    ket1 = re.sub(r'\s+', ' ', str(first.get('Keterangan', '')).strip().upper())
    ket2 = re.sub(r'\s+', ' ', str(first.get('Keterangan Lain', '')).strip().upper())
    return (g_norm, kbj, nbj, ket1, ket2, items_sorted)

def _bersihkan_blok_ganda_dalam_1_file_mb(df, fname, log_rows):
    """Skenario C: Hapus blok transaksi identik berulang dengan Kode & Gudang sama dalam 1 file."""
    if df is None or df.empty or 'Kode' not in df.columns:
        return df

    df_work = df.copy().reset_index(drop=True)
    kode_s = df_work['Kode'].fillna('').astype(str).str.strip()
    gudang_s = df_work['Gudang'].fillna('').astype(str).apply(normalisasi_gudang)
    key_series = list(zip(gudang_s, kode_s))

    # Deteksi blok berurutan (contiguous runs)
    runs = []
    start = 0
    for i in range(1, len(key_series) + 1):
        if i == len(key_series) or key_series[i] != key_series[start]:
            runs.append((key_series[start], start, i))
            start = i

    # C1: Kode & Gudang sama muncul terpisah di >1 blok dalam 1 file -> simpan blok terakhir
    last_run_pos = {}
    for r_i, (k_key, _, _) in enumerate(runs):
        if k_key[1] and k_key[1].lower() not in ['nan', 'none', 'nat']:
            last_run_pos[k_key] = r_i

    keep_indices = []
    for r_i, (k_key, s_idx, e_idx) in enumerate(runs):
        if k_key[1] and k_key[1].lower() not in ['nan', 'none', 'nat'] and last_run_pos.get(k_key) != r_i:
            first_r = df_work.iloc[s_idx]
            log_rows.append({
                'Sumber Master': 'Master MB',
                'Jenis Duplikat': 'Duplikat Blok Sama Kode Dalam 1 File',
                'Tindakan Sistem': 'Dihapus Otomatis (Blok Berulang dalam 1 File)',
                'No. Transaksi / Kode': k_key[1],
                'Tanggal': first_r.get('Tanggal', ''),
                'Gudang': first_r.get('Gudang', ''),
                'File Dihapus': fname,
                'File Dipakai': fname,
                'Keterangan': 'Kode MB yang sama muncul lebih dari 1 blok terpisah dalam file yang sama'
            })
            continue

        # C2: Dalam 1 blok kontigu, cek apakah seluruh isi baris terduplikasi persis k kali (k >= 2)
        sub_df = df_work.iloc[s_idx:e_idx]
        n_rows = len(sub_df)
        dipotong = False
        if k_key[1] and n_rows >= 2:
            tuples_all = [_row_tuple_mb(sub_df.iloc[j]) for j in range(n_rows)]
            for k_rep in range(2, n_rows + 1):
                if n_rows % k_rep != 0:
                    continue
                chunk_len = n_rows // k_rep
                first_chunk = tuples_all[:chunk_len]
                # Pastikan chunk punya minimal 1 baris item valid
                has_valid = any(t[0] and t[0] not in BARIS_DIABAIKAN_MB and 'PEKERJAAN' not in t[0] for t in first_chunk)
                if not has_valid:
                    continue
                if all(tuples_all[c * chunk_len:(c + 1) * chunk_len] == first_chunk for c in range(1, k_rep)):
                    keep_indices.extend(range(s_idx, s_idx + chunk_len))
                    first_r = df_work.iloc[s_idx]
                    log_rows.append({
                        'Sumber Master': 'Master MB',
                        'Jenis Duplikat': 'Duplikat Blok Sama Kode Dalam 1 File',
                        'Tindakan Sistem': f'Dihapus Otomatis (Blok Tercetak {k_rep}x Identik)',
                        'No. Transaksi / Kode': k_key[1],
                        'Tanggal': first_r.get('Tanggal', ''),
                        'Gudang': first_r.get('Gudang', ''),
                        'File Dihapus': fname,
                        'File Dipakai': fname,
                        'Keterangan': f'Seluruh {chunk_len} baris transaksi tercetak berulang {k_rep}x identik dalam 1 file'
                    })
                    dipotong = True
                    break

        if not dipotong:
            keep_indices.extend(range(s_idx, e_idx))

    return df_work.iloc[keep_indices].reset_index(drop=True)

def deduplikasi_lintas_file_mb(list_tuple_mb):
    """
    Menerima list of (nama_file, df_mb_siap).
    - Skenario C: Sama Kode MB tercetak >1 blok identik dalam 1 file -> otomatis dibersihkan & dicatat ke log.
    - Skenario B: Sama (Gudang, Kode MB) lintas file -> otomatis dibersihkan (versi file terakhir) & dicatat 1 baris ke log.
    - Skenario A: Beda Kode MB di gudang sama tapi seluruh isi kelompok identik -> TIDAK dihapus, ditandai di map & log.
    """
    if not list_tuple_mb:
        return pd.DataFrame(), pd.DataFrame(columns=KOLOM_LOG_DUPLIKAT), {}

    log_rows = []
    cleaned_tuples = []

    # 1. Jalankan Skenario C per file
    for fname, df in list_tuple_mb:
        df_c = _bersihkan_blok_ganda_dalam_1_file_mb(df, fname, log_rows)
        cleaned_tuples.append((fname, df_c))

    # 2. Scan Skenario B: duplikat (Gudang_Norm, Kode) lintas file (keep='last')
    last_occ = {}
    signatures = {}
    for idx_f, (fname, df) in enumerate(cleaned_tuples):
        if df is None or df.empty or 'Kode' not in df.columns:
            continue
        df_tmp = df.copy()
        df_tmp['_k'] = df_tmp['Kode'].fillna('').astype(str).str.strip()
        df_tmp['_g'] = df_tmp['Gudang'].fillna('').astype(str).apply(normalisasi_gudang)
        for (g_norm, k_val), df_sub in df_tmp.groupby(['_g', '_k'], sort=False):
            if not k_val or k_val.lower() in ['nan', 'none', 'nat']:
                continue
            doc_key = (g_norm, k_val)
            last_occ[doc_key] = (idx_f, fname)
            signatures[(idx_f, doc_key)] = _signature_dokumen_mb(df_sub, include_header=False)

    df_final_list = []
    for idx_f, (fname, df) in enumerate(cleaned_tuples):
        if df is None or df.empty:
            continue
        if 'Kode' not in df.columns:
            df_final_list.append(df)
            continue

        keys_row = list(zip(
            df['Gudang'].fillna('').astype(str).apply(normalisasi_gudang),
            df['Kode'].fillna('').astype(str).str.strip()
        ))
        mask_buang = [
            (k[1] != '' and k[1].lower() not in ['nan', 'none', 'nat'] and k in last_occ and last_occ[k][0] != idx_f)
            for k in keys_row
        ]

        if any(mask_buang):
            df_buang = df[mask_buang]
            sudah_dicatat = set()
            for _, r in df_buang.iterrows():
                g_norm = normalisasi_gudang(r.get('Gudang', ''))
                k_val = str(r.get('Kode', '')).strip()
                doc_key = (g_norm, k_val)
                if doc_key in sudah_dicatat:
                    continue
                sudah_dicatat.add(doc_key)
                kept_idx, kept_fname = last_occ[doc_key]
                sama = (signatures.get((idx_f, doc_key)) == signatures.get((kept_idx, doc_key)))
                jenis_dup = (
                    'Duplikat Lintas File (Identik)'
                    if sama else
                    'Duplikat Lintas File (Isi Berbeda - Indikasi Revisi ERP)'
                )
                ket_txt = (
                    'Seluruh item & qty identik antar file (overlap upload)'
                    if sama else
                    'Kode transaksi sama tetapi rincian item/qty berbeda antar file (cek revisi ERP)'
                )
                log_rows.append({
                    'Sumber Master': 'Master MB',
                    'Jenis Duplikat': jenis_dup,
                    'Tindakan Sistem': 'Dihapus Otomatis (Pakai Versi File Terakhir)',
                    'No. Transaksi / Kode': k_val,
                    'Tanggal': r.get('Tanggal', ''),
                    'Gudang': r.get('Gudang', ''),
                    'File Dihapus': fname,
                    'File Dipakai': kept_fname,
                    'Keterangan': ket_txt
                })

            mask_keep = [not m for m in mask_buang]
            df_final_list.append(df[mask_keep].reset_index(drop=True))
        else:
            df_final_list.append(df)

    df_gabungan = pd.concat(df_final_list, ignore_index=True) if df_final_list else pd.DataFrame()

    # 3. Scan Skenario A: Beda Kode MB di gudang sama tapi seluruh isi kelompok identik (mengecualikan Kode & Tanggal)
    duplikat_mb_map = {}
    if not df_gabungan.empty and 'Kode' in df_gabungan.columns:
        df_scan = df_gabungan.copy()
        df_scan['_k'] = df_scan['Kode'].fillna('').astype(str).str.strip()
        df_scan['_g'] = df_scan['Gudang'].fillna('').astype(str).apply(normalisasi_gudang)

        sig_to_kodes = {}
        meta_per_doc = {}
        for (g_norm, k_val), df_sub in df_scan.groupby(['_g', '_k'], sort=False):
            if not k_val or k_val.lower() in ['nan', 'none', 'nat']:
                continue
            full_sig = _signature_dokumen_mb(df_sub, include_header=True)
            items_sig = full_sig[-1]
            if not items_sig:
                continue
            if full_sig not in sig_to_kodes:
                sig_to_kodes[full_sig] = []
            if k_val not in sig_to_kodes[full_sig]:
                sig_to_kodes[full_sig].append(k_val)
            first_r = df_sub.iloc[0]
            meta_per_doc[(g_norm, k_val)] = (first_r.get('Tanggal', ''), first_r.get('Gudang', ''))

        for full_sig, list_kodes in sig_to_kodes.items():
            if len(list_kodes) > 1:
                g_norm = full_sig[0]
                for k_val in list_kodes:
                    pasangan = [x for x in list_kodes if x != k_val]
                    pasangan_str = ", ".join(pasangan)
                    duplikat_mb_map[(g_norm, k_val)] = pasangan_str
                    tgl_v, gdg_v = meta_per_doc.get((g_norm, k_val), ('', g_norm))
                    log_rows.append({
                        'Sumber Master': 'Master MB',
                        'Jenis Duplikat': 'Indikasi Duplikat MB Identik (Beda Kode)',
                        'Tindakan Sistem': 'TIDAK Dihapus (Cek Input Ganda / ERP)',
                        'No. Transaksi / Kode': k_val,
                        'Tanggal': tgl_v,
                        'Gudang': gdg_v,
                        'File Dihapus': '-',
                        'File Dipakai': '-',
                        'Keterangan': f'Seluruh isi kelompok transaksi identik dengan Kode MB: {pasangan_str}'
                    })

    df_log = pd.DataFrame(log_rows, columns=KOLOM_LOG_DUPLIKAT)
    return df_gabungan, df_log, duplikat_mb_map

def bersihkan_master_pibc(file_excel):
    file_excel.seek(0)
    df_pibc = pd.read_excel(file_excel)
    if len(df_pibc) > 1:
        df_pibc = df_pibc.iloc[:-1].reset_index(drop=True)
    df_pibc.columns = [str(col).strip() for col in df_pibc.columns]
    df_pibc = df_pibc.replace(r'^\s*$', pd.NA, regex=True).dropna(how='all').reset_index(drop=True)
    return df_pibc

def _signature_dokumen_pg(df_sub):
    sig = []
    for _, r in df_sub.iterrows():
        sig.append((
            str(r.get('PIBC_PG', '')).strip().upper(),
            str(r.get('Kode_Item_PG', '')).strip().upper(),
            str(r.get('Nama_Item_PG', '')).strip().upper(),
            round(bersihkan_angka_indo(r.get('Qty_PG')), 3)
        ))
    return sorted(sig)

def bersihkan_master_pg(daftar_file_pg, return_log=False):
    """
    Membersihkan dan menggabungkan Master PG tanpa forward fill.
    Deduplikasi otomatis lintas file murni berdasarkan (Gudang_Norm, No_Transaksi_PG)
    dengan mempertahankan versi file urutan terakhir, dan mencatat ringkas 1 baris per dokumen.
    """
    if not isinstance(daftar_file_pg, list):
        daftar_file_pg = [daftar_file_pg]
        
    kolom_output = [
        'No_Transaksi_PG', 'Tanggal_PG', 'Gudang_PG', 'Keterangan_PG',
        'PIBC_PG', 'Kode_Item_PG', 'Nama_Item_PG', 'Qty_PG', 'Status_PG'
    ]
    per_file_dfs = []
    
    for idx_f, file in enumerate(daftar_file_pg):
        fname = getattr(file, 'name', f'File_PG_{idx_f + 1}')
        try:
            file.seek(0)
            df = pd.read_excel(file, header=3)
        except Exception:
            continue
            
        if df.empty:
            continue
            
        col_no_trans = cari_kolom_akurat(df, ['NOTRANSAKSI', 'KODETRANSAKSI'])
        col_tanggal = cari_kolom_akurat(df, ['TANGGAL', 'TGL', 'DATE'])
        col_masuk = cari_kolom_akurat(df, ['MASUK'])
        col_ket_det = cari_kolom_akurat(df, ['KETERANGANDET'])
        col_kode_var = cari_kolom_akurat(df, ['KODEVARIANT', 'KODEITEM'])
        col_nama_var = cari_kolom_akurat(df, ['NAMAVARIANT', 'NAMAITEM'])
        col_gudang = cari_kolom_akurat(df, ['GUDANG'])
        col_status = cari_kolom_akurat(df, ['STATUS'])

        col_ket_utama = None
        for col in df.columns:
            if col == col_ket_det:
                continue
            col_clean = re.sub(r'[^A-Z0-9]', '', str(col).upper())
            if col_clean == 'KETERANGAN':
                col_ket_utama = col
                break
        if not col_ket_utama:
            for col in df.columns:
                if col == col_ket_det:
                    continue
                col_clean = re.sub(r'[^A-Z0-9]', '', str(col).upper())
                if 'KETERANGAN' in col_clean and 'DET' not in col_clean:
                    col_ket_utama = col
                    break
        
        if not col_masuk: 
            continue
        
        df['TEMP_MASUK'] = df[col_masuk].apply(bersihkan_angka_indo)
        df_bersih = df[df['TEMP_MASUK'] > 0].copy()
        
        rename_dict = {}
        if col_no_trans: rename_dict[col_no_trans] = 'No_Transaksi_PG'
        if col_tanggal: rename_dict[col_tanggal] = 'Tanggal_PG'
        if col_gudang: rename_dict[col_gudang] = 'Gudang_PG'
        if col_ket_utama: rename_dict[col_ket_utama] = 'Keterangan_PG'
        if col_ket_det: rename_dict[col_ket_det] = 'PIBC_PG'
        if col_kode_var: rename_dict[col_kode_var] = 'Kode_Item_PG'
        if col_nama_var: rename_dict[col_nama_var] = 'Nama_Item_PG'
        if col_status: rename_dict[col_status] = 'Status_PG'
        
        df_bersih.rename(columns=rename_dict, inplace=True)
        df_bersih['Qty_PG'] = df_bersih['TEMP_MASUK']
        
        for col_str in ['No_Transaksi_PG', 'Gudang_PG', 'Keterangan_PG', 'PIBC_PG', 'Kode_Item_PG', 'Nama_Item_PG', 'Status_PG']:
            if col_str in df_bersih.columns:
                df_bersih[col_str] = df_bersih[col_str].fillna('').astype(str).str.strip()
                df_bersih.loc[df_bersih[col_str].str.lower().isin(['nan', 'none', 'nat']), col_str] = ''

        for kf in kolom_output:
            if kf not in df_bersih.columns:
                df_bersih[kf] = ""

        if col_status:
            df_bersih = df_bersih[
                df_bersih['Status_PG'].astype(str).str.strip().str.lower() == 'active'
            ].reset_index(drop=True)

        if not df_bersih.empty:
            per_file_dfs.append((fname, df_bersih[kolom_output].copy()))

    log_rows = []
    last_occ = {}
    signatures = {}

    for idx_f, (fname, df_f) in enumerate(per_file_dfs):
        df_cek = df_f.copy()
        df_cek['_nt'] = df_cek['No_Transaksi_PG'].str.strip().str.upper()
        df_cek['_gd'] = df_cek['Gudang_PG'].apply(normalisasi_gudang)

        for (gd_k, nt_k), df_sub in df_cek.groupby(['_gd', '_nt'], sort=False):
            if not nt_k:
                continue
            doc_key = (gd_k, nt_k)
            last_occ[doc_key] = (idx_f, fname)
            signatures[(idx_f, doc_key)] = _signature_dokumen_pg(df_sub)

    all_df_dedup = []
    for idx_f, (fname, df_f) in enumerate(per_file_dfs):
        keys_series = list(zip(
            df_f['Gudang_PG'].apply(normalisasi_gudang),
            df_f['No_Transaksi_PG'].str.strip().str.upper()
        ))
        mask_buang = [
            (k[1] != '' and k in last_occ and last_occ[k][0] != idx_f)
            for k in keys_series
        ]

        if any(mask_buang):
            df_buang = df_f[mask_buang]
            sudah_dicatat = set()
            for _, r in df_buang.iterrows():
                doc_key = (normalisasi_gudang(r['Gudang_PG']), r['No_Transaksi_PG'].strip().upper())
                if doc_key in sudah_dicatat:
                    continue
                sudah_dicatat.add(doc_key)
                kept_idx, kept_fname = last_occ[doc_key]
                sama = (signatures.get((idx_f, doc_key)) == signatures.get((kept_idx, doc_key)))
                jenis_dup = (
                    'Duplikat Lintas File (Identik)'
                    if sama else
                    'Duplikat Lintas File (Isi Berbeda - Indikasi Revisi ERP)'
                )
                ket_txt = (
                    'Seluruh item & qty identik antar file (overlap upload)'
                    if sama else
                    'No. Transaksi sama tetapi rincian item/qty berbeda antar file (cek revisi ERP)'
                )
                log_rows.append({
                    'Sumber Master': 'Master PG',
                    'Jenis Duplikat': jenis_dup,
                    'Tindakan Sistem': 'Dihapus Otomatis (Pakai Versi File Terakhir)',
                    'No. Transaksi / Kode': r['No_Transaksi_PG'],
                    'Tanggal': r['Tanggal_PG'],
                    'Gudang': r['Gudang_PG'],
                    'File Dihapus': fname,
                    'File Dipakai': kept_fname,
                    'Keterangan': ket_txt
                })

            mask_keep = [not m for m in mask_buang]
            all_df_dedup.append(df_f[mask_keep].reset_index(drop=True))
        else:
            all_df_dedup.append(df_f)

    if all_df_dedup:
        df_gabungan = pd.concat(all_df_dedup, ignore_index=True)
    else:
        df_gabungan = pd.DataFrame(columns=kolom_output)

    df_log = pd.DataFrame(log_rows, columns=KOLOM_LOG_DUPLIKAT)
    if return_log:
        return df_gabungan, df_log
    return df_gabungan

def bersihkan_master_mesin(file_excel):
    if file_excel is None:
        return None
    file_excel.seek(0)
    dict_sheets = pd.read_excel(file_excel, sheet_name=None)
    
    daftar_df = []
    for _, df_sheet in dict_sheets.items():
        if df_sheet is None or df_sheet.empty:
            continue
        df_sheet.columns = [str(col).strip() for col in df_sheet.columns]
        df_sheet = df_sheet.replace(r'^\s*$', pd.NA, regex=True).dropna(how='all').reset_index(drop=True)
        
        col_mesin = cari_kolom_akurat(df_sheet, ['MESIN', 'NAMAMESIN', 'KODEMESIN'])
        col_jenis = cari_kolom_akurat(df_sheet, ['JENISMESIN', 'JENIS', 'KATEGORI'])
        
        if not col_mesin or not col_jenis:
            continue
            
        df_sub = df_sheet[[col_mesin, col_jenis]].copy()
        df_sub.columns = ['Mesin', 'JENIS MESIN']
        df_sub = df_sub.dropna(subset=['Mesin', 'JENIS MESIN']).reset_index(drop=True)
        df_sub['Mesin'] = df_sub['Mesin'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df_sub['JENIS MESIN'] = df_sub['JENIS MESIN'].astype(str).str.strip().str.upper()
        daftar_df.append(df_sub)
        
    if not daftar_df:
        return pd.DataFrame(columns=['Mesin', 'JENIS MESIN'])
        
    df_out = pd.concat(daftar_df, ignore_index=True)
    df_out['_key'] = df_out['Mesin'].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()))
    df_out = df_out.drop_duplicates(subset=['_key'], keep='last').drop(columns=['_key']).reset_index(drop=True)
    return df_out
