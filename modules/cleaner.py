import pandas as pd
import re

KOLOM_LOG_DUPLIKAT = [
    'Sumber Master', 'Jenis Duplikat', 'Tindakan Sistem', 'Nama File Asal',
    'No. Transaksi / Kode', 'Tanggal', 'Gudang', 'PIBC / Keterangan',
    'Kode Item', 'Nama Item', 'Qty (GR)'
]

BARIS_DIABAIKAN_MB = {
    "Total Material (Details)", "Total Pekerjaan (Works)",
    "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor",
    "Kode Pekerjaan", "Nama Pekerjaan"
}

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

def _signature_dokumen_mb(df_sub):
    sig = []
    for _, r in df_sub.iterrows():
        ki = str(r.get('Kode Item', '')).strip().upper()
        ni = str(r.get('Nama Item', '')).strip().upper()
        if ki in BARIS_DIABAIKAN_MB or 'PEKERJAAN' in ki:
            continue
        q_bb = round(bersihkan_angka_indo(r.get('Qty BB Standar (GR)')), 3)
        q_tgt = round(bersihkan_angka_indo(r.get('Qty Target Standar (GR)')), 3)
        sig.append((ki, ni, q_bb, q_tgt))
    return sorted(sig)

def deduplikasi_lintas_file_mb(list_tuple_mb):
    """
    Menerima list of (nama_file, df_mb_siap).
    - Duplikat Kode MB lintas file: otomatis dibersihkan (pakai versi file urutan terakhir) & dicatat ke log.
    - Baris identik dalam 1 file: TIDAK dihapus (tetap dihitung), hanya dicatat ke log sebagai peringatan.
    """
    if not list_tuple_mb:
        return pd.DataFrame(), pd.DataFrame(columns=KOLOM_LOG_DUPLIKAT)

    log_rows = []
    last_occ = {}
    signatures = {}

    for idx_f, (fname, df) in enumerate(list_tuple_mb):
        if df is None or df.empty or 'Kode' not in df.columns:
            continue

        # 1. Cek baris identik dalam 1 file (tidak dihapus)
        df_cek = df.copy()
        df_cek['_kode_c'] = df_cek['Kode'].fillna('').astype(str).str.strip()
        df_cek['_ki_c'] = df_cek['Kode Item'].fillna('').astype(str).str.strip() if 'Kode Item' in df_cek.columns else ''
        df_cek['_ni_c'] = df_cek['Nama Item'].fillna('').astype(str).str.strip() if 'Nama Item' in df_cek.columns else ''
        df_cek['_g_c'] = df_cek['Gudang'].fillna('').astype(str).str.strip().str.upper() if 'Gudang' in df_cek.columns else ''
        df_cek['_q_c'] = df_cek['Qty BB Standar (GR)'].apply(lambda x: round(bersihkan_angka_indo(x), 3)) if 'Qty BB Standar (GR)' in df_cek.columns else 0.0

        valid_mask = (
            (df_cek['_kode_c'] != '') &
            (~df_cek['_kode_c'].str.lower().isin(['nan', 'none', 'nat'])) &
            (df_cek['_ki_c'] != '') &
            (~df_cek['_ki_c'].isin(BARIS_DIABAIKAN_MB)) &
            (~df_cek['_ki_c'].str.upper().str.contains('PEKERJAAN', na=False)) &
            (df_cek['_q_c'] > 0)
        )
        df_valid = df_cek[valid_mask]
        if not df_valid.empty:
            dup_in_mask = df_valid.duplicated(subset=['_kode_c', '_g_c', '_ki_c', '_ni_c', '_q_c'], keep='first')
            for _, r in df_valid[dup_in_mask].iterrows():
                log_rows.append({
                    'Sumber Master': 'Master MB',
                    'Jenis Duplikat': 'Duplikat Dalam 1 File (Baris Identik)',
                    'Tindakan Sistem': 'TIDAK Dihapus (Tetap Dihitung - Cek ERP/Input Ganda)',
                    'Nama File Asal': fname,
                    'No. Transaksi / Kode': r.get('Kode', ''),
                    'Tanggal': r.get('Tanggal', ''),
                    'Gudang': r.get('Gudang', ''),
                    'PIBC / Keterangan': r.get('Keterangan', ''),
                    'Kode Item': r.get('Kode Item', ''),
                    'Nama Item': r.get('Nama Item', ''),
                    'Qty (GR)': r['_q_c']
                })

        # 2. Catat posisi terakhir & signature per Kode MB untuk dedup lintas file
        for k_val, df_sub in df_cek.groupby('_kode_c', sort=False):
            if not k_val or k_val.lower() in ['nan', 'none', 'nat']:
                continue
            last_occ[k_val] = (idx_f, fname)
            signatures[(idx_f, k_val)] = _signature_dokumen_mb(df_sub)

    # 3. Buang salinan Kode MB di file yang lebih awal (keep='last') dan catat ke log
    df_final_list = []
    for idx_f, (fname, df) in enumerate(list_tuple_mb):
        if df is None or df.empty:
            continue
        if 'Kode' not in df.columns:
            df_final_list.append(df)
            continue

        kode_series = df['Kode'].fillna('').astype(str).str.strip()
        kode_buang = {
            k for k in kode_series.unique()
            if k in last_occ and last_occ[k][0] != idx_f
        }

        if kode_buang:
            df_buang = df[kode_series.isin(kode_buang)]
            for _, r in df_buang.iterrows():
                k_val = str(r.get('Kode', '')).strip()
                ki = str(r.get('Kode Item', '')).strip()
                if ki in BARIS_DIABAIKAN_MB or 'PEKERJAAN' in ki.upper():
                    continue
                kept_idx, kept_fname = last_occ[k_val]
                sama = (signatures.get((idx_f, k_val)) == signatures.get((kept_idx, k_val)))
                jenis_dup = (
                    'Duplikat Lintas File (Identik)'
                    if sama else
                    'Duplikat Lintas File (Isi Berbeda - Indikasi Revisi ERP)'
                )
                q_val = bersihkan_angka_indo(r.get('Qty BB Standar (GR)'))
                log_rows.append({
                    'Sumber Master': 'Master MB',
                    'Jenis Duplikat': jenis_dup,
                    'Tindakan Sistem': f'Dihapus Otomatis (Pakai versi file: {kept_fname})',
                    'Nama File Asal': fname,
                    'No. Transaksi / Kode': k_val,
                    'Tanggal': r.get('Tanggal', ''),
                    'Gudang': r.get('Gudang', ''),
                    'PIBC / Keterangan': r.get('Keterangan', ''),
                    'Kode Item': ki,
                    'Nama Item': r.get('Nama Item', ''),
                    'Qty (GR)': q_val if q_val > 0 else None
                })

            df_keep = df[~kode_series.isin(kode_buang)].reset_index(drop=True)
            df_final_list.append(df_keep)
        else:
            df_final_list.append(df)

    df_gabungan = pd.concat(df_final_list, ignore_index=True) if df_final_list else pd.DataFrame()
    df_log = pd.DataFrame(log_rows, columns=KOLOM_LOG_DUPLIKAT)
    return df_gabungan, df_log

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

        # Filter Status == 'active' per file sebelum cek duplikat
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
        # 1. Cek baris identik dalam 1 file (tidak dihapus)
        df_cek = df_f.copy()
        df_cek['_nt'] = df_cek['No_Transaksi_PG'].str.strip().str.upper()
        df_cek['_gd'] = df_cek['Gudang_PG'].str.strip().str.upper()
        df_cek['_pb'] = df_cek['PIBC_PG'].str.strip().str.upper()
        df_cek['_ki'] = df_cek['Kode_Item_PG'].str.strip().str.upper()
        df_cek['_ni'] = df_cek['Nama_Item_PG'].str.strip().str.upper()
        df_cek['_qt'] = df_cek['Qty_PG'].apply(lambda x: round(bersihkan_angka_indo(x), 3))

        valid_in = df_cek[df_cek['_nt'] != '']
        if not valid_in.empty:
            dup_in_mask = valid_in.duplicated(subset=['_nt', '_gd', '_pb', '_ki', '_ni', '_qt'], keep='first')
            for _, r in valid_in[dup_in_mask].iterrows():
                log_rows.append({
                    'Sumber Master': 'Master PG',
                    'Jenis Duplikat': 'Duplikat Dalam 1 File (Baris Identik)',
                    'Tindakan Sistem': 'TIDAK Dihapus (Tetap Dihitung - Cek ERP/Input Ganda)',
                    'Nama File Asal': fname,
                    'No. Transaksi / Kode': r['No_Transaksi_PG'],
                    'Tanggal': r['Tanggal_PG'],
                    'Gudang': r['Gudang_PG'],
                    'PIBC / Keterangan': r['PIBC_PG'] or r['Keterangan_PG'],
                    'Kode Item': r['Kode_Item_PG'],
                    'Nama Item': r['Nama_Item_PG'],
                    'Qty (GR)': r['_qt']
                })

        # 2. Catat posisi terakhir & signature per (No_Transaksi_PG, Gudang_PG)
        for (nt_k, gd_k), df_sub in df_cek.groupby(['_nt', '_gd'], sort=False):
            if not nt_k:
                continue
            doc_key = (nt_k, gd_k)
            last_occ[doc_key] = (idx_f, fname)
            signatures[(idx_f, doc_key)] = _signature_dokumen_pg(df_sub)

    # 3. Buang dokumen duplikat lintas file dari file lebih awal (keep='last') & catat ke log
    all_df_dedup = []
    for idx_f, (fname, df_f) in enumerate(per_file_dfs):
        keys_series = list(zip(
            df_f['No_Transaksi_PG'].str.strip().str.upper(),
            df_f['Gudang_PG'].str.strip().str.upper()
        ))
        mask_buang = [
            (k[0] != '' and k in last_occ and last_occ[k][0] != idx_f)
            for k in keys_series
        ]

        if any(mask_buang):
            df_buang = df_f[mask_buang]
            for _, r in df_buang.iterrows():
                doc_key = (r['No_Transaksi_PG'].strip().upper(), r['Gudang_PG'].strip().upper())
                kept_idx, kept_fname = last_occ[doc_key]
                sama = (signatures.get((idx_f, doc_key)) == signatures.get((kept_idx, doc_key)))
                jenis_dup = (
                    'Duplikat Lintas File (Identik)'
                    if sama else
                    'Duplikat Lintas File (Isi Berbeda - Indikasi Revisi ERP)'
                )
                log_rows.append({
                    'Sumber Master': 'Master PG',
                    'Jenis Duplikat': jenis_dup,
                    'Tindakan Sistem': f'Dihapus Otomatis (Pakai versi file: {kept_fname})',
                    'Nama File Asal': fname,
                    'No. Transaksi / Kode': r['No_Transaksi_PG'],
                    'Tanggal': r['Tanggal_PG'],
                    'Gudang': r['Gudang_PG'],
                    'PIBC / Keterangan': r['PIBC_PG'] or r['Keterangan_PG'],
                    'Kode Item': r['Kode_Item_PG'],
                    'Nama Item': r['Nama_Item_PG'],
                    'Qty (GR)': round(bersihkan_angka_indo(r['Qty_PG']), 3)
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
