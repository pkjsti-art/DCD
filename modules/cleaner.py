import pandas as pd
import re

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

def bersihkan_master_pibc(file_excel):
    file_excel.seek(0)
    df_pibc = pd.read_excel(file_excel)
    if len(df_pibc) > 1:
        df_pibc = df_pibc.iloc[:-1].reset_index(drop=True)
    df_pibc.columns = [str(col).strip() for col in df_pibc.columns]
    df_pibc = df_pibc.replace(r'^\s*$', pd.NA, regex=True).dropna(how='all').reset_index(drop=True)
    return df_pibc

def bersihkan_master_pg(daftar_file_pg):
    if not isinstance(daftar_file_pg, list):
        daftar_file_pg = [daftar_file_pg]
        
    all_df_bersih = []
    ada_kolom_status = False
    
    for file in daftar_file_pg:
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
        
        if not col_masuk: 
            continue

        if col_status:
            ada_kolom_status = True
        
        df['TEMP_MASUK'] = df[col_masuk].apply(bersihkan_angka_indo)
        df_bersih = df[df['TEMP_MASUK'] > 0].copy()
        
        rename_dict = {}
        if col_no_trans: rename_dict[col_no_trans] = 'No_Transaksi_PG'
        if col_tanggal: rename_dict[col_tanggal] = 'Tanggal_PG'
        if col_gudang: rename_dict[col_gudang] = 'Gudang_PG'
        if col_ket_det: rename_dict[col_ket_det] = 'PIBC_PG'
        if col_kode_var: rename_dict[col_kode_var] = 'Kode_Item_PG'
        if col_nama_var: rename_dict[col_nama_var] = 'Nama_Item_PG'
        if col_status: rename_dict[col_status] = 'Status_PG'
        
        df_bersih.rename(columns=rename_dict, inplace=True)
        df_bersih['Qty_PG'] = df_bersih['TEMP_MASUK']
        
        if 'No_Transaksi_PG' in df_bersih.columns:
            df_bersih['No_Transaksi_PG'] = df_bersih['No_Transaksi_PG'].fillna('').astype(str).str.strip()
            df_bersih.loc[df_bersih['No_Transaksi_PG'].str.lower().isin(['nan', 'none', 'nat']), 'No_Transaksi_PG'] = ''

        kolom_inti = ['No_Transaksi_PG', 'Tanggal_PG', 'Gudang_PG', 'PIBC_PG', 'Kode_Item_PG', 'Nama_Item_PG', 'Qty_PG', 'Status_PG']
        for kf in kolom_inti:
            if kf not in df_bersih.columns:
                df_bersih[kf] = ""
                
        all_df_bersih.append(df_bersih[kolom_inti])
        
    kolom_output = ['No_Transaksi_PG', 'Tanggal_PG', 'Gudang_PG', 'PIBC_PG', 'Kode_Item_PG', 'Nama_Item_PG', 'Qty_PG', 'Status_PG']
    if all_df_bersih:
        df_gabungan = pd.concat(all_df_bersih, ignore_index=True)
        if ada_kolom_status and 'Status_PG' in df_gabungan.columns:
            df_gabungan = df_gabungan[
                df_gabungan['Status_PG'].astype(str).str.strip().str.lower() == 'active'
            ].reset_index(drop=True)
        return df_gabungan
    else:
        return pd.DataFrame(columns=kolom_output)

def bersihkan_master_mesin(file_excel):
    if file_excel is None:
        return None
    file_excel.seek(0)
    df_mesin = pd.read_excel(file_excel)
    df_mesin.columns = [str(col).strip() for col in df_mesin.columns]
    df_mesin = df_mesin.replace(r'^\s*$', pd.NA, regex=True).dropna(how='all').reset_index(drop=True)
    
    col_mesin = cari_kolom_akurat(df_mesin, ['MESIN', 'NAMAMESIN', 'KODEMESIN'])
    col_jenis = cari_kolom_akurat(df_mesin, ['JENISMESIN', 'JENIS', 'KATEGORI'])
    
    if not col_mesin or not col_jenis:
        return pd.DataFrame(columns=['Mesin', 'JENIS MESIN'])
        
    df_out = df_mesin[[col_mesin, col_jenis]].copy()
    df_out.columns = ['Mesin', 'JENIS MESIN']
    df_out = df_out.dropna(subset=['Mesin']).reset_index(drop=True)
    df_out['Mesin'] = df_out['Mesin'].astype(str).str.strip()
    df_out['JENIS MESIN'] = df_out['JENIS MESIN'].astype(str).str.strip().str.upper()
    return df_out

def deteksi_duplikasi_mb(list_df_mb):
    """Mendeteksi Kode MB yang muncul di lebih dari 1 file upload."""
    if len(list_df_mb) <= 1:
        return []
    kode_per_file = {}
    duplikat = set()
    for idx_f, df in enumerate(list_df_mb):
        if 'Kode' not in df.columns:
            continue
        kodes = set(df['Kode'].dropna().astype(str).str.strip())
        kodes = {k for k in kodes if k and k.lower() not in ['nan', 'none', 'nat']}
        for k in kodes:
            if k in kode_per_file:
                duplikat.add(k)
            else:
                kode_per_file[k] = idx_f
    return sorted(list(duplikat))

def deteksi_duplikasi_pg(daftar_file_pg):
    """Mendeteksi No. Transaksi PG yang muncul di lebih dari 1 file upload."""
    if not isinstance(daftar_file_pg, list) or len(daftar_file_pg) <= 1:
        return []
    trans_seen = set()
    duplikat = set()
    for file in daftar_file_pg:
        try:
            file.seek(0)
            df = pd.read_excel(file, header=3)
        except Exception:
            continue
        if df.empty:
            continue
        col_no_trans = cari_kolom_akurat(df, ['NOTRANSAKSI', 'KODETRANSAKSI'])
        if not col_no_trans:
            continue
        trans_file = set(df[col_no_trans].dropna().astype(str).str.strip())
        trans_file = {t for t in trans_file if t and t.lower() not in ['nan', 'none', 'nat']}
        for t in trans_file:
            if t in trans_seen:
                duplikat.add(t)
            else:
                trans_seen.add(t)
    return sorted(list(duplikat))
