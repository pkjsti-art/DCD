import pandas as pd
import re

# =============================================================================
# FUNGSI BANTUAN (HELPER) UNTUK CLEANING
# =============================================================================
def bersihkan_angka_indo(val):
    """Mengubah format angka Excel (ribuan titik, desimal koma) menjadi Float"""
    if pd.isna(val) or val is None: return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('.', '').replace(',', '.')
    try: return float(val_str)
    except: return 0.0

def cari_kolom_akurat(df, daftar_kemungkinan):
    """Mencari nama kolom asli di Excel mengabaikan titik, spasi, atau garis bawah"""
    for kem in daftar_kemungkinan:
        kem_clean = re.sub(r'[^A-Z0-9]', '', kem.upper())
        for col in df.columns:
            col_clean = re.sub(r'[^A-Z0-9]', '', str(col).upper())
            if kem_clean in col_clean:
                return col
    return None

# =============================================================================
# FUNGSI CLEANING UTAMA
# =============================================================================

def bersihkan_master_mb(file_excel):
    """
    Mendeteksi baris header otomatis, membuang baris judul di atas header,
    membuang footer, dan melakukan forward fill pada kolom-kolom utama MB.
    """
    df_raw = pd.read_excel(file_excel, header=None)
    
    header_row_index = 0
    for idx, row in df_raw.iterrows():
        row_str_list = [str(val).strip() for val in row.values]
        if any(k in row_str_list for k in ['Kode', 'Kode Item', 'Target Qty', 'Kode Barang Jadi']):
            header_row_index = idx
            break
            
    df_mb = pd.read_excel(file_excel, header=header_row_index)
    df_mb.columns = [str(col).strip() for col in df_mb.columns]
    df_mb = df_mb.dropna(how='all').reset_index(drop=True)
    
    if len(df_mb) > 4:
        df_mb = df_mb.iloc[:-4].reset_index(drop=True)
        
    kolom_utama = ['Gudang', 'Kode', 'Kode Barang Jadi', 'Nama Barang Jadi', 'Keterangan', 'Keterangan Lain']
    for nama_kolom in kolom_utama:
        if nama_kolom in df_mb.columns:
            df_mb[nama_kolom] = df_mb[nama_kolom].ffill()
            
    return df_mb


def bersihkan_master_pibc(file_excel):
    """
    Membersihkan Master PIBC / Dyelot (menghapus baris total di akhir).
    """
    df_pibc = pd.read_excel(file_excel)
    if len(df_pibc) > 1:
        df_pibc = df_pibc.iloc[:-1].reset_index(drop=True)
    df_pibc.columns = [str(col).strip() for col in df_pibc.columns]
    return df_pibc


def bersihkan_master_pg(daftar_file_pg):
    """
    Membersihkan Master PG / Pindah Gudang secara agresif dan digabung jadi 1 file:
    1. Menghapus 3 baris judul di atas (menggunakan header=3).
    2. Menghapus baris yang nilai MASUK-nya 0.
    3. Menstandarisasi nama kolom.
    4. Menggabungkan semua file PG yang diupload menjadi satu DataFrame siap pakai.
    """
    # Jika yang diinput cuma 1 file (bukan list), jadikan list agar loop tidak error
    if not isinstance(daftar_file_pg, list):
        daftar_file_pg = [daftar_file_pg]
        
    all_df_bersih = []
    
    for file in daftar_file_pg:
        try:
            # header=3 otomatis membuang 3 baris teratas (index 0,1,2) 
            # dan menjadikan baris ke-4 sebagai nama kolom
            df = pd.read_excel(file, header=3)
        except Exception as e:
            continue
            
        if df.empty:
            continue
            
        # Cari nama kolom intinya agar kebal terhadap perubahan nama spasi/titik
        col_masuk = cari_kolom_akurat(df, ['MASUK'])
        col_ket_det = cari_kolom_akurat(df, ['KETERANGANDET'])
        col_kode_var = cari_kolom_akurat(df, ['KODEVARIANT', 'KODEITEM'])
        col_nama_var = cari_kolom_akurat(df, ['NAMAVARIANT', 'NAMAITEM'])
        col_gudang = cari_kolom_akurat(df, ['GUDANG'])
        
        if not col_masuk: 
            continue
        
        # Bersihkan format angka agar terdeteksi 0 dengan benar
        df['TEMP_MASUK'] = df[col_masuk].apply(bersihkan_angka_indo)
        
        # HAPUS MUTLAK BARIS YANG MASUKNYA = 0
        df_bersih = df[df['TEMP_MASUK'] > 0].copy()
        
        # Rename kolom ke nama baku internal sistem
        rename_dict = {}
        if col_gudang: rename_dict[col_gudang] = 'Gudang_PG'
        if col_ket_det: rename_dict[col_ket_det] = 'PIBC_PG'
        if col_kode_var: rename_dict[col_kode_var] = 'Kode_Item_PG'
        if col_nama_var: rename_dict[col_nama_var] = 'Nama_Item_PG'
        
        df_bersih.rename(columns=rename_dict, inplace=True)
        df_bersih['Qty_PG'] = df_bersih['TEMP_MASUK']
        
        # Filter hanya kolom inti yang kita butuhkan untuk memperingan memori
        kolom_inti = ['Gudang_PG', 'PIBC_PG', 'Kode_Item_PG', 'Nama_Item_PG', 'Qty_PG']
        for kf in kolom_inti:
            if kf not in df_bersih.columns:
                df_bersih[kf] = ""
                
        all_df_bersih.append(df_bersih[kolom_inti])
        
    # GABUNGKAN SEMUA FILE MENJADI 1 DATAFRAME MATANG
    if all_df_bersih:
        df_gabungan = pd.concat(all_df_bersih, ignore_index=True)
        return df_gabungan
    else:
        # Jika gagal atau datanya kosong, buat format kosong yang aman
        return pd.DataFrame(columns=['Gudang_PG', 'PIBC_PG', 'Kode_Item_PG', 'Nama_Item_PG', 'Qty_PG'])
