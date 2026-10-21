import pandas as pd

def bersihkan_master_mb(file_excel):
    """
    Fungsi untuk membersihkan file Master Manufacture Basic (MB).
    Secara otomatis mendeteksi baris header yang benar, menghapus baris total di akhir,
    dan melakukan Forward Fill pada kolom utama berdasarkan kesamaan 'Kode'.
    """
    df_raw = pd.read_excel(file_excel, header=None)
    
    header_row_index = 0
    for idx, row in df_raw.iterrows():
        row_str_list = [str(val).strip() for val in row.values]
        if 'Kode' in row_str_list or 'Kode Item' in row_str_list or 'Target Qty' in row_str_list:
            header_row_index = idx
            break
            
    df_mb = pd.read_excel(file_excel, header=header_row_index)
    df_mb.columns = [str(col).strip() for col in df_mb.columns]
    df_mb = df_mb.dropna(how='all').reset_index(drop=True)
    
    # Menghapus 4 baris terakhir (total/footer) jika mencukupi
    if len(df_mb) > 4:
        df_mb = df_mb.iloc[:-4].reset_index(drop=True)
        
    # Forward Fill kolom utama berdasarkan kesamaan 'Kode'
    kolom_utama = ['Gudang', 'Kode', 'Kode Barang Jadi', 'Nama Barang Jadi', 'Keterangan', 'Keterangan Lain']
    for nama_kolom in kolom_utama:
        if nama_kolom in df_mb.columns:
            df_mb[nama_kolom] = df_mb[nama_kolom].ffill()
            
    return df_mb


def bersihkan_master_pibc(file_excel):
    """
    Fungsi untuk membersihkan file Master PIBC/Dyelot (Menghapus baris total di akhir)[cite: 3].
    """
    df_pibc = pd.read_excel(file_excel)
    if len(df_pibc) > 1:
        df_pibc = df_pibc.iloc[:-1].reset_index(drop=True)
    df_pibc.columns = [str(col).strip() for col in df_pibc.columns]
    return df_pibc


def bersihkan_master_pg(file_excel):
    """
    Fungsi untuk membersihkan file Master PG / Pindah Gudang (Menghapus 3 baris judul)[cite: 3].
    """
    df_pg = pd.read_excel(file_excel)
    if len(df_pg) > 3:
        df_pg = df_pg.iloc[3:].reset_index(drop=True)
    df_pg.columns = [str(col).strip() for col in df_pg.columns]
    return df_pg
