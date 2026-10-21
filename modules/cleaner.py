import pandas as pd

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


def bersihkan_master_pg(file_excel):
    """
    Membersihkan Master PG / Pindah Gudang (menghapus 3 baris judul di atas).
    """
    df_pg = pd.read_excel(file_excel)
    if len(df_pg) > 3:
        df_pg = df_pg.iloc[3:].reset_index(drop=True)
    df_pg.columns = [str(col).strip() for col in df_pg.columns]
    return df_pg
