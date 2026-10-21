import pandas as pd

def bersihkan_master_mb(file_excel):
    """
    Fungsi untuk membersihkan file Master Manufacture Basic (MB).
    Aturan:
    1. Menghapus 4 baris pertama dan 4 baris terakhir (judul & total awal/akhir).
    2. Melakukan Forward Fill pada kolom utama berdasarkan kesamaan 'Kode'.
    """
    df_mb = pd.read_excel(file_excel)
    
    # Menghapus 4 baris pertama dan 4 baris terakhir jika baris data mencukupi
    if len(df_mb) > 8:
        df_mb = df_mb.iloc[4:-4].reset_index(drop=True)
        
    # Daftar kolom yang wajib di-forward fill berdasarkan kolom 'Kode'
    kolom_utama = ['Gudang', 'Kode', 'Kode Barang Jadi', 'Nama Barang Jadi', 'Keterangan', 'Keterangan Lain']
    
    for nama_kolom in kolom_utama:
        if nama_kolom in df_mb.columns:
            df_mb[nama_kolom] = df_mb[nama_kolom].ffill()
            
    return df_mb


def bersihkan_master_pibc(file_excel):
    """
    Fungsi untuk membersihkan file Master PIBC/Dyelot.
    Aturan: Menghapus 1 baris terakhir karena merupakan baris total.
    """
    df_pibc = pd.read_excel(file_excel)
    if len(df_pibc) > 1:
        df_pibc = df_pibc.iloc[:-1].reset_index(drop=True)
    return df_pibc


def bersihkan_master_pg(file_excel):
    """
    Fungsi untuk membersihkan file Master PG (Pindah Gudang).
    Aturan: Menghapus 3 baris pertama karena merupakan judul file.
    """
    df_pg = pd.read_excel(file_excel)
    if len(df_pg) > 3:
        df_pg = df_pg.iloc[3:].reset_index(drop=True)
    return df_pg
