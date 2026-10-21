import pandas as pd

def siapkan_kolom_standar(df_mb):
    """
    Fungsi untuk menambahkan kolom baru dan mengonversi satuan ke GR pada Data Master MB.
    """
    # 1. Konversi Qty Target ke Satuan GR
    if 'Target Unit' in df_mb.columns and 'Target Qty' in df_mb.columns:
        def konversi_target_gr(row):
            unit = str(row['Target Unit']).strip().upper()
            qty = row['Target Qty']
            try:
                qty_angka = float(qty)
            except:
                qty_angka = 0.0
            
            # Jika bukan GR, kalikan 1000
            if unit != 'GR':
                return qty_angka * 1000
            else:
                return qty_angka
                
        df_mb['Qty Target Standar (GR)'] = df_mb.apply(konversi_target_gr, axis=1)

    # 2. Konversi Qty BB ke Satuan GR
    if 'Unit' in df_mb.columns and 'Qty' in df_mb.columns:
        def konversi_bb_gr(row):
            unit = str(row['Unit']).strip().upper()
            qty = row['Qty']
            try:
                qty_angka = float(qty)
            except:
                qty_angka = 0.0
            
            if unit != 'GR':
                return qty_angka * 1000
            else:
                return qty_angka
                
        df_mb['Qty BB Standar (GR)'] = df_mb.apply(konversi_bb_gr, axis=1)

    # Menambahkan kolom kosong penampung hasil crosscheck
    df_mb['Crosscheck Qty'] = ''
    df_mb['Qty PIBC'] = 0.0
    df_mb['Qty PG'] = 0.0
    df_mb['Selisih (GR)'] = 0.0
    df_mb['Status'] = ''
    df_mb['Status PIBC'] = ''
    df_mb['Rincian Benang'] = ''
    df_mb['Cek Variant'] = ''

    return df_mb


def filter_baris_diabaikan(df_mb):
    """
    Fungsi untuk menyaring baris total/overhead yang harus dihiraukan 
    pada kolom 'Kode Item' (tidak dihapus, hanya dilewati saat pengecekan).
    """
    baris_diabaikan = [
        "Total Material (Details)", 
        "Total Pekerjaan (Works)", 
        "Total (Details + Works)", 
        "Overhead Cost", 
        "TOTAL HARGA", 
        "Labor"
    ]
    
    if 'Kode Item' in df_mb.columns:
        mask_valid = ~df_mb['Kode Item'].astype(str).isin(baris_diabaikan)
        return df_mb[mask_valid]
    
    return df_mb
