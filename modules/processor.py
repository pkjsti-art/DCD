import pandas as pd
import re

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
            if unit != 'GR':
                return qty_angka * 1000
            else:
                return qty_angka
        df_mb['Qty Target Standar (GR)'] = df_mb.apply(konversi_target_gr, axis=1)
    else:
        df_mb['Qty Target Standar (GR)'] = 0.0

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
    else:
        df_mb['Qty BB Standar (GR)'] = 0.0

    # Menambahkan kolom penampung hasil crosscheck
    df_mb['Crosscheck Qty'] = 'CORRECT' # Default awal
    df_mb['Qty PIBC'] = 0.0
    df_mb['Qty PG'] = 0.0
    df_mb['Selisih (GR)'] = 0.0
    df_mb['Status'] = 'COMPLATE'
    df_mb['Status PIBC'] = 'KODE PIBC ADA'
    df_mb['Rincian Benang'] = ''
    df_mb['Cek Variant'] = 'CLEAN'

    return df_mb


def ekstrak_kode_pibc(teks):
    """
    Fungsi cerdas menggunakan Regex untuk mendeteksi dan mengambil kode PIBC
    dari teks keterangan yang bervariasi (misal: TU-PIBC-260000410, PIBC-2610000410-TU, dll).
    """
    if pd.isna(teks):
        return None
    teks_str = str(teks)
    
    # Pola regex untuk mencari format PIBC diikuti angka/karakter
    pola = r'(PIBC-[A-Za-z0-9#\-\.]+)'
    hasil = re.search(pola, teks_str, re.IGNORECASE)
    
    if hasil:
        # Membersihkan karakter titik atau spasi di akhir jika ada
        kode = hasil.group(1).strip().rstrip('.')
        return kode
    return None


def jalankan_proses_crosscheck(df_mb, df_pibc, df_pg):
    """
    Fungsi utama untuk mengeksekusi logika crosscheck data MB dengan Master PIBC dan Master PG.
    """
    baris_diabaikan = [
        "Total Material (Details)", "Total Pekerjaan (Works)", 
        "Total (Details + Works)", "Overhead Cost", "TOTAL HARGA", "Labor"
    ]
    
    # Melakukan iterasi atau pemrosesan per baris / kelompok kode transaksi MB
    for idx, row in df_mb.iterrows():
        kode_item = str(row.get('Kode Item', '')).strip()
        
        # 1. Lewati baris total/overhead yang harus dihiraukan
        if kode_item in baris_diabaikan:
            df_mb.at[idx, 'Crosscheck Qty'] = 'DIHIRAUKAN'
            continue
            
        gudang = str(row.get('Gudang', '')).strip().upper()
        keterangan_lain = str(row.get('Keterangan Lain', ''))
        keterangan = str(row.get('Keterangan', ''))
        
        # 2. Deteksi Kode PIBC dari Keterangan
        teks_gabungan = keterangan + " " + keterangan_lain
        kode_pibc_ditemukan = ekstrak_kode_pibc(teks_gabungan)
        
        if not kode_pibc_ditemukan:
            df_mb.at[idx, 'Status PIBC'] = 'KODE PIBC TIDAK ADA'
            
        # 3. Logika Pengecekan Benang Berdasarkan Gudang
        if 'PRODUKSI SOFTCONE' in gudang or kode_item.startswith(('TWP', 'MWP', 'TBM', 'MBB', 'TBB')):
            qty_target = row.get('Qty Target Standar (GR)', 0.0)
            qty_bb = row.get('Qty BB Standar (GR)', 0.0)
            selisih = qty_bb - qty_target
            df_mb.at[idx, 'Selisih (GR)'] = selisih
            
            if abs(selisih) > 0.01:
                df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
            else:
                df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                
            df_mb.at[idx, 'Rincian Benang'] = kode_item
            
        # 4. Logika Pengecekan Chemical & Dyestuff (Gudang Mesin Dyeing & Lab)
        elif 'DYEING' in gudang or 'LAB' in gudang:
            # Pencocokan sederhana dengan Master PIBC jika kolom tersedia
            if df_pibc is not None and not df_pibc.empty and kode_pibc_ditemukan:
                # Contoh logika matching dengan Master PIBC
                df_match_pibc = df_pibc[df_pibc.astype(str).apply(lambda x: x.str.contains(kode_pibc_ditemukan, case=False, na=False)).any(axis=1)]
                
                if not df_match_pibc.empty:
                    # Ambil nilai actual jika ditemukan
                    qty_actual = float(df_match_pibc.iloc[0].get('Actual', 0.0) or 0.0)
                    qty_bb = row.get('Qty BB Standar (GR)', 0.0)
                    df_mb.at[idx, 'Qty PIBC'] = qty_actual
                    
                    selisih = qty_bb - qty_actual
                    df_mb.at[idx, 'Selisih (GR)'] = selisih
                    
                    if abs(selisih) > 0.01:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'INCORRECT'
                    else:
                        df_mb.at[idx, 'Crosscheck Qty'] = 'CORRECT'
                else:
                    # Jika tidak ada di PIBC, tandai item tidak ditemukan sementara
                    df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'
            else:
                df_mb.at[idx, 'Crosscheck Qty'] = 'ITEM TIDAK DI TEMUKAN'

    return df_mb
