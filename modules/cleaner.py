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
