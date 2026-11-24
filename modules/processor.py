def parse_tanggal_aman(val):
    if pd.isna(val) or val is None:
        return pd.NaT
    if isinstance(val, (pd.Timestamp, datetime)):
        return val
    if hasattr(val, 'year') and hasattr(val, 'month') and hasattr(val, 'day'):
        return pd.Timestamp(val)
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'nat']:
        return pd.NaT
    if re.match(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}', s):
        return pd.to_datetime(s, dayfirst=False, errors='coerce')
    if re.match(r'^\d{1,2}[-/]\d{1,2}[-/]\d{4}', s):
        return pd.to_datetime(s, dayfirst=True, errors='coerce')
    return pd.to_datetime(s, dayfirst=True, errors='coerce')

def hitung_masa_dan_rentang(df, col_tgl, nama_kolom_masa):
    df_out = df.copy()
    hari_ini = get_hari_ini_wib()
    
    if col_tgl and col_tgl in df_out.columns:
        if pd.api.types.is_datetime64_any_dtype(df_out[col_tgl]):
            tgl_parsed = df_out[col_tgl]
        else:
            tgl_parsed = df_out[col_tgl].apply(parse_tanggal_aman)
            
        masa_list = []
        rentang_list = []
        for dt in tgl_parsed:
            if pd.isna(dt):
                masa_list.append(pd.NA)
                rentang_list.append("")
            else:
                selisih = int((hari_ini - dt.date()).days)
                masa_list.append(selisih)
                if selisih <= 3:
                    rentang_list.append("0 - 3 Hari")
                elif selisih <= 7:
                    rentang_list.append("4 - 7 Hari")
                else:
                    rentang_list.append("> 7 Hari")
                    
        df_out[nama_kolom_masa] = pd.Series(masa_list, index=df_out.index, dtype="Int64")
        df_out['Rentang Hari'] = rentang_list
        df_out = df_out.sort_values(by=nama_kolom_masa, ascending=False, na_position='last').reset_index(drop=True)
    else:
        df_out[nama_kolom_masa] = pd.Series([pd.NA] * len(df_out), dtype="Int64")
        df_out['Rentang Hari'] = ""
        
    return df_out
