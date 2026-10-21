def is_item_benang(kode_item, gudang=""):
    """
    Menentukan apakah suatu item adalah benang berdasarkan kode item dan gudangnya:
    - Gudang Mesin Dyeing & LAB/RND: Benang HANYA TWP, MWP, TBM, MBB (TBB adalah Chemical).
    - Gudang Softcone / Lainnya: Benang mencakup TWP, MWP, TBM, MBB, dan TBB.
    """
    k_upper = str(kode_item).strip().upper()
    g_upper = str(gudang).strip().upper()
    
    # Prefix utama benang
    prefix_utama = ('TWP', 'MWP', 'TBM', 'MBB')
    
    if k_upper.startswith(prefix_utama):
        return True
        
    # Jika berada di Mesin Dyeing atau LAB/RND, TBB BUKAN benang (dianggap chemical)
    if 'MESIN DYEING' in g_upper or 'LAB' in g_upper or 'RND' in g_upper:
        return False
        
    # Untuk Softcone atau gudang lain, TBB adalah benang
    if k_upper.startswith('TBB'):
        return True
        
    return False
