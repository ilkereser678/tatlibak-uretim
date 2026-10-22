def urun_dagilimi(detay, hedef, toplam, pencere=60, orijinal_sira=None, haric_urunler=None):
    if detay is None or detay.empty:
        return pd.DataFrame(), "Veri yok"
    
    son = detay["Tarih"].max()
    d = detay[detay["Tarih"] > son - pd.Timedelta(days=pencere)].copy()
    
    haric_norm = {_norm(x) for x in (haric_urunler or [])}
    if haric_norm:
        d = d[~d["Cesit"].map(_norm).isin(haric_norm)]
    if d.empty:
        return pd.DataFrame(), "Veri yok"

    # Hafta sonu / Hafta içi filtresi (Cuma-Pazar vs Pazartesi-Perşembe)
    hs = hedef.dayofweek >= 4
    maske = d["Tarih"].dt.dayofweek.ge(4) == hs
    d_filt = d[maske]
    if d_filt.empty:
        d_filt = d  # Filtre boş kalırsa tüm döneme dön

    # Çeşit bazlı son dönem toplamları
    g = d_filt.groupby("Cesit")["Adet"].sum()
    
    mevcut = list(g.index)
    cesitler = [c for c in (orijinal_sira or []) if c in mevcut]
    cesitler += [c for c in mevcut if c not in cesitler]
    
    g = g.reindex(cesitler).fillna(0)
    toplam_satis = g.sum()
    
    if toplam_satis <= 0:
        return pd.DataFrame(), "Veri yok"

    # Oransal dağılımı hesapla ve hedeflenen toplam üretime (örn: 2092 adete) sabitle
    oranlar = g / toplam_satis
    dagilim_adetleri = oranlar * float(toplam)
    
    sonuc = yuvarla_topla(dagilim_adetleri.to_numpy(), float(toplam))
    tablo = pd.DataFrame({"Önerilen Üretim (Adet)": sonuc}, index=cesitler)
    tablo.loc["TOPLAM"] = tablo.sum()
    tablo.index.name = "Ürün Çeşidi"
    return tablo.astype(int), ("Cuma-Pazar" if hs else "Pazartesi-Perşembe")