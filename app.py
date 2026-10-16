import datetime as dt
import glob
import io
import json
import os
import sys
from urllib.parse import urlencode
from urllib.request import urlopen

import numpy as np
import pandas as pd
import streamlit as st

try:
    from zoneinfo import ZoneInfo
    BUGUN = dt.datetime.now(ZoneInfo("Europe/Istanbul")).date()
except Exception:
    BUGUN = dt.date.today()

st.set_page_config(page_title="Tatlıbak Üretim Planlama", page_icon="🍰", layout="wide")

# ----------------------------------------------------------------------------
# AYARLAR VE DOSYA İSİMLERİ
# ----------------------------------------------------------------------------
GUNLUK_DOSYALAR = ["2026 pasta satış.xlsx", "gunluk_satis.xlsx", "gunluk_satis.csv"]
DETAY_DOSYALAR = ["2026 pasta satış.xlsx", "urun_detay.xlsx", "urun_detay.csv"]
GECMIS_DOSYALAR = ["geçmiş yıllar data.xlsx", "gecmis_arsiv.xlsx"]
EKLENEN_DOSYA = "eklenen_gunluk.csv"
DETAY_EKLENEN_DOSYA = "eklenen_detay.csv"

GUN_TR = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

BAYRAMLAR = [
    ("2023-04-21", "2023-04-23"), ("2023-06-28", "2023-07-01"),
    ("2024-04-10", "2024-04-12"), ("2024-06-16", "2024-06-19"),
    ("2025-03-30", "2025-04-01"), ("2025-06-06", "2025-06-09"),
    ("2026-03-20", "2026-03-22"), ("2026-05-27", "2026-05-30"),
    ("2027-03-09", "2027-03-11"), ("2027-05-16", "2027-05-19"),
]

GRUPLAR = {
    "Gün tipi": [f"gun_{k}" for k in range(7)],
    "Maaş / ay günü": ["ay_basi", "ay_ortasi", "ay_sonu"],
    "Mevsim": ["mevsim_sin1", "mevsim_cos1", "mevsim_sin2", "mevsim_cos2"],
    "Bayram / özel gün": ["bayram", "bayram_oncesi", "resmi_tatil", "ozel_gun"],
    "Hava": ["sicaklik", "yagmurlu"],
    "Trend / menü": ["trend", "menu_sonrasi"],
}

SEVIYELER = {
    "Dengeli (%50) - fire az, kaçan satış olabilir": 0.5,
    "Temkinli (%70)": 0.7,
    "Güvenli (%80)": 0.8,
    "Çok güvenli (%90) - fire artar": 0.9,
}

YONTEM_AD = {"Baz": "Baz çizgi", "Model": "İstatistiksel model", "Karma": "Karma (baz + model)"}


# ----------------------------------------------------------------------------
# YARDIMCI FONKSİYONLAR
# ----------------------------------------------------------------------------
def _norm(s):
    s = str(s).strip().replace("İ", "i").replace("I", "i").lower()
    for a, b in zip("ıüşöçğâî", "iusocgai"):
        s = s.replace(a, b)
    return s


def _evet(v):
    return str(v).strip().lower() in ("evet", "e", "1", "true", "x", "var", "yes", "bitti")


def _sayi(v):
    if isinstance(v, str):
        v = v.replace(".", "").replace(",", ".")
    return pd.to_numeric(v, errors="coerce")


def _tarih(seri):
    t = pd.to_datetime(seri, dayfirst=True, errors="coerce")
    return t.dt.normalize().astype("datetime64[ns]")


def mape(gercek, tahmin):
    gercek = np.asarray(gercek, float)
    tahmin = np.asarray(tahmin, float)
    return float(np.mean(np.abs(gercek - tahmin) / gercek))


def nth_weekday(yil, ay, hafta_gunu, n):
    d = dt.date(yil, ay, 1)
    kayma = (hafta_gunu - d.weekday()) % 7
    return d + dt.timedelta(days=kayma + 7 * (n - 1))


def ozel_gunler(yillar, ek_metin):
    bayram, oncesi, resmi, ozel = set(), set(), set(), set()
    for b, e in BAYRAMLAR:
        bs, be = dt.date.fromisoformat(b), dt.date.fromisoformat(e)
        for i in range((be - bs).days + 1):
            bayram.add(bs + dt.timedelta(days=i))
        for i in range(1, 4):
            oncesi.add(bs - dt.timedelta(days=i))
    for y in yillar:
        for m, g in [(1, 1), (4, 23), (5, 1), (5, 19), (7, 15), (8, 30), (10, 29)]:
            resmi.add(dt.date(y, m, g))
        for m, g in [(2, 14), (3, 8), (11, 24), (12, 31)]:
            ozel.add(dt.date(y, m, g))
        ozel.add(nth_weekday(y, 5, 6, 2))
        ozel.add(nth_weekday(y, 6, 6, 3))
    for satir in (ek_metin or "").splitlines():
        try:
            ozel.add(dt.date.fromisoformat(satir.strip()))
        except ValueError:
            pass
    return {"bayram": bayram, "oncesi": oncesi, "resmi": resmi, "ozel": ozel}


# ----------------------------------------------------------------------------
# VERİ OKUMA VE BİRLEŞTİRME
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def tablo_oku(ad, ham):
    if ad.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "cp1254", "latin1"):
            try:
                return pd.read_csv(io.BytesIO(ham), sep=None, engine="python", encoding=enc)
            except Exception:
                continue
        raise ValueError("CSV okunamadı")
    xls = pd.ExcelFile(io.BytesIO(ham))
    sheet_name = [s for s in xls.sheet_names if "siparis" in _norm(s) or "satis" in _norm(s)]
    sheet_name = sheet_name[0] if sheet_name else xls.sheet_names[0]
    return xls.parse(sheet_name, header=None)


def kaynak_bul(yuklenen, adaylar, anahtarlar, haric=()):
    if yuklenen is not None:
        return yuklenen.name, yuklenen.getvalue()
    bulunan = [f for f in adaylar if os.path.exists(f)]
    if not bulunan:
        for f in sorted(glob.glob("*.xlsx") + glob.glob("*.csv")):
            n = _norm(f)
            if any(a in n for a in anahtarlar) and not any(h in n for h in haric):
                bulunan.append(f)
    if bulunan:
        with open(bulunan[0], "rb") as fh:
            return bulunan[0], fh.read()
    return None, None


def matris_isle_2026(df):
    detay_rows = []
    gunluk_rows = []
    
    for col_idx in range(1, df.shape[1]):
        t = pd.to_datetime(df.iloc[1, col_idx], errors="coerce")
        if pd.isna(t):
            continue
        
        gunluk_toplam = 0
        for row_idx in range(2, df.shape[0]):
            p_ad = str(df.iloc[row_idx, 0]).strip()
            val = _sayi(df.iloc[row_idx, col_idx])
            if pd.isna(val) or val < 0:
                val = 0
            
            if "toplam" in _norm(p_ad):
                gunluk_toplam = val
                continue
            
            if p_ad and p_ad.lower() != "nan" and p_ad != "":
                detay_rows.append({"Tarih": t, "Cesit": p_ad, "Adet": val})
                gunluk_toplam += val
                
        gunluk_rows.append({"Tarih": t, "Adet": gunluk_toplam, "Bitti": False, "Not": ""})
        
    df_gunluk = pd.DataFrame(gunluk_rows).dropna(subset=["Tarih", "Adet"])
    df_detay = pd.DataFrame(detay_rows).dropna(subset=["Tarih", "Adet"])
    return df_gunluk, df_detay


def gecmis_veri_oku():
    bulunan = [f for f in GECMIS_DOSYALAR if os.path.exists(f)]
    if not bulunan:
        return pd.DataFrame(columns=["Tarih", "Adet", "Bitti", "Not"])
    
    tum_gecmis = []
    for dosya in bulunan:
        try:
            xls = pd.ExcelFile(dosya)
            for sheet in xls.sheet_names:
                df_raw = xls.parse(sheet, header=None)
                t_col_idx, a_col_idx = 0, 1
                for c_idx in range(df_raw.shape[1]):
                    col_vals = df_raw.iloc[:, c_idx]
                    parsed_dates = pd.to_datetime(col_vals, errors="coerce")
                    if parsed_dates.notna().mean() > 0.5:
                        t_col_idx = c_idx
                        break
                        
                for c_idx in range(df_raw.shape[1]):
                    if c_idx == t_col_idx:
                        continue
                    num_vals = pd.to_numeric(df_raw.iloc[:, c_idx], errors="coerce")
                    if num_vals.notna().mean() > 0.5:
                        a_col_idx = c_idx
                        break
                
                temp = pd.DataFrame({
                    "Tarih": _tarih(df_raw.iloc[:, t_col_idx]),
                    "Adet": df_raw.iloc[:, a_col_idx].map(_sayi),
                    "Bitti": False,
                    "Not": f"Arşiv ({sheet})"
                }).dropna(subset=["Tarih", "Adet"])
                tum_gecmis.append(temp)
        except Exception:
            pass
            
    if tum_gecmis:
        return pd.concat(tum_gecmis).drop_duplicates("Tarih", keep="last")
    return pd.DataFrame(columns=["Tarih", "Adet", "Bitti", "Not"])


def eklenen_detay_oku():
    if os.path.exists(DETAY_EKLENEN_DOSYA):
        try:
            return pd.read_csv(DETAY_EKLENEN_DOSYA, parse_dates=["Tarih"])
        except Exception:
            pass
    return pd.DataFrame(columns=["Tarih", "Cesit", "Adet"])


def eklenen_gunluk_oku():
    detaylar = eklenen_detay_oku()
    if not detaylar.empty:
        g = detaylar.groupby("Tarih")["Adet"].sum().reset_index()
        g["Bitti"] = False
        g["Not"] = "Uygulamadan girildi"
        return g
    if os.path.exists(EKLENEN_DOSYA):
        try:
            e = pd.read_csv(EKLENEN_DOSYA)
            e["Tarih"] = _tarih(e["Tarih"])
            e["Bitti"] = e["Bitti"].map(_evet)
            e["Not"] = e["Not"].fillna("").astype(str)
            return e[["Tarih", "Adet", "Bitti", "Not"]]
        except Exception:
            pass
    return pd.DataFrame({"Tarih": pd.Series(dtype="datetime64[ns]"), "Adet": pd.Series(dtype=float),
                         "Bitti": pd.Series(dtype=bool), "Not": pd.Series(dtype=str)})


# ----------------------------------------------------------------------------
# HAVA DURUMU (Open-Meteo)
# ----------------------------------------------------------------------------
def _hava_df(js):
    g = js["daily"]
    return pd.DataFrame({
        "Tarih": pd.to_datetime(g["time"]).astype("datetime64[ns]"),
        "sicaklik": g["temperature_2m_mean"],
        "yagis": g["precipitation_sum"],
    })


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def hava_getir(lat, lon, bas_iso, bit_iso):
    parcalar = []
    bas = dt.date.fromisoformat(bas_iso)
    arsiv_son = min(dt.date.fromisoformat(bit_iso), BUGUN - dt.timedelta(days=6))
    daily = "temperature_2m_mean,precipitation_sum"
    try:
        if bas <= arsiv_son:
            q = urlencode({"latitude": lat, "longitude": lon, "start_date": bas.isoformat(),
                           "end_date": arsiv_son.isoformat(), "daily": daily, "timezone": "auto"})
            with urlopen("https://archive-api.open-meteo.com/v1/archive?" + q, timeout=20) as r:
                parcalar.append(_hava_df(json.load(r)))
    except Exception:
        pass
    try:
        q = urlencode({"latitude": lat, "longitude": lon, "daily": daily, "timezone": "auto",
                       "past_days": 10, "forecast_days": 16})
        with urlopen("https://api.open-meteo.com/v1/forecast?" + q, timeout=20) as r:
            parcalar.append(_hava_df(json.load(r)))
    except Exception:
        pass
    if not parcalar:
        return None
    h = pd.concat(parcalar).drop_duplicates("Tarih", keep="first")
    return h.sort_values("Tarih").reset_index(drop=True)


# ----------------------------------------------------------------------------
# ÖZELLİKLER VE MODEL (Ridge Regresyon)
# ----------------------------------------------------------------------------
def ozellik_olustur(df, cfg):
    d = pd.to_datetime(df["Tarih"])
    X = pd.DataFrame(index=df.index)
    for k in range(7):
        X[f"gun_{k}"] = (d.dt.dayofweek == k).astype(float)
    gun = d.dt.day
    X["ay_basi"] = (gun <= 5).astype(float)
    X["ay_ortasi"] = gun.between(10, 15).astype(float)
    X["ay_sonu"] = (gun >= 26).astype(float)
    doy = d.dt.dayofyear.to_numpy()
    for h in (1, 2):
        X[f"mevsim_sin{h}"] = np.sin(2 * np.pi * h * doy / 365.25)
        X[f"mevsim_cos{h}"] = np.cos(2 * np.pi * h * doy / 365.25)
    t = d.dt.date
    k = cfg["kum"]
    X["bayram"] = t.isin(k["bayram"]).astype(float)
    X["bayram_oncesi"] = t.isin(k["oncesi"]).astype(float)
    X["resmi_tatil"] = t.isin(k["resmi"]).astype(float)
    X["ozel_gun"] = t.isin(k["ozel"]).astype(float)
    X["trend"] = (d - cfg["t0"]).dt.days / 365.0
    X["menu_sonrasi"] = (d >= cfg["menu"]).astype(float) if cfg["menu"] is not None else 0.0
    if cfg["use_weather"]:
        X["sicaklik"] = df["sicaklik"].astype(float)
        X["yagmurlu"] = (df["yagis"].astype(float) >= 1.0).astype(float)
    return X


def fit_model(X, y_log, alpha=10.0):
    mu = X.mean()
    sd = X.std(ddof=0)
    sd = sd.where(sd > 1e-9, 1.0)
    Z = ((X - mu) / sd).to_numpy(float)
    y_log = np.asarray(y_log, float)
    yb = float(y_log.mean())
    A = Z.T @ Z + alpha * np.eye(Z.shape[1])
    w = np.linalg.solve(A, Z.T @ (y_log - yb))
    return {"w": w, "b": yb, "mu": mu, "sd": sd, "cols": list(X.columns)}


def predict_model(mod, X):
    Z = ((X[mod["cols"]] - mod["mu"]) / mod["sd"]).to_numpy(float)
    return np.exp(Z @ mod["w"] + mod["b"])


def baz_cizgi(v, hedef):
    gecmis = v[v["Tarih"] < hedef]
    ayni = gecmis[gecmis["Tarih"].dt.dayofweek == hedef.dayofweek].tail(6)["Y"]
    if len(ayni) >= 3:
        return float(ayni.mean())
    return float(gecmis.tail(28)["Y"].mean())


@st.cache_data(show_spinner="Geriye dönük test çalışıyor...")
def geriye_donuk_test(v, X, alphalar=(1.0, 3.0, 10.0, 30.0), test_gun=56, min_egitim=120):
    n = len(v)
    if n < min_egitim + 14:
        return None
    bas = max(min_egitim, n - test_gun)
    y = v["Y"].to_numpy(float)
    y_log = np.log(y)
    gercek = y[bas:]
    baz = np.array([baz_cizgi(v.iloc[:i], v["Tarih"].iloc[i]) for i in range(bas, n)])
    tahminler = {}
    for a in alphalar:
        tah = []
        for i in range(bas, n):
            mod = fit_model(X.iloc[:i], y_log[:i], a)
            tah.append(float(predict_model(mod, X.iloc[[i]])[0]))
        tahminler[a] = np.array(tah)
    mape_a = {a: mape(gercek, p) for a, p in tahminler.items()}
    en_iyi = min(mape_a, key=mape_a.get)
    model = tahminler[en_iyi]
    bt = pd.DataFrame({
        "Tarih": v["Tarih"].iloc[bas:].to_numpy(),
        "Gercek": gercek, "Baz": baz, "Model": model, "Karma": (baz + model) / 2,
    })
    return {
        "bt": bt, "alpha": en_iyi,
        "mape": {k: mape(bt["Gercek"], bt[k]) for k in ("Baz", "Model", "Karma")},
    }


def tahmin_et(cfg, veri, mod, yontem, tarih, sic, yag):
    satir = pd.DataFrame({"Tarih": [tarih], "sicaklik": [sic], "yagis": [yag]})
    xr = ozellik_olustur(satir, cfg)
    baz = baz_cizgi(veri, tarih)
    model = float(predict_model(mod, xr)[0]) if mod is not None else baz
    sec = {"Baz": baz, "Model": model, "Karma": (baz + model) / 2}[yontem]
    return sec, model, baz, xr


def aciklama(mod, xr):
    z = ((xr[mod["cols"]] - mod["mu"]) / mod["sd"]).iloc[0]
    katki = pd.Series(mod["w"], index=mod["cols"]) * z
    satirlar = []
    for grup, cols in GRUPLAR.items():
        c = katki[[x for x in cols if x in katki.index]].sum()
        pct = (np.exp(c) - 1) * 100
        if abs(pct) >= 1:
            satirlar.append((grup, float(pct)))
    return sorted(satirlar, key=lambda s: -abs(s[1]))


# ----------------------------------------------------------------------------
# ÜRÜN DAĞILIMI
# ----------------------------------------------------------------------------
def yuvarla_topla(degerler, toplam):
    taban = np.floor(degerler).astype(int)
    kalan = int(round(toplam)) - int(taban.sum())
    sira = np.argsort(-(degerler - taban))
    for i in sira[: max(kalan, 0)]:
        taban[i] += 1
    return taban


def urun_dagilimi(detay, hedef, toplam, pencere=60, k_cesit=300.0):
    if detay is None or detay.empty:
        return pd.DataFrame(), "Veri yok"
    son = detay["Tarih"].max()
    d = detay[detay["Tarih"] > son - pd.Timedelta(days=pencere)]
    hs = hedef.dayofweek >= 4
    maske = d["Tarih"].dt.dayofweek.ge(4) == hs
    cesitler = list(pd.unique(d["Cesit"]))

    genel = d.groupby("Cesit")["Adet"].sum().reindex(cesitler).fillna(0)
    genel_pay = genel / (genel.sum() if genel.sum() > 0 else 1)
    g = d[maske].groupby("Cesit")["Adet"].sum().reindex(cesitler).fillna(0)
    cesit_pay = (g + k_cesit * genel_pay) / (g.sum() + k_cesit)

    if cesit_pay.sum() > 0:
        cesit_pay = cesit_pay / cesit_pay.sum()
        
    sonuc = yuvarla_topla((cesit_pay * toplam).to_numpy(), toplam)
    
    tablo = pd.DataFrame({"Önerilen Üretim (Adet)": sonuc}, index=cesitler)
    tablo.loc["TOPLAM"] = tablo.sum()
    tablo.index.name = "Ürün Çeşidi"
    return tablo.astype(int), ("Cuma-Pazar" if hs else "Pazartesi-Perşembe")


# ----------------------------------------------------------------------------
# UYGULAMA ARAYÜZÜ (STREAMLIT)
# ----------------------------------------------------------------------------
def uygulama():
    st.title("🍰 Tatlıbak Yaş Pasta - Akıllı Üretim Planlama Motoru")

    with st.sidebar:
        st.header("⚙️ Ayarlar & Veri")
        lat = st.number_input("Enlem", value=41.01, format="%.2f")
        lon = st.number_input("Boylam", value=28.97, format="%.2f")
        hava_kullan = st.checkbox("Hava durumunu modele kat", value=True)
        bitti_duzeltme = st.slider("Ürün biten günlerde talep düzeltmesi (%)", 0, 30, 10)
        pencere = st.slider("Ürün dağılımı geriye bakış (gün)", 30, 120, 60)
        menu_var = st.checkbox("Menü değişikliği tarihini belirt")
        menu_tarihi = None
        if menu_var:
            menu_tarihi = pd.Timestamp(st.date_input("Yeni menünün başlangıcı", value=BUGUN - dt.timedelta(days=60)))
        ek_ozel = st.text_area("Ek özel günler (her satıra YYYY-AA-GG)", height=60)

    gecmis_df = gecmis_veri_oku()
    
    ad, ham = kaynak_bul(None, GUNLUK_DOSYALAR, ("2026", "satis", "pasta"))
    gunluk_2026 = pd.DataFrame(columns=["Tarih", "Adet", "Bitti", "Not"])
    detay_2026 = None
    
    if ad is not None:
        try:
            df_matris = tablo_oku(ad, ham)
            gunluk_2026, detay_2026 = matris_isle_2026(df_matris)
        except Exception as e:
            st.error(f"2026 satış dosyası okunamadı: {e}")
            
    ek_gunluk = eklenen_gunluk_oku()
    ek_detay = eklenen_detay_oku()
    
    if not ek_detay.empty:
        detay_2026 = pd.concat([detay_2026, ek_detay]).drop_duplicates(["Tarih", "Cesit"], keep="last")
        
    tum_veri_listesi = [gecmis_df, gunluk_2026, ek_gunluk]
    veri = pd.concat([df for df in tum_veri_listesi if not df.empty]).drop_duplicates("Tarih", keep="last").sort_values("Tarih").reset_index(drop=True)

    if len(veri) < 28:
        st.error(f"En az 28 günlük veri gerekli, şu an {len(veri)} gün var. Lütfen '2026 pasta satış.xlsx' dosyasının klasörde olduğundan emin olun.")
        return

    veri["Y"] = veri["Adet"] * np.where(veri["Bitti"], 1 + bitti_duzeltme / 100.0, 1.0)

    hava = None
    if hava_kullan:
        hava = hava_getir(round(lat, 2), round(lon, 2), veri["Tarih"].min().date().isoformat(),
                          (BUGUN + dt.timedelta(days=16)).isoformat())
    use_weather = False
    if hava is not None and veri["Tarih"].isin(hava["Tarih"]).mean() >= 0.7:
        veri = veri.merge(hava, on="Tarih", how="left")
        ay = veri["Tarih"].dt.month
        for c in ("sicaklik", "yagis"):
            veri[c] = veri[c].fillna(veri.groupby(ay)[c].transform("mean")).fillna(veri[c].mean())
        use_weather = True
    else:
        veri["sicaklik"] = np.nan
        veri["yagis"] = np.nan

    yillar = range(veri["Tarih"].min().year - 1, BUGUN.year + 3)
    cfg = {
        "kum": ozel_gunler(yillar, ek_ozel),
        "menu": menu_tarihi,
        "t0": veri["Tarih"].min(),
        "use_weather": use_weather,
    }
    X = ozellik_olustur(veri, cfg)

    test = geriye_donuk_test(veri[["Tarih", "Y"]], X)
    mod = None
    yontem = "Baz"
    if test is not None:
        yontem = min(test["mape"], key=test["mape].get)
        mod = fit_model(X, np.log(veri["Y"].to_numpy(float)), test["alpha"])

    def hava_degeri(tarih):
        if use_weather:
            s = hava[hava["Tarih"] == tarih]
            if len(s) and pd.notna(s["sicaklik"].iloc[0]):
                return float(s["sicaklik"].iloc[0]), float(s["yagis"].iloc[0] or 0.0), "tahmin"
            m = veri[veri["Tarih"].dt.month == tarih.month]
            return float(m["sicaklik"].mean()), 0.0, "mevsim ortalaması"
        return 15.0, 0.0, "yok"

    def carpan_hesapla(seviye):
        if test is None or len(test["bt"]) < 20:
            return 1.10, False
        oran = test["bt"]["Gercek"] / test["bt"][yontem]
        return float(np.quantile(oran, seviye)), True

    tab1, tab2, tab3, tab4 = st.tabs(["🚀 Üretim Planı", "📈 Model Performansı", "📝 Çeşit Bazlı Günlük Giriş", "📂 Veri Arşivi"])

    with tab1:
        c1, c2 = st.columns([1, 2])
        with c1:
            hedef_d = st.date_input("Üretim tarihi", value=BUGUN + dt.timedelta(days=1))
            hedef = pd.Timestamp(hedef_d)
            seviye_ad = st.selectbox("Güvenlik seviyesi", list(SEVIYELER), index=1)
            seviye = SEVIYELER[seviye_ad]
            hava_elle = st.checkbox("Hava durumunu elle gir")
            sic, yag, hava_kaynak = hava_degeri(hedef)
            if hava_elle:
                sic = st.slider("Sıcaklık (°C)", -5, 40, int(round(sic)))
                yag = 5.0 if st.checkbox("Yağmurlu", value=yag >= 1) else 0.0
                hava_kaynak = "elle"
            st.info(f"{GUN_TR[hedef.dayofweek]} | Hava: {sic:.0f}°C ({hava_kaynak})")

        sec, model_t, baz_t, xr = tahmin_et(cfg, veri, mod, yontem, hedef, sic, yag)
        carpan, olculu = carpan_hesapla(seviye)
        uretim = int(round(sec * carpan))

        with c2:
            m1, m2, m3 = st.columns(3)
            m1.metric("Önerilen toplam üretim", f"{uretim} adet")
            m2.metric("Beklenen talep", f"{int(round(sec))} adet")
            if olculu:
                oran = test["bt"]["Gercek"] / test["bt"][yontem]
                lo, hi = np.quantile(oran, 0.1) * sec, np.quantile(oran, 0.9) * sec
                m3.metric("Olası aralık (%80)", f"{int(round(lo))} - {int(round(hi))}")
            else:
                m3.metric("Olası aralık", "hesaplanamadı")
            
            # --- DETAYLI TAHMİN AÇIKLAMA RAPORU ---
            st.markdown("### 🧠 Akıllı Tahmin ve Karşılaştırma Analitiği")
            st.markdown(f"* **Kullanılan Yöntem:** `{YONTEM_AD[yontem]}` (Geçmiş veriler üzerinde yapılan geriye dönük testlerde en düşük hata oranını veren yöntem seçilmiştir).")
            st.markdown(f"* **Veri Havuzu Temeli:** Model, geçmiş yılların arşivleri ve 2026 dönemi dahil toplam **{len(veri)} günlük** gerçekleşen satış verisini taramaktadır.")
            
            # Baz çizgi ve model karşılaştırması
            st.markdown(f"* **Baz Çizgi (Son 6 {GUN_TR[hedef.dayofweek]} Ortalaması):** `{int(round(baz_t))}` adet")
            if mod is not None:
                st.markdown(f"* **İstatistiksel Regresyon Modeli:** `{int(round(model_t))}` adet")
            
            if olculu:
                mape_val = test['mape'][yontem] * 100
                st.markdown(f"* **Güvenlik Çarpanı & Hata Payı:** Modelin geçmiş testlerdeki ortalama yüzde hatası (MAPE) **%{mape_val:.1f}** olarak ölçülmüştür. Seçilen güvenlik seviyesi (`{seviye_ad}`) ile bu hata payı dengelenerek son üretim adedine ulaşılmıştır.")

            if mod is not None:
                st.markdown("**Neden Bu Talep Bekleniyor? (Etken Faktörler):**")
                etki = aciklama(mod, xr)
                if etki:
                    for grup, pct in etki[:4]:
                        st.markdown(f"  - {grup}: **{pct:+.0f}%** etki")
                else:
                    st.markdown("  - Belirgin ekstrem bir etken yok, tipik sezon ortalaması.")

        st.markdown("### 📋 Çeşit Dağılımı")
        if detay_2026 is None or detay_2026.empty:
            st.info("Detay ürün dağılımı için '2026 pasta satış.xlsx' dosyası kullanılıyor.")
        else:
            tablo, grup_ad = urun_dagilimi(detay_2026, hedef, uretim, pencere)
            st.dataframe(tablo, use_container_width=True)
            st.caption(f"Pay dağılımı son {pencere} günlük satışlardan ({grup_ad} günleri) hesaplanmıştır. Toplam = {uretim} adet.")

        st.markdown("### 🗓️ Önümüzdeki 7 gün")
        satirlar = []
        for i in range(7):
            td = hedef + pd.Timedelta(days=i)
            s_i, y_i, _ = hava_degeri(td)
            tahmin_i = tahmin_et(cfg, veri, mod, yontem, td, s_i, y_i)[0]
            satirlar.append({
                "Tarih": td.strftime("%d.%m.%Y"), "Gün": GUN_TR[td.dayofweek],
                "Hava": f"{s_i:.0f}°C", "Beklenen talep": int(round(tahmin_i)), "Önerilen üretim": int(round(tahmin_i * carpan)),
            })
        st.dataframe(pd.DataFrame(satirlar), hide_index=True)

    with tab2:
        st.subheader("Model Geriye Dönük Test (Backtest) Performansı")
        if test is not None:
            mm = test["mape"]
            a, b, c = st.columns(3)
            a.metric("Baz çizgi hatası", f"%{mm['Baz'] * 100:.1f}")
            b.metric("Model hatası", f"%{mm['Model'] * 100:.1f}")
            c.metric("Karma hatası", f"%{mm['Karma'] * 100:.1f}")
            st.success(f"Seçilen en iyi yöntem: **{YONTEM_AD[yontem]}**")
            st.line_chart(test["bt"].rename(columns={"Gercek": "Gerçek", "Baz": "Baz çizgi"}).set_index("Tarih")[["Gerçek", "Model"]])

    with tab3:
        st.subheader("📝 Çeşit Bazlı Günlük Satış Girişi")
        st.caption("Her akşam o gün üretilen/satılan çeşitleri adet bazında girin. Bu veriler gelecekteki tahminleri ve ürün dağılımını besleyecektir.")
        
        # Referans ürün listesini bulalım
        ornek_urunler = []
        if detay_2026 is not None and not detay_2026.empty:
            ornek_urunler = sorted(detay_2026["Cesit"].unique().tolist())
        else:
            ornek_urunler = ["BATON ÇİKOLATALI", "BÜYÜK TURTA ÇİKOLATALI (ÇİKOLATA ÇİLEK OREO)", "BÜYÜK TURTA ÇİKOLATALI (ÇİKOLATA MUZLU)"]

        with st.form("detay_giris_form"):
            g_tarih = st.date_input("Satış Tarihi", value=BUGUN)
            st.markdown("---")
            
            girilen_degerler = {}
            # İki sütunlu şık bir form yapısı
            cols = st.columns(2)
            for idx, urun in enumerate(ornek_urunler):
                col_target = cols[idx % 2]
                girilen_degerler[urun] = col_target.number_input(f"{urun}", min_value=0, value=0, step=1, key=f"inp_{urun}")
                
            st.markdown("---")
            kaydet_detay = st.form_submit_button("💾 Çeşit Bazlı Satışları Kaydet")
            
        if kaydet_detay:
            yeni_satirlar = []
            toplam_adet_hesap = 0
            for urun, adet in girilen_degerler.items():
                if adet > 0:
                    yeni_satirlar.append({"Tarih": pd.Timestamp(g_tarih), "Cesit": urun, "Adet": adet})
                    toplam_adet_hesap += adet
                    
            if toplam_adet_hesap > 0:
                yeni_df = pd.DataFrame(yeni_satirlar)
                mevcut_detay = eklenen_detay_oku()
                birlesik_detay = pd.concat([mevcut_detay, yeni_df]).drop_duplicates(["Tarih", "Cesit"], keep="last")
                birlesik_detay.to_csv(DETAY_EKLENEN_DOSYA, index=False)
                st.success(f"✅ {g_tarih:%d.%m.%Y} tarihli satışlar başarıyla kaydedildi! Toplam Satış: {toplam_adet_hesap} adet.")
                st.rerun()
            else:
                st.warning("Lütfen en az bir ürün için 0'dan büyük adet girin.")

    with tab4:
        st.subheader("Birleştirilmiş Veri Havuzu")
        st.caption(f"Toplam kayıtlı gün sayısı: {len(veri)}")
        goster = veri.copy()
        goster["Tarih"] = goster["Tarih"].dt.strftime("%d.%m.%Y")
        st.dataframe(goster.iloc[::-1], hide_index=True)


try:
    uygulama()
except Exception as ex:
    st.error(f"Hata: {ex}")
    st.exception(ex)