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
# AYARLAR
# ----------------------------------------------------------------------------
GUNLUK_DOSYALAR = ["gunluk_satis.xlsx", "gunluk_satis.csv"]
DETAY_DOSYALAR = ["urun_detay.xlsx", "urun_detay.csv"]
EKLENEN_DOSYA = "eklenen_gunluk.csv"  # uygulamadan girilen günlük satışlar

GUN_TR = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

# Dini bayramlar (başlangıç, bitiş - dahil). Yeni yıl için buraya ekleyin.
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
        for i in range(1, 4):  # arife dahil bayramdan önceki 3 gün
            oncesi.add(bs - dt.timedelta(days=i))
    for y in yillar:
        for m, g in [(1, 1), (4, 23), (5, 1), (5, 19), (7, 15), (8, 30), (10, 29)]:
            resmi.add(dt.date(y, m, g))
        for m, g in [(2, 14), (3, 8), (11, 24), (12, 31)]:
            ozel.add(dt.date(y, m, g))  # Sevgililer, Kadınlar, Öğretmenler, Yılbaşı arifesi
        ozel.add(nth_weekday(y, 5, 6, 2))  # Anneler Günü
        ozel.add(nth_weekday(y, 6, 6, 3))  # Babalar Günü
    for satir in (ek_metin or "").splitlines():
        try:
            ozel.add(dt.date.fromisoformat(satir.strip()))
        except ValueError:
            pass
    return {"bayram": bayram, "oncesi": oncesi, "resmi": resmi, "ozel": ozel}


# ----------------------------------------------------------------------------
# VERİ OKUMA
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
    return xls.parse(xls.sheet_names[0])


def kaynak_bul(yuklenen, adaylar, anahtarlar, haric=()):
    """Yüklenen dosya > tam isim eşleşmesi > klasördeki benzer isimli dosya."""
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


def gunluk_hazirla(ham):
    n = {c: _norm(c) for c in ham.columns}
    tarih_c = next((c for c, x in n.items() if "tarih" in x), ham.columns[0])
    bitti_c = next((c for c, x in n.items() if "bitti" in x or "tukendi" in x), None)
    not_c = next((c for c, x in n.items() if x.startswith("not") or "aciklama" in x), None)
    dis = {tarih_c, bitti_c, not_c}
    top_c = next(
        (c for c, x in n.items() if c not in dis and any(k in x for k in ("toplam", "adet", "satis"))),
        None,
    )
    if top_c is None:
        aday = [c for c in ham.columns if c not in dis]
        if not aday:
            raise ValueError(f"Toplam adet sütunu bulunamadı. Bulunan sütunlar: {list(ham.columns)}")
        top_c = aday[0]
    out = pd.DataFrame({
        "Tarih": _tarih(ham[tarih_c]),
        "Adet": ham[top_c].map(_sayi),
        "Bitti": ham[bitti_c].map(_evet) if bitti_c else False,
        "Not": ham[not_c].fillna("").astype(str) if not_c else "",
    })
    out = out.dropna(subset=["Tarih", "Adet"])
    out = out[out["Adet"] > 0]
    out["Bitti"] = out["Bitti"].astype(bool)
    if out.empty:
        raise ValueError(
            "Geçerli satır bulunamadı. Tarih ve adet sütunlarını kontrol edin. "
            f"Bulunan sütunlar: {list(ham.columns)}"
        )
    return out.drop_duplicates("Tarih", keep="last").sort_values("Tarih").reset_index(drop=True)


def detay_hazirla(ham):
    n = {c: _norm(c) for c in ham.columns}
    tarih_c = next((c for c, x in n.items() if "tarih" in x), ham.columns[0])
    boy_c = next((c for c, x in n.items() if "boy" in x), None)
    adet_c = next((c for c, x in n.items() if "adet" in x), None)
    cesit_c = next(
        (c for c, x in n.items() if c not in (tarih_c, boy_c, adet_c) and any(k in x for k in ("cesit", "urun", "pasta"))),
        None,
    )
    if None in (boy_c, adet_c, cesit_c):
        raise ValueError(
            "Ürün detay dosyasında Tarih, Çeşit, Boy, Adet sütunları bulunamadı. "
            f"Bulunan sütunlar: {list(ham.columns)}"
        )
    out = pd.DataFrame({
        "Tarih": _tarih(ham[tarih_c]),
        "Cesit": ham[cesit_c].astype(str).str.strip(),
        "Boy": ham[boy_c].astype(str).str.strip(),
        "Adet": ham[adet_c].map(_sayi),
    }).dropna()
    out = out[out["Adet"] > 0].reset_index(drop=True)
    if out.empty:
        raise ValueError("Ürün detay dosyasında geçerli satır yok.")
    return out


def eklenen_oku():
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
# ÖRNEK VERİ VE ŞABLON (dosyanız hazır değilken uygulamayı denemek için)
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def ornek_veri():
    rng = np.random.default_rng(7)
    gunler = pd.date_range(BUGUN - dt.timedelta(days=730), BUGUN - dt.timedelta(days=1))
    dow = np.array([0.95, 0.95, 1.0, 1.0, 1.1, 1.2, 1.1])
    y, bitti = [], []
    for d in gunler:
        v = 1300 * dow[d.dayofweek] * (1.05 if d.month in (11, 12, 1, 2) else 0.95)
        v *= 1.06 if 10 <= d.day <= 15 else 1.0
        y.append(int(v * rng.normal(1, 0.07)))
        bitti.append("Evet" if rng.random() < 0.08 else "Hayır")
    g = pd.DataFrame({"Tarih": gunler, "Toplam Adet": y, "Ürün Bitti mi": bitti, "Not": ""})
    satir = []
    for d in gunler[-60:]:
        for c in ["Fıstıklı", "Çikolatalı", "Meyveli", "Frambuazlı", "Karamelli", "Muzlu", "Limonlu", "Kakaolu"]:
            for b, s in zip(["Küçük", "Orta", "Büyük"], [0.5, 0.35, 0.15]):
                satir.append([d, c, b, int(rng.integers(2, 12) * s * 3) + 1])
    return g, pd.DataFrame(satir, columns=["Tarih", "Çeşit", "Boy", "Adet"])


def sablon_xlsx(df):
    b = io.BytesIO()
    df.to_excel(b, index=False)
    return b.getvalue()


# ----------------------------------------------------------------------------
# HAVA DURUMU (Open-Meteo, ücretsiz, anahtar gerektirmez)
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
# ÖZELLİKLER VE MODEL (yalnız numpy; ek kütüphane gerekmez)
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


def fit_model(X, y_log, alpha):
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
    """Aynı gün tipinin son 6 değerinin ortalaması."""
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


def urun_dagilimi(detay, hedef, toplam, pencere, k_cesit=300.0, k_boy=100.0):
    son = detay["Tarih"].max()
    d = detay[detay["Tarih"] > son - pd.Timedelta(days=pencere)]
    hs = hedef.dayofweek >= 4  # Cuma-Cumartesi-Pazar grubu
    maske = d["Tarih"].dt.dayofweek.ge(4) == hs
    cesitler = list(pd.unique(d["Cesit"]))
    boylar = list(pd.unique(d["Boy"]))

    genel = d.groupby("Cesit")["Adet"].sum().reindex(cesitler)
    genel_pay = genel / genel.sum()
    g = d[maske].groupby("Cesit")["Adet"].sum().reindex(cesitler).fillna(0)
    cesit_pay = (g + k_cesit * genel_pay) / (g.sum() + k_cesit)

    pay = np.zeros((len(cesitler), len(boylar)))
    for i, c in enumerate(cesitler):
        dc = d[d["Cesit"] == c]
        gb = dc.groupby("Boy")["Adet"].sum().reindex(boylar).fillna(0)
        gb_pay = gb / gb.sum() if gb.sum() > 0 else pd.Series(1 / len(boylar), index=boylar)
        gh = dc[dc["Tarih"].dt.dayofweek.ge(4) == hs].groupby("Boy")["Adet"].sum().reindex(boylar).fillna(0)
        boy_pay = (gh + k_boy * gb_pay) / (gh.sum() + k_boy)
        pay[i, :] = cesit_pay[c] * boy_pay.to_numpy()
    pay = pay / pay.sum()
    sonuc = yuvarla_topla((pay * toplam).ravel(), toplam).reshape(pay.shape)
    tablo = pd.DataFrame(sonuc, index=cesitler, columns=boylar)
    tablo["Toplam"] = tablo.sum(axis=1)
    tablo.loc["TOPLAM"] = tablo.sum()
    tablo.index.name = "Çeşit"
    return tablo.astype(int), ("Cuma-Pazar" if hs else "Pazartesi-Perşembe")


# ----------------------------------------------------------------------------
# UYGULAMA
# ----------------------------------------------------------------------------
def uygulama():
    st.title("🍰 Tatlıbak Yaş Pasta - Üretim Planlama")

    # ---------------- Yan menü ----------------
    with st.sidebar:
        st.header("⚙️ Ayarlar")
        ornek = st.checkbox("Örnek veriyle dene (gerçek veri değil)", value=False,
                            help="Kendi dosyalarınız hazır değilken uygulamayı denemek içindir.")
        st.caption("Veri dosyaları (isteğe bağlı; yüklenmezse GitHub'daki dosyalar okunur)")
        yuk_gunluk = st.file_uploader("Günlük toplam satış (xlsx/csv)", type=["xlsx", "csv"], key="u1")
        yuk_detay = st.file_uploader("Ürün detay (xlsx/csv)", type=["xlsx", "csv"], key="u2")
        st.markdown("---")
        lat = st.number_input("Enlem", value=41.01, format="%.2f", help="Firmanızın bulunduğu şehir. Varsayılan: İstanbul")
        lon = st.number_input("Boylam", value=28.97, format="%.2f")
        hava_kullan = st.checkbox("Hava durumunu modele kat", value=True)
        bitti_duzeltme = st.slider(
            "Ürün biten günlerde talep düzeltmesi (%)", 0, 30, 10,
            help="Ürün tükenen günün satışı gerçek talepten düşüktür. Bu günlerin değeri bu oranda yukarı çekilir.",
        )
        pencere = st.slider("Ürün dağılımı için geriye bakış (gün)", 30, 120, 60)
        menu_var = st.checkbox("Menü değişikliği tarihini belirt")
        menu_tarihi = None
        if menu_var:
            menu_tarihi = pd.Timestamp(st.date_input("Yeni menünün başlangıcı", value=BUGUN - dt.timedelta(days=60)))
        ek_ozel = st.text_area("Ek özel günler / kampanyalar (her satıra YYYY-AA-GG)", height=80)
        with st.expander("Sorun giderme"):
            st.write(f"Python {sys.version.split()[0]} | pandas {pd.__version__} | numpy {np.__version__} | streamlit {st.__version__}")
            st.write("Klasördeki dosyalar:", sorted(glob.glob("*"))[:40])

    # ---------------- Veriyi yükle ----------------
    if ornek:
        g_ham, d_ham = ornek_veri()
        gunluk = gunluk_hazirla(g_ham)
        detay = detay_hazirla(d_ham)
        st.warning("ÖRNEK VERİ kullanılıyor. Sayılar gerçek değil, yalnız uygulamayı denemek içindir.")
    else:
        ad, ham = kaynak_bul(yuk_gunluk, GUNLUK_DOSYALAR, ("gunluk", "satis"), ("detay", "urun", "eklenen"))
        if ad is None:
            st.error("Günlük satış dosyası bulunamadı.")
            st.markdown(
                "GitHub'daki `app.py` ile **aynı klasöre** `gunluk_satis.xlsx` dosyasını ekleyin "
                "ya da soldan yükleyin. Gerekli sütunlar: **Tarih | Toplam Adet | Ürün Bitti mi | Not** "
                "(son ikisi isteğe bağlı).\n\n"
                "Dosyanız hazır değilse soldaki **Örnek veriyle dene** kutusunu işaretleyin."
            )
            c1, c2 = st.columns(2)
            g_ham, d_ham = ornek_veri()
            c1.download_button(
                "⬇️ Günlük satış şablonu (xlsx)",
                sablon_xlsx(g_ham.head(5).assign(Tarih=lambda x: x["Tarih"].dt.strftime("%Y-%m-%d"))),
                file_name="gunluk_satis_sablon.xlsx",
            )
            c2.download_button(
                "⬇️ Ürün detay şablonu (xlsx)",
                sablon_xlsx(d_ham.head(5).assign(Tarih=lambda x: x["Tarih"].dt.strftime("%Y-%m-%d"))),
                file_name="urun_detay_sablon.xlsx",
            )
            return
        try:
            gunluk = gunluk_hazirla(tablo_oku(ad, ham))
        except Exception as ex:
            st.error(f"Günlük satış dosyası ({ad}) okunamadı: {ex}")
            return
        eklenen = eklenen_oku()
        if len(eklenen):
            gunluk = pd.concat([gunluk, eklenen]).drop_duplicates("Tarih", keep="last").sort_values("Tarih").reset_index(drop=True)
        detay = None
        dad, dham = kaynak_bul(yuk_detay, DETAY_DOSYALAR, ("detay", "urun"))
        if dad is not None:
            try:
                detay = detay_hazirla(tablo_oku(dad, dham))
            except Exception as ex:
                st.sidebar.warning(f"Ürün detay dosyası ({dad}) okunamadı: {ex}")
        st.caption(f"Okunan dosya: {ad}" + (f" | ürün detay: {dad}" if dad else ""))

    if len(gunluk) < 28:
        st.error(f"En az 28 günlük veri gerekli, şu an {len(gunluk)} gün var.")
        return

    veri = gunluk.copy()
    veri["Y"] = veri["Adet"] * np.where(veri["Bitti"], 1 + bitti_duzeltme / 100.0, 1.0)

    # Hava durumu
    hava = None
    if hava_kullan:
        hava = hava_getir(round(lat, 2), round(lon, 2), veri["Tarih"].min().date().isoformat(),
                          (BUGUN + dt.timedelta(days=16)).isoformat())
    use_weather = False
    if hava is not None and veri["Tarih"].isin(hava["Tarih"]).mean() >= 0.8:
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
        yontem = min(test["mape"], key=test["mape"].get)
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

    # ---------------- Sekmeler ----------------
    tab1, tab2, tab3, tab4 = st.tabs(["🚀 Üretim Planı", "📈 Model Performansı", "📝 Günlük Giriş", "📂 Veri"])

    # ---- Sekme 1: Plan ----
    with tab1:
        c1, c2 = st.columns([1, 2])
        with c1:
            hedef_d = st.date_input("Üretim tarihi", value=BUGUN + dt.timedelta(days=1))
            hedef = pd.Timestamp(hedef_d)
            seviye_ad = st.selectbox("Güvenlik seviyesi", list(SEVIYELER), index=1,
                                     help="Talebin üretimi aşma olasılığını belirler. Yüksek seviye = daha az kaçan satış, daha çok fire.")
            seviye = SEVIYELER[seviye_ad]
            hava_elle = st.checkbox("Hava durumunu elle gir")
            sic, yag, hava_kaynak = hava_degeri(hedef)
            if hava_elle:
                sic = st.slider("Sıcaklık (°C)", -5, 40, int(round(sic)))
                yag = 5.0 if st.checkbox("Yağmurlu", value=yag >= 1) else 0.0
                hava_kaynak = "elle"
            st.info(f"{GUN_TR[hedef.dayofweek]} | Hava: {sic:.0f}°C, {'yağmurlu' if yag >= 1 else 'yağışsız'} ({hava_kaynak})")

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
            aciklama_metni = (
                f"Üretim = beklenen talep × {carpan:.2f} (geçmiş hatalardan ölçüldü)." if olculu
                else "Yeterli test verisi yok; geçici olarak %10 pay eklendi."
            )
            st.caption(f"Yöntem: **{YONTEM_AD[yontem]}**. {aciklama_metni}")

            ayni = veri[(veri["Tarih"] < hedef) & (veri["Tarih"].dt.dayofweek == hedef.dayofweek)].tail(6)
            if len(ayni):
                st.markdown(
                    f"**Son {len(ayni)} {GUN_TR[hedef.dayofweek]}:** "
                    + ", ".join(str(int(x)) for x in ayni["Adet"])
                    + f"  (ortalama {int(round(ayni['Adet'].mean()))})"
                )

            if mod is not None:
                st.markdown("**Neden bu tahmin?** (ortalama güne göre etkiler)")
                etki = aciklama(mod, xr)
                if etki:
                    for grup, pct in etki:
                        st.markdown(f"- {grup}: **{pct:+.0f}%**")
                else:
                    st.markdown("- Belirgin bir etken yok, ortalama gün.")

            st.markdown("### ⚠️ Uyarılar")
            uyarilar = []
            son_veri = veri["Tarih"].max()
            bosluk = (hedef - son_veri).days
            if bosluk > 2:
                uyarilar.append(("warning", f"Son veri {son_veri:%d.%m.%Y}. Arada {bosluk - 1} gün girilmemiş, tahmin bayat olabilir."))
            if test is None:
                uyarilar.append(("info", f"Elinizde {len(veri)} günlük veri var. Model için en az ~135 gün gerekir; şimdilik yalnız baz çizgi kullanılıyor."))
            else:
                bt = test["bt"]
                son28 = bt.tail(28)
                m28 = mape(son28["Gercek"], son28[yontem])
                if m28 > 0.15:
                    uyarilar.append(("warning", f"Son 4 haftada ortalama hata %{m28 * 100:.0f}. Tahmine temkinli yaklaşın."))
                b7 = float((bt.tail(7)["Gercek"] / bt.tail(7)[yontem]).mean())
                if b7 > 1.10:
                    uyarilar.append(("warning", f"Model son 7 günde ortalama %{(b7 - 1) * 100:.0f} düşük tahmin etti. Talep yükseliyor olabilir."))
                elif b7 < 0.90:
                    uyarilar.append(("warning", f"Model son 7 günde ortalama %{(1 - b7) * 100:.0f} yüksek tahmin etti. Talep düşüyor olabilir."))
            bitti14 = veri.tail(14)["Bitti"].mean()
            if bitti14 > 0.2:
                uyarilar.append(("warning", f"Son 14 günün %{bitti14 * 100:.0f}'inde ürün tükenmiş. Güvenlik seviyesini yükseltmeyi düşünün."))
            if baz_t > 0 and abs(sec / baz_t - 1) > 0.25:
                uyarilar.append(("warning", f"Tahmin, son haftaların aynı gün ortalamasından %{abs(sec / baz_t - 1) * 100:.0f} farklı. Takvim/hava etkisini kontrol edin."))
            if sec > veri["Adet"].max():
                uyarilar.append(("info", "Tahmin, geçmişteki en yüksek günün üzerinde."))
            t = hedef.date()
            k = cfg["kum"]
            if t in k["bayram"]:
                uyarilar.append(("info", "Seçilen gün bayram günü."))
            elif t in k["oncesi"]:
                uyarilar.append(("info", "Bayram öncesi (arife dahil) dönem."))
            if t in k["ozel"]:
                uyarilar.append(("info", "Özel gün (Anneler Günü, Sevgililer Günü, ek günleriniz vb.)."))
            if hava_kullan and not use_weather:
                uyarilar.append(("info", "Hava durumu verisi alınamadı, model havasız çalışıyor."))
            if not uyarilar:
                st.success("Dikkat çeken bir durum yok.")
            for seviye_u, metin in uyarilar:
                (st.warning if seviye_u == "warning" else st.info)(metin)

        st.markdown("### 📋 Çeşit / Boy Dağılımı")
        if detay is None or detay.empty:
            st.info("Ürün detay dosyası (`urun_detay.xlsx`) yok. Eklerseniz çeşit ve boy dağılımı da hesaplanır.")
        else:
            tablo, grup_ad = urun_dagilimi(detay, hedef, uretim, pencere)
            st.dataframe(tablo)
            st.caption(f"Pay, son {pencere} günlük satışlardan ({grup_ad} günlerine göre) hesaplandı. Toplam = {uretim} adet.")

        st.markdown("### 🗓️ Önümüzdeki 7 gün")
        satirlar = []
        for i in range(7):
            td = hedef + pd.Timedelta(days=i)
            s_i, y_i, _ = hava_degeri(td)
            tahmin_i = tahmin_et(cfg, veri, mod, yontem, td, s_i, y_i)[0]
            satirlar.append({
                "Tarih": td.strftime("%d.%m.%Y"), "Gün": GUN_TR[td.dayofweek],
                "Hava": f"{s_i:.0f}°C" + (" 🌧" if y_i >= 1 else ""),
                "Beklenen talep": int(round(tahmin_i)), "Önerilen üretim": int(round(tahmin_i * carpan)),
            })
        st.dataframe(pd.DataFrame(satirlar), hide_index=True)

    # ---- Sekme 2: Performans ----
    with tab2:
        st.subheader("Model kendini nasıl sınadı?")
        st.caption(
            "Her gün için, o güne kadarki veriyle tahmin yapılıp gerçek satışla karşılaştırıldı. "
            "MAPE = ortalama yüzde hata (düşük iyi). Geçmiş hava durumu gerçek değerlerle kullanıldığı için "
            "gerçek hayatta hata biraz daha yüksek olabilir."
        )
        if test is None:
            st.info(f"Test için en az ~135 günlük veri gerekli. Şu an {len(veri)} gün var.")
        else:
            mm = test["mape"]
            a, b, c = st.columns(3)
            a.metric("Baz çizgi hatası", f"%{mm['Baz'] * 100:.1f}")
            b.metric("Model hatası", f"%{mm['Model'] * 100:.1f}")
            c.metric("Karma hatası", f"%{mm['Karma'] * 100:.1f}")
            st.success(f"Seçilen yöntem: **{YONTEM_AD[yontem]}** (en düşük hata). Test edilen gün sayısı: {len(test['bt'])}.")
            g = test["bt"].rename(columns={"Gercek": "Gerçek", "Baz": "Baz çizgi"}).set_index("Tarih")
            st.line_chart(g[["Gerçek", "Baz çizgi", "Model", "Karma"]])
            son = test["bt"].tail(14).copy()
            son["Seçilen tahmin"] = son[yontem].round(0).astype(int)
            son["Sapma %"] = ((son["Seçilen tahmin"] / son["Gercek"] - 1) * 100).round(1)
            son["Tarih"] = pd.to_datetime(son["Tarih"]).dt.strftime("%d.%m.%Y")
            son["Gerçek"] = son["Gercek"].round(0).astype(int)
            st.markdown("**Son 14 gün: tahmin ve gerçek**")
            st.dataframe(son[["Tarih", "Gerçek", "Seçilen tahmin", "Sapma %"]], hide_index=True)

    # ---- Sekme 3: Günlük giriş ----
    with tab3:
        st.subheader("Günlük satış girişi")
        if ornek:
            st.info("Örnek veri modunda kayıt kapalı. Kendi verinizle çalışırken bu bölüm açılır.")
        else:
            st.caption("Her akşam o günün toplam satışını girin. Model bir sonraki hesaplamada bunu kullanır.")
            with st.form("giris", clear_on_submit=True):
                g_tarih = st.date_input("Tarih", value=BUGUN)
                g_adet = st.number_input("Toplam satış adedi", min_value=0, step=1)
                g_bitti = st.checkbox("Gün içinde ürün tükendi")
                g_not = st.text_input("Not (kampanya, kapalı gün vb.)")
                kaydet = st.form_submit_button("Kaydet")
            if kaydet:
                if g_adet <= 0:
                    st.error("Satış adedi 0'dan büyük olmalı.")
                else:
                    yeni = pd.DataFrame({"Tarih": [pd.Timestamp(g_tarih)], "Adet": [int(g_adet)],
                                         "Bitti": [bool(g_bitti)], "Not": [g_not]})
                    birlesik = pd.concat([eklenen_oku(), yeni]).drop_duplicates("Tarih", keep="last")
                    birlesik.to_csv(EKLENEN_DOSYA, index=False)
                    st.success(f"{g_tarih:%d.%m.%Y} kaydedildi.")
                    st.rerun()
            st.warning(
                "Streamlit Cloud'da bu giriş uygulama yeniden başlayınca silinebilir. "
                "Düzenli olarak **Veri** sekmesinden güncel dosyayı indirip GitHub'daki `gunluk_satis` dosyasının yerine yükleyin."
            )

    # ---- Sekme 4: Veri ----
    with tab4:
        st.subheader("Günlük satış verisi")
        gun_aralik = (veri["Tarih"].max() - veri["Tarih"].min()).days + 1
        eksik = gun_aralik - len(veri)
        c1, c2, c3 = st.columns(3)
        c1.metric("Kayıtlı gün", len(veri))
        c2.metric("İlk - son tarih", f"{veri['Tarih'].min():%d.%m.%y} - {veri['Tarih'].max():%d.%m.%y}")
        c3.metric("Aradaki eksik gün", eksik)
        if eksik > 0:
            st.caption("Eksik günler (kapalı gün ya da girilmemiş) modele dahil edilmez; sorun değil ama çok fazlaysa kontrol edin.")
        goster = gunluk.copy()
        goster["Tarih"] = goster["Tarih"].dt.strftime("%d.%m.%Y")
        goster["Bitti"] = goster["Bitti"].map({True: "Evet", False: "Hayır"})
        st.dataframe(goster.iloc[::-1].rename(columns={"Adet": "Toplam Adet", "Bitti": "Ürün Bitti mi"}), hide_index=True)
        indir = gunluk.rename(columns={"Adet": "Toplam Adet", "Bitti": "Ürün Bitti mi"}).copy()
        indir["Ürün Bitti mi"] = indir["Ürün Bitti mi"].map({True: "Evet", False: "Hayır"})
        st.download_button("⬇️ Güncel veriyi indir (CSV)", indir.to_csv(index=False).encode("utf-8-sig"),
                           file_name="gunluk_satis.csv", mime="text/csv")
        if detay is not None:
            st.subheader("Ürün detay verisi")
            st.caption(f"{detay['Tarih'].min():%d.%m.%Y} - {detay['Tarih'].max():%d.%m.%Y}, "
                       f"{detay['Cesit'].nunique()} çeşit, {detay['Boy'].nunique()} boy")
            st.dataframe(detay.tail(200), hide_index=True)


try:
    uygulama()
except Exception as ex:
    st.error(f"Beklenmeyen hata: {type(ex).__name__}: {ex}")
    st.exception(ex)
