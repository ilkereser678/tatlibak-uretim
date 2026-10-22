# -*- coding: utf-8 -*-
"""
Tatlıbak Yaş Pasta - Akıllı Üretim Planlama (app.py) - v2: kendini her gün geliştiren sürüm

Kurulum:   pip install streamlit pandas numpy openpyxl
Çalıştır:  streamlit run app.py
"""
import datetime as dt
import glob
import io
import json
import os
import re
import sqlite3
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
# AYARLAR VE SABİTLER
# ----------------------------------------------------------------------------
GUNLUK_DOSYALAR = ["2026 pasta satış.xlsx", "gunluk_satis.xlsx", "gunluk_satis.csv"]
GECMIS_DOSYALAR = ["geçmiş yıllar data.xlsx", "gecmis_arsiv.xlsx"]
DB_YOLU = os.environ.get("TATLIBAK_DB", "tatlibak.db")
ESKI_DETAY_CSV = "eklenen_detay.csv"
ESKI_GUNLUK_CSV = "eklenen_gunluk.csv"

VARSAYILAN_HARIC = ["BATON ÇİKOLATALI"]

GUN_TR = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
AY_TR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
         "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]

BAYRAMLAR = [
    ("2023-04-21", "2023-04-23"), ("2023-06-28", "2023-07-01"),
    ("2024-04-10", "2024-04-12"), ("2024-06-16", "2024-06-19"),
    ("2025-03-30", "2025-04-01"), ("2025-06-06", "2025-06-09"),
    ("2026-03-20", "2026-03-22"), ("2026-05-27", "2026-05-30"),
    ("2027-03-09", "2027-03-11"), ("2027-05-16", "2027-05-19"),
]

OLAY_ETIKET = {
    "bayram": "Bayram",
    "bayram_oncesi": "Bayram öncesi",
    "bayram_sonrasi": "Bayram sonrası",
    "resmi_tatil": "Resmi tatil",
    "ozel_sevgililer": "Sevgililer Günü",
    "ozel_kadinlar": "Kadınlar Günü",
    "ozel_anneler": "Anneler Günü",
    "ozel_babalar": "Babalar Günü",
    "ozel_ogretmenler": "Öğretmenler Günü",
    "ozel_yilbasi": "Yılbaşı",
    "ozel_cumhuriyet": "Cumhuriyet Bayramı",
    "ozel_diger": "Özel gün (elle eklenen)",
}

GRUPLAR = {
    "Gün tipi": [f"gun_{k}" for k in range(7)],
    "Maaş / ay günü": ["ay_basi", "ay_ortasi", "ay_sonu"],
    "Mevsim": ["mevsim_sin1", "mevsim_cos1", "mevsim_sin2", "mevsim_cos2"],
    "Bayram / özel gün": list(OLAY_ETIKET.keys()),
    "Hava": ["sicaklik", "yagmurlu", "sic_sapma"],
    "Trend / menü": ["trend", "menu_sonrasi"],
    "Geçmiş seviye": ["son7", "lag7"],
}

MALIYET_ANAHTAR = "Maliyete göre (fire / kaçan satış TL)"
SEVIYELER = {
    "Dengeli (%50) - fire az, kaçan satış olabilir": 0.5,
    "Temkinli (%70)": 0.7,
    "Güvenli (%80)": 0.8,
    "Çok güvenli (%90) - fire artar": 0.9,
    MALIYET_ANAHTAR: None,
}

YONTEM_AD = {"Baz": "Baz çizgi", "Model": "İstatistiksel model", "YoY": "Geçen yıl bazlı",
             "Karma": "Öğrenen karma (ağırlıklı)"}


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
    if v is None or isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return np.nan
    if isinstance(v, str):
        s = v.strip()
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
            s = s.replace(".", "")
        v = s
    return pd.to_numeric(v, errors="coerce")


def _hucre_tarih(x):
    if isinstance(x, (pd.Timestamp, dt.datetime, dt.date, np.datetime64)):
        t = pd.Timestamp(x)
    elif isinstance(x, str):
        t = pd.to_datetime(x.strip(), dayfirst=True, errors="coerce")
    else:
        return pd.NaT
    if pd.isna(t):
        return pd.NaT
    return t.normalize()


def _dt64(s):
    return pd.to_datetime(s).astype("datetime64[ns]")


def _tarih_seri(seri):
    return _dt64(pd.to_datetime(seri.map(_hucre_tarih), errors="coerce"))


def mape(gercek, tahmin):
    gercek = np.asarray(gercek, float)
    tahmin = np.asarray(tahmin, float)
    return float(np.mean(np.abs(gercek - tahmin) / gercek))


def yanlilik(gercek, tahmin):
    gercek = np.asarray(gercek, float)
    tahmin = np.asarray(tahmin, float)
    return float(np.mean((tahmin - gercek) / gercek))


def nth_weekday(yil, ay, hafta_gunu, n):
    d = dt.date(yil, ay, 1)
    kayma = (hafta_gunu - d.weekday()) % 7
    return d + dt.timedelta(days=kayma + 7 * (n - 1))


EPOCH = pd.Timestamp("2000-01-01")


def _gno(t):
    return int((pd.Timestamp(t) - EPOCH).days)


def _bin(x):
    return f"{int(round(x)):,}".replace(",", ".")


def _bos_gunluk():
    return pd.DataFrame({"Tarih": pd.Series(dtype="datetime64[ns]"), "Adet": pd.Series(dtype=float),
                         "Bitti": pd.Series(dtype=bool), "Not": pd.Series(dtype=str)})


def _bos_detay():
    return pd.DataFrame({"Tarih": pd.Series(dtype="datetime64[ns]"), "Cesit": pd.Series(dtype=str),
                         "Adet": pd.Series(dtype=float)})


def _df(df, **kw):
    try:
        st.dataframe(df, width="stretch", **kw)
    except Exception:
        st.dataframe(df, use_container_width=True, **kw)


def _editor(df, **kw):
    try:
        return st.data_editor(df, width="stretch", **kw)
    except Exception:
        return st.data_editor(df, use_container_width=True, **kw)


# ----------------------------------------------------------------------------
# ÖZEL GÜNLER
# ----------------------------------------------------------------------------
def ozel_gun_kumeleri(yillar, ek_metin):
    bayram, oncesi, sonrasi, resmi = set(), set(), set(), set()
    for b, e in BAYRAMLAR:
        bs, be = dt.date.fromisoformat(b), dt.date.fromisoformat(e)
        for i in range((be - bs).days + 1):
            bayram.add(bs + dt.timedelta(days=i))
        for i in range(1, 4):
            oncesi.add(bs - dt.timedelta(days=i))
        for i in range(1, 3):
            sonrasi.add(be + dt.timedelta(days=i))
    olaylar = {ad: set() for ad in OLAY_ETIKET if ad.startswith("ozel_")}
    for y in yillar:
        for m, g in [(1, 1), (4, 23), (5, 1), (5, 19), (7, 15), (8, 30)]:
            resmi.add(dt.date(y, m, g))
        olaylar["ozel_sevgililer"].add(dt.date(y, 2, 14))
        olaylar["ozel_cumhuriyet"].update({dt.date(y, 10, 28), dt.date(y, 10, 29)})
        olaylar["ozel_kadinlar"].add(dt.date(y, 3, 8))
        olaylar["ozel_ogretmenler"].add(dt.date(y, 11, 24))
        olaylar["ozel_yilbasi"].update({dt.date(y, 12, 30), dt.date(y, 12, 31)})
        olaylar["ozel_anneler"].add(nth_weekday(y, 5, 6, 2))
        olaylar["ozel_babalar"].add(nth_weekday(y, 6, 6, 3))
    for satir in (ek_metin or "").splitlines():
        try:
            olaylar["ozel_diger"].add(dt.date.fromisoformat(satir.strip()))
        except ValueError:
            pass
    return {"bayram": bayram, "oncesi": oncesi, "sonrasi": sonrasi, "resmi": resmi, "olaylar": olaylar}


def aktif_olaylar(d, kum):
    r = []
    if d in kum["bayram"]:
        r.append("bayram")
    if d in kum["oncesi"]:
        r.append("bayram_oncesi")
    if d in kum["sonrasi"]:
        r.append("bayram_sonrasi")
    if d in kum["resmi"]:
        r.append("resmi_tatil")
    for ad, kume in kum["olaylar"].items():
        if d in kume:
            r.append(ad)
    return r


# ----------------------------------------------------------------------------
# VERİTABANI (SQLite)
# ----------------------------------------------------------------------------
def _baglan():
    return sqlite3.connect(DB_YOLU)


def _eski_csv_goc(con):
    try:
        if os.path.exists(ESKI_DETAY_CSV) and con.execute("SELECT COUNT(*) FROM detay").fetchone()[0] == 0:
            d = pd.read_csv(ESKI_DETAY_CSV)
            d["Tarih"] = pd.to_datetime(d["Tarih"], errors="coerce")
            d = d.dropna(subset=["Tarih", "Cesit", "Adet"])
            con.executemany(
                "INSERT OR REPLACE INTO detay VALUES (?,?,?)",
                [(t.date().isoformat(), str(c), float(a)) for t, c, a in zip(d["Tarih"], d["Cesit"], d["Adet"])])
        if os.path.exists(ESKI_GUNLUK_CSV) and con.execute("SELECT COUNT(*) FROM gunluk").fetchone()[0] == 0:
            e = pd.read_csv(ESKI_GUNLUK_CSV)
            e["Tarih"] = pd.to_datetime(e["Tarih"], errors="coerce")
            e = e.dropna(subset=["Tarih", "Adet"])
            con.executemany(
                "INSERT OR REPLACE INTO gunluk VALUES (?,?,?,?)",
                [(t.date().isoformat(), float(a), int(_evet(b)), "" if pd.isna(n) else str(n))
                 for t, a, b, n in zip(e["Tarih"], e["Adet"], e.get("Bitti", False), e.get("Not", ""))])
        con.commit()
    except Exception:
        pass


def db_hazirla():
    con = _baglan()
    try:
        con.executescript("""
            CREATE TABLE IF NOT EXISTS detay (
                tarih TEXT, cesit TEXT, adet REAL, PRIMARY KEY (tarih, cesit));
            CREATE TABLE IF NOT EXISTS gunluk (
                tarih TEXT PRIMARY KEY, adet REAL, bitti INTEGER DEFAULT 0, notu TEXT DEFAULT '');
            CREATE TABLE IF NOT EXISTS tahmin_log (
                hedef TEXT PRIMARY KEY, beklenen REAL, uretim REAL, baz REAL, yontem TEXT, kayit TEXT);
            CREATE TABLE IF NOT EXISTS ozel_carpan (
                tarih TEXT PRIMARY KEY, carpan REAL, aciklama TEXT DEFAULT '');
            CREATE TABLE IF NOT EXISTS ogrenme_log (
                son_veri TEXT PRIMARY KEY, n_gun INTEGER, mape_baz REAL, mape_model REAL, mape_yoy REAL,
                mape_karma REAL, mape_secilen REAL, yontem TEXT, alpha REAL, yarimomur REAL,
                dz_kullan INTEGER, kayit TEXT);
            CREATE TABLE IF NOT EXISTS uretim_gercek (
                tarih TEXT PRIMARY KEY, uretilen REAL, fire REAL DEFAULT 0);
        """)
        con.commit()
        _eski_csv_goc(con)
    finally:
        con.close()


def sql_oku(sorgu):
    con = _baglan()
    try:
        return pd.read_sql_query(sorgu, con)
    finally:
        con.close()


def detay_oku():
    df = sql_oku("SELECT tarih AS Tarih, cesit AS Cesit, adet AS Adet FROM detay")
    if df.empty:
        return _bos_detay()
    df["Tarih"] = _dt64(df["Tarih"])
    df["Adet"] = pd.to_numeric(df["Adet"], errors="coerce")
    return df.dropna(subset=["Tarih", "Adet"])


def meta_oku():
    df = sql_oku("SELECT tarih AS Tarih, adet AS AdetToplam, bitti AS Bitti, notu AS Notu FROM gunluk")
    if df.empty:
        return pd.DataFrame({"Tarih": pd.Series(dtype="datetime64[ns]"), "AdetToplam": pd.Series(dtype=float),
                             "Bitti": pd.Series(dtype=bool), "Not": pd.Series(dtype=str)})
    df["Tarih"] = _dt64(df["Tarih"])
    df["Bitti"] = df["Bitti"].fillna(0).astype(int).astype(bool)
    df["Not"] = df["Notu"].fillna("").astype(str)
    return df[["Tarih", "AdetToplam", "Bitti", "Not"]]


def detay_yaz(tarih, satirlar):
    iso = pd.Timestamp(tarih).date().isoformat()
    con = _baglan()
    try:
        con.execute("DELETE FROM detay WHERE tarih=?", (iso,))
        con.executemany("INSERT INTO detay VALUES (?,?,?)", [(iso, str(c), float(a)) for c, a in satirlar])
        con.commit()
    finally:
        con.close()


def detay_toplu_yaz(df):
    for t, g in df.groupby("Tarih"):
        detay_yaz(t, list(zip(g["Cesit"], g["Adet"])))


def meta_yaz(tarih, bitti, notu):
    iso = pd.Timestamp(tarih).date().isoformat()
    con = _baglan()
    try:
        con.execute(
            "INSERT INTO gunluk (tarih, adet, bitti, notu) VALUES (?, NULL, ?, ?) "
            "ON CONFLICT(tarih) DO UPDATE SET bitti=excluded.bitti, notu=excluded.notu",
            (iso, int(bool(bitti)), notu or ""))
        con.commit()
    finally:
        con.close()


def toplam_yaz(tarih, adet):
    iso = pd.Timestamp(tarih).date().isoformat()
    con = _baglan()
    try:
        con.execute(
            "INSERT INTO gunluk (tarih, adet, bitti, notu) VALUES (?, ?, 0, '') "
            "ON CONFLICT(tarih) DO UPDATE SET adet=excluded.adet", (iso, float(adet)))
        con.commit()
    finally:
        con.close()


def log_yaz(hedef, beklenen, uretim, baz, yontem):
    con = _baglan()
    try:
        con.execute(
            "INSERT OR REPLACE INTO tahmin_log VALUES (?,?,?,?,?,?)",
            (pd.Timestamp(hedef).date().isoformat(), float(beklenen), float(uretim), float(baz), yontem,
             dt.datetime.now().isoformat(timespec="seconds")))
        con.commit()
    finally:
        con.close()


def log_oku():
    df = sql_oku("SELECT hedef AS Hedef, beklenen AS Beklenen, uretim AS Uretim, baz AS Baz, "
                 "yontem AS Yontem, kayit AS Kayit FROM tahmin_log ORDER BY hedef DESC")
    df["Hedef"] = _dt64(df["Hedef"])
    kayit_gun = pd.to_datetime(df["Kayit"], errors="coerce").dt.normalize()
    df["Canli"] = (kayit_gun < df["Hedef"]).fillna(False).astype(bool)
    return df


def ogrenme_yaz(son_veri, n_gun, mp, secilen_mape, yontem, alpha, yarimomur, dz_kullan):
    con = _baglan()
    try:
        con.execute("INSERT OR REPLACE INTO ogrenme_log VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                    (pd.Timestamp(son_veri).date().isoformat(), int(n_gun), mp.get("Baz"), mp.get("Model"),
                     mp.get("YoY"), mp.get("Karma"), float(secilen_mape), yontem, float(alpha),
                     float(yarimomur), int(bool(dz_kullan)), dt.datetime.now().isoformat(timespec="seconds")))
        con.commit()
    finally:
        con.close()


def ogrenme_oku():
    df = sql_oku("SELECT son_veri AS Veri, n_gun AS NGun, mape_baz AS Baz, mape_model AS Model, mape_yoy AS YoY, "
                 "mape_karma AS Karma, mape_secilen AS Secilen, yontem AS Yontem, alpha AS Alpha, "
                 "yarimomur AS Yarimomur, dz_kullan AS Duzeltme FROM ogrenme_log ORDER BY son_veri")
    df["Veri"] = _dt64(df["Veri"])
    return df


def uretim_yaz(tarih, uretilen, fire):
    con = _baglan()
    try:
        con.execute("INSERT OR REPLACE INTO uretim_gercek VALUES (?,?,?)",
                    (pd.Timestamp(tarih).date().isoformat(), float(uretilen), float(fire)))
        con.commit()
    finally:
        con.close()


def uretim_oku():
    df = sql_oku("SELECT tarih AS Tarih, uretilen AS Uretilen, fire AS Fire FROM uretim_gercek ORDER BY tarih")
    df["Tarih"] = _dt64(df["Tarih"])
    return df


def carpan_tablo_oku():
    df = sql_oku("SELECT tarih AS Tarih, carpan AS Carpan, aciklama AS Aciklama FROM ozel_carpan ORDER BY tarih")
    df["Tarih"] = _dt64(df["Tarih"])
    df["Aciklama"] = df["Aciklama"].fillna("").astype(str)
    return df


def carpan_tablo_yaz(df):
    satirlar = []
    for t, c, a in zip(df["Tarih"], df["Carpan"], df["Aciklama"]):
        if pd.isna(t) or pd.isna(c):
            continue
        satirlar.append((pd.Timestamp(t).date().isoformat(), float(c), "" if pd.isna(a) else str(a)))
    con = _baglan()
    try:
        con.execute("DELETE FROM ozel_carpan")
        con.executemany("INSERT OR REPLACE INTO ozel_carpan VALUES (?,?,?)", satirlar)
        con.commit()
    finally:
        con.close()


def eklenen_gunluk_oku(detay_db, meta_db):
    parcalar = []
    if not detay_db.empty:
        g = detay_db.groupby("Tarih", as_index=False)["Adet"].sum()
        g["Bitti"] = False
        g["Not"] = "Uygulamadan girildi (çeşit bazlı)"
        parcalar.append(g)
    if not meta_db.empty:
        t = meta_db[meta_db["AdetToplam"].notna()].copy()
        if not detay_db.empty:
            t = t[~t["Tarih"].isin(detay_db["Tarih"].unique())]
        if not t.empty:
            t = t.rename(columns={"AdetToplam": "Adet"})[["Tarih", "Adet"]]
            t["Bitti"] = False
            t["Not"] = "Uygulamadan girildi (toplam)"
            parcalar.append(t)
    if not parcalar:
        return _bos_gunluk()
    return pd.concat(parcalar, ignore_index=True)[["Tarih", "Adet", "Bitti", "Not"]]


# ----------------------------------------------------------------------------
# VERİ OKUMA
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def tablo_oku(ad, ham):
    if ad.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "cp1254", "latin1"):
            try:
                return pd.read_csv(io.BytesIO(ham), sep=None, engine="python", encoding=enc, header=None)
            except Exception:
                continue
        raise ValueError("CSV okunamadı")
    xls = pd.ExcelFile(io.BytesIO(ham))
    adaylar = [s for s in xls.sheet_names if "siparis" in _norm(s) or "satis" in _norm(s)]
    sayfa = adaylar[0] if adaylar else xls.sheet_names[0]
    return xls.parse(sayfa, header=None)


def kaynak_bul(adaylar, anahtarlar, haric=()):
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


def matris_isle(df):
    bas = None
    for r in range(min(6, df.shape[0])):
        n = sum(pd.notna(_hucre_tarih(df.iloc[r, c])) for c in range(1, df.shape[1]))
        if n >= 3:
            bas = r
            break
    if bas is None:
        raise ValueError("Tarih başlık satırı bulunamadı")

    son_kolon = {}
    for c in range(1, df.shape[1]):
        t = _hucre_tarih(df.iloc[bas, c])
        if pd.notna(t):
            son_kolon[t] = c

    urun_satirlari, urun_sirasi = [], []
    for r in range(bas + 1, df.shape[0]):
        ad = str(df.iloc[r, 0]).strip()
        if not ad or ad.lower() == "nan" or "toplam" in _norm(ad):
            continue
        urun_satirlari.append((r, ad))
        if ad not in urun_sirasi:
            urun_sirasi.append(ad)

    gunluk_rows, detay_rows = [], []
    for t, c in son_kolon.items():
        degerler = [(ad, _sayi(df.iloc[r, c])) for r, ad in urun_satirlari]
        if all(pd.isna(v) for _, v in degerler):
            continue
        gun = {}
        for ad, v in degerler:
            v = 0.0 if (pd.isna(v) or v < 0) else float(v)
            gun[ad] = gun.get(ad, 0.0) + v
        toplam = sum(gun.values())
        if toplam <= 0:
            continue
        gunluk_rows.append({"Tarih": t, "Adet": toplam, "Bitti": False, "Not": ""})
        for ad, v in gun.items():
            detay_rows.append({"Tarih": t, "Cesit": ad, "Adet": v})

    g = pd.DataFrame(gunluk_rows) if gunluk_rows else _bos_gunluk()
    d = pd.DataFrame(detay_rows) if detay_rows else _bos_detay()
    if not g.empty:
        g["Tarih"] = _dt64(g["Tarih"])
    if not d.empty:
        d["Tarih"] = _dt64(d["Tarih"])
    return g, d, urun_sirasi


def _arsiv_sayfa_coz(raw, sheet):
    if raw.shape[0] < 3 or raw.shape[1] < 2:
        return None, "çok küçük / boş, atlandı"

    t_idx, t_ser = None, None
    for c in range(raw.shape[1]):
        p = raw.iloc[:, c].map(_hucre_tarih)
        gecerli = p.dropna()
        dolu = int(raw.iloc[:, c].notna().sum())
        if len(gecerli) >= 5 and dolu > 0 and len(gecerli) / dolu >= 0.6:
            yillar = pd.to_datetime(gecerli).dt.year
            if yillar.between(2015, 2035).mean() >= 0.9:
                t_idx, t_ser = c, pd.to_datetime(p, errors="coerce")
                break

    if t_idx is None:
        try:
            g, _, _ = matris_isle(raw)
            if not g.empty:
                g = g.copy()
                g["Not"] = f"Arşiv ({sheet})"
                return g, f"matris formatı okundu ({len(g)} gün)"
        except Exception:
            pass
        return None, "tarih sütunu bulunamadı, atlandı"

    maske = t_ser.notna()
    adaylar = []
    for c in range(raw.shape[1]):
        if c == t_idx:
            continue
        s = pd.to_numeric(raw.loc[maske, c].map(_sayi), errors="coerce")
        if s.notna().mean() >= 0.6 and s.median() > 0:
            baslik = " ".join(_norm(x) for x in raw.iloc[:5, c] if isinstance(x, str))
            oncelik = any(k in baslik for k in ("adet", "toplam", "satis", "siparis", "pasta"))
            adaylar.append((oncelik, float(s.median()), c, s))
    if not adaylar:
        return None, "adet sütunu bulunamadı, atlandı"
    adaylar.sort(key=lambda x: (x[0], x[1]), reverse=True)
    _, _, a_idx, s = adaylar[0]

    temp = pd.DataFrame({"Tarih": t_ser[maske], "Adet": s}).dropna()
    temp = temp[temp["Adet"] > 0]
    if temp.empty:
        return None, "geçerli satır yok, atlandı"
    tekrar = int(temp["Tarih"].duplicated().sum())
    temp = temp.groupby("Tarih", as_index=False)["Adet"].sum()
    temp["Tarih"] = _dt64(temp["Tarih"])
    temp["Bitti"] = False
    temp["Not"] = f"Arşiv ({sheet})"
    bilgi = (f"tarih sütunu={t_idx}, adet sütunu={a_idx}, {len(temp)} gün "
             f"({temp['Tarih'].min():%d.%m.%Y} - {temp['Tarih'].max():%d.%m.%Y})")
    if tekrar:
        bilgi += f"; aynı tarihe ait {tekrar} fazla satır toplandı"
    return temp[["Tarih", "Adet", "Bitti", "Not"]], bilgi


def gecmis_veri_oku():
    bulunan = [f for f in GECMIS_DOSYALAR if os.path.exists(f)]
    if not bulunan:
        return _bos_gunluk(), ["Arşiv dosyası bulunamadı (aranan: " + ", ".join(GECMIS_DOSYALAR) + ")"]
    log, parcalar = [], []
    for dosya in bulunan:
        try:
            xls = pd.ExcelFile(dosya)
        except Exception as e:
            log.append(f"{dosya}: açılamadı ({e})")
            continue
        for sheet in xls.sheet_names:
            try:
                raw = xls.parse(sheet, header=None)
                temp, bilgi = _arsiv_sayfa_coz(raw, sheet)
                log.append(f"{dosya} / {sheet}: {bilgi}")
                if temp is not None:
                    parcalar.append(temp)
            except Exception as e:
                log.append(f"{dosya} / {sheet}: hata ({e})")
    if not parcalar:
        return _bos_gunluk(), log
    tum = pd.concat(parcalar, ignore_index=True).drop_duplicates("Tarih", keep="last")
    return tum, log


def yukleneni_coz(ad, ham):
    raw = tablo_oku(ad, ham)
    for r in range(min(6, raw.shape[0])):
        hucre = [_norm(x) for x in raw.iloc[r].tolist()]
        i_t = next((i for i, h in enumerate(hucre) if h == "tarih"), None)
        i_c = next((i for i, h in enumerate(hucre) if h in ("cesit", "urun", "urun adi", "cesit adi")), None)
        i_a = next((i for i, h in enumerate(hucre) if h in ("adet", "miktar", "siparis")), None)
        if None not in (i_t, i_c, i_a):
            alt = raw.iloc[r + 1:]
            d = pd.DataFrame({
                "Tarih": _tarih_seri(alt.iloc[:, i_t]),
                "Cesit": alt.iloc[:, i_c].astype(str).str.strip(),
                "Adet": pd.to_numeric(alt.iloc[:, i_a].map(_sayi), errors="coerce"),
            }).dropna(subset=["Tarih", "Adet"])
            d = d[(d["Cesit"] != "") & (d["Cesit"].str.lower() != "nan") & (d["Adet"] >= 0)]
            d = d.groupby(["Tarih", "Cesit"], as_index=False)["Adet"].sum()
            return d, "Uzun format (Tarih / Çeşit / Adet) algılandı."
    g, d, _ = matris_isle(raw)
    return d, "Matris format (satır = ürün, sütun = tarih) algılandı."


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
def hava_getir(bas_iso, bit_iso):
    lat, lon = 41.01, 28.97
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
    h["Tarih"] = _dt64(h["Tarih"])
    return h.sort_values("Tarih").reset_index(drop=True)


# ----------------------------------------------------------------------------
# ÖZELLİKLER VE MODEL
# ----------------------------------------------------------------------------
def gecmis_ozellikleri(tarihler, ly):
    ort = float(np.mean(list(ly.values()))) if ly else 0.0
    s7, l7 = [], []
    for t in tarihler:
        g = _gno(t)
        son = []
        for x in range(g - 1, g - 15, -1):
            if x in ly:
                son.append(ly[x])
                if len(son) == 7:
                    break
        m = float(np.mean(son)) if len(son) >= 3 else np.nan
        lg = np.nan
        for k in (7, 14, 21, 28, 35):
            if (g - k) in ly:
                lg = ly[g - k]
                break
        if np.isnan(m):
            m = lg if not np.isnan(lg) else ort
        if np.isnan(lg):
            lg = m
        s7.append(m)
        l7.append(lg)
    return np.array(s7, float), np.array(l7, float)


def ozellik_olustur(df, cfg):
    d = pd.to_datetime(df["Tarih"])
    X = pd.DataFrame(index=df.index)
    for k in range(7):
        X[f"gun_{k}"] = (d.dt.dayofweek == k).astype(float)
    gun = d.dt.day
    X["ay_basi"] = (gun <= 5).astype(float)
    X["ay_ortasi"] = gun.between(10, 15).astype(float)
    X["ay_sonu"] = (gun >= 26).astype(float)
    if cfg["use_season"]:
        doy = d.dt.dayofyear.to_numpy()
        for h in (1, 2):
            X[f"mevsim_sin{h}"] = np.sin(2 * np.pi * h * doy / 365.25)
            X[f"mevsim_cos{h}"] = np.cos(2 * np.pi * h * doy / 365.25)
    t = d.dt.date
    k = cfg["kum"]
    X["bayram"] = t.isin(k["bayram"]).astype(float)
    X["bayram_oncesi"] = t.isin(k["oncesi"]).astype(float)
    X["bayram_sonrasi"] = t.isin(k["sonrasi"]).astype(float)
    X["resmi_tatil"] = t.isin(k["resmi"]).astype(float)
    for ad, kume in k["olaylar"].items():
        X[ad] = t.isin(kume).astype(float)
    X["trend"] = (d - cfg["t0"]).dt.days / 365.0
    X["menu_sonrasi"] = (d >= cfg["menu"]).astype(float) if cfg["menu"] is not None else 0.0
    if cfg.get("ly"):
        s7, l7 = gecmis_ozellikleri(d, cfg["ly"])
        X["son7"] = s7
        X["lag7"] = l7
    if cfg["use_weather"]:
        X["sicaklik"] = df["sicaklik"].astype(float)
        X["yagmurlu"] = (df["yagis"].astype(float) >= 1.0).astype(float)
        ik, gen = cfg.get("iklim", {}), cfg.get("iklim_genel", 15.0)
        norm = d.dt.month.map(lambda m: ik.get(m, gen)).to_numpy(float)
        X["sic_sapma"] = df["sicaklik"].astype(float).to_numpy() - norm
    return X


def fit_model(X, y_log, alpha=10.0, yarimomur=0.0, dayanikli=True):
    cols = list(X.columns)
    A = X.to_numpy(float)
    y = np.asarray(y_log, float)
    n = len(y)
    w = np.ones(n)
    if yarimomur and "trend" in cols:
        gun_no = A[:, cols.index("trend")] * 365.0
        w = 0.5 ** ((gun_no.max() - gun_no) / yarimomur)
    w = w / w.mean()

    def _coz(w):
        mu = (A * w[:, None]).sum(0) / w.sum()
        sd = np.sqrt((((A - mu) ** 2) * w[:, None]).sum(0) / w.sum())
        sd = np.where(sd > 1e-9, sd, 1.0)
        Z = (A - mu) / sd
        yb = float((w * y).sum() / w.sum())
        ZW = Z * w[:, None]
        coef = np.linalg.solve(Z.T @ ZW + alpha * np.eye(Z.shape[1]), ZW.T @ (y - yb))
        return mu, sd, Z, yb, coef

    mu, sd, Z, yb, coef = _coz(w)
    res = y - (yb + Z @ coef)
    sigma = float(1.4826 * np.median(np.abs(res - np.median(res))))
    if dayanikli and n >= 30:
        c = max(1.5 * sigma, 0.03)
        hw = np.where(np.abs(res) <= c, 1.0, c / np.maximum(np.abs(res), 1e-9))
        w2 = w * hw
        w2 = w2 / w2.mean()
        mu, sd, Z, yb, coef = _coz(w2)
        res = y - (yb + Z @ coef)
        sigma = float(1.4826 * np.median(np.abs(res - np.median(res))))
    return {"w": coef, "b": yb, "mu": pd.Series(mu, index=cols), "sd": pd.Series(sd, index=cols),
            "cols": cols, "sigma": sigma}


def predict_model(mod, X):
    Z = ((X[mod["cols"]] - mod["mu"]) / mod["sd"]).to_numpy(float)
    return np.exp(np.clip(Z @ mod["w"] + mod["b"], 0.0, 15.0))


def baz_cizgi(v, hedef):
    gecmis = v[v["Tarih"] < hedef]
    if gecmis.empty:
        return float(v["Y"].mean())
    ayni = gecmis[(gecmis["Tarih"].dt.dayofweek == hedef.dayofweek)
                  & (gecmis["Tarih"] >= hedef - pd.Timedelta(days=70))].tail(6)["Y"]
    if len(ayni) >= 3:
        return float(ayni.mean())
    son = gecmis[gecmis["Tarih"] >= hedef - pd.Timedelta(days=42)]
    if len(son) < 7:
        son = gecmis.tail(28)
    return float(son["Y"].mean())


def yoy_tahmin(yd, hedef):
    g = _gno(hedef)
    parcalar = [(yd.get(g - 364 + k), w) for k, w in ((-7, 0.25), (0, 0.5), (7, 0.25))]
    parcalar = [(v, w) for v, w in parcalar if v is not None]
    if not parcalar or sum(w for _, w in parcalar) < 0.5:
        return np.nan
    gecen = sum(v * w for v, w in parcalar) / sum(w for _, w in parcalar)
    gunler = [x for x in range(g - 1, g - 60, -1) if x in yd][:28]
    ortak = [x for x in gunler if (x - 364) in yd]
    if len(gunler) < 14 or len(ortak) < 7:
        return np.nan
    pay = sum(yd[x] for x in ortak)
    payda = sum(yd[x - 364] for x in ortak)
    if payda <= 0:
        return np.nan
    return float(gecen * np.clip(pay / payda, 0.6, 1.8))


def oz_duzeltme(gercek, tahmin, hs, pencere=21, min_n=8, kuvvet=0.5, sinir=0.08):
    gercek = np.asarray(gercek, float)
    tahmin = np.asarray(tahmin, float)
    lr = np.log(gercek / tahmin)
    ls = float(np.log(1 + sinir))
    n = len(lr)
    duz = tahmin.copy()
    for j in range(n):
        gec = [lr[k] for k in range(j) if hs[k] == hs[j]][-pencere:]
        if len(gec) >= min_n:
            duz[j] = tahmin[j] * float(np.exp(np.clip(kuvvet * np.mean(gec), -ls, ls)))
    faktor = {}
    for tip in (False, True):
        gec = [lr[k] for k in range(n) if hs[k] == tip][-pencere:]
        faktor[tip] = float(np.exp(np.clip(kuvvet * np.mean(gec), -ls, ls))) if len(gec) >= min_n else 1.0
    return duz, faktor


@st.cache_data(show_spinner="Geriye dönük test çalışıyor...")
def geriye_donuk_test(v, X, alphalar=(1.0, 3.0, 10.0, 30.0), yarimomurler=(0.0, 365.0),
                      test_gun=70, min_egitim=60):
    n = len(v)
    if n < min_egitim + 14:
        return None
    bas = max(min_egitim, n - test_gun)
    y = v["Y"].to_numpy(float)
    y_log = np.log(y)
    tarihler = v["Tarih"]
    idx = list(range(bas, n))
    gercek = y[bas:]
    yd = {_gno(t): float(a) for t, a in zip(tarihler, y)}
    baz = np.array([baz_cizgi(v.iloc[:i], tarihler.iloc[i]) for i in idx])
    yoy = np.array([yoy_tahmin(yd, tarihler.iloc[i]) for i in idx])

    en_iyi, en_m, model = None, np.inf, None
    for hl in yarimomurler:
        for a in alphalar:
            tah = np.array([
                float(np.clip(predict_model(fit_model(X.iloc[:i], y_log[:i], a, hl), X.iloc[[i]])[0],
                              0.6 * baz[j], 1.6 * baz[j]))
                for j, i in enumerate(idx)])
            m = mape(gercek, tah)
            if m < en_m:
                en_iyi, en_m, model = (a, hl), m, tah

    yoy_ok = np.isfinite(yoy)
    n_yoy = int(yoy_ok.sum())
    kapsama = float(yoy_ok.mean())

    def _agirliklar(maske, anahtarlar, tah):
        hatalar = {k: mape(gercek[maske], tah[k][maske]) for k in anahtarlar}
        ham_w = np.array([1.0 / (hatalar[k] ** 2 + 1e-6) for k in anahtarlar])
        ham_w = ham_w / ham_w.sum()
        w_ = 0.7 * ham_w + 0.3 / len(anahtarlar)
        return {k: float(x) for k, x in zip(anahtarlar, w_)}

    tum = np.ones(len(gercek), bool)
    bilesen = {"Baz": baz, "Model": model}
    w2 = _agirliklar(tum, ("Baz", "Model"), bilesen)
    karma = w2["Baz"] * baz + w2["Model"] * model
    agirlik = dict(w2)
    if n_yoy >= 8:
        bilesen["YoY"] = np.where(yoy_ok, yoy, baz)
        w3 = _agirliklar(yoy_ok, ("Baz", "Model", "YoY"), bilesen)
        karma3 = w3["Baz"] * baz + w3["Model"] * model + w3["YoY"] * bilesen["YoY"]
        karma = np.where(yoy_ok, karma3, karma)
        agirlik = w3
    mp = {k: mape(gercek, p) for k, p in bilesen.items()}
    bilesen["Karma"] = karma
    mp["Karma"] = mape(gercek, karma)
    yontem = min(mp, key=mp.get)

    hs = tarihler.iloc[bas:].dt.dayofweek.to_numpy() >= 4
    ham = bilesen[yontem]
    duz, faktor = oz_duzeltme(gercek, ham, hs)
    mape_ham, mape_duz = mp[yontem], mape(gercek, duz)
    kullan = bool(mape_duz < mape_ham * 0.99)
    secilen = duz if kullan else ham
    bt = pd.DataFrame({"Tarih": tarihler.iloc[bas:].to_numpy(), "Gercek": gercek, **bilesen,
                       "Ham": ham, "Secilen": secilen})
    return {
        "bt": bt, "alpha": float(en_iyi[0]), "yarimomur": float(en_iyi[1]), "yontem": yontem,
        "mape": mp, "bias": {k: yanlilik(gercek, bilesen[k]) for k in bilesen},
        "w": agirlik, "w2": w2, "yoy_kapsama": kapsama, "yoy_gun": n_yoy,
        "dz_kullan": kullan, "dz_faktor": faktor,
        "mape_ham": mape_ham, "mape_duz": mape_duz, "mape_secilen": mape(gercek, secilen),
    }


def tahmin_et(cfg, veri, mod, test, tarih, sic, yag):
    satir = pd.DataFrame({"Tarih": [tarih], "sicaklik": [sic], "yagis": [yag]})
    xr = ozellik_olustur(satir, cfg)
    baz = baz_cizgi(veri, tarih)
    model = float(np.clip(predict_model(mod, xr)[0], 0.6 * baz, 1.6 * baz)) if mod is not None else baz
    bil = {"Baz": baz, "Model": model}
    if test is None:
        bil["Karma"] = (baz + model) / 2
        return bil["Baz"], model, baz, xr, bil, {"Baz": 0.5, "Model": 0.5}
    w = test["w2"]
    if "YoY" in test["w"]:
        yv = yoy_tahmin(cfg["yd"], tarih)
        if np.isfinite(yv):
            bil["YoY"] = float(yv)
            w = test["w"]
    bil["Karma"] = sum(w[k] * bil[k] for k in w)
    sec = bil[test["yontem"]] if test["yontem"] in bil else bil["Karma"]
    return sec, model, baz, xr, bil, w


# ----------------------------------------------------------------------------
# YORUM ÜRETİMİ
# ----------------------------------------------------------------------------
def etki_tablosu(mod, xr):
    if mod is None:
        return {}
    z = ((xr[mod["cols"]] - mod["mu"]) / mod["sd"]).iloc[0]
    katki = pd.Series(mod["w"], index=mod["cols"]) * z
    sonuc = {}
    for ad, kolonlar in GRUPLAR.items():
        k = [c for c in kolonlar if c in katki.index]
        if k:
            sonuc[ad] = float(np.exp(katki[k].sum()) - 1.0)
    return sonuc


def referanslar(yd, hedef):
    g = _gno(hedef)
    ayni = [yd[g - 7 * k] for k in range(1, 5) if (g - 7 * k) in yd]
    return {"gecen_hafta": yd.get(g - 7), "ort4": float(np.mean(ayni)) if len(ayni) >= 2 else None,
            "gecen_yil": yd.get(g - 364)}


def olay_etkileri(veri, kum):
    sat = {}
    tarih = veri["Tarih"].reset_index(drop=True)
    for i in range(14, len(veri)):
        olaylar = aktif_olaylar(tarih.iloc[i].date(), kum)
        if not olaylar:
            continue
        baz = baz_cizgi(veri.iloc[:i], tarih.iloc[i])
        if baz <= 0:
            continue
        lift = float(veri["Y"].iloc[i]) / baz - 1.0
        for o in olaylar:
            sat.setdefault(o, []).append(lift)
    if not sat:
        return pd.DataFrame(columns=["n", "medyan", "min", "max"])
    return pd.DataFrame({o: {"n": len(v), "medyan": float(np.median(v)), "min": float(min(v)), "max": float(max(v))}
                         for o, v in sat.items()}).T


def anomali_bul(test, esik=3.0):
    if test is None:
        return pd.DataFrame()
    bt = test["bt"]
    r = np.log(bt["Gercek"].to_numpy(float) / bt["Secilen"].to_numpy(float))
    med = float(np.median(r))
    mad = max(float(np.median(np.abs(r - med))) * 1.4826, 0.03)
    z = (r - med) / mad
    m = np.abs(z) >= esik
    out = bt.loc[m, ["Tarih", "Gercek", "Secilen"]].copy()
    out["Sapma"] = (np.exp(r[m]) - 1.0)
    return out.reset_index(drop=True)


def guven_etiketi(P, ufuk):
    yg = P.get("yg")
    puan = 2 if yg is None else (0 if yg < 0.08 else 1 if yg < 0.15 else 2)
    if P["gozlenmeyen"]:
        puan += 1
    if P["hava_kaynak"] == "mevsim ortalaması":
        puan += 1
    if ufuk > 3:
        puan += 1
    return ["🟢 Yüksek", "🟡 Orta", "🔴 Düşük"][min(puan, 2)]


def yonetici_ozeti(P, ref):
    td = P["tarih"]
    s = (f"**{td:%d.%m.%Y} {GUN_TR[td.dayofweek]}** için **~{_bin(P['uretim'])} adet** üretim önerilir. "
         f"Beklenen talep {_bin(P['beklenen'])} adet")
    if P.get("lo") is not None:
        s += f" (olası aralık {_bin(P['lo'])} - {_bin(P