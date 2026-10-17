# -*- coding: utf-8 -*-
"""
Tatlıbak Yaş Pasta - Akıllı Üretim Planlama (app.py)

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
    "resmi_tatil": "Resmi tatil",
    "ozel_sevgililer": "Sevgililer Günü",
    "ozel_kadinlar": "Kadınlar Günü",
    "ozel_anneler": "Anneler Günü",
    "ozel_babalar": "Babalar Günü",
    "ozel_ogretmenler": "Öğretmenler Günü",
    "ozel_yilbasi": "Yılbaşı",
    "ozel_diger": "Özel gün (elle eklenen)",
}

GRUPLAR = {
    "Gün tipi": [f"gun_{k}" for k in range(7)],
    "Maaş / ay günü": ["ay_basi", "ay_ortasi", "ay_sonu"],
    "Mevsim": ["mevsim_sin1", "mevsim_cos1", "mevsim_sin2", "mevsim_cos2"],
    "Bayram / özel gün": list(OLAY_ETIKET.keys()),
    "Hava": ["sicaklik", "yagmurlu"],
    "Trend / menü": ["trend", "menu_sonrasi"],
}

MALIYET_ANAHTAR = "Maliyete göre (fire / kaçan satış TL)"
SEVIYELER = {
    "Dengeli (%50) - fire az, kaçan satış olabilir": 0.5,
    "Temkinli (%70)": 0.7,
    "Güvenli (%80)": 0.8,
    "Çok güvenli (%90) - fire artar": 0.9,
    MALIYET_ANAHTAR: None,
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
    bayram, oncesi, resmi = set(), set(), set()
    for b, e in BAYRAMLAR:
        bs, be = dt.date.fromisoformat(b), dt.date.fromisoformat(e)
        for i in range((be - bs).days + 1):
            bayram.add(bs + dt.timedelta(days=i))
        for i in range(1, 4):
            oncesi.add(bs - dt.timedelta(days=i))
    olaylar = {ad: set() for ad in OLAY_ETIKET if ad.startswith("ozel_")}
    for y in yillar:
        for m, g in [(1, 1), (4, 23), (5, 1), (5, 19), (7, 15), (8, 30), (10, 29)]:
            resmi.add(dt.date(y, m, g))
        olaylar["ozel_sevgililer"].add(dt.date(y, 2, 14))
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
    return {"bayram": bayram, "oncesi": oncesi, "resmi": resmi, "olaylar": olaylar}


def aktif_olaylar(d, kum):
    r = []
    if d in kum["bayram"]:
        r.append("bayram")
    if d in kum["oncesi"]:
        r.append("bayram_oncesi")
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
                 "yontem AS Yontem, kayit AS Kayit FROM tahmin_log ORDER BY hedef")
    df["Hedef"] = _dt64(df["Hedef"])
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
    X["resmi_tatil"] = t.isin(k["resmi"]).astype(float)
    for ad, kume in k["olaylar"].items():
        X[ad] = t.isin(kume).astype(float)
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
        for j, i in enumerate(range(bas, n)):
            mod = fit_model(X.iloc[:i], y_log[:i], a)
            ham = float(predict_model(mod, X.iloc[[i]])[0])
            tah.append(float(np.clip(ham, 0.6 * baz[j], 1.6 * baz[j])))
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
        "bias": {k: yanlilik(bt["Gercek"], bt[k]) for k in ("Baz", "Model", "Karma")},
    }


def retro_tahmin(veri, X, i, yontem, alpha):
    v = veri[["Tarih", "Y"]].iloc[:i]
    hedef = veri["Tarih"].iloc[i]
    baz = baz_cizgi(v, hedef)
    if yontem == "Baz" or i < 60:
        return baz
    mod = fit_model(X.iloc[:i], np.log(v["Y"].to_numpy(float)), alpha)
    m = float(np.clip(predict_model(mod, X.iloc[[i]])[0], 0.6 * baz, 1.6 * baz))
    return m if yontem == "Model" else (baz + m) / 2


def tahmin_et(cfg, veri, mod, yontem, tarih, sic, yag):
    satir = pd.DataFrame({"Tarih": [tarih], "sicaklik": [sic], "yagis": [yag]})
    xr = ozellik_olustur(satir, cfg)
    baz = baz_cizgi(veri, tarih)
    if mod is not None:
        model = float(np.clip(predict_model(mod, xr)[0], 0.6 * baz, 1.6 * baz))
    else:
        model = baz
    sec = {"Baz": baz, "Model": model, "Karma": (baz + model) / 2}[yontem]
    return sec, model, baz, xr


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


def yorum_uret(P, mod, gozlenmeyen, use_weather):
    etkiler = etki_tablosu(mod, P["xr"])
    yorumlar = []

    if use_weather:
        yd = f"{P['sic']:.0f}°C" + (f", {P['yag']:.1f} mm yağış" if P["yag"] >= 1.0 else ", yağış beklenmiyor")
        satir = f"🌤️ **Hava:** {yd} ({P['hava_kaynak']})."
        e = etkiler.get("Hava")
        if e is not None:
            if abs(e) >= 0.03:
                satir += f" Modele göre talebi yaklaşık %{abs(e) * 100:.0f} {'artırıyor' if e > 0 else 'azaltıyor'}."
            else:
                satir += " Modele göre belirgin bir etkisi yok."
        yorumlar.append(satir)
    else:
        yorumlar.append("🌤️ **Hava:** Hava durumu bu tahmine dahil edilmedi.")

    if P["olaylar"]:
        etiket = ", ".join(OLAY_ETIKET[o] for o in P["olaylar"])
        satir = f"🎉 **Özel gün:** {etiket}."
        e = etkiler.get("Bayram / özel gün")
        if e is not None and abs(e) >= 0.03 and not P["gozlenmeyen"]:
            satir += f" Modelin geçmişten öğrendiği etki yaklaşık %{e * 100:+.0f}."
        if P["gozlenmeyen"]:
            ad = ", ".join(OLAY_ETIKET[o] for o in P["gozlenmeyen"])
            satir += (f" ⚠️ Model **{ad}** gününü geçmiş verilerde hiç görmedi, etkisini bilmiyor. "
                      "Geçen yılki benzer güne bakıp 'Özel Gün Çarpanları' bölümünden elle çarpan girin.")
        yorumlar.append(satir)

    e = etkiler.get("Maaş / ay günü")
    if e is not None and abs(e) >= 0.03:
        yorumlar.append(f"💰 **Ay içi konum:** Ay başı/ortasi/sonu etkisi yaklaşık %{e * 100:+.0f}.")

    if mod is not None and "trend" in mod["cols"]:
        i = mod["cols"].index("trend")
        yillik = float(np.exp(mod["w"][i] / mod["sd"]["trend"]) - 1.0)
        if abs(yillik) >= 0.05:
            yorumlar.append(f"📈 **Trend:** Model yıllık yaklaşık %{yillik * 100:+.0f} bir eğilim görüyor.")

    if mod is not None and P["baz"] > 0:
        fark = P["model"] / P["baz"] - 1.0
        if abs(fark) >= 0.20:
            yorumlar.append(f"🔎 **Model / baz farkı:** Model, son haftalardaki aynı gün ortalamasından "
                            f"%{abs(fark) * 100:.0f} {'yüksek' if fark > 0 else 'düşük'} tahmin ediyor. Nedenini kontrol edin.")

    if abs(P["manuel"] - 1.0) > 1e-9:
        yorumlar.append(f"🎯 **Elle çarpan:** Beklenen talebe ×{P['manuel']:.2f} uygulandı.")
    return yorumlar


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


def urun_dagilimi(detay, hedef, toplam, pencere=60, k_cesit=300.0, orijinal_sira=None, haric_urunler=None):
    if detay is None or detay.empty:
        return pd.DataFrame(), "Veri yok"
    son = detay["Tarih"].max()
    d = detay[detay["Tarih"] > son - pd.Timedelta(days=pencere)].copy()
    haric_norm = {_norm(x) for x in (haric_urunler or [])}
    if haric_norm:
        d = d[~d["Cesit"].map(_norm).isin(haric_norm)]
    if d.empty:
        return pd.DataFrame(), "Veri yok"

    hs = hedef.dayofweek >= 4
    maske = d["Tarih"].dt.dayofweek.ge(4) == hs

    mevcut = list(pd.unique(d["Cesit"]))
    cesitler = [c for c in (orijinal_sira or []) if c in mevcut]
    cesitler += [c for c in mevcut if c not in cesitler]

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
# VERİ GİRİŞİ VE ARŞİV SEKMELERİ
# ----------------------------------------------------------------------------
def giris_sekmesi(urunler, detay_tum, meta_db):
    st.subheader("📝 Günlük Sipariş / Satış Girişi")

    with st.expander("📤 Dosyadan yükle (Excel / CSV)", expanded=False):
        st.caption("Desteklenen biçimler: (1) 'Tarih', 'Çeşit', 'Adet' sütunlu liste, "
                   "(2) 2026 dosyasındaki gibi satır = ürün, sütun = tarih matrisi.")
        up = st.file_uploader("Sipariş dosyası", type=["xlsx", "csv"], key="up_siparis")
        if up is not None:
            try:
                d_up, bilgi = yukleneni_coz(up.name, up.getvalue())
                st.write(bilgi)
                if d_up.empty:
                    st.warning("Dosyada eklenebilecek veri bulunamadı.")
                else:
                    ozet = d_up.groupby("Tarih", as_index=False)["Adet"].sum()
                    ozet["Tarih"] = ozet["Tarih"].dt.strftime("%d.%m.%Y")
                    ozet = ozet.rename(columns={"Adet": "Günlük toplam"})
                    _df(ozet, hide_index=True)
                    st.caption("Aynı tarihe daha önce girilmiş çeşit verisi varsa bu dosyadakiyle değiştirilir.")
                    if st.button("✅ Veritabanına ekle", key="btn_up_ekle"):
                        detay_toplu_yaz(d_up)
                        st.session_state["_mesaj"] = (f"{d_up['Tarih'].nunique()} günlük veri eklendi "
                                                      f"(toplam {int(d_up['Adet'].sum())} adet).")
                        st.rerun()
            except Exception as e:
                st.error(f"Dosya okunamadı: {e}")

    st.markdown("---")
    g_tarih = st.date_input("Satış / sipariş tarihi", value=BUGUN, key="giris_tarih")
    ts = pd.Timestamp(g_tarih)
    iso = g_tarih.isoformat()

    mevcut = detay_tum[detay_tum["Tarih"] == ts].set_index("Cesit")["Adet"] if not detay_tum.empty else pd.Series(dtype=float)
    tablo = pd.DataFrame({"Ürün": urunler, "Adet": [int(mevcut.get(u, 0)) for u in urunler]})
    m = meta_db[meta_db["Tarih"] == ts]
    bitti_varsayilan = bool(m["Bitti"].iloc[0]) if len(m) else False
    not_varsayilan = str(m["Not"].iloc[0]) if len(m) else ""

    if len(mevcut):
        st.info(f"Bu tarih için {int(mevcut.sum())} adetlik kayıt var; aşağıda düzenleyip kaydedebilirsiniz.")

    st.caption("Adet sütununa yazın; **Tab / Enter / ok tuşlarıyla** hücreler arasında ilerleyebilirsiniz.")
    with st.form(f"form_{iso}"):
        duz = _editor(
            tablo, hide_index=True, disabled=["Ürün"], key=f"ed_{iso}", height=min(36 * (len(tablo) + 1) + 4, 900),
            column_config={"Adet": st.column_config.NumberColumn("Adet", min_value=0, step=1, format="%d")})
        bitti = st.checkbox("Bu gün bazı ürünler bitti (gerçek talep satıştan yüksek olabilir)", value=bitti_varsayilan)
        notu = st.text_input("Not (kampanya, etkinlik, aksaklık vb.)", value=not_varsayilan)
        kaydet = st.form_submit_button("💾 Çeşit Bazlı Satışları Kaydet")

    if kaydet:
        satirlar = [(u, int(a)) for u, a in zip(duz["Ürün"], duz["Adet"].fillna(0)) if a and a > 0]
        detay_yaz(ts, satirlar)
        meta_yaz(ts, bitti, notu)
        toplam = sum(a for _, a in satirlar)
        st.session_state["_mesaj"] = (f"{g_tarih:%d.%m.%Y} kaydedildi. Toplam: {toplam} adet." if toplam > 0
                                      else f"{g_tarih:%d.%m.%Y} için çeşit kaydı temizlendi.")
        st.rerun()

    with st.expander("Sadece günlük toplam gir (çeşit bilgisi yoksa)"):
        toplam_giris = st.number_input("Günlük toplam adet", min_value=0, value=0, step=1, key=f"top_{iso}")
        if st.button("Toplamı kaydet", key=f"btn_top_{iso}"):
            if toplam_giris <= 0:
                st.warning("0'dan büyük bir değer girin.")
            elif len(mevcut):
                st.error("Bu tarihte çeşit bazlı kayıt var; yukarıdaki tablodan düzenleyin.")
            else:
                toplam_yaz(ts, toplam_giris)
                st.session_state["_mesaj"] = f"{g_tarih:%d.%m.%Y} günlük toplamı kaydedildi: {int(toplam_giris)} adet."
                st.rerun()


def arsiv_sekmesi(veri, gecmis_log):
    st.subheader("Birleştirilmiş Veri Havuzu")
    if veri.empty:
        st.info("Henüz veri yok.")
    else:
        st.caption(f"Toplam kayıtlı gün: {len(veri)} | "
                   f"{veri['Tarih'].min():%d.%m.%Y} - {veri['Tarih'].max():%d.%m.%Y}")
        goster = veri.copy()
        goster["Tarih"] = goster["Tarih"].dt.strftime("%d.%m.%Y")
        _df(goster.iloc[::-1], hide_index=True)
        st.download_button("⬇️ Veriyi CSV olarak indir", goster.to_csv(index=False).encode("utf-8-sig"),
                           file_name="birlesik_veri.csv", mime="text/csv")
    with st.expander("Geçmiş yıllar arşivi okuma raporu (doğru okunduğunu buradan kontrol edin)"):
        for satir in gecmis_log:
            st.write("• " + satir)
    if os.path.exists(DB_YOLU):
        with open(DB_YOLU, "rb") as fh:
            st.download_button("⬇️ Veritabanı yedeği (.db)", fh.read(), file_name="tatlibak_yedek.db")
    st.caption(f"Uygulamadan girilen veriler '{DB_YOLU}' dosyasında saklanır. Düzenli yedek alın.")


# ----------------------------------------------------------------------------
# UYGULAMA ARAYÜZÜ
# ----------------------------------------------------------------------------
def uygulama():
    db_hazirla()
    st.title("🍰 Tatlıbak Yaş Pasta - Akıllı Üretim Planlama Motoru")

    mesaj = st.session_state.pop("_mesaj", None)
    if mesaj:
        st.success(mesaj)
    uyari_kutu = st.container()

    with st.sidebar:
        st.header("⚙️ Ayarlar & Veri")
        st.info("📍 Konum: **İstanbul** (sabit)")

        st.markdown("### 🚫 Üretimden kalkan ürünler")
        haric_metin = st.text_area("Her satıra bir ürün adı", value="\n".join(VARSAYILAN_HARIC), height=70)
        haric_liste = [x.strip() for x in haric_metin.splitlines() if x.strip()]

        st.markdown("---")
        hava_kullan = st.checkbox("Hava durumunu modele kat", value=True)
        bitti_duzeltme = st.slider("Ürün biten günlerde talep düzeltmesi (%)", 0, 30, 10)
        pencere = st.slider("Ürün dağılımı geriye bakış (gün)", 30, 120, 60)
        menu_var = st.checkbox("Menü değişikliği tarihini belirt")
        menu_tarihi = None
        if menu_var:
            menu_tarihi = pd.Timestamp(st.date_input("Yeni menünün başlangıcı", value=BUGUN - dt.timedelta(days=60)))
        ek_ozel = st.text_area("Diğer özel günler (her satıra YYYY-AA-GG)", height=60)

        st.markdown("---")
        st.markdown("### 🛡️ Güvenlik seviyesi")
        seviye_ad = st.selectbox("Seviye", list(SEVIYELER), index=1)
        if seviye_ad == MALIYET_ANAHTAR:
            fire_tl = st.number_input("Fire maliyeti (TL / adet)", min_value=1.0, value=100.0, step=10.0)
            kacan_tl = st.number_input("Kaçan satış kaybı (TL / adet)", min_value=1.0, value=100.0, step=10.0)
            seviye = float(np.clip(kacan_tl / (kacan_tl + fire_tl), 0.5, 0.95))
        else:
            seviye = SEVIYELER[seviye_ad]
        esik = st.slider("Sapma uyarı eşiği (%)", 10, 60, 25)

    gecmis_df, gecmis_log = gecmis_veri_oku()

    ad, ham = kaynak_bul(GUNLUK_DOSYALAR, ("2026", "satis", "pasta"))
    gunluk_2026, detay_2026, orijinal_urun_sirasi = _bos_gunluk(), _bos_detay(), []
    if ad is not None:
        try:
            gunluk_2026, detay_2026, orijinal_urun_sirasi = matris_isle(tablo_oku(ad, ham))
        except Exception as e:
            st.error(f"'{ad}' okunamadı: {e}")

    detay_db = detay_oku()
    meta_db = meta_oku()
    if not detay_db.empty and not detay_2026.empty:
        detay_2026 = detay_2026[~detay_2026["Tarih"].isin(detay_db["Tarih"].unique())]
    detay_parcalar = [x for x in (detay_2026, detay_db) if not x.empty]
    detay_tum = pd.concat(detay_parcalar, ignore_index=True) if detay_parcalar else _bos_detay()
    for c in detay_db["Cesit"].unique():
        if c not in orijinal_urun_sirasi:
            orijinal_urun_sirasi.append(c)

    ek_gunluk = eklenen_gunluk_oku(detay_db, meta_db)
    parcalar = [df for df in (gecmis_df, gunluk_2026, ek_gunluk) if df is not None and not df.empty]
    if parcalar:
        veri = pd.concat(parcalar, ignore_index=True)
        veri["Tarih"] = _dt64(veri["Tarih"])
        veri["Adet"] = pd.to_numeric(veri["Adet"], errors="coerce")
        veri = veri.dropna(subset=["Tarih", "Adet"])
        
        # --- GÜNCELLENEN FİLTRE: Toplam sipariş miktarı 10'un altında olan sahte/çapa günleri tahminlemeye dahil etme ---
        veri = veri[veri["Adet"] >= 10]
        
        veri["Bitti"] = veri["Bitti"].map(lambda x: bool(x) if pd.notna(x) else False)
        veri["Not"] = veri["Not"].fillna("").astype(str)
        veri = veri.drop_duplicates("Tarih", keep="last").sort_values("Tarih").reset_index(drop=True)
    else:
        veri = _bos_gunluk()

    if not meta_db.empty and not veri.empty:
        bitti_tarihler = meta_db.loc[meta_db["Bitti"], "Tarih"].tolist()
        veri["Bitti"] = veri["Bitti"] | veri["Tarih"].isin(bitti_tarihler)
        notlar = {t: n for t, n in zip(meta_db["Tarih"], meta_db["Not"]) if n}
        veri["Not"] = [notlar.get(t, n) for t, n in zip(veri["Tarih"], veri["Not"])]

    if len(veri) < 28:
        st.error(f"Tahmin için en az 28 günlük veri gerekli, şu an {len(veri)} gün var.")
        sek = st.tabs(["📝 Veri Girişi", "📂 Veri Arşivi"])
        with sek[0]:
            giris_sekmesi(orijinal_urun_sirasi, detay_tum, meta_db)
        with sek[1]:
            arsiv_sekmesi(veri, gecmis_log)
        return

    haric_norm = {_norm(x) for x in haric_liste}
    aktif_urunler = [u for u in orijinal_urun_sirasi if _norm(u) not in haric_norm]

    ay_sayisi = int(veri["Tarih"].dt.month.nunique())
    with st.sidebar:
        st.markdown("---")
        mevsim_kullan = st.checkbox("Mevsim terimlerini kullan", value=ay_sayisi >= 12, key="mevsim_kullan")

    veri["Y"] = veri["Adet"] * np.where(veri["Bitti"], 1 + bitti_duzeltme / 100.0, 1.0)

    hava = None
    if hava_kullan:
        hava = hava_getir(veri["Tarih"].min().date().isoformat(), (BUGUN + dt.timedelta(days=16)).isoformat())
    use_weather = False
    if hava is not None and veri["Tarih"].isin(hava["Tarih"]).mean() >= 0.7:
        veri = veri.merge(hava, on="Tarih", how="left")
        ay = veri["Tarih"].dt.month
        for c in ("sicaklik", "yagis"):
            veri[c] = veri[c].fillna(veri.groupby(ay)[c].transform("mean")).fillna(veri[c].mean())
        use_weather = True

    yillar = range(veri["Tarih"].min().year - 1, BUGUN.year + 3)
    cfg = {
        "kum": ozel_gun_kumeleri(yillar, ek_ozel),
        "menu": menu_tarihi,
        "t0": veri["Tarih"].min(),
        "use_weather": use_weather,
        "use_season": bool(mevsim_kullan),
    }
    X = ozellik_olustur(veri, cfg)
    gozlenmeyen_tum = {c for c in OLAY_ETIKET if X[c].sum() == 0}

    test = geriye_donuk_test(veri[["Tarih", "Y"]], X)
    mod, yontem = None, "Baz"
    alpha = 10.0
    if test is not None:
        yontem = min(test["mape"], key=test["mape"].get)
        alpha = test["alpha"]
        mod = fit_model(X, np.log(veri["Y"].to_numpy(float)), alpha)

    ozel_df = carpan_tablo_oku()
    ozel_carpan = {pd.Timestamp(t).date(): float(c) for t, c in zip(ozel_df["Tarih"], ozel_df["Carpan"])}
    ay_gozlem = veri["Tarih"].dt.month.value_counts().to_dict()
    veri_tarihleri = set(veri["Tarih"])

    def hava_degeri(tarih):
        if not use_weather:
            return 15.0, 0.0, "yok"
        s = hava[hava["Tarih"] == tarih]
        if len(s) and pd.notna(s["sicaklik"].iloc[0]):
            yag = s["yagis"].iloc[0]
            yag = 0.0 if pd.isna(yag) else float(yag)
            kaynak = "gerçekleşen" if tarih.date() < BUGUN else "tahmin"
            return float(s["sicaklik"].iloc[0]), yag, kaynak
        m = hava[hava["Tarih"].dt.month == tarih.month]["sicaklik"].mean()
        if pd.isna(m):
            m = veri["sicaklik"].mean()
        if pd.isna(m):
            m = 15.0
        return float(m), 0.0, "mevsim ortalaması"

    def oran_serisi(hedef_ts):
        bt = test["bt"]
        hs = hedef_ts.dayofweek >= 4
        maske = (bt["Tarih"].dt.dayofweek >= 4) == hs
        alt = bt[maske] if maske.sum() >= 15 else bt
        return alt["Gercek"] / alt[yontem]

    def carpan_hesapla(sv, hedef_ts):
        if test is None or len(test["bt"]) < 20:
            return 1.10, False
        return float(np.quantile(oran_serisi(hedef_ts), sv)), True

    def planla(td, sic=None, yag=None, kaynak=None, manuel=None):
        if sic is None:
            sic, yag, kaynak = hava_degeri(td)
        sec, model_t, baz_t, xr = tahmin_et(cfg, veri, mod, yontem, td, sic, yag)
        if manuel is None:
            manuel = ozel_carpan.get(td.date(), 1.0)
        beklenen = sec * manuel
        carpan, olculu = carpan_hesapla(seviye, td)
        olaylar = aktif_olaylar(td.date(), cfg["kum"])
        return {
            "tarih": td, "sic": sic, "yag": yag, "hava_kaynak": kaynak, "sec": sec, "model": model_t,
            "baz": baz_t, "xr": xr, "olaylar": olaylar,
            "gozlenmeyen": [o for o in olaylar if o in gozlenmeyen_tum], "manuel": manuel,
            "beklenen": beklenen, "carpan": carpan, "olculu": olculu,
            "uretim": int(round(beklenen * carpan)),
        }

    existing_logs = set(log_oku()["Hedef"])
    for vt in veri["Tarih"]:
        if vt not in existing_logs:
            p_ot = planla(vt)
            log_yaz(vt, p_ot["beklenen"], p_ot["uretim"], p_ot["baz"], yontem)

    bugun_ts = pd.Timestamp(BUGUN)
    yaklasan = [planla(bugun_ts + pd.Timedelta(days=i)) for i in range(1, 8)]
    yarin = yaklasan[0]
    if yarin["tarih"] not in veri_tarihleri:
        log_yaz(yarin["tarih"], yarin["beklenen"], yarin["uretim"], yarin["baz"], yontem)

    log_df = log_oku()
    gl = log_df.merge(veri[["Tarih", "Y"]], left_on="Hedef", right_on="Tarih", how="inner")
    if not gl.empty:
        gl["Hata"] = (gl["Beklenen"] - gl["Y"]) / gl["Y"]

    son_i = len(veri) - 1
    son_tarih = veri["Tarih"].iloc[son_i]
    son_gercek = float(veri["Y"].iloc[son_i])
    esl = log_df[log_df["Hedef"] == son_tarih]
    if len(esl):
        son_beklenen, son_kaynak = float(esl["Beklenen"].iloc[0]), "kayıtlı tahmin"
    else:
        son_beklenen = retro_tahmin(veri, X, son_i, yontem, alpha) * ozel_carpan.get(son_tarih.date(), 1.0)
        son_kaynak = "geriye dönük hesap"
    son_sapma = (son_gercek - son_beklenen) / son_beklenen if son_beklenen > 0 else 0.0

    uyarilar = []
    son_gun = son_tarih.date()
    if son_gun < BUGUN - dt.timedelta(days=1):
        uyarilar.append(("uyari", f"Son yüklenen veri **{son_gun:%d.%m.%Y}**; {(BUGUN - son_gun).days} günlük veri eksik."))
    elif son_gun < BUGUN:
        uyarilar.append(("bilgi", f"Bugünün ({BUGUN:%d.%m.%Y}) siparişleri henüz yüklenmedi."))

    if abs(son_sapma) * 100 > esik:
        uyarilar.append(("uyari", f"**{son_tarih:%d.%m.%Y}** gerçekleşen **{int(round(son_gercek))}**, "
                                  f"beklenen ~{int(round(son_beklenen))} ({son_kaynak}): %{abs(son_sapma) * 100:.0f} sapma."))

    with uyari_kutu:
        if uyarilar:
            for sev, m in uyarilar:
                {"hata": st.error, "uyari": st.warning, "bilgi": st.info}[sev](m)

    # ------------------------- SEKMELER -------------------------
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🚀 Üretim Planı", "📈 Model Performansı", "🎯 Tahmin Takip", "📝 Veri Girişi", "🔔 Uyarılar", "📂 Veri Arşivi"])

    with tab1:
        c1, c2 = st.columns([1, 2])
        with c1:
            varsayilan = max(BUGUN + dt.timedelta(days=1), son_gun + dt.timedelta(days=1))
            hedef_d = st.date_input("Üretim / tahmin tarihi", value=varsayilan)
            hedef = pd.Timestamp(hedef_d)

            hava_elle = st.checkbox("Hava durumunu elle gir")
            sic, yag, hava_kaynak = hava_degeri(hedef)
            if hava_elle:
                sic = st.slider("Sıcaklık (°C)", -5, 40, int(round(sic)))
                yag = 5.0 if st.checkbox("Yağmurlu", value=yag >= 1) else 0.0
                hava_kaynak = "elle"
            st.info(f"{GUN_TR[hedef.dayofweek]} | Hava: {sic:.0f}°C ({hava_kaynak})")

            kayitli = float(ozel_carpan.get(hedef.date(), 1.0))
            elle_carpan = st.number_input("Elle çarpan", min_value=0.3, max_value=4.0, value=kayitli, step=0.05)

        P = planla(hedef, sic, yag, hava_kaynak, manuel=elle_carpan)

        with c1:
            if hedef in veri_tarihleri:
                st.warning("Bu tarih için gerçek veri var.")
            elif st.button("💾 Bu tahmini takibe kaydet"):
                log_yaz(hedef, P["beklenen"], P["uretim"], P["baz"], yontem)
                st.session_state["_mesaj"] = f"{hedef:%d.%m.%Y} tahmini takip kaydına eklendi."
                st.rerun()

        with c2:
            m1, m2, m3 = st.columns(3)
            m1.metric("Önerilen toplam üretim", f"{P['uretim']} adet")
            m2.metric("Beklenen talep", f"{int(round(P['beklenen']))} adet")
            if P["olculu"]:
                oran = oran_serisi(hedef)
                lo, hi = np.quantile(oran, 0.1) * P["beklenen"], np.quantile(oran, 0.9) * P["beklenen"]
                m3.metric("Olası aralık (%80)", f"{int(round(lo))} - {int(round(hi))}")
            else:
                m3.metric("Olası aralık", "hesaplanamadı")

            st.markdown("### 🧠 Akıllı Tahmin ve Yorum Analizi")
            for y in yorum_uret(P, mod, P["gozlenmeyen"], use_weather):
                st.markdown(f"- {y}")

            st.markdown(f"* **Kullanılan yöntem:** `{YONTEM_AD[yontem]}`")
            st.markdown(f"* **Baz çizgi:** `{int(round(P['baz']))}` adet")

        st.markdown("### 📋 Çeşit Dağılımı (Aktif Ürünler)")
        if detay_tum.empty:
            st.info("Çeşit dağılımı için ürün bazlı veri bulunamadı.")
        else:
            tablo, grup_ad = urun_dagilimi(detay_tum, hedef, P["uretim"], pencere,
                                           orijinal_sira=orijinal_urun_sirasi, haric_urunler=haric_liste)
            if not tablo.empty:
                _df(tablo)
                st.caption(f"Toplam = {P['uretim']} adet.")

        st.markdown("### 🗓️ Önümüzdeki 7 gün")
        satirlar = []
        for i in range(7):
            td = hedef + pd.Timedelta(days=i)
            p = P if i == 0 else planla(td)
            satirlar.append({
                "Tarih": td.strftime("%d.%m.%Y"), "Gün": GUN_TR[td.dayofweek], "Hava": f"{p['sic']:.0f}°C",
                "Beklenen talep": int(round(p["beklenen"])), "Önerilen üretim": p["uretim"],
            })
        _df(pd.DataFrame(satirlar), hide_index=True)

    with tab2:
        st.subheader("Model Geriye Dönük Test Performansı")
        if test is not None:
            mm, bb = test["mape"], test["bias"]
            a, b, c = st.columns(3)
            a.metric("Baz çizgi hatası", f"%{mm['Baz'] * 100:.1f}")
            b.metric("Model hatası", f"%{mm['Model'] * 100:.1f}")
            c.metric("Karma hatası", f"%{mm['Karma'] * 100:.1f}")
            st.line_chart(test["bt"].rename(columns={"Gercek": "Gerçek", "Baz": "Baz çizgi"}).set_index("Tarih")[["Gerçek", "Model"]])

    with tab3:
        st.subheader("🎯 Kaydedilen Tahminler ve Gerçekleşen Karşılaştırması")
        st.caption("Sistemin ürettiği tahminlerin o günkü gerçek satış/siparişlerle kıyaslandığı ana takip ekranı.")
        
        log_df_guncel = log_oku()
        if log_df_guncel.empty:
            st.info("Henüz kaydedilmiş tahmin bulunmuyor.")
        else:
            takip_df = log_df_guncel.merge(veri[["Tarih", "Y"]], left_on="Hedef", right_on="Tarih", how="left")
            takip_df["Gerçekleşen"] = takip_df["Y"]
            takip_df["Sapma (Adet)"] = takip_df["Beklenen"] - takip_df["Gerçekleşen"]
            takip_df["Sapma (%)"] = ((takip_df["Beklenen"] - takip_df["Gerçekleşen"]) / takip_df["Gerçekleşen"] * 100).round(1)
            
            gerceklesen_var = takip_df[takip_df["Gerçekleşen"].notna()]
            if not gerceklesen_var.empty:
                t1, t2, t3 = st.columns(3)
                t1.metric("Toplam Kayıtlı Tahmin", len(takip_df))
                t2.metric("Verisi Girilmiş Gün Sayısı", len(gerceklesen_var))
                ortalama_hata = np.mean(np.abs(gerceklesen_var["Sapma (%)"]))
                t3.metric("Ortalama Mutlak Sapma", f"%{ortalama_hata:.1f}")

            st.markdown("---")
            goster_takip = takip_df[["Hedef", "Beklenen", "Uretim", "Gerçekleşen", "Sapma (Adet)", "Sapma (%)", "Yontem", "Kayit"]].copy()
            goster_takip["Hedef"] = goster_takip["Hedef"].dt.strftime("%d.%m.%Y")
            goster_takip["Beklenen"] = goster_takip["Beklenen"].round(0).astype("Int64")
            goster_takip["Uretim"] = goster_takip["Uretim"].round(0).astype("Int64")
            goster_takip["Gerçekleşen"] = goster_takip["Gerçekleşen"].round(0).astype("Int64")
            goster_takip["Sapma (Adet)"] = goster_takip["Sapma (Adet)"].round(0).astype("Int64")
            goster_takip["Sapma (%)"] = goster_takip["Sapma (%)"].map(lambda x: f"%{x:+.1f}" if pd.notna(x) else "Bekleniyor")
            
            _df(goster_takip.sort_values("Hedef", ascending=False), hide_index=True)
            st.download_button("⬇️ Tahmin Takip Verisini İndir", goster_takip.to_csv(index=False).encode("utf-8-sig"),
                               file_name="tahmin_takip.csv", mime="text/csv")

    with tab4:
        giris_sekmesi(aktif_urunler, detay_tum, meta_db)

    with tab5:
        st.subheader("🔔 Sistem Uyarıları ve Takip Özeti")
        if not gl.empty:
            st.line_chart(gl.set_index("Hedef")[["Beklenen", "Y"]].rename(columns={"Y": "Gerçek"}))
        else:
            st.info("Henüz karşılaştırma verisi yok.")

    with tab6:
        arsiv_sekmesi(veri, gecmis_log)


try:
    uygulama()
except Exception as ex:
    st.error(f"Hata: {ex}")
    st.exception(ex)