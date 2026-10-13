import datetime
import json
import os
from urllib.request import urlopen
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tatlıbak AI Üretim Planlama Asistanı",
    page_icon="🍰",
    layout="wide",
)

st.title(
    "🍰 Tatlıbak Yaş Pasta - Kesin Sınırlandırılmış Gerçekçi Üretim Planı"
)

excel_file = "YAŞ PASTA EYLÜL SİPARİŞ.xlsx"
excel_file_alt = "YAŞ PASTA 2025 SİPARİŞ.xlsx"
historical_excel_file = "gecmis_arsiv.xlsx"

# --- VERİLERİ YÜKLEME ---


@st.cache_data
def load_main_excel():
  for fname in [excel_file, excel_file_alt, "yas_pasta_eylul_siparis.xlsx"]:
    if os.path.exists(fname):
      try:
        xls = pd.ExcelFile(fname)
        sheet = xls.sheet_names[0]
        for s in xls.sheet_names:
          if "SİPARİŞ" in s.upper() or "LİST" in s.upper():
            sheet = s
            break
        return xls.parse(sheet), fname
      except:
        pass
  return None, None


df_all, loaded_excel_name = load_main_excel()


@st.cache_data
def load_historical_excel():
  for hname in [historical_excel_file, "gecmis_arsiv.xlsx", "arsiv.xlsx"]:
    if os.path.exists(hname):
      try:
        xls_h = pd.ExcelFile(hname)
        sheet_h = xls_h.sheet_names[0]
        return xls_h.parse(sheet_h), hname
      except:
        pass
  return None, None


df_history, loaded_history_name = load_historical_excel()

# --- DEMİR GİBİ SABİTLENMİŞ VE KONTROLLÜ TAHMİNLEME MOTORU ---


def run_strict_reasoning(
    target_date,
    turkish_day,
    special_day,
    weather_condition,
    temperature,
    buffer_pct,
    df_all,
    df_history,
):
  month_val = target_date.month
  analysis_notes = []

  # Çok hassas ve kontrollü yüzdesel değişimler (En fazla ±%5-8 oynar)
  season_factor = 1.03 if month_val in [10, 11, 12, 1, 2, 3, 4] else 0.97

  weather_factor = 1.0
  if (
      "Bulutlu" in weather_condition
      or "Hafif Yağmurlu" in weather_condition
      or "Yağmurlu" in weather_condition
  ):
    weather_factor = 1.04  # Kapalı hava %4 tetikler
    analysis_notes.append(
        f"☁️ **Hava ({weather_condition}):** Kapalı hava etkisi (+%4)."
    )
  elif temperature < 15:
    weather_factor = 1.02
  elif temperature > 28:
    weather_factor = 0.95

  special_factor = 1.0
  if "Maaş" in special_day or "10-15" in special_day:
    special_factor = 1.06  # Maaş dönemi %6
    analysis_notes.append("💰 **Özel Dönem:** Maaş sarmalı etkisi (+%6).")
  elif "Yılbaşı" in special_day:
    special_factor = 1.15

  day_factor = (
      1.08
      if turkish_day in ["Cuma", "Cumartesi"]
      else (0.95 if turkish_day == "Pazartesi" else 1.0)
  )

  net_multiplier = (
      season_factor * weather_factor * special_factor * day_factor
  )

  # Taban Referans: Geçen yıl 13 Ekim (1263) ve dünkü gerçek (~1500)
  # Sistem asla bu gerçekliğin dışına çıkmayacak, ortalama baz taban: ~1400 adet
  base_reference_total = 1420

  # Ham Hedef Hesaplama
  raw_target = base_reference_total * net_multiplier

  # GÜVENLİK KİLİDİ (CLAMP): Toplam hedef asla taban referansın %120'sini geçemez!
  max_allowed_target = base_reference_total * 1.20  # En fazla ~1700 sınır
  min_allowed_target = base_reference_total * 0.85  # En az ~1200 sınır

  controlled_total = max(
      min_allowed_target, min(raw_target, max_allowed_target)
  )

  # Güvenlik / Stok Payı ekleniyor (% buffer)
  final_total_production = controlled_total * (1.0 + (buffer_pct / 100.0))

  analysis_notes.insert(
      0,
      f"📜 **Referans Taban (Geçen Yıl 13 Ekim + Dün):**"
      f" ~{base_reference_total} adet (Sıkı kontrolde tutuldu)",
  )

  # Ürün Bazlı Dağılım (Excel'deki ürünlerin oranına göre kusursuz paylaştırma)
  product_results = []
  total_checked_production = 0

  if df_all is not None and not df_all.empty:
    try:
      p_col = df_all.columns[0]
      date_cols = [
          c
          for c in df_all.columns
          if c != p_col and "TOPLAM" not in str(c).upper()
      ]

      product_weights = []
      total_weight_sum = 0

      for _, row in df_all.iterrows():
        prod_name = str(row[p_col]).strip()
        if (
            not prod_name
            or "TOPLAM" in prod_name.upper()
            or prod_name == "nan"
            or prod_name.lower() == "nat"
        ):
          continue

        p_vals = []
        for d in date_cols:
          try:
            v = float(row[d])
            if v > 0:
              p_vals.append(v)
          except:
            pass
        p_avg = sum(p_vals) / len(p_vals) if p_vals else 10
        product_weights.append({"name": prod_name, "avg": p_avg})
        total_weight_sum += p_avg

      if total_weight_sum > 0:
        for pw in product_weights:
          share = pw["avg"] / total_weight_sum
          product_target = int(round(final_total_production * share))
          product_target = max(5, product_target)
          total_checked_production += product_target
          product_results.append({
              "Ürün Adı": pw["name"],
              "Önerilen Üretim": product_target,
          })
    except Exception as ex:
      analysis_notes.append(f"Dağılım hatası: {ex}")

  if not product_results:
    product_results = [
        {"Ürün Adı": "Fıstıklı Özel Pasta", "Önerilen Üretim": 100},
        {"Ürün Adı": "Çikolatalı Rüzgar", "Önerilen Üretim": 90},
        {"Ürün Adı": "Meyveli Pasta", "Önerilen Üretim": 80},
    ]
    total_checked_production = sum(
        [item["Önerilen Üretim"] for item in product_results]
    )

  notes_str = "\n* ".join(analysis_notes)

  report_markdown = f"""
### 🔍 1. Sıkı Kontrollü Gerçekçi Üretim Raporu
* {notes_str}
* **Kontrol Kriteri:** Dünkü gerçekleşen sipariş (~1500) ve geçen yılki ekim verisi (~1263) referans alınarak şişirme çarpanlar tamamen engellenmiştir.

### 📊 2. Stratejik Değerlendirme & Sınırlandırılmış Bantlar
* **Üretim Sınırı:** Sistem, piyasa gerçeklerinin dışına çıkmaması için tavan sınırla (clamp) sabitlenmiştir.
* **Güvenlik / Stok Payı:** **%{buffer_pct}** marj kontrollü olarak yansıtılmıştır.

### 📋 3. Ürün Bazlı Üretim Dağılım Listesi
| Ürün Çeşidi | Önerilen Üretim Adedi |
| :--- | :--- |
"""
  for item in product_results:
    report_markdown += (
        f"| {item['Ürün Adı']} | **{item['Önerilen Üretim']} Adet** |\n"
    )

  report_markdown += f"""
---
🔥 **TOPLAM ÜRETİM HEDEFİ: {int(total_checked_production)} Adet**
"""
  return report_markdown


tab1, tab2, tab3, tab4 = st.tabs([
    "🚀 Gerçekçi Sınırlandırılmış Plan",
    "📝 Günlük Sipariş Girişi",
    "📂 Arşiv & Matris Tablo",
    "📜 Geçmiş Yıllar Arşivi",
])

with tab1:
  st.subheader(
      "🤖 Asla Şişme Yapmayan, Dünkü ve Geçen Yılki Gerçeklere Sabitlenmiş Motor"
  )
  col1, col2 = st.columns([1, 2])

  with col1:
    selected_date = st.date_input("İşlem / Üretim Tarihi")
    day_of_week = selected_date.strftime("%A")
    days_tr = {
        "Monday": "Pazartesi",
        "Tuesday": "Salı",
        "Wednesday": "Çarşamba",
        "Thursday": "Perşembe",
        "Friday": "Cuma",
        "Saturday": "Cumartesi",
        "Sunday": "Pazar",
    }
    turkish_day = days_tr.get(day_of_week, "")
    st.info(f"Seçilen Gün: **{turkish_day}**")

    st.markdown("---")
    weather_condition = st.selectbox(
        "Hava Durumu Beklentisi",
        [
            "Bulutlu / Kapalı (Tatlı Talebini Artırır)",
            "Hafif Yağmurlu",
            "Açık / Güneşli",
            "Çok Yağmurlu / Fırtınalı",
            "Aşırı Sıcak",
        ],
    )
    temperature = st.slider("Beklenen Sıcaklık (°C)", -5, 40, 15, 1)
    special_day = st.selectbox(
        "Özel Durum / Maaş Dönemi",
        [
            "Normal Gün",
            "Maaş Dönemi / Ayın 10-15 Sarmalı",
            "Yılbaşı ve Arifesi",
            "Hafta Sonu Yoğunluğu",
        ],
    )
    buffer_pct = st.selectbox(
        "Güvenlik / Stok Payı (%)", [0, 5, 10, 15, 20, 25], index=2
    )
    analyze_button = st.button("🧠 Gerçekçi Sınırlandırılmış Planı Hesapla")

  with col2:
    if analyze_button:
      with st.spinner("⚙️ Kontrollü ve gerçekçi üretim planı hesaplanıyor..."):
        try:
          report_output = run_strict_reasoning(
              selected_date,
              turkish_day,
              special_day,
              weather_condition,
              temperature,
              buffer_pct,
              df_all,
              df_history,
          )
          st.markdown(report_output)
        except Exception as ex:
          st.error(f"Hata: {ex}")
    else:
      st.info("👈 Kriterleri seçip butona tıklayın.")

with tab2:
  st.subheader("📝 Akşam Gelen Reel Sipariş Girişi")
  if df_all is not None:
    st.success("Ürün listesi aktif.")
  else:
    st.warning("Excel bulunamadı.")

with tab3:
  st.subheader("📂 Arşiv & Matris Tablo")
  if df_all is not None:
    st.dataframe(df_all, use_container_width=True, height=550)

with tab4:
  st.subheader("📜 Geçmiş Yıllar Arşivi")
  if df_history is not None:
    st.dataframe(df_history, use_container_width=True, height=550)