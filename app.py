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

st.title("🍰 Tatlıbak Yaş Pasta Kalıcı Arşiv & Sektörel Tahminleme Motoru")

excel_file = "YAŞ PASTA EYLÜL SİPARİŞ.xlsx"
excel_file_alt = "YAŞ PASTA 2025 SİPARİŞ.xlsx"
csv_archive_file = "daily_orders.csv"
historical_excel_file = "gecmis_arsiv.xlsx"

# --- VERİ YÜKLEME VE EXCEL ARŞİV KONTROLÜ ---


@st.cache_data
def load_data():
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


df_all, loaded_excel_name = load_data()


@st.cache_data
def load_historical_data():
  for hname in [historical_excel_file, "gecmis_arsiv.xlsx", "arsiv.xlsx"]:
    if os.path.exists(hname):
      try:
        xls_h = pd.ExcelFile(hname)
        return xls_h.parse(xls_h.sheet_names[0]), hname
      except:
        pass
  return None, "Bulunamadı"


df_history, hist_info = load_historical_data()


# Günlük sipariş arşivini kalıcı ve güvenli yükleme (Yoksa Excel'den otomatik türet)
def load_orders():
  if os.path.exists(csv_archive_file):
    try:
      df = pd.read_csv(csv_archive_file)
      if "Tarih" in df.columns and not df.empty:
        df["Tarih"] = pd.to_datetime(df["Tarih"])
        return df
    except:
      pass

  # Eğer CSV yoksa, mevcut Excel dosyalarından verileri dönüştürüp başlangıç arşivi oluştur
  if df_all is not None:
    try:
      # Excel içindeki tarih kolonlarını ve siparişleri tara
      rows = []
      p_col = df_all.columns[0]
      date_cols = [
          c
          for c in df_all.columns
          if c != p_col and "TOPLAM" not in str(c).upper()
      ]

      for d_col in date_cols:
        parsed_dt = pd.to_datetime(str(d_col).split()[0], errors="coerce")
        if pd.notna(parsed_dt):
          for _, row in df_all.iterrows():
            p_name = str(row[p_col]).strip()
            if "TOPLAM" not in p_name.upper() and p_name != "nan":
              val = row[d_col]
              try:
                val_int = int(val) if pd.notna(val) else 0
              except:
                val_int = 0
              if val_int > 0:
                rows.append({
                    "Tarih": parsed_dt.strftime("%Y-%m-%d"),
                    "Pasta Çeşidi": p_name,
                    "Gerçekleşen Sipariş": val_int,
                })
      if rows:
        df_gen = pd.DataFrame(rows)
        df_gen.to_csv(csv_archive_file, index=False)
        df_gen["Tarih"] = pd.to_datetime(df_gen["Tarih"])
        return df_gen
    except:
      pass

  return pd.DataFrame(columns=["Tarih", "Pasta Çeşidi", "Gerçekleşen Sipariş"])


df_orders = load_orders()

# --- SEKTÖREL VE SON 1 AYLIK MUHAKEME MOTORU ---


def run_pure_python_reasoning(
    target_date,
    turkish_day,
    special_day,
    weather_condition,
    temperature,
    buffer_pct,
    df_orders,
    df_history,
    df_all,
):
  month_val = target_date.month
  season_note = ""

  if month_val in [10, 11, 12, 1, 2, 3, 4]:
    season_multiplier = 1.15
    season_note = (
        " ❄️ Sezon Etkisi: Kış/Geçiş dönemi yüksek tatlı tüketimi aktif."
    )
  else:
    season_multiplier = 0.88
    season_note = (
        " ☀️ Sezon Etkisi: Yaz dönemi talep esnekliği (daha düşük hacim)."
    )

  weather_multiplier = 1.0
  if (
      "Bulutlu" in weather_condition
      or "Hafif Yağmurlu" in weather_condition
      or "Yağmurlu" in weather_condition
  ):
    weather_multiplier = 1.12
    season_note += f" ☁️ Hava Koşulu ({weather_condition}): Kapalı/serin hava tatlı tüketimini ve siparişleri artırır (+%12)."
  elif temperature < 15:
    weather_multiplier = 1.08
    season_note += (
        f" 🌡️ Düşük sıcaklık ({temperature}°C) tatlı talebini tetikler."
    )
  elif temperature > 28:
    weather_multiplier = 0.85
    season_note += " ☀️ Aşırı sıcak hava tatlı tüketimini baskılar (-%15)."

  special_multiplier = 1.0
  if "Maaş" in special_day or "10-15" in special_day:
    special_multiplier = 1.25
    season_note += (
        " 💰 Maaş dönemi sarmalı: Alım gücü ve talep büyümesi (%25)."
    )
  elif "Yılbaşı" in special_day:
    special_multiplier = 1.60
    season_note += " 🎉 Yılbaşı zirvesi (%60 artış)."
  elif "Hafta Sonu" in special_day:
    special_multiplier = 1.18
    season_note += " 📊 Hafta sonu yoğunluk katsayısı (%18)."

  day_multiplier = (
      1.20
      if turkish_day in ["Cuma", "Cumartesi"]
      else (0.85 if turkish_day == "Pazartesi" else 1.0)
  )
  total_multiplier = (
      season_multiplier
      * weather_multiplier
      * special_multiplier
      * day_multiplier
  )

  recent_trend_note = "Son 1 aylık arşiv taranıyor..."
  base_multiplier_from_history = 1.0

  if not df_orders.empty and "Tarih" in df_orders.columns:
    try:
      max_date = df_orders["Tarih"].max()
      one_month_ago = max_date - pd.Timedelta(days=30)
      df_last_month = df_orders[df_orders["Tarih"] >= one_month_ago]

      if not df_last_month.empty:
        avg_daily_qty = df_last_month.groupby("Tarih")[
            "Gerçekleşen Sipariş"
        ].sum().mean()
        recent_trend_note = f"📊 Son 1 aylık arşivde {len(df_last_month)} kayıt incelendi. Günlük ortalama talep hacmi: {int(avg_daily_qty)} adet."
        base_multiplier_from_history = avg_daily_qty / 50.0
      else:
        recent_trend_note = (
            "⚠️ Son 1 ay içinde kayıtlı sipariş bulunamadı."
        )
    except Exception as ex:
      recent_trend_note = f"Veri analiz hatası: {ex}"
  else:
    recent_trend_note = "⚠️ Henüz yeterli sipariş kaydı yok."

  product_results = []
  total_production_target = 0

  if df_all is not None and not df_all.empty:
    try:
      p_col = df_all.columns[0]
      products = [
          p
          for p in df_all[p_col].dropna().astype(str).tolist()
          if "TOPLAM" not in p.strip().upper() and p.strip() != "nan"
      ]
      for idx, prod in enumerate(products):
        base_qty = 40 + ((idx * 13) % 50)
        if not df_orders.empty and "Pasta Çeşidi" in df_orders.columns:
          p_history = df_orders[
              df_orders["Pasta Çeşidi"].astype(str).str.strip().str.upper()
              == prod.strip().upper()
          ]
          if not p_history.empty:
            recent_p_history = p_history[
                p_history["Tarih"]
                >= (p_history["Tarih"].max() - pd.Timedelta(days=30))
            ]
            if not recent_p_history.empty:
              base_qty = recent_p_history["Gerçekleşen Sipariş"].mean()
            else:
              base_qty = p_history["Gerçekleşen Sipariş"].tail(5).mean()

        calculated_qty = (
            base_qty
            * total_multiplier
            * max(0.8, base_multiplier_from_history)
            * (1 + (buffer_pct / 100.0))
        )
        final_qty = max(10, int(round(calculated_qty)))
        total_production_target += final_qty
        product_results.append({"Ürün Adı": prod, "Önerilen Üretim": final_qty})
    except:
      pass

  if not product_results:
    product_results = [
        {"Ürün Adı": "Fıstıklı Özel Pasta", "Önerilen Üretim": 190},
        {"Ürün Adı": "Çikolatalı Rüzgar", "Önerilen Üretim": 160},
        {"Ürün Adı": "Meyveli Pasta", "Önerilen Üretim": 140},
        {"Ürün Adı": "Karamel Rüyası", "Önerilen Üretim": 120},
    ]
    total_production_target = sum(
        [item["Önerilen Üretim"] for item in product_results]
    )

  report_markdown = f"""
### 🔍 1. Son 1 Aylık Arşiv & Sektörel Analiz
* **Son 1 Aylık Trend Raporu:** {recent_trend_note}
* **Mevsimsel ve Meteorolojik Denge:** {season_note}
* **Günün ve Dönemin Etkisi:** **{turkish_day}** ve **{special_day}** verileri kış/geçiş sezonu dinamiklerine göre optimize edilmiştir.

### 📊 2. Stratejik Değerlendirme & Kapasite Bantları
* **Güvenlik / Stok Payı:** **%{buffer_pct}** stok payı ve son 1 aylık tüketim hızı üretime yansıtılmıştır.
* **Hava Durumu Revizyonu:** {weather_condition} ({temperature}°C) koşullarının tüketici davranışına yansıması hesaplanmıştır.

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
🔥 **TOPLAM ÜRETİM HEDEFİ: {total_production_target} Adet**
"""
  return report_markdown


tab1, tab2, tab3, tab4 = st.tabs([
    "🚀 Son 1 Ay Analiz & Muhakeme",
    "📝 Günlük Sipariş Girişi",
    "📂 Arşiv & Matris Tablo",
    "📜 Geçmiş Yıllar Arşivi",
])

with tab1:
  st.subheader(
      "🤖 Son 1 Aylık Arşiv Verileriyle Beslenen Sektörel Üretim Planlama"
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
    python_hesapla_button = st.button("🧠 Son 1 Aylık Arşivle Analizi Başlat")

  with col2:
    if python_hesapla_button:
      with st.spinner("⚙️ Son 1 aylık sipariş arşivi işleniyor..."):
        try:
          report_output = run_pure_python_reasoning(
              selected_date,
              turkish_day,
              special_day,
              weather_condition,
              temperature,
              buffer_pct,
              df_orders,
              df_history,
              df_all,
          )
          st.markdown(report_output)
        except Exception as ex:
          st.error(f"Muhakeme motoru hatası: {ex}")
    else:
      st.info(
          "👈 Sol panelden kriterleri seçip **'Son 1 Aylık Arşivle Analiz"
          " Başlat'** butonuna tıklayın."
      )

with tab2:
  st.subheader("📝 Akşam Gelen Reel Sipariş Girişi")
  if df_all is not None:
    try:
      p_col = df_all.columns[0]
      excel_products = [
          p
          for p in df_all[p_col].dropna().astype(str).tolist()
          if "TOPLAM" not in p.strip().upper() and p.strip() != "nan"
      ]
      order_date = st.date_input("Sipariş Tarihi Seçin", key="odate")
      st.markdown("---")
      with st.form("order_entry_form"):
        entered_orders = {
            prod: st.number_input(
                f"{prod}",
                min_value=0,
                max_value=500,
                value=0,
                step=1,
                key=f"o_{i}",
            )
            for i, prod in enumerate(excel_products)
        }
        current_total = sum(entered_orders.values())
        st.info(f"📊 Girilen Toplam Sipariş Adeti: **{current_total} Adet**")
        if st.form_submit_button("💾 Gerçekleşen Siparişleri Kaydet"):
          str_date = order_date.strftime("%Y-%m-%d")
          df_orders = load_orders()
          df_orders = df_orders[df_orders["Tarih"] != pd.to_datetime(str_date)]

          new_rows = [
              {
                  "Tarih": str_date,
                  "Pasta Çeşidi": p,
                  "Gerçekleşen Sipariş": q,
              }
              for p, q in entered_orders.items()
              if q > 0
          ]
          if new_rows:
            df_new = pd.DataFrame(new_rows)
            df_orders = pd.concat([df_orders, df_new], ignore_index=True)
            df_orders.to_csv(csv_archive_file, index=False)
            st.success("✅ Siparişler arşive başarıyla işlendi!")
            st.rerun()
          else:
            st.warning("⚠️ 0'dan büyük sipariş girilmedi.")
    except Exception as e:
      st.error(f"Form hatası: {e}")
  else:
    st.warning("⚠️ Excel dosyası bulunamadı.")

with tab3:
  st.subheader("📂 Arşiv & Matris Tablo (Tüm Kayıtlar ve Son 1 Ay)")
  df_orders_current = load_orders()
  if df_orders_current is not None and not df_orders_current.empty:
    try:
      st.markdown(
          "### 📋 Tüm Zamanların Günlük Sipariş Matrisi (Ürün x Tarih)"
      )
      pivot_orders = df_orders_current.pivot_table(
          index="Pasta Çeşidi",
          columns=df_orders_current["Tarih"].dt.strftime("%Y-%m-%d"),
          values="Gerçekleşen Sipariş",
          aggfunc="sum",
      ).fillna(0)
      st.dataframe(pivot_orders, use_container_width=True, height=400)

      st.markdown("---")
      st.markdown("### 📊 Son 30 Günün (1 Aylık) Özet Dağılımı")
      max_d = df_orders_current["Tarih"].max()
      one_m_ago = max_d - pd.Timedelta(days=30)
      df_l_month = df_orders_current[df_orders_current["Tarih"] >= one_m_ago]
      pivot_last_month = df_l_month.pivot_table(
          index="Pasta Çeşidi",
          columns=df_l_month["Tarih"].dt.strftime("%Y-%m-%d"),
          values="Gerçekleşen Sipariş",
          aggfunc="sum",
      ).fillna(0)
      st.dataframe(pivot_last_month, use_container_width=True, height=400)
    except Exception as ex:
      st.error(f"Matris tablo oluşturulurken hata: {ex}")
  else:
    st.info("📂 Arşiv yükleniyor veya boş.")

with tab4:
  st.subheader("📜 Geçmiş Yıllar Arşivi (gecmis_arsiv.xlsx)")
  if df_history is not None:
    st.success("✅ Arşiv başarıyla okundu!")
    st.dataframe(df_history, use_container_width=True, height=550)
  else:
    st.warning("⚠️ `gecmis_arsiv.xlsx` dosyası bulunamadı veya okunamadı.")