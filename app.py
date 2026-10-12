import datetime
import json
import os
from urllib.request import urlopen
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tatlıbak Üretim Planlama Asistanı", page_icon="🍰", layout="wide"
)

st.title("🍰 Tatlıbak Yaş Pasta Üretim & Sipariş Planlama Asistanı")

excel_file = "YAŞ PASTA EYLÜL SİPARİŞ.xlsx"
csv_archive_file = "daily_orders.csv"
csv_weather_file = "daily_weather.csv"
historical_excel_file = "gecmis_arsiv.xlsx"


@st.cache_data
def load_data():
  xls = pd.ExcelFile(excel_file)
  df = xls.parse("Sipariş List")
  return df


try:
  df_all = load_data()
except Exception as e:
  st.error(f"Excel dosyası yüklenirken hata oluştu: {e}")


@st.cache_data
def load_historical_data():
  if os.path.exists(historical_excel_file):
    try:
      xls_h = pd.ExcelFile(historical_excel_file)
      sheet_name = xls_h.sheet_names[0]
      df_h = xls_h.parse(sheet_name)
      return df_h, sheet_name
    except Exception as ex:
      return None, str(ex)
  return None, "Dosya bulunamadı"


df_history, hist_info = load_historical_data()

if os.path.exists(csv_archive_file):
  df_orders = pd.read_csv(csv_archive_file)
  df_orders["Tarih"] = pd.to_datetime(df_orders["Tarih"])
else:
  df_orders = pd.DataFrame(
      columns=["Tarih", "Pasta Çeşidi", "Gerçekleşen Sipariş"]
  )
  df_orders["Tarih"] = pd.to_datetime(df_orders["Tarih"])

if os.path.exists(csv_weather_file):
  df_weather = pd.read_csv(csv_weather_file)
  df_weather["Tarih"] = pd.to_datetime(df_weather["Tarih"])
else:
  df_weather = pd.DataFrame(
      columns=["Tarih", "Hava Durumu", "Sıcaklık (°C)"]
  )
  df_weather["Tarih"] = pd.to_datetime(df_weather["Tarih"])


# İstanbul için 3 günlük saatlik (17:00-21:00 akşam yoğunluğu) hava durumu tahmini çeken fonksiyon
@st.cache_data(ttl=3600)
def fetch_3day_evening_weather():
  forecast_days = []
  try:
    url = "https://wttr.in/Istanbul?format=j1"
    response = urlopen(url, timeout=5)
    data = json.loads(response.read().decode("utf-8"))

    for idx, day_data in enumerate(data.get("weather", [])[:3]):
      date_str = day_data.get("date")
      hourly_list = day_data.get("hourly", [])

      evening_temps = []
      evening_descs = []

      for h in hourly_list:
        try:
          hour_val = int(h["time"]) // 100
          if 17 <= hour_val <= 21:
            evening_temps.append(int(h["tempC"]))
            evening_descs.append(h["weatherDesc"][0]["value"])
        except:
          pass

      if evening_temps:
        avg_t = int(sum(evening_temps) / len(evening_temps))
        common_d = max(set(evening_descs), key=evening_descs.count)
      else:
        avg_t = int(day_data.get("avgtempC", 20))
        common_d = "Açık"

      tr_desc = {
          "Sunny": "Açık",
          "Clear": "Açık",
          "Partly cloudy": "Parçalı Bulutlu",
          "Cloudy": "Bulutlu",
          "Overcast": "Çok Bulutlu",
          "Light rain": "Hafif Yağmurlu",
          "Patchy rain possible": "Hafif Yağmurlu",
          "Moderate rain": "Yağmurlu",
          "Heavy rain": "Çok Yağmurlu",
      }.get(common_d, common_d)

      day_name_map = {
          0: "Bugün",
          1: "Yarın",
          2: "2 Gün Sonra (Satış Günü)",
      }
      forecast_days.append({
          "Gün": day_name_map.get(idx, f"+{idx} Gün"),
          "Tarih": date_str,
          "Akşam Sıcaklık (17:00-21:00)": f"{avg_t}°C",
          "Hava Durumu": tr_desc,
      })
  except Exception:
    pass

  return forecast_days


@st.cache_data(ttl=86400)
def fetch_historical_weather(target_date_str):
  try:
    url = f"https://wttr.in/Istanbul?date={target_date_str}&format=j1"
    response = urlopen(url, timeout=4)
    data = json.loads(response.read().decode("utf-8"))
    weather_info = data.get("weather", [{}])[0]
    avg_temp = weather_info.get("avgtempC", "Bilinmiyor")
    desc_list = weather_info.get("hourly", [{}])[0].get("weatherDesc", [{}])
    desc = desc_list[0].get("value", "Parçalı Bulutlu") if desc_list else "Normal"
    return f"{avg_temp}°C, {desc}"
  except Exception:
    return "Mevsim Normalleri (18°C - Parçalı Bulutlu)"


three_day_forecast = fetch_3day_evening_weather()

# SEKME YAPISI
tab1, tab2, tab3, tab4 = st.tabs([
    "🚀 Üretim Tahmini & Hesaplama",
    "📝 Günlük Sipariş Girişi",
    "📂 Arşiv & Matris Tablo",
    "📜 Geçmiş Yıllar Arşivi",
])

with tab1:
  st.subheader("Günlük Veri Girişi ve Yarınki Üretim Planı")

  # --- 3 GÜNLÜK HAVA DURUMU BİLGİ PANELİ ---
  if three_day_forecast:
    st.markdown("🌤️ **İstanbul 3 Günlük Akşam Yoğunluk (17:00 - 21:00) Hava Rehberi**")
    f_cols = st.columns(len(three_day_forecast))
    for i, day_info in enumerate(three_day_forecast):
      with f_cols[i]:
        st.metric(
            label=f"{day_info['Gün']} ({day_info['Tarih']})",
            value=day_info["Akşam Sıcaklık (17:00-21:00)"],
            delta=day_info["Hava Durumu"],
        )
    st.markdown("---")

  col1, col2 = st.columns([1, 2])

  with col1:
    selected_date = st.date_input("İşlem Tarihi")
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
    st.info(f"Seçilen Gün: **{days_tr.get(day_of_week, '')}**")

    st.markdown("---")
    st.markdown("🌤️ **Üretim / Satış Günü Hava Koşulları Seçimi**")

    weather_condition = st.selectbox(
        "Hava Durumu Etkisi",
        [
            "Açık / Güneşli (Normal)",
            "Bulutlu / Hafif Yağmurlu (Talep Artışı +%)",
            "Çok Yağmurlu / Fırtınalı (Talep Düşüşü -%)",
            "Aşırı Sıcak / Güneşli (Ağır Tatlı Azalır)",
        ],
    )

    temperature = st.slider(
        "Beklenen Sıcaklık (°C) [Yukarıdaki tablodan bakabilirsiniz]",
        min_value=-5,
        max_value=40,
        value=20,
        step=1,
    )

    special_day = st.selectbox(
        "Özel Gün Durumu",
        [
            "Normal Gün",
            "Yılbaşı ve Arifesi",
            "Sevgililer Günü (14 Şubat)",
            "Anneler Günü",
            "Resmi Tatil / Bayram (29 Ekim, 23 Nisan vb.)",
            "Hafta Sonu / Cuma Akşamı Yoğunluğu",
        ],
    )
    buffer_pct = st.selectbox(
        "Stok / Güvenlik Payı Ekle (%)", [0, 5, 10, 15, 20, 25], index=2
    )
    hesapla_button = st.button("🚀 Yarın İçin Tahmini Üretim Listesi Oluştur")

  with col2:
    if hesapla_button:
      environmental_multiplier = 1.0
      if "Bulutlu" in weather_condition:
        environmental_multiplier += 0.05
      elif "Çok Yağmurlu" in weather_condition:
        environmental_multiplier -= 0.08
      elif "Aşırı Sıcak" in weather_condition or temperature > 30:
        environmental_multiplier -= 0.05

      special_multiplier = 1.0
      if "Yılbaşı" in special_day:
        special_multiplier = 1.40
      elif "Sevgililer" in special_day or "Anneler" in special_day:
        special_multiplier = 1.25
      elif "Resmi Tatil" in special_day:
        special_multiplier = 1.15
      elif "Hafta Sonu" in special_day:
        special_multiplier = 1.10

      raw_products = df_all.iloc[1:, 0].dropna().astype(str).values
      valid_products = [
          p for p in raw_products if "TOPLAM" not in p.strip().upper()
      ]

      numeric_df = df_all.iloc[1:, 1:].apply(pd.to_numeric, errors="coerce")
      valid_numeric_df = numeric_df.iloc[: len(valid_products)]

      base_vals = (
          valid_numeric_df.mean(axis=1).fillna(0)
          if not valid_numeric_df.empty
          else pd.Series([0] * len(valid_products))
      )

      combined_multiplier = (
          environmental_multiplier
          * special_multiplier
          * (1 + (buffer_pct / 100.0))
      )
      forecast_vals = (base_vals * combined_multiplier).round().astype(int)
      target_col_name = (
          "Tahmini Üretim (Net / Stoksuz)"
          if buffer_pct == 0
          else f"Tahmini Üretim (%{buffer_pct} Stoklu + Akıllı Etki)"
      )

      forecast_df = pd.DataFrame({
          "Pasta Çeşidi": valid_products,
          target_col_name: forecast_vals,
      })
      forecast_df = forecast_df[
          ~forecast_df["Pasta Çeşidi"]
          .astype(str)
          .str.contains("BATON", na=False)
      ]

      subtotal = int(forecast_df[target_col_name].sum())

      total_row_df = pd.DataFrame({
          "Pasta Çeşidi": ["TOPLAM PASTA SAYISI"],
          target_col_name: [subtotal],
      })
      forecast_df = pd.concat([forecast_df, total_row_df], ignore_index=True)

      st.metric(
          label="Önerilen Toplam Pasta Üretimi",
          value=f"{subtotal} Adet",
      )
      st.dataframe(forecast_df, use_container_width=True, height=450)

      # --- PROFESYONEL VE NET YÖNETİCİ ÖZETİ ---
      historical_insight_text = (
          "Geçmiş yıl arşivi (`gecmis_arsiv.xlsx`) okunamadı."
      )
      if df_history is not None:
        try:
          target_month = selected_date.month
          target_day = selected_date.day
          target_weekday = selected_date.weekday()
          week_of_month = (target_day - 1) // 7 + 1

          date_col = df_history.columns[0]
          df_history["Parsed_Date"] = pd.to_datetime(
              df_history[date_col], errors="coerce"
          )
          df_history["Weekday"] = df_history["Parsed_Date"].dt.weekday
          df_history["Week_Of_Month"] = (
              df_history["Parsed_Date"].dt.day - 1
          ) // 7 + 1