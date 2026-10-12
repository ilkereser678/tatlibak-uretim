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


@st.cache_data
def load_data():
  xls = pd.ExcelFile(excel_file)
  df = xls.parse("Sipariş List")
  return df


try:
  df_all = load_data()
except Exception as e:
  st.error(f"Excel dosyası yüklenirken hata oluştu: {e}")

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


# Harici paket istemeyen, doğrudan web servisinden İstanbul hava durumu çekme fonksiyonu
@st.cache_data(ttl=3600)
def fetch_live_weather():
  try:
    url = "https://wttr.in/Istanbul?format=j1"
    response = urlopen(url, timeout=5)
    data = json.loads(response.read().decode("utf-8"))
    current = data["current_condition"][0]
    temp = int(current["temp_C"])
    desc = current["weatherDesc"][0]["value"]
    return temp, desc
  except Exception:
    return 20, "Clear"


live_temp, live_desc = fetch_live_weather()

# Türkçe hava durumu çevirisi
weather_translation = {
    "Sunny": "Açık",
    "Clear": "Açık",
    "Partly cloudy": "Parçalı Bulutlu",
    "Cloudy": "Bulutlu",
    "Overcast": "Çok Bulutlu",
    "Light rain": "Hafif Yağmurlu",
    "Patchy rain possible": "Hafif Yağmurlu",
    "Moderate rain": "Yağmurlu",
    "Heavy rain": "Çok Yağmurlu",
}
tr_live_desc = weather_translation.get(live_desc, "Açık / Güneşli (Normal)")

tab1, tab2, tab3 = st.tabs([
    "🚀 Üretim Tahmini & Hesaplama",
    "📝 Günlük Sipariş Girişi",
    "📂 Arşiv & Matris Tablo",
])

with tab1:
  st.subheader("Günlük Veri Girişi ve Yarınki Üretim Planı")
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
    st.markdown("🌤️ **İstanbul Hava & Çevre Koşulları (Canlı Otomatik)**")

    weather_condition = st.selectbox(
        "Hava Durumu Etkisi",
        [
            "Açık / Güneşli (Normal)",
            "Bulutlu / Hafif Yağmurlu (Talep Artışı +%)",
            "Çok Yağmurlu / Fırtınalı (Talep Düşüşü -%)",
            "Aşırı Sıcak / Güneşli (Ağır Tatlı Azalır)",
        ],
        index=(
            1
            if "Bulut" in tr_live_desc
            else (2 if "Yağmur" in tr_live_desc else 0)
        ),
    )

    temperature = st.slider(
        "Beklenen Sıcaklık (°C)",
        min_value=-5,
        max_value=40,
        value=int(live_temp),
        step=1,
    )
    st.caption(
        f"🌐 İstanbul anlık web servis verisi: **{live_temp}°C, {live_desc}**"
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

      products = df_all.iloc[1:, 0].values
      numeric_df = df_all.iloc[1:, 1:].apply(pd.to_numeric, errors="coerce")

      base_vals = (
          numeric_df.mean(axis=1).fillna(0)
          if not numeric_df.empty
          else pd.Series([0] * len(products))
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
          "Pasta Çeşidi": products,
          target_col_name: forecast_vals,
      })
      forecast_df = forecast_df.dropna(subset=["Pasta Çeşidi"])
      forecast_df = forecast_df[
          ~forecast_df["Pasta Çeşidi"]
          .astype(str)
          .str.contains("BATON", na=False)
      ]

      total_forecast = int(forecast_df[target_col_name].sum())

      st.metric(
          label="Önerilen Toplam Pasta Üretimi",
          value=f"{total_forecast} Adet",
      )
      st.dataframe(forecast_df, use_container_width=True, height=450)
    else:
      st.info(
          "👈 Sol panelden kriterleri seçip **'Yarın İçin Tahmini Üretim Listesi"
          " Oluştur'** butonuna tıklayın."
      )

with tab2:
  st.subheader("📝 Akşam Gelen Reel Sipariş Girişi")
  if df_all is not None:
    desired_order = [
        "BÜYÜK TURTA ÇİKOLATALI (ÇİKOLATA ÇİLEK OREO)",
        "BÜYÜK TURTA ÇİKOLATALI (ÇİKOLATA MUZLU)",
        "BÜYÜK TURTA ÇİKOLATALI (ÇİKOLATA MUZLU ÇİLEKLİ)",
        "BÜYÜK TURTA ÇİKOLATALI (ÇİKOLATA SOSLU)",
        "BÜYÜK TURTA FRAMBUAZ (FRAMBUAZ ÇİLEK)",
        "BÜYÜK TURTA MUZLU (MUZLU TURTA)",
        "ORTA TURTA ÇİKOLATALI (ÇİKOLATA ÇİLEK OREO)",
        "ORTA TURTA ÇİKOLATALI (ÇİKOLATA MUZLU)",
        "ORTA TURTA ÇİKOLATALI (ÇİKOLATA MUZLU ÇİLEKLİ)",
        "ORTA TURTA ÇİKOLATALI (ÇİKOLATA SOSLU)",
        "ORTA TURTA MEYVELİ",
        "ORTA TURTA FRAMBUAZ (FRAMBUAZ ÇİLEK)",
        "ORTA TURTA MUZLU (MUZLU TURTA)",
        "SIFIR BOY PASTA ÇİKOLATALI (ÇİKOLATA ÇİLEK OREO)",
        "SIFIR BOY PASTA ÇİKOLATALI (ÇİKOLATA MUZLU)",
        "SIFIR BOY PASTA ÇİKOLATALI (ÇİKOLATA MUZLU ÇİLEKLİ)",
        "SIFIR BOY PASTA ÇİKOLATALI (ÇİKOLATA SOSLU)",
        "SIFIR BOY PASTA FRAMBUAZ (FRAMBUAZ ÇİLEK)",
        "SIFIR BOY PASTA MUZLU (MUZLU TURTA)",
        "SIFIR BOY PASTA MEYVELİ",
        "BÜYÜK TURTA ÇİLEKLİ OREO PASTA",
        "ORTA TURTA ÇİLEKLİ OREO PASTA",
        "SIFIR TURTA ÇİLEKLİ OREO PASTA",
    ]
    excel_products = df_all.iloc[1:, 0].dropna().astype(str).tolist()
    products_list = [p for p in desired_order if p in excel_products]
    for p in excel_products:
      if p not in products_list:
        products_list.append(p)

    order_date = st.date_input("Sipariş Tarihi Seçin", key="order_date_input")
    st.markdown("---")
    st.markdown("🌤️ **O Günün Hava & Çevre Koşulları**")
    entry_weather = st.selectbox(
        "Hava Durumu",
        [
            "Açık",
            "Bulutlu",
            "Hafif Çisenti",
            "Hafif Yağmurlu",
            "Çok Yağmurlu",
            "Çok Bulutlu",
        ],
        key="entry_weather",
    )
    entry_temp = st.number_input(
        "Sıcaklık (°C)", min_value=-10, max_value=45, value=20, step=1, key="entry_temp"
    )

    st.markdown("---")
    with st.form("order_entry_form"):
      entered_orders = {}
      for idx, prod in enumerate(products_list):
        entered_orders[prod] = st.number_input(
            f"{prod}", min_value=0, max_value=500, value=0, step=1, key=f"ord_{idx}"
        )

      current_total = sum(entered_orders.values())
      st.info(f"📊 Girilen Toplam Sipariş Adeti: **{current_total} Adet**")
      submit_order = st.form_submit_button(
          "💾 Gerçekleşen Siparişleri ve Hava Durumunu Kaydet"
      )

      if submit_order:
        new_rows = []
        str_date = order_date.strftime("%Y-%m-%d")
        if not df_orders.empty:
          df_orders = df_orders[df_orders["Tarih"] != pd.to_datetime(str_date)]

        for prod, qty in entered_orders.items():
          if qty > 0:
            new_rows.append({
                "Tarih": str_date,
                "Pasta Çeşidi": prod,
                "Gerçekleşen Sipariş": qty,
            })

        if new_rows:
          new_df = pd.DataFrame(new_rows)
          new_df["Tarih"] = pd.to_datetime(new_df["Tarih"])
          df_orders = pd.concat([df_orders, new_df], ignore_index=True)
          df_orders_save = df_orders.copy()
          df_orders_save["Tarih"] = df_orders_save["Tarih"].dt.strftime(
              "%Y-%m-%d"
          )
          df_orders_save.to_csv(csv_archive_file, index=False)

          if not df_weather.empty:
            df_weather = df_weather[
                df_weather["Tarih"] != pd.to_datetime(str_date)
            ]
          new_weather_row = pd.DataFrame([{
              "Tarih": pd.to_datetime(str_date),
              "Hava Durumu": entry_weather,
              "Sıcaklık (°C)": f"{entry_temp}°C",
          }])
          df_weather = pd.concat([df_weather, new_weather_row], ignore_index=True)
          df_weather_save = df_weather.copy()
          df_weather_save["Tarih"] = df_weather_save["Tarih"].dt.strftime(
              "%Y-%m-%d"
          )
          df_weather_save.to_csv(csv_weather_file, index=False)
          st.success("✅ Kayıt başarıyla tamamlandı!")
          st.rerun()
        else:
          st.warning("⚠️ 0'dan büyük sipariş girilmedi.")

with tab3:
  st.subheader("📂 Arşiv ve Matris Tablo")
  if not df_orders.empty:
    st.markdown("---")
    st.subheader("✏️ Geçmiş Gün Verisini Düzenle")
    edit_date_list = sorted(
        df_orders["Tarih"].dt.strftime("%d.%m.%Y").unique().tolist(), reverse=True
    )
    if edit_date_list:
      selected_edit_date_str = st.selectbox(
          "Düzenlenecek Tarihi Seçin", edit_date_list, key="edit_date_sel"
      )
      dt_edit_obj = pd.to_datetime(selected_edit_date_str, format="%d.%m.%Y")

      current_w_val, current_t_val = "Açık", 20
      if not df_weather.empty:
        w_match = df_weather[df_weather["Tarih"] == dt_edit_obj]
        if not w_match.empty:
          current_w_val = w_match.iloc[0].get("Hava Durumu", "Açık")
          try:
            current_t_val = int(
                str(w_match.iloc[0].get("Sıcaklık (°C)", "20"))
                .replace("°C", "")
                .strip()
            )
          except:
            current_t_val = 20

      col_e1, col_