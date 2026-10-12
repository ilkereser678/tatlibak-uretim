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

      # --- AKILLI YÖNETİCİ ÖZETİ & TREND ANALİZİ ---
      historical_insight_text = (
          "Geçmiş yıl arşivi (`gecmis_arsiv.xlsx`) okunamadı."
      )
      if df_history is not None:
        try:
          target_month = selected_date.month
          target_day = selected_date.day

          date_col = df_history.columns[0]
          df_history["Parsed_Date"] = pd.to_datetime(
              df_history[date_col], errors="coerce"
          )
          num_cols = df_history.select_dtypes(
              include=["number", "float", "int"]
          ).columns

          if not num_cols.empty:
            # Geçmiş yılın o günkü satışı
            matched_hist = df_history[
                (df_history["Parsed_Date"].dt.month == target_month)
                & (df_history["Parsed_Date"].dt.day == target_day)
            ]
            hist_val = (
                int(matched_hist[num_cols].sum().sum())
                if not matched_hist.empty
                else 0
            )

            # Geçmiş yılın genel ortalaması ile bu yılın (df_all) genel ortalamasını kıyaslayarak büyüme trendini hesaplayalım
            hist_avg_daily = df_history[num_cols].mean().mean()
            current_avg_daily = valid_numeric_df.mean().mean()

            trend_note = ""
            if hist_avg_daily > 0:
              growth_rate = (
                  (current_avg_daily - hist_avg_daily) / hist_avg_daily
              ) * 100
              if growth_rate > 0:
                trend_note = (
                    f" Ancak bu yıl genel satış hacminiz geçen yıla oranla"
                    f" **%{growth_rate:.1f} daha yüksek** seyretmektedir."
                )
              else:
                trend_note = (
                    f" Bu yıl genel satış hacminiz geçen yıla oranla"
                    f" **%{abs(growth_rate):.1f} daha düşük** seyretmektedir."
                )

            if hist_val > 0:
              historical_insight_text = (
                  f"Geçen yıl bugün ({target_day}.{target_month}.2025) toplam"
                  f" **{hist_val} adet** satış gerçekleşmişti.{trend_note} Bu"
                  f" veriler ışığında bugünkü optimize üretim hedefiniz"
                  f" **{subtotal} adet** olarak belirlenmiştir."
              )
            else:
              historical_insight_text = (
                  f"Geçmiş yıl arşivinde {target_day}.{target_month} için"
                  f" birebir kayıt bulunamadı, ancak genel büyüme trendi ({trend_note})"
                  f" analize yansıtıldı."
              )
          else:
            historical_insight_text = (
                "Geçmiş yıl verisinde uygun sayısal sütun bulunamadı."
            )
        except Exception as ex:
          historical_insight_text = (
              f"Geçmiş yıl verisi analiz edilirken hata oluştu: {ex}"
          )

      st.markdown("---")
      st.info(
          "🤖 **Tatlıbak Akıllı Yönetici Özeti & Büyüme Trendi Analizi**\n\n"
          f"- **Seçilen Kriterler:** {days_tr.get(day_of_week, '')} günü için"
          f" **{special_day}** ve hava koşulu **'{weather_condition}'**"
          f" ({temperature}°C) baz alınmıştır.\n- **Geçmiş Yıl & Trend"
          f" Kıyaslaması:** {historical_insight_text}\n- **Operasyonel Öneri:**"
          f" Şube sipariş ortalamaları, yıllık büyüme trendi ve güvenlik payı"
          f" (%{buffer_pct}) harmanlanarak toplam **{subtotal} adet** üretim"
          " optimize edilmiştir."
      )

    else:
      st.info(
          "👈 Sol panelden 3 günlük hava tablosuna bakarak kriterleri seçip"
          " **'Yarın İçin Tahmini Üretim Listesi Oluştur'** butonuna tıklayın."
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
    excel_products = [
        p
        for p in df_all.iloc[1:, 0].dropna().astype(str).tolist()
        if "TOPLAM" not in p.strip().upper()
    ]
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
  st.subheader("📂 Arşiv & Matris Tablo")
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

      col_e1, col_e2 = st.columns(2)
      with col_e1:
        edit_weather = st.selectbox(
            "Hava Durumu Güncelle",
            [
                "Açık",
                "Bulutlu",
                "Hafif Çisenti",
                "Hafif Yağmurlu",
                "Çok Yağmurlu",
                "Çok Bulutlu",
            ],
            index=(
                [
                    "Açık",
                    "Bulutlu",
                    "Hafif Çisenti",
                    "Hafif Yağmurlu",
                    "Çok Yağmurlu",
                    "Çok Bulutlu",
                ].index(current_w_val)
                if current_w_val
                in [
                    "Açık",
                    "Bulutlu",
                    "Hafif Çisenti",
                    "Hafif Yağmurlu",
                    "Çok Yağmurlu",
                    "Çok Bulutlu",
                ]
                else 0
            ),
            key="edit_w_sel",
        )
      with col_e2:
        edit_temp = st.number_input(
            "Sıcaklık (°C) Güncelle",
            min_value=-10,
            max_value=45,
            value=current_t_val,
            step=1,
            key="edit_t_sel",
        )

      day_records = df_orders[df_orders["Tarih"] == dt_edit_obj]
      with st.form("edit_record_form"):
        updated_qtys = {}
        for _, row in day_records.iterrows():
          updated_qtys[row["Pasta Çeşidi"]] = st.number_input(
              f"{row['Pasta Çeşidi']}",
              min_value=0,
              max_value=500,
              value=int(row["Gerçekleşen Sipariş"]),
              step=1,
          )
        if st.form_submit_button("🔄 Güncelle"):
          df_orders = df_orders[df_orders["Tarih"] != dt_edit_obj]
          new_edit_rows = [
              {
                  "Tarih": dt_edit_obj,
                  "Pasta Çeşidi": p,
                  "Gerçekleşen Sipariş": q,
              }
              for p, q in updated_qtys.items()
              if q > 0
          ]
          if new_edit_rows:
            df_orders = pd.concat(
                [df_orders, pd.DataFrame(new_edit_rows)], ignore_index=True
            )
            df_orders.to_csv(csv_archive_file, index=False)

          if os.path.exists(csv_weather_file):
            df_w_up = pd.read_csv(csv_weather_file)
            df_w_up["Tarih"] = pd.to_datetime(df_w_up["Tarih"])
            df_w_up = df_w_up[df_w_up["Tarih"] != dt_edit_obj]
          else:
            df_w_up = pd.DataFrame(
                columns=["Tarih", "Hava Durumu", "Sıcaklık (°C)"]
            )
          new_w_row = pd.DataFrame([{
              "Tarih": dt_edit_obj,
              "Hava Durumu": edit_weather,
              "Sıcaklık (°C)": f"{edit_temp}°C",
          }])
          updated_weather_df = pd.concat(
              [df_w_up, new_w_row], ignore_index=True
          )
          updated_weather_df.to_csv(csv_weather_file, index=False)
          st.success("✅ Güncellendi!")
          st.rerun()

  # --- BİRLEŞTİRİLMİŞ MATRİS TABLO ---
  st.markdown("---")
  st.subheader("📂 Birleştirilmiş Ana Arşiv Tablosu")

  if df_all is not None:
    master_df = df_all.loc[:, ~df_all.columns.str.contains("^Unnamed")]
    product_col_name = master_df.columns[0]
    master_df = master_df[
        ~master_df[product_col_name]
        .astype(str)
        .str.upper()
        .str.contains("TOPLAM", na=False)
    ]

    weather_dict = {}
    if not df_weather.empty:
      for _, w_row in df_weather.iterrows():
        d_str = pd.to_datetime(w_row["Tarih"]).strftime("%Y-%m-%d")
        w_cond = w_row.get("Hava Durumu", "Açık")
        t_val = str(w_row.get("Sıcaklık (°C)", "20")).replace("°C", "").strip()
        weather_dict[d_str] = f"{t_val}°C - {w_cond}"

    if not df_orders.empty:
      df_orders["Clean_Product"] = (
          df_orders["Pasta Çeşidi"].astype(str).str.strip().str.upper()
      )
      pivot_orders = df_orders.pivot_table(
          index="Clean_Product",
          columns=df_orders["Tarih"].dt.strftime("%Y-%m-%d"),
          values="Gerçekleşen Sipariş",
          aggfunc="sum",
      ).fillna(0)

      master_df["Clean_Product"] = (
          master_df[product_col_name].astype(str).str.strip().str.upper()
      )

      for date_col in pivot_orders.columns:
        if date_col not in master_df.columns:
          master_df[date_col] = 0
        mapping_dict = pivot_orders[date_col].to_dict()
        master_df[date_col] = (
            master_df["Clean_Product"].map(mapping_dict).fillna(0)
        )

      master_df = master_df.drop(columns=["Clean_Product"])

    rename_cols = {}
    date_row_data = {product_col_name: "TARİH"}

    for col in master_df.columns:
      if col != product_col_name:
        col_str = str(col).split()[0]
        try:
          formatted_date_str = pd.to_datetime(col_str).strftime("%d.%m.%y")
        except:
          formatted_date_str = col_str

        if col_str in weather_dict:
          new_header = weather_dict[col_str]
          rename_cols[col] = new_header
          date_row_data[new_header] = formatted_date_str
        else:
          rename_cols[col] = str(col)
          date_row_data[str(col)] = formatted_date_str

    master_df = master_df.rename(columns=rename_cols)
    master_df = pd.concat(
        [pd.DataFrame([date_row_data]), master_df], ignore_index=True
    )

    total_row_data = {product_col_name: "TOPLAM PASTA SAYISI"}
    for col in master_df.columns:
      if col != product_col_name:
        numeric_col = pd.to_numeric(master_df[col].iloc[1:], errors="coerce")
        if not numeric_col.isna().all():
          total_row_data[col] = int(numeric_col.sum())
        else:
          total_row_data[col] = 0

    master_df = pd.concat(
        [master_df, pd.DataFrame([total_row_data])], ignore_index=True
    )

    for col in master_df.columns:
      if col != product_col_name:
        try:
          temp_num = pd.to_numeric(master_df[col], errors="coerce")
          if not temp_num.isna().all():
            mask = master_df.index != 0
            master_df.loc[mask, col] = (
                temp_num[mask].fillna(0).astype(int).astype(str)
            )
            master_df.loc[master_df.index == 0, col] = date_row_data.get(
                col, master_df.loc[0, col]
            )
        except:
          pass

    st.dataframe(master_df, use_container_width=True, height=500)

with tab4:
  st.subheader("📜 Geçmiş Yıllar Arşivi (gecmis_arsiv.xlsx)")
  if df_history is not None:
    st.success("✅ Arşiv başarıyla yüklendi ve okundu!")
    st.dataframe(df_history, use_container_width=True, height=550)
  else:
    st.warning(
        "⚠️ `gecmis_arsiv.xlsx` dosyası bulunamadı veya okunamadı. Lütfen"
        " dosyanın ana klasörde olduğundan emin olun."
    )
    if hist_info:
      st.info(f"Detay: {hist_info}")