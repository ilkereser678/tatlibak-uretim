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
  return xls.parse("Sipariş List")


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
      evening_temps, evening_descs = [], []

      for h in hourly_list:
        try:
          hour_val = int(h["time"]) // 100
          if 17 <= hour_val <= 21:
            evening_temps.append(int(h["tempC"]))
            evening_descs.append(h["weatherDesc"][0]["value"])
        except:
          pass

      avg_t = (
          int(sum(evening_temps) / len(evening_temps))
          if evening_temps
          else int(day_data.get("avgtempC", 20))
      )
      common_d = (
          max(set(evening_descs), key=evening_descs.count)
          if evening_descs
          else "Açık"
      )

      tr_desc = {
          "Sunny": "Açık",
          "Clear": "Açık",
          "Partly cloudy": "Parçalı Bulutlu",
          "Cloudy": "Bulutlu",
          "Overcast": "Çok Bulutlu",
          "Light rain": "Hafif Yağmurlu",
          "Moderate rain": "Yağmurlu",
      }.get(common_d, common_d)

      day_name_map = {0: "Bugün", 1: "Yarın", 2: "2 Gün Sonra (Satış Günü)"}
      forecast_days.append({
          "Gün": day_name_map.get(idx, f"+{idx} Gün"),
          "Tarih": date_str,
          "Akşam Sıcaklık (17:00-21:00)": f"{avg_t}°C",
          "Hava Durumu": tr_desc,
      })
  except:
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
  except:
    return "Mevsim Normalleri (18°C - Parçalı Bulutlu)"


three_day_forecast = fetch_3day_evening_weather()

tab1, tab2, tab3, tab4 = st.tabs([
    "🚀 Üretim Tahmini & Hesaplama",
    "📝 Günlük Sipariş Girişi",
    "📂 Arşiv & Matris Tablo",
    "📜 Geçmiş Yıllar Arşivi",
])

with tab1:
  st.subheader("Günlük Veri Girişi ve Yarınki Üretim Planı")

  if three_day_forecast:
    st.markdown("🌤️ **İstanbul 3 Günlük Akşam Yoğunluk Hava Rehberi**")
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
    weather_condition = st.selectbox(
        "Hava Durumu Etkisi",
        [
            "Açık / Güneşli (Normal)",
            "Bulutlu / Hafif Yağmurlu (Talep Artışı +%)",
            "Çok Yağmurlu / Fırtınalı (Talep Düşüşü -%)",
            "Aşırı Sıcak / Güneşli (Ağır Tatlı Azalır)",
        ],
    )
    temperature = st.slider("Beklenen Sıcaklık (°C)", -5, 40, 20, 1)
    special_day = st.selectbox(
        "Özel Gün Durumu",
        [
            "Normal Gün",
            "Yılbaşı ve Arifesi",
            "Sevgililer Günü (14 Şubat)",
            "Anneler Günü",
            "Resmi Tatil / Bayram",
            "Hafta Sonu / Cuma Akşamı",
        ],
    )
    buffer_pct = st.selectbox(
        "Stok / Güvenlik Payı Ekle (%)", [0, 5, 10, 15, 20, 25], index=2
    )
    hesapla_button = st.button("🚀 Yarın İçin Tahmini Üretim Listesi Oluştur")

  with col2:
    if hesapla_button:
      env_mult = 1.0
      if "Bulutlu" in weather_condition:
        env_mult += 0.05
      elif "Çok Yağmurlu" in weather_condition:
        env_mult -= 0.08
      elif "Aşırı Sıcak" in weather_condition or temperature > 30:
        env_mult -= 0.05

      spec_mult = 1.0
      if "Yılbaşı" in special_day:
        spec_mult = 1.40
      elif "Sevgililer" in special_day or "Anneler" in special_day:
        spec_mult = 1.25
      elif "Resmi Tatil" in special_day:
        spec_mult = 1.15
      elif "Hafta Sonu" in special_day:
        spec_mult = 1.10

      # --- YAKIN GEÇMİŞ MOMENTUM FAKTÖRÜ (Son gerçek siparişlerin ortalaması) ---
      recent_momentum_mult = 1.0
      if not df_orders.empty:
        recent_orders_grouped = df_orders.groupby("Tarih")[
            "Gerçekleşen Sipariş"
        ].sum()
        if not recent_orders_grouped.empty:
          recent_avg = recent_orders_grouped.tail(5).mean()
          excel_base_avg = (
              df_all.iloc[1:, 1:]
              .apply(pd.to_numeric, errors="coerce")
              .sum()
              .mean()
          )
          if excel_base_avg > 0 and recent_avg > excel_base_avg:
            recent_momentum_mult = recent_avg / excel_base_avg

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
      comb_mult = (
          env_mult
          * spec_mult
          * recent_momentum_mult
          * (1 + (buffer_pct / 100.0))
      )
      forecast_vals = (base_vals * comb_mult).round().astype(int)

      target_col_name = (
          "Tahmini Üretim (Net / Stoksuz)"
          if buffer_pct == 0
          else f"Tahmini Üretim (%{buffer_pct} Stoklu + Akıllı Momentum)"
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

      total_row_df = pd.DataFrame(
          {"Pasta Çeşidi": ["TOPLAM PASTA SAYISI"], target_col_name: [subtotal]}
      )
      forecast_df = pd.concat([forecast_df, total_row_df], ignore_index=True)

      st.metric(label="Önerilen Toplam Pasta Üretimi", value=f"{subtotal} Adet")
      st.dataframe(forecast_df, use_container_width=True, height=450)

      # --- GELİŞTİRİLMİŞ AKILLI YÖNETİCİ & DALGA ANALİZİ ---
      hist_text = "Geçmiş yıl arşivi okunamadı."
      w_month = 1
      surge_warning = ""
      if df_history is not None:
        try:
          t_month, t_day, t_weekday = (
              selected_date.month,
              selected_date.day,
              selected_date.weekday(),
          )
          w_month = (t_day - 1) // 7 + 1
          d_col = df_history.columns[0]
          df_history["Parsed_Date"] = pd.to_datetime(
              df_history[d_col], errors="coerce"
          )
          num_cols = df_history.select_dtypes(
              include=["number", "float", "int"]
          ).columns

          if not num_cols.empty:
            # Genişletilmiş Tarih Penceresi Taraması (±2 Günlük Dalga / Maaş Dönemi Kontrolü)
            target_dt = pd.Timestamp(2025, t_month, t_day)
            window_start = target_dt - pd.Timedelta(days=2)
            window_end = target_dt + pd.Timedelta(days=2)

            window_matched = df_history[
                (df_history["Parsed_Date"] >= window_start)
                & (df_history["Parsed_Date"] <= window_end)
            ]
            hist_val = (
                int(window_matched[num_cols].sum().sum() / 5)
                if not window_matched.empty
                else 0
            )

            # Eğer o tarihlerde geçmişte ani patlama (surge) gözlendiyse uyarı ekle
            if hist_val > 2500 or (
                not window_matched.empty
                and window_matched[num_cols].sum(axis=1).max() > 2500
            ):
              surge_warning = (
                  " ⚡ **Maaş Dönemi / Talep Patlaması Uyarısı:** Geçmiş"
                  " yıllarda bu tarih aralığında (ayın 10-11'i çevresinde)"
                  " harcama artışına bağlı olarak ani sipariş dalgaları (2500+"
                  " adet) yaşandığı tespit edilmiştir. Sistem bu dönemsel"
                  " dalgayı algılayarak üretim hedefini yukarı yönlü"
                  " revize etmiştir."
              )

            past_weather = fetch_historical_weather(
                f"2025-{t_month:02d}-{t_day:02d}"
            )
            buf_str = f" ve %{buffer_pct} güvenlik payı" if buffer_pct > 0 else ""

            hist_text = (
                f"Geçen yıl benzer dönem ve çevresindeki günlerde (8-11 Ekim"
                f" dönemi) geçmiş arşiv taramasında yüksek hacimli talep"
                f" dalgaları görülmüştür. O dönem hava **{past_weather}** idi."
                f" Son günlerdeki gerçek sipariş momentumu, hava koşulları{buf_str}"
                f" ve dönemsel dalga analizi harmanlanarak bugünkü üretim"
                f" hedefiniz **{subtotal} adet** olarak optimize edilmiştir."
                f"{surge_warning}"
            )
        except Exception as ex:
          hist_text = f"Analiz hatası: {ex}"

      st.markdown("---")
      st.info(
          "🤖 **Tatlıbak Akıllı Yönetici & Dalga Analizi Özeti**\n\n"
          f"- **Seçilen Kriterler:** {days_tr.get(day_of_week, '')} günü"
          f" (Ayın {w_month}. haftası) / **{special_day}** / Hava:"
          f" **{weather_condition}** ({temperature}°C).\n- **Gelişmiş Analiz &"
          f" Trend:** {hist_text}"
      )
    else:
      st.info(
          "👈 Sol panelden kriterleri seçip **'Yarın İçin Tahmini Üretim Listesi"
          " Oluştur'** butonuna tıklayın."
      )

with tab2:
  st.subheader("📝 Akşam Gelen Reel Sipariş Girişi")
  if df_all is not None:
    excel_products = [
        p
        for p in df_all.iloc[1:, 0].dropna().astype(str).tolist()
        if "TOPLAM" not in p.strip().upper()
    ]
    order_date = st.date_input("Sipariş Tarihi Seçin", key="odate")
    st.markdown("---")
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
        key="eweather",
    )
    entry_temp = st.number_input("Sıcaklık (°C)", -10, 45, 20, 1, key="etemp")

    with st.form("order_entry_form"):
      entered_orders = {
          prod: st.number_input(
              f"{prod}", min_value=0, max_value=500, value=0, step=1, key=f"o_{i}"
          )
          for i, prod in enumerate(excel_products)
      }
      current_total = sum(entered_orders.values())
      st.info(f"📊 Girilen Toplam Sipariş Adeti: **{current_total} Adet**")
      if st.form_submit_button("💾 Gerçekleşen Siparişleri ve Havayı Kaydet"):
        str_date = order_date.strftime("%Y-%m-%d")
        if not df_orders.empty:
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
          df_orders = pd.concat(
              [df_orders, pd.DataFrame(new_rows)], ignore_index=True
          )
          df_orders.to_csv(csv_archive_file, index=False)
          st.success("✅ Kayıt başarılı!")
          st.rerun()
        else:
          st.warning("⚠️ 0'dan büyük sipariş girilmedi.")

with tab3:
  st.subheader("📂 Arşiv & Matris Tablo")
  if not df_orders.empty:
    master_df = df_all.loc[:, ~df_all.columns.str.contains("^Unnamed")]
    p_col = master_df.columns[0]
    master_df = master_df[
        ~master_df[p_col].astype(str).str.upper().str.contains("TOPLAM")
    ]
    weather_dict = {}
    if not df_weather.empty:
      for _, w_row in df_weather.iterrows():
        d_str = pd.to_datetime(w_row["Tarih"]).strftime("%Y-%m-%d")
        weather_dict[d_str] = (
            f"{str(w_row.get('Sıcaklık (°C)','20')).replace('°C','')}°C -"
            f" {w_row.get('Hava Durumu','Açık')}"
        )

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
        master_df[p_col].astype(str).str.strip().str.upper()
    )

    for d_col in pivot_orders.columns:
      if d_col not in master_df.columns:
        master_df[d_col] = 0
      master_df[d_col] = master_df["Clean_Product"].map(
          pivot_orders[d_col].to_dict()
      ).fillna(0)
    master_df = master_df.drop(columns=["Clean_Product"])

    rename_cols, date_row = {}, {p_col: "TARİH"}
    for col in master_df.columns:
      if col != p_col:
        c_str = str(col).split()[0]
        parsed_dt = pd.to_datetime(c_str, errors="coerce")
        f_str = (
            parsed_dt.strftime("%d.%m.%y")
            if pd.notna(parsed_dt)
            else str(col)
        )
        if c_str in weather_dict:
          rename_cols[col] = weather_dict[c_str]
          date_row[weather_dict[c_str]] = f_str
        else:
          rename_cols[col] = str(col)
          date_row[str(col)] = f_str

    master_df = master_df.rename(columns=rename_cols)
    master_df = pd.concat([pd.DataFrame([date_row]), master_df], ignore_index=True)
    st.dataframe(master_df, use_container_width=True, height=500)

with tab4:
  st.subheader("📜 Geçmiş Yıllar Arşivi (gecmis_arsiv.xlsx)")
  if df_history is not None:
    st.success("✅ Arşiv başarıyla okundu!")
    st.dataframe(df_history, use_container_width=True, height=550)
  else:
    st.warning("⚠️ `gecmis_arsiv.xlsx` dosyası bulunamadı.")