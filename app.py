import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tatlıbak Üretim Planlama Asistanı", page_icon="🍰", layout="wide"
)

st.title("🍰 Tatlıbak Yaş Pasta Üretim & Sipariş Planlama Asistanı")

# Excel و CSV dosya yolları
excel_file = "YAŞ PASTA EYLÜL SİPARİŞ.xlsx"
csv_archive_file = "daily_orders.csv"


@st.cache_data
def load_data():
  xls = pd.ExcelFile(excel_file)
  df = xls.parse("Sipariş List")
  return df


try:
  df_all = load_data()
except Exception as e:
  st.error(f"Excel dosyası yüklenirken hata oluştu: {e}")

# CSV Arşiv dosyasını okuma veya oluşturma
if os.path.exists(csv_archive_file):
  df_orders = pd.read_csv(csv_archive_file)
  df_orders["Tarih"] = pd.to_datetime(df_orders["Tarih"])
else:
  df_orders = pd.DataFrame(
      columns=["Tarih", "Pasta Çeşidi", "Gerçekleşen Sipariş"]
  )
  df_orders["Tarih"] = pd.to_datetime(df_orders["Tarih"])

# Sekme Yapısı (Tabs)
tab1, tab2, tab3 = st.tabs([
    "🚀 Üretim Tahmini & Hesaplama",
    "📝 Günlük Sipariş Girişi",
    "📂 Arşiv & Geçmiş Siparişler",
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
    tr_day = days_tr.get(day_of_week, "")
    st.info(f"Seçilen Gün: **{tr_day}**")

    # Hava Durumu & Sıcaklık Parametreleri
    st.markdown("---")
    st.markdown("🌤️ **Hava & Çevre Koşulları**")
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
        "Beklenen Sıcaklık (°C)", min_value=-5, max_value=40, value=20, step=1
    )

    # Özel Günler ve Resmi Tatiller
    st.markdown("🎉 **Özel Gün & Tatil Seçenekleri**")
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

    total_actual_pasta = st.number_input(
        "Bugünkü Gerçekleşen Toplam Pasta", min_value=0, max_value=5000, value=2138
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

      st.success(
          f"📋 {selected_date.strftime('%d.%m.%Y')} İçin Tahmin Hazırlandı! |"
          f" Hava: {weather_condition} ({temperature}°C) | Özel Gün: {special_day}"
      )

      products = df_all.iloc[1:, 0].values
      numeric_df = df_all.iloc[1:, 1:].apply(pd.to_numeric, errors="coerce")

      if not numeric_df.empty:
        base_vals = numeric_df.mean(axis=1).fillna(0)
      else:
        base_vals = pd.Series([0] * len(products))

      combined_multiplier = (
          environmental_multiplier
          * special_multiplier
          * (1 + (buffer_pct / 100.0))
      )
      forecast_vals = (base_vals * combined_multiplier).round().astype(int)

      if buffer_pct == 0:
        target_col_name = "Tahmini Üretim (Net / Stoksuz)"
      else:
        target_col_name = f"Tahmini Üretim (%{buffer_pct} Stoklu + Akıllı Etki)"

      forecast_df = pd.DataFrame({
          "Pasta Çeşidi": products,
          "Baz Ortalama": base_vals.round().astype(int),
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
          label=(
              f"Önerilen Toplam Pasta Üretimi (Stok: %{buffer_pct} + Hava/Özel"
              " Gün Etkili)"
          ),
          value=f"{total_forecast} Adet",
      )
      st.dataframe(forecast_df, use_container_width=True, height=450)
    else:
      st.info(
          "👈 Sol panelden tarih, hava durumu, özel gün ve stok yüzdesini seçip"
          " **'Yarın İçin Tahmini Üretim Listesi Oluştur'** butonuna tıklayın."
      )

with tab2:
  st.subheader("📝 Akşam Gelen Reel Sipariş Girişi")
  st.write(
      "Bu ekrandan akşam gelen gerçek sipariş adetlerini belirlediğiniz"
      " sıralamaya göre ilgili tarihe işleyebilir ve arşive"
      " kaydedebilirsiniz. (TAB tuşuyla hızlıca ilerleyebilirsiniz)"
  )

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
    st.write("### Ürün Bazlı Sipariş Adetleri (Tek Sütun / Hızlı Giriş)")

    with st.form("order_entry_form"):
      entered_orders = {}
      for idx, prod in enumerate(products_list):
        entered_orders[prod] = st.number_input(
            f"{prod}", min_value=0, max_value=500, value=0, step=1, key=f"ord_{idx}"
        )

      current_total = sum(entered_orders.values())
      st.info(f"📊 Girilen Toplam Sipariş Adeti: **{current_total} Adet**")

      submit_order = st.form_submit_button(
          "💾 Gerçekleşen Siparişleri Kalıcı Olarak Kaydet"
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
          st.success(
              f"✅ {order_date.strftime('%d.%m.%Y')} tarihli toplam"
              f" **{current_total}** adet sipariş arşive başarıyla"
              " kaydedildi!"
          )
        else:
          st.warning("⚠️ 0'dan büyük sipariş girilmedi.")

with tab3:
  st.subheader("📂 Arşiv, Geçmiş Siparişler ve Birleştirilmiş Matris Tablo")

  if not df_orders.empty:
    st.write("### 🔍 Arşivde Filtreleme ve Arama")
    f_col1, f_col2 = st.columns(2)

    with f_col1:
      min_d = df_orders["Tarih"].min().date()
      max_d = df_orders["Tarih"].max().date()
      date_range = st.date_input(
          "Görüntülenecek Tarih Aralığı", value=(min_d, max_d)
      )

    with f_col2:
      search_product = st.text_input(
          "Ürün Adına Göre Ara (Örn: Baton, Çikolatalı)"
      )

    filtered_df = df_orders.copy()
    if isinstance(date_range, tuple) and len(date_range) == 2:
      start_date, end_date = date_range
      filtered_df = filtered_df[
          (filtered_df["Tarih"].dt.date >= start_date)
          & (filtered_df["Tarih"].dt.date <= end_date)
      ]

    if search_product:
      filtered_df = filtered_df[
          filtered_df["Pasta Çeşidi"]
          .str.contains(search_product, case=False, na=False)
      ]

    display_df = filtered_df.copy()
    display_df["Tarih"] = display_df["Tarih"].dt.strftime("%d.%m.%Y")

    st.markdown("---")
    st.write("### 📊 Filtrelenmiş Günlük Sipariş Arşivi")
    st.dataframe(display_df, use_container_width=True, height=350)

    st.markdown("---")
    st.write("### 🗑️ Kayıt Düzenleme / Silme")
    unique_dates_str = display_df["Tarih"].unique().tolist()
    if unique_dates_str:
      selected_del_date_str = st.selectbox(
          "Silmek İstediğiniz Tarihi Seçin (GG.AA.YYYY)", unique_dates_str
      )

      if st.button("🗑️ Seçilen Tarihe Ait Tüm Kayıtları Sil"):
        dt_obj = pd.to_datetime(selected_del_date_str, format="%d.%m.%Y")
        df_orders = df_orders[df_orders["Tarih"] != dt_obj]

        df_orders_save = df_orders.copy()
        df_orders_save["Tarih"] = df_orders_save["Tarih"].dt.strftime(
              "%Y-%m-%d"
          )
        df_orders_save.to_csv(csv_archive_file, index=False)
        st.success(
            f"✅ {selected_del_date_str} tarihli kayıtlar arşivden silindi!"
        )
        st.rerun()

  # --- DÜZELTİLMİŞ: Excel'in Kendi 'TOPLAM PASTA SAYISI' Satırını Dinamik Güncelleyen Matris Tablo ---
  st.markdown("---")
  st.subheader("📂 Birleştirilmiş Ana Arşiv Tablosu (Excel + Yeni Günler)")

  if df_all is not None:
    master_df = df_all.copy()
    product_col_name = master_df.columns[0]

    # 1. Yeni girilen CSV siparişlerini matrise ekleyelim
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
          mapping_dict = pivot_orders[date_col].to_dict()
          master_df[date_col] = (
              master_df["Clean_Product"].map(mapping_dict).fillna(0)
          )

      master_df = master_df.drop(columns=["Clean_Product"])

    # 2. Orijinal Excel'deki "TOPLAM PASTA SAYISI" satırını bulup tüm sayısal sütunlar için güncel toplamı yazdıralım
    total_row_mask = (
        master_df[product_col_name]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.contains("TOPLAM PASTA SAYISI", na=False)
    )

    if total_row_mask.any():
      # Ürün çeşitleri satırlarını alalım (Toplam satırı hariç)
      products_only_df = master_df[~total_row_mask]

      # Tüm sayısal sütunlar için (hem eski Excel sütunları hem yeni eklenenler) toplamları hesaplayalım
      for col in master_df.columns:
        if col != product_col_name:
          numeric_col = pd.to_numeric(products_only_df[col], errors="coerce")
          if not numeric_col.isna().all():
            master_df.loc[total_row_mask, col] = numeric_col.sum()

    st.dataframe(master_df, use_container_width=True, height=480)