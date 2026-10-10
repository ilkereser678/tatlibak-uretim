import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tatlıbak Üretim Planlama Asistanı", page_icon="🍰", layout="wide"
)

st.title("🍰 Tatlıbak Yaş Pasta Üretim & Sipariş Planlama Asistanı")

# Excel ve CSV dosya yolları
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
else:
  df_orders = pd.DataFrame(
      columns=["Tarih", "Pasta Çeşidi", "Gerçekleşen Sipariş"]
  )

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

    weather_condition = st.selectbox(
        "Yarının Hava Durumu / Koşul",
        [
            "Açık / Güneşli",
            "Az Bulutlu",
            "Parçalı Bulutlu",
            "Hafif Yağmurlu / Çisenti",
            "Yağmurlu",
        ],
    )

    total_actual_pasta = st.number_input(
        "Bugünkü Gerçekleşen Toplam Pasta", min_value=0, max_value=5000, value=2138
    )

    # Güvenlik / Stok Payı Yüzde Seçeneği
    buffer_pct = st.selectbox(
        "Stok / Güvenlik Payı Ekle (%)", [5, 10, 15, 20, 25], index=1
    )

    hesapla_button = st.button("🚀 Yarın İçin Tahmini Üretim Listesi Oluştur")

  with col2:
    if hesapla_button:
      st.success(
          f"📋 {selected_date.strftime('%d.%m.%Y')} Sonrası İçin Tahmin"
          f" Hazırlandı! (Stok Payı: %{buffer_pct})"
      )

      products = df_all.iloc[1:, 0].values
      numeric_df = df_all.iloc[1:, 1:].apply(pd.to_numeric, errors="coerce")

      if not numeric_df.empty:
        base_vals = numeric_df.mean(axis=1).fillna(0)
      else:
        base_vals = pd.Series([0] * len(products))

      multiplier = 1 + (buffer_pct / 100.0)
      forecast_vals = (base_vals * multiplier).round().astype(int)

      forecast_df = pd.DataFrame({
          "Pasta Çeşidi": products,
          "Baz Ortalama": base_vals.round().astype(int),
          f"Tahmini Üretim (%{buffer_pct} Stoklu)": forecast_vals,
      })

      forecast_df = forecast_df.dropna(subset=["Pasta Çeşidi"])
      forecast_df = forecast_df[
          ~forecast_df["Pasta Çeşidi"]
          .astype(str)
          .str.contains("BATON", na=False)
      ]

      total_forecast = forecast_df[
          f"Tahmini Üretim (%{buffer_pct} Stoklu)"
      ].sum()

      st.metric(
          label=f"Önerilen Toplam Pasta Üretimi (Stoklu: %{buffer_pct})",
          value=f"{total_forecast} Adet",
      )
      st.dataframe(forecast_df, use_container_width=True, height=450)
    else:
      st.info(
          "👈 Sol panelden tarih, stok yüzdesi ve verileri seçip **'Yarın İçin"
          " Tahmini Üretim Listesi Oluştur'** butonuna tıklayın."
      )

with tab2:
  st.subheader("📝 Akşam Gelen Reel Sipariş Girişi")
  st.write(
      "Bu ekrandan akşam gelen gerçek sipariş adetlerini ilgili tarihe"
      " işleyebilir ve doğrudan arşive kaydedebilirsiniz."
  )

  if df_all is not None:
    products_list = df_all.iloc[1:, 0].dropna().astype(str).tolist()

    order_date = st.date_input("Sipariş Tarihi Seçin", key="order_date_input")

    st.markdown("---")
    st.write("### Ürün Bazlı Sipariş Adetleri")

    with st.form("order_entry_form"):
      entered_orders = {}
      cols = st.columns(2)
      for idx, prod in enumerate(products_list):
        col_idx = idx % 2
        with cols[col_idx]:
          entered_orders[prod] = st.number_input(
              f"{prod}", min_value=0, max_value=500, value=0, step=1, key=f"ord_{idx}"
          )

      submit_order = st.form_submit_button(
          "💾 Gerçekleşen Siparişleri Kalıcı Olarak Kaydet"
      )

      if submit_order:
        new_rows = []
        str_date = order_date.strftime("%Y-%m-%d")

        # Aynı tarihe ait eski kayıtlar varsa temizleyelim (güncelleme mantığı için)
        if not df_orders.empty and "Tarih" in df_orders.columns:
          df_orders = df_orders[df_orders["Tarih"] != str_date]

        for prod, qty in entered_orders.items():
          if qty > 0:  # Sadece sıfırdan büyükleri kaydedelim
            new_rows.append({
                "Tarih": str_date,
                "Pasta Çeşidi": prod,
                "Gerçekleşen Sipariş": qty,
            })

        if new_rows:
          new_df = pd.DataFrame(new_rows)
          df_orders = pd.concat([df_orders, new_df], ignore_index=True)
          df_orders.to_csv(csv_archive_file, index=False)
          st.success(
              f"✅ {order_date.strftime('%d.%m.%Y')} tarihli siparişler arşive"
              " başarıyla kaydedildi!"
          )
        else:
          st.warning(
              "⚠️ Kaydedilecek 0'dan büyük sipariş adedi girilmediği için"
              " kayıt yapılmadı."
          )

with tab3:
  st.subheader("📂 Arşiv, Geçmiş Siparişler ve Kayıt Yönetimi")

  # Kayıt silme veya düzenleme alanı
  if not df_orders.empty:
    st.write("### 🗑️ Kayıt Düzenleme / Silme")
    unique_dates = df_orders["Tarih"].unique().tolist()
    selected_del_date = st.selectbox(
        "Silmek veya Güncellemek İstediğiniz Tarihi Seçin", unique_dates
    )

    if st.button("🗑️ Seçilen Tarihe Ait Kayıtları Sil"):
      df_orders = df_orders[df_orders["Tarih"] != selected_del_date]
      df_orders.to_csv(csv_archive_file, index=False)
      st.success(
          f"✅ {selected_del_date} tarihli kayıtlar arşivden silindi. Sayfayı"
          " yenileyebilirsiniz."
      )
      st.rerun()

    st.markdown("---")
    st.write("### 📊 Kaydedilen Tüm Sipariş Arşivi (`daily_orders.csv`)")
    st.dataframe(df_orders, use_container_width=True, height=400)
  else:
    st.info(
        "Henüz arşive kaydedilmiş günlük sipariş bulunmuyor. 'Günlük Sipariş"
        " Girişi' sekmesinden veri ekleyebilirsiniz."
    )

  st.markdown("---")
  st.subheader("📂 Orijinal Excel Dosyası Arşivi")
  if df_all is not None:
    st.dataframe(df_all, use_container_width=True, height=400)