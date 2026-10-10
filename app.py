import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tatlıbak Üretim Planlama Asistanı", page_icon="🍰", layout="wide"
)

st.title("🍰 Tatlıbak Yaş Pasta Üretim & Sipariş Planlama Asistanı")

# Excel dosyasını yükleme
excel_file = "YAŞ PASTA EYLÜL SİPARİŞ.xlsx"


@st.cache_data
def load_data():
  xls = pd.ExcelFile(excel_file)
  df = xls.parse("Sipariş List")
  return df


try:
  df_all = load_data()
except Exception as e:
  st.error(f"Excel dosyası yüklenirken hata oluştu: {e}")

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

      # Pasta isimleri 0. sütunda, 1. satırdan başlar
      products = df_all.iloc[1:, 0].values

      # Verileri sayısal formata dönüştürüyoruz (metin kaynaklı 0 sorununu çözer)
      numeric_df = df_all.iloc[1:, 1:].apply(pd.to_numeric, errors="coerce")

      if not numeric_df.empty:
        base_vals = numeric_df.mean(axis=1).fillna(0)
      else:
        base_vals = pd.Series([0] * len(products))

      # Seçilen yüzdeye göre stok payı ekleyip yuvarlıyoruz
      multiplier = 1 + (buffer_pct / 100.0)
      forecast_vals = (base_vals * multiplier).round().astype(int)

      forecast_df = pd.DataFrame({
          "Pasta Çeşidi": products,
          "Baz Ortalama": base_vals.round().astype(int),
          f"Tahmini Üretim (%{buffer_pct} Stoklu)": forecast_vals,
      })

      # Boş satırları ve BATON ürünlerini filtreleyelim
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
      "Bu ekrandan akşam gelen gerçek siparişleri girebilir ve doğrudan"
      " sisteme kaydedebilirsiniz."
  )
  st.info("Sipariş giriş ekranı yapım aşamasında...")

with tab3:
  st.subheader("Geçmiş Arşiv ve Excel Tabloları")
  if df_all is not None:
    st.dataframe(df_all, use_container_width=True, height=500)