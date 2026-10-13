import datetime
import json
import os
import random
import google.generativeai as genai
from urllib.request import Request, urlopen
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tatlıbak AI Üretim Planlama Asistanı",
    page_icon="🍰",
    layout="wide",
)

st.title(
    "🍰 Tatlıbak Yaş Pasta Gerçek Gemini AI & Muhakeme Planlama Asistanı"
)

# --- RESMİ GOOGLE SDK İLE API BAĞLANTISI ---
api_key = None
try:
  api_key = st.secrets.get("GEMINI_API_KEY")
except:
  pass

if not api_key:
  api_key = os.environ.get("GEMINI_API_KEY", "")

if api_key:
  try:
    genai.configure(api_key=api_key)
  except:
    pass


def call_gemini_api(prompt_text):
  if not api_key:
    return "⚠️ Gemini API anahtarı bulunamadı! Lütfen secrets.toml dosyanızı kontrol edin."
  try:
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt_text)
    return response.text
  except Exception as e:
    return f"Gemini API Bağlantı Hatası: {e}"


excel_file = "YAŞ PASTA EYLÜL SİPARİŞ.xlsx"
csv_archive_file = "daily_orders.csv"
csv_weather_file = "daily_weather.csv"
historical_excel_file = "gecmis_arsiv.xlsx"


@st.cache_data
def load_data():
  if not os.path.exists(excel_file):
    return None
  try:
    xls = pd.ExcelFile(excel_file)
    sheet = xls.sheet_names[0] if xls.sheet_names else "Sheet1"
    for s in xls.sheet_names:
      if "SİPARİŞ" in s.upper() or "LİST" in s.upper():
        sheet = s
        break
    return xls.parse(sheet)
  except Exception as e:
    return None


df_all = load_data()
if df_all is None:
  st.warning(
      f"⚠️ `{excel_file}` dosyası bulunamadı veya okunamadı. Lütfen dosyanın"
      " depoda olduğundan emin olun."
  )


@st.cache_data
def load_historical_data():
  if os.path.exists(historical_excel_file):
    try:
      xls_h = pd.ExcelFile(historical_excel_file)
      sheet_name = xls_h.sheet_names[0]
      df_h = xls_h.parse(sheet_name)
      return df_h, sheet_name
    except:
      return None, "Hata"
  return None, "Bulunamadı"


df_history, hist_info = load_historical_data()

if os.path.exists(csv_archive_file):
  try:
    df_orders = pd.read_csv(csv_archive_file)
    df_orders["Tarih"] = pd.to_datetime(df_orders["Tarih"])
  except:
    df_orders = pd.DataFrame(
        columns=["Tarih", "Pasta Çeşidi", "Gerçekleşen Sipariş"]
    )
else:
  df_orders = pd.DataFrame(
      columns=["Tarih", "Pasta Çeşidi", "Gerçekleşen Sipariş"]
  )

if os.path.exists(csv_weather_file):
  try:
    df_weather = pd.read_csv(csv_weather_file)
    df_weather["Tarih"] = pd.to_datetime(df_weather["Tarih"])
  except:
    df_weather = pd.DataFrame(
        columns=["Tarih", "Hava Durumu", "Sıcaklık (°C)"]
    )
else:
  df_weather = pd.DataFrame(
      columns=["Tarih", "Hava Durumu", "Sıcaklık (°C)"]
  )


@st.cache_data(ttl=3600)
def fetch_3day_evening_weather():
  forecast_days = []
  try:
    url = "https://wttr.in/Istanbul?format=j1"
    response = urlopen(url, timeout=3)
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


three_day_forecast = fetch_3day_evening_weather()

tab1, tab2, tab3, tab4 = st.tabs([
    "🚀 Gemini AI Muhakeme Tahmini",
    "📝 Günlük Sipariş Girişi",
    "📂 Arşiv & Matris Tablo",
    "📜 Geçmiş Yıllar Arşivi",
])

with tab1:
  st.subheader(
      "🤖 Gerçek Gemini AI Zekası ile Muhakemeli Üretim Planlama Asistanı"
  )

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
            "Açık / Güneşli (Normal)",
            "Bulutlu / Hafif Yağmurlu",
            "Çok Yağmurlu / Fırtınalı",
            "Aşırı Sıcak",
        ],
    )
    temperature = st.slider("Beklenen Sıcaklık (°C)", -5, 40, 20, 1)
    special_day = st.selectbox(
        "Özel Durum / Maaş Dönemi",
        [
            "Normal Gün",
            "Maaş Dönemi / Ayın 10-15 Sarmalı",
            "Yılbaşı ve Arifesi",
            "Sevgililer Günü / Anneler Günü",
            "Resmi Tatil / Bayram",
            "Hafta Sonu Yoğunluğu",
        ],
    )
    buffer_pct = st.selectbox(
        "Güvenlik / Stok Payı (%)", [0, 5, 10, 15, 20, 25], index=2
    )
    ai_hesapla_button = st.button(
        "🧠 Gemini AI ile Muhakemeli Üretim Planı Oluştur"
    )

  with col2:
    if ai_hesapla_button:
      with st.spinner(
          "🧠 Gemini AI arşivi, geçmiş ekim döngülerini (8-13 Ekim) ve sipariş"
          " momentumunu muhakeme ediyor..."
      ):
        try:
          recent_orders_summary = (
              df_orders.tail(20).to_string()
              if not df_orders.empty
              else "Henüz günlük sipariş girilmemiş."
          )
          history_summary = (
              df_history.head(30).to_string()
              if df_history is not None
              else "Geçmiş arşiv yok."
          )
          base_excel_summary = (
              df_all.head(30).to_string()
              if df_all is not None
              else "Excel verisi yok."
          )

          prompt = f"""
                    Sen Tatlıbak pastacılık üretim planlama sisteminin kıdemli yapay zeka muhakeme uzmanısın. Asla ezbere ve sabit rakamlar verme. Seçilen günün özelliklerine, haftanın hangi gününe denk geldiğine, geçmişteki ekim ayı dalgalarına (örneğin 8 Ekim 1317, 9 Ekim 1452, 10 Ekim 2241, 11 Ekim 2339 zirvesi, 12 Ekim 2019, 13 Ekim 1263 düşüşü gibi geçmiş verilerdeki non-linear sarmallara) ve şubelerden gelen sipariş momentumuna dayanarak her defasında tamamen o güne özel özgün üretim planı kurgula ve geçmiş yıla göre uyarılar ver.
                    
                    HEDEF ANALİZ KRİTERLERİ:
                    - Hedef Tarih ve Gün: {selected_date.strftime('%d.%m.%Y')} ({turkish_day})
                    - Özel Durum / Dönem: {special_day}
                    - Hava Durumu & Sıcaklık: {weather_condition}, {temperature}°C
                    - Güvenlik / Stok Payı: %{buffer_pct}
                    
                    VERİ KAYNAKLARI:
                    1. Son Girilen Gerçek Siparişler:
                    {recent_orders_summary}
                    
                    2. Geçmiş Yıllar Arşivi:
                    {history_summary}
                    
                    3. Ürün Çeşit Yapısı:
                    {base_excel_summary}
                    
                    GÖREVLERİN:
                    1. Geçmiş arşivdeki benzer dönemlerin (örneğin Ekim ortası sarmalları) artış ve azalış eğrilerini inceleyerek bugünkü durumla kıyasla.
                    2. Yöneticiye şu formatta muhakemeli bir rapor sun:
                       - **Muhakeme & Arşiv Kıyaslama Uyarısı:** Geçmiş yılın aynı dönemiyle bugünkü trendi kıyaslayarak dikkat edilmesi gereken noktaları ve uyarıları açıkla.
                       - **Üretim Listesi Raporu:** Ürün bazlı nihai üretim adetleri ve en altta TOPLAM PASTA SAYISI.
                    """

          ai_text = call_gemini_api(prompt)
          st.markdown(ai_text)
        except Exception as ex:
          st.error(f"Gemini AI analizi sırasında hata oluştu: {ex}")
    else:
      st.info(
          "👈 Sol panelden kriterleri seçip **'Gemini AI ile Muhakemeli Üretim"
          " Planı Oluştur'** butonuna tıklayın."
      )

with tab2:
  st.subheader("📝 Akşam Gelen Reel Sipariş Girişi")
  if df_all is not None:
    try:
      excel_products = [
          p
          for p in df_all.iloc[1:, 0].dropna().astype(str).tolist()
          if "TOPLAM" not in p.strip().upper()
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
    except Exception as e:
      st.error(f"Sipariş formu yüklenirken hata: {e}")
  else:
    st.warning("⚠️ Excel verisi yüklenemediği için sipariş formu açılamıyor.")

with tab3:
  st.subheader("📂 Arşiv & Matris Tablo")
  if df_all is not None and not df_orders.empty:
    try:
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
      master_df = pd.concat(
          [pd.DataFrame([date_row]), master_df], ignore_index=True
      )
      st.dataframe(master_df, use_container_width=True, height=500)
    except Exception as ex:
      st.error(f"Matris tablo oluşturulurken hata: {ex}")
  else:
      st.info("Henüz günlük sipariş arşivi oluşmadı veya Excel yüklenemedi.")

with tab4:
  st.subheader("📜 Geçmiş Yıllar Arşivi (gecmis_arsiv.xlsx)")
  if df_history is not None:
    st.success("✅ Arşiv başarıyla okundu!")
    st.dataframe(df_history, use_container_width=True, height=550)
  else:
    st.warning("⚠️ `gecmis_arsiv.xlsx` dosyası bulunamadı veya okunamadı.")