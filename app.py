import datetime
import json
import os
import random
import google.generativeai as genai  # Resmi Google SDK
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

# Resmi SDK'yı yapılandır (AQ... formatındaki yeni nesil anahtarları doğrudan destekler)
if api_key:
  genai.configure(api_key=api_key)


def call_gemini_api(prompt_text):
  if not api_key:
    return (
        "⚠️ Gemini API anahtarı bulunamadı! Lütfen .streamlit/secrets.toml"
        " dosyanıza AQ... ile başlayan anahtarınızı ekleyin."
    )

  try:
    # Gemini 1.5 Flash modelini resmi SDK ile çağırıyoruz
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt_text)
    return response.text
  except Exception as e:
    return (
        f"Gemini API Bağlantı Hatası: {e}. Lütfen anahtar yetkilerini kontrol"
        " edin."
    )


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

# (Arayüz sekme ve form kodları aynen devam eder...)