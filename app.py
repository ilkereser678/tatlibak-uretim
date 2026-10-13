İmport datetime as dt
import io
import json
import os
from urllib.parse import urlencode
from urllib.request import urlopen

import numpy as np
import pandas as pd
import streamlit as st
from sklearn.linear_model import Ridge

try:
from zoneinfo import ZoneInfo

BUGUN = dt.datetime.now(ZoneInfo("Europe/Istanbul")).date()
except Exception:
BUGUN = dt.date.today()

st.set_page_config(page_title="Tatlıbak Üretim Planlama", page_icon="🍰", layout="wide")

# ----------------------------------------------------------------------------
# AYARLAR
# ----------------------------------------------------------------------------
GUNLUK_DOSYALAR = ["gunluk_satis.xlsx", "gunluk_satis.csv"]
DETAY_DOSYALAR = ["urun_detay.xlsx", "urun_detay.csv"]
EKLENEN_DOSYA = "eklenen_gunluk.csv" # uygulamadan girilen günlük satışlar

GUN_TR = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

# Dini bayramlar (başlangıç, bitiş - dahil). Yeni yıl için buraya ekleyin.
BAYRAMLAR = [
("2023-04-21", "2023-04-23"), ("2023-06-28", "2023-07-01"),
("2024-04-10", "2024-04-12"), ("2024-06-16", "2024-06-19"),
("2025-03-30", "2025-04-01"), ("2025-06-06", "2025-06-09"),
("2026-03-20", "2026-03-22"), ("2026-05-27", "2026-05-30"),