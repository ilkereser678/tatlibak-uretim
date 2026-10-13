import datetime
import json
import os
from urllib.error import HTTPError
import urllib.request
from google.auth.transport.requests import Request
from google.oauth2 import service_account
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tatlıbak AI Üretim Planlama Asistanı",
    page_icon="🍰",
    layout="wide",
)

st.title(
    "🍰 Tatlıbak Yaş Pasta Doğrudan REST OAuth Gemini Muhakeme ve Uyarı Motoru"
)

# --- DOĞRUDAN SERVİS HESABI OAUTH TOKEN ÜRETME VE REST BAĞLANTISI ---
credentials_path = "credentials.json"


def get_gemini_rest_response(prompt_text):
  if not os.path.exists(credentials_path):
    return (
        "⚠️ `credentials.json` dosyası bulunamadı. Lütfen ana dizine"
        " eklediğinizden emin olun."
    )

  try:
    # 1. Servis hesabı üzerinden Google OAuth 2.0 Access Token alıyoruz
    credentials = service_account.Credentials.from_service_account_file(
        credentials_path,
        scopes=["https://www.googleapis.com/auth/generative-language"],
    )
    auth_request = Request()
    credentials.refresh(auth_request)
    access_token = credentials.token

    # 2. Gemini 1.5 Flash REST API endpoint'ini hazırlıyoruz
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"

    payload = {"contents": [{"parts": [{"text": prompt_text}]}]}

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )

    # 3. İsteği doğrudan gönderip yanıtı alıyoruz
    with urllib.request.urlopen(req) as response:
      res_data = json.loads(response.read().decode("utf-8"))
      # Yanıt yapısından modelin metnini çekiyoruz
      candidate = res_data.get("candidates", [{}])[0]
      content_part = (
          candidate.get("content", {}).get("parts", [{}])[0].get("text", "")
      )
      return (
          content_part
          if content_part
          else "Model boş bir yanıt döndürdü."
      )

  except HTTPError as e:
    error_body = e.read().decode("utf-8")
    return f"Gemini REST API HTTP Hatası ({e.code}): {error_body}"
  except Exception as ex:
    return f"Kimlik Doğrulama / Bağlantı Hatası: {ex}"