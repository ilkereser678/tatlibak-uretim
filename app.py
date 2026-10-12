import os
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
          "Baz Ortalama": base_vals.round().astype(int),
          target_col_name: forecast_vals,
      })
      forecast_df = forecast_df.dropna(subset=["Pasta Çeşidi"])
      forecast_df = forecast_df[
          ~forecast_df["Pasta Çeşidi"]
          .astype(str)
          .str.contains("BATON", na=False)
      ]

      st.metric(
          label="Önerilen Toplam Pasta Üretimi",
          value=f"{int(forecast_df[target_col_name].sum())} Adet",
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
            index=0,
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
          pd.concat([df_w_up, new_w_row], ignore_index=True).to_csv(
              csv_weather_file, index=False
          )
          st.success("✅ Güncellendi!")
          st.rerun()

  # --- BİRLEŞTİRİLMİŞ MATRİS TABLO (KÜSURATSIZ VE TAM SAYI FORMATI) ---
  st.markdown("---")
  st.subheader("📂 Birleştirilmiş Ana Arşiv Tablosu")

  if df_all is not None:
    master_df = df_all.loc[:, ~df_all.columns.str.contains("^Unnamed")]
    product_col_name = master_df.columns[0]

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

    # Toplam satırını otomatik güncelleme
    total_row_mask = (
        master_df[product_col_name]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.contains("TOPLAM PASTA SAYISI", na=False)
    )

    if total_row_mask.any():
      products_only_df = master_df[~total_row_mask].iloc[1:]
      for col in master_df.columns:
        if col != product_col_name:
          numeric_col = pd.to_numeric(products_only_df[col], errors="coerce")
          if not numeric_col.isna().all():
            master_df.loc[total_row_mask, col] = int(numeric_col.sum())

    # --- KÜSURATLARI VE .0 İFADELERİNİ TAMAMEN KALDIRMA ---
    for col in master_df.columns:
      if col != product_col_name:
        # Önce sayısal formata çevir, NaN olanları (tarih satırı vb.) orijinalinde bırak
        numeric_series = pd.to_numeric(master_df[col], errors="coerce")
        # Null olmayan sayıları integer'a çevirip string yapıyoruz (böylece .0 görünmez)
        master_df[col] = numeric_series.apply(
            lambda x: str(int(x)) if pd.notnull(x) else master_df.loc[x.index if hasattr(x, 'index') else master_df.index, col] if False else master_df[col]
        )
        
        # Daha temiz bir döngü ile hücre bazlı dönüştürme:
        for idx in master_df.index:
          val = master_df.loc[idx, col]
          try:
            # Eğer sayıya dönüşebiliyorsa tam sayı yapalım (TARİH satırı hariç)
            if idx != 0 and pd.notnull(val) and str(val).strip() != "":
              num_val = float(val)
              master_df.loc[idx, col] = int(num_val)
          except:
            pass

    st.dataframe(master_df, use_container_width=True, height=500)