import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Clan Event Suche", layout="wide")
st.title("🏆 Total Battle — Debugging")

excel_path = "Master_Total_Battle.xlsx"

if os.path.exists(excel_path):
  try:
    # 1. Alle verfügbaren Tabellenblätter (Reiter) anzeigen
    xls = pd.ExcelFile(excel_path)
    st.write("📂 **Verfügbare Tabellenblätter (Sheets) in der Excel:**", xls.sheet_names)

    # 2. Daten laden (header=1 anpassen falls nötig)
    df = pd.read_excel(excel_path, header=1)
    
    st.write("📊 **Gefundene Spalten in dieser Tabelle:**", df.columns.tolist())

    # 3. Die Tabelle anzeigen (horizontal scrollbar)
    st.dataframe(df, use_container_width=True)

  except Exception as e:
    st.error(f"Fehler beim Lesen der Excel-Datei: {e}")
else:
  st.error("Die Excel-Datei wurde nicht gefunden!")
