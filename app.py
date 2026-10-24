import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Clan Event Suche", layout="wide")
st.title("🏆 Total Battle — Clan Event Abfrage")

excel_path = "Master_Total_Battle.xlsx"

if os.path.exists(excel_path):
  try:
    # Wir laden das Blatt "Truhen_Tracker" und nutzen header=1 für die Spaltenüberschriften
    df = pd.read_excel(excel_path, sheet_name="Truhen_Tracker", header=1)

    # Suchfeld anzeigen
    spieler_name = st.text_input("🔍 Spielername eingeben:")

    if spieler_name:
      # Alle Spalten nach dem Namen durchsuchen
      mask = df.astype(str).apply(
          lambda col: col.str.contains(spieler_name, case=False, na=False)
      )
      ergebnis = df[mask.any(axis=1)]

      if not ergebnis.empty:
        st.success(f"  {len(ergebnis)} Treffer für '{spieler_name}' gefunden!")
        st.dataframe(ergebnis, use_container_width=True)
      else:
        st.warning(f"Keine Einträge für '{spieler_name}' gefunden.")
    else:
      st.info("💡 Gib oben einen Namen ein, um die Daten zu durchsuchen.")
      st.dataframe(df.head(20), use_container_width=True)

  except Exception as e:
    st.error(f"Fehler beim Lesen der Excel-Datei: {e}")
else:
  st.error("Die Excel-Datei wurde nicht gefunden!")
