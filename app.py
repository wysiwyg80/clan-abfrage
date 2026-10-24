import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Clan Event Suche", layout="wide")
st.title("🏆 Total Battle — Clan Event & Tracker Abfrage")

excel_path = "Master_Total_Battle.xlsx"

if os.path.exists(excel_path):
  try:
    # 1. Alle verfügbaren Tabellenblätter auslesen
    xls = pd.ExcelFile(excel_path)
    sheet_names = xls.sheet_names

    # 2. Ein Auswahlmenü (Dropdown) für die Reiter oben in der App erstellen
    selected_sheet = st.selectbox("📂 Wähle den Bereich aus:", sheet_names)

    # 3. Das ausgewählte Blatt laden (header=1 für die richtige Zeile)
    df = pd.read_excel(excel_path, sheet_name=selected_sheet, header=1)

    st.markdown("---")

    # 4. Suchfeld anzeigen (sucht im aktuell ausgewählten Tabellenblatt)
    spieler_name = st.text_input(
        f"🔍 Suche in '{selected_sheet}' (Spielername oder Begriff eingeben):"
    )

    if spieler_name:
      # Alle Spalten nach dem Begriff durchsuchen
      mask = df.astype(str).apply(
          lambda col: col.str.contains(spieler_name, case=False, na=False)
      )
      ergebnis = df[mask.any(axis=1)]

      if not ergebnis.empty:
        st.success(
            f"✅ {len(ergebnis)} Treffer für '{spieler_name}' in"
            f" '{selected_sheet}' gefunden!"
        )
        st.dataframe(ergebnis, use_container_width=True)
      else:
        st.warning(
            f"Keine Einträge für '{spieler_name}' in '{selected_sheet}' gefunden."
        )
    else:
      st.info(
          f"💡 Aktives Blatt: **{selected_sheet}**. Gib oben einen Suchbegriff ein"
          " oder scrolle durch die Tabelle:"
      )
      st.dataframe(df, use_container_width=True)

  except Exception as e:
    st.error(f"Fehler beim Lesen der Excel-Datei: {e}")
else:
  st.error("Die Excel-Datei wurde nicht gefunden!")
