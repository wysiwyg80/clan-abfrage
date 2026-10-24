import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bad in Blut - Clan Tracker", layout="wide")

# --- CSS-Styling für den "Bad in Blut"-Effekt ---
st.markdown(
    """
    <style>
    /* Hintergrund der App etwas abdunkeln für den perfekten Look (optional) */
    .stApp {
        background-color: #0e1117;
    }
    
    /* Bluteffekt-Titel */
    .blood-title {
        font-size: 42px;
        font-weight: 900;
        color: #b30000; /* Dunkles Blutrot */
        text-align: center;
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.8), 
                     0 0 15px rgba(255, 0, 0, 0.6), 
                     0 0 30px rgba(179, 0, 0, 0.4);
        margin-bottom: 20px;
        font-family: 'Cinzel', 'Times New Roman', serif;
    }
    </style>
    
    <div class="blood-title">🩸 Bad in Blut — Total Battle Clan Tracker 🩸</div>
    """,
    unsafe_allow_html=True,
)

excel_path = "Master_Total_Battle.xlsx"

if os.path.exists(excel_path):
  try:
    xls = pd.ExcelFile(excel_path)
    sheet_names = xls.sheet_names

    selected_sheet = st.selectbox("📂 Wähle den Bereich aus:", sheet_names)
    df = pd.read_excel(excel_path, sheet_name=selected_sheet, header=1)

    st.markdown("---")

    spieler_name = st.text_input(
        f"🔍 Suche in '{selected_sheet}' (Spielername eingeben):"
    )

    if spieler_name:
      mask = df.astype(str).apply(
          lambda col: col.str.contains(spieler_name, case=False, na=False)
      )
      ergebnis = df[mask.any(axis=1)]

      if not ergebnis.empty:
        st.success(f"✅ {len(ergebnis)} Treffer gefunden!")
        st.dataframe(ergebnis, use_container_width=True)
      else:
        st.warning("Keine Einträge gefunden.")
    else:
      st.info(f"💡 Aktives Blatt: **{selected_sheet}**")
      st.dataframe(df, use_container_width=True)

  except Exception as e:
    st.error(f"Fehler beim Lesen der Excel-Datei: {e}")
else:
  st.error("Die Excel-Datei wurde nicht gefunden!")
