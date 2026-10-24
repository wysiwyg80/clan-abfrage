import datetime
import openpyxl
import pandas as pd
import streamlit as st

# ==========================================
# 1. PAGE CONFIGURATION & DESIGN (WEISS & ROT)
# ==========================================
st.set_page_config(
    page_title="BAD LegioXEquentris — Clan Tracker",
    page_icon="🩸",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS: Weißer Hintergrund + Blutrote Akzente + Dunkler Text für optimale Lesbarkeit
st.markdown(
    """
    <style>
    /* Gesamter Hintergrund weiß */
    .stApp {
        background-color: #FFFFFF;
        color: #111111;
    }
    
    /* Hauptüberschrift: BAD LegioXEquentris in kräftigem Blutrot */
    .main-title {
        color: #B22222 !important;
        text-shadow: 0px 0px 8px rgba(178, 34, 34, 0.2);
        text-align: center;
        font-size: 2.5rem;
        font-weight: 900;
        margin-bottom: 5px;
        letter-spacing: 1px;
    }
    
    /* Untertitel / Clan-Motto */
    .sub-title {
        color: #444444 !important;
        text-align: center;
        font-size: 1.2rem;
        font-style: italic;
        font-weight: 600;
        margin-bottom: 30px;
    }
    
    /* Alle Beschriftungen & Texte in schwarzer/dunkler Schrift */
    label, p, span, div, .stSelectbox label, .stTextInput label {
        color: #111111 !important;
        font-size: 15px;
        font-weight: 600;
    }
    
    /* Dropdown-Menü (Bereichsauswahl) mit rotem Rahmen auf weißem Grund */
    div[data-baseweb="select"] > div {
        background-color: #FAFAFA !important;
        color: #111111 !important;
        border: 2px solid #B22222 !important;
        border-radius: 8px;
    }
    
    /* Suchleiste für Spielernamen */
    div[data-baseweb="input"] > div {
        background-color: #FAFAFA !important;
        color: #111111 !important;
        border: 2px solid #B22222 !important;
        border-radius: 8px;
    }
    
    /* Input Text */
    input {
        color: #111111 !important;
    }
    
    /* Streamlit Tabellen-Design (Dataframe) */
    [data-testid="stDataFrame"] {
        border: 1px solid #E0E0E0;
        border-radius: 8px;
        background-color: #FFFFFF;
    }
    
    /* Info-Boxen in Rot/Weiß */
    .stAlert {
        background-color: #FFF5F5 !important;
        border: 1px solid #B22222 !important;
        color: #B22222 !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# ==========================================
# 2. DATEN EINLESEN & FORMELN BERECHNEN
# ==========================================
@st.cache_data(ttl=60)
def load_and_process_data(file_path="Master_Total_Battle.xlsx"):
  try:
    wb = openpyxl.load_workbook(file_path, data_only=True)

    # 1. Feed-Punkte berechnen (Truhen)
    feed_points = {}
    if "Erfasste_Truhen_Feed" in wb.sheetnames:
      ws_feed = wb["Erfasste_Truhen_Feed"]
      for r in range(2, ws_feed.max_row + 1):
        pname = ws_feed.cell(row=r, column=3).value
        pts = ws_feed.cell(row=r, column=6).value or 0
        if pname:
          pname_str = str(pname).strip()
          feed_points[pname_str] = feed_points.get(pname_str, 0) + int(pts)

    # 2. Turmbau-Punkte berechnen
    turm_points = {}
    if "Turmbau_Tracker" in wb.sheetnames:
      ws_turm = wb["Turmbau_Tracker"]
      for r in range(2, ws_turm.max_row + 1):
        pname = ws_turm.cell(row=r, column=2).value
        pts = ws_turm.cell(row=r, column=5).value or 0
        if pname:
          pname_str = str(pname).strip()
          turm_points[pname_str] = int(pts) if str(pts).isdigit() else 0

    # 3. Truhen_Tracker verarbeiten
    df_truhen = pd.read_excel(file_path, sheet_name="Truhen_Tracker")

    for idx, row in df_truhen.iterrows():
      pname = str(row["Spielername"]).strip()
      unter_100m = str(row.get("Unter_100M", "Nein")).strip()

      ziel = 12500 if unter_100m == "Ja" else 15000
      df_truhen.at[idx, "Ziel_Quote"] = ziel

      t_pts = feed_points.get(pname, 0)
      turm_pts = turm_points.get(pname, 0)
      g_pts = t_pts + turm_pts

      df_truhen.at[idx, "Truhen_Punkte"] = t_pts
      df_truhen.at[idx, "Turmbau_Punkte"] = turm_pts
      df_truhen.at[idx, "Gesamt_Punkte"] = g_pts
      df_truhen.at[idx, "Quote_Erreicht"] = "Ja" if g_pts >= ziel else "Nein"

    # Alle Tabellenblätter laden
    sheets = {}
    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
      if sheet_name == "Truhen_Tracker":
        sheets[sheet_name] = df_truhen
      else:
        sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name).fillna(
            ""
        )

    return sheets
  except Exception as e:
    st.error(f"❌ Fehler beim Laden der Excel-Datei: {e}")
    return None


# ==========================================
# 3. STREAMLIT USER INTERFACE
# ==========================================
st.markdown(
    '<div class="main-title">BAD LegioXEquentris</div>', unsafe_allow_html=True
)
st.markdown(
    '<div class="sub-title">🌸 Lieber Blumen im Asphalt als Blut im Gras 🩸</div>',
    unsafe_allow_html=True,
)

data_sheets = load_and_process_data()

if data_sheets:
  # Bereichs-Auswahl (Dropdown)
  selected_sheet = st.selectbox(
      "📁 Wähle den Bereich aus:", options=list(data_sheets.keys()), index=0
  )

  df_active = data_sheets[selected_sheet]

  # Suchleiste für Spielernamen
  search_term = st.text_input(
      f"🔍 Suche in '{selected_sheet}' (Spielername eingeben):", ""
  )

  # Filter anwenden
  if search_term and "Spielername" in df_active.columns:
    df_filtered = df_active[
        df_active["Spielername"]
        .astype(str)
        .str.contains(search_term, case=False, na=False)
    ]
  else:
    df_filtered = df_active

  st.info(f"💡 Aktives Blatt: **{selected_sheet}**")

  # Tabelle anzeigen
  st.dataframe(df_filtered, use_container_width=True, hide_index=True)
