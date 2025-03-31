import streamlit as st
import pandas as pd
import difflib
import re
from datetime import datetime

# Artikeldaten laden
df = pd.read_excel("Artikelliste.xlsx")

# Spalten vereinheitlichen
df.columns = df.columns.str.lower()

# Bezeichnung aus mehreren Spalten zusammensetzen
beschreibung_cols = ["hersteller", "serie", "typ", "größe", "variante"]
df["bezeichnung"] = df[beschreibung_cols].astype(str).apply(lambda row: " ".join(row.values).strip(), axis=1)

# Artikel aus Texteingabe extrahieren
@st.cache_data
def finde_passende_artikel(eingabe, df):
    artikel_ergebnisse = []
    teile = [t.strip() for t in eingabe.split(",")]

    for teil in teile:
        menge_search = re.search(r"(\d+(?:[\.,]\d+)?)(\s*(stk|st\.?|meter|m|mtr))?", teil.lower())
        menge = menge_search.group(1).replace(",", ".") if menge_search else "1"
        try:
            menge = int(float(menge))
        except:
            menge = 1

        beschreibungen = df["bezeichnung"].astype(str)
        beschreibungen_liste = beschreibungen.str.lower().tolist()
        match = difflib.get_close_matches(teil.lower(), beschreibungen_liste, n=1, cutoff=0.3)

        if match:
            treffer = df[beschreibungen.str.lower() == match[0]].iloc[0]
            artikel_ergebnisse.append({
                "Artikelnummer": treffer["artikel nr."],
                "Bezeichnung": treffer["bezeichnung"],
                "Menge": menge,
                "Einheit": "STK" if menge > 1 else "MTR",
                "EAN": treffer.get("ean", "")
            })

    return artikel_ergebnisse

# UGL-Datei generieren
def erstelle_ugl(artikel_liste):
    heute = datetime.today().strftime("%Y%m%d")
    lines = [
        f"KOPWORO01    PMKA01    BE{' '*65}4001926        {heute}EUR04.00Bestellung{' '*26}{heute}",
        "ADR1684                                                        Lager                         Ludwigstr. 81-85                 63110 Rodgau - Jugesheim"
    ]

    for i, art in enumerate(artikel_liste, 1):
        menge_str = str(art["Menge"]).rjust(11, "0") + "0"
        pos = str(i).rjust(3, "0")
        lines.append(
            f"POA00000000{pos}000000000{pos}{art['Artikelnummer'][:15].ljust(15)}     {menge_str}{art['Bezeichnung'][:60].ljust(60)}000000000000           0000000000 H                   {art['Einheit']}2L"
        )

    lines.append("END")
    return "\n".join(lines)

# Streamlit UI
st.title("🧰 UGL-Generator für freie Artikeleingaben")

freitext = st.text_area("Gib hier die Artikel ein:",
                        "Kupferrohr 22 8 Meter, Press Bogen 22 90, 50er HT Rohr 0,5 m 3 Stck")

if st.button("🔍 Artikel erkennen und UGL erstellen"):
    artikel = finde_passende_artikel(freitext, df)

    if not artikel:
        st.warning("Keine passenden Artikel gefunden.")
    else:
        st.success("Folgende Artikel wurden erkannt:")
        st.table(artikel)

        ugl_text = erstelle_ugl(artikel)
        st.download_button(
            label="📄 UGL-Datei herunterladen",
            data=ugl_text,
            file_name="ugl.001",
            mime="text/plain"
        )
