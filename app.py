import streamlit as st
import pandas as pd
import difflib
import re
from datetime import datetime
from rapidfuzz import fuzz

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
        menge = 1
        einheit = "STK"

        # Mengenextraktion robuster
        menge_search = re.search(r"(\d+(?:[\.,]\d+)?)(\s*(stk|st\.?|stück|meter|m|mtr))?", teil.lower())
        if menge_search:
            menge_text = menge_search.group(1).replace(",", ".")
            try:
                menge = int(float(menge_text))
            except:
                menge = 1
            if menge_search.group(3):
                einheit_raw = menge_search.group(3).lower()
                if "m" in einheit_raw:
                    einheit = "MTR"

        # Beste Übereinstimmung per RapidFuzz
        beschreibungen = df["bezeichnung"].astype(str)
        beste_score = 0
        bester_treffer = None

        for _, row in df.iterrows():
            score = fuzz.token_set_ratio(teil.lower(), str(row["bezeichnung"]).lower())
            if score > beste_score:
                beste_score = score
                bester_treffer = row

        if bester_treffer is not None and beste_score >= 60:
            artikel_ergebnisse.append({
                "Artikelnummer": str(bester_treffer.get("ean", "")),
                "Bezeichnung": bester_treffer["bezeichnung"],
                "Menge": menge,
                "Einheit": einheit,
                "EAN": bester_treffer.get("ean", "")
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
        artikelnummer = str(art['EAN'])[:15].ljust(15)
        bezeichnung = str(art['Bezeichnung'])[:60].ljust(60)
        lines.append(
            f"POA00000000{pos}000000000{pos}{artikelnummer}     {menge_str}{bezeichnung}000000000000           0000000000 H                   {art['Einheit']}2L"
        )

    lines.append("END")
    return "\n".join(lines)

# Streamlit UI
st.title("🧰 UGL-Generator für freie Artikeleingaben")

freitext = st.text_area("🎤 Du kannst hier auch per Sprache diktieren:",
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
