import streamlit as st
import pandas as pd
from io import BytesIO

# 📌 Titre de l'application
st.title("📊 Consolidateur d'occurrences (Excel ➡️ Excel)")

# 📝 Mode d'emploi
st.markdown("""
**Étapes** :
1. Téléversez un classeur **Excel (.xlsx)** contenant **une feuille par liste de mots-clés**.  
   • Chaque cellule peut contenir soit un mot-clé unique, soit une liste séparée par `|`.  
2. Cliquez sur **« Consolider & Télécharger »**.  
3. Récupérez un fichier **XLSX** avec une **seule feuille** listant chaque mot-clé et son nombre d'occurrences.
""")

# 🔹 Widget d’upload
xlsx_file = st.file_uploader("📂 Choisissez votre fichier Excel", type=["xlsx"])

if xlsx_file and st.button("🚀 Consolider & Télécharger"):
    try:
        # 1️⃣ Charger toutes les feuilles en dictionnaire de DataFrames
        sheets_dict = pd.read_excel(xlsx_file, sheet_name=None, header=None, engine="openpyxl")

        tokens = []
        for sheet_name, df in sheets_dict.items():
            # 2️⃣ Flatten de toutes les valeurs de la feuille
            for cell in df.values.flatten():
                if pd.isna(cell):
                    continue
                # découpe éventuelle par « | »
                for part in str(cell).split("|"):
                    part = part.strip()
                    if part:
                        tokens.append(part)

        if not tokens:
            st.warning("⚠️ Aucun mot-clé détecté dans le classeur.")
        else:
            # 3️⃣ Comptage des occurrences
            counts_df = (
                pd.Series(tokens)
                .value_counts()
                .reset_index(name="Occurrences")
                .rename(columns={"index": "Mot Clé"})
            )

            # 4️⃣ Aperçu
            st.subheader("🔎 Aperçu des occurrences consolidées")
            st.dataframe(counts_df, use_container_width=True)

            # 5️⃣ Export en XLSX en mémoire
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                counts_df.to_excel(writer, sheet_name="Occurrences", index=False)
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger le XLSX consolidé",
                data=buffer,
                file_name="occurrences_consolidées.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
