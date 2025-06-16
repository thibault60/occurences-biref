import streamlit as st
import pandas as pd
from io import BytesIO

# 📌 Titre de l'application
st.title("📊 Fusion simple des occurrences (Excel ➡️ Excel)")

st.markdown("""
👉 Téléversez un classeur **.xlsx** avec *une feuille par thématique*.  
Le script extrait tous les mots‑clés (séparateur `|` géré), déduplique dans l'ordre, puis
produit **un seul fichier Excel** qui contient uniquement la liste d'occurrences —
*aucun autre champ n'est ajouté*. Chaque ligne correspond à une feuille source; il n'y a
pas d'en‑tête.
""")

xlsx_file = st.file_uploader("📂 Charger le fichier Excel", type=["xlsx"])

if xlsx_file and st.button("🚀 Fusionner & Télécharger"):
    try:
        sheets_dict = pd.read_excel(xlsx_file, sheet_name=None, header=None, engine="openpyxl")

        lignes = []
        for _, df in sheets_dict.items():
            tokens = []
            for cell in df.values.flatten():
                if pd.isna(cell):
                    continue
                for part in str(cell).split("|"):
                    part = part.strip()
                    if part:
                        tokens.append(part)
            if tokens:
                unique_tokens = list(dict.fromkeys(tokens))
                lignes.append([" | ".join(unique_tokens)])  # stocké comme liste pour DataFrame sans header

        if not lignes:
            st.warning("⚠️ Aucun mot-clé détecté.")
        else:
            # DataFrame sans intitulé de colonne
            result_df = pd.DataFrame(lignes)
            st.subheader("🔎 Aperçu (occurrences seules)")
            st.dataframe(result_df, use_container_width=True, hide_index=True)

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                result_df.to_excel(writer, sheet_name="Occurrences", index=False, header=False)
            buffer.seek(0)

            st.download_button(
                "📥 Télécharger le fichier XLSX",
                data=buffer,
                file_name="occurrences_uniques.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
