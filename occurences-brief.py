import streamlit as st
import pandas as pd
from io import BytesIO

# 📌 Titre de l'application
st.title("📊 Consolidation Mot Clé Principal + Occurrences (Excel ➡️ Excel)")

st.markdown("""
**Mode d'emploi** :
1️⃣ Téléversez un classeur **.xlsx** comportant **une feuille par URL/cluster**.  
   • Les cellules **A5** et **B5** doivent contenir le *mot-clé principal* (ex. « robe » et « longue »).  
   • Les mots-clés secondaires peuvent être répartis dans la feuille (séparateur `|` géré).  
2️⃣ Cliquez sur **« Fusionner & Télécharger »**.  
3️⃣ Le fichier généré possède deux colonnes :  
   • **Mot Clé Principal**  
   • **Occurrences** (liste dédupliquée, séparée par ` | `).
""")

# 🔹 Upload du classeur
xlsx_file = st.file_uploader("📂 Charger le fichier Excel", type=["xlsx"])

if xlsx_file and st.button("🚀 Fusionner & Télécharger"):
    try:
        # 1️⃣ Lecture de toutes les feuilles
        sheets_dict = pd.read_excel(xlsx_file, sheet_name=None, header=None, engine="openpyxl")

        lignes = []
        for sheet_name, df in sheets_dict.items():
            # --- Mot-clé principal (A5 + B5) ---
            try:
                mot_cle_principal = f"{df.iloc[4, 0]} {df.iloc[4, 1]}".strip()
            except Exception:
                mot_cle_principal = ""  # fallback si structure inattendue

            # --- Collecte de tous les mots-clés ---
            tokens = []
            for cell in df.values.flatten():
                if pd.isna(cell):
                    continue
                for part in str(cell).split("|"):
                    part = part.strip()
                    if part:
                        tokens.append(part)

            if tokens:
                unique_tokens = list(dict.fromkeys(tokens))  # déduplication en conservant l'ordre
                lignes.append({
                    "Mot Clé Principal": mot_cle_principal,
                    "Occurrences": " | ".join(unique_tokens)
                })

        # 2️⃣ Vérification et export
        if not lignes:
            st.warning("⚠️ Aucun mot-clé détecté dans le classeur.")
        else:
            result_df = pd.DataFrame(lignes, columns=["Mot Clé Principal", "Occurrences"])
            st.subheader("🔎 Aperçu de la consolidation")
            st.dataframe(result_df, use_container_width=True, hide_index=True)

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                result_df.to_excel(writer, sheet_name="Consolidé", index=False, header=True)
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger le fichier XLSX consolidé",
                data=buffer,
                file_name="motcles_consolides.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
