import streamlit as st
import pandas as pd
from io import BytesIO

# 📌 Titre de l'application
st.title("📊 Consolidateur de mots‑clés (feuille ➡️ unique)")

# 📝 Mode d'emploi
st.markdown("""
1. Téléversez un classeur **Excel (.xlsx)** comportant **une feuille par famille de mots‑clés**.  
   • Les mots‑clés peuvent être dans une ou plusieurs colonnes et/ou séparés par `|`.  
2. Cliquez sur **« Fusionner & Télécharger »**.  
3. Un fichier **XLSX** est généré avec **deux colonnes** :  
   • **Mot Clé** → le nom de la feuille  
   • **Occurrences** → la liste unique des mots‑clés, séparée par ` | `.
""")

# 🔹 Upload du classeur
xlsx_file = st.file_uploader("📂 Sélectionnez votre fichier Excel", type=["xlsx"])

if xlsx_file and st.button("🚀 Fusionner & Télécharger"):
    try:
        # 1️⃣ Charger toutes les feuilles
        sheets_dict = pd.read_excel(xlsx_file, sheet_name=None, header=None, engine="openpyxl")

        lignes = []
        for sheet_name, df in sheets_dict.items():
            # 2️⃣ Récupérer toutes les cellules non‑vides
            tokens = []
            for cell in df.values.flatten():
                if pd.isna(cell):
                    continue
                for part in str(cell).split("|"):
                    part = part.strip()
                    if part:
                        tokens.append(part)

            if not tokens:
                st.info(f"ℹ️ Feuille « {sheet_name} » ignorée (vide).")
                continue

            # 3️⃣ Déduplication en conservant l'ordre
            unique_tokens = list(dict.fromkeys(tokens))
            lignes.append({
                "Mot Clé": sheet_name,
                "Occurrences": " | ".join(unique_tokens)
            })

        if not lignes:
            st.warning("⚠️ Aucun mot‑clé trouvé dans le classeur.")
        else:
            # 4️⃣ DataFrame consolidé
            result_df = pd.DataFrame(lignes)
            st.subheader("🔎 Aperçu consolidé")
            st.dataframe(result_df, use_container_width=True)

            # 5️⃣ Export XLSX en mémoire
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                result_df.to_excel(writer, sheet_name="Consolidé", index=False)
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger le XLSX consolidé",
                data=buffer,
                file_name="mots_cles_consolides.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
